# folio 仕様書

## 1. 概要

folio は Excel VBA アドインとして動作する案件管理ツール。
開いている Excel ワークブックのテーブル（ListObject）をリアルタイムで読み書きし、
メールアーカイブや案件フォルダと突合して一画面で管理する。

## 2. 基本原則

- **正本は Excel テーブルそのもの**。JSON 中間ファイルやスナップショットは持たない
- **フィールド検出はセルデータから行う**。VarType・NumberFormat で型判定する。フィールド名ヒューリスティクスやハードコードは使わない
- **リアルタイム双方向同期**。フォーム編集→即テーブル書き込み、テーブル変更→SheetChange で即反映
- **全変更をログに記録**。ローカル編集・外部変更・メール/フォルダ変化すべてを Change Log に流す
- **設定は Excel 内部に保存**。`_folio_config` / `_folio_log` 隠しシートを使う
- **WinAPI 禁止**。`Declare Function` はブロックされるため使用しない

## 3. 技術スタック

| レイヤー | 技術 |
|---------|------|
| アドイン本体 | Excel VBA (.xlsm) |
| GUI | MSForms UserForm（ランタイム生成コントロール） |
| データ読み書き | ListObject API（直接セルアクセス） |
| メール連携 | Outlook COM (下書き作成のみ) |
| メールエクスポート | FolioMailExport.bas (Outlook VBA 側、folio 本体には組み込まない) |
| 案件フォルダ | Dir$ / GetAttr / FileLen / FileDateTime (FSO 最小限) |
| 設定保存 | 隠しシート + JSON シリアライズ |
| BEワーカー | 別プロセスの Excel.Application (Visible=False) |
| BE→FE通信 | FEの隠しシートへの直接書き込み + Workbook_SheetChange |
| タイマー | Application.OnTime |

## 4. アーキテクチャ

### 4.1 FE/BE 分離

```text
FE: folio.xlsm (ユーザーの Excel インスタンス)
  ├── UI (frmFolio, frmSettings, etc.)
  ├── 設定管理 (FolioConfig)
  ├── FE側キャッシュ (FolioData の m_fe* 変数)
  │     └── 隠しシートからの読み込み（ローカル、O(1)）
  └── 隠しシート群
        _folio_config   設定 JSON
        _folio_log      変更ログ
        _folio_signal   シグナル (clock, version, timing)
        _folio_mail     メールレコード
        _folio_mail_idx メールインデックス
        _folio_cases    案件名一覧
        _folio_files    案件ファイルツリー
        _folio_diff     差分ログ

BE: 別プロセスの Excel.Application (Visible=False)
  ├── FolioWorker     スキャン実行 + FEシートへ書き込み
  └── FolioData       メール/案件フォルダの読み込み・差分検知
```

### 4.2 BE→FE通信フロー

```text
BE (FolioWorker)                          FE (ThisWorkbook)
  │                                         │
  │ スキャン (FolioData.Refresh*)           │
  │ ↓                                       │
  │ FEの隠しシートに配列一括書き込み        │
  │  g_feWb.Sheets("_folio_mail").Range = data
  │  g_feWb.Sheets("_folio_signal").Range = ver
  │                                      ───→ Workbook_SheetChange 発火
  │                                         │  └→ frmFolio.OnFolioSheetChange(shName)
  │                                         │       ├→ _folio_signal: version変化 → LoadDataFromLocalSheets()
  │                                         │       └→ _folio_diff: LogDiffsFromSheet()
  │                                         │
  │                                         │ LoadDataFromLocalSheets():
  │                                         │   ローカルシートを Range.Value で一括読み取り
  │                                         │   → m_feMailRecords, m_feMailIndex, etc. に格納
  │                                         │   → UI更新 (ステータス, メールタブ, ファイルタブ)
```

### 4.3 BEワーカー起動シーケンス

```text
ユーザー: Alt+F8 → Folio_ShowPanel
  → FolioMain.Folio_ShowPanel()
    → EnsureFolioSheets() (隠しシート作成)
    → frmFolio.Show vbModeless
      → UserForm_Initialize()
        → Application.OnTime Now, "FolioMain.DeferredStartup"
          → frmFolio.DoPollCycle()
            → FolioMain.StartWorker(mailFolder, caseRoot, matchField, matchMode)
              → CreateObject("Excel.Application")      ← 3-5秒ブロック (既知の課題)
              → Workbooks.Open ThisWorkbook.FullName
              → g_workerApp.Run "FolioWorker.WorkerEntryPoint", ..., ThisWorkbook
                → (BEプロセスで) WorkerInitialScan → ScheduleNextPoll (5秒間隔)
```

### 4.4 BEワーカータイマー

| タイマー | 間隔 | 内容 |
|---------|------|------|
| WorkerPollCallback | 5秒 | メール/案件の差分スキャン → 変更あればFEシート更新 |
| ClockCallback | 1秒 | FEの `_folio_signal` A1 に時刻を書き込み（※スキャン中は発火しない既知の課題） |

## 5. モジュール構成

### 5.1 FolioMain.bas

エントリポイントとBEワーカー管理。

- `Folio_ShowPanel()` — メインフォームを表示（Alt+F8 から実行）
- `Folio_ShowSettings()` — 設定フォームを表示
- `DeferredStartup()` — フォーム表示後に OnTime で呼ばれ、ワーカーを起動
- `StartWorker()` — BEプロセス作成 + xlsm をReadOnlyで開く + WorkerEntryPoint 呼び出し
- `StopWorker()` — BE を Quit
- `CleanupZombieWorker()` — PIDファイルから前回のBEプロセスを強制終了

**安全機構:**
- `g_formLoaded` — フォームがロードされているかのフラグ。`frmFolio.Visible` を参照すると VB_PredeclaredId=True のフォームが自動再生成されるため、直接参照を避ける
- `g_forceClose` — ワークブック閉じ時にフォームの UI 状態保存をスキップするフラグ
- PIDファイル (`.folio_cache/_worker.pid`) — ゾンビBE検知用

### 5.2 FolioWorker.bas

BEプロセスで実行されるスキャンモジュール。

- `WorkerEntryPoint(mailFolder, caseRoot, matchField, matchMode, feWorkbook)` — BE初期化。`g_feWb` にFEのWorkbook参照を保持
- `WorkerInitialScan()` — 全データスキャン + FEシート書き込み + シグナル通知
- `WorkerPollCallback()` — 5秒間隔で差分スキャン。変更があればFEシート更新
- `UpdateConfig()` — FE側の設定変更時にBEのスキャン設定を更新
- `WorkerStop()` — タイマーキャンセル + FE参照解放

**FEシート書き込みメソッド:**

| メソッド | 対象シート | 内容 |
|---------|-----------|------|
| WriteMailToFE | _folio_mail | メールレコード (10列: entry_id, sender_email, ...) |
| WriteMailIndexToFE | _folio_mail_idx | メールインデックス (2列: key, entry_id) |
| WriteCasesToFE | _folio_cases | 案件名一覧 (1列) |
| WriteCaseFilesToFE | _folio_files | ファイルツリー (7列: case_id, file_name, ...) |
| WriteDiffToFE | _folio_diff | 差分ログ (4列: action, type, id, description) |
| WriteSignalToFE | _folio_signal | バージョン + タイミング情報 |
| WriteClockToFE | _folio_signal | 時刻文字列 (A1) |

### 5.3 FolioData.bas

BE側とFE側の両方で使用されるデータモジュール。

**BE側（ワーカープロセスで使用）:**
- `m_mailRecords`, `m_mailIndex`, `m_caseNames`, `m_caseFiles` — BE側キャッシュ
- `RefreshMailData(folderPath)` — Dir$ ベースのメールスキャン。`meta.json` をパースして差分検知
- `RefreshCaseNames(rootPath)` — 案件フォルダ一覧の差分検知
- `RefreshCaseFiles(rootPath)` — 案件ファイルツリーの差分スキャン（フォルダ mod time で変更検知）
- `GetMailRecords()`, `GetMailIndex()`, `GetCaseNames()`, etc. — FolioWorker からのアクセサ

**FE側（フォーム表示プロセスで使用）:**
- `m_feMailRecords`, `m_feMailIndex`, `m_feCaseNames`, `m_feCaseFiles` — FE側キャッシュ
- `LoadFromLocalSheets(wb)` — FEの隠しシートから Range.Value 一括読み取り → Dictionary に格納
- `GetMailCount()`, `GetCaseCount()` — FE側カウント
- `FindMailRecords(keyValue, matchField, matchMode)` — O(1) インデックス引きでメール検索

**テーブル操作（FEのみ）:**
- `FindTable(wb, name)`, `ReadTableRecords(tbl)`, `WriteTableCell(tbl, row, col, val)`

### 5.4 frmFolio.frm

メインフォーム。全コントロールをランタイム生成する。

**レイアウト（3カラム）:**

```text
┌──────────┬─────────────────┬──────────┐
│ 左       │ 中央            │ 右       │
│ ソース   │ MultiPage       │ Change   │
│ フィルタ │  基本: 編集欄   │ Log      │
│ レコード │  メール: 一覧   │          │
│ リスト   │  ファイル: ツリー│          │
├──────────┼─────────────────┤          │
│ 件数     │ ボタン群        │          │
└──────────┴─────────────────┴──────────┘
  ステータスバー | Active (v1 mail:200 cases:50) | 時刻
```

**主要メソッド:**

| メソッド | 責務 |
|---------|------|
| `UserForm_Initialize` | レイアウト構築、設定読み込み、DeferredStartup をスケジュール |
| `SwitchSource` | テーブル切替、フィールド設定初期化、ワーカー起動/再設定 |
| `OnFolioSheetChange` | BEの書き込みを検知。シグナル変化時にデータ読み込み |
| `LoadDataFromLocalSheets` | FE隠しシートから Dictionary に読み込み + UI更新 |
| `LogDiffsFromSheet` | _folio_diff シートから差分ログを読み取り表示 |
| `DoPollCycle` | ワーカー起動（初回のみ） |
| `OnFieldChanged` | FieldEditor からのコールバック。テーブル書き込み + ログ記録 |
| `OnTableChanged` | SheetWatcher からのコールバック。テーブル変更検知 |

### 5.5 FieldEditor.cls

WithEvents パターンでテキストボックスの Change イベントを捕捉する。

- `SetValue(val)` — m_suppressing=True で値セット（Change イベント抑制）
- `RefreshIfChanged(val)` — 外部変更検知用。prevValue と異なり、かつユーザー未編集なら更新

### 5.6 SheetWatcher.cls

WithEvents でデータテーブル（ListObject）の変更を監視する。

- `Watch(ws, tableName, callback)` — 対象シートの Change イベントを監視
- テーブル範囲と交差する変更があれば `callback.OnTableChanged` を呼ぶ

### 5.7 FolioConfig.bas

プロファイルベースの設定管理。`_folio_config` 隠しシートに JSON で保存。

**設定 JSON 構造:**

```json
{
  "self_address": "",
  "mail_folder": "C:\\...\\mail",
  "case_folder_root": "C:\\...\\cases",
  "sources": {
    "anken": {
      "key_column": "案件ID",
      "display_name_column": "団体名",
      "mail_link_column": "メールアドレス",
      "folder_link_column": "案件ID",
      "mail_match_field": "sender_email",
      "mail_match_mode": "exact",
      "field_settings": { ... }
    }
  },
  "ui_state": { "window_width": 870, ... }
}
```

### 5.8 FolioChangeLog.bas

`_folio_log` 隠しシートに変更履歴を永続化。最大 5,000 行ローテーション。

### 5.9 FolioHelpers.bas

汎用ユーティリティ。JSON パーサー/シリアライザー、Dictionary ヘルパー、ファイル操作。

### 5.10 FolioBundler.bas

全モジュールのコードを単一の `.bas` インストーラーファイルにエクスポート。

### 5.11 FolioMailExport.bas (Outlook VBA)

Outlook VBA にインポートして使うメールエクスポートモジュール。**folio 本体には組み込まない。**

### 5.12 ErrorHandler.cls

全モジュール共通のエラーハンドラ。`Enter` / `OK` / `Catch` パターン。

## 6. データフロー

### 6.1 フォーム編集 → テーブル書き込み

```text
ユーザーがテキストボックスを編集
  → FieldEditor.m_txt_Change()
    → frmFolio.OnFieldChanged(field, old, new, "local")
      → FolioData.WriteTableCell(tbl, rowIdx, field, newVal)
      → FolioChangeLog.AddLogEntry(...)
      → Undo スタックに追加
```

### 6.2 外部変更 → フォーム反映

```text
データ xlsx のセルが変更される
  → SheetWatcher.m_ws_Change()
    → frmFolio.OnTableChanged()
      → RefreshCurrentRecord()
        → 各 FieldEditor の値をテーブルセルと比較
        → 差分あり → OnFieldChanged(field, old, new, "external")
```

### 6.3 メール/フォルダ変化の検知

```text
BE: WorkerPollCallback() (5秒間隔)
  → FolioData.RefreshMailData() / RefreshCaseNames() / RefreshCaseFiles()
  → 変更あり → WriteMailToFE / WriteCasesToFE / WriteDiffToFE
  → WriteVersionToFE(newVersion)
FE: Workbook_SheetChange("_folio_signal")
  → OnFolioSheetChange → LoadDataFromLocalSheets
  → UI更新
```

## 7. GUI 仕様

### 7.1 ウィンドウサイズ

設定から復元（`window_width`, `window_height`）。frmResize で変更。

### 7.2 中央カラム — タブ

- **基本タブ**: テーブルの各フィールドを TextBox で表示・編集。フィールド名プレフィクスでタブ自動分割
- **メールタブ**: 選択レコードに紐づくメール一覧 + 本文プレビュー + 添付ファイル
- **ファイルタブ**: 選択レコードに紐づく案件フォルダのファイルツリー

### 7.3 Undo

Ctrl+Z で直前のローカル編集を元に戻す（最大50件スタック）。

## 8. ビルド

### 8.1 Build-Addin.ps1

VBA ソースファイルから `folio.xlsm` を生成する。

**処理:**
1. 空のワークブックを作成
2. .bas モジュールをインポート
3. UserForm を作成しコードを注入
4. .cls モジュールを作成しコードを注入
5. ThisWorkbook に `Workbook_BeforeClose` + `Workbook_SheetChange` を注入
6. `_folio_config` / `_folio_log` + FEデータ用隠しシートを作成
7. サンプルパスで初期設定を書き込み
8. .xlsm で保存

### 8.2 samplerun.bat

```bat
Build-Addin.ps1 -Sample  # ビルド
start folio.xlsm          # アドイン起動
start sample/folio-sample.xlsx  # データソース起動
```

## 9. 設計上の制約と決定

| 項目 | 決定 | 理由 |
|------|------|------|
| WinAPI 禁止 | VBA 標準機能のみ | 本番環境のポリシーで Declare がブロックされる |
| BEワーカー分離 | 別プロセスの Excel.Application | メール/フォルダスキャンがFEをブロックしないように |
| シート直書き | BE→FEの隠しシートに配列書き込み | TSVファイル経由より速く、SheetChange で即時通知 |
| FEデータ読み取りはローカルのみ | FE側の隠しシートから読む | cross-process COM 呼び出しを避ける |
| ControlSource 不使用 | コードベースで読み書き | 異なるワークブック間で動作しない |
| frmFolio.Visible 直接参照禁止 | g_formLoaded フラグで管理 | VB_PredeclaredId=True で自動再生成される |
| フォーム閉じ時に全参照解放 | CleanupRefs で Nothing 代入 | 外部ワークブックの ListObject 参照が残ると COM デッドロック |

## 10. 既知の課題とリファクタリング計画

### 10.1 WinAPI 違反

`FolioMain.bas` で `GetWindowThreadProcessId` を使用中（ゾンビPID取得）。
WinAPI 禁止ルールに違反。PID取得を VBA 標準機能で代替するか、ゾンビ検知方式を見直す必要あり。

### 10.2 FE起動時フリーズ (3-5秒)

`StartWorker` 内の `CreateObject("Excel.Application")` + `Workbooks.Open` が同期実行でFEをブロック。
VBA標準では非同期化できない。フォーム表示前に実行して体感を改善する案あり。

### 10.3 BE時計のスキャン中停止

BEの `ClockCallback` (1秒) がスキャン中に発火しない（VBA単一スレッド）。
時計をFE側の `Application.OnTime` に移す案あり。

### 10.4 FolioData.bas の肥大化

BE側キャッシュ + FE側キャッシュ + テーブル操作が1ファイルに混在（880行超）。
役割ごとにモジュール分割を検討。

### 10.5 旧コードの残骸

- `FolioData.bas` のコメント「FE reads _signal.txt to detect changes」— 旧TSV方式の記述が残存
- `FolioData.LoadFromSheets(beWb)` — 旧方式のBEワークブック直読みメソッドが残存
- `FolioData.GetCaseFilesTsvContent()` — TSV出力用メソッドが残存
- `frmFolio.frm` のコメント「Worker Signal (via _signal.txt file polling)」「Worker SheetChange callback (called by WorkerWatcher)」— 旧方式のコメントが残存
- `src/WorkerWatcher.cls` — 未使用ファイル（ビルド対象外）。削除可
