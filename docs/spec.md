# folio 仕様書

## 1. 概要

folio は Excel VBA アドインとして動作する案件管理ツール。
開いている Excel ワークブックのテーブル（ListObject）をリアルタイムで読み書きし、
メールアーカイブや案件フォルダと突合して一画面で管理する。

## 2. 基本原則

- **正本は Excel テーブルそのもの**。JSON 中間ファイルやスナップショットは持たない
- **フィールド検出はセルデータから行う**。VarType・NumberFormat で型判定する。フィールド名ヒューリスティクスやハードコードは使わない
- **リアルタイム双方向同期**。フォーム編集→即テーブル書き込み、テーブル変更→ポーリングでフォーム反映
- **全変更をログに記録**。ローカル編集・外部変更・メール/フォルダ変化すべてを Change Log に流す
- **設定は Excel 内部に保存**。`_folio_config` / `_folio_log` 隠しシートを使う

## 3. 技術スタック

| レイヤー | 技術 |
|---------|------|
| アドイン本体 | Excel VBA (.xlsm / .xlam) |
| GUI | MSForms UserForm（ランタイム生成コントロール） |
| データ読み書き | ListObject API（直接セルアクセス） |
| メール連携 | Outlook COM (下書き作成のみ) |
| 案件フォルダ | FileSystemObject |
| 設定保存 | 隠しシート + JSON シリアライズ |
| ポーリング | Application.OnTime |
| リサイズ | Win32 API (SetWindowLong) |

## 4. アーキテクチャ

```text
Excel ワークブック (*.xlsx)
  └── テーブル (ListObject)
        │
        │  直接読み書き
        ▼
folio アドイン (*.xlsm)
  ├── frmFolio (メインフォーム)
  │     ├── 左: ソース選択 + フィルタ + レコード一覧
  │     ├── 中: タブ (基本 / メール / ファイル / 通知)
  │     └── 右: Change Log
  ├── FieldEditor (WithEvents で変更検知)
  ├── FolioConfig (プロファイル管理)
  ├── FolioData (テーブル・メール・フォルダ読み書き)
  ├── FolioChangeLog (変更ログ永続化)
  ├── FolioMain (エントリポイント + ポーリング)
  ├── FolioHelpers (JSON・Dict・ファイル操作)
  ├── FolioOutlook (Outlook 下書き作成)
  └── _folio_config / _folio_log (隠しシート)
```

## 5. モジュール構成

### 5.1 FolioMain.bas

エントリポイントとポーリングタイマー。

- `Folio_ShowPanel()` — メインフォームを表示（Alt+F8 から実行）
- `Folio_ShowSettings()` — 設定フォームを表示
- `PollCallback()` — Application.OnTime で定期実行。frmFolio.DoPollCycle を呼ぶ

### 5.2 frmFolio.frm

メインフォーム。全コントロールをランタイム生成する。

**レイアウト（3カラム）:**

```text
┌──────────┬─────────────────┬──────────┐
│ 左       │ 中央            │ 右       │
│ ソース   │ MultiPage       │ Change   │
│ フィルタ │  基本: 編集欄   │ Log      │
│ レコード │  メール: 一覧   │          │
│ リスト   │  ファイル: ツリー│          │
│          │  通知: 設定     │          │
├──────────┼─────────────────┤          │
│ 件数     │ Sync | Settings │          │
└──────────┴─────────────────┴──────────┘
  ステータスバー
```

**主要メソッド:**

| メソッド | 責務 |
|---------|------|
| `UserForm_Initialize` | レイアウト構築、設定読み込み、ソース一覧取得 |
| `SwitchSource` | テーブル切替、フィールド設定初期化、レコード一覧更新 |
| `UpdateRecordList` | フィルタ適用、一覧表示（in_list カラムを表示） |
| `UpdateDetail` | 選択レコードの詳細表示（FillFieldEditors + タブ更新） |
| `FillFieldEditors` | テーブルセルから直接読み取り、FieldEditor に値セット |
| `OnFieldChanged` | FieldEditor からのコールバック。テーブル書き込み + ログ記録 |
| `DoPollCycle` | リサイズ検知 + RefreshCurrentRecord + RefreshJoinedData |
| `RefreshCurrentRecord` | 全 FieldEditor の値をテーブルと照合、差分があれば更新 |
| `RefreshJoinedData` | メール・フォルダを再スキャンし、件数変化をログ |

**データワークブック検出:**
- ThisWorkbook 以外で ListObject を持つワークブックを自動検出

### 5.3 FieldEditor.cls

WithEvents パターンでテキストボックスの Change イベントを捕捉する。

```text
TextBox.Change
  → m_txt_Change()
    → (m_suppressing なら無視)
    → m_form.OnFieldChanged(fieldName, oldVal, newVal, origin)
```

- `SetValue(val)` — m_suppressing=True で値セット（Change イベント抑制）
- `RefreshIfChanged(val)` — 外部変更検知用。prevValue と異なり、かつユーザー未編集なら更新

**origin 判定:**
- `m_isReadOnly = False` → `"local"`（ユーザー編集）
- `m_isReadOnly = True` → `"external"`（外部変更の反映）

### 5.4 FolioConfig.bas

プロファイルベースの設定管理。`_folio_config` 隠しシートに JSON で保存。

**シート構造:**

| 行 | A列 | B列 |
|----|-----|-----|
| 1 | active_profile | default |
| 3 | profile_name | config_json |
| 4 | default | {JSON} |
| 5 | (追加プロファイル) | {JSON} |

**設定 JSON 構造:**

```json
{
  "self_address": "",
  "mail_folder": "C:\\...\\mail",
  "case_folder_root": "C:\\...\\cases",
  "poll_interval": 5,
  "sources": {
    "anken": {
      "key_column": "案件ID",
      "display_name_column": "団体名",
      "mail_link_column": "メールアドレス",
      "folder_link_column": "案件ID",
      "field_settings": {
        "案件ID": { "type": "text", "in_list": true, "editable": false, "multiline": false },
        "団体名": { "type": "text", "in_list": true, "editable": true, "multiline": false }
      }
    }
  },
  "ui_state": {
    "window_width": 870,
    "window_height": 540,
    "left_width": 250,
    "right_width": 250,
    "selected_source": "",
    "search_text": ""
  }
}
```

**自動検出:**

| 項目 | ロジック |
|------|---------|
| key_column | 先頭50行で全値がユニーク＆非空の最初のカラム |
| display_name_column | key_column の次のテキスト型カラム |
| mail_link_column | 値に `@` を含む最初のカラム |
| folder_link_column | key_column と同じ |
| field type | `VarType(cell.Value)`: vbDate→date, vbDouble等→number, その他→text |
| multiline | 先頭10行に改行または100文字超の値があれば true |

### 5.5 FolioData.bas

テーブル・メール・フォルダのデータアクセス。

**テーブル操作:**
- `GetWorkbookTableNames(wb)` — ワークブック内の全テーブル名を返す
- `FindTable(wb, name)` — 名前でテーブルを検索
- `ReadTableRecords(tbl)` — 全行を Dictionary の Collection として返す
- `WriteTableCell(tbl, rowIdx, colName, val)` — 単一セルに書き込む
- `GetTableColumnNames(tbl)` — `_` プレフィクスを除外したカラム名一覧

**メールアーカイブ:**
- `ReadMailArchive(folderPath)` — フォルダを再帰スキャンし、meta.json を持つフォルダをメールレコードとして収集
- 添付ファイルパスを絶対パスに解決

**案件フォルダ:**
- `ReadCaseFolders(rootPath)` — ルート直下のサブフォルダを案件として再帰スキャン
- `CreateCaseFolder(rootPath, caseId, displayName)` — 案件フォルダを作成

**結合:**
- `FindJoinedRecords(records, keyField, keyValue, matchMode)` — exact / domain マッチ

### 5.6 FolioChangeLog.bas

`_folio_log` 隠しシートに変更履歴を永続化。

**シート構造:**

| A | B | C | D | E | F | G |
|---|---|---|---|---|---|---|
| timestamp | source | key | field | old_value | new_value | origin |

**仕様:**
- 最大 5,000 行。超過時は古い行を削除（ローテーション）
- `GetRecentEntries(count)` — 最新 N 件を取得（新しい順）
- `FormatLogLine(entry)` — 表示用フォーマット

**origin 値:**

| origin | 意味 |
|--------|------|
| local | ユーザーがフォームで編集 |
| external | ポーリングで検知した外部変更 |
| scan | メール/フォルダの件数変化 |

### 5.7 FolioHelpers.bas

汎用ユーティリティ。

- **JSON パーサー/シリアライザー** — 外部ライブラリ不要の自前実装
- **Dictionary ヘルパー** — `NewDict`, `DictStr`, `DictObj`, `DictBool`, `DictLng`, `DictPut`
- **ファイル操作** — `ReadTextFile`, `WriteTextFile`, `EnsureFolder`, `FileExists`, `FolderExists`
- **文字列** — `SafeName`, `FormatFieldValue`, `GetFieldLabel`, `GetFieldGroup`

**FormatFieldValue の型判定:**
- `date`: `VarType(val) = vbDate` の場合のみ `yyyy/mm/dd` にフォーマット
- `number`: IsNumeric なら `#,0` / `#,0.##`
- それ以外: CStr

### 5.8 FolioOutlook.bas

Outlook COM 経由で返信下書きを作成。送信はしない。

### 5.9 ErrorHandler.cls

全モジュール共通のエラーハンドラ。Immediate ウィンドウにトレースを出力。

```vba
Dim eh As New ErrorHandler: eh.Enter "ModuleName", "ProcName"
On Error GoTo ErrHandler
' ...
eh.OK: Exit Sub
ErrHandler: eh.Catch
```

### 5.10 frmSettings.frm

設定フォーム（モーダル）。

- **Paths タブ**: Self address, Mail folder, Case folder, Poll interval
- **Sources タブ**: ソース選択、key/name/mail/folder カラム設定、フィールド設定

## 6. データフロー

### 6.1 フォーム編集 → テーブル書き込み

```text
ユーザーがテキストボックスを編集
  → FieldEditor.m_txt_Change()
    → frmFolio.OnFieldChanged(field, old, new, "local")
      → FolioData.WriteTableCell(tbl, rowIdx, field, newVal)
      → FolioChangeLog.AddLogEntry(...)
      → Change Log UI に追加
      → Undo スタックに追加
```

### 6.2 外部変更 → フォーム反映

```text
Application.OnTime (N秒間隔)
  → FolioMain.PollCallback()
    → frmFolio.DoPollCycle()
      → RefreshCurrentRecord()
        → 各 FieldEditor の値をテーブルセルと比較
        → 差分あり → FieldEditor.RefreshIfChanged(newVal)
          → TextBox.Text を更新
          → Change イベント発火
          → OnFieldChanged(field, old, new, "external")
            → ログ記録（テーブル書き込みはしない）
```

### 6.3 メール/フォルダ変化の検知

```text
DoPollCycle()
  → RefreshJoinedData()
    → ReadMailArchive() / ReadCaseFolders()
    → 前回件数と比較
    → 差分あり → ログ記録 ("scan")
    → タブ表示を更新
```

## 7. GUI 仕様

### 7.1 レイアウト

- 3カラム構成（左/中央/右）、リサイズ可能
- デフォルトサイズ: 870 x 540
- Win32 API で WS_THICKFRAME を付与しリサイズ対応
- ポーリングでサイズ変更を検知し RepositionControls を呼ぶ

### 7.2 左カラム

- **ソース選択** (ComboBox): データワークブック内のテーブル一覧
- **フィルタ** (TextBox): テキスト入力で即時絞り込み
- **レコード一覧** (ListBox): `in_list=true` のカラムを表示。key_column + display_name_column が基本

### 7.3 中央カラム — タブ

**基本タブ:**
- テーブルの各フィールドを TextBox で表示・編集
- editable=false のフィールドは読み取り専用
- multiline=true のフィールドは高さを広げる
- フィールドはテーブルのカラム順で表示

**メールタブ:**
- 選択レコードに紐づくメール一覧
- 添付ファイル一覧（ダブルクリックで開く）

**ファイルタブ:**
- 選択レコードに紐づく案件フォルダのファイルツリー
- ツリー表示（インデントで階層表現）
- ダブルクリックでファイル/フォルダを開く
- フォルダ作成ボタン

**通知タブ:**
- （将来拡張用）

### 7.4 右カラム — Change Log

- 全変更をリアルタイム表示
- 表示形式: `時刻 | ソース | キー | フィールド | 旧値→新値 | origin`
- 新しいエントリが上
- クリアボタンあり

### 7.5 Undo

- Ctrl+Z で直前のローカル編集を元に戻す
- Undo スタック（最大50件）
- WriteTableCell で元の値を書き戻し、FieldEditor の Change イベントでログ記録

### 7.6 ステータスバー

- 最後の変更情報を表示: `origin: fieldName @ hh:nn:ss`

## 8. ポーリング

- デフォルト間隔: 5秒（設定で変更可能、最小1秒）
- `Application.OnTime` で実装（VBA で唯一の非同期的メカニズム）
- フォーム表示中のみ動作（`g_pollActive` フラグで制御）
- フォーム閉じ時に `Application.OnTime ..., False` でキャンセル

**ポーリング内容:**
1. フォームリサイズ検知
2. 現在レコードの全フィールドをテーブルと照合
3. メールアーカイブ再スキャン
4. 案件フォルダ再スキャン

## 9. ビルド

### 9.1 Build-Addin.ps1

VBA ソースファイルから `folio.xlsm` を生成する。

**前提条件:**
- Excel > トラストセンター > 「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」が ON

**処理:**
1. 空のワークブックを作成
2. .bas モジュールをインポート
3. UserForm を作成しコードを注入
4. .cls モジュールを作成しコードを注入
5. `_folio_config` / `_folio_log` 隠しシートを作成
6. サンプルパスで初期設定を書き込み
7. .xlsm で保存

### 9.2 Build-Sample.ps1

元データ（shinsa リポジトリの Excel ファイル）から `folio-sample.xlsx` を生成する。

**処理:**
1. 各ソース Excel を開き、値と書式をコピー
2. ListObject（テーブル）を作成
3. メール・案件フォルダをコピー

### 9.3 Test-Compile.ps1

ビルド済み xlsm にテストモジュールを注入し、各プロシージャの動作を検証する。

### 9.4 実行方法

```bat
build-addin.bat    :: folio.xlsm を生成
build-sample.bat   :: folio-sample.xlsx + サンプルデータを生成
run.bat            :: sample を開いてから addin を開く
```

## 10. ディレクトリ構成

```text
folio/
├── docs/
│   └── spec.md              ← この文書
├── src/
│   ├── FolioMain.bas
│   ├── FolioConfig.bas
│   ├── FolioData.bas
│   ├── FolioHelpers.bas
│   ├── FolioChangeLog.bas
│   ├── FolioOutlook.bas
│   ├── ErrorHandler.cls
│   ├── FieldEditor.cls
│   ├── frmFolio.frm
│   └── frmSettings.frm
├── scripts/
│   ├── Build-Addin.ps1
│   ├── Build-Sample.ps1
│   └── Test-Compile.ps1
├── sample/                   ← ビルド生成物（git 管理）
│   ├── folio-sample.xlsx
│   ├── mail/
│   └── cases/
├── build-addin.bat
├── build-sample.bat
├── run.bat
└── .gitignore                ← *.xlsm, *.xlam, ~$* を除外
```

## 11. 設計上の制約と決定

| 項目 | 決定 | 理由 |
|------|------|------|
| ControlSource 不使用 | コードベースで読み書き | ControlSource は異なるワークブック間で動作しない |
| スナップショット不使用 | テーブルを直接読み書き | 中間キャッシュは複雑性を増すだけ |
| VBA 単一スレッド | Application.OnTime でポーリング | VBA に非同期メカニズムがない |
| JSON 自前実装 | FolioHelpers 内 | VBA に標準 JSON ライブラリがない |
| フィールド名ヒューリスティクス禁止 | VarType/NumberFormat で型判定 | `IsDate("R06-001")` が日本語ロケールで True を返す問題 |
| 3-way merge 不使用 | 最後の書き込みが勝つ | Excel テーブル直接アクセスのため競合は OneDrive 側で管理 |

## 12. 今後の課題

- [ ] フィルタの絞り込み動作の検証
- [ ] フォルダツリーの階層表示の検証
- [ ] メール添付ファイル表示の検証
- [ ] Excel 閉じ時のフリーズ対策の検証（OnTime キャンセル）
- [ ] .xlam 形式での配布対応
- [ ] 通知タブの実装
- [ ] キーボードショートカットの充実
