# folio 仕様書

最終更新: 2026-03-15

## 1. 概要

folio は Excel VBA の案件管理ツール。テーブル (ListObject) をリアルタイムで読み書きし、メールアーカイブ・案件フォルダと突合して一画面で管理する。

## 2. 基本原則

- **正本は Excel テーブル**。中間ファイルは持たない
- **フィールド検出はセルデータから**。VarType・NumberFormat で型判定
- **リアルタイム双方向同期**。フォーム編集→即テーブル書き込み、テーブル変更→SheetChange で即反映
- **全変更をログに記録**。Change Log (ListObject, 5000行ローテーション)
- **設定は Dictionary キャッシュ**。起動時シートから一括ロード、終了時シリアライズ
- **WinAPI 禁止**。VBA 標準 + COM 標準のみ (Scripting.Dictionary, ADODB.Stream, WScript.Shell)

## 3. アーキテクチャ

### 3.1 FE/BE 分離

```
FE: folio.xlsm (ユーザーの Excel)
  ├── UI (frmFolio, frmSettings, frmResize)
  ├── FolioMain       エントリポイント + BE管理
  ├── FolioData       FE側キャッシュ + テーブル操作
  ├── FolioLib        Config + ChangeLog + ユーティリティ
  ├── ErrorHandler    エラー + ログ蓄積
  ├── FieldEditor     WithEvents テキストボックス
  ├── SheetWatcher    WithEvents テーブル監視
  └── 隠しシート群

BE: 別プロセスの Excel.Application (Visible=False)
  └── FolioWorker     スイッチ式スキャン + FEシート書き込み + リクエスト応答
```

### 3.2 BE→FE通信

BE が FE の隠しシートに `.Value` で書き込み → FE の `Workbook_SheetChange` が発火。

### 3.3 FE→BE通信 (リクエスト/レスポンス)

FE が BE の `_folio_request` シートに書き込み → BE の `SheetChange` → `Application.OnTime` で非同期処理 → 結果を FE のシートに書き込み → FE の `SheetChange` で受信。

### 3.4 スイッチ式スキャンループ

5秒ポーリングではなく、1秒チャンク + 1秒 Yield の連続ループ:

```
DoScanChunk (1秒枠, ラウンドロビン)
  TASK_MAIL:  manifest.tsv mtime チェック → 変化時は再読み
  TASK_CASES: root mtime チェック → 変化時は Dir$ 列挙
  TASK_WRITE: 変更があれば FE シートに書き込み + バージョン進行
  → OnTime → YieldCallback

YieldCallback
  時計更新 (WriteClockToFE)
  リクエスト確認・応答 (ProcessRequest)
  → OnTime(+1s) → DoScanChunk
```

ラウンドロビン (`g_nextTask`) で各タスクに公平に実行機会を与え、飢餓を防ぐ。通常運用は全タスク μs で通過。

## 4. 隠しシート

| シート | 用途 |
|--------|------|
| `_folio_config` | 設定 KV (起動時 Dict にロード) |
| `_folio_sources` | ソーステーブル設定 |
| `_folio_fields` | フィールド設定 |
| `_folio_log` | 変更ログ (ListObject "FolioLog") |
| `_folio_signal` | A1:時刻, B1:バージョン, C1:タイミング |
| `_folio_mail` | メールレコード (10列) |
| `_folio_mail_idx` | メールインデックス (key→entry_id) |
| `_folio_cases` | 案件名一覧 |
| `_folio_files` | ケースファイル (オンデマンド応答) |
| `_folio_diff` | 差分ログ |
| `_folio_request` | FE→BEリクエスト (A1:id, B1:type, C1+:params) |

## 5. メールスキャン

### manifest.tsv (高速パス)

エクスポータ (FolioMailExport) がメール追加時に `manifest.tsv` へ1行追記。スキャナは manifest.tsv の mtime をチェックし、変化時のみ全件再読みする (10万件 ~887ms)。

### Dir$ fallback (マイグレーション)

manifest.tsv がない場合、Dir$ + meta.json で全件スキャンし manifest.tsv を自動生成。初回のみ発生。

## 6. ケースファイル (オンデマンド)

起動時の全件プリロード廃止。ユーザーが案件を選択したとき:
1. FE → `_folio_request` に `case_files` + case_id を書き込み
2. BE → Dir$ で該当フォルダを走査 (~4ms)
3. BE → FE の `_folio_files` に結果を書き込み
4. FE → `SheetChange` で受信、ファイルタブを更新

## 7. 設定管理

- 起動時: 3つの隠しシート (`_folio_config`, `_folio_sources`, `_folio_fields`) → Dictionary に一括ロード
- 実行中: Dictionary 引き (O(1))。変更は Dict に書き込み + `m_dirty` フラグ
- 終了時: `BeforeWorkbookClose` → `SaveToSheets` でシートにシリアライズ

## 8. エラーハンドリング

```vba
Dim eh As New ErrorHandler: eh.Enter "Module", "Proc"
On Error GoTo ErrHandler
eh.Log "processing " & count & " items"
...
eh.OK: Exit Sub
ErrHandler: eh.Catch  ' エラー情報 + 蓄積ログ全件を Debug.Print
```

正常時は `OK` でログクリア。エラー時は `Catch` で蓄積ログ全件 + 経過時間を出力。

## 9. 制約

| 項目 | 決定 | 理由 |
|------|------|------|
| WinAPI 禁止 | VBA 標準 + COM 標準のみ | 本番環境のポリシーで Declare がブロック |
| BE分離 | 別プロセス Excel.Application | スキャンが FE をブロックしない |
| シート直書き | BE→FE 隠しシート + SheetChange | TSV ファイル経由より速い |
| ControlSource 不使用 | コードで読み書き | 異なるワークブック間で非対応 |
| frmFolio.Visible 直参照禁止 | g_formLoaded フラグ | VB_PredeclaredId=True で自動再生成される |
