# Project Rules

## Language
- ランタイムのコードはすべてVBA。PowerShell, VBScript等は禁止（ビルドスクリプトは除く）
- WinAPI (Declare Function) は禁止。VBA標準機能のみ使用すること

## Architecture
- FE (folio.xlsm) = UI＋設定のみ。データキャッシュはすべてBE
- BE = 別プロセスの Excel.Application (Visible=False)。FolioWorker がスキャン実行
- BE→FE通信はFEの隠しシートへの直接書き込み（`_folio_signal`, `_folio_mail` 等）
- FEは `Workbook_SheetChange` イベントで変更を検知（ポーリング不要）
- FEのデータ読み取りはDictionary引き（O(1)）

## Build & Test
- ビルド: `powershell -ExecutionPolicy Bypass -File scripts/Build-Addin.ps1`
- サンプル実行: `samplerun.bat`（ビルド→xlsm＋folio-sample.xlsx を開く）
- テスト: `powershell -ExecutionPolicy Bypass -File scripts/Test-Worker.ps1`
