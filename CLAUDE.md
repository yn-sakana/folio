# Project Rules

## Language
- ランタイムのコードはすべてVBA。PowerShell, VBScript等は禁止（ビルドスクリプトは除く）
- WinAPI (Declare Function) も禁止。VBA標準機能のみ使用すること

## Architecture
- FEはUI＋設定のみ。データキャッシュはすべてBE
- BE→FE通信はTSVファイル経由（COM越し/シート経由は禁止）
- FEのデータ読み取りはDictionary引き（O(1)）
