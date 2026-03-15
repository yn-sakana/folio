# folio

Excel VBA 案件管理アドイン。開いているワークブックのテーブル (ListObject) をリアルタイムで読み書きし、メールアーカイブ・案件フォルダと突合して一画面で管理する。

## 基本原則

- **正本は Excel テーブルそのもの** — 中間ファイルやスナップショットは持たない
- **フィールド検出はセルデータから** — `VarType` / `NumberFormat` で型判定。ハードコード禁止
- **全変更をログに記録** — ローカル編集・外部変更・メール/フォルダ変化すべてを Change Log に流す
- **設定は Dictionary キャッシュ** — 起動時にシートから一括ロード、終了時にシリアライズ
- **WinAPI 禁止** — VBA 標準 + COM 標準オブジェクトのみ

## セットアップ

### 前提条件

- Windows + Excel (Microsoft 365 / 2021以降)
- Excel > ファイル > オプション > トラストセンター > マクロの設定 > **「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」を ON**

### ビルド・実行

```bat
samplerun.bat            # ビルド → xlsm + サンプルを開く
build-addin.bat          # folio.xlsm のみ生成
build-sample.bat         # folio-sample.xlsx + サンプルデータ生成
```

### 使い方

1. `samplerun.bat` を実行
2. `Alt+F8` → `Folio_ShowPanel`
3. 左上のドロップダウンからテーブルを選択
4. レコード選択 → 中央タブで編集・メール確認・ファイル確認
5. 右カラムに変更ログがリアルタイム表示

## アーキテクチャ

```
FE: folio.xlsm (ユーザーの Excel インスタンス)
  ├── frmFolio          メインUI (3カラム, ランタイム生成)
  ├── frmSettings       設定ダイアログ
  ├── frmResize         リサイズダイアログ
  ├── FolioMain         エントリポイント + BE管理
  ├── FolioData         FE側キャッシュ + テーブル操作
  ├── FolioLib          Config(Dict) + ChangeLog(ListObject) + ユーティリティ
  ├── FieldEditor       WithEvents テキストボックス
  ├── SheetWatcher      WithEvents データテーブル監視
  ├── ErrorHandler      エラー + ログ蓄積
  └── 隠しシート群      _folio_config, _folio_log, _folio_signal, _folio_mail, ...

BE: 別プロセスの Excel.Application (Visible=False)
  └── FolioWorker       スイッチ式スキャン + FEシート書き込み + リクエスト応答
```

### 通信フロー

```
BE→FE: FEの隠しシートに .Value 書き込み → Workbook_SheetChange 発火
FE→BE: BEの _folio_request シートに書き込み → SheetChange → OnTime で非同期処理
```

### スイッチ式スキャンループ

```
DoScanChunk (1秒枠)          YieldCallback
  ├ mail: manifest.tsv mtime → 時計更新
  ├ cases: root mtime        → リクエスト応答
  └ write: 変更あれば FE更新  → 次の DoScanChunk をスケジュール
```

ラウンドロビンで各タスクに公平に実行機会を与える。通常運用は全タスク μs で通過。

## モジュール一覧 (10)

| モジュール | 種別 | 責務 |
|-----------|------|------|
| FolioMain.bas | bas | エントリポイント, BE起動/停止, PID管理, FE→BE リクエスト送信 |
| FolioWorker.bas | bas | BE: スイッチ式スキャン, manifest/Dir$読み, FEシート書き込み, リクエスト処理 |
| FolioData.bas | bas | FE: 隠しシートからのキャッシュ読み込み, テーブル読み書き |
| FolioLib.bas | bas | Config(Dict), ChangeLog(ListObject), JSON, Dict, ファイルI/O |
| frmFolio.frm | frm | メインUI (左:一覧, 中:タブ詳細, 右:ログ) |
| frmSettings.frm | frm | 設定UI (パス, ソース, フィールド) |
| frmResize.frm | frm | リサイズUI (ScrollBar) |
| ErrorHandler.cls | cls | エラートレース + ログ蓄積 (エラー時に全ログ出力) |
| FieldEditor.cls | cls | WithEvents テキストボックスバインディング |
| SheetWatcher.cls | cls | WithEvents データテーブル変更監視 |

## メールアーカイブ

### フォルダ構造

```
mail_folder/
  ├── manifest.tsv        # メール一覧 (エクスポータが追記、スキャナの唯一のソース)
  └── 20260315_120000_Subject/
        ├── meta.json     # メタデータ
        ├── body.txt      # 本文
        ├── mail.msg      # 元メール
        └── attachment.pdf
```

### Outlook エクスポータ

`src/outlook/FolioMailExport.bas` + `frmMailExport.frm` を Outlook VBA にインポートして使う。folio 本体には組み込まない。エクスポート時に `manifest.tsv` へ自動追記。

## テスト

```powershell
powershell -ExecutionPolicy Bypass -File scripts/Test-Refactoring.ps1
powershell -ExecutionPolicy Bypass -File scripts/Test-Worker.ps1
```

## ディレクトリ構成

```
folio/
├── src/                      VBA ソース (10 モジュール)
│   ├── outlook/              Outlook VBA 用 (folio 本体には組み込まない)
├── scripts/
│   ├── Build-Addin.ps1       folio.xlsm 生成
│   ├── Build-Sample.ps1      サンプルデータ生成
│   ├── Test-Refactoring.ps1  スモークテスト (32項目)
│   └── Test-Worker.ps1       BEワーカーE2Eテスト
├── sample/                   サンプルデータ
├── docs/                     設計ドキュメント
├── bench/                    ベンチマークスクリプト
├── samplerun.bat             ビルド + サンプル起動
└── CLAUDE.md                 プロジェクトルール
```
