# folio

Excel VBA 案件管理アドイン。開いているワークブックのテーブル (ListObject) をリアルタイムで読み書きし、メールアーカイブ・案件フォルダと突合して一画面で管理する。

## 基本原則

- **正本は Excel テーブルそのもの** — 中間ファイルやスナップショットは持たない
- **フィールド検出はセルデータから** — `VarType` / `NumberFormat` で型判定。ハードコード禁止
- **全変更をログに記録** — ローカル編集・外部変更・メール/フォルダ変化すべてを Change Log に流す
- **設定は Excel 内部に保存** — `_folio_config` / `_folio_log` 隠しシート + JSON

## セットアップ

### 前提条件

- Windows + Excel (Microsoft 365 / 2021以降)
- Excel > ファイル > オプション > トラストセンター > マクロの設定 > **「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」を ON**

### ビルド・実行

```bat
build-addin.bat          # folio.xlsm を生成
build-sample.bat         # folio-sample.xlsx + サンプルデータを生成
samplerun.bat            # ビルド → xlsm + サンプルを開く
```

### 使い方

1. `samplerun.bat` を実行（またはビルド済み `folio.xlsm` とデータ用 `.xlsx` を開く）
2. `Alt+F8` → `Folio_ShowPanel` を実行
3. 左上のドロップダウンからテーブルを選択
4. レコード選択 → 中央タブで編集・メール確認・ファイル確認
5. 右カラムに変更ログがリアルタイム表示

### キーボードショートカット

| キー | 操作 |
|------|------|
| Ctrl+F | フィルタにフォーカス |
| Ctrl+Z | 直前の編集を元に戻す |
| F5 | メール/フォルダを再スキャン |

## アーキテクチャ

```
FE: folio.xlsm (ユーザーが操作する Excel インスタンス)
  ├── frmFolio          メインフォーム (3カラム, 全コントロール ランタイム生成)
  ├── frmSettings       設定フォーム (プロファイル・パス・フィールド設定)
  ├── FolioMain         エントリポイント + BEワーカー管理
  ├── FolioConfig       プロファイル管理 (隠しシート + JSON)
  ├── FolioData         テーブル操作 + FE側キャッシュ (隠しシートから読み込み)
  ├── FolioChangeLog    変更ログ (5000行ローテーション)
  ├── FolioHelpers      JSON パーサー, Dict, ファイル操作
  ├── FolioBundler      全コードを単一 .bas にエクスポート
  ├── FieldEditor       WithEvents で TextBox 変更検知
  ├── SheetWatcher      WithEvents でデータテーブル変更検知
  ├── ErrorHandler      エラートレース (Immediate ウィンドウ)
  └── 隠しシート群      _folio_config, _folio_log, _folio_signal, _folio_mail, ...

BE: 別プロセスの Excel.Application (Visible=False)
  ├── FolioWorker       スキャン実行 + FEシートへの書き込み
  └── FolioData         メール/案件フォルダの読み込み・キャッシュ
```

### BE→FE通信フロー

```
BE (FolioWorker)                          FE (frmFolio)
  │                                         │
  │  スキャン実行                            │
  │  FolioData.RefreshMailData()            │
  │  FolioData.RefreshCaseNames()           │
  │                                         │
  │  FEの隠しシートに直接書き込み            │
  │  g_feWb.Worksheets("_folio_mail") ─────→│ Workbook_SheetChange 発火
  │  g_feWb.Worksheets("_folio_signal") ───→│ frmFolio.OnFolioSheetChange()
  │                                         │   └→ LoadDataFromLocalSheets()
  │                                         │      (ローカルシートから O(1) 読み取り)
```

## モジュール一覧

| モジュール | 種別 | 責務 |
|-----------|------|------|
| FolioMain.bas | 標準 | エントリポイント (`Folio_ShowPanel`), BEワーカー起動/停止, ゾンビPID管理 |
| FolioConfig.bas | 標準 | プロファイル CRUD, フィールド自動検出, 設定 JSON 管理 |
| FolioData.bas | 標準 | BE: メール/案件スキャン, FE: 隠しシートからの読み込み, テーブル読み書き |
| FolioWorker.bas | 標準 | BEプロセスで実行。スキャン→FEシート書き込み→シグナル通知 |
| FolioChangeLog.bas | 標準 | `_folio_log` シートへの変更記録, ローテーション |
| FolioHelpers.bas | 標準 | JSON, Dictionary ヘルパー, ファイル I/O, 文字列操作 |
| FolioBundler.bas | 標準 | 全モジュールを単一 .bas インストーラーにエクスポート |
| ErrorHandler.cls | クラス | `Enter` / `OK` / `Catch` パターンでエラートレース |
| FieldEditor.cls | クラス | WithEvents テキストボックスバインディング, 双方向変更検知 |
| SheetWatcher.cls | クラス | WithEvents でデータテーブルの変更を監視 |
| frmFolio.frm | フォーム | メイン UI (左: 一覧, 中: タブ詳細, 右: ログ) |
| frmSettings.frm | フォーム | 設定 UI (パス, ソース, フィールド設定) |
| frmFilter.frm | フォーム | フィルタ条件設定 |
| frmDraft.frm | フォーム | メール下書き作成 |
| frmBulkDraft.frm | フォーム | 一括メール下書き作成 |
| frmResize.frm | フォーム | ウィンドウサイズ変更 |

## メールアーカイブ

folio はローカルフォルダのメールアーカイブを読み込む。各メールは以下の構造:

```
mail_folder/
  └── mail_0001/
        ├── meta.json       # メタデータ (sender, subject, received_at, ...)
        ├── body.txt        # 本文
        ├── mail.msg        # 元メール (optional)
        └── attachment.pdf  # 添付ファイル
```

### Outlook からのエクスポート

`src/FolioMailExport.bas` と `src/frmMailExport.frm` を Outlook VBA にインポートして使う（folio 本体には組み込まない）。

## 配布

`FolioBundler` で全コードを単一 `.bas` ファイルにエクスポートできる:

```
Alt+F8 → Folio_Export
```

## ディレクトリ構成

```
folio/
├── src/                      VBA ソース
│   ├── *.bas / *.cls / *.frm
│   ├── FolioMailExport.bas      Outlook VBA 用 (folio 本体には組み込まない)
│   └── frmMailExport.frm       Outlook VBA 用設定フォーム
├── scripts/
│   ├── Build-Addin.ps1       folio.xlsm 生成
│   ├── Build-Sample.ps1      folio-sample.xlsx + サンプルデータ生成
│   ├── Test-Worker.ps1       BEワーカーE2Eテスト
│   └── Test-Compile.ps1      ビルド検証
├── sample/                   サンプルデータ (git 管理)
│   ├── folio-sample.xlsx
│   ├── mail/
│   └── cases/
├── docs/
│   └── spec.md               詳細仕様書
├── build-addin.bat
├── build-sample.bat
├── samplerun.bat             ビルド + サンプル起動
├── .gitattributes            VBA ファイルの CRLF 強制
└── .gitignore                *.xlsm, *.xlam 除外
```
