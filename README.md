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
run.bat                  # ビルド → xlsm + サンプルを開く
```

### 使い方

1. `run.bat` を実行（またはビルド済み `folio.xlsm` とデータ用 `.xlsx` を開く）
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
データワークブック (*.xlsx)
  └── テーブル (ListObject)
        ↕  直接読み書き
folio アドイン (*.xlsm)
  ├── frmFolio          メインフォーム (3カラム, 全コントロール ランタイム生成)
  ├── frmSettings       設定フォーム (プロファイル・パス・フィールド設定)
  ├── FieldEditor       WithEvents で TextBox 変更検知
  ├── FolioMain         エントリポイント + Application.OnTime ポーリング
  ├── FolioConfig       プロファイル管理 (隠しシート + JSON)
  ├── FolioData         テーブル・メール・フォルダ I/O
  ├── FolioChangeLog    変更ログ (5000行ローテーション)
  ├── FolioHelpers      JSON パーサー, Dict, ファイル操作
  ├── FolioBundler      全コードを単一 .bas にエクスポート
  └── ErrorHandler      エラートレース (Immediate ウィンドウ)
```

## モジュール一覧

| モジュール | 種別 | 責務 |
|-----------|------|------|
| FolioMain.bas | 標準 | エントリポイント (`Folio_ShowPanel`), ポーリングタイマー |
| FolioConfig.bas | 標準 | プロファイル CRUD, フィールド自動検出, 設定 JSON 管理 |
| FolioData.bas | 標準 | テーブル読み書き, メールアーカイブ読み込み, 案件フォルダ走査 |
| FolioChangeLog.bas | 標準 | `_folio_log` シートへの変更記録, ローテーション |
| FolioHelpers.bas | 標準 | JSON, Dictionary ヘルパー, ファイル I/O, 文字列操作 |
| FolioBundler.bas | 標準 | 全モジュールを単一 .bas インストーラーにエクスポート |
| FolioSampleBuilder.bas | 標準 | サンプルデータ (テーブル・メール・フォルダ) をゼロから生成 |
| ErrorHandler.cls | クラス | `Enter` / `OK` / `Catch` パターンでエラートレース |
| FieldEditor.cls | クラス | WithEvents テキストボックスバインディング, 双方向変更検知 |
| frmFolio.frm | フォーム | メイン UI (左: 一覧, 中: タブ詳細, 右: ログ) |
| frmSettings.frm | フォーム | 設定 UI (パス, ソース, フィールド設定) |

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

```
1. Outlook で Alt+F11 → FolioMailExport.bas, frmMailExport.frm をインポート
2. FolioMail_Setup で出力先・アカウント・フォルダ・期間を設定
3. FolioMail_Run で手動エクスポート、または自動エクスポート (ThisOutlookSession)
4. 出力先を folio の mail_folder 設定に指定
```

差分エクスポート対応（meta.json 存在チェック）。設定は `%APPDATA%\FolioMailExport\.foliomail.json` に保存。

## 配布

`FolioBundler` で全コードを単一 `.bas` ファイルにエクスポートできる:

```
Alt+F8 → Folio_Export
```

生成された `.bas` を配布先の Excel にインポートし、`Install_Folio` マクロを実行すると、モジュール・隠しシートが自動作成される。

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
│   └── Test-Compile.ps1      ビルド検証
├── sample/                   サンプルデータ (git 管理)
│   ├── folio-sample.xlsx
│   ├── mail/
│   └── cases/
├── docs/
│   └── spec.md               詳細仕様書
├── build-addin.bat
├── build-sample.bat
├── run.bat                   ビルド + 起動
├── .gitattributes            VBA ファイルの CRLF 強制
└── .gitignore                *.xlsm, *.xlam 除外
```
