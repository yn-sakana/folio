# frmFolioV2 設計構想

## 方針
- ウィンドウなしコントロール（Label + TextBox + ScrollBar）のみで構成
- ListBox / MultiPage / ComboBox / Frame は使わない
- Z位置・リサイズ・スプリッターが完全に自由
- 既存 frmFolio はそのまま残す。V2 は別フォームとして並行開発

## エントリポイント
- `Folio_ShowPanel2` → `frmFolioV2.Show vbModeless`
- 既存の `Folio_ShowPanel` は変更しない

## コントロール構成（全て Label / TextBox / ScrollBar）

### 3カラムレイアウト
| 領域 | 旧コントロール | V2 実装 |
|------|---------------|---------|
| 左: ソース選択 | ComboBox | Label（表示）+ Label（▼ボタン）+ Label群（ドロップダウン） |
| 左: フィルタ | TextBox | TextBox（そのまま、ウィンドウなし） |
| 左: レコードリスト | ListBox | Label群 + ScrollBar（自前選択・ホバー） |
| 中央: タブバー | MultiPage | Label群（タブヘッダ）+ Label/TextBox（ページ内容切り替え） |
| 中央: Detail | TextBox群 | TextBox群（そのまま） |
| 中央: Mail リスト | ListBox | Label群 + ScrollBar |
| 中央: Files ツリー | ListBox | Label群 + ScrollBar |
| 右: ログ | ListBox | Label群 + ScrollBar |
| スプリッター | Label | Label（test-splitter.xlsm と同じ） |
| リサイズハンドル | Label | Label（test-splitter.xlsm と同じ） |
| ステータスバー | Label | Label |
| ボタン | CommandButton | Label（クリックイベント） |

### 自前リストの仕組み
- 固定数の Label 行（表示領域 ÷ 行高さ）を事前作成
- ScrollBar.Value でオフセットを管理
- データ配列[offset + i] → Label(i).Caption にバインド
- 選択行: BackColor 変更
- ホバー: MouseMove で検出、BackColor 変更

### 自前タブの仕組み
- タブヘッダ: Label 群を横並び。選択中タブは BackColor / Font.Bold で区別
- ページ内容: 全ページの Label/TextBox を重ねて配置。タブ切替時に Visible 切替

### 自前ドロップダウンの仕組み
- 閉じ状態: Label（現在値表示）+ Label（▼）
- 開き状態: Label群をフォーム上にオーバーレイ表示（全て同一レイヤーなのでZ位置問題なし）
- 選択: Label クリックで値セット → ドロップダウン非表示

## 段階的実装
1. **Phase 1**: 3カラム + スプリッター + リサイズハンドル + ステータスバー（空の箱）
2. **Phase 2**: 自前リスト（レコードリスト）+ フィルタ + ソース選択
3. **Phase 3**: タブ + Detail ページ（FieldEditor 連携）
4. **Phase 4**: Mail / Files / Log タブ
5. **Phase 5**: 既存 frmFolio を frmFolioV2 で置き換え

## 懸念事項
- Label 数が多い（リスト行数 × カラム数）→ パフォーマンス要検証
- TextBox のフォーカス管理（Tab キー移動等）
- ComboBox 自前実装のキーボード操作
