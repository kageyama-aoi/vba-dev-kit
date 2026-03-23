# SPEC - 技術仕様・要件定義

## 機能要件

### 1. RenameShapesSequentially の命名ルール改修（実装済み）

命名フォーマット：`FLOW_{種別}_{通し番号3桁}`

| 種別 | 対象 | 例 |
|---|---|---|
| `Process` | 矩形（msoShapeRectangle） | `FLOW_Process_001` |
| `Decision` | 菱形（msoShapeDiamond） | `FLOW_Decision_002` |
| `StartEnd` | 角丸矩形（msoShapeRoundedRectangle） | `FLOW_StartEnd_003` |
| `Arrow` | コネクタ | `FLOW_Arrow_004` |
| `Shape` | その他 | `FLOW_Shape_005` |
| 除外 | ボタン系（msoFormControl / msoOLEControlObject） | 変更しない |

---

### 2. リネーム後に図形一覧シートへ書き出し（実装済み・拡張予定）

#### 概要
`RenameShapesSequentially` 実行後、リネーム結果を専用シートに書き出す。

#### 書き出し先シート
- シート名：`図形一覧`（定数 `LIST_SHEET_NAME` で管理）
- シートが存在しない場合：新規作成
- シートが存在する場合：内容を全消去して上書き

#### 書き出し内容

| 列 | ヘッダー | 内容 | 例 |
|---|---|---|---|
| A | 図形名 | シェイプ名 | `FLOW_Process_001` |
| B | 種別 | GetShapeKind の戻値 | `Process` |
| C | テキスト | 図形内の文字列（なければ空欄） | `注文受付` |
| D | 行 | Top 座標から換算したExcel行番号 | `12` |
| E | 列 | Left 座標から換算したExcel列アルファベット | `C` |

- 並び順：Top 座標順（リネームと同じ順序）
- ボタン系は除外（リネーム対象外のため）
- コネクタはテキストなしが通常のため C 列は空欄になることが多い

---

### 3. フローシートへのボタン自動配置（実装済み）

#### ボタンレイアウト（変更後：横1行並び）
- 11個すべてを1行に横並び（幅60pt固定、はみ出しは許容）
- グループ間隔（`‖`）は3番目と4番目の間に維持
- ボタン幅60pt × 高さ22pt、フォント8pt（Meiryo UI）、文字色：黒

---

### 4. 完了アナウンス（ステータスバー表示）

以下3マクロの実行完了後、ステータスバーに処理結果を表示する。

| マクロ | 表示内容（例） |
|---|---|
| `InjectTextFromList` | `テキスト注入完了：5件` |
| `RenameShapesSequentially` | `連番命名完了：8件` |
| `SetParagraphAndFont` | `フォント設定完了：3件` |

- 表示時間：3秒後に自動消去（`Application.OnTime` で復元）
- 表示中は他のステータスバーメッセージを上書きしない（処理後に `False` に戻す）

---

### 5. 作業結果シートへの自動移動（全マクロ共通）

- 全マクロ実行後、`GetTargetSheet` で取得した対象シートをアクティブにして移動する
- 対象ブックが現在のアクティブブックと異なる場合のみ移動（同一ブックは移動不要）

---

### 6. 1行目除外（RenameShapesSequentially・SelectShapesInColumnRange）

#### RenameShapesSequentially
- Top 座標が1行目のセル下端（`ws.Rows(1).Height`）以内の図形はリネーム対象外

#### SelectShapesInColumnRange
- Top 座標が1行目のセル下端以内の図形は選択対象外

---

## 非機能要件
- `図形一覧` シートは操作対象ブック（`targetWb`）内に作成する
- 書き出し後は `図形一覧` シートをアクティブにしない（フローシートのままにする）

## 技術構成
- 変更対象ファイル：`vba-files/Class/ThisWorkbook.cls`
