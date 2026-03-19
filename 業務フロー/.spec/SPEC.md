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

#### 変更対象
- `vba-files/Class/ThisWorkbook.cls`
- 定数追加：`LIST_SHEET_NAME = "図形一覧"`
- 関数変更：`RenameShapesSequentially`（書き出しロジックを末尾に追加）

## 非機能要件
- `図形一覧` シートは操作対象ブック（`targetWb`）内に作成する
- 書き出し後は `図形一覧` シートをアクティブにしない（フローシートのままにする）

## 技術構成
- 変更対象ファイル：`vba-files/Class/ThisWorkbook.cls`
