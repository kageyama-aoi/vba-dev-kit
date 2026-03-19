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

---

### 3. フローシートへのボタン自動配置（#12）

#### 概要
`SetupFlowButtons` を実行すると、業務フローブックのフローシート（Config の「フローシート名」）上部に全11マクロの実行ボタンが配置される。

#### ボタンレイアウト
| 行 | ボタン（左→右） |
|---|---|
| 上段 | 処理 / 判断 / 開始-終了 ‖ 自動接続 / 縦連続接続 / 横連続接続 |
| 下段 | フォント設定 / 列範囲選択 / 接続線切替 / テキスト注入 / 連番命名 |

- `‖` = グループ間隔（図形配置系と接続系の区切り）
- ボタン幅60pt × 高さ22pt、フォント8pt（Meiryo UI）、色：青系 #4682B4
- `OnAction = "'マクロブック名'!マクロ名"` を動的に設定

#### SetupMacroBookButton
- `ThisWorkbook` の Config シートに「業務フローブックにボタンを配置」ボタンを追加
- `OnAction = "SetupFlowButtons"`
- ボタン色：グリーン系（RGB 34,139,34）
- 再実行時は既存ボタンを削除して再配置

#### 変更対象
- `vba-files/Class/ThisWorkbook.cls`
  - 追加: `SetupFlowButtons`（公開マクロ）
  - 追加: `SetupMacroBookButton`（公開マクロ）

## 非機能要件
- `図形一覧` シートは操作対象ブック（`targetWb`）内に作成する
- 書き出し後は `図形一覧` シートをアクティブにしない（フローシートのままにする）

## 技術構成
- 変更対象ファイル：`vba-files/Class/ThisWorkbook.cls`
