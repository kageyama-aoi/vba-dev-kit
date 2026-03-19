# SPEC - 技術仕様・要件定義

## 機能要件

### RenameShapesSequentially の命名ルール改修

#### 命名フォーマット
```
FLOW_{種別}_{通し番号3桁}
```

#### 種別の判定ルール

| 種別キーワード | 対象シェイプの条件 |
|---|---|
| `Process` | `msoShapeRectangle`（矩形） |
| `Decision` | `msoShapeDiamond`（菱形） |
| `StartEnd` | `msoShapeRoundedRectangle`（角丸矩形） |
| `Arrow` | `shp.Connector = True`（コネクタ） |
| `Shape` | 上記に該当しないその他の図形 |
| 除外 | `msoFormControl` / `msoOLEControlObject`（ボタン系、名前変更しない） |

#### 番号の振り方
- 種別をまたいで **全体通し番号**（Top 座標順でソート後に採番）
- 例：`FLOW_Process_001`, `FLOW_Arrow_002`, `FLOW_Decision_003`

## 非機能要件
- 既存の `InjectTextFromList` は図形名で引くため、命名ルール変更後は定義シートの A 列も新フォーマットに合わせる必要がある（マクロ側の変更は不要）
- ボタン系は前回修正通り除外継続

## 技術構成
- 変更対象ファイル：`vba-files/Class/ThisWorkbook.cls`
- 変更対象関数：`RenameShapesSequentially`
