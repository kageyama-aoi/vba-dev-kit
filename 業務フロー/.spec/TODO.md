# TODO - タスクリスト

## 優先度：高
- [ ] `RenameShapesSequentially` に図形一覧シート書き出し機能を追加
  - 定数 `LIST_SHEET_NAME = "図形一覧"` を追加
  - シートが存在しない場合は新規作成、存在する場合は上書き
  - A列=図形名、B列=種別、1行目はヘッダー

## 優先度：中

## 優先度：低

## 完了済み
- [x] 初期セットアップ
- [x] バグ修正：RenameShapesSequentially でボタン図形を除外（#1）
- [x] 文字化けコメント・文字列リテラルを日本語で書き直し
- [x] UTF-8/Shift-JIS 変換スクリプト作成（vba-import.ps1 / vba-export.ps1）
- [x] RenameShapesSequentially の命名ルールをシェイプ種別対応に改修（#2）
