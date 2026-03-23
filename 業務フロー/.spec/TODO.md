# TODO - タスクリスト

## 優先度：高

- [ ] T1: SetupFlowButtons のボタンレイアウトを横1行に変更（要望1）
- [ ] T2: InjectTextFromList・RenameShapesSequentially・SetParagraphAndFont に完了アナウンス追加（要望2）
- [ ] T3: 全マクロ共通で作業結果シートへ自動移動（要望3）
- [ ] T4: RenameShapesSequentially・SelectShapesInColumnRange で1行目の図形を除外（要望4）

## 優先度：中

## 優先度：低

## 完了済み
- [x] 初期セットアップ
- [x] バグ修正：RenameShapesSequentially でボタン図形を除外（#1）
- [x] 文字化けコメント・文字列リテラルを日本語で書き直し
- [x] UTF-8/Shift-JIS 変換スクリプト作成（vba-import.ps1 / vba-export.ps1）
- [x] RenameShapesSequentially の命名ルールをシェイプ種別対応に改修（#2）
- [x] 図形一覧シートへの書き出し機能追加（ExportShapeList）
- [x] リファクタリング：IsButtonShape / GetShapeKind / SortShapesByAxis 共通化
- [x] 図形一覧に テキスト・行・列 情報追加
- [x] 図形一覧の列順変更 / Config 化
- [x] Config.csv テンプレート作成
- [x] フローシートへのマクロ実行ボタン自動配置（SetupFlowButtons / SetupMacroBookButton）(#12)
- [x] READMEシート自動生成（SetupReadmeSheet）(#13)
- [x] コンパイルエラー修正：Font.Color.RGB → Font.Fill.ForeColor.RGB（#14）
- [x] ボタン重なり・OnActionブック名欠落修正（#15）
- [x] OnActionにThisWorkbook.プレフィックス追加（#16）
- [x] FlowStyleのシェイプ文字色を黒に設定（#18）
