# TODO - タスクリスト

## 優先度：高

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
- [x] ボタン横1行化・完了アナウンス・シート移動・1行目除外（#19）
- [x] ボタンをグループ別に色分け（緑/青/オレンジ）（#20）
- [x] リファクタリング：コード重複の解消・共通ヘルパー整理（#21）
- [x] バグ修正3件・RunDiagnostics 診断マクロ追加（#22）
