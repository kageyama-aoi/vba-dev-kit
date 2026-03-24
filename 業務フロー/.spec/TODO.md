# TODO - タスクリスト

## 優先度：高

### GenerateFlowFromSheet（フロー自動生成）
- [ ] T1: 定義シートの読み込み処理（A〜E列、ヘッダースキップ、空行終了）
- [ ] T2: 既存図形の確認ロジック（ボタン・1行目除外で件数カウント → 警告ダイアログ）
- [ ] T3: 図形配置ループ（種別→msoShape変換、縦一列自動座標計算、FlowStyle適用）
- [ ] T4: 接続線の生成（YES/NO接続先を走査、ConnectTwoShapes 呼び出し、Auto判定）
- [ ] T5: ステータスバー完了アナウンス（図形数・接続線数を表示）
- [ ] T6: SetupFlowButtons にボタン追加（「フロー生成」）

### ExportFlowData（フロー情報エクスポート）
- [ ] T7: コネクタ走査で接続関係を Dictionary にマッピング（接続元→接続先1/2）
- [ ] T8: ExportShapeList を拡張（F列：接続先1、G列：接続先2 を追加）
- [ ] T9: CSV 書き出し処理（UTF-8 BOM付き、ファイル名に実行時刻、targetWb のフォルダに保存）
- [ ] T10: 完了 MsgBox（件数 ＋ 保存先パスを表示）
- [ ] T11: SetupFlowButtons にボタン追加（「ﾌﾛｰ出力」）

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
