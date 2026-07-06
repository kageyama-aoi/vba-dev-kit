## 注意事項（VBA開発ナレッジ）

以下のVBA特有の制約・注意事項を必ず守って実装してください。

> **読むタイミング：** 新規モジュール作成時は **A** → Excelを操作するときは **B** → 挙動がおかしいときは **C** → マクロブック自身とは別のブックを操作する構成なら **D** も

---

### 目次（一覧）

#### A：モジュール骨格（新規作成時に確認）

| # | 項目 | 一言メモ |
|---|------|---------|
| A-1 | モジュールレベル宣言の位置 | `Private Const` / `Dim` は最初の Sub より前に書く。関数間はコンパイルエラーになる |
| A-2 | モジュール変数のスコープ | モジュールレベルは Private 明示。定数は Private Const に |
| A-3 | 変数宣言の位置 | Dim はすべて関数先頭にまとめる |
| A-4 | パフォーマンス設定の定型句 | ScreenUpdating / Calculation / EnableEvents を Cleanup とセットで |
| A-5 | エラー発生箇所の特定（errContext） | フェーズ名をメッセージに含める |

#### B：Excel操作パターン（セル・シートを触るとき）

| # | 項目 | 一言メモ |
|---|------|---------|
| B-1 | Excel操作（パフォーマンス） | ループ内で Cells を直接触らない。配列経由で |
| B-2 | UsedRangeの使用禁止 | 最終行は `End(xlUp).Row` で取る |
| B-3 | Variant配列の扱い | Range.Value は2次元・1始まり。単一セルは IsArray で判定 |
| B-4 | 配列書き込み | Resize.Value = 配列 で一括書き込み |
| B-5 | 処理進捗の表示（StatusBar） | 長い処理は StatusBar で状況を見せる |
| B-6 | 空行で途中終了するループ | 「空になったら終了」は途中の空行で止まる。最終行を先に取って For で回す |
| B-7 | ボタン系図形の除外 | `ws.Shapes` ループはフォーム/ActiveXコントロールも含む。Type チェックで除外 |
| B-8 | ループ内 ReDim Preserve の性能 | 都度拡張は O(n²)。件数既知なら1回で ReDim、不明で大量なら Collection |
| B-9 | フォント色プロパティの混同 | TextFrame2 配下は `.Font.Fill.ForeColor.RGB`。`Font.Color.RGB` はエラー |

#### C：コーディング落とし穴（VBA固有の罠）

| # | 項目 | 一言メモ |
|---|------|---------|
| C-1 | 文字コード（UTF-8） | 読み込みは ADODB.Stream。BOM除去も忘れずに |
| C-2 | 固定サイズ配列と動的配列 | ReDim Preserve したいなら `Dim arr()` で動的宣言 |
| C-3 | And / Or は短絡評価しない | 配列境界チェックと配列アクセスは条件を分けて書く |
| C-4 | ByRefエラー対策 | Variant要素を String 引数に渡す時は CStr() か ByVal |
| C-5 | Dictionaryへのオブジェクト格納 | Collection等を Dictionary に入れるときは必ず `Set` を付ける |
| C-6 | `Private Type` は Public Sub の引数型に使用不可 | 標準モジュールでも `Private Type` を Public Sub の引数にすると「Sub または Function が定義されていません」になる。`Public Type` にすること |
| C-7 | ループ内のオブジェクト初期化忘れ | `On Error Resume Next` で取得する前に `Set obj = Nothing`。前の反復の値が残って誤判定する |
| C-8 | `On Error Resume Next` の範囲が広すぎる | 有効範囲は1〜2行に絞る。`Err.Number = 0` 判定は前後に `Err.Clear` |
| C-9 | 数値変換前の型チェック欠如 | `CLng` / `CInt` の前に キャンセル → `IsNumeric` → 範囲 の3段階チェック |
| C-10 | `Optional` キーワードの付け忘れ | 複数箇所から呼ぶ Sub の引数は Optional + デフォルト値を早めに検討 |
| C-11 | `Scripting.Dictionary` の生成忘れ | `Object` 宣言だけでは Nothing。`Set dict = CreateObject(...)` を宣言直後に |
| C-12 | `OnAction` / `OnTime` のマクロパス書式 | `'ブック名'!モジュール名.Sub名` が基本形。クラスモジュールはモジュール名必須 |

#### D：外部ブック操作マクロ（条件付き）

> **適用条件：** マクロブック自身とは別のブック（業務ファイル等）を操作対象とする設計の場合のみ。
> マクロが埋め込まれたブック自身を操作するだけの単体マクロには適用不要。

| # | 項目 | 一言メモ |
|---|------|---------|
| D-1 | `ActiveSheet` の直接使用禁止 | 呼び出し元ブックのシートを指して「システムエラー &H80004005」になる |
| D-2 | 新規 Public Sub の冒頭テンプレート | 設定読み込み → シート取得 → `Is Nothing` チェックの3点セット |
| D-3 | ボタン除外＋ボタン配置行の除外 | 種別除外だけでは不十分。配置行の高さでも除外する |
| D-4 | `OnAction` のクラスモジュール プレフィックス | モジュール名なしだとボタンが反応しない |
| D-5 | `Application.OnTime` のプレフィックス | モジュール名なしだとサイレントに失敗し StatusBar が残る |
| D-6 | 図形名の重複 | Dictionary キーに図形名を使うと同名は後勝ち上書き。エラーにならず気づきにくい |
| D-7 | 定義シートの空行による途中終了 | 外部ブックの定義シートは空行を想定して最終行方式で読む |

---

## セクション A：モジュール骨格

### ■ A-1 モジュールレベル宣言の位置

・`Private Const` / `Dim` などのモジュールレベル宣言は、**最初の Sub / Function より前**にすべてまとめること
・Sub / Function の間に挟むと「End Sub, End Function または End Property 以降には、コメントのみが記述できます」コンパイルエラーになる
・既存モジュールに定数を追加するときは、ファイル先頭の定数ブロックに追記すること

```vba
' NG：関数の間に Private Const を置くとコンパイルエラー
End Function

Private Const NEW_CONST As String = "値"   ' ← ここに置くと NG

Sub NextSub()

' OK：先頭の定数ブロックにまとめる
Private Const EXISTING_CONST As String = "既存"
Private Const NEW_CONST       As String = "値"   ' ← 先頭ブロックに追加

Function FirstFunc()
```

【実例】WriteRunLog 追加時に `Private Const RUN_LOG_SHEET_NAME` を `ResolveFilePath` の直後に置いたため、
  コンパイルエラーが発生した。定数を先頭の定数ブロックへ移動して解消。

---

### ■ A-2 モジュール変数のスコープ

・モジュールレベルの変数は `Private` を明示すること（`Dim` のみは暗黙 Public 相当）
・複数箇所で使う文字列定数は `Private Const` にまとめること

```vba
Private Const OUTPUT_SHEET_NAME As String = "テーブル定義"
Private IMAGE_SCALE As Double
```

---

### ■ A-3 変数宣言の位置

・`Dim` はすべてサブルーチン・関数の**先頭**にまとめること
・ループ内・条件分岐内の `Dim` は VBA では関数スコープになるが、
  意図が伝わらず混乱するため禁止

---

### ■ A-4 パフォーマンス設定の定型句

Sub の先頭と Cleanup に必ず入れること。
エラーが発生しても必ず復元されるよう `On Error GoTo Cleanup` とセットで使うこと。

```vba
' --- 先頭 ---
Application.ScreenUpdating = False
Application.Calculation    = xlCalculationManual
Application.EnableEvents   = False

' --- Cleanup ラベル内 ---
Application.ScreenUpdating = True
Application.Calculation    = xlCalculationAutomatic
Application.EnableEvents   = True
Application.StatusBar      = False
```

---

### ■ A-5 エラー発生箇所の特定（errContext）

大きな処理ではフェーズ名を `errContext` 変数で管理し、エラーメッセージに含めること。
「エラーが発生しました」だけでは原因特定に時間がかかる。

```vba
Dim errContext As String
errContext = "初期化"

errContext = "データ取得"
' ...

Cleanup:
If Err.Number <> 0 Then
    MsgBox "エラー（" & errContext & "）: " & Err.Description, vbExclamation
End If
```

---

## セクション B：Excel操作パターン

### ■ B-1 Excel操作（パフォーマンス）

・Excelセルへのアクセスは遅いため、ループ内で Cells を直接操作しないこと
・データは一度配列に格納して処理すること
・Excelへの書き込みは最後にまとめて行うこと

---

### ■ B-2 UsedRangeの使用禁止

・UsedRangeは不正確になる場合があるため使用しないこと
・最終行の取得は以下の方法を使用すること

```vba
lastRow = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
```

---

### ■ B-3 Variant配列の扱い

・Range.Valueで取得したデータは2次元配列になる
・配列は1始まり（1-based）である
・ループは `UBound(data, 1)` を使用すること
・配列アクセスは `data(i, 1)` のように2次元で行うこと
・単一セルの場合は配列にならないため `IsArray` で判定すること

---

### ■ B-4 配列書き込み

・セルに1件ずつ書き込まず、配列をそのままRangeに代入すること

```vba
' NG（行数分のCOM呼び出しが発生）
For i = 1 To 1000
    ws.Cells(i, 1).Value = data(i)
Next i

' OK（1回のCOM呼び出しで完了）
ws.Range("A1").Resize(1000, 1).Value = data
```

---

### ■ B-5 処理進捗の表示（StatusBar）

処理中かどうかユーザーが判断できるよう StatusBar を活用すること。

```vba
Application.StatusBar = "処理中... テーブル X / Y"
```

処理完了後は必ず `Application.StatusBar = False` でリセットすること。

---

### ■ B-6 空行で途中終了するループ

・「A列が空になったら終了」のループは、途中に空行があると後半データが読まれない
・ユーザーが直接編集するシートを読む際は空行の混入を想定すること
・`End(xlUp).Row` で最終行を取得してから `For` ループが安全

```vba
' NG：途中に空行があると止まる
Do While ws.Cells(i, 1).Value <> ""

' OK：最終行を先に取得してから For ループ
lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
For i = 2 To lastRow
    If ws.Cells(i, 1).Value = "" Then GoTo NextRow
    ' 処理
NextRow:
Next i
```

---

### ■ B-7 ボタン系図形の除外

・`ws.Shapes` をループするとフォームコントロールや ActiveX コントロールも含まれる
・名前変更・削除・エクスポートでボタンを巻き込む（特に削除は消したボタンの再配置が手間）
・図形を操作するループには必ずボタン系の除外条件を入れること

```vba
' NG：ボタンも対象になる
For Each shp In ws.Shapes
    shp.Name = "FLOW_" & i

' OK：Type チェックで除外
For Each shp In ws.Shapes
    If shp.Type = msoFormControl Or shp.Type = msoOLEControlObject Then GoTo Skip
    shp.Name = "FLOW_" & i
Skip:
Next shp
```

---

### ■ B-8 ループ内 `ReDim Preserve` の性能

・件数不明の配列を `ReDim Preserve` でループ都度拡張すると、毎回全体コピーが走り O(n²) で著しく遅くなる
・件数が事前にわかる場合は1回だけ `ReDim` する
・数百件を超える可能性があるなら `Collection` や `Dictionary` を検討する（小規模＝数十件なら実害なし）

```vba
' 動くが遅い
For Each shp In ws.Shapes
    count = count + 1
    ReDim Preserve arr(1 To count)  ' 毎回全体をコピー
Next shp

' OK：件数が事前にわかる場合は1回だけ ReDim
ReDim arr(1 To ws.Shapes.Count)

' OK：件数不明で大量になりうる場合は Collection
Dim col As New Collection
For Each shp In ws.Shapes : col.Add shp : Next shp
```

---

### ■ B-9 フォント色プロパティの混同

・TextFrame2 配下では `Font.Color.RGB` は使えない（AI が頻繁に混同するパターン）
・AutoShape の TextFrame2 でフォント色 → `.Font.Fill.ForeColor.RGB`
・フォームコントロールのキャプション文字色 → `.TextFrame.Characters.Font.Color`（別系統）

```vba
' NG：TextFrame2 配下で Font.Color.RGB はエラーになる
shp.TextFrame2.TextRange.Font.Color.RGB = RGB(0, 0, 0)

' OK：TextFrame2 配下でのフォント色
shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)

' OK：フォームコントロール（ボタン）の文字色は別系統
btn.TextFrame.Characters.Font.Color = RGB(255, 255, 255)
```

---

## セクション C：コーディング落とし穴

### ■ C-1 文字コード（UTF-8）

・VBAはUTF-8を直接読み込めないため、ファイル読み込みは ADODB.Stream を使用すること
・Charset は "UTF-8" を指定すること
・UTF-8 BOM が含まれる可能性があるため、先頭文字のBOM除去を行うこと

```vba
If Left(content, 1) = Chr(65279) Then content = Mid(content, 2)
```

---

### ■ C-2 固定サイズ配列と動的配列

・`ReDim Preserve` を使う配列は `Dim arr()` で動的宣言すること
・`Dim arr(99)` と宣言した後の `ReDim` はコンパイルエラーになる

```vba
' NG
Dim arr(99) As String
ReDim Preserve arr(0 To colCount - 1)  ' → コンパイルエラー

' OK
Dim arr() As String
ReDim arr(0 To 99)
ReDim Preserve arr(0 To colCount - 1)  ' → OK
```

---

### ■ C-3 And / Or は短絡評価しない（配列境界外アクセスに注意）

・VBAの `And` / `Or` は**両辺を必ず評価する**（C言語の `&&` / `||` とは異なる）
・`Do While i < rowCount And arr(i) = val` のような書き方は、
  `i = rowCount`（配列の上限外）のときでも `arr(i)` を評価して「インデックスが有効範囲にありません」エラーになる
・配列アクセスを含む条件は必ず分けて書くこと

```vba
' NG：i が UBound を超えると arr(i) でエラー
Do While i < rowCount And arr(i) = tName
    i = i + 1
Loop

' OK：先に範囲チェックしてから配列アクセス
Do While i < rowCount
    If arr(i) <> tName Then Exit Do
    i = i + 1
Loop
```

【実例】v1.3 で ReDim Preserve による配列縮小を追加したことで顕在化。
  それ以前は配列が MAX_ROWS サイズのままだったため、範囲外が空文字を返して
  偶然ループを抜けていた。

---

### ■ C-4 ByRefエラー対策

・VBAの引数はデフォルトで ByRef であるため、型が一致しないとエラーになる
・Variant配列の要素（`arr(i)` など）を String 引数に渡す場合は `CStr()` で変換すること
・または引数に `ByVal` を付けて自動変換させること

```vba
' NG（Variant要素をByRef Stringに渡すとエラー）
Call MyFunc(varArr(i))

' OK パターン1：CStr で明示変換
Call MyFunc(CStr(varArr(i)))

' OK パターン2：ByVal で受ける
Sub MyFunc(ByVal colName As String)
```

---

### ■ C-5 Dictionaryへのオブジェクト格納（Set の省略禁止）

・Scripting.Dictionary にオブジェクト（Collection・Worksheet 等）を格納するときは必ず `Set` を付けること
・`Set` なしで代入すると、VBA はオブジェクトの**既定プロパティの値**を格納しようとする
・`Collection` の既定プロパティは `Item`（要引数）のため、引数なしで呼ばれて **Error 450「引数の数が一致していません」** が発生する
・エラーが `On Error GoTo Cleanup` に捕まるため「なぜそのフェーズで引数エラー？」と原因特定が難しくなる

```vba
' NG：Collection の既定プロパティ（Item）を引数なしで呼ぼうとして Error 450
If Not dict.Exists(key) Then dict(key) = New Collection

' OK：Set でオブジェクト参照を格納する
If Not dict.Exists(key) Then Set dict(key) = New Collection
```

【実例】v1.6 で indexDetails（Scripting.Dictionary）に Collection を格納する処理を追加した際、
  3箇所すべてで `Set` が漏れ、ALTER TABLE 行のパース中に Error 450 が発生した。

---

### ■ C-6 `Private Type` は Public Sub の引数型に使用不可

・VBA 標準モジュール（.bas）で `Private Type` を定義しても、`Public Sub / Function` の引数型・戻り値型には使えない
・コンパイルが通らず、その Sub/Function を呼び出している行で **「Sub または Function が定義されていません」** エラーになる
・同じモジュール内の呼び出しでも発生するため、原因が分かりにくい
・`Public Type` にすれば解消する（他モジュールから参照されても問題ない）

```vba
' NG：Private Type を Public Sub の引数に使うとコンパイルエラー
Private Type ColRow
    tblName As String
    colName As String
End Type

Public Sub WriteDefinitions(rows() As ColRow)  ' ← コンパイル不可
    ...
End Sub

' OK：Public Type にする
Public Type ColRow
    tblName As String
    colName As String
End Type

Public Sub WriteDefinitions(rows() As ColRow)  ' ← OK
    ...
End Sub
```

【実例】R-3 で ColRow 型を導入した際に `Private Type` にしたため、
  `ParseSqlFile` / `ApplyKeyInfo` 等すべての Public Sub がコンパイル不可となり、
  呼び出し元の `UBound(keyCols)` を含む行で「Sub または Function が定義されていません」が発生した。

---

### ■ C-7 ループ内のオブジェクト初期化忘れ

・`On Error Resume Next` でオブジェクト取得に失敗しても、前のループ反復の値が変数に残り、意図せず処理が通ってしまう
・`On Error Resume Next` でオブジェクトを取得する箇所は直前に `Set obj = Nothing` を入れること
・ループ終了後にも `Set obj = Nothing`（後続処理への値の持ち越し防止）

```vba
' NG：shp が前の反復の値を保持したまま次の判定に入る
Set shp = ws.Shapes(shpName)   ' 見つからなくてもエラーが無視される
If Not shp Is Nothing Then     ' 前の shp が残っていて誤って True になる！

' OK：毎反復の頭で必ず Nothing リセット
Set shp = Nothing
On Error Resume Next
Set shp = ws.Shapes(shpName)
On Error GoTo 0
If Not shp Is Nothing Then ...
```

---

### ■ C-8 `On Error Resume Next` の範囲が広すぎる

・`On Error Resume Next` を有効にしたまま複数の処理を書くと、別の処理のエラーまで無視する
・有効範囲は最小限（1〜2行）に絞り、すぐ `On Error GoTo 0` に戻すこと
・`Err.Number = 0` で成否を判定するパターンは、直前に `Err.Clear`・直後にも `Err.Clear` がないと前のエラーの残滓で誤判定する

```vba
' NG：Err.Number が前の操作の残滓を持っている
On Error Resume Next
Set wsDef = targetWb.Worksheets("SheetName")
Call SomeOtherFunction()          ' ここのエラーも無視される
If Err.Number = 0 Then ...        ' 前のエラーが残っていて誤判定！
On Error GoTo 0

' OK：範囲を最小限に絞り、Err.Clear を適切に挟む
On Error Resume Next
Set wsDef = targetWb.Worksheets("SheetName")
On Error GoTo 0

On Error Resume Next
Err.Clear
srcName = shp.ConnectorFormat.BeginConnectedShape.Name
If Err.Number = 0 And srcName <> "" Then ...
Err.Clear
On Error GoTo 0
```

---

### ■ C-9 数値変換前の型チェック欠如

・`InputBox` の戻り値を直接 `CLng` / `CInt` / `CDbl` に渡すと、非数値・キャンセル入力でクラッシュする
・キャンセル（空文字）チェック → `IsNumeric` → 範囲チェック（下限・上限）の3段階で確認すること

```vba
' NG："abc" や空文字でクラッシュ
idx = CLng(InputBox("番号を入力してください"))

' OK：3段階チェック（キャンセル → 数値か → 範囲内か）
If ans = "" Then Exit Sub
If Not IsNumeric(ans) Then MsgBox "半角数字で入力してください。" : Exit Sub
idx = CLng(ans)
If idx < 1 Or idx > maxCount Then MsgBox "不正な番号です。" : Exit Sub
```

---

### ■ C-10 `Optional` キーワードの付け忘れ

・引数を省略して呼びたい場面が後から生じたとき、`Optional` がなければコンパイルエラーになる
・複数箇所から呼ばれる Sub/Function の引数は `Optional` を早めに検討すること
・`Optional` 引数には必ずデフォルト値を指定すること
・「引数が渡されたか」を判定したい場合は `Variant` 型にして `IsMissing(引数名)` を使う

```vba
' NG：省略して呼べない
Sub MySub(shp As Shape, txt As String)

' OK：Optional にして省略可能にする
Sub MySub(shp As Shape, Optional txt As String = "")
```

---

### ■ C-11 `Scripting.Dictionary` の生成忘れ

・`Object` 型で宣言しただけでは `Nothing` のまま。`Set` + `CreateObject` を忘れると実行時エラーになる
・`Set dict = CreateObject("Scripting.Dictionary")` を宣言直後に書くこと
・アクセス前に `dict.Exists(key)` チェックを入れること
・オブジェクトを値に格納する場合は `Set dict(key) = obj`（→ C-5 参照）

```vba
' NG：CreateObject を忘れる
Dim dict As Object
dict("key") = value   ' エラー

' OK：必ず Set + CreateObject で生成
Dim dict As Object
Set dict = CreateObject("Scripting.Dictionary")
dict("key") = value
```

---

### ■ C-12 `OnAction` / `Application.OnTime` のマクロパス書式

・マクロパスの書式を間違えるとボタンが反応しない・OnTime が発火しない
・`'ブック名'!モジュール名.Sub名` が基本形。ブック名はシングルクォートで囲む
・ブック名は `ThisWorkbook.Name` で動的に取得する（ハードコードしない）
・クラスモジュール内の Sub は必ず `モジュール名.` プレフィックスが必要

```vba
' NG：ブック名なし
btn.OnAction = "MyMacro"

' NG：モジュール名なし（クラスモジュール内の Sub は参照不可）
btn.OnAction = "'" & ThisWorkbook.Name & "'!MyMacro"

' OK：標準モジュール
btn.OnAction = "'" & ThisWorkbook.Name & "'!Module1.MyMacro"

' OK：クラスモジュール（ThisWorkbook など）
btn.OnAction = "'" & ThisWorkbook.Name & "'!ThisWorkbook.MyMacro"

' Application.OnTime も同様
Application.OnTime Now + TimeValue("00:00:03"), _
    "'" & ThisWorkbook.Name & "'!ThisWorkbook.ClearStatusBar"
```

---

## セクション D：外部ブック操作マクロ（条件付き）

> **適用条件：** レビュー・実装対象のマクロが「マクロブック自身とは別のブック（業務ファイル等）を操作対象とする設計」になっている場合のみ適用する。
> 具体的には、マクロブックからボタンやスクリプトで外部ブックのシートを操作する構成。
> マクロが埋め込まれたブック自身を操作するだけの単体マクロには適用不要。

### ■ D-1 `ActiveSheet` の直接使用禁止

・外部ブックからマクロを呼び出す場合、`ActiveSheet` は呼び出し元ブックのシートを指してしまい「システムエラー &H80004005」が発生する
・`ActiveSheet` / `ActiveWorkbook` の直接参照は禁止
・操作対象ブックを特定するヘルパー関数（例：`GetTargetSheet`）経由でシートを取得すること

```vba
' NG：外部ブックから呼び出すと ActiveSheet がマクロブック側を指す
Set ws = ActiveSheet

' OK：操作対象ブックを明示的に解決して取得する
Set targetWb = （対象ブックを特定するロジック）
Set ws = targetWb.Worksheets("フローシート")
```

---

### ■ D-2 新規 Public Sub の冒頭テンプレート

・新しい Public Sub を追加するときは、設定読み込み → シート取得 → `Is Nothing` チェックの3点セットを冒頭に入れること

```vba
' NG：設定読み込みとシート取得・エラーチェックが抜けている
Sub NewMacro()
    ws.Shapes.AddShape ...   ' ws が Nothing でクラッシュ

' OK：3点セット
Sub NewMacro()
    ' (1) 設定を読み込む（Config シート等から）
    If Not LoadConfig() Then Exit Sub
    ' (2) 操作対象シートを解決する
    Set ws = GetTargetSheet(m_FlowSheet)
    ' (3) 取得失敗時は即終了
    If ws Is Nothing Then Exit Sub
```

---

### ■ D-3 ボタン除外 ＋ ボタン配置行の除外

・シートの特定行にボタンを横並び配置している場合、図形ループでは「ボタン種別除外」（B-7）と「ボタン行除外」の両方が必要

```vba
' NG：種別除外だけでは不十分
If Not IsButtonShape(shp) Then ...

' OK：種別除外 + ボタン配置行（例：1行目）の高さで除外を組み合わせる
row1H = ws.Rows(1).Height
If Not IsButtonShape(shp) And shp.Top >= row1H Then ...
```

---

### ■ D-4 `OnAction` のクラスモジュール プレフィックス

・マクロをクラスモジュール（`ThisWorkbook` や `Sheet1`）内に定義している場合、`OnAction` に `モジュール名.` プレフィックスがないとボタンが反応しない（書式の基本形は C-12 参照）

```vba
' NG：モジュール名なし
btn.OnAction = "'" & macroBook & "'!Flow_Process"

' OK：クラスモジュールの Sub は モジュール名. が必要
btn.OnAction = "'" & macroBook & "'!ThisWorkbook.Flow_Process"
```

---

### ■ D-5 `Application.OnTime` のクラスモジュール プレフィックス

・`OnAction` と同様、モジュール名なしだと**サイレントに失敗**し StatusBar が残り続ける

```vba
' NG：モジュール名なし（サイレントに失敗し StatusBar が残り続ける）
Application.OnTime Now + TimeValue("00:00:03"), _
    "'" & ThisWorkbook.Name & "'!ClearStatusBar"

' OK：クラスモジュール内の Sub は モジュール名. が必要
Application.OnTime Now + TimeValue("00:00:03"), _
    "'" & ThisWorkbook.Name & "'!ThisWorkbook.ClearStatusBar"
```

---

### ■ D-6 図形名の重複

・Dictionary でキーに図形名を使う場合、定義データに同名が2件あると後勝ちで上書きされ、接続先が変わる。エラーにならず気づきにくい
・図形名（キー）に重複がないかコード上でチェックするか、ドキュメントで禁止すること
・接続先参照で `dict.Exists(key)` が False の場合に処理をスキップする実装は「エラーなしで無視」になるため、その旨をコメントまたはログに残すこと

---

### ■ D-7 定義シートの空行による途中終了

・外部ブックから読み込む定義シートに空行があるとループが途中終了する（B-6 と同じ対策：最終行を先に取得してから For ループ）

```vba
' NG：途中に空行があると止まる
Do While wsDef.Cells(rowIdx, 1).Value <> ""

' OK：最終行を先に取得してから For ループにする
lastRow = wsDef.Cells(wsDef.Rows.Count, 1).End(xlUp).Row
For rowIdx = 2 To lastRow
    If wsDef.Cells(rowIdx, 1).Value = "" Then GoTo NextRow
NextRow:
Next rowIdx
```
