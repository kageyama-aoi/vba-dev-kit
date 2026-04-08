## 注意事項（VBA開発ナレッジ）

以下のVBA特有の制約・注意事項を必ず守って実装してください。

---

### ■ 文字コード（UTF-8）

・VBAはUTF-8を直接読み込めないため、ファイル読み込みは ADODB.Stream を使用すること
・Charset は "UTF-8" を指定すること
・UTF-8 BOM が含まれる可能性があるため、先頭文字のBOM除去を行うこと

```vba
If Left(content, 1) = Chr(65279) Then content = Mid(content, 2)
```

---

### ■ Excel操作（パフォーマンス）

・Excelセルへのアクセスは遅いため、ループ内で Cells を直接操作しないこと
・データは一度配列に格納して処理すること
・Excelへの書き込みは最後にまとめて行うこと

---

### ■ パフォーマンス設定の定型句

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

### ■ 処理進捗の表示（StatusBar）

処理中かどうかユーザーが判断できるよう StatusBar を活用すること。

```vba
Application.StatusBar = "処理中... テーブル X / Y"
```

処理完了後は必ず `Application.StatusBar = False` でリセットすること。

---

### ■ エラー発生箇所の特定（errContext）

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

### ■ 変数宣言の位置

・`Dim` はすべてサブルーチン・関数の**先頭**にまとめること
・ループ内・条件分岐内の `Dim` は VBA では関数スコープになるが、
  意図が伝わらず混乱するため禁止

---

### ■ 固定サイズ配列と動的配列

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

### ■ UsedRangeの使用禁止

・UsedRangeは不正確になる場合があるため使用しないこと
・最終行の取得は以下の方法を使用すること

```vba
lastRow = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
```

---

### ■ Variant配列の扱い

・Range.Valueで取得したデータは2次元配列になる
・配列は1始まり（1-based）である
・ループは `UBound(data, 1)` を使用すること
・配列アクセスは `data(i, 1)` のように2次元で行うこと
・単一セルの場合は配列にならないため `IsArray` で判定すること

---

### ■ 配列書き込み

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

### ■ ByRefエラー対策

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

### ■ モジュール変数のスコープ

・モジュールレベルの変数は `Private` を明示すること（`Dim` のみは暗黙 Public 相当）
・複数箇所で使う文字列定数は `Private Const` にまとめること

```vba
Private Const OUTPUT_SHEET_NAME As String = "テーブル定義"
Private IMAGE_SCALE As Double
```
