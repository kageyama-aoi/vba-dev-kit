# AI生成 VBAコード セルフチェックノウハウ資料【汎用版】

> 作成日: 2026-03-24
> 対象: どんな VBA マクロプロジェクトにも使える共通チェック項目

---

## はじめに：なぜ AIのコードはバグが出やすいか

AIは「それっぽく動くコード」を生成するのが得意ですが、以下の点で人間よりも見落としがちです：

- Excel/VBA の**実行環境の特殊性**（型変換の暗黙挙動、エラー処理の罠など）
- **プロパティ名の微妙な違い**（TextFrame vs TextFrame2 など）
- **引数設計の見落とし**（Optional / IsMissing の使い忘れ）

---

## チェックリスト早見表

| # | カテゴリ | チェック項目 |
|---|---------|------------|
| 1 | オブジェクト初期化 | ループ内で `Set obj = Nothing` してから取得しているか |
| 2 | エラー処理 | `On Error Resume Next` の範囲と `Err.Clear` は適切か |
| 3 | 数値変換 | `CLng` / `CInt` の前に `IsNumeric` チェックがあるか |
| 4 | フォント色設定 | `Font.Color.RGB` と `Font.Fill.ForeColor.RGB` を混同していないか |
| 5 | Optional 引数 | 省略可能な引数に `Optional` キーワードがついているか |
| 6 | ReDim Preserve | 件数不明の配列を `ReDim Preserve` でループ都度拡張しているか |
| 7 | Dictionary 生成 | `Set obj = CreateObject(...)` を忘れていないか |
| 8 | 空行によるループ終了 | `Do While .Cells(i,1) <> ""` 方式で途中の空行に対応できているか |
| 9 | ボタン系図形の除外 | 図形ループで `msoFormControl` / `msoOLEControlObject` を除外しているか |
| 10 | OnAction / OnTime パス | マクロ名が `'ブック名'!モジュール名.Sub名` の形式になっているか |

---

## 詳細解説

---

### ✅ チェック 1：ループ内のオブジェクト初期化忘れ

**バグパターン：** `On Error Resume Next` でオブジェクト取得に失敗しても、前のループ反復の値が変数に残り、意図せず処理が通ってしまう

```vba
' ❌ バグあり：shp が前の反復の値を保持したまま次の判定に入る
Do While wsDef.Cells(i, 1).Value <> ""
    shpName = Trim(wsDef.Cells(i, 1).Value)
    On Error Resume Next
    Set shp = ws.Shapes(shpName)   ' 見つからなくてもエラーが無視される
    On Error GoTo 0
    If Not shp Is Nothing Then     ' 前の shp が残っていて誤って True になる！
        Call DoSomething(shp)
    End If
    i = i + 1
Loop
```

```vba
' ✅ 修正：毎反復の頭で必ず Nothing リセット
Do While wsDef.Cells(i, 1).Value <> ""
    shpName = Trim(wsDef.Cells(i, 1).Value)
    Set shp = Nothing              ' ← ここでリセット
    On Error Resume Next
    Set shp = ws.Shapes(shpName)
    On Error GoTo 0
    If Not shp Is Nothing Then
        Call DoSomething(shp)
        Set shp = Nothing
    End If
    i = i + 1
Loop
```

**チェックポイント：**
- `On Error Resume Next` でオブジェクトを取得している箇所は直前に `Set obj = Nothing`
- ループが終わった後にも `Set obj = Nothing`（後続処理への値の持ち越し防止）

---

### ✅ チェック 2：`On Error Resume Next` の範囲が広すぎる

**バグパターン：** `On Error Resume Next` を有効にしたまま複数の処理を書くと、別の処理のエラーまで無視する。特に `Err.Number = 0` で成否を判定するパターンは `Err.Clear` がないと誤判定する

```vba
' ❌ 危険パターン：Err.Number が前の操作の残滓を持っている
On Error Resume Next
Set wsDef = targetWb.Worksheets("SheetName")   ' ここでエラーが起きても...
Call SomeOtherFunction()                        ' ここのエラーも無視される
srcName = shp.ConnectorFormat.BeginConnectedShape.Name
If Err.Number = 0 Then ...    ' ← 前のエラーが残っていて誤判定！
On Error GoTo 0
```

```vba
' ✅ 正しいパターン：範囲を最小限に絞り、Err.Clear を適切に挟む
On Error Resume Next
Set wsDef = targetWb.Worksheets("SheetName")
On Error GoTo 0   ' ← できるだけ早く元に戻す

' ループ内での判定が必要な場合
On Error Resume Next
Err.Clear                                                  ' ← 判定前にクリア
srcName = shp.ConnectorFormat.BeginConnectedShape.Name
dstName = shp.ConnectorFormat.EndConnectedShape.Name
If Err.Number = 0 And srcName <> "" Then ...
Err.Clear                                                  ' ← 次の反復に持ち越さない
On Error GoTo 0
```

**チェックポイント：**
- `On Error Resume Next` の有効範囲を最小限に絞る（1〜2行で `On Error GoTo 0` に戻す）
- ループ内で `Err.Number = 0` 判定をするなら、直前に `Err.Clear`・直後にも `Err.Clear`
- 関数から抜ける前に `On Error GoTo 0` で必ず元に戻す

---

### ✅ チェック 3：数値変換前の型チェック欠如

**バグパターン：** `InputBox` の戻り値を直接 `CLng` / `CInt` に渡すと、非数値・キャンセル入力でクラッシュする

```vba
' ❌ バグあり："abc" や空文字でクラッシュ
Dim ans As String
ans = InputBox("番号を入力してください")
Dim idx As Long
idx = CLng(ans)   ' ← 型変換エラー！
```

```vba
' ✅ 修正：3段階チェック（キャンセル → 数値か → 範囲内か）
Dim ans As String
ans = InputBox("番号を入力してください")
If ans = "" Then Exit Sub                    ' キャンセル
If Not IsNumeric(ans) Then
    MsgBox "半角数字で入力してください。", vbExclamation
    Exit Sub
End If
Dim idx As Long
idx = CLng(ans)
If idx < 1 Or idx > maxCount Then           ' 範囲チェック
    MsgBox "不正な番号です。", vbExclamation
    Exit Sub
End If
```

**チェックポイント：**
- `CLng` / `CInt` / `CDbl` の前は必ず `IsNumeric` チェック
- InputBox のキャンセル（`ans = ""`）は `IsNumeric` の前に確認する（`IsNumeric("")` は False だが意図を明確にするため）
- 数値の範囲チェック（下限・上限）もセットで入れる

---

### ✅ チェック 4：フォント色プロパティの混同

**バグパターン：** VBA の図形フォント色には2系統のプロパティがあり、間違えると色が変わらない・エラーになる。AI はこれを頻繁に混同する

```vba
' ❌ TextFrame2 配下で Font.Color.RGB は使えない（エラーになる）
shp.TextFrame2.TextRange.Font.Color.RGB = RGB(0, 0, 0)

' ❌ TextFrame2 配下で Fill なしの ForeColor も効かないことがある
shp.TextFrame2.TextRange.Font.ForeColor.RGB = RGB(0, 0, 0)

' ✅ TextFrame2 配下でのフォント色（Fill.ForeColor.RGB が正しい）
shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
```

```vba
' フォームコントロール（ボタン）の文字色は別系統
' ✅ TextFrame 経由（旧オブジェクトモデル）
btn.TextFrame.Characters.Font.Color = RGB(255, 255, 255)
```

**チェックポイント：**
- **AutoShape の TextFrame2** でフォント色を設定するとき → `.Font.Fill.ForeColor.RGB`
- **フォームコントロール** のキャプション文字色 → `.TextFrame.Characters.Font.Color`
- `Font.Color.RGB` という書き方は `TextFrame2` 配下では機能しない点に注意

---

### ✅ チェック 5：`Optional` キーワードの付け忘れ

**バグパターン：** 引数を省略して呼びたい場面が後から生じたとき、`Optional` がなければコンパイルエラーになる

```vba
' ❌ txt が必須なので Call MySub(shp) のように省略して呼べない
Sub MySub(shp As Shape, txt As String)

' ✅ Optional にして省略可能にする
Sub MySub(shp As Shape, Optional txt As String = "")
```

**チェックポイント：**
- 複数箇所から呼ばれる Sub/Function の引数は `Optional` を早めに検討する
- `Optional` 引数には必ずデフォルト値を指定する（`= ""` / `= 0` / `= False`）
- 「引数が渡されたか」を判定したい場合は `Variant` 型にして `IsMissing(引数名)` を使う

---

### ✅ チェック 6：ループ内 `ReDim Preserve` の扱い

**バグパターン：** 件数が不明なコレクションを配列に詰める際、`ReDim Preserve` をループ都度呼ぶのは正しく動くが、件数が多いと O(n²) で著しく遅くなる

```vba
' 動くが遅い：1000件なら1000回コピーが走る
Dim arr() As String
Dim count As Long
count = 0
For Each shp In ws.Shapes
    count = count + 1
    ReDim Preserve arr(1 To count)   ' 毎回全体をコピー
    arr(count) = shp.Name
Next shp
```

```vba
' ✅ 件数が事前にわかる場合：1回だけ ReDim
Dim n As Long
n = ws.Shapes.Count
ReDim arr(1 To n)
For i = 1 To n
    arr(i) = ws.Shapes(i).Name
Next i

' ✅ 件数不明でも大量になりうる場合：Collection を使う
Dim col As New Collection
For Each shp In ws.Shapes
    col.Add shp.Name
Next shp
```

**チェックポイント：**
- 図形数・行数が数百件を超える可能性があるなら `Collection` や `Dictionary` を検討
- 小規模（数十件）なら `ReDim Preserve` でも実害はない
- `ReDim Preserve` は配列の最終次元しか変更できない点も要確認

---

### ✅ チェック 7：`Scripting.Dictionary` の生成忘れ

**バグパターン：** `Object` 型で宣言しただけでは `Nothing` のまま。`Set` と `CreateObject` を忘れると実行時エラー

```vba
' ❌ CreateObject を忘れて使おうとする
Dim dict As Object
dict("key") = value   ' ← オブジェクト変数が設定されていませんエラー

' ✅ 必ず Set + CreateObject で生成
Dim dict As Object
Set dict = CreateObject("Scripting.Dictionary")
dict("key") = value
```

**チェックポイント：**
- `Set dict = CreateObject("Scripting.Dictionary")` を宣言直後に書く習慣を持つ
- オブジェクトを Dictionary の値に格納する場合は `Set dict(key) = obj`（`Set` が必要）
- アクセス前に `dict.Exists(key)` チェックを入れることで未登録キーのエラーを防ぐ

---

### ✅ チェック 8：空行で途中終了するループ

**バグパターン：** `A列が空になったら終了` のループは、ユーザーが途中の行を削除したり空行を挿入すると後半データが読まれない

```vba
' ❌ 途中に空行があると止まる
Dim i As Long
i = 2
Do While ws.Cells(i, 1).Value <> ""
    ' i 行目の処理
    i = i + 1
Loop
```

```vba
' ✅ 最終行を先に取得してから For ループ
Dim lastRow As Long
lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
Dim i As Long
For i = 2 To lastRow
    If ws.Cells(i, 1).Value = "" Then GoTo NextRow   ' 空行は読み飛ばす
    ' i 行目の処理
NextRow:
Next i
```

**チェックポイント：**
- ユーザーが直接編集するシートを読む際は空行の混入を想定する
- `End(xlUp).Row` で最終行を取得してから `For` ループが安全
- 「空行＝データ終端」として扱いたいなら仕様として明記し、ユーザーに徹底させる

---

### ✅ チェック 9：ボタン系図形の除外忘れ

**バグパターン：** `ws.Shapes` をループすると、フォームコントロール（ボタン）や ActiveX コントロールも含まれる。名前変更・削除・エクスポートの処理でボタンを巻き込んでしまう

```vba
' ❌ ボタンも対象になってしまう
For Each shp In ws.Shapes
    shp.Name = "FLOW_" & i   ' フォームコントロールのボタン名も変わる！
Next shp
```

```vba
' ✅ Type チェックでボタン系を除外
For Each shp In ws.Shapes
    If shp.Type = msoFormControl Or shp.Type = msoOLEControlObject Then
        GoTo SkipShape
    End If
    shp.Name = "FLOW_" & i
SkipShape:
Next shp
```

**チェックポイント：**
- 図形を対象に処理するループには必ずボタン系の除外条件を入れる
- `msoFormControl`：フォームコントロール（旧来のボタン）
- `msoOLEControlObject`：ActiveX コントロール（コマンドボタン等）
- 削除処理は特に注意（一度消したボタンの再配置は手間がかかる）

---

### ✅ チェック 10：`OnAction` / `Application.OnTime` のマクロパス書式

**バグパターン：** ボタンの `OnAction` や `Application.OnTime` に設定するマクロパスの書式を間違えると、ボタンが反応しない・OnTime が発火しない

```vba
' ❌ パターン1：ブック名なし（他のブックが開いていると誤動作する）
btn.OnAction = "MyMacro"

' ❌ パターン2：ブック名あり・モジュール名なし（標準モジュールなら動くが、
'              クラスモジュール内の Sub は参照できない）
btn.OnAction = "'" & ThisWorkbook.Name & "'!MyMacro"

' ✅ 標準モジュールの Sub
btn.OnAction = "'" & ThisWorkbook.Name & "'!Module1.MyMacro"

' ✅ クラスモジュール（ThisWorkbook など）内の Sub
btn.OnAction = "'" & ThisWorkbook.Name & "'!ThisWorkbook.MyMacro"
```

```vba
' Application.OnTime も同様
' ❌ ブック名なし
Application.OnTime Now + TimeValue("00:00:03"), "ThisWorkbook.ClearStatusBar"

' ✅ 正しい形式
Application.OnTime Now + TimeValue("00:00:03"), _
    "'" & ThisWorkbook.Name & "'!ThisWorkbook.ClearStatusBar"
```

**チェックポイント：**
- `'ブック名'!モジュール名.Sub名` が基本形。ブック名はシングルクォートで囲む（日本語・スペース対応）
- ブック名は `ThisWorkbook.Name` で動的に取得する（ファイル名変更に強くなる）
- クラスモジュール（`ThisWorkbook`、`Sheet1` 等）内の Sub は必ず `モジュール名.` プレフィックスが必要
- `TimeValue` に渡す文字列は `"00:00:03"` の形式（ゼロ埋め2桁）に統一する

---

## まとめ：コードを受け取ったときのチェック手順

```
1. まずコンパイルを通す（Alt+F11 → デバッグ → コンパイル）
2. 以下を上から順にチェック：
   □ On Error Resume Next の範囲が最小限か、Err.Clear が適切に入っているか
   □ ループ内で Set obj = Nothing をしてから On Error Resume Next → 取得しているか
   □ InputBox など外部入力を CLng / CInt する前に IsNumeric があるか
   □ TextFrame2 配下のフォント色が .Font.Fill.ForeColor.RGB になっているか
   □ 省略して呼びたい引数に Optional がついているか
   □ 図形ループに msoFormControl / msoOLEControlObject の除外があるか
   □ OnAction / OnTime のパスが '書式'!モジュール名.Sub名 になっているか
3. 小さい単位で動作確認する（1マクロずつ実行）
```

---

*このドキュメントは新しいバグパターンが発見されたら随時更新する。*
