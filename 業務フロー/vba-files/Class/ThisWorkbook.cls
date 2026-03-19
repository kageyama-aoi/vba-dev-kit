VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'=============================================================================
' モジュール名  : FlowMacro
' 作成目的      : Excelシート上に業務フローの図形・コネクタを作成するマクロ群
'
' 【前提】
'   このマクロは「マクロを直接埋め込まない業務フロー用Excelブック」に対して
'   外部から呼び出すブックのボタン操作を通じて操作することを想定しています。
'   呼び出しブックにボタンを配置し、押下時に各マクロを呼び出してください。
'
' 【ActiveSheet について】─────────────────────────────────
'   外部ブックからマクロを呼び出す場合、ActiveSheet は呼び出し元ブックの
'   シートを指してしまい「システムエラー &H80004005」が発生することがあります。
'
'   このモジュールでは GetTargetSheet という専用ヘルパーで操作対象シートを
'   一元管理しています。対象シートの解決ロジックはそこだけ変更してください。
'
'   【現在の解決ロジック（GetTargetSheet 関数）】
'     1. 開いているブックの中から対象ブックキーワード（Config シート設定）を名前に含むものを検索
'        （ThisWorkbook 自身は除外。単体ファイル用途はフォールバックへ）
'     2. 該当ブックが 1 件 → そのブックを使用
'        該当ブックが複数 → InputBox でユーザーに選ばせる
'        該当なし         → ThisWorkbook をフォールバックとして使用
'     3. sheetName が明示指定されていればそのシートを返す
'        省略時は ActiveSheet を返す
'
'   ブック名の判定キーワードは Config シートの「対象ブックキーワード」で変更できます。
'   図形シートのシート名は Config シートの「フローシート名」で変更できます。
'
' 【公開マクロ一覧】────────────────────────────────────────
'
'  ■図形配置系
'   Flow_Process              … 「処理」図形（矩形）を選択セル位置に配置
'   Flow_Decision             … 「判断」図形（菱形）を選択セル位置に配置
'   Flow_StartEnd             … 「開始/終了」図形（角丸矩形）を選択セル位置に配置
'
'  ■スタイル系
'   FlowStyle                 … 図形の共通スタイル（塗り・罫線・フォント）を適用
'
'  ■接続系
'   ConnectSelectedShapes_Auto          … 選択した 2 図形を位置関係で自動接続
'   ConnectMultipleShapes_TopToBottom   … 選択した複数図形を上から下へ連続接続
'   ConnectMultipleShapes_LeftToRight   … 選択した複数図形を左から右へ連続接続
'
'  ■補助ユーティリティ系
'   SetParagraphAndFont       … 選択図形のフォント・段落書式を一括設定
'   SelectShapesInColumnRange … InputBox で指定した列範囲内の図形を一括選択
'   ToggleConnectorType       … コネクタの種類（直線 ⇔ 折れ線）を切り替え
'   InjectTextFromList        … 定義シートのリストから図形にテキストを一括注入
'   RenameShapesSequentially  … シート上の図形を Top 順に連番命名
'
'  ■内部ヘルパー（直接実行は非推奨）■
'   GetTargetSheet              … 操作対象シートを解決して返す（ActiveSheet問題の対策）
'   GetSelectedShapes           … 選択図形を切り出して Shape 配列に格納
'   ConnectTwoShapes_Vertical   … 2 図形を上下方向で接続
'   ConnectTwoShapes_Horizontal … 2 図形を左右方向で接続
'
' 【使い方（基本フロー）】──────────────────────────────────
'   1. 図形を配置したいセルをクリックして選択する
'   2. 配置マクロを実行（Flow_Process / Flow_Decision / Flow_StartEnd）
'      → 選択セルの位置に図形が作成される
'   3. 接続したい 2 図形を Ctrl+クリックで選択する
'   4. ConnectSelectedShapes_Auto を実行
'      → 位置関係（縦 or 横）を自動判定して接続する
'
' 【複数図形を一括接続する場合】────────────────────────────
'   ・縦に並んでいる場合 → ConnectMultipleShapes_TopToBottom
'   ・横に並んでいる場合 → ConnectMultipleShapes_LeftToRight
'
' 【定数について】────────────────────────────────────────────
'   モジュール冒頭の「■ 設定定数」セクションで図形サイズ・色・フォントを
'   一括管理しています。デザイン変更はそこだけ編集してください。
'
' 【変更履歴】──────────────────────────────────────────────────
'   v1.0  初版作成
'   v2.0  定数化・エラー処理・コメント整備
'   v3.0  図形追加（判断・開始終了）/ 選択および接続マクロ追加
'   v4.0  コメント・ドキュメント整備 / SelectShapesInColumnRange をInputBox対応に変更
'   v4.1  外部ブック呼び出し時の ActiveSheet 問題に対応（GetTargetSheet 新設）
'   v4.2  GetTargetSheet に sheetName 引数追加 / 複数ブック競合ロジック修正
'   v4.3  RenameShapesSequentially でボタン系図形（フォームコントロール/ActiveX）を除外
'=============================================================================

Option Explicit


'=============================================================================
' ■ 設定定数（デザイン変更はここを編集する）
'=============================================================================

' --- Config シートの名前（このシート名だけはコードで固定）---
Private Const CONFIG_SHEET_NAME As String = "Config"

' --- 処理・判断図形サイズ（センチメートル）---
Private Const SHAPE_WIDTH_CM  As Double = 2.5
Private Const SHAPE_HEIGHT_CM As Double = 1.2

' --- 開始/終了図形サイズ（センチメートル）---
Private Const SE_WIDTH_CM  As Double = 2.5
Private Const SE_HEIGHT_CM As Double = 0.8

' --- 図形の塗り色（RGB）---
Private Const FILL_R As Long = 255
Private Const FILL_G As Long = 255
Private Const FILL_B As Long = 255

' --- 罫線の色（RGB）---
Private Const BORDER_R As Long = 0
Private Const BORDER_G As Long = 0
Private Const BORDER_B As Long = 0

' --- 罫線の太さ（ポイント）---
Private Const BORDER_WEIGHT As Single = 0.75

' --- フォント設定 ---
Private Const FONT_NAME As String = "Meiryo UI"
Private Const FONT_SIZE As Single = 9

' --- コネクタ（矢印）の色と太さ ---
Private Const CONN_R      As Long = 0
Private Const CONN_G      As Long = 0
Private Const CONN_B      As Long = 0
Private Const CONN_WEIGHT As Single = 1


'=============================================================================
' ■ Config シートから読み込む設定値（実行時に LoadConfig で初期化）
'=============================================================================
Private m_Keyword   As String   ' 対象ブックキーワード
Private m_FlowSheet As String   ' フローシート名
Private m_DefSheet  As String   ' 定義シート名
Private m_ListSheet As String   ' 一覧シート名


'=============================================================================
' ■ 接続ポイント番号の定数
'   上=1 / 右=2 / 下=3 / 左=4
'=============================================================================
Private Const CP_TOP    As Long = 1
Private Const CP_RIGHT  As Long = 2
Private Const CP_BOTTOM As Long = 3
Private Const CP_LEFT   As Long = 4


'=============================================================================
' Sub Flow_Process
'=============================================================================
Sub Flow_Process()

    If Not LoadConfig() Then Exit Sub

    Dim ws  As Worksheet
    Dim shp As Shape

    Set ws = GetTargetSheet(m_FlowSheet)
    If ws Is Nothing Then Exit Sub

    Set shp = ws.Shapes.AddShape( _
        msoShapeRectangle, _
        Selection.Left, _
        Selection.Top, _
        Application.CentimetersToPoints(SHAPE_WIDTH_CM), _
        Application.CentimetersToPoints(SHAPE_HEIGHT_CM))

    Call FlowStyle(shp, "処理")

End Sub


'=============================================================================
' Sub Flow_Decision
'=============================================================================
Sub Flow_Decision()

    If Not LoadConfig() Then Exit Sub

    Dim ws  As Worksheet
    Dim shp As Shape

    Set ws = GetTargetSheet(m_FlowSheet)
    If ws Is Nothing Then Exit Sub

    Set shp = ws.Shapes.AddShape( _
        msoShapeDiamond, _
        Selection.Left, _
        Selection.Top, _
        Application.CentimetersToPoints(SHAPE_WIDTH_CM), _
        Application.CentimetersToPoints(SHAPE_HEIGHT_CM))

    Call FlowStyle(shp, "判断")

End Sub


'=============================================================================
' Sub Flow_StartEnd
'=============================================================================
Sub Flow_StartEnd()

    If Not LoadConfig() Then Exit Sub

    Dim ws  As Worksheet
    Dim shp As Shape

    Set ws = GetTargetSheet(m_FlowSheet)
    If ws Is Nothing Then Exit Sub

    Set shp = ws.Shapes.AddShape( _
        msoShapeRoundedRectangle, _
        Selection.Left, _
        Selection.Top, _
        Application.CentimetersToPoints(SE_WIDTH_CM), _
        Application.CentimetersToPoints(SE_HEIGHT_CM))

    Call FlowStyle(shp, "開始")

End Sub


'=============================================================================
' Sub FlowStyle
'=============================================================================
Sub FlowStyle(shp As Shape, txt As String)

    If shp.Connector Then Exit Sub

    With shp
        .Fill.ForeColor.RGB = RGB(FILL_R, FILL_G, FILL_B)

        .Line.ForeColor.RGB = RGB(BORDER_R, BORDER_G, BORDER_B)
        .Line.Weight = BORDER_WEIGHT

        With .TextFrame2
            .TextRange.Font.Name = FONT_NAME
            .TextRange.Font.NameFarEast = FONT_NAME
            .TextRange.Font.NameAscii = FONT_NAME
            .TextRange.Font.Size = FONT_SIZE
            .VerticalAnchor = msoAnchorMiddle
            .TextRange.ParagraphFormat.Alignment = msoAlignCenter
            .TextRange.Text = txt
        End With

    End With

End Sub


'=============================================================================
' Sub ConnectSelectedShapes_Auto
'=============================================================================
Sub ConnectSelectedShapes_Auto()

    If Not LoadConfig() Then Exit Sub

    If TypeName(Selection) <> "DrawingObjects" Then
        MsgBox "図形を選択してから実行してください。", vbExclamation, "選択エラー"
        Exit Sub
    End If

    If Selection.ShapeRange.count <> 2 Then
        MsgBox "図形を 2 つ選択してください。", vbExclamation, "選択エラー"
        Exit Sub
    End If

    Dim shp1 As Shape
    Dim shp2 As Shape
    Set shp1 = Selection.ShapeRange(1)
    Set shp2 = Selection.ShapeRange(2)

    Dim cx1 As Double, cy1 As Double
    Dim cx2 As Double, cy2 As Double

    cx1 = shp1.Left + shp1.Width / 2
    cy1 = shp1.Top + shp1.Height / 2
    cx2 = shp2.Left + shp2.Width / 2
    cy2 = shp2.Top + shp2.Height / 2

    Dim dx As Double
    Dim dy As Double
    dx = cx1 - cx2
    dy = cy1 - cy2

    If Abs(dx) > Abs(dy) Then
        If dx < 0 Then
            Call ConnectTwoShapes_Horizontal(shp1, shp2)
        Else
            Call ConnectTwoShapes_Horizontal(shp2, shp1)
        End If
    Else
        If dy < 0 Then
            Call ConnectTwoShapes_Vertical(shp1, shp2)
        Else
            Call ConnectTwoShapes_Vertical(shp2, shp1)
        End If
    End If

End Sub


'=============================================================================
' Sub ConnectMultipleShapes_TopToBottom
'=============================================================================
Sub ConnectMultipleShapes_TopToBottom()

    If Not LoadConfig() Then Exit Sub

    Dim arr() As Shape
    Dim n     As Long

    If Not GetSelectedShapes(arr, n) Then Exit Sub

    Call SortShapesByAxis(arr, n, "Top")

    Dim i As Long
    For i = 1 To n - 1
        Call ConnectTwoShapes_Vertical(arr(i), arr(i + 1))
    Next i

End Sub


'=============================================================================
' Sub ConnectMultipleShapes_LeftToRight
'=============================================================================
Sub ConnectMultipleShapes_LeftToRight()

    If Not LoadConfig() Then Exit Sub

    Dim arr() As Shape
    Dim n     As Long

    If Not GetSelectedShapes(arr, n) Then Exit Sub

    Call SortShapesByAxis(arr, n, "Left")

    Dim i As Long
    For i = 1 To n - 1
        Call ConnectTwoShapes_Horizontal(arr(i), arr(i + 1))
    Next i

End Sub


'=============================================================================
'
'  ── 内部ヘルパー（直接実行は非推奨） ──
'
'=============================================================================

'=============================================================================
' Function LoadConfig  【内部ヘルパー・直接実行は非推奨】
' 概要 : Config シートから設定値を読み込み、モジュール変数に格納する
' 戻値 : True = 正常読み込み / False = エラー（処理を中断すること）
'
' Config シートの形式（ThisWorkbook 内）：
'   1行目  … ヘッダー（読み飛ばし）
'   A列    … 設定名
'   B列    … 値
'   設定名 ：対象ブックキーワード / フローシート名 / 定義シート名 / 一覧シート名
'=============================================================================
Private Function LoadConfig() As Boolean

    LoadConfig = False

    m_Keyword   = GetConfigValue("対象ブックキーワード")
    If m_Keyword = "" Then Exit Function

    m_FlowSheet = GetConfigValue("フローシート名")
    If m_FlowSheet = "" Then Exit Function

    m_DefSheet  = GetConfigValue("定義シート名")
    If m_DefSheet = "" Then Exit Function

    m_ListSheet = GetConfigValue("一覧シート名")
    If m_ListSheet = "" Then Exit Function

    LoadConfig = True

End Function


'=============================================================================
' Function GetConfigValue  【内部ヘルパー・直接実行は非推奨】
' 概要 : Config シートから指定キーの値を返す
'        シートが存在しない / キーが見つからない / 値が空の場合はエラー表示して "" を返す
'=============================================================================
Private Function GetConfigValue(key As String) As String

    GetConfigValue = ""

    ' Config シートの存在確認
    Dim wsConfig As Worksheet
    On Error Resume Next
    Set wsConfig = ThisWorkbook.Worksheets(CONFIG_SHEET_NAME)
    On Error GoTo 0

    If wsConfig Is Nothing Then
        MsgBox "「" & CONFIG_SHEET_NAME & "」シートが見つかりません。" & vbCrLf & _
               "マクロブックに Config シートを作成してください。", _
               vbExclamation, "設定エラー"
        Exit Function
    End If

    ' A列でキーを検索（2行目以降）
    Dim i   As Long
    Dim val As String

    For i = 2 To 1000
        If wsConfig.Cells(i, 1).Value = "" Then Exit For
        If wsConfig.Cells(i, 1).Value = key Then
            val = Trim(wsConfig.Cells(i, 2).Value)
            If val = "" Then
                MsgBox "Config シートの「" & key & "」の値が空です。設定してください。", _
                       vbExclamation, "設定エラー"
                Exit Function
            End If
            GetConfigValue = val
            Exit Function
        End If
    Next i

    MsgBox "Config シートに「" & key & "」の設定が見つかりません。", _
           vbExclamation, "設定エラー"

End Function

'=============================================================================
' Function GetTargetSheet  【内部ヘルパー・直接実行は非推奨】
' 概要 : 操作対象の Worksheet を解決して返す
'
' 引数 : sheetName … 取得したいシート名（省略時は ActiveSheet を返す）
' 戻値 : Worksheet オブジェクト（解決失敗時は Nothing を返しメッセージを表示）
'
' 解決ロジック:
'   1. 対象ブックキーワード（Config シートで設定）を名前に含むブックを検索（ThisWorkbook は除外）
'   2. 0件 → ThisWorkbook をフォールバック
'      1件 → そのブックを使用
'      複数 → InputBox でユーザーに選ばせる
'   3. sheetName 指定あり → 該当シートを返す（なければエラーメッセージ）
'      sheetName 省略    → ActiveSheet を返す
'=============================================================================
Private Function GetTargetSheet(Optional sheetName As String = "") As Worksheet

    Dim wb         As Workbook
    Dim matched()  As Workbook
    Dim matchCount As Long
    Dim targetWb   As Workbook

    matchCount = 0

    ' --- 対象ブックキーワードを名前に含むブックを検索（ThisWorkbook は除外）---
    For Each wb In Application.Workbooks
        If InStr(wb.Name, m_Keyword) > 0 Then
            If wb.Name <> ThisWorkbook.Name Then
                matchCount = matchCount + 1
                ReDim Preserve matched(1 To matchCount)
                Set matched(matchCount) = wb
            End If
        End If
    Next wb

    ' --- ヒット数でターゲットブックを決定 ---
    Select Case matchCount

        Case 0
            ' 外部ブックなし → 単体ファイル用途とみなして ThisWorkbook を使用
            Set targetWb = ThisWorkbook

        Case 1
            Set targetWb = matched(1)

        Case Else
            ' 複数ヒット → ユーザーに選ばせる
            Dim i       As Long
            Dim choices As String
            choices = "対象ブックを番号で選んでください:" & vbCrLf
            For i = 1 To matchCount
                choices = choices & i & ": " & matched(i).Name & vbCrLf
            Next i

            Dim ans As String
            ans = InputBox(choices, "対象ブックの選択", "1")
            If ans = "" Then
                Set GetTargetSheet = Nothing
                Exit Function
            End If

            Dim idx As Long
            idx = CLng(ans)
            If idx < 1 Or idx > matchCount Then
                MsgBox "不正な番号です。", vbExclamation, "入力エラー"
                Set GetTargetSheet = Nothing
                Exit Function
            End If

            Set targetWb = matched(idx)

    End Select

    ' --- シート名指定の有無で返すシートを選択 ---
    If sheetName <> "" Then
        On Error GoTo ErrNoSheet
        Set GetTargetSheet = targetWb.Worksheets(sheetName)
        On Error GoTo 0
    Else
        Set GetTargetSheet = targetWb.ActiveSheet
    End If

    Exit Function

ErrNoSheet:
    MsgBox "シート「" & sheetName & "」が見つかりません。" & vbCrLf & _
           "Config シートの「フローシート名」/「定義シート名」を確認してください。", _
           vbExclamation, "シート取得エラー"
    Set GetTargetSheet = Nothing

End Function


'=============================================================================
' Sub SortShapesByAxis  【内部ヘルパー・直接実行は非推奨】
' 概要 : Shape 配列を指定軸の座標でバブルソートする
' 引数 : arr()  … Shape 配列（参照渡し・ソート結果で上書き）
'         n      … 配列の要素数
'         axis   … ソートキー "Top"（上から下）/ "Left"（左から右）
'=============================================================================
Private Sub SortShapesByAxis(arr() As Shape, n As Long, axis As String)

    Dim i   As Long
    Dim j   As Long
    Dim tmp As Shape

    For i = 1 To n - 1
        For j = i + 1 To n
            Dim v1 As Double
            Dim v2 As Double
            If axis = "Top" Then
                v1 = arr(i).Top
                v2 = arr(j).Top
            Else
                v1 = arr(i).Left
                v2 = arr(j).Left
            End If
            If v1 > v2 Then
                Set tmp = arr(i)
                Set arr(i) = arr(j)
                Set arr(j) = tmp
            End If
        Next j
    Next i

End Sub


'=============================================================================
' Function IsButtonShape  【内部ヘルパー・直接実行は非推奨】
' 概要 : シェイプがボタン系（フォームコントロール / ActiveX）かどうかを返す
' 戻値 : True = ボタン系（操作対象外）/ False = 通常図形
'=============================================================================
Private Function IsButtonShape(shp As Shape) As Boolean

    IsButtonShape = (shp.Type = msoFormControl Or shp.Type = msoOLEControlObject)

End Function


'=============================================================================
' Function GetShapeKind  【内部ヘルパー・直接実行は非推奨】
' 概要 : シェイプの種別を文字列で返す
' 戻値 : "Process" / "Decision" / "StartEnd" / "Arrow" / "Shape"
'=============================================================================
Private Function GetShapeKind(shp As Shape) As String

    If shp.Connector Then
        GetShapeKind = "Arrow"
    ElseIf shp.AutoShapeType = msoShapeRectangle Then
        GetShapeKind = "Process"
    ElseIf shp.AutoShapeType = msoShapeDiamond Then
        GetShapeKind = "Decision"
    ElseIf shp.AutoShapeType = msoShapeRoundedRectangle Then
        GetShapeKind = "StartEnd"
    Else
        GetShapeKind = "Shape"
    End If

End Function


'=============================================================================
' Function GetSelectedShapes  【内部ヘルパー・直接実行は非推奨】
'=============================================================================
Private Function GetSelectedShapes(arr() As Shape, n As Long) As Boolean

    GetSelectedShapes = False

    If TypeName(Selection) <> "DrawingObjects" Then
        MsgBox "図形を選択してから実行してください。", vbExclamation, "選択エラー"
        Exit Function
    End If

    n = Selection.ShapeRange.count

    If n < 2 Then
        MsgBox "2 つ以上の図形を選択してください。", vbExclamation, "選択エラー"
        Exit Function
    End If

    ReDim arr(1 To n)

    Dim i As Long
    For i = 1 To n
        Set arr(i) = Selection.ShapeRange(i)
    Next i

    GetSelectedShapes = True

End Function


'=============================================================================
' Sub ConnectTwoShapes_Vertical  【内部ヘルパー・直接実行は非推奨】
'=============================================================================
Private Sub ConnectTwoShapes_Vertical(shp1 As Shape, shp2 As Shape)

    Dim ws   As Worksheet
    Dim conn As Shape

    Set ws = GetTargetSheet(m_FlowSheet)
    If ws Is Nothing Then Exit Sub

    Set conn = ws.Shapes.AddConnector(msoConnectorStraight, 0, 0, 0, 0)

    With conn
        .Line.ForeColor.RGB = RGB(CONN_R, CONN_G, CONN_B)
        .Line.Weight = CONN_WEIGHT
        .Line.EndArrowheadStyle = msoArrowheadTriangle

        .ConnectorFormat.BeginConnect shp1, CP_BOTTOM
        .ConnectorFormat.EndConnect shp2, CP_TOP

        .RerouteConnections
    End With

End Sub


'=============================================================================
' Sub ConnectTwoShapes_Horizontal  【内部ヘルパー・直接実行は非推奨】
'=============================================================================
Private Sub ConnectTwoShapes_Horizontal(shp1 As Shape, shp2 As Shape)

    Dim ws   As Worksheet
    Dim conn As Shape

    Set ws = GetTargetSheet(m_FlowSheet)
    If ws Is Nothing Then Exit Sub

    Set conn = ws.Shapes.AddConnector(msoConnectorStraight, 0, 0, 0, 0)

    With conn
        .Line.ForeColor.RGB = RGB(CONN_R, CONN_G, CONN_B)
        .Line.Weight = CONN_WEIGHT
        .Line.EndArrowheadStyle = msoArrowheadTriangle

        .ConnectorFormat.BeginConnect shp1, CP_RIGHT
        .ConnectorFormat.EndConnect shp2, CP_LEFT

        .RerouteConnections
    End With

End Sub


'=============================================================================
'
'  ── 補助ユーティリティ ──
'
'=============================================================================

'=============================================================================
' Sub SetParagraphAndFont
'=============================================================================
Sub SetParagraphAndFont()

    If Not LoadConfig() Then Exit Sub

    Dim shp As Shape

    If TypeName(Selection) <> "DrawingObjects" Then
        MsgBox "図形を選択してください", vbExclamation, "選択エラー"
        Exit Sub
    End If

    For Each shp In Selection.ShapeRange

        If shp.TextFrame2.HasText Then

            With shp.TextFrame2.TextRange.ParagraphFormat
                .Alignment = msoAlignCenter
                .SpaceWithin = 12
                .FirstLineIndent = 0
                .LeftIndent = 0
                .RightIndent = 0
            End With

            With shp.TextFrame2.TextRange.Font
                .Name = FONT_NAME
                .Size = FONT_SIZE
            End With

        End If

    Next shp

End Sub


'=============================================================================
' Sub SelectShapesInColumnRange
'=============================================================================
Sub SelectShapesInColumnRange()

    If Not LoadConfig() Then Exit Sub

    Dim ws As Worksheet
    Set ws = GetTargetSheet(m_FlowSheet)
    If ws Is Nothing Then Exit Sub

    Dim inputStart As String
    Dim inputEnd   As String

    inputStart = InputBox("開始列を入力してください（例: B, BG）", "列範囲の指定（開始）", "BG")
    If inputStart = "" Then Exit Sub

    inputEnd = InputBox("終了列を入力してください（例: F, BN）", "列範囲の指定（終了）", "BN")
    If inputEnd = "" Then Exit Sub

    Dim startCol As Double
    Dim endCol   As Double

    On Error GoTo ErrInvalidColumn
    startCol = ws.Columns(inputStart).Left
    endCol = ws.Columns(inputEnd).Left + ws.Columns(inputEnd).Width
    On Error GoTo 0

    Dim shp        As Shape
    Dim shpNames() As String
    Dim count      As Long

    count = 0

    For Each shp In ws.Shapes
        If shp.Left >= startCol And shp.Left <= endCol Then
            count = count + 1
            ReDim Preserve shpNames(1 To count)
            shpNames(count) = shp.Name
        End If
    Next shp

    If count > 0 Then
        ws.Shapes.Range(shpNames).Select
    Else
        MsgBox "対象図形なし", vbInformation, "完了"
    End If

    Exit Sub

ErrInvalidColumn:
    MsgBox "列名が認識できません。アルファベットで入力してください（例: B, BG）", _
           vbExclamation, "入力エラー"

End Sub


'=============================================================================
' Sub ToggleConnectorType
'=============================================================================
Sub ToggleConnectorType()

    If Not LoadConfig() Then Exit Sub

    Dim shp As Shape

    If TypeName(Selection) <> "DrawingObjects" Then
        MsgBox "図形を選択してください", vbExclamation, "選択エラー"
        Exit Sub
    End If

    For Each shp In Selection.ShapeRange

        If shp.Connector Then

            Select Case shp.ConnectorFormat.Type

                Case msoConnectorStraight
                    shp.ConnectorFormat.Type = msoConnectorElbow

                Case msoConnectorElbow
                    shp.ConnectorFormat.Type = msoConnectorStraight

                Case msoConnectorCurve
                    shp.ConnectorFormat.Type = msoConnectorStraight

            End Select

        End If

    Next shp

End Sub


'=============================================================================
' Sub InjectTextFromList
' 概要 : 定義シートのリストを読み、図形に対応するテキストと書式を一括適用する
'
' 定義シート（Config シートの「定義シート名」で指定）の形式：
'   1行目 … ヘッダー行（読み飛ばす）
'   A列   … 図形の名前（例: Flow_001）
'   B列   … 挿入するテキスト
'
' 備考 : 図形名は RenameShapesSequentially で事前に付与しておく
'         テキストを注入すると同時に FlowStyle を呼んで書式を適用する
'=============================================================================
Sub InjectTextFromList()

    If Not LoadConfig() Then Exit Sub

    Dim wsDef    As Worksheet
    Dim ws       As Worksheet
    Dim shp      As Shape
    Dim i        As Long
    Dim shpName  As String
    Dim txt      As String
    Dim hitCount As Long

    Set wsDef = GetTargetSheet(m_DefSheet)
    If wsDef Is Nothing Then Exit Sub

    Set ws = GetTargetSheet(m_FlowSheet)
    If ws Is Nothing Then Exit Sub

    hitCount = 0
    i = 2   ' 1行目はヘッダーとして読み飛ばし

    Do While wsDef.Cells(i, 1).Value <> ""

        shpName = Trim(wsDef.Cells(i, 1).Value)
        txt = Trim(wsDef.Cells(i, 2).Value)

        On Error Resume Next
        Set shp = ws.Shapes(shpName)
        On Error GoTo 0

        If Not shp Is Nothing Then
            Call FlowStyle(shp, txt)
            hitCount = hitCount + 1
            Set shp = Nothing
        End If

        i = i + 1
    Loop

    MsgBox hitCount & " 件の図形にテキストと書式を適用しました。", _
           vbInformation, "完了"

End Sub


'=============================================================================
' Sub RenameShapesSequentially
' 概要 : シート上の図形を Top 座標順に "FLOW_{種別}_{通し番号3桁}" 形式で命名する
' 備考 : InjectTextFromList の前に実行する
'         番号は種別をまたいで通し番号で採番する
'         ボタン（フォームコントロール / ActiveX）は除外（名前変更しない）
'
' 命名ルール：
'   FLOW_Process_001  … 矩形（処理）
'   FLOW_Decision_002 … 菱形（判断）
'   FLOW_StartEnd_003 … 角丸矩形（開始/終了）
'   FLOW_Arrow_004    … コネクタ（矢印）
'   FLOW_Shape_005    … その他の図形
'=============================================================================
Sub RenameShapesSequentially()

    If Not LoadConfig() Then Exit Sub

    Dim ws  As Worksheet
    Dim arr() As Shape
    Dim n   As Long
    Dim i   As Long, j As Long
    Dim tmp As Shape

    Set ws = GetTargetSheet(m_FlowSheet)
    If ws Is Nothing Then Exit Sub

    If ws.Shapes.count = 0 Then
        MsgBox "図形がありません", vbInformation
        Exit Sub
    End If

    ' ボタン系を除外して収集
    n = 0
    Dim k As Long
    For k = 1 To ws.Shapes.count
        If Not IsButtonShape(ws.Shapes(k)) Then
            n = n + 1
            ReDim Preserve arr(1 To n)
            Set arr(n) = ws.Shapes(k)
        End If
    Next k

    ' Top 座標順にソート
    For i = 1 To n - 1
        For j = i + 1 To n
            If arr(i).Top > arr(j).Top Then
                Set tmp = arr(i): Set arr(i) = arr(j): Set arr(j) = tmp
            End If
        Next j
    Next i

    ' 種別を判定して通し番号で命名
    For i = 1 To n
        arr(i).Name = "FLOW_" & GetShapeKind(arr(i)) & "_" & Format(i, "000")
    Next i

    MsgBox n & " 件の図形を連番で命名しました。", vbInformation, "完了"

    ' --- 図形一覧シートに書き出す ---
    Call ExportShapeList(ws.Parent, arr, n)

End Sub


'=============================================================================
' Sub ExportShapeList  【内部ヘルパー・直接実行は非推奨】
' 概要 : リネーム済み図形の一覧を「図形一覧」シートに書き出す
'        シートが存在しない場合は新規作成、存在する場合は上書き
'
' 書き出し形式：
'   1行目  … ヘッダー（A: 図形名 / B: 種別）
'   2行目以降 … 図形名と種別（Top 座標順）
'=============================================================================
Private Sub ExportShapeList(wb As Workbook, arr() As Shape, n As Long)

    Dim wsList As Worksheet
    Dim ws     As Worksheet

    Set ws = wb.Worksheets(m_FlowSheet)

    ' シートの取得または新規作成
    On Error Resume Next
    Set wsList = wb.Worksheets(m_ListSheet)
    On Error GoTo 0

    If wsList Is Nothing Then
        Set wsList = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.count))
        wsList.Name = m_ListSheet
    Else
        wsList.Cells.ClearContents
    End If

    ' ヘッダー行
    wsList.Cells(1, 1).Value = "図形名"
    wsList.Cells(1, 2).Value = "テキスト"
    wsList.Cells(1, 3).Value = "種別"
    wsList.Cells(1, 4).Value = "行"
    wsList.Cells(1, 5).Value = "列"

    ' データ行：図形名・テキスト・種別・位置を書き出す
    Dim i   As Long
    Dim txt As String

    For i = 1 To n
        ' テキスト取得（コネクタ等でエラーになる場合は空欄）
        On Error Resume Next
        txt = arr(i).TextFrame2.TextRange.Text
        If Err.Number <> 0 Then txt = ""
        On Error GoTo 0

        wsList.Cells(i + 1, 1).Value = arr(i).Name
        wsList.Cells(i + 1, 2).Value = txt
        wsList.Cells(i + 1, 3).Value = GetShapeKind(arr(i))
        wsList.Cells(i + 1, 4).Value = TopToRowNum(ws, arr(i).Top)
        wsList.Cells(i + 1, 5).Value = LeftToColLetter(ws, arr(i).Left)
    Next i

    ' 列幅を自動調整
    wsList.Columns("A:E").AutoFit

End Sub


'=============================================================================
' Function TopToRowNum  【内部ヘルパー・直接実行は非推奨】
' 概要 : Top 座標（ポイント）から最も近い Excel 行番号を返す
'=============================================================================
Private Function TopToRowNum(ws As Worksheet, topPt As Double) As Long

    Dim cum As Double
    Dim r   As Long

    cum = 0
    For r = 1 To 2000
        cum = cum + ws.Rows(r).Height
        If cum >= topPt Then
            TopToRowNum = r
            Exit Function
        End If
    Next r

    TopToRowNum = r

End Function


'=============================================================================
' Function LeftToColLetter  【内部ヘルパー・直接実行は非推奨】
' 概要 : Left 座標（ポイント）から最も近い Excel 列アルファベットを返す
'=============================================================================
Private Function LeftToColLetter(ws As Worksheet, leftPt As Double) As String

    Dim cum As Double
    Dim c   As Long

    cum = 0
    For c = 1 To 500
        cum = cum + ws.Columns(c).Width
        If cum >= leftPt Then
            LeftToColLetter = Split(ws.Cells(1, c).Address(True, False), "$")(0)
            Exit Function
        End If
    Next c

    LeftToColLetter = Split(ws.Cells(1, c).Address(True, False), "$")(0)

End Function
