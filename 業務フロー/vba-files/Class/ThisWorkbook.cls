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
' ���W���[����  : FlowMacro
' �쐬�ړI      : Excel�V�[�g��ɋƖ��t���[�}�`�E�R�l�N�^���쐬����}�N���W
'
' �y�O��z
'   ���̃}�N���́u�}�N���𒼐ږ��ߍ��߂Ȃ��Ɩ��t���[�pExcel�u�b�N�v�ɑ΂���
'   �O���̌Ăяo�����u�b�N����{�^������ʂ��đ��삷�邱�Ƃ�z�肵�Ă��܂��B
'   �Ăяo�����u�b�N�Ƀ{�^����z�u���A��������e�}�N�����Ăяo���Ă��������B
'
' �yActiveSheet �ɂ��āz����������������������������������������������������
'   �O���u�b�N����}�N�����Ăяo���ꍇ�AActiveSheet ���Ăяo�����u�b�N��
'   �V�[�g���w���Ă��܂��u�V�X�e���G���[ &H80004005�v���������邱�Ƃ�����܂��B
'
'   ���̃��W���[���ł� GetTargetSheet �Ƃ������ʃw���p�[�ő���ΏۃV�[�g��
'   �ꌳ�Ǘ����Ă��܂��B�ΏۃV�[�g�̉������W�b�N�͂���������ύX���Ă��������B
'
'   �y���݂̉������W�b�N�iGetTargetSheet ���j�z
'     1. �J���Ă���u�b�N�̒����� TARGET_BOOK_KEYWORD �𖼑O�Ɋ܂ނ��̂�T��
'        �iThisWorkbook ���g�͏��O�B����t�@�C���^�p���̓t�H�[���o�b�N�ցj
'     2. ���������u�b�N�� 1 �� �� ���̃u�b�N���g�p
'        ���������u�b�N������ �� InputBox �Ń��[�U�[�ɑI�΂���
'        ������Ȃ� �� ThisWorkbook ���t�H�[���o�b�N�Ƃ��Ďg�p
'     3. sheetName �������w�肳��Ă���΂��̃V�[�g��Ԃ�
'        �ȗ����� ActiveSheet ��Ԃ�
'
'   �u�b�N���̔���L�[���[�h�͒萔 TARGET_BOOK_KEYWORD �ŕύX�ł��܂��B
'   �}�`�V�[�g�̃V�[�g���͒萔 FLOW_SHEET_NAME �ŕύX�ł��܂��B
'
' �y���^�}�N���ꗗ�z����������������������������������������������������������
'
'  ���}�`�z�u��
'   Flow_Process              �c �u�����v�}�`�i��`�j��I���Z���ʒu�ɔz�u
'   Flow_Decision             �c �u���f�v�}�`�i�H�`�j��I���Z���ʒu�ɔz�u
'   Flow_StartEnd             �c �u�J�n�^�I���v�}�`�i�p�ۋ�`�j��I���Z���ʒu�ɔz�u
'
'  ���X�^�C����
'   FlowStyle                 �c �}�`�̋��ʃX�^�C���i�h��E�g���E�t�H���g�j��K�p
'
'  ���ڑ���
'   ConnectSelectedShapes_Auto          �c �I������ 2 �}�`���ʒu�֌W�Ŏ����ڑ�
'   ConnectMultipleShapes_TopToBottom   �c �I�����������}�`���と�����ɘA���ڑ�
'   ConnectMultipleShapes_LeftToRight   �c �I�����������}�`�������E���ɘA���ڑ�
'
'  ���⏕���[�e�B���e�B��
'   SetParagraphAndFont       �c �I��}�`�̃t�H���g�E�i���������ꊇ�ݒ�
'   SelectShapesInColumnRange �c InputBox �Ŏw�肵����͈͓��̐}�`���ꊇ�I��
'   ToggleConnectorType       �c �R�l�N�^�̐���i���� �� �G���{�j��؂�ւ�
'   InjectTextFromList        �c ��`�V�[�g�̃��X�g����}�`�Ƀe�L�X�g���ꊇ����
'   RenameShapesSequentially  �c �V�[�g��̐}�`�� Top ���ɘA�Ԗ�������
'
'  �������w���p�[�i���ڎ��s�͔񐄏��j��
'   GetTargetSheet              �c ����ΏۃV�[�g���������ĕԂ��iActiveSheet���̑΍�j
'   GetSelectedShapes           �c �I��}�`�����؂��� Shape �z��Ɋi�[
'   ConnectTwoShapes_Vertical   �c 2 �}�`���と�������Őڑ�
'   ConnectTwoShapes_Horizontal �c 2 �}�`�������E�����Őڑ�
'
' �y�g�����i��{�t���[�j�z��������������������������������������������������
'   1. �}�`��z�u�������Z�����N���b�N���đI������
'   2. �z�u�}�N�������s�iFlow_Process / Flow_Decision / Flow_StartEnd�j
'      �� �I���Z���̈ʒu�ɐ}�`���쐬�����
'   3. �ڑ������� 2 �}�`�� Ctrl+�N���b�N�őI������
'   4. ConnectSelectedShapes_Auto �����s
'      �� �ʒu�֌W�i�c or ���j���������肵�Ė�������
'
' �y�����}�`���ꊇ�ڑ�����ꍇ�z��������������������������������������������
'   �E�c�ɕ���ł���ꍇ �� ConnectMultipleShapes_TopToBottom
'   �E���ɕ���ł���ꍇ �� ConnectMultipleShapes_LeftToRight
'
' �y�萔�ɂ��āz����������������������������������������������������������������
'   ���W���[���`���́u�� �ݒ�萔�v�Z�N�V�����Ő}�`�T�C�Y�E�F�E�t�H���g��
'   �ꊇ�Ǘ����Ă��܂��B�f�U�C���ύX�͂���������ҏW���Ă��������B
'
' �y�ύX�����z������������������������������������������������������������������������
'   v1.0  ���ō쐬
'   v2.0  �萔���E�G���[�����E�R�����g����
'   v3.0  �}�`�ǉ��i���f�E�J�n�I���j/ �����ѐڑ��}�N���ǉ�
'   v4.0  �R�����g�E�h�L�������g���� / SelectShapesInColumnRange ��InputBox�Ή��ɕύX
'   v4.1  �O���u�b�N�Ăяo������ ActiveSheet ���ɑΉ��iGetTargetSheet �����j
'   v4.2  GetTargetSheet �� sheetName �����ǉ� / �����u�b�N�딚�����C��
'=============================================================================

Option Explicit


'=============================================================================
' �� �ݒ�萔�i�f�U�C���ύX�͂�����ҏW����j
'=============================================================================

' --- ����Ώۃu�b�N�̔���L�[���[�h ---
Private Const TARGET_BOOK_KEYWORD As String = "�Ɩ��t���["

' --- �}�`��z�u����V�[�g�̖��O ---
'     ���̃V�[�g���}�`�z�u�E�ڑ��̑ΏۂɂȂ�
Private Const FLOW_SHEET_NAME As String = "Sheet1"

' --- ��`�V�[�g�̖��O ---
'     InjectTextFromList ���Q�Ƃ���}�`���E�e�L�X�g�̑Ή��\�V�[�g
Private Const DEF_SHEET_NAME As String = "�t���[��`"

' --- �����E���f�}�`�T�C�Y�i�Z���`���[�g���j---
Private Const SHAPE_WIDTH_CM  As Double = 2.5
Private Const SHAPE_HEIGHT_CM As Double = 1.2

' --- �J�n�^�I���}�`�T�C�Y�i�Z���`���[�g���j---
Private Const SE_WIDTH_CM  As Double = 2.5
Private Const SE_HEIGHT_CM As Double = 0.8

' --- �}�`�̓h��F�iRGB�j---
Private Const FILL_R As Long = 255
Private Const FILL_G As Long = 255
Private Const FILL_B As Long = 255

' --- �g���̐F�iRGB�j---
Private Const BORDER_R As Long = 0
Private Const BORDER_G As Long = 0
Private Const BORDER_B As Long = 0

' --- �g���̑����i�|�C���g�j---
Private Const BORDER_WEIGHT As Single = 0.75

' --- �t�H���g�ݒ� ---
Private Const FONT_NAME As String = "Meiryo UI"
Private Const FONT_SIZE As Single = 9

' --- �R�l�N�^�i���j�̐F�Ƒ��� ---
Private Const CONN_R      As Long = 0
Private Const CONN_G      As Long = 0
Private Const CONN_B      As Long = 0
Private Const CONN_WEIGHT As Single = 1


'=============================================================================
' �� �ڑ��|�C���g�ԍ��̒萔
'   ��=1 / �E=2 / ��=3 / ��=4
'=============================================================================
Private Const CP_TOP    As Long = 1
Private Const CP_RIGHT  As Long = 2
Private Const CP_BOTTOM As Long = 3
Private Const CP_LEFT   As Long = 4


'=============================================================================
' Sub Flow_Process
'=============================================================================
Sub Flow_Process()

    Dim ws  As Worksheet
    Dim shp As Shape

    Set ws = GetTargetSheet(FLOW_SHEET_NAME)
    If ws Is Nothing Then Exit Sub

    Set shp = ws.Shapes.AddShape( _
        msoShapeRectangle, _
        Selection.Left, _
        Selection.Top, _
        Application.CentimetersToPoints(SHAPE_WIDTH_CM), _
        Application.CentimetersToPoints(SHAPE_HEIGHT_CM))

    Call FlowStyle(shp, "����")

End Sub


'=============================================================================
' Sub Flow_Decision
'=============================================================================
Sub Flow_Decision()

    Dim ws  As Worksheet
    Dim shp As Shape

    Set ws = GetTargetSheet(FLOW_SHEET_NAME)
    If ws Is Nothing Then Exit Sub

    Set shp = ws.Shapes.AddShape( _
        msoShapeDiamond, _
        Selection.Left, _
        Selection.Top, _
        Application.CentimetersToPoints(SHAPE_WIDTH_CM), _
        Application.CentimetersToPoints(SHAPE_HEIGHT_CM))

    Call FlowStyle(shp, "���f")

End Sub


'=============================================================================
' Sub Flow_StartEnd
'=============================================================================
Sub Flow_StartEnd()

    Dim ws  As Worksheet
    Dim shp As Shape

    Set ws = GetTargetSheet(FLOW_SHEET_NAME)
    If ws Is Nothing Then Exit Sub

    Set shp = ws.Shapes.AddShape( _
        msoShapeRoundedRectangle, _
        Selection.Left, _
        Selection.Top, _
        Application.CentimetersToPoints(SE_WIDTH_CM), _
        Application.CentimetersToPoints(SE_HEIGHT_CM))

    Call FlowStyle(shp, "�J�n")

End Sub


'=============================================================================
' Sub FlowStyle
'=============================================================================
Sub FlowStyle(shp As Shape, txt As String)

    Debug.Print "FlowStyle�J�n shp.Name=" & shp.Name & " Connector=" & shp.Connector
    If shp.Connector Then Exit Sub

    With shp
        Debug.Print "Fill"
        .Fill.ForeColor.RGB = RGB(FILL_R, FILL_G, FILL_B)

        Debug.Print "Line"
        .Line.ForeColor.RGB = RGB(BORDER_R, BORDER_G, BORDER_B)
        .Line.Weight = BORDER_WEIGHT

        With .TextFrame2
            Debug.Print "Font.Name"
            .TextRange.Font.Name = FONT_NAME
            Debug.Print "Font.NameFarEast"
            .TextRange.Font.NameFarEast = FONT_NAME
            Debug.Print "Font.NameAscii"
            .TextRange.Font.NameAscii = FONT_NAME
            Debug.Print "Font.Size"
            .TextRange.Font.Size = FONT_SIZE
            Debug.Print "VerticalAnchor"
            .VerticalAnchor = msoAnchorMiddle
            Debug.Print "Alignment"
            .TextRange.ParagraphFormat.Alignment = msoAlignCenter
            Debug.Print "Text"
            .TextRange.Text = txt
        End With

    End With

    Debug.Print "FlowStyle����"

End Sub


'=============================================================================
' Sub ConnectSelectedShapes_Auto
'=============================================================================
Sub ConnectSelectedShapes_Auto()

    If TypeName(Selection) <> "DrawingObjects" Then
        MsgBox "�}�`��I�����Ă�����s���Ă��������B", vbExclamation, "�I���G���["
        Exit Sub
    End If

    If Selection.ShapeRange.count <> 2 Then
        MsgBox "�}�`�� 2 �����I�����Ă��������B", vbExclamation, "�I���G���["
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

    Dim arr() As Shape
    Dim n     As Long

    If Not GetSelectedShapes(arr, n) Then Exit Sub

    Dim i   As Long
    Dim j   As Long
    Dim tmp As Shape

    For i = 1 To n - 1
        For j = i + 1 To n
            If arr(i).Top > arr(j).Top Then
                Set tmp = arr(i)
                Set arr(i) = arr(j)
                Set arr(j) = tmp
            End If
        Next j
    Next i

    For i = 1 To n - 1
        Call ConnectTwoShapes_Vertical(arr(i), arr(i + 1))
    Next i

End Sub


'=============================================================================
' Sub ConnectMultipleShapes_LeftToRight
'=============================================================================
Sub ConnectMultipleShapes_LeftToRight()

    Dim arr() As Shape
    Dim n     As Long

    If Not GetSelectedShapes(arr, n) Then Exit Sub

    Dim i   As Long
    Dim j   As Long
    Dim tmp As Shape

    For i = 1 To n - 1
        For j = i + 1 To n
            If arr(i).Left > arr(j).Left Then
                Set tmp = arr(i)
                Set arr(i) = arr(j)
                Set arr(j) = tmp
            End If
        Next j
    Next i

    For i = 1 To n - 1
        Call ConnectTwoShapes_Horizontal(arr(i), arr(i + 1))
    Next i

End Sub


'=============================================================================
'
'  ������ �����w���p�[�i���ڎ��s�͔񐄏��j ������
'
'=============================================================================

'=============================================================================
' Function GetTargetSheet  �m�����w���p�[�E���ڎ��s�͔񐄏��n
' �T�v : ����Ώۂ� Worksheet ���������ĕԂ�
'
' ���� : sheetName �c �擾�������V�[�g���i�ȗ����� ActiveSheet ��Ԃ��j
' �ߒl : Worksheet �I�u�W�F�N�g�i�������s���� Nothing ��Ԃ����b�Z�[�W��\���j
'
' �������W�b�N:
'   1. TARGET_BOOK_KEYWORD �𖼑O�Ɋ܂ރu�b�N�����W�iThisWorkbook �͏��O�j
'   2. 0�� �� ThisWorkbook ���t�H�[���o�b�N
'      1�� �� ���̃u�b�N���g�p
'      ���� �� InputBox �Ń��[�U�[�ɑI�΂���
'   3. sheetName �w�肠�� �� �Y���V�[�g��Ԃ��i�Ȃ���΃G���[���b�Z�[�W�j
'      sheetName �ȗ�    �� ActiveSheet ��Ԃ�
'=============================================================================
Private Function GetTargetSheet(Optional sheetName As String = "") As Worksheet

    Dim wb         As Workbook
    Dim matched()  As Workbook
    Dim matchCount As Long
    Dim targetWb   As Workbook

    matchCount = 0

    ' --- TARGET_BOOK_KEYWORD �𖼑O�Ɋ܂ރu�b�N�����W�iThisWorkbook �͏��O�j---
    For Each wb In Application.Workbooks
        If InStr(wb.Name, TARGET_BOOK_KEYWORD) > 0 Then
            If wb.Name <> ThisWorkbook.Name Then
                matchCount = matchCount + 1
                ReDim Preserve matched(1 To matchCount)
                Set matched(matchCount) = wb
            End If
        End If
    Next wb

    ' --- �q�b�g���Ń^�[�Q�b�g�u�b�N������ ---
    Select Case matchCount

        Case 0
            ' �O���u�b�N�Ȃ� �� ����t�@�C���^�p�Ƃ݂Ȃ��� ThisWorkbook ���g�p
            Set targetWb = ThisWorkbook

        Case 1
            Set targetWb = matched(1)

        Case Else
            ' �����q�b�g �� ���[�U�[�ɑI�΂���
            Dim i       As Long
            Dim choices As String
            choices = "�Ώۃu�b�N��ԍ��őI��ł�������:" & vbCrLf
            For i = 1 To matchCount
                choices = choices & i & ": " & matched(i).Name & vbCrLf
            Next i

            Dim ans As String
            ans = InputBox(choices, "�Ώۃu�b�N�̑I��", "1")
            If ans = "" Then
                Set GetTargetSheet = Nothing
                Exit Function
            End If

            Dim idx As Long
            idx = CLng(ans)
            If idx < 1 Or idx > matchCount Then
                MsgBox "�����Ȕԍ��ł��B", vbExclamation, "���̓G���["
                Set GetTargetSheet = Nothing
                Exit Function
            End If

            Set targetWb = matched(idx)

    End Select

    ' --- �V�[�g���w��̗L���ŕԂ��V�[�g������ ---
    If sheetName <> "" Then
        On Error GoTo ErrNoSheet
        Set GetTargetSheet = targetWb.Worksheets(sheetName)
        On Error GoTo 0
    Else
        Set GetTargetSheet = targetWb.ActiveSheet
    End If

    Exit Function

ErrNoSheet:
    MsgBox "�V�[�g�u" & sheetName & "�v��������܂���B" & vbCrLf & _
           "�萔 FLOW_SHEET_NAME / DEF_SHEET_NAME ���m�F���Ă��������B", _
           vbExclamation, "�V�[�g�擾�G���["
    Set GetTargetSheet = Nothing

End Function


'=============================================================================
' Function GetSelectedShapes  �m�����w���p�[�E���ڎ��s�͔񐄏��n
'=============================================================================
Private Function GetSelectedShapes(arr() As Shape, n As Long) As Boolean

    GetSelectedShapes = False

    If TypeName(Selection) <> "DrawingObjects" Then
        MsgBox "�}�`��I�����Ă�����s���Ă��������B", vbExclamation, "�I���G���["
        Exit Function
    End If

    n = Selection.ShapeRange.count

    If n < 2 Then
        MsgBox "2 �ȏ�̐}�`��I�����Ă��������B", vbExclamation, "�I���G���["
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
' Sub ConnectTwoShapes_Vertical  �m�����w���p�[�E���ڎ��s�͔񐄏��n
'=============================================================================
Private Sub ConnectTwoShapes_Vertical(shp1 As Shape, shp2 As Shape)

    Dim ws   As Worksheet
    Dim conn As Shape

    Set ws = GetTargetSheet(FLOW_SHEET_NAME)
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
' Sub ConnectTwoShapes_Horizontal  �m�����w���p�[�E���ڎ��s�͔񐄏��n
'=============================================================================
Private Sub ConnectTwoShapes_Horizontal(shp1 As Shape, shp2 As Shape)

    Dim ws   As Worksheet
    Dim conn As Shape

    Set ws = GetTargetSheet(FLOW_SHEET_NAME)
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
'  ������ �⏕���[�e�B���e�B ������
'
'=============================================================================

'=============================================================================
' Sub SetParagraphAndFont
'=============================================================================
Sub SetParagraphAndFont()

    Dim shp As Shape

    If TypeName(Selection) <> "DrawingObjects" Then
        MsgBox "�}�`��I�����Ă�������", vbExclamation, "�I���G���["
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
                .Name = "���C���I"
                .Size = 9
            End With

        End If

    Next shp

End Sub


'=============================================================================
' Sub SelectShapesInColumnRange
'=============================================================================
Sub SelectShapesInColumnRange()

    Dim ws As Worksheet
    Set ws = GetTargetSheet(FLOW_SHEET_NAME)
    If ws Is Nothing Then Exit Sub

    Dim inputStart As String
    Dim inputEnd   As String

    inputStart = InputBox("�J�n�����͂��Ă��������i��: B, BG�j", "��͈͂̎w��i�J�n�j", "BG")
    If inputStart = "" Then Exit Sub

    inputEnd = InputBox("�I�������͂��Ă��������i��: F, BN�j", "��͈͂̎w��i�I���j", "BN")
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
        MsgBox "�Ώې}�`�Ȃ�", vbInformation, "����"
    End If

    Exit Sub

ErrInvalidColumn:
    MsgBox "�񖼂�����������܂���B�A���t�@�x�b�g�œ��͂��Ă��������i��: B, BG�j", _
           vbExclamation, "���̓G���["

End Sub


'=============================================================================
' Sub ToggleConnectorType
'=============================================================================
Sub ToggleConnectorType()

    Dim shp As Shape

    If TypeName(Selection) <> "DrawingObjects" Then
        MsgBox "�}�`��I�����Ă�������", vbExclamation, "�I���G���["
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
' �T�v : ��`�V�[�g�̃��X�g��ǂ݁A�}�`���ɑΉ�����e�L�X�g�Ə������ꊇ�K�p����
'
' ��`�V�[�g�iDEF_SHEET_NAME�j�̌`���F
'   1�s�� �c �w�b�_�[�s�i�ǂݔ�΂����j
'   A��   �c �}�`�̖��O�i��: Flow_001�j
'   B��   �c ��������e�L�X�g
'
' ���l : �}�`���� RenameShapesSequentially �Ŏ��O�ɕt�^���Ă���
'         �e�L�X�g�𒍓�����Ɠ����� FlowStyle ���Ă�ŏ�������
'=============================================================================
Sub InjectTextFromList()

    Dim wsDef    As Worksheet
    Dim ws       As Worksheet
    Dim shp      As Shape
    Dim i        As Long
    Dim shpName  As String
    Dim txt      As String
    Dim hitCount As Long

    Set wsDef = GetTargetSheet(DEF_SHEET_NAME)
    If wsDef Is Nothing Then Exit Sub

    Set ws = GetTargetSheet(FLOW_SHEET_NAME)
    If ws Is Nothing Then Exit Sub

    hitCount = 0
    i = 2   ' 1�s�ڂ̓w�b�_�[�Ƃ��ēǂݔ�΂�

    Do While wsDef.Cells(i, 1).Value <> ""

        shpName = Trim(wsDef.Cells(i, 1).Value)
        txt = Trim(wsDef.Cells(i, 2).Value)
        
        ' ���m�F�p
        Debug.Print "i=" & i & " shpName=[" & shpName & "] txt=[" & txt & "]"

        On Error Resume Next
        Set shp = ws.Shapes(shpName)
        On Error GoTo 0

        ' ���m�F�p
        Debug.Print "shp Is Nothing: " & (shp Is Nothing)

        If Not shp Is Nothing Then
            Call FlowStyle(shp, txt)
            hitCount = hitCount + 1
            Set shp = Nothing
        End If

        i = i + 1
    Loop

    MsgBox hitCount & " �̐}�`�Ƀe�L�X�g�Ə�����K�p���܂����B", _
           vbInformation, "����"

End Sub


'=============================================================================
' Sub RenameShapesSequentially
' �T�v : �V�[�g��̐}�`�� Top ���W���� "Flow_001"�`�̘A�ԂŖ���������
' ���l : InjectTextFromList �̑O�Ɏ��s����
'         �R�l�N�^���܂߂ĘA�Ԃ��U����
'=============================================================================
Sub RenameShapesSequentially()

    Dim ws  As Worksheet
    Dim arr() As Shape
    Dim n   As Long
    Dim i   As Long, j As Long
    Dim tmp As Shape

    Set ws = GetTargetSheet(FLOW_SHEET_NAME)
    If ws Is Nothing Then Exit Sub

    If ws.Shapes.count = 0 Then
        MsgBox "�}�`������܂���", vbInformation
        Exit Sub
    End If

    n = 0
    Dim k As Long
    For k = 1 To ws.Shapes.count
        If ws.Shapes(k).Type <> msoFormControl And _
           ws.Shapes(k).Type <> msoOLEControlObject Then
            n = n + 1
            ReDim Preserve arr(1 To n)
            Set arr(n) = ws.Shapes(k)
        End If
    Next k

    For i = 1 To n - 1
        For j = i + 1 To n
            If arr(i).Top > arr(j).Top Then
                Set tmp = arr(i): Set arr(i) = arr(j): Set arr(j) = tmp
            End If
        Next j
    Next i

    For i = 1 To n
        arr(i).Name = "Flow_" & Format(i, "000")
    Next i

    MsgBox n & " �̐}�`��A�ԂŖ������܂����B", vbInformation, "����"

End Sub

