VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PatternSelectForm
   Caption         =   "作成シートの選択"
   ClientHeight    =   5895
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4335
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "PatternSelectForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=============================================================================
' フォーム名  : PatternSelectForm
' 作成目的    : SetupTestScenarioSheets から呼び出すパターン選択フォーム
'
' 【使い方】
'   With PatternSelectForm
'       Set .Patterns = <Collection>   ' 抽出済みパターン一覧をセット
'       .Show
'       If .Cancelled Then
'           ' キャンセル処理
'       Else
'           ' .SelectedPatterns に選択結果が入っている
'       End If
'   End With
'   Unload PatternSelectForm
'=============================================================================
Option Explicit

' ─── 入出力プロパティ ─────────────────────────────────────
Public Patterns         As Collection   ' 呼び出し元がセット
Public SelectedPatterns As Collection   ' OK押下後に呼び出し元が読む
Public Cancelled        As Boolean      ' キャンセルされたか

' ─── コントロール（WithEvents で動的追加後もイベント受信） ──
Private WithEvents btnOK     As MSForms.CommandButton
Private WithEvents btnCancel As MSForms.CommandButton
Private WithEvents btnAll    As MSForms.CommandButton
Private WithEvents btnNone   As MSForms.CommandButton
Private lstPatterns          As MSForms.ListBox

'=============================================================================
' ■ フォーム初期化
'=============================================================================
Private Sub UserForm_Initialize()

    Cancelled = True  ' デフォルトはキャンセル扱い

    Dim itemCount As Long
    If Not Patterns Is Nothing Then itemCount = Patterns.Count

    ' ── フォームサイズ ──────────────────────────────────────
    Dim listH As Long
    listH = WorksheetFunction.Min(itemCount, 15) * 18 + 10
    If listH < 80 Then listH = 80

    Dim fW As Long: fW = 300
    Dim fH As Long: fH = listH + 134
    Me.Width = fW + 14
    Me.Height = fH + 28

    ' ── タイトルラベル ──────────────────────────────────────
    Dim lbl As MSForms.Label
    Set lbl = Me.Controls.Add("Forms.Label.1")
    With lbl
        .Caption = "作成するシートを選択してください（" & itemCount & " 件）"
        .Top = 6: .Left = 6: .Width = fW - 4: .Height = 18
        .FontBold = True
    End With

    ' ── リストボックス ──────────────────────────────────────
    Set lstPatterns = Me.Controls.Add("Forms.ListBox.1", "lstPatterns")
    With lstPatterns
        .Top = 28: .Left = 6: .Width = fW - 4: .Height = listH
        .MultiSelect = fmMultiSelectMulti
        .BorderStyle = fmBorderStyleSingle
    End With

    Dim btnTop As Long: btnTop = listH + 34

    ' ── 全選択ボタン ────────────────────────────────────────
    Set btnAll = Me.Controls.Add("Forms.CommandButton.1", "btnAll")
    With btnAll
        .Caption = "全選択": .Top = btnTop: .Left = 6: .Width = 64: .Height = 22
    End With

    ' ── 全解除ボタン ────────────────────────────────────────
    Set btnNone = Me.Controls.Add("Forms.CommandButton.1", "btnNone")
    With btnNone
        .Caption = "全解除": .Top = btnTop: .Left = 76: .Width = 64: .Height = 22
    End With

    ' ── OK ボタン ───────────────────────────────────────────
    Set btnOK = Me.Controls.Add("Forms.CommandButton.1", "btnOK")
    With btnOK
        .Caption = "OK": .Top = btnTop + 30: .Left = fW - 148: .Width = 68: .Height = 26
        .Default = True
    End With

    ' ── キャンセルボタン ────────────────────────────────────
    Set btnCancel = Me.Controls.Add("Forms.CommandButton.1", "btnCancel")
    With btnCancel
        .Caption = "キャンセル": .Top = btnTop + 30: .Left = fW - 74: .Width = 78: .Height = 26
        .Cancel = True
    End With

End Sub

'=============================================================================
' ■ Activate 時にリストを構築（Patterns セット後に発火するため）
'=============================================================================
Private Sub UserForm_Activate()
    If lstPatterns.ListCount > 0 Then Exit Sub  ' 2回目以降はスキップ
    If Patterns Is Nothing Then Exit Sub
    Dim i As Long
    For i = 1 To Patterns.Count
        lstPatterns.AddItem CStr(Patterns(i))
        lstPatterns.Selected(i - 1) = True
    Next i
End Sub

'=============================================================================
' ■ ボタンイベント
'=============================================================================
Private Sub btnAll_Click()
    Dim i As Long
    For i = 0 To lstPatterns.ListCount - 1
        lstPatterns.Selected(i) = True
    Next i
End Sub

Private Sub btnNone_Click()
    Dim i As Long
    For i = 0 To lstPatterns.ListCount - 1
        lstPatterns.Selected(i) = False
    Next i
End Sub

Private Sub btnOK_Click()
    ' 選択数チェック
    Dim selCount As Long
    Dim i As Long
    For i = 0 To lstPatterns.ListCount - 1
        If lstPatterns.Selected(i) Then selCount = selCount + 1
    Next i

    If selCount = 0 Then
        MsgBox "少なくとも1件選択してください。", vbExclamation, "選択なし"
        Exit Sub
    End If

    ' 選択結果を収集
    Set SelectedPatterns = New Collection
    For i = 0 To lstPatterns.ListCount - 1
        If lstPatterns.Selected(i) Then
            SelectedPatterns.Add lstPatterns.List(i)
        End If
    Next i

    Cancelled = False
    Me.Hide
End Sub

Private Sub btnCancel_Click()
    Cancelled = True
    Me.Hide
End Sub

'=============================================================================
' ■ × ボタンでもキャンセル扱い
'=============================================================================
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancelled = True
        Cancel = 1
        Me.Hide
    End If
End Sub
