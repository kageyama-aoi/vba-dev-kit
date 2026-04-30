Attribute VB_Name = "TestScenarioSetup"
'=============================================================================
' モジュール名  : TestScenarioSetup
' 作成目的      : テストシナリオシートからパターン番号を抽出し、
'                 各パターン用シートの自動生成・シートリンク一覧の作成を行う
'
' 【前提】
'   このマクロは「マクロブック」から「別のテスト用ブック」を操作します。
'   既存の FlowMacro (ThisWorkbook.cls) とは完全に独立したモジュールです。
'
' 【公開マクロ一覧】
'   SetupTestScenarioSheets … メインエントリ。ファイル選択から全処理を実行する
'
'=============================================================================
Option Explicit

' ─── 定数 ───────────────────────────────────────────────
Private Const TARGET_SHEET_NAME  As String = "テストシナリオ"
Private Const PATTERN_NO_LABEL   As String = "パターンNo"
Private Const LINK_SHEET_NAME    As String = "シートリンク"
Private Const LOG_SHEET_NAME     As String = "ログ"
Private Const MAX_SHEET_NAME_LEN As Integer = 31

' シート名に使用できない文字
Private Const INVALID_CHARS      As String = "\/:*?[]"

'=============================================================================
' ■ メインエントリ
'=============================================================================
Public Sub SetupTestScenarioSheets()

    Dim targetWb         As Workbook
    Dim ws               As Worksheet
    Dim foundCell        As Range
    Dim patterns         As Collection
    Dim createdSheets    As Collection
    Dim selectedPatterns As Collection
    Dim errContext       As String

    ' ── Step1: ファイル選択 & オープン（キャンセルは正常終了）──
    Set targetWb = OpenTargetWorkbook()
    If targetWb Is Nothing Then Exit Sub

    On Error GoTo Cleanup
    Application.ScreenUpdating = False

    ' ── Step2: テストシナリオシート確認 ─────────────────────────
    errContext = "テストシナリオシート確認"
    If Not SheetExists(targetWb, TARGET_SHEET_NAME) Then
        MsgBox "「" & TARGET_SHEET_NAME & "」シートが見つかりません。" & vbCrLf & _
               "対象ブック：" & targetWb.Name, vbExclamation, "シートが見つかりません"
        GoTo Cleanup
    End If
    Set ws = targetWb.Worksheets(TARGET_SHEET_NAME)

    ' ── Step3: パターンNoセル検索 ────────────────────────────────
    errContext = "パターンNoセル検索"
    Set foundCell = FindPatternNoCell(ws)
    If foundCell Is Nothing Then
        MsgBox "「" & PATTERN_NO_LABEL & "」という項目がシート内に見つかりません。" & vbCrLf & _
               "テストシナリオシートのフォーマットを確認してください。", vbExclamation, "項目が見つかりません"
        GoTo Cleanup
    End If
    WriteLog targetWb, "パターンNoセル発見：" & foundCell.Address(False, False)

    ' ── Step4: パターン番号抽出 ──────────────────────────────────
    errContext = "パターン番号抽出"
    Set patterns = ExtractPatterns(foundCell)
    If patterns.Count = 0 Then
        MsgBox "「" & PATTERN_NO_LABEL & "」の下に値が見つかりませんでした。", vbExclamation, "データなし"
        GoTo Cleanup
    End If

    ' ── Step5: 重複チェック ──────────────────────────────────────
    errContext = "重複チェック"
    CheckDuplicates patterns, targetWb

    ' ── Step5.5: パターン選択フォーム ───────────────────────────
    errContext = "パターン選択"
    Set selectedPatterns = ShowPatternSelectForm(patterns)
    If selectedPatterns Is Nothing Then
        WriteLog targetWb, "パターン選択：ユーザーがキャンセル"
        GoTo Cleanup
    End If

    ' ── Step6: パターン用シート作成 ──────────────────────────────
    errContext = "シート作成"
    Set createdSheets = CreatePatternSheets(targetWb, selectedPatterns)

    ' ── Step7: シートリンクシート作成 ────────────────────────────
    errContext = "シートリンク作成"
    CreateLinkSheet targetWb, createdSheets

    ' ── 完了 ─────────────────────────────────────────────────────
    WriteLog targetWb, "処理完了："  & createdSheets.Count & " 件のシートを作成。時刻：" & Format(Now, "yyyy/mm/dd hh:mm:ss")
    MsgBox "処理が完了しました。" & vbCrLf & _
           "作成シート数：" & createdSheets.Count & " 件" & vbCrLf & _
           "詳細は「" & LOG_SHEET_NAME & "」シートを確認してください。", vbInformation, "完了"

    ActivateLinkSheet targetWb

Cleanup:
    Application.ScreenUpdating = True
    If Err.Number <> 0 Then
        MsgBox "エラー（" & errContext & "）: " & Err.Description, vbCritical, "エラー"
    End If

End Sub

'=============================================================================
' ■ パターン選択フォームの表示（#33）
'=============================================================================
Private Function ShowPatternSelectForm(patterns As Collection) As Collection

    With PatternSelectForm
        Set .Patterns = patterns
        .Show
        If .Cancelled Then
            Set ShowPatternSelectForm = Nothing
        Else
            Set ShowPatternSelectForm = .SelectedPatterns
        End If
    End With
    Unload PatternSelectForm

End Function

'=============================================================================
' ■ シートリンクシートをアクティブにする（#34 #35）
'=============================================================================
Private Sub ActivateLinkSheet(targetWb As Workbook)

    If SheetExists(targetWb, LINK_SHEET_NAME) Then
        targetWb.Worksheets(LINK_SHEET_NAME).Activate
    End If

End Sub

'=============================================================================
' ■ ファイル選択 & オープン
'=============================================================================
Private Function OpenTargetWorkbook() As Workbook

    Dim filePath As String
    filePath = Application.GetOpenFilename( _
        FileFilter:="Excel ファイル (*.xlsx;*.xlsm;*.xls),*.xlsx;*.xlsm;*.xls", _
        Title:="操作対象のブックを選択してください")

    If filePath = "False" Then
        Set OpenTargetWorkbook = Nothing
        Exit Function
    End If

    On Error Resume Next
    Dim wb As Workbook
    Set wb = Workbooks.Open(filePath)
    On Error GoTo 0

    If wb Is Nothing Then
        MsgBox "ブックを開けませんでした。" & vbCrLf & filePath, vbCritical, "エラー"
        Set OpenTargetWorkbook = Nothing
    Else
        Set OpenTargetWorkbook = wb
    End If

End Function

'=============================================================================
' ■ シート存在確認
'=============================================================================
Private Function SheetExists(wb As Workbook, sheetName As String) As Boolean

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    SheetExists = Not (ws Is Nothing)

End Function

'=============================================================================
' ■ パターンNoセル検索
'=============================================================================
Private Function FindPatternNoCell(ws As Worksheet) As Range

    Dim foundCell As Range
    Set foundCell = ws.Cells.Find( _
        What:=PATTERN_NO_LABEL, _
        LookIn:=xlValues, _
        LookAt:=xlWhole, _
        MatchCase:=False)

    Set FindPatternNoCell = foundCell

End Function

'=============================================================================
' ■ パターン番号抽出（見つかったセルの結合範囲末尾の次行から空白まで）
'=============================================================================
Private Function ExtractPatterns(patternNoCell As Range) As Collection

    Dim col     As New Collection
    Dim r       As Long
    Dim val     As String
    Dim startRow As Long

    ' 結合セル対応：MergeArea の末尾行の次行から開始
    ' 結合なし（1行）の場合も MergeArea.Rows.Count = 1 なので同じ計算で動く
    Dim mergeArea As Range
    Set mergeArea = patternNoCell.MergeArea
    startRow = mergeArea.Row + mergeArea.Rows.Count

    Dim ws As Worksheet
    Set ws = patternNoCell.Parent
    Dim c  As Long
    c = mergeArea.Column  ' 結合の左端列を使用

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, c).End(xlUp).Row
    For r = startRow To lastRow
        val = CStr(ws.Cells(r, c).Value)
        If val <> "" Then col.Add val
    Next r

    Set ExtractPatterns = col

End Function

'=============================================================================
' ■ 重複チェック
'=============================================================================
Private Sub CheckDuplicates(patterns As Collection, targetWb As Workbook)

    Dim i As Long, j As Long
    Dim dupList As String
    Dim checked() As Boolean
    ReDim checked(1 To patterns.Count)

    For i = 1 To patterns.Count
        If Not checked(i) Then
            For j = i + 1 To patterns.Count
                If patterns(i) = patterns(j) Then
                    If InStr(dupList, patterns(i)) = 0 Then
                        dupList = dupList & "  ・" & patterns(i) & vbCrLf
                    End If
                    checked(j) = True
                End If
            Next j
        End If
    Next i

    If dupList <> "" Then
        MsgBox "以下のパターン番号が重複しています。" & vbCrLf & _
               "重複分は1枚のみ作成し、処理を続行します。" & vbCrLf & vbCrLf & _
               dupList, vbExclamation, "重複あり"
        WriteLog targetWb, "重複値あり：" & vbCrLf & dupList
    Else
        WriteLog targetWb, "重複チェック：重複なし"
    End If

End Sub

'=============================================================================
' ■ シート名バリデーション
'=============================================================================
Private Function ValidateSheetName(sheetName As String) As String
    ' 戻り値：空文字 = OK、それ以外 = エラーメッセージ

    Dim i As Long
    Dim c As String

    ' 長さチェック
    If Len(sheetName) > MAX_SHEET_NAME_LEN Then
        ValidateSheetName = "シート名は " & MAX_SHEET_NAME_LEN & " 文字以内にしてください。" & _
            "（現在 " & Len(sheetName) & " 文字）"
        Exit Function
    End If

    ' 空文字チェック
    If Trim(sheetName) = "" Then
        ValidateSheetName = "シート名を空にすることはできません。"
        Exit Function
    End If

    ' 使用不可文字チェック
    For i = 1 To Len(INVALID_CHARS)
        c = Mid(INVALID_CHARS, i, 1)
        If InStr(sheetName, c) > 0 Then
            ValidateSheetName = "「" & c & "」はシート名に使用できない文字です。" & _
                "使用できない文字：" & INVALID_CHARS
            Exit Function
        End If
    Next i

    ValidateSheetName = ""  ' OK

End Function

'=============================================================================
' ■ パターン用シート作成
'=============================================================================
Private Function CreatePatternSheets(targetWb As Workbook, patterns As Collection) As Collection

    Dim createdSheets As New Collection
    Dim addedNames    As New Collection  ' 重複スキップ用（今回実行分）
    Dim sheetName     As String
    Dim errMsg        As String
    Dim i             As Long
    Dim alreadyAdded  As Boolean
    Dim n             As Long
    Dim newWs         As Worksheet

    For i = 1 To patterns.Count
        sheetName = CStr(patterns(i))

        ' 今回実行内での重複（2枚目以降）はスキップ
        alreadyAdded = False
        For n = 1 To addedNames.Count
            If addedNames(n) = sheetName Then
                alreadyAdded = True
                Exit For
            End If
        Next n
        If alreadyAdded Then GoTo NextPattern

        ' バリデーション
        errMsg = ValidateSheetName(sheetName)
        If errMsg <> "" Then
            MsgBox "シート名「" & sheetName & "」は作成できません。" & vbCrLf & vbCrLf & _
                   "【問題点】" & vbCrLf & errMsg, vbExclamation, "シート名エラー"
            WriteLog targetWb, "スキップ（無効名）：" & sheetName & " → " & errMsg
            GoTo NextPattern
        End If

        ' 既存シートの重複チェック
        If SheetExists(targetWb, sheetName) Then
            MsgBox "シート「" & sheetName & "」はすでに存在するためスキップします。", _
                   vbExclamation, "既存シート"
            WriteLog targetWb, "スキップ（既存）：" & sheetName
            GoTo NextPattern
        End If

        ' シート作成
        Set newWs = targetWb.Worksheets.Add(After:=targetWb.Worksheets(targetWb.Worksheets.Count))
        newWs.Name = sheetName
        createdSheets.Add sheetName
        addedNames.Add sheetName
        WriteLog targetWb, "作成：" & sheetName

NextPattern:
    Next i

    Set CreatePatternSheets = createdSheets

End Function

'=============================================================================
' ■ シートリンクシートの作成
'=============================================================================
Private Sub CreateLinkSheet(targetWb As Workbook, createdSheets As Collection)

    Dim ws As Worksheet

    ' 既存なら上書き
    If SheetExists(targetWb, LINK_SHEET_NAME) Then
        Set ws = targetWb.Worksheets(LINK_SHEET_NAME)
        ws.Cells.Clear
    Else
        Set ws = targetWb.Worksheets.Add(After:=targetWb.Worksheets(targetWb.Worksheets.Count))
        ws.Name = LINK_SHEET_NAME
    End If

    ' ヘッダー
    ws.Cells(1, 2).Value = "シート名"
    ws.Cells(1, 3).Value = "リンク"
    ws.Rows(1).Font.Bold = True

    ' データ行
    Dim i   As Long
    Dim row As Long
    For i = 1 To createdSheets.Count
        row = i + 1
        ws.Cells(row, 2).Value = createdSheets(i)
        ws.Cells(row, 3).Formula = "=HYPERLINK(""#'""&B" & row & "&""'!A1"",""⇒""&B" & row & ")"
    Next i

    ' 列幅調整
    ws.Columns("B:C").AutoFit

    WriteLog targetWb, LINK_SHEET_NAME & " シートを作成：" & createdSheets.Count & " 件"

End Sub

'=============================================================================
' ■ ログ記録
'=============================================================================
Private Sub WriteLog(targetWb As Workbook, message As String)

    Dim ws As Worksheet

    If SheetExists(targetWb, LOG_SHEET_NAME) Then
        Set ws = targetWb.Worksheets(LOG_SHEET_NAME)
    Else
        Set ws = targetWb.Worksheets.Add(After:=targetWb.Worksheets(targetWb.Worksheets.Count))
        ws.Name = LOG_SHEET_NAME
        ws.Cells(1, 1).Value = "日時"
        ws.Cells(1, 2).Value = "メッセージ"
        ws.Rows(1).Font.Bold = True
    End If

    ' 最終行の次に追記
    Dim nextRow As Long
    nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    ws.Cells(nextRow, 1).Value = Format(Now, "yyyy/mm/dd hh:mm:ss")
    ws.Cells(nextRow, 2).Value = message

End Sub

'=============================================================================
' ■ 空のパターンシートを一括削除 #32
'=============================================================================
Public Sub CleanupEmptyPatternSheets()

    Dim targetWb     As Workbook
    Dim linkWs       As Worksheet
    Dim emptySheets  As New Collection
    Dim linkRows     As New Collection
    Dim r            As Long
    Dim lastRow      As Long
    Dim sheetName    As String
    Dim ws           As Worksheet
    Dim listMsg      As String
    Dim i            As Long
    Dim answer       As VbMsgBoxResult
    Dim deletedCount As Long
    Dim errContext   As String

    ' ── Step1: ファイル選択 & オープン（キャンセルは正常終了）──
    Set targetWb = OpenTargetWorkbook()
    If targetWb Is Nothing Then Exit Sub

    On Error GoTo Cleanup
    Application.ScreenUpdating = False

    ' ── Step2: シートリンクシートの確認 ──────────────────────────
    errContext = "シートリンクシート確認"
    If Not SheetExists(targetWb, LINK_SHEET_NAME) Then
        MsgBox "「" & LINK_SHEET_NAME & "」シートが見つかりません。" & vbCrLf & _
               "先に SetupTestScenarioSheets を実行してください。", vbExclamation, "シートなし"
        GoTo Cleanup
    End If

    ' ── Step3: シートリンクからパターンシート名を収集 ─────────────
    errContext = "空シート収集"
    Set linkWs = targetWb.Worksheets(LINK_SHEET_NAME)
    lastRow = linkWs.Cells(linkWs.Rows.Count, 2).End(xlUp).Row
    For r = 2 To lastRow
        sheetName = CStr(linkWs.Cells(r, 2).Value)
        If sheetName = "" Then GoTo NextLinkRow

        If SheetExists(targetWb, sheetName) Then
            Set ws = targetWb.Worksheets(sheetName)
            If IsSheetEmpty(ws) Then
                emptySheets.Add sheetName
                linkRows.Add r
            End If
        End If
NextLinkRow:
    Next r

    ' ── Step4: 削除候補がなければ終了 ────────────────────────────
    If emptySheets.Count = 0 Then
        MsgBox "削除対象の空シートはありませんでした。", vbInformation, "対象なし"
        GoTo Cleanup
    End If

    ' ── Step5: 確認ダイアログ ────────────────────────────────────
    errContext = "削除確認"
    For i = 1 To emptySheets.Count
        listMsg = listMsg & "  ・" & emptySheets(i) & vbCrLf
    Next i

    answer = MsgBox("以下の空シート（" & emptySheets.Count & " 件）を削除します。" & vbCrLf & vbCrLf & _
                    listMsg & vbCrLf & _
                    "よろしいですか？", vbYesNo + vbExclamation, "空シートの削除確認")
    If answer = vbNo Then
        WriteLog targetWb, "空シート削除：ユーザーがキャンセル"
        GoTo Cleanup
    End If

    ' ── Step6: シート削除 & シートリンク行削除（逆順で処理）────
    errContext = "シート削除"
    Application.DisplayAlerts = False
    deletedCount = 0

    For i = emptySheets.Count To 1 Step -1
        targetWb.Worksheets(emptySheets(i)).Delete
        linkWs.Rows(linkRows(i)).Delete
        deletedCount = deletedCount + 1
        WriteLog targetWb, "削除：" & emptySheets(i)
    Next i

    RebuildLinkSheetFormulas linkWs

    ' ── Step7: 完了通知 ──────────────────────────────────────────
    WriteLog targetWb, "空シート削除完了：" & deletedCount & " 件。時刻：" & Format(Now, "yyyy/mm/dd hh:mm:ss")
    MsgBox deletedCount & " 件の空シートを削除しました。", vbInformation, "削除完了"

    ActivateLinkSheet targetWb

Cleanup:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    If Err.Number <> 0 Then
        MsgBox "エラー（" & errContext & "）: " & Err.Description, vbCritical, "エラー"
    End If

End Sub

'=============================================================================
' ■ シートが空かどうかを判定
'=============================================================================
Private Function IsSheetEmpty(ws As Worksheet) As Boolean
    IsSheetEmpty = ws.Cells.Find("*", LookIn:=xlValues) Is Nothing
End Function

'=============================================================================
' ■ シートリンクの HYPERLINK 数式を行番号に合わせて再構築
'=============================================================================
Private Sub RebuildLinkSheetFormulas(linkWs As Worksheet)

    Dim r As Long
    Dim lastRow As Long
    lastRow = linkWs.Cells(linkWs.Rows.Count, 2).End(xlUp).Row
    For r = 2 To lastRow
        If linkWs.Cells(r, 2).Value <> "" Then
            linkWs.Cells(r, 3).Formula = _
                "=HYPERLINK(""#'""&B" & r & "&""'!A1"",""⇒""&B" & r & ")"
        End If
    Next r

End Sub
