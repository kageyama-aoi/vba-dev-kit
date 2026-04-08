# VBAツール標準テンプレート

## 目的

Excelを使ったツールを作成する。

## 設計ルール

以下の設計を必ず守る。

* Main処理は処理の流れだけを書く
* 実処理は関数に分割する
* Excelセルアクセス回数を最小化する
* 配列で処理して最後にExcelへ書き込む
* エラー処理を入れる
* パフォーマンス設定（ScreenUpdating等）をセットで管理する

---

## モジュールヘッダーテンプレート

新規モジュール作成時は必ずこのヘッダーをファイル先頭に入れること。

```vba
'=============================================================================
' モジュール名  : （モジュール名）
' 作成目的      : （このモジュールが何をするか1〜2行で）
'
' 【前提】──────────────────────────────────────────────────────
'   ・（動作に必要な前提条件）
'
' 【公開マクロ一覧】────────────────────────────────────────────
'
'  ■メイン
'   MainProcName   … （説明）
'
' 【内部ヘルパー（直接実行は非推奨）】
'   HelperName     … （説明）
'
' 【変更履歴】──────────────────────────────────────────────────
'   v1.0  初版作成
'=============================================================================
```

---

## コード構造テンプレート

```vba
Option Explicit

' 定数はモジュール先頭に集約
Private Const OUTPUT_SHEET As String = "結果"

'=============================================================================
' メイン処理
'=============================================================================
Sub Main()

    '--- 変数宣言（必ずサブルーチン先頭にまとめる）---
    Dim errContext As String
    Dim data       As Variant
    Dim result     As Variant

    errContext = "初期化"

    '--- パフォーマンス設定（On Error とセットで必ず入れる）---
    On Error GoTo Cleanup
    Application.ScreenUpdating = False
    Application.Calculation    = xlCalculationManual
    Application.EnableEvents   = False

    '1 初期処理
    errContext = "初期処理"
    Application.StatusBar = "初期処理中..."
    Call Initialize

    '2 データ取得
    errContext = "データ取得"
    Application.StatusBar = "データ取得中..."
    data = GetData()

    '3 データ処理
    errContext = "データ処理"
    Application.StatusBar = "処理中..."
    result = ProcessData(data)

    '4 Excel出力
    errContext = "Excel出力"
    Application.StatusBar = "シートへ書き込み中..."
    Call WriteResult(result)

    MsgBox "処理が完了しました"

Cleanup:
    '--- 必ず復元する（エラー時も正常時も通る）---
    Application.ScreenUpdating = True
    Application.Calculation    = xlCalculationAutomatic
    Application.EnableEvents   = True
    Application.StatusBar      = False

    If Err.Number <> 0 Then
        MsgBox "エラー（" & errContext & "）: " & Err.Description, vbExclamation
    End If

End Sub


'=============================================================================
' 初期処理
'=============================================================================
Sub Initialize()

End Sub


'=============================================================================
' データ取得
' 戻り値 : 取得データ（2D Variant 配列）
'=============================================================================
Function GetData() As Variant

End Function


'=============================================================================
' データ処理
' 引数   : data … GetData の戻り値
' 戻り値 : 処理済みデータ
'=============================================================================
Function ProcessData(data As Variant) As Variant

End Function


'=============================================================================
' Excel出力
' ・result を2D配列として一括書き込みする（セル単体アクセス禁止）
'=============================================================================
Sub WriteResult(result As Variant)

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(OUTPUT_SHEET)

    ' 一括書き込み（ループしない）
    ws.Range("A2").Resize(UBound(result, 1), UBound(result, 2)).Value = result

End Sub


'=============================================================================
' Utility：最終行取得
'=============================================================================
Function GetLastRow(ws As Worksheet, col As Long) As Long

    GetLastRow = ws.Cells(ws.Rows.Count, col).End(xlUp).Row

End Function
```

---

## 関数コメントテンプレート

各関数の直上に必ずこの形式でコメントを入れること。

```vba
'-----------------------------------------------------------------------------
' （この関数が何をするか）
' 引数 : argName … （説明）
' 戻り値 : （戻り値の説明。Sub の場合は省略）
'-----------------------------------------------------------------------------
```
