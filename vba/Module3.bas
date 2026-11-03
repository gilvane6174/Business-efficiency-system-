Attribute VB_Name = "Module3"
Option Explicit

Sub 手動入力内容クリア()
    Dim ws As Worksheet
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False

    ' ダッシュボードシート名。必要なら変更してください。
    Set ws = ThisWorkbook.Worksheets("ダッシュボード")

    With ws
        ' C・E・G列は40～70行、I～L列は結合想定でまとめてクリア
        Union(.Range("C40:C70"), _
              .Range("E40:E70"), _
              .Range("G40:G70"), _
              .Range("I40:L70")).ClearContents
    End With

CleanExit:
    Application.ScreenUpdating = True
    Exit Sub
ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "クリアに失敗しました。シート名・シート保護を確認してください。", vbExclamation
End Sub

Sub 本日のセルへ移動()
    Dim wb As Workbook
    Dim wsDash As Worksheet
    Dim wsTarget As Worksheet
    Dim tgtName As String

    On Error GoTo ErrHandler
    Application.ScreenUpdating = False

    Set wb = ThisWorkbook
    Set wsDash = wb.Worksheets("ダッシュボード")  ' ←必要ならシート名調整

    tgtName = Trim$(CStr(wsDash.Range("B11").Value))
    If Len(tgtName) = 0 Then
        With wsDash.Range("D2")
            .Value = " ※ B11 に今月のシート名がありません。確認してください。"
            .Font.Color = vbRed
        End With
        GoTo CleanExit
    End If

    On Error Resume Next
    Set wsTarget = wb.Worksheets(tgtName)
    On Error GoTo ErrHandler
    If wsTarget Is Nothing Then
        With wsDash.Range("D2")
            .Value = " ※ '" & tgtName & "' シートが見つかりませんでした。シートを新規作成か、確認をしてください。"
            .Font.Color = vbRed
        End With
        GoTo CleanExit
    End If

    If wsTarget.Visible <> xlSheetVisible Then wsTarget.Visible = xlSheetVisible
    Application.Goto wsTarget.Range("A4"), True
    Exit Sub

CleanExit:
    Application.ScreenUpdating = True
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "移動に失敗しました。B11とシート名をご確認ください。", vbExclamation
End Sub

