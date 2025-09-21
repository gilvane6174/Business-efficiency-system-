# VBA Macros Included in the Business Efficiency System

This page documents the VBA macros embedded in the Excel system.  
They provide convenient functions such as clearing manual inputs and jumping directly to today’s sheet.

---

## Macro 1: Clear Manual Inputs

```vba
Option Explicit

Sub 手動入力内容クリア()
    Dim ws As Worksheet
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False

    ' Dashboard sheet name. Change if necessary.
    Set ws = ThisWorkbook.Worksheets("ダッシュボード")

    With ws
        ' Clear C, E, G columns (rows 40–70), and merged I–L columns
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
    MsgBox "Clear operation failed. Please check sheet name or protection settings.", vbExclamation
End Sub
```

### Purpose
- Clears specific ranges in the dashboard sheet with a single click.  
- Designed for recurring manual inputs that need resetting each month.  

---

## Macro 2: Jump to Today’s Sheet

```vba
Option Explicit

Sub 本日のセルへ移動()
    Dim wb As Workbook
    Dim wsDash As Worksheet
    Dim wsTarget As Worksheet
    Dim tgtName As String

    On Error GoTo ErrHandler
    Application.ScreenUpdating = False

    Set wb = ThisWorkbook
    Set wsDash = wb.Worksheets("ダッシュボード")  ' Adjust sheet name if necessary

    tgtName = Trim$(CStr(wsDash.Range("B11").Value))
    If Len(tgtName) = 0 Then
        With wsDash.Range("D2")
            .Value = " ※ B11 does not contain this month’s sheet name. Please check."
            .Font.Color = vbRed
        End With
        GoTo CleanExit
    End If

    On Error Resume Next
    Set wsTarget = wb.Worksheets(tgtName)
    On Error GoTo ErrHandler
    If wsTarget Is Nothing Then
        With wsDash.Range("D2")
            .Value = " ※ '" & tgtName & "' sheet not found. Please create or confirm the sheet."
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
    MsgBox "Failed to move. Please check B11 and sheet name.", vbExclamation
End Sub
```

### Purpose
- Reads the current month’s sheet name from cell **B11** in the dashboard.  
- Automatically navigates to that sheet and selects cell **A4**.  
- Provides clear error messages if the sheet is missing or misnamed.  

---

## Notes
- Both macros require **VBA to be enabled** in Excel.  
- Sheet names can be adjusted as needed.  
- Error handling ensures the workbook remains stable even if misconfigurations occur.  

---
