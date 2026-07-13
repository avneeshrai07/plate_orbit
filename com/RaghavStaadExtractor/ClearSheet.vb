Imports System.Runtime.InteropServices
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel

Public Class ClearSheet

    Private Const ProtectionPassword As String = "2022"

    ' Clear and format Sheet1
    Public Sub ClearAndResetSheet1()
        ClearSpecificSheet(1)
    End Sub

    ' Clear and format Sheet2
    Public Sub ClearAndResetSheet2()
        ClearSpecificSheet(2)
    End Sub

    ' Clear both Sheet1 and Sheet2
    Public Sub ClearAllSheets()
        Dim excelApp As Application = CType(Marshal.GetActiveObject("Excel.Application"), Application)
        Dim wb As Workbook = excelApp.ActiveWorkbook

        If MsgBox("Do you want to proceed with clearing all sheets?", vbOKCancel + vbQuestion, "Confirm Action") <> vbOK Then
            MsgBox("Action canceled by the user.", MsgBoxStyle.Information, "Canceled")
            Return
        End If

        For i As Integer = 1 To Math.Min(2, wb.Sheets.Count)
            Dim ws As Worksheet = CType(wb.Sheets(i), Worksheet)

            ws.Unprotect(ProtectionPassword)
            If i = 1 Then
                ClearSheet1(ws)
            ElseIf i = 2 Then
                ClearSheet2(ws)
            End If
            ws.Protect(ProtectionPassword)
        Next
    End Sub

    ' Core method to clear either Sheet1 or Sheet2
    Private Sub ClearSpecificSheet(sheetIndex As Integer)
        Dim excelApp As Application = CType(Marshal.GetActiveObject("Excel.Application"), Application)
        Dim wb As Workbook = excelApp.ActiveWorkbook
        If sheetIndex < 1 OrElse sheetIndex > wb.Sheets.Count Then
            MsgBox("Invalid sheet index.", MsgBoxStyle.Critical, "Error")
            Return
        End If

        Dim ws As Worksheet = CType(wb.Sheets(sheetIndex), Worksheet)
        ws.Unprotect(ProtectionPassword)

        Dim message As String = If(sheetIndex = 1, "Do you want to proceed with clearing Master Data List?", "Do you want to proceed with clearing Summary List?")
        If MsgBox(message, vbOKCancel + vbQuestion, "Confirm Action") <> vbOK Then
            MsgBox("Action canceled by the user.", MsgBoxStyle.Information, "Canceled")
            ws.Protect(ProtectionPassword)
            Return
        End If

        If sheetIndex = 1 Then
            ClearSheet1(ws)
        ElseIf sheetIndex = 2 Then
            ClearSheet2(ws)
        End If

        ws.Protect(ProtectionPassword)
    End Sub

    ' Actual clearing logic for Sheet1
    Private Sub ClearSheet1(ws As Excel.Worksheet)
        ' Preserve values from key cells
        Dim valAS2 = ws.Range("AS2").Value
        Dim valAS3 = ws.Range("AS3").Value
        Dim valAS4 = ws.Range("AS4").Value
        Dim valAS5 = ws.Range("AS5").Value
        Dim valAS6 = ws.Range("AS6").Value
        Dim valAR7 = ws.Range("AR7").Value
        Dim valAS7 = ws.Range("AS7").Value
        Dim valAW2 = ws.Range("AW2").Value

        ' Clear all cells and reset background color
        ws.Cells.Clear()
        ws.Cells.Interior.Color = System.Drawing.Color.White.ToArgb()

        ' Restore preserved values
        ws.Range("AS2").Value = valAS2
        ws.Range("AS3").Value = valAS3
        ws.Range("AS4").Value = valAS4
        ws.Range("AS5").Value = valAS5
        ws.Range("AS6").Value = valAS6
        ws.Range("AR7").Value = valAR7
        ws.Range("AS7").Value = valAS7
        ws.Range("AW2").Value = valAW2

        ' Lock all cells by default
        ws.Cells.Locked = True

        ' Unlock the specific cells that should remain editable after protection
        ws.Range("AS2").Locked = False
        ws.Range("AS3").Locked = False
        ws.Range("AS4").Locked = False
        ws.Range("AS5").Locked = False
        ws.Range("AS6").Locked = False
        ws.Range("AR7").Locked = False
        ws.Range("AS7").Locked = False
        ws.Range("AW2").Locked = False

        ' Format merged header cell
        With ws.Range("C2", "AM5")
            .Merge()
            .Value = "A Vision Brought to Life by Raghav" & vbCrLf & "PRESS BUTTONS FOR [STAAD OPERATIONS]"
            .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
            .VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
            .Font.Size = 14
            .Font.Bold = True
            .Font.Color = RGB(0, 0, 0)
            .Interior.Color = RGB(0, 176, 240)
            .Borders.LineStyle = Excel.XlLineStyle.xlContinuous
            .Borders.Weight = Excel.XlBorderWeight.xlThin
        End With

        ' Instructional rows
        Dim infoRows As String() = {
            "A1: Extract structural data directly from STAAD model.",
            "A2: Generate a comprehensive Material Summary List.",
            "SAVEAS: Save current output specifically for Beam details.",
            "MPS: Redirect to the MPS (Material Processing Sheet).",
            "CLEAR: Clears the active worksheet only.",
            "CLEAR ALL: Clears all worksheets in the workbook.",
            "DATABASE: Assign Raghav Database to Plate Nova (one-time setup only).",
            "DRAW: Activate Drawing Mode for custom visual outputs.",
            "RESET: Revert to default interface and settings.",
            "CREATED BY: Edit project or job-related metadata (button-enabled).",
            "FRONT: Generate front elevation view (ideal for frames and full 3D models).",
            "LEFT: Generate left side elevation (recommended for side views).",
            "RIGHT: Generate right side elevation (recommended for opposite side views).",
            "TOP: Generate top view (ideal for roof and anchor bolt plans).",
            "PLACEMENT: Enter U for upward text, D for downward, N for no text.",
            "TEXT HEIGHT: Specify desired text height.",
            "TEXT COLOR: Specify text color (0 = multi-color, 100=Enable group mode,101-Enble property+group mode)."
        }

        Dim startRow As Integer = 9
        Dim useBlack As Boolean = True

        For Each rowText As String In infoRows
            Dim cell As Excel.Range = ws.Cells(startRow, 1)
            cell.Value = rowText

            If useBlack Then
                cell.Font.Color = RGB(0, 0, 0)        ' Black
            Else
                cell.Font.Color = RGB(0, 0, 146)      ' Dark Blue
            End If

            useBlack = Not useBlack
            startRow += 1
        Next
    End Sub

    ' Actual clearing logic for Sheet2
    Private Sub ClearSheet2(ws As Worksheet)
        ws.Cells.Clear()
        ws.Cells.Interior.Color = RGB(255, 255, 255)

        ' Lock all cells by default
        ws.Cells.Locked = True

        With ws.Range("A2:F4")
            .Merge()
            .Value = "A Vision Brought to Life by Raghav" & vbCrLf & "PRESS BUTTONS FOR [MATERIAL SUMMARY]"
            .HorizontalAlignment = XlHAlign.xlHAlignCenter
            .VerticalAlignment = XlVAlign.xlVAlignCenter
            .Font.Size = 14
            .Font.Bold = True
            .Font.Color = RGB(0, 0, 0)
            .Interior.Color = RGB(0, 176, 240)
            .Borders.LineStyle = XlLineStyle.xlContinuous
            .Borders.Weight = XlBorderWeight.xlThin
        End With

        Dim infoRows() As String = {
            "A1: Extract structural data directly from STAAD model.",
            "A2: Generate a comprehensive Material Summary List.",
            "SAVEAS: Save current output specifically for Beam details.",
            "MPS: Redirect to the MPS (Material Processing Sheet).",
            "CLEAR: Clears the active worksheet only.",
            "CLEAR ALL: Clears all worksheets in the workbook.",
            "DATABASE: Assign Raghav Database to Plate Nova (one-time setup only).",
            "DRAW: Activate Drawing Mode for custom visual outputs.",
            "RESET: Revert to default interface and settings.",
            "CREATED BY: Edit project or job-related metadata (button-enabled).",
            "FRONT: Generate front elevation view (ideal for frames and full 3D models).",
            "LEFT: Generate left side elevation (recommended for side views).",
            "RIGHT: Generate right side elevation (recommended for opposite side views).",
            "TOP: Generate top view (ideal for roof and anchor bolt plans).",
            "PLACEMENT: Enter U for upward text, D for downward, N for no text.",
            "TEXT HEIGHT: Specify desired text height.",
            "TEXT COLOR: Specify text color (0 = multi-color, recommended)."
        }

        Dim rowIdx As Integer = 9
        For Each text As String In infoRows
            ws.Cells(rowIdx, "A").Value = text
            rowIdx += 1
        Next
    End Sub

End Class