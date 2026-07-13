Imports System.Runtime.InteropServices
Imports Microsoft.Office.Interop.Excel

Namespace RaghavStaadExtractor

    <ComVisible(True)>
    <Guid("D2A2B1E4-44C2-41C3-9A54-9A48F77D1234")> ' Replace with your own GUID
    <ProgId("RaghavStaadExtractor.RaghavDatabase")>
    Public Class RaghavDatabase

        Private Const pwd As String = "2022"

        ' === Unlock sheets for data review ===
        <ComVisible(True)>
        Public Sub DataRewiewUnlocked()
            Dim app As Application = Nothing
            Try
                app = CType(Marshal.GetActiveObject("Excel.Application"), Application)
                Dim currentSheet As Worksheet = app.ActiveSheet
                Dim wb As Workbook = app.ActiveWorkbook

                app.ScreenUpdating = False
                app.EnableEvents = False
                app.Calculation = XlCalculation.xlCalculationManual

                wb.Unprotect(pwd)

                For i As Integer = 2 To wb.Sheets.Count
                    Dim ws As Worksheet = CType(wb.Sheets(i), Worksheet)
                    ws.Unprotect(pwd)
                    ws.Visible = XlSheetVisibility.xlSheetVisible
                    ws.Cells.Locked = True
                    ws.Protect(Password:=pwd,
                               AllowFormattingCells:=False,
                               AllowDeletingRows:=False,
                               AllowInsertingRows:=False,
                               AllowSorting:=False,
                               AllowFiltering:=False)
                Next

                wb.Protect(Password:=pwd, Structure:=True)
                currentSheet.Activate()

                app.ScreenUpdating = True
                MsgBox("Step Inside: The Data Room is Now Open.", vbInformation)
            Catch ex As Exception
                MsgBox("Error in DataRewiewUnlocked: " & ex.Message, vbCritical)
            Finally
                If app IsNot Nothing Then
                    app.ScreenUpdating = True
                    app.EnableEvents = True
                    app.Calculation = XlCalculation.xlCalculationAutomatic
                End If
            End Try
        End Sub

        ' === Generate Unit Weight Database ===
        <ComVisible(True)>
        Public Sub LetsGenrateUnitWt()
            Dim app As Application = Nothing
            Try
                app = CType(Marshal.GetActiveObject("Excel.Application"), Application)
                Dim wb As Workbook = app.ActiveWorkbook
                Dim wsDest As Worksheet = CType(wb.Sheets("Sheet1"), Worksheet)
                Dim wsSrc As Worksheet
                Dim pasteRow As Long = 6
                Dim serialNum As Long = 1

                app.ScreenUpdating = False
                app.EnableEvents = False
                app.Calculation = XlCalculation.xlCalculationManual

                wb.Unprotect(pwd)
                For i As Integer = 1 To wb.Sheets.Count
                    Dim ws As Worksheet = CType(wb.Sheets(i), Worksheet)
                    ws.Visible = XlSheetVisibility.xlSheetVisible
                    ws.Unprotect(pwd)
                Next

                wsDest.Cells.Clear()

                With wsDest.Range("A1:E1")
                    .Merge()
                    .Value = "RAGHAV DATABASE"
                    .HorizontalAlignment = XlHAlign.xlHAlignCenter
                    .VerticalAlignment = XlVAlign.xlVAlignCenter
                    .Font.Bold = True
                    .Font.Size = 14
                    .Interior.Color = RGB(247, 199, 172)
                End With

                wsDest.Range("A2:E3").Merge()
                wsDest.Range("A2:E3").Value = ""
                wsDest.Range("A4:E4").Merge()
                wsDest.Range("A4:E4").Interior.Color = RGB(247, 199, 172)

                With wsDest.Range("A5:E5")
                    .Value = {"S.No", "Name", "Staad Name", "AREA Ax(cm2)", "Kg/M"}
                    .Font.Bold = True
                    .Font.Color = RGB(0, 0, 0)
                    .Interior.Color = RGB(173, 216, 230)
                End With

                wsDest.Range("A1:E4").Borders.LineStyle = XlLineStyle.xlContinuous

                For i As Integer = 2 To wb.Sheets.Count
                    wsSrc = CType(wb.Sheets(i), Worksheet)
                    Dim lastRowSrc As Long = wsSrc.Cells(wsSrc.Rows.Count, "B").End(XlDirection.xlUp).Row

                    If lastRowSrc >= 3 Then
                        Dim rngToCopy As Range = wsSrc.Range("B3:D" & lastRowSrc)
                        wsDest.Cells(pasteRow, "B").Resize(rngToCopy.Rows.Count, 3).Value = rngToCopy.Value

                        For r As Long = pasteRow To pasteRow + rngToCopy.Rows.Count - 1
                            wsDest.Cells(r, "A").Value = serialNum
                            If IsNumeric(wsDest.Cells(r, "D").Value) Then
                                wsDest.Cells(r, "E").Value = Math.Round(CDbl(wsDest.Cells(r, "D").Value) * 0.785, 3)
                            Else
                                wsDest.Cells(r, "E").Value = "N/A"
                            End If
                            serialNum += 1
                        Next

                        pasteRow = wsDest.Cells(wsDest.Rows.Count, "B").End(XlDirection.xlUp).Row + 1
                    End If
                Next

                Dim finalRow As Long = wsDest.Cells(wsDest.Rows.Count, "A").End(XlDirection.xlUp).Row
                wsDest.Range("A6:E" & finalRow).Borders.LineStyle = XlLineStyle.xlContinuous
                wsDest.Columns("A:E").AutoFit()

                For i As Integer = 2 To wb.Sheets.Count
                    Dim ws As Worksheet = CType(wb.Sheets(i), Worksheet)
                    ws.Visible = XlSheetVisibility.xlSheetVeryHidden
                    ws.Protect(Password:=pwd,
                               AllowFormattingCells:=False,
                               AllowDeletingRows:=False,
                               AllowInsertingRows:=False,
                               AllowSorting:=False,
                               AllowFiltering:=False)
                Next

                wsDest.Unprotect(pwd)
                wsDest.Cells.Locked = False
                wb.Protect(pwd, Structure:=True)

                app.ScreenUpdating = True
                MsgBox("All set! PlateNova can now use the Raghav Database.", vbInformation)

            Catch ex As Exception
                MsgBox("Error in LetsGenrateUnitWt: " & ex.Message, vbCritical)
            Finally
                If app IsNot Nothing Then
                    app.ScreenUpdating = True
                    app.EnableEvents = True
                    app.Calculation = XlCalculation.xlCalculationAutomatic
                End If
            End Try
        End Sub

        ' === Unprotect sheets but allow adding rows ===
        <ComVisible(True)>
        Public Sub UnprotectSheetsAllowAdd()
            Dim app As Application = Nothing
            Try
                app = CType(Marshal.GetActiveObject("Excel.Application"), Application)
                Dim wb As Workbook = app.ActiveWorkbook
                Dim currentSheet As Worksheet = app.ActiveSheet

                app.ScreenUpdating = False
                app.EnableEvents = False
                app.Calculation = XlCalculation.xlCalculationManual

                wb.Unprotect(pwd)

                For Each ws As Worksheet In wb.Worksheets
                    ws.Visible = XlSheetVisibility.xlSheetVisible
                    ws.Unprotect(pwd)
                    ws.Protect(Password:=pwd,
                               AllowFormattingCells:=True,
                               AllowInsertingRows:=True,
                               AllowDeletingRows:=True,
                               AllowSorting:=True,
                               AllowFiltering:=True)
                Next

                currentSheet.Activate()
                app.ScreenUpdating = True
                MsgBox("Blank pages await your brilliance — go ahead, add those sheets and let ideas flow!", vbInformation)
            Catch ex As Exception
                MsgBox("Error in UnprotectSheetsAllowAdd: " & ex.Message, vbCritical)
            Finally
                If app IsNot Nothing Then
                    app.ScreenUpdating = True
                    app.EnableEvents = True
                    app.Calculation = XlCalculation.xlCalculationAutomatic
                End If
            End Try
        End Sub

    End Class
End Namespace