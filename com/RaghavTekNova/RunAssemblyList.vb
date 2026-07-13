Imports System.IO
Imports System.Runtime.InteropServices
Imports Microsoft.Office.Interop.Excel
Imports Microsoft.VisualBasic

Namespace RaghavTekNova

    <ComVisible(True)>
    <Guid("B19A8C6D-5D2E-4F6D-9F24-91F7C1234567")> ' Replace with your own unique GUID
    <ProgId("RaghavTekNova.RunAssemblyList")>
    Public Class RunAssemblyList

        Public Sub New()
            EnsureLicense()                     ' ✅ Existing licensing
            CheckExpirationDateAndBackdate()   ' 🔐 Expiry + backdate tamper check
        End Sub

        Public Sub A2_RUN_PR_ASSEMBLY_LIST_VB()

            Dim excelApp As Application = Nothing
            Dim workbook As Workbook = Nothing
            Dim wsSource As Worksheet = Nothing
            Dim wsTarget As Worksheet = Nothing
            Dim lastRow As Long, targetStartRow As Long, serialNum As Long, targetLastRow As Long
            Dim dataRange As Range = Nothing
            Dim totalSum As Double
            Dim i As Long
            Dim processResult As Boolean = True

            Try
                excelApp = CType(Marshal.GetActiveObject("Excel.Application"), Application)
                workbook = excelApp.ThisWorkbook

                ' === Unlock sheets first ===
                UnlockAllSheets(excelApp, workbook)

                ' === Run SetPrintPreferences ===
                SetPrintPreferences(excelApp)

                ' Set the source and target worksheets
                wsSource = CType(workbook.Sheets("Sheet1"), Worksheet)  ' Sheet1
                wsTarget = CType(workbook.Sheets("Sheet6"), Worksheet)  ' Sheet6

                ' Find the last row in the target sheet (Sheet6)
                lastRow = wsTarget.Cells(wsTarget.Rows.Count, "A").End(XlDirection.xlUp).Row

                ' Ensure there is sufficient data to copy
                If lastRow < 8 Then
                    MsgBox("Not enough data in the target sheet!", MsgBoxStyle.Exclamation)
                    processResult = False
                    Exit Sub
                End If

                ' Define the starting row in Sheet1 for pasting data
                targetStartRow = 8

                ' Clear existing data in the source sheet (Sheet1)
                wsSource.Cells.Clear()

                ' Add project information to the designated cells
                With wsSource
                    .Range("C2").Value = "PROJECT NO:-"
                    .Range("C3").Value = "PROJECT NAME:-"
                    .Range("C4").Value = "CLIENT NAME:-"
                    .Range("C5").Value = "DATE:-"
                    .Range("F5").Value = "REVISION:-"
                    .Range("J3").Value = "TOTAL:-"
                    .Range("J4").Value = "WEIGHT(KG)"
                    .Range("K4").Value = "AREA(SQM)"
                End With

                ' Format the project information text
                With wsSource.Range("A1:K5")
                    .Font.Bold = True
                    .Font.Size = 14
                End With

                ' Set background colors (same as VBA)
                wsSource.Range("A1:B1,C1:G1,H1:I1,J1:K1").Interior.Color = RGB(248, 203, 173)
                wsSource.Range("A7:M7").Interior.Color = RGB(173, 216, 230)
                wsSource.Range("J3").Interior.Color = RGB(226, 239, 218)
                wsSource.Range("J4:K4").Interior.Color = RGB(0, 176, 240)
                wsSource.Range("A6").Interior.Color = RGB(248, 203, 173)
                wsSource.Range("L6,L1").Interior.Color = RGB(248, 203, 173)
                wsSource.Range("M6,M1").Interior.Color = RGB(248, 203, 173)
                wsSource.Range("J5,K5").Interior.Color = RGB(169, 208, 142)

                ' Clear contents before merging to avoid data loss warning
                wsSource.Range("A2:B2,A3:B3,A4:B4,A5:B5").ClearContents()

                ' Merge and format cells for project headers and details
                wsSource.Range("A2:B5").Merge()
                wsSource.Range("D2:I2").Merge()
                wsSource.Range("D3:I3").Merge()
                wsSource.Range("D4:I4").Merge()
                wsSource.Range("J3:K3").Merge()

                With wsSource.Range("A1:K1")
                    .Merge()
                    .Value = "BILL OF MATERIAL (ASSEMBLY LIST)"
                    .HorizontalAlignment = XlHAlign.xlHAlignCenter
                    .VerticalAlignment = XlVAlign.xlVAlignCenter
                    .Font.Bold = True
                    .Font.Size = 15
                End With

                With wsSource.Range("I2:I5,J2:J5,K5")
                    .HorizontalAlignment = XlHAlign.xlHAlignLeft
                    .VerticalAlignment = XlVAlign.xlVAlignCenter
                    .Font.Bold = True
                    .Font.Size = 14
                End With

                ' Merge and leave row 6 blank
                With wsSource.Range("A6:K6")
                    .Merge()
                    .Value = ""
                End With

                ' Write headers in row 7
                wsSource.Range("A7").Value = "SR.NO."
                wsSource.Range("B7").Value = "CATEGORY"
                wsSource.Range("C7").Value = "ASSEMBLY NAME"
                wsSource.Range("D7").Value = "PROFILE"
                wsSource.Range("E7").Value = "ASSEMBLY MARK"
                wsSource.Range("F7").Value = "LENGTH"
                wsSource.Range("G7").Value = "QTY."
                wsSource.Range("H7").Value = "UNIT WEIGHT"
                wsSource.Range("I7").Value = "TOTAL WEIGHT"
                wsSource.Range("J7").Value = "UNIT AREA"
                wsSource.Range("K7").Value = "TOTAL AREA"
                wsSource.Range("L7").Value = "PARTS IN 1 ASSEMBLY"
                wsSource.Range("M7").Value = "DRAWINGS"
                wsSource.Range("N7").Value = "LOCKED"
                wsSource.Range("O7").Value = "GUID"

                ' Format header row
                With wsSource.Range("A7:N7")
                    .Font.Bold = True
                    .HorizontalAlignment = XlHAlign.xlHAlignCenter
                    .VerticalAlignment = XlVAlign.xlVAlignCenter
                    .Font.Size = 12
                End With

                ' Adjust column widths
                wsSource.Columns("A").ColumnWidth = 6    ' SR.NO.
                wsSource.Columns("B").ColumnWidth = 22   ' CATEGORY
                wsSource.Columns("C").ColumnWidth = 25   ' ASSEMBLY NAME
                wsSource.Columns("D").ColumnWidth = 25   ' PROFILE
                wsSource.Columns("E").ColumnWidth = 18   ' ASSEMBLY MARK
                wsSource.Columns("F").ColumnWidth = 15   ' LENGTH
                wsSource.Columns("G").ColumnWidth = 10   ' QTY.
                wsSource.Columns("H").ColumnWidth = 15   ' UNIT WEIGHT
                wsSource.Columns("I").ColumnWidth = 20   ' TOTAL WEIGHT
                wsSource.Columns("J").ColumnWidth = 18   ' UNIT AREA
                wsSource.Columns("K").ColumnWidth = 18   ' TOTAL AREA
                wsSource.Columns("L").ColumnWidth = 25   ' PARTS IN 1 ASSEMBLY
                wsSource.Columns("M").ColumnWidth = 18   ' DRAWINGS
                wsSource.Columns("N").ColumnWidth = 18   ' LOCKED
                wsSource.Columns("O").ColumnWidth = 30   ' GUID

                ' Copy data from the target sheet (Sheet6) to the source sheet (Sheet1)
                dataRange = wsTarget.Range("A8:N" & lastRow)
                dataRange.Copy(wsSource.Cells(targetStartRow, 2)) ' Paste data starting from column B

                ' Add serial numbers in column A of the source sheet (Sheet1)
                For serialNum = 1 To dataRange.Rows.Count
                    wsSource.Cells(targetStartRow + serialNum - 1, 1).Value = serialNum
                Next serialNum

                ' Apply borders to the entire used range in the source sheet
                With wsSource.UsedRange.Borders
                    .LineStyle = XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = XlBorderWeight.xlThin
                End With

                ' Align content from row 7 to column O to the left
                With wsSource.Range("A7:O" & wsSource.Rows.Count)
                    .HorizontalAlignment = XlHAlign.xlHAlignLeft
                    .VerticalAlignment = XlVAlign.xlVAlignCenter
                End With

                ' Adjust row heights for better readability
                wsSource.Rows("6:" & lastRow + targetStartRow - 1).RowHeight = 18

                ' Autofit remaining columns if necessary
                wsSource.Columns("O:Z").AutoFit()

                ' Find the last row with data in the source sheet
                targetLastRow = wsSource.Cells(wsSource.Rows.Count, "A").End(XlDirection.xlUp).Row

                ' Perform SUM operations and display results
                totalSum = excelApp.WorksheetFunction.Sum(wsTarget.Range("H8:H" & lastRow - 1))
                wsSource.Range("J5").Value = totalSum
                totalSum = excelApp.WorksheetFunction.Sum(wsTarget.Range("J8:J" & lastRow - 1))
                wsSource.Range("K5").Value = totalSum

                ' Display values from the target sheet in the source sheet
                wsSource.Range("D2").Value = wsTarget.Range("E1").Value
                wsSource.Range("D3").Value = wsTarget.Range("E2").Value
                wsSource.Range("D4").Value = wsTarget.Range("E3").Value
                wsSource.Range("D5").Value = wsTarget.Range("E4").Value

                ' Clear contents before merging the last row
                wsSource.Range("A" & targetLastRow & ":K" & targetLastRow).ClearContents()

                ' Merge and clear the last row
                With wsSource.Range("A" & targetLastRow & ":K" & targetLastRow)
                    .Merge()
                    .Value = ""
                End With

                ' Remove borders after the last merged line
                With wsSource.Range("A" & targetLastRow + 1 & ":O" & wsSource.Rows.Count)
                    .Borders.LineStyle = XlLineStyle.xlLineStyleNone
                End With

                ' Freeze panes to keep top 7 rows visible
                If excelApp.ActiveWindow IsNot Nothing Then
                    excelApp.ActiveWindow.FreezePanes = True
                End If

                ' Hide columns L to O
                wsSource.Columns("L:O").Hidden = True

                ' === Lock sheets back ===
                LockAllSheets(excelApp, workbook)

                ' === Restore Excel UI for proper operation ===
                excelApp.ScreenUpdating = True
                excelApp.Calculation = -4105     ' xlCalculationAutomatic
                excelApp.EnableEvents = True
                excelApp.DisplayAlerts = True

                ' === Force Sheet1 as active to avoid wrong window context ===
                workbook.Sheets("Sheet1").Select() ' ✅ Stay on Sheet1
                Dim win As Object = excelApp.ActiveWindow

                ' === Clean up UI ===
                win.DisplayGridlines = False               ' Hide gridlines (Sheet1)
                win.DisplayHeadings = False                ' Hide row/column labels
                excelApp.DisplayFormulaBar = True         ' Show formula bar
                win.DisplayWorkbookTabs = True            ' Hide sheet tabs

                ' Optional: Hide scrollbars
                'win.DisplayHorizontalScrollBar = False
                'win.DisplayVerticalScrollBar = False

                ' === Hide ribbon safely ===
                Try
                    excelApp.ExecuteExcel4Macro("SHOW.TOOLBAR(""Ribbon"", False)")
                Catch ex As Exception
                    ' Ribbon hiding might fail in some Excel versions, continue silently
                End Try

                ' === Position the view (scroll and zoom only) ===
                win.ScrollRow = 1
                win.ScrollColumn = 1
                win.Zoom = 100

                ' === Set DisplayGridlines = False for each sheet WITHOUT activating ===
                For i = 1 To workbook.Windows.Count
                    Dim wnd As Object = workbook.Windows(i)
                    If wnd IsNot Nothing Then
                        Try
                            wnd.DisplayGridlines = False
                        Catch ex As Exception
                            ' Continue if window access fails
                        End Try
                    End If
                Next

                ' === Auto Save workbook at the end ===
                AutoSaveWorkbook(workbook)

                MsgBox("Assembly list formatted successfully!", MsgBoxStyle.Information)

            Catch ex As Exception
                MsgBox("Error: " & ex.Message, MsgBoxStyle.Critical)
                processResult = False
            Finally
                ' Clean up COM objects
                If dataRange IsNot Nothing Then
                    Marshal.ReleaseComObject(dataRange)
                End If
                If wsTarget IsNot Nothing Then
                    Marshal.ReleaseComObject(wsTarget)
                End If
                If wsSource IsNot Nothing Then
                    Marshal.ReleaseComObject(wsSource)
                End If
                If workbook IsNot Nothing Then
                    Marshal.ReleaseComObject(workbook)
                End If
                If excelApp IsNot Nothing Then
                    Marshal.ReleaseComObject(excelApp)
                End If
            End Try
        End Sub

        ' === AutoSave Workbook Method ===
        Private Sub AutoSaveWorkbook(workbook As Workbook)
            Try
                If workbook IsNot Nothing Then
                    workbook.Save()
                End If
            Catch ex As Exception
                ' Handle save errors silently or log them
            End Try
        End Sub

        ' === Helper: Unlock all sheets ===
        Private Sub UnlockAllSheets(app As Application, wb As Workbook)
            Dim currentSheet As Worksheet = Nothing
            Dim calcMode As XlCalculation = app.Calculation

            Try
                currentSheet = app.ActiveSheet
                app.ScreenUpdating = False
                app.EnableEvents = False
                app.Calculation = XlCalculation.xlCalculationManual

                For Each ws As Worksheet In wb.Worksheets
                    Try
                        ws.Unprotect("2022")
                    Catch ex As Exception
                        ' Continue if sheet is already unprotected or password is wrong
                    End Try
                Next

                If currentSheet IsNot Nothing Then
                    currentSheet.Activate()
                End If
            Catch ex As Exception
                ' Handle any errors during unlock
            Finally
                app.Calculation = calcMode
                app.EnableEvents = True
                app.ScreenUpdating = True

                If currentSheet IsNot Nothing Then
                    Marshal.ReleaseComObject(currentSheet)
                End If
            End Try
        End Sub

        Private Sub EnsureLicense()
            Dim validator As New ExecutionValidation()
            If Not validator.IsLicenceValid() Then
                Throw New UnauthorizedAccessException("License check failed.")
            End If
        End Sub

        Private Sub CheckExpirationDateAndBackdate()
            Dim expiryDate As New DateTime(2027, 4, 9)
            Dim today As DateTime = DateTime.Now.Date

            ' 🔒 Path to store last run date (user-specific and hidden)
            Dim filePath As String = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "RaghavStaad.last")

            ' Check if file exists
            If File.Exists(filePath) Then
                Dim lastRunText As String = File.ReadAllText(filePath).Trim()
                Dim lastRunDate As DateTime
                If DateTime.TryParse(lastRunText, lastRunDate) Then

                    ' ⛔ Backdate tampering detected
                    If today < lastRunDate Then
                        Throw New UnauthorizedAccessException("System date has been tampered. Please correct your clock.")
                    End If
                End If
            End If

            ' Save today's date as last run (overwrite or create)
            File.WriteAllText(filePath, today.ToString("yyyy-MM-dd"))

            ' 🔐 Expiry check
            If today > expiryDate Then
                Throw New UnauthorizedAccessException("Software expired. Please contact support for an updated version.")
            End If
        End Sub


        ' === SetPrintPreferences method ===
        Private Sub SetPrintPreferences(app As Application)
            Dim lastRow As Long
            Dim ws As Worksheet = Nothing

            Try
                ' Set ws to the active sheet
                ws = app.ActiveSheet

                ' Find the last used row in the active worksheet
                lastRow = ws.Cells(ws.Rows.Count, "A").End(XlDirection.xlUp).Row

                ' Set the print area to include from row 1 to the last filled row + 1
                ws.PageSetup.PrintArea = "$A$1:$O$" & lastRow + 1

                ' Custom Page Setup
                With ws.PageSetup
                    .Orientation = XlPageOrientation.xlLandscape        ' Set landscape orientation
                    .Zoom = False                                       ' Disable zoom
                    .FitToPagesWide = 1                                ' Fit to 1 page wide
                    .FitToPagesTall = False                            ' Fit to as many pages tall as needed
                    ' Margins (in inches)
                    .TopMargin = app.InchesToPoints(0.2)               ' Top margin: 0.2 inch
                    .BottomMargin = app.InchesToPoints(0.1)            ' Bottom margin: 0.1 inch
                    .LeftMargin = app.InchesToPoints(0.2)              ' Left margin: 0.2 inch
                    .RightMargin = app.InchesToPoints(0.2)             ' Right margin: 0.2 inch
                    .HeaderMargin = app.InchesToPoints(0)              ' Header margin: 0 inch
                    .FooterMargin = app.InchesToPoints(0)              ' Footer margin: 0 inch
                    ' Header and Footer
                    .CenterHorizontally = True                         ' Center horizontally
                    .CenterVertically = False                          ' Center vertically
                    .PrintTitleRows = "$6:$7"                          ' Repeat rows 6 to 7 on each page
                    .PaperSize = XlPaperSize.xlPaperA4                 ' Set paper size to A4
                    ' Custom Header
                    .CenterHeader = ""                                 ' Centered header empty
                    .LeftHeader = ""                                   ' Left header empty
                    .RightHeader = ""                                  ' Right header empty
                    ' Custom Footer
                    .LeftFooter = "Page &P of &N"                      ' Left footer with page number
                    .CenterFooter = ""                                 ' Center footer empty
                    .RightFooter = "Raghav"                            ' Right footer with Name
                End With

            Catch ex As Exception
                ' Handle any errors during print setup
            Finally
                If ws IsNot Nothing Then
                    Marshal.ReleaseComObject(ws)
                End If
            End Try
        End Sub

        ' === Helper: Lock all sheets ===
        Private Sub LockAllSheets(app As Application, wb As Workbook)
            Dim currentSheet As Worksheet = Nothing
            Dim calcMode As XlCalculation = app.Calculation

            Try
                currentSheet = app.ActiveSheet
                app.ScreenUpdating = False
                app.EnableEvents = False
                app.Calculation = XlCalculation.xlCalculationManual

                For Each ws As Worksheet In wb.Worksheets
                    Try
                        ws.Protect("2022", UserInterfaceOnly:=True)
                    Catch ex As Exception
                        ' Continue if protection fails
                    End Try
                Next

                If currentSheet IsNot Nothing Then
                    currentSheet.Activate()
                End If
            Catch ex As Exception
                ' Handle any errors during lock
            Finally
                app.Calculation = calcMode
                app.EnableEvents = True
                app.ScreenUpdating = True

                If currentSheet IsNot Nothing Then
                    Marshal.ReleaseComObject(currentSheet)
                End If
            End Try
        End Sub

        Private Function RGB(ByVal Red As Integer, ByVal Green As Integer, ByVal Blue As Integer) As Long
            Return Red + (Green * 256) + (Blue * 65536)
        End Function
    End Class

End Namespace