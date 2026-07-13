Imports System
Imports System.Runtime.InteropServices
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Windows.Forms
Imports System.Collections.Generic
Imports Microsoft.VisualBasic

<ComVisible(True)>
<ClassInterface(ClassInterfaceType.AutoDual)>
<ProgId("RaghavTekNova.ClearSheets")>
Public Class ClearSheets
    Implements IDisposable

    Private xlApp As Excel.Application = Nothing
    Private isAppCreatedInternally As Boolean = False
    Private Const SHEET_PASSWORD As String = "2022"

#Region "Color Constants"
    ' Define colors using RGB function for better readability
    Private Shared ReadOnly LIGHT_GRAY As Integer = RGB(192, 192, 192)  ' Was 12611584
    Private Shared ReadOnly WHITE As Integer = RGB(255, 255, 255)       ' Was 16777215
    Private Shared ReadOnly BLACK As Integer = RGB(0, 0, 0)
#End Region

#Region "Initialization and Cleanup"

    Public Sub New()
    End Sub

    Private Sub InitializeExcelApplication()
        If xlApp Is Nothing Then
            Try
                xlApp = CType(Marshal.GetActiveObject("Excel.Application"), Excel.Application)
            Catch ex As Exception
                Try
                    xlApp = New Excel.Application()
                    isAppCreatedInternally = True
                    xlApp.Visible = True
                Catch ex2 As Exception
                    MessageBox.Show("Failed to initialize Excel application: " & ex2.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
            End Try
        End If
    End Sub

    Public Sub Dispose() Implements IDisposable.Dispose
        If xlApp IsNot Nothing Then
            Try
                If isAppCreatedInternally Then
                    xlApp.Quit()
                End If
                Marshal.ReleaseComObject(xlApp)
            Catch
            Finally
                xlApp = Nothing
            End Try
        End If
    End Sub

#End Region

#Region "Public Methods (Exposed to COM)"

    <ComVisible(True)>
    Public Sub ClearAllSheets()
        InitializeExcelApplication()
        If xlApp Is Nothing OrElse xlApp.ActiveWorkbook Is Nothing Then
            MessageBox.Show("Excel is not available or no workbook is open.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        Dim originalWb As Excel.Workbook = Nothing
        Dim ws As Excel.Worksheet = Nothing
        Dim currentSheet As Excel.Worksheet = Nothing
        Dim calcMode As Excel.XlCalculation

        Try
            originalWb = xlApp.ActiveWorkbook
            Dim response As DialogResult = MessageBox.Show("Do you want to proceed with clearing data?", "Confirm Action", MessageBoxButtons.OKCancel, MessageBoxIcon.Question)
            If response <> DialogResult.OK Then
                MessageBox.Show("Action canceled by the user.", "Canceled", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            calcMode = xlApp.Calculation
            xlApp.ScreenUpdating = False
            xlApp.EnableEvents = False
            xlApp.Calculation = Excel.XlCalculation.xlCalculationManual
            xlApp.DisplayAlerts = False

            ' Store current active sheet
            currentSheet = CType(xlApp.ActiveSheet, Excel.Worksheet)

            ' Unlock all sheets
            UnlockAllSheetsInternal(originalWb)

            ' Activate a temporary sheet (not Sheet1) to allow clearing Sheet1
            For Each ws In originalWb.Sheets
                If ws.Index <> 1 Then
                    ws.Activate()
                    Exit For
                End If
            Next

            ' Clear Sheets 1 to 11
            Dim wsCount As Integer = originalWb.Sheets.Count
            For i As Integer = 1 To Math.Min(11, wsCount)
                ws = SafeGetWorksheet(originalWb, i)
                If ws IsNot Nothing Then
                    ws.Cells.Clear()
                    SafeReleaseComObject(ws)
                End If
            Next

            ' Set content for Sheets 1, 2, 3, 4, 5, 10
            Dim targetSheets As Integer() = {1, 2, 3, 4, 5, 10}
            For Each i As Integer In targetSheets
                If i <= wsCount Then
                    ws = SafeGetWorksheet(originalWb, i)
                    If ws IsNot Nothing Then
                        FormatSheet(ws, "PEEYUSH RAGHAV'S CREATION")
                        SafeReleaseComObject(ws)
                    End If
                End If
            Next

            ' Set values in C1 for specific sheets
            For Each ws In originalWb.Worksheets
                Select Case ws.Name
                    Case "ASSEMBLY LIST"
                        ws.Range("C1").Value = "ASSEMBLY LIST"
                    Case "BOLT LIST"
                        ws.Range("C1").Value = "BOLT LIST"
                    Case "BUILDING WEIGHT SUMMARY"
                        ws.Range("C1").Value = "BUILDING WEIGHT SUMMARY"
                End Select
                SafeReleaseComObject(ws)
            Next

            ' Relock all sheets
            LockAllSheetsInternal(originalWb)

            ' Reactivate Sheet1
            originalWb.Sheets(1).Activate()

            MessageBox.Show("All sheets cleared successfully.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            MessageBox.Show($"Error clearing all sheets: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            xlApp.ScreenUpdating = True
            xlApp.EnableEvents = True
            xlApp.Calculation = calcMode
            xlApp.DisplayAlerts = True
            SafeReleaseComObject(currentSheet)
            SafeReleaseComObject(ws)
            SafeReleaseComObject(originalWb)
        End Try
    End Sub

    <ComVisible(True)>
    Public Sub A8_Clear_Sheet1()
        ' TODO: Implement CheckExpiration_And_Delete if needed
        ClearIndividualSheet(1, "Assembly List", "PRESS BUTTONS FOR [ASSEMBLY LIST]", {
            "A1:-ASSEMBLY LIST ALIGNMENT", "A2:-ASSEMBLY LIST CREATION",
            "A3:-REMOVE DUPLICATE & CHECK DRAWING STATUS", "(PRESS 2 TIMES FOR FINAL)",
            "A3R:-SEPARATE REMOVE DUPLICATE", "A3D:-SEPARATE CHECK DRAWING STATUS",
            "A4:-REMOVE ASSEMBLY PROFILE", "GUID:-GENERATE GLOBAL UNIQUE ID FOR SEARCH ASSEMBLY IN MODEL",
            "(GENERATE GUID BEFORE PRESS A3)", "(IF PRESSED ALREADY THEN PRESS A2 BEFORE GUID)",
            "SaveAs:-GENERATE FILE IN BOQ OR SAVE AS BREAKUP FOLDER",
            "CLEAR:-FOR CLEAR INDIVIDUAL SHEET", "CLEAR ALL:-FOR CLEAR ALL SHEETS"})
    End Sub

    <ComVisible(True)>
    Public Sub B8_Clear_Sheet2()
        ' TODO: Implement CheckExpiration_And_Delete if needed
        ClearIndividualSheet(2, "Bolt List", "PRESS BUTTONS FOR [BOLT LIST]", {
            "B1:-BOLT LIST ALIGNMENT", "B2:-ASSIGNED BOLT ALIGNMENT",
            "B3:-BOLT LIST CREATION", "B4:-GET BOLT WEIGHT", "B5:-ARRANGE BOLT DATA",
            "SaveAs:-GENERATE FILE IN BOQ OR SAVE AS BREAKUP FOLDER",
            "CLEAR:-FOR CLEAR INDIVIDUAL SHEET", "CLEAR ALL:-FOR CLEAR ALL SHEETS"})
    End Sub

    <ComVisible(True)>
    Public Sub S3_Clear_Sheet3()
        ' TODO: Implement CheckExpiration_And_Delete if needed
        ClearIndividualSheet(3, "Boq Summary List", "PRESS BUTTONS FOR [BOQ SUMMARY LIST]", {
            "FINAL:-GET BOQ SUMMARY LIST",
            "SaveAs:-GENERATE FILE IN BOQ OR SAVE AS BREAKUP FOLDER",
            "CLEAR:-FOR CLEAR INDIVIDUAL SHEET", "CLEAR ALL:-FOR CLEAR ALL SHEETS",
            "NOTE:-PRESS FINAL ONLY AFTER SPLIT & BOLT LIST (ITS SUMMARY OF WEIGHT)"})
    End Sub

    <ComVisible(True)>
    Public Sub P7_Clear_Sheet4()
        ' TODO: Implement CheckExpiration_And_Delete if needed
        ClearIndividualSheet(4, "Part List", "PRESS BUTTONS FOR [PART LIST]", {
            "P1:-PART LIST ALIGNMENT", "P2:-PART LIST CREATION",
            "P3:-MATERIAL BREAKUP GENERATION (IN SHEET5)",
            "SaveAs:-GENERATE FILE IN BOQ OR SAVE AS BREAKUP FOLDER",
            "CLEAR:-FOR CLEAR INDIVIDUAL SHEET", "CLEAR ALL:-FOR CLEAR ALL SHEETS"})
    End Sub

    <ComVisible(True)>
    Public Sub P8_Clear_Sheet5()
        ' TODO: Implement CheckExpiration_And_Delete if needed
        ClearIndividualSheet(5, "Material Summary List", "[MATERIAL SUMMARY LIST]", {
            "SUMMARY:-GET MOST REFINED MATERIAL SUMMARY",
            "SaveAs:-GENERATE FILE IN BOQ OR SAVE AS BREAKUP FOLDER",
            "CLEAR:-FOR CLEAR INDIVIDUAL SHEET", "CLEAR ALL:-FOR CLEAR ALL SHEETS"})
    End Sub

    <ComVisible(True)>
    Public Sub V4_Clear_Sheet10()
        ' TODO: Implement CheckExpiration_And_Delete if needed
        ClearIndividualSheet(10, "Drawing Tracking List", "PRESS BUTTONS FOR [DRAWING TRACKING LIST]", {
            "V1:-TRACKING LIST ALIGNMENT", "V2:-ASSIGNED BOLT ALIGNMENT",
            "V3:-BOLT LIST CREATION", "B4:-GET BOLT WEIGHT", "B5:-ARRANGE BOLT DATA",
            "SaveAs:-GENERATE FILE IN BOQ OR SAVE AS BREAKUP FOLDER",
            "CLEAR:-FOR CLEAR INDIVIDUAL SHEET", "CLEAR ALL:-FOR CLEAR ALL SHEETS"})
    End Sub

#End Region

#Region "Private Helper Methods"

    Private Sub ClearIndividualSheet(sheetIndex As Integer, sheetDescription As String, headerText As String, infoRows As String())
        InitializeExcelApplication()
        If xlApp Is Nothing OrElse xlApp.ActiveWorkbook Is Nothing Then
            MessageBox.Show("Excel is not available or no workbook is open.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        Dim originalWb As Excel.Workbook = Nothing
        Dim ws As Excel.Worksheet = Nothing
        Dim currentSheet As Excel.Worksheet = Nothing
        Dim currentIndex As Integer
        Dim calcMode As Excel.XlCalculation

        Try
            originalWb = xlApp.ActiveWorkbook
            currentSheet = CType(xlApp.ActiveSheet, Excel.Worksheet)
            currentIndex = currentSheet.Index

            Dim response As DialogResult = MessageBox.Show($"Do you want to proceed with clearing {sheetDescription}?", "Confirm Action", MessageBoxButtons.OKCancel, MessageBoxIcon.Question)
            If response <> DialogResult.OK Then
                MessageBox.Show("Action canceled by the user.", "Canceled", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            calcMode = xlApp.Calculation
            xlApp.ScreenUpdating = False
            xlApp.EnableEvents = False
            xlApp.Calculation = Excel.XlCalculation.xlCalculationManual
            xlApp.DisplayAlerts = False

            ' Unlock all sheets
            UnlockAllSheetsInternal(originalWb)

            ws = SafeGetWorksheet(originalWb, sheetIndex)
            If ws Is Nothing Then
                MessageBox.Show($"The sheet at index {sheetIndex} was not found or cannot be accessed.", "Sheet Not Found", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            ws.Cells.Clear()
            ws.Cells.Interior.Color = WHITE

            With ws.Range("C1:E5")
                .Merge()
                .Value = headerText
                .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                .VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                .Font.Size = 14
                .Font.Bold = True
                .Font.Color = BLACK
                .Interior.Color = LIGHT_GRAY
                With .Borders
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .Weight = Excel.XlBorderWeight.xlThin
                End With
            End With

            Dim rowIdx As Integer = 8
            For Each info In infoRows
                ws.Cells(rowIdx, "A").Value = info
                rowIdx += 1
            Next

            ' Delete extra sheets
            DeleteExtraSheets(originalWb)

            ' Relock all sheets
            LockAllSheetsInternal(originalWb)

            ' Reactivate appropriate sheet
            If currentIndex <= 11 Then
                Try
                    Dim reactWs As Excel.Worksheet = SafeGetWorksheet(originalWb, currentIndex)
                    If reactWs IsNot Nothing Then
                        reactWs.Activate()
                        SafeReleaseComObject(reactWs)
                    Else
                        ws.Activate()
                    End If
                Catch
                    ws.Activate()
                End Try
            Else
                ws.Activate()
            End If

            MessageBox.Show($"{sheetDescription} cleared successfully.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            MessageBox.Show($"Error clearing {sheetDescription}: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            xlApp.ScreenUpdating = True
            xlApp.EnableEvents = True
            xlApp.Calculation = calcMode
            xlApp.DisplayAlerts = True
            SafeReleaseComObject(currentSheet)
            SafeReleaseComObject(ws)
            SafeReleaseComObject(originalWb)
        End Try
    End Sub

    Private Sub FormatSheet(ws As Excel.Worksheet, headerText As String)
        Try
            ws.Range("A7:E7").Merge()
            ws.Range("C1:E6").Merge()
            With ws.Range("A7")
                .Value = headerText
                .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                .VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                .Font.Bold = True
                .Interior.Color = LIGHT_GRAY
            End With

            Dim infoRows As String() = {
                "A1:-ASSEMBLY LIST ALIGNMENT", "A2:-ASSEMBLY LIST CREATION",
                "A3:-REMOVE DUPLICATE & CHECK DRAWING STATUS", "(PRESS 2 TIMES FOR FINAL)",
                "A3R:-SEPARATE REMOVE DUPLICATE", "A3D:-SEPARATE CHECK DRAWING STATUS",
                "A4:-REMOVE ASSEMBLY PROFILE", "GUID:-GENERATE GLOBAL UNIQUE ID FOR SEARCH ASSEMBLY IN MODEL",
                "(GENERATE GUID BEFORE PRESS A3)", "(IF PRESSED ALREADY THEN PRESS A2 BEFORE GUID)",
                "B1:-BOLT LIST ALIGNMENT", "B2:-ASSIGNED BOLT ALIGNMENT",
                "B3:-BOLT LIST CREATION", "B4:-GET BOLT WEIGHT", "B5:-ARRANGE BOLT DATA",
                "P1:-PART LIST ALIGNMENT", "P2:-PART LIST CREATION",
                "P3:-MATERIAL BREAKUP GENERATION", "Summary:-GET MATERIAL SUMMARY",
                "SaveAs:-GENERATE FILE IN BOQ OR SAVE AS BREAKUP FOLDER",
                "CLEAR:-FOR CLEAR INDIVIDUAL SHEET", "CLEAR ALL:-FOR CLEAR ALL SHEETS"}

            Dim rowIdx As Integer = 8
            For Each info In infoRows
                ws.Cells(rowIdx, "A").Value = info
                rowIdx += 1
            Next

            ' Set default formatting for all cells FIRST
            ws.Cells.Interior.Color = WHITE
            ws.Cells.Font.Bold = False
            ws.Cells.Font.Color = BLACK
            ws.Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignGeneral
            ws.Cells.VerticalAlignment = Excel.XlVAlign.xlVAlignTop

            ' Then apply specific formatting to A7:E7 (this will override the default)
            With ws.Range("A7:E7")
                .Font.Bold = True
                .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                .VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                .Font.Color = WHITE
                .Interior.Color = RGB(10, 50, 100)
            End With

        Catch ex As Exception
            MessageBox.Show($"Error formatting sheet: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub UnlockAllSheetsInternal(wb As Excel.Workbook)
        Dim ws As Excel.Worksheet = Nothing
        Try
            For Each ws In wb.Worksheets
                Try
                    ws.Unprotect(SHEET_PASSWORD)
                Catch
                    ' Continue if unprotect fails
                End Try
                SafeReleaseComObject(ws)
            Next
        Catch ex As Exception
            MessageBox.Show($"Error unlocking sheets: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub LockAllSheetsInternal(wb As Excel.Workbook)
        Dim ws As Excel.Worksheet = Nothing
        Try
            For Each ws In wb.Worksheets
                Try
                    ws.Protect(SHEET_PASSWORD, UserInterfaceOnly:=True)
                Catch
                    ' Continue if protect fails
                End Try
                SafeReleaseComObject(ws)
            Next
        Catch ex As Exception
            MessageBox.Show($"Error locking sheets: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub DeleteExtraSheets(wb As Excel.Workbook)
        Dim ws As Excel.Worksheet = Nothing
        Try
            Dim wsCount As Integer = wb.Sheets.Count
            If wsCount > 11 Then
                xlApp.DisplayAlerts = False
                For i As Integer = wsCount To 12 Step -1
                    ws = CType(wb.Sheets(i), Excel.Worksheet)
                    ws.Delete()
                    SafeReleaseComObject(ws)
                Next
                xlApp.DisplayAlerts = True
            End If
        Catch ex As Exception
            MessageBox.Show($"Error deleting extra sheets: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Function SafeGetWorksheet(wb As Excel.Workbook, sheetIndex As Integer) As Excel.Worksheet
        Try
            Return CType(wb.Sheets(sheetIndex), Excel.Worksheet)
        Catch
            Return Nothing
        End Try
    End Function

    Private Sub SafeReleaseComObject(ByRef obj As Object)
        If obj IsNot Nothing Then
            Try
                Marshal.ReleaseComObject(obj)
            Catch
            Finally
                obj = Nothing
            End Try
        End If
    End Sub

#End Region

End Class