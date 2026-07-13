Imports System.Collections.Generic
Imports System.Runtime.InteropServices
Imports System.IO
Imports WinForms = System.Windows.Forms
Imports Excel = Microsoft.Office.Interop.Excel

<ComVisible(True)>
<Guid("C13B569E-E3DE-4881-8D2C-BA60590F5B0D")>
<ClassInterface(ClassInterfaceType.AutoDual)>
Public Class SplitAntiSplit

    ' ================================
    ' Constructor with license validation
    ' ================================
    Public Sub New()
        EnsureLicense()                     ' ✅ Existing licensing
        CheckExpirationDateAndBackdate()   ' 🔐 Expiry + backdate tamper check
    End Sub

    ' ================================
    ' License validation methods
    ' ================================
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

    ' ================================
    ' Generate sheets by category - OPTIMIZED
    ' ================================
    <ComVisible(True)>
    Public Sub GenerateSheetsByCategory()
        Dim xlApp As Excel.Application = Nothing
        Dim wb As Excel.Workbook = Nothing
        Dim wsSource As Excel.Worksheet = Nothing
        Dim wsNew As Excel.Worksheet = Nothing
        Dim originalSheet As Excel.Worksheet = Nothing

        ' Data structures for bulk operations
        Dim lastRow As Integer = 0
        Dim uniqueCategories As New List(Of String)
        Dim categorySet As New HashSet(Of String)
        Dim categoryData As New Dictionary(Of String, List(Of Integer))
        Dim sourceDataArray As Object(,) = Nothing
        Dim headerArray As Object(,) = Nothing

        ' Performance counters
        Dim category As String = ""
        Dim i As Integer = 0

        Try
            xlApp = CType(Marshal.GetActiveObject("Excel.Application"), Excel.Application)
            wb = xlApp.ActiveWorkbook
            wsSource = CType(wb.ActiveSheet, Excel.Worksheet)
            originalSheet = CType(xlApp.ActiveSheet, Excel.Worksheet)

            ' PERFORMANCE: disable screen updating, calculations, alerts
            xlApp.ScreenUpdating = False
            xlApp.Calculation = Excel.XlCalculation.xlCalculationManual
            xlApp.DisplayAlerts = False
            xlApp.EnableEvents = False ' Additional performance boost

            ' Find last used row in column A
            lastRow = wsSource.Cells(wsSource.Rows.Count, "A").End(Excel.XlDirection.xlUp).Row
            If lastRow < 8 Then
                WinForms.MessageBox.Show("No data rows found (starting at row 8).", "Info", WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Information)
                Return
            End If

            ' ===== OPTIMIZATION 1: Read all data into memory arrays at once =====
            ' Read header rows (1-7) into array
            headerArray = CType(wsSource.Range("A1:K7").Value, Object(,))

            ' Read all data rows (8 to lastRow) into array for faster processing
            If lastRow >= 8 Then
                sourceDataArray = CType(wsSource.Range("A8:K" & lastRow).Value, Object(,))
            End If

            ' ===== OPTIMIZATION 2: Single pass to collect categories and organize data =====
            If sourceDataArray IsNot Nothing Then
                For i = 1 To sourceDataArray.GetLength(0) ' Array is 1-based
                    Dim categoryValue As Object = sourceDataArray(i, 2) ' Column B
                    If categoryValue IsNot Nothing Then
                        Dim catVal As String = Trim(CStr(categoryValue))
                        If catVal <> "" Then
                            If categorySet.Add(catVal) Then
                                uniqueCategories.Add(catVal)
                                categoryData(catVal) = New List(Of Integer)
                            End If
                            ' Store row index for this category
                            categoryData(catVal).Add(i)
                        End If
                    End If
                Next
            End If

            ' Delete all sheets after Sheet 11 (preserve 1..11)
            If wb.Sheets.Count >= 12 Then
                For i = wb.Sheets.Count To 12 Step -1
                    CType(wb.Sheets(i), Excel.Worksheet).Delete()
                Next
            End If

            ' Delete any existing sheets that match category names
            xlApp.DisplayAlerts = False
            For Each category In uniqueCategories
                Try
                    Dim wsExisting As Excel.Worksheet = CType(wb.Sheets(category), Excel.Worksheet)
                    If wsExisting IsNot Nothing Then
                        wsExisting.Delete()
                        Marshal.ReleaseComObject(wsExisting)
                    End If
                Catch
                    ' ignore if not present
                End Try
            Next

            ' ===== OPTIMIZATION 3: Process each category with bulk operations =====
            For Each category In uniqueCategories
                ' Add new sheet
                wsNew = CType(wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count)), Excel.Worksheet)

                ' Set sheet name safely
                Try
                    wsNew.Name = category
                Catch
                    Try
                        wsNew.Name = GetSafeSheetName(category, wb)
                    Catch
                        ' leave default name
                    End Try
                End Try

                ' ===== OPTIMIZATION 4: Copy headers with full formatting =====
                wsSource.Rows("1:7").Copy()
                wsNew.Rows("1:7").PasteSpecial(Excel.XlPasteType.xlPasteAll)
                xlApp.CutCopyMode = False

                ' ===== OPTIMIZATION 5: Ultra-fast internal filtering with bulk data+formatting copy =====
                Dim categoryRows As List(Of Integer) = categoryData(category)
                If categoryRows.Count > 0 Then
                    ' Create filtered data array for this category
                    Dim filteredData(categoryRows.Count - 1, 10) As Object ' 0-based, 11 columns

                    ' Build filtered data array in memory (super fast)
                    For rowIdx = 0 To categoryRows.Count - 1
                        Dim sourceRowIdx As Integer = categoryRows(rowIdx)
                        filteredData(rowIdx, 0) = rowIdx + 1 ' SR.NO
                        For colIdx = 1 To 10 ' Columns B-K from source data
                            filteredData(rowIdx, colIdx) = sourceDataArray(sourceRowIdx, colIdx + 1)
                        Next
                    Next

                    ' Write all data in one shot (fastest possible)
                    Dim targetRange As Excel.Range = wsNew.Range("A8").Resize(categoryRows.Count, 11)
                    targetRange.Value = filteredData

                    ' ===== FAST FORMATTING COPY: Copy formatting from source rows =====
                    ' Copy formatting template from first data row of source
                    wsSource.Rows(8).Copy()
                    wsNew.Range("A8:K8").PasteSpecial(Excel.XlPasteType.xlPasteFormats)
                    xlApp.CutCopyMode = False

                    ' Apply formatting to all data rows at once using FillDown equivalent
                    If categoryRows.Count > 1 Then
                        Dim formatRange As Excel.Range = wsNew.Range("A8:K" & (7 + categoryRows.Count))
                        wsNew.Range("A8:K8").Copy()
                        formatRange.PasteSpecial(Excel.XlPasteType.xlPasteFormats)
                        xlApp.CutCopyMode = False
                    End If
                End If

                ' ===== OPTIMIZATION 6: Calculate sums using array data instead of Excel functions =====
                Dim sumI As Double = 0
                Dim sumK As Double = 0
                For Each rowIdx In categoryData(category)
                    Try
                        If sourceDataArray(rowIdx, 9) IsNot Nothing Then ' Column I
                            sumI += CDbl(sourceDataArray(rowIdx, 9))
                        End If
                        If sourceDataArray(rowIdx, 11) IsNot Nothing Then ' Column K
                            sumK += CDbl(sourceDataArray(rowIdx, 11))
                        End If
                    Catch
                        ' ignore conversion errors
                    End Try
                Next

                wsNew.Range("J5").Value = sumI
                wsNew.Range("K5").Value = sumK

                ' Determine target last row after data copy
                Dim targetLast As Integer = 7 + categoryData(category).Count

                ' ===== OPTIMIZATION 7: Copy row heights (formatting preserved from copy operations above) =====
                ' Header row heights are already copied with the header copy operation
                ' Data row heights are already copied with the AutoFilter/row copy operations

                ' Copy shapes (keep original logic but optimize)
                CopyShapesOptimized(wsSource, wsNew)

                ' ===== OPTIMIZATION 8: Copy column widths from source sheet =====
                For col As Integer = 1 To 11 ' Columns A to K
                    wsNew.Columns(col).ColumnWidth = wsSource.Columns(col).ColumnWidth
                Next

                ' Remove form controls/OLE objects
                Try
                    Dim shapesToDelete As New List(Of Excel.Shape)
                    For Each shp As Excel.Shape In wsNew.Shapes
                        If shp.Type = Microsoft.Office.Core.MsoShapeType.msoFormControl OrElse
                           shp.Type = Microsoft.Office.Core.MsoShapeType.msoOLEControlObject Then
                            shapesToDelete.Add(shp)
                        End If
                    Next
                    For Each shp In shapesToDelete
                        shp.Delete()
                    Next
                Catch
                End Try

                ' Freeze panes, format, and setup (keep original logic)
                SetupSheetFormatting(wsNew, xlApp, category, targetLast)

                ' Release COM object
                If wsNew IsNot Nothing Then
                    Try
                        Marshal.ReleaseComObject(wsNew)
                    Catch
                    End Try
                    wsNew = Nothing
                End If
            Next

            WinForms.MessageBox.Show("Sheets generated successfully for each category.", "Information", WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Information)

        Catch ex As Exception
            WinForms.MessageBox.Show("Error: " & ex.Message, "Error", WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Error)
        Finally
            ' Restore original state
            Try
                If originalSheet IsNot Nothing Then originalSheet.Activate()
            Catch
            End Try

            If xlApp IsNot Nothing Then
                xlApp.DisplayAlerts = True
                xlApp.ScreenUpdating = True
                xlApp.Calculation = Excel.XlCalculation.xlCalculationAutomatic
                xlApp.EnableEvents = True
            End If

            ' Clean up COM objects
            Try
                If wsSource IsNot Nothing Then Marshal.ReleaseComObject(wsSource)
                If wb IsNot Nothing Then Marshal.ReleaseComObject(wb)
            Catch
            End Try
            xlApp = Nothing
        End Try
    End Sub

    ' ================================
    ' Optimized row heights copy
    ' ================================
    <ComVisible(True)>
    Private Sub CopyRowHeightsBulk(wsSource As Excel.Worksheet, wsNew As Excel.Worksheet, categoryRowIndices As List(Of Integer))
        Try
            ' Copy header row heights (1-7) in bulk
            For r As Integer = 1 To 7
                Try
                    wsNew.Rows(r).RowHeight = wsSource.Rows(r).RowHeight
                Catch
                End Try
            Next

            ' Copy data row heights based on original positions
            Dim destRow As Integer = 8
            For Each sourceRowIdx In categoryRowIndices
                Try
                    wsNew.Rows(destRow).RowHeight = wsSource.Rows(sourceRowIdx + 7).RowHeight ' +7 because array is 0-based but rows start at 8
                Catch
                End Try
                destRow += 1
            Next
        Catch
        End Try
    End Sub

    ' ================================
    ' Optimized shapes copying
    ' ================================
    <ComVisible(True)>
    Private Sub CopyShapesOptimized(wsSource As Excel.Worksheet, wsNew As Excel.Worksheet)
        Try
            wsNew.Activate()
            wsNew.Range("A1").Select()

            ' Get all shapes to copy first (avoid collection modification during iteration)
            Dim shapesToCopy As New List(Of Excel.Shape)
            For Each sourceShape As Excel.Shape In wsSource.Shapes
                If sourceShape.Type <> Microsoft.Office.Core.MsoShapeType.msoFormControl AndAlso
                   sourceShape.Type <> Microsoft.Office.Core.MsoShapeType.msoOLEControlObject Then
                    shapesToCopy.Add(sourceShape)
                End If
            Next

            ' Copy shapes
            For Each sourceShape In shapesToCopy
                Try
                    sourceShape.Copy()
                    wsNew.Paste()
                    Dim newShape As Excel.Shape = wsNew.Shapes(wsNew.Shapes.Count)

                    ' Restore geometry in bulk
                    With newShape
                        .Left = sourceShape.Left
                        .Top = sourceShape.Top
                        .Width = sourceShape.Width
                        .Height = sourceShape.Height
                    End With

                    If sourceShape.TopLeftCell IsNot Nothing AndAlso sourceShape.TopLeftCell.Address = "$A$1" Then
                        newShape.Placement = Excel.XlPlacement.xlFreeFloating
                    End If

                    Try
                        newShape.Name = sourceShape.Name & "_" & wsNew.Name
                    Catch
                    End Try
                Catch
                    ' ignore individual shape errors
                End Try
            Next
        Catch
        End Try
    End Sub

    ' ================================
    ' Sheet formatting setup
    ' ================================
    <ComVisible(True)>
    Private Sub SetupSheetFormatting(wsNew As Excel.Worksheet, xlApp As Excel.Application, category As String, targetLast As Integer)
        Try
            ' Freeze top 7 rows
            wsNew.Activate()
            With xlApp.ActiveWindow
                .SplitColumn = 0
                .SplitRow = 7
                .FreezePanes = True
            End With

            ' AutoFit rows only
            wsNew.Rows.AutoFit()

            ' Show category name
            wsNew.Range("J3").Value = category

            ' Hide columns L:O
            wsNew.Columns("L:O").EntireColumn.Hidden = True

            ' Set print area and page setup
            If targetLast < 8 Then targetLast = 7
            wsNew.PageSetup.PrintArea = "$A$1:$O$" & (targetLast + 1).ToString()

            With wsNew.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesWide = 1
                .FitToPagesTall = False
                .TopMargin = xlApp.InchesToPoints(0.2)
                .BottomMargin = xlApp.InchesToPoints(0.3)
                .LeftMargin = xlApp.InchesToPoints(0.2)
                .RightMargin = xlApp.InchesToPoints(0.2)
                .HeaderMargin = xlApp.InchesToPoints(0)
                .FooterMargin = xlApp.InchesToPoints(0)
                .CenterHorizontally = True
                .CenterVertically = False
                .PrintTitleRows = "$6:$7"
                .PaperSize = Excel.XlPaperSize.xlPaperA4
                .LeftFooter = "Page &P of &N"
                .RightFooter = "Raghav"
                .AlignMarginsHeaderFooter = True
            End With
        Catch
        End Try
    End Sub

    ' ================================
    ' Copy row heights helper (kept for compatibility)
    ' ================================
    <ComVisible(True)>
    Public Sub CopyRowHeights(wsSource As Excel.Worksheet, wsNew As Excel.Worksheet, startRow As Integer, endRow As Integer)
        Try
            For r As Integer = startRow To endRow
                Try
                    wsNew.Rows(r).RowHeight = wsSource.Rows(r).RowHeight
                Catch
                End Try
            Next
        Catch
        End Try
    End Sub

    ' ================================
    ' Delete and renumber sheets - OPTIMIZED
    ' ================================
    <ComVisible(True)>
    Public Sub DeleteAndRenumberSheets()
        Dim xlApp As Excel.Application = Nothing

        Try
            xlApp = CType(Marshal.GetActiveObject("Excel.Application"), Excel.Application)
            xlApp.ScreenUpdating = False
            xlApp.DisplayAlerts = False
            xlApp.EnableEvents = False

            Dim wb As Excel.Workbook = xlApp.ActiveWorkbook

            ' Delete sheets from 12 onward
            If wb.Sheets.Count >= 12 Then
                For i As Integer = wb.Sheets.Count To 12 Step -1
                    CType(wb.Sheets(i), Excel.Worksheet).Delete()
                Next
            End If

            ' Renumber remaining sheets
            For i As Integer = 1 To wb.Sheets.Count
                Try
                    CType(wb.Sheets(i), Excel.Worksheet).Name = "Sheet" & i
                Catch
                End Try
            Next

            ' Activate first sheet
            Try
                CType(wb.Sheets(1), Excel.Worksheet).Activate()
                CType(wb.Sheets(1), Excel.Worksheet).Range("A1").Select()
            Catch
            End Try

            WinForms.MessageBox.Show("Sheets Generated by Split deleted successfully.", "Information", WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Information)

        Catch ex As Exception
            WinForms.MessageBox.Show("Error in DeleteAndRenumberSheets: " & ex.Message, "Error", WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Error)
        Finally
            If xlApp IsNot Nothing Then
                xlApp.DisplayAlerts = True
                xlApp.ScreenUpdating = True
                xlApp.EnableEvents = True
            End If
        End Try
    End Sub

    ' ================================
    ' Helper to return a safe sheet name
    ' ================================
    Private Function GetSafeSheetName(rawName As String, wb As Excel.Workbook) As String
        If rawName Is Nothing Then rawName = "Sheet"
        Dim s As String = rawName

        ' Remove invalid characters
        Dim invalidChars As String() = {":", "\", "/", "?", "*", "[", "]"}
        For Each ch As String In invalidChars
            s = s.Replace(ch, "_")
        Next

        If s.Length > 31 Then s = s.Substring(0, 31)

        ' Handle duplicate names
        Dim base As String = s
        Dim suffix As Integer = 1

        Try
            While True
                Dim exists As Boolean = False
                For Each sh As Excel.Worksheet In wb.Sheets
                    If String.Compare(sh.Name, s, StringComparison.OrdinalIgnoreCase) = 0 Then
                        exists = True
                        Exit For
                    End If
                Next
                If Not exists Then Exit While

                s = base
                Dim suffixStr As String = "_" & suffix.ToString()
                Dim remain As Integer = 31 - suffixStr.Length
                If s.Length > remain Then s = s.Substring(0, remain)
                s = s & suffixStr
                suffix += 1
            End While
        Catch
        End Try

        Return s
    End Function

End Class