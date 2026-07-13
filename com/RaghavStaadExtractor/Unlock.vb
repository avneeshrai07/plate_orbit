Imports System
Imports System.Collections.Generic
Imports System.IO
Imports System.Runtime.InteropServices
Imports Microsoft.Office.Interop.Excel

<ComVisible(True)>
<Guid("F1234567-89AB-4CDE-8123-ABCDEF123456")>
<ProgId("RaghavStaadExtractor.Unlock")>
Public Class Unlock

    ' Unlock specific range E2:AQ5
    Public Sub UnlockSheetRange()
        Dim excelApp As Application = Nothing
        Dim workbook As Workbook = Nothing
        Dim sheet As Worksheet = Nothing
        Try
            excelApp = Marshal.GetActiveObject("Excel.Application")
            workbook = excelApp.ActiveWorkbook
            sheet = CType(workbook.Sheets("Sheet1"), Worksheet)
            ' Temporarily unprotect the sheet
            sheet.Unprotect("2022")
            ' Unlock only the target range
            Dim unlockRange As Range = sheet.Range("E2:AQ5")
            unlockRange.Locked = False
            ' Re-protect the sheet while allowing unlocked cells to be edited
            sheet.Protect("2022", DrawingObjects:=True, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True)
        Catch ex As Exception
            MsgBox("Error unlocking specific range: " & ex.Message)
        Finally
            If sheet IsNot Nothing Then Marshal.ReleaseComObject(sheet)
            If workbook IsNot Nothing Then Marshal.ReleaseComObject(workbook)
            If excelApp IsNot Nothing Then Marshal.ReleaseComObject(excelApp)
            GC.Collect()
            GC.WaitForPendingFinalizers()
        End Try
    End Sub

    ' Toggle filter method
    Public Sub ToggleUniqueFilter()
        Dim excelApp As Application = Nothing
        Dim workbook As Workbook = Nothing
        Dim sheet As Worksheet = Nothing
        Try
            excelApp = Marshal.GetActiveObject("Excel.Application")
            workbook = excelApp.ActiveWorkbook
            sheet = CType(workbook.Sheets("Sheet1"), Worksheet)

            ' Unprotect sheet
            Try
                sheet.Unprotect("2022")
            Catch
            End Try

            ' Check if filtered data already exists (look for "FILTERED DATA" marker)
            Dim lastRow As Long = sheet.Cells(sheet.Rows.Count, "E").End(XlDirection.xlUp).Row
            Dim hasFilteredData As Boolean = False

            ' Look for the filtered data marker
            For i As Long = 9 To lastRow
                If sheet.Cells(i, 5).Value IsNot Nothing AndAlso
                   sheet.Cells(i, 5).Value.ToString().Contains("FILTERED DATA") Then
                    hasFilteredData = True
                    Exit For
                End If
            Next

            If hasFilteredData Then
                ' Remove filtered data and show original rows
                RemoveFilteredData(sheet)
            Else
                ' Apply unique filter by creating filtered data below original
                CreateFilteredData(sheet)
            End If

        Catch ex As Exception
            MsgBox("Error in ToggleUniqueFilter: " & ex.Message)
        Finally
            Try
                If sheet IsNot Nothing Then sheet.Protect("2022", UserInterfaceOnly:=True)
            Catch
            End Try
            If sheet IsNot Nothing Then Marshal.ReleaseComObject(sheet)
            If workbook IsNot Nothing Then Marshal.ReleaseComObject(workbook)
            If excelApp IsNot Nothing Then Marshal.ReleaseComObject(excelApp)
            GC.Collect()
            GC.WaitForPendingFinalizers()
        End Try
    End Sub

    ' Create filtered data below original data
    Private Sub CreateFilteredData(sheet As Worksheet)
        ' Find last row with data in column E
        Dim lastRow As Long = sheet.Cells(sheet.Rows.Count, "E").End(XlDirection.xlUp).Row

        If lastRow < 9 Then
            MsgBox("No data found from row 9 onwards in column E")
            Exit Sub
        End If

        ' Hide all original rows from 9 to lastRow
        For i As Long = 9 To lastRow
            sheet.Rows(i).Hidden = True
        Next

        ' Collect unique values with their row numbers for proper numerical sorting
        Dim uniqueValueRows As New Dictionary(Of String, Long)

        ' First pass: collect unique values and their first occurrence row
        For i As Long = 9 To lastRow
            Dim cellValue As String = CStr(sheet.Cells(i, 5).Value) ' Column E = 5
            If Not String.IsNullOrEmpty(cellValue) AndAlso cellValue.StartsWith("R") Then
                If Not uniqueValueRows.ContainsKey(cellValue) Then
                    uniqueValueRows.Add(cellValue, i)
                End If
            End If
        Next

        ' Convert to list for custom sorting
        Dim valueRowPairs As New List(Of KeyValuePair(Of String, Long))(uniqueValueRows)

        ' Sort using custom comparer for proper numerical order (R1, R2, R3, R8, R15, R20...)
        valueRowPairs.Sort(Function(x, y) CompareAlphanumeric(x.Key, y.Key))

        ' Add a separator row
        Dim startFilterRow As Long = lastRow + 1
        sheet.Cells(startFilterRow, 5).Value = "UNIQUE PROPERTIES FILTERED DATA - DO NOT EDIT"
        sheet.Range(sheet.Cells(startFilterRow, 5), sheet.Cells(startFilterRow, 43)).Interior.Color = RGB(255, 255, 255) ' WHITE background

        ' Copy each unique row to the filtered section
        Dim currentFilterRow As Long = startFilterRow + 1

        For Each pair In valueRowPairs
            Dim sourceRow As Long = pair.Value

            ' Copy the entire row
            Dim sourceRange As Range = sheet.Range(sheet.Cells(sourceRow, 1), sheet.Cells(sourceRow, 43))
            Dim targetRange As Range = sheet.Range(sheet.Cells(currentFilterRow, 1), sheet.Cells(currentFilterRow, 43))

            sourceRange.Copy(targetRange)

            currentFilterRow += 1
        Next

        ' Unlock copy columns for filtered data
        UnlockCopyColumns(sheet, currentFilterRow - 1)

        MsgBox("Filter Applied - Unique rows displayed below original data")
    End Sub

    ' Remove filtered data and show original rows
    Private Sub RemoveFilteredData(sheet As Worksheet)
        ' Find the last row with data
        Dim lastRow As Long = sheet.Cells(sheet.Rows.Count, "E").End(XlDirection.xlUp).Row

        ' Find the filtered data section
        Dim filterStartRow As Long = 0
        For i As Long = 9 To lastRow
            If sheet.Cells(i, 5).Value IsNot Nothing AndAlso
               sheet.Cells(i, 5).Value.ToString().Contains("UNIQUE PROPERTIES") Then
                filterStartRow = i
                Exit For
            End If
        Next

        If filterStartRow = 0 Then
            ' No filtered data found, just show all rows
            sheet.Rows.Hidden = False
            MsgBox("Filter Removed")
            Exit Sub
        End If

        ' Delete all rows from the filter start to the end
        sheet.Range(sheet.Rows(filterStartRow), sheet.Rows(lastRow)).Delete()

        ' Show all original rows
        sheet.Rows.Hidden = False

        ' Lock copy columns back
        LockCopyColumns(sheet)

        MsgBox("Filter Removed - Filtered data deleted")
    End Sub

    ' Save filtered data to new file (updated to work with new filtered data location)
    ' Save filtered data to new file
    Public Sub SaveAs_filter()
        Dim excelApp As Application = Nothing
        Dim sourceWorkbook As Workbook = Nothing
        Dim sourceSheet As Worksheet = Nothing
        Dim newWorkbook As Workbook = Nothing
        Dim newSheet As Worksheet = Nothing
        Dim folderPath As String
        Dim fileName As String
        Dim fullFilePath As String
        Dim overwriteChoice As MsgBoxResult
        Dim i As Integer = 1

        Try
            ' Get Excel instance with improved error handling
            excelApp = GetExcelInstance()
            If excelApp Is Nothing Then Exit Sub

            excelApp.ScreenUpdating = False
            excelApp.DisplayAlerts = False
            excelApp.EnableEvents = False

            sourceWorkbook = excelApp.ActiveWorkbook
            If sourceWorkbook Is Nothing Then
                MsgBox("No active workbook found.", MsgBoxStyle.Critical)
                Exit Sub
            End If

            sourceSheet = TryCast(sourceWorkbook.Sheets("Sheet1"), Worksheet)
            If sourceSheet Is Nothing Then
                MsgBox("Sheet1 not found in the active workbook.", MsgBoxStyle.Critical)
                Exit Sub
            End If

            ' Unprotect original temporarily
            SafeUnprotectSheet(sourceSheet)

            ' Prepare folder path
            folderPath = Path.Combine(sourceWorkbook.Path, "PROPERTY")
            If Not Directory.Exists(folderPath) Then Directory.CreateDirectory(folderPath)

            ' Get base filename from E3
            Dim baseName As String = GetCellTextSafely(sourceSheet.Range("E3"))
            If String.IsNullOrWhiteSpace(baseName) Then baseName = "Unnamed"

            fileName = baseName & " PROPERTIES.xlsx"
            fullFilePath = Path.Combine(folderPath, fileName)

            ' Handle file exists
            If File.Exists(fullFilePath) Then
                overwriteChoice = MsgBox("File already exists. Overwrite?", MsgBoxStyle.YesNoCancel + MsgBoxStyle.Question)
                If overwriteChoice = MsgBoxResult.Yes Then
                    Try
                        File.Delete(fullFilePath)
                    Catch ex As Exception
                        MsgBox("Error deleting existing file: " & ex.Message, MsgBoxStyle.Critical)
                        Exit Sub
                    End Try
                ElseIf overwriteChoice = MsgBoxResult.No Then
                    Do While File.Exists(Path.Combine(folderPath, baseName & " PROPERTIES (" & i & ").xlsx"))
                        i += 1
                    Loop
                    fileName = baseName & " PROPERTIES (" & i & ").xlsx"
                    fullFilePath = Path.Combine(folderPath, fileName)
                Else
                    GoTo CleanUp
                End If
            End If

            ' Create new workbook
            newWorkbook = excelApp.Workbooks.Add()
            newSheet = CType(newWorkbook.Sheets(1), Worksheet)
            newSheet.Name = "Properties"

            ' Merge cells B1:D1 in the new sheet
            Try
                Dim mergeRange As Range = newSheet.Range("B1:D1")
                mergeRange.Merge()
                mergeRange.HorizontalAlignment = XlHAlign.xlHAlignCenter
                mergeRange.VerticalAlignment = XlVAlign.xlVAlignCenter
            Catch ex As Exception
                ' Ignore merge errors
            End Try

            ' Copy data from specific columns and only visible rows (excluding top 6 rows)
            CopyVisibleColumnsData(sourceSheet, newSheet, excelApp)

            ' Save the new workbook
            Try
                newWorkbook.SaveAs(fullFilePath, XlFileFormat.xlOpenXMLWorkbook)
                newWorkbook.Close(False)
                newWorkbook = Nothing
            Catch ex As Exception
                MsgBox("Error saving the workbook: " & ex.Message, MsgBoxStyle.Critical)
            End Try

            MsgBox("Unique Properties SaveAs Successfully")

CleanUp:
            ' Clean up
            If Not newSheet Is Nothing Then Marshal.ReleaseComObject(newSheet)
            If Not newWorkbook Is Nothing Then
                newWorkbook.Close(False)
                Marshal.ReleaseComObject(newWorkbook)
            End If

            ' Reprotect original
            SafeProtectSheet(sourceSheet)

            If Not sourceSheet Is Nothing Then Marshal.ReleaseComObject(sourceSheet)
            If Not sourceWorkbook Is Nothing Then Marshal.ReleaseComObject(sourceWorkbook)

            ' Restore Excel settings
            If Not excelApp Is Nothing Then
                excelApp.ScreenUpdating = True
                excelApp.DisplayAlerts = True
                excelApp.EnableEvents = True
            End If

        Catch ex As Exception
            MsgBox("Unexpected error: " & ex.Message, MsgBoxStyle.Critical)
        Finally
            If Not excelApp Is Nothing Then Marshal.ReleaseComObject(excelApp)
        End Try
    End Sub

    ' Copy only specific columns from visible rows (starting from row 7, skipping top 6 rows)
    Private Sub CopyVisibleColumnsData(sourceSheet As Worksheet, targetSheet As Worksheet, excelApp As Application)
        Try
            ' Find the last row with data
            Dim lastRow As Long = sourceSheet.Cells(sourceSheet.Rows.Count, "E").End(XlDirection.xlUp).Row
            If lastRow < 7 Then Exit Sub ' No data after row 6

            Dim targetRow As Long = 1
            Dim sourceColumns As Integer() = {5, 37, 38, 39} ' E=5, AK=37, AL=38, AM=39
            Dim targetColumns As String() = {"A", "B", "C", "D"} ' Target columns

            ' Dictionary to track processed merged cells to avoid duplicates
            Dim processedMerges As New Dictionary(Of String, Boolean)

            ' Copy data from visible rows only (starting from row 7 to skip top 6 rows)
            For sourceRowIndex As Long = 7 To lastRow
                ' Check if the row is visible (not hidden)
                If Not sourceSheet.Rows(sourceRowIndex).Hidden Then

                    ' Copy data from each specified column with formatting
                    For colIndex As Integer = 0 To sourceColumns.Length - 1
                        Dim sourceRange As Range = sourceSheet.Cells(sourceRowIndex, sourceColumns(colIndex))
                        Dim targetRange As Range = targetSheet.Cells(targetRow, colIndex + 1)

                        ' Copy value only if not empty or if it's the top-left cell of a merge
                        If Not IsNothing(sourceRange.Value) OrElse sourceRange.MergeCells Then
                            targetRange.Value = sourceRange.Value
                        End If

                        ' Copy formatting using direct property copying to avoid clipboard issues
                        With sourceRange
                            targetRange.Font.Name = .Font.Name
                            targetRange.Font.Size = .Font.Size
                            targetRange.Font.Bold = .Font.Bold
                            targetRange.Font.Italic = .Font.Italic
                            targetRange.Font.Color = .Font.Color
                            targetRange.Interior.Color = .Interior.Color
                            targetRange.HorizontalAlignment = .HorizontalAlignment
                            targetRange.VerticalAlignment = .VerticalAlignment
                            targetRange.WrapText = .WrapText
                            targetRange.NumberFormat = .NumberFormat

                            ' Copy borders
                            targetRange.Borders(XlBordersIndex.xlEdgeLeft).LineStyle = .Borders(XlBordersIndex.xlEdgeLeft).LineStyle
                            targetRange.Borders(XlBordersIndex.xlEdgeLeft).Weight = .Borders(XlBordersIndex.xlEdgeLeft).Weight
                            targetRange.Borders(XlBordersIndex.xlEdgeRight).LineStyle = .Borders(XlBordersIndex.xlEdgeRight).LineStyle
                            targetRange.Borders(XlBordersIndex.xlEdgeRight).Weight = .Borders(XlBordersIndex.xlEdgeRight).Weight
                            targetRange.Borders(XlBordersIndex.xlEdgeTop).LineStyle = .Borders(XlBordersIndex.xlEdgeTop).LineStyle
                            targetRange.Borders(XlBordersIndex.xlEdgeTop).Weight = .Borders(XlBordersIndex.xlEdgeTop).Weight
                            targetRange.Borders(XlBordersIndex.xlEdgeBottom).LineStyle = .Borders(XlBordersIndex.xlEdgeBottom).LineStyle
                            targetRange.Borders(XlBordersIndex.xlEdgeBottom).Weight = .Borders(XlBordersIndex.xlEdgeBottom).Weight
                        End With

                        ' Handle merged cells properly
                        If sourceRange.MergeCells Then
                            Try
                                Dim mergeArea As Range = sourceRange.MergeArea
                                Dim mergeKey As String = $"{sourceRowIndex}_{sourceColumns(colIndex)}"

                                ' Only process if this merge hasn't been handled yet
                                If Not processedMerges.ContainsKey(mergeKey) Then
                                    processedMerges.Add(mergeKey, True)

                                    ' Calculate how many rows this merge spans in visible rows only
                                    Dim visibleRowsInMerge As Integer = 0
                                    Dim mergeStartRow As Long = mergeArea.Row
                                    Dim mergeEndRow As Long = mergeArea.Row + mergeArea.Rows.Count - 1

                                    ' Count visible rows within the merge area that are >= row 7
                                    For checkRow As Long = Math.Max(mergeStartRow, 7) To mergeEndRow
                                        If Not sourceSheet.Rows(checkRow).Hidden Then
                                            visibleRowsInMerge += 1
                                        End If
                                    Next

                                    ' If more than 1 visible row, create merge in target
                                    If visibleRowsInMerge > 1 Then
                                        Dim targetMergeRange As Range = targetSheet.Range(targetRange, targetRange.Offset(visibleRowsInMerge - 1, 0))
                                        If Not targetMergeRange.MergeCells Then
                                            targetMergeRange.Merge()
                                            targetMergeRange.HorizontalAlignment = sourceRange.HorizontalAlignment
                                            targetMergeRange.VerticalAlignment = sourceRange.VerticalAlignment
                                        End If
                                    End If
                                End If
                            Catch ex As Exception
                                ' Ignore merge errors and continue
                                System.Diagnostics.Debug.WriteLine("Merge error: " & ex.Message)
                            End Try
                        End If
                    Next

                    ' Copy row height
                    Try
                        targetSheet.Rows(targetRow).RowHeight = sourceSheet.Rows(sourceRowIndex).RowHeight
                    Catch
                        ' Ignore row height errors
                    End Try

                    targetRow += 1
                End If
            Next

            ' Set column widths to match source columns
            Try
                targetSheet.Columns("A:A").ColumnWidth = sourceSheet.Columns("E:E").ColumnWidth
                targetSheet.Columns("B:B").ColumnWidth = sourceSheet.Columns("AK:AK").ColumnWidth
                targetSheet.Columns("C:C").ColumnWidth = sourceSheet.Columns("AL:AL").ColumnWidth
                targetSheet.Columns("D:D").ColumnWidth = sourceSheet.Columns("AM:AM").ColumnWidth
            Catch
                ' If width setting fails, auto-fit
                targetSheet.Columns.AutoFit()
            End Try

        Catch ex As Exception
            MsgBox("Error copying visible columns data: " & ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub

    ' Helper methods for Excel operations
    Private Function GetExcelInstance() As Application
        Try
            Return CType(Marshal.GetActiveObject("Excel.Application"), Application)
        Catch ex As Exception
            MsgBox("Excel is not running or could not be accessed: " & ex.Message, MsgBoxStyle.Critical)
            Return Nothing
        End Try
    End Function


    ' Copy only from filtered section (below the original data)
    Private Sub CopyFilteredColumnsData(sourceSheet As Worksheet, targetSheet As Worksheet, excelApp As Application)
        Try
            ' Find the filtered data section
            Dim lastRow As Long = sourceSheet.Cells(sourceSheet.Rows.Count, "E").End(XlDirection.xlUp).Row
            Dim filterStartRow As Long = 0
            Dim filterEndRow As Long = 0

            ' Find the start of filtered data
            For i As Long = 9 To lastRow
                If sourceSheet.Cells(i, 5).Value IsNot Nothing AndAlso
                   sourceSheet.Cells(i, 5).Value.ToString().Contains("FILTERED DATA") Then
                    filterStartRow = i + 1 ' Data starts after the marker
                    Exit For
                End If
            Next

            If filterStartRow = 0 Then
                MsgBox("No filtered data found", MsgBoxStyle.Exclamation)
                Exit Sub
            End If

            ' Find the end of filtered data (next empty cell in column E after filter start)
            For i As Long = filterStartRow To lastRow
                If sourceSheet.Cells(i, 5).Value Is Nothing OrElse
                   String.IsNullOrEmpty(sourceSheet.Cells(i, 5).Value.ToString()) Then
                    filterEndRow = i - 1
                    Exit For
                End If
            Next

            If filterEndRow = 0 Then filterEndRow = lastRow

            Dim targetRow As Long = 1
            Dim sourceColumns As Integer() = {5, 37, 38, 39} ' E=5, AK=37, AL=38, AM=39
            Dim targetColumns As String() = {"A", "B", "C", "D"} ' Target columns

            ' Copy data from filtered rows only
            For sourceRowIndex As Long = filterStartRow To filterEndRow
                ' Copy data from each specified column with formatting
                For colIndex As Integer = 0 To sourceColumns.Length - 1
                    Dim sourceRange As Range = sourceSheet.Cells(sourceRowIndex, sourceColumns(colIndex))
                    Dim targetRange As Range = targetSheet.Cells(targetRow, colIndex + 1)

                    ' Copy value only if not empty
                    If Not IsNothing(sourceRange.Value) Then
                        targetRange.Value = sourceRange.Value
                    End If

                    ' Copy formatting using direct property copying to avoid clipboard issues
                    With sourceRange
                        targetRange.Font.Name = .Font.Name
                        targetRange.Font.Size = .Font.Size
                        targetRange.Font.Bold = .Font.Bold
                        targetRange.Font.Italic = .Font.Italic
                        targetRange.Font.Color = .Font.Color
                        targetRange.Interior.Color = .Interior.Color
                        targetRange.HorizontalAlignment = .HorizontalAlignment
                        targetRange.VerticalAlignment = .VerticalAlignment
                        targetRange.WrapText = .WrapText
                        targetRange.NumberFormat = .NumberFormat

                        ' Copy borders
                        targetRange.Borders(XlBordersIndex.xlEdgeLeft).LineStyle = .Borders(XlBordersIndex.xlEdgeLeft).LineStyle
                        targetRange.Borders(XlBordersIndex.xlEdgeLeft).Weight = .Borders(XlBordersIndex.xlEdgeLeft).Weight
                        targetRange.Borders(XlBordersIndex.xlEdgeRight).LineStyle = .Borders(XlBordersIndex.xlEdgeRight).LineStyle
                        targetRange.Borders(XlBordersIndex.xlEdgeRight).Weight = .Borders(XlBordersIndex.xlEdgeRight).Weight
                        targetRange.Borders(XlBordersIndex.xlEdgeTop).LineStyle = .Borders(XlBordersIndex.xlEdgeTop).LineStyle
                        targetRange.Borders(XlBordersIndex.xlEdgeTop).Weight = .Borders(XlBordersIndex.xlEdgeTop).Weight
                        targetRange.Borders(XlBordersIndex.xlEdgeBottom).LineStyle = .Borders(XlBordersIndex.xlEdgeBottom).LineStyle
                        targetRange.Borders(XlBordersIndex.xlEdgeBottom).Weight = .Borders(XlBordersIndex.xlEdgeBottom).Weight
                    End With
                Next

                ' Copy row height
                Try
                    targetSheet.Rows(targetRow).RowHeight = sourceSheet.Rows(sourceRowIndex).RowHeight
                Catch
                    ' Ignore row height errors
                End Try

                targetRow += 1
            Next

            ' Set column widths to match source columns
            Try
                targetSheet.Columns("A:A").ColumnWidth = sourceSheet.Columns("E:E").ColumnWidth
                targetSheet.Columns("B:B").ColumnWidth = sourceSheet.Columns("AK:AK").ColumnWidth
                targetSheet.Columns("C:C").ColumnWidth = sourceSheet.Columns("AL:AL").ColumnWidth
                targetSheet.Columns("D:D").ColumnWidth = sourceSheet.Columns("AM:AM").ColumnWidth
            Catch
                ' If width setting fails, auto-fit
                targetSheet.Columns.AutoFit()
            End Try

        Catch ex As Exception
            MsgBox("Error copying filtered columns data: " & ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub

    ' Custom alphanumeric comparison for proper sorting (R1, R2, R3, R8, R15, R20...)
    Private Function CompareAlphanumeric(x As String, y As String) As Integer
        ' Extract letter and number parts
        Dim xLetter As String = ""
        Dim xNumber As Integer = 0
        Dim yLetter As String = ""
        Dim yNumber As Integer = 0

        ExtractLetterNumber(x, xLetter, xNumber)
        ExtractLetterNumber(y, yLetter, yNumber)

        ' First compare letters
        Dim letterCompare As Integer = String.Compare(xLetter, yLetter)
        If letterCompare <> 0 Then
            Return letterCompare
        End If

        ' If letters are same, compare numbers numerically
        Return xNumber.CompareTo(yNumber)
    End Function

    ' Helper to extract letter and number from strings like "R15", "A5", etc.
    Private Sub ExtractLetterNumber(input As String, ByRef letter As String, ByRef number As Integer)
        If String.IsNullOrEmpty(input) Then
            letter = ""
            number = 0
            Return
        End If

        Dim i As Integer = 0
        ' Extract letter part
        While i < input.Length AndAlso Not Char.IsDigit(input(i))
            i += 1
        End While

        letter = If(i > 0, input.Substring(0, i), "")

        ' Extract number part
        If i < input.Length Then
            Dim numberStr As String = input.Substring(i)
            Integer.TryParse(numberStr, number)
        Else
            number = 0
        End If
    End Sub

    ' Helper method to unlock copy columns when filter is applied
    Private Sub UnlockCopyColumns(sheet As Worksheet, lastFilteredRow As Long)
        Try
            ' Find the filtered data section
            Dim lastRow As Long = sheet.Cells(sheet.Rows.Count, "E").End(XlDirection.xlUp).Row
            Dim filterStartRow As Long = 0

            ' Find the start of filtered data
            For i As Long = 9 To lastRow
                If sheet.Cells(i, 5).Value IsNot Nothing AndAlso
                   sheet.Cells(i, 5).Value.ToString().Contains("FILTERED DATA") Then
                    filterStartRow = i + 1 ' Data starts after the marker
                    Exit For
                End If
            Next

            If filterStartRow = 0 Then Exit Sub

            ' Unlock Column E in filtered section
            Dim rangeE As Range = sheet.Range($"E{filterStartRow}:E{lastFilteredRow}")
            rangeE.Locked = False

            ' Unlock Columns AK, AL, AM in filtered section
            Dim rangeAK As Range = sheet.Range($"AK{filterStartRow}:AK{lastFilteredRow}")
            Dim rangeAL As Range = sheet.Range($"AL{filterStartRow}:AL{lastFilteredRow}")
            Dim rangeAM As Range = sheet.Range($"AM{filterStartRow}:AM{lastFilteredRow}")

            rangeAK.Locked = False
            rangeAL.Locked = False
            rangeAM.Locked = False

        Catch ex As Exception
            MsgBox("Error unlocking copy columns: " & ex.Message)
        End Try
    End Sub

    ' Helper method to lock copy columns when filter is removed
    Private Sub LockCopyColumns(sheet As Worksheet)
        Try
            ' Find the full range and lock it back
            Dim lastRow As Long = sheet.Cells(sheet.Rows.Count, "E").End(XlDirection.xlUp).Row

            ' Lock Column E from row 9 onwards
            Dim rangeE As Range = sheet.Range($"E9:E{lastRow}")
            rangeE.Locked = True

            ' Lock Columns AK, AL, AM from row 9 onwards
            Dim rangeAK As Range = sheet.Range($"AK9:AK{lastRow}")
            Dim rangeAL As Range = sheet.Range($"AL9:AL{lastRow}")
            Dim rangeAM As Range = sheet.Range($"AM9:AM{lastRow}")

            rangeAK.Locked = True
            rangeAL.Locked = True
            rangeAM.Locked = True

        Catch ex As Exception
            MsgBox("Error locking copy columns: " & ex.Message)
        End Try
    End Sub

    Private Sub SafeUnprotectSheet(ws As Worksheet)
        If ws Is Nothing Then Return
        Try
            ws.Unprotect("2022")
            If ws.Parent IsNot Nothing Then
                Dim wb As Workbook = CType(ws.Parent, Workbook)
                wb.Unprotect("2022")
                Marshal.ReleaseComObject(wb)
            End If
        Catch
            ' Ignore protection errors
        End Try
    End Sub

    Private Sub SafeProtectSheet(ws As Worksheet)
        If ws Is Nothing Then Return
        Try
            ws.Protect(Password:="2022", DrawingObjects:=True, Contents:=True, Scenarios:=True)
            If ws.Parent IsNot Nothing Then
                Dim wb As Workbook = CType(ws.Parent, Workbook)
                wb.Protect(Password:="2022", Structure:=True, Windows:=False)
                Marshal.ReleaseComObject(wb)
            End If
        Catch
            ' Ignore protection errors
        End Try
    End Sub

    Private Function GetCellTextSafely(rng As Range) As String
        If rng Is Nothing Then Return String.Empty
        Try
            Return rng.Text.ToString().Trim()
        Catch
            Return String.Empty
        Finally
            If Not rng Is Nothing Then Marshal.ReleaseComObject(rng)
        End Try
    End Function

End Class