Imports System.IO
Imports System.Runtime.InteropServices
Imports Microsoft.Office.Interop.Excel

Namespace RaghavTekNova

    <ComVisible(True)>
    <Guid("A12B34CD-5678-4EF9-ABCD-1234567890AB")> ' ⚠️ Replace with your own GUID
    <ProgId("RaghavTekNova.DuplicateAndUnlockedCheck")>
    Public Class DuplicateAndUnlockedCheck
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

        ' === MASTER: Do both checks & remove duplicates (single popup) ===
        Public Sub A3_Remove_Duplicate_Drawing_Check_Combine()
            Try
                UnlockAllSheets()

                Dim message As String = ""
                message &= A3_Drawing_Check()
                message &= A3_Remove_Duplicate()

                If message <> "" Then MsgBox(message, MsgBoxStyle.Information)

                LockAllSheets()
            Catch ex As Exception
                MsgBox("Error: " & ex.Message, MsgBoxStyle.Critical)
            End Try
        End Sub

        ' === OPTIMIZED: Drawing Highlight Check ===
        Public Function A3_Drawing_Check() As String
            Dim xlApp As Application = CType(Marshal.GetActiveObject("Excel.Application"), Application)
            Dim ws As Worksheet = CType(xlApp.ActiveWorkbook.Sheets("Sheet1"), Worksheet)

            Dim lastRow As Long = ws.Cells(ws.Rows.Count, "E").End(XlDirection.xlUp).Row

            ' Read all data into arrays for faster processing
            Dim colMValues As Object(,) = ws.Range("M8:M" & lastRow).Value2
            Dim colNValues As Object(,) = ws.Range("N8:N" & lastRow).Value2

            ' Get references to ranges for batch formatting
            Dim colERange As Range = ws.Range("E8:E" & lastRow)

            Dim highlightCountGolden As Long = 0
            Dim highlightCountGreen As Long = 0

            ' Store formatting changes to apply in batch
            Dim goldenRows As New List(Of Long)
            Dim greenRows As New List(Of Long)

            For i As Long = 1 To lastRow - 7 ' Array is 1-based, rows start at 8
                Dim colMValue As String = If(colMValues(i, 1) Is Nothing, "", colMValues(i, 1).ToString())
                Dim colNValue As String = If(colNValues(i, 1) Is Nothing, "", colNValues(i, 1).ToString())

                ' Optimized string cleaning
                colMValue = colMValue.Trim().Replace(ChrW(160), "")
                colNValue = colNValue.Trim().Replace(ChrW(160), "")

                If colMValue <> "" AndAlso colNValue = "0" Then
                    goldenRows.Add(i + 7) ' Convert back to actual row number
                    highlightCountGolden += 1
                ElseIf colMValue = "" AndAlso colNValue = "0" Then
                    greenRows.Add(i + 7)
                    highlightCountGreen += 1
                End If
            Next

            ' Apply formatting in batches
            If goldenRows.Count > 0 Then
                For Each rowNum In goldenRows
                    Dim cell As Range = ws.Cells(rowNum, 5)
                    cell.Font.Color = RGB(192, 0, 0)
                    cell.Interior.Color = RGB(255, 217, 102)
                Next
            End If

            If greenRows.Count > 0 Then
                For Each rowNum In greenRows
                    Dim cell As Range = ws.Cells(rowNum, 5)
                    cell.Font.Color = RGB(192, 0, 0)
                    cell.Interior.Color = RGB(226, 239, 218)
                Next
            End If

            ' Build message string
            Dim message As String = ""
            If highlightCountGolden > 0 Then message &= highlightCountGolden & " Drawings are unlocked in the Tekla model!" & vbCrLf
            If highlightCountGreen > 0 Then message &= highlightCountGreen & " Drawings are still missing in the Tekla model!" & vbCrLf

            Return message
        End Function

        ' === OPTIMIZED: Remove Duplicates ===
        Public Function A3_Remove_Duplicate() As String
            Dim xlApp As Application = CType(Marshal.GetActiveObject("Excel.Application"), Application)
            Dim ws As Worksheet = CType(xlApp.ActiveWorkbook.Sheets("Sheet1"), Worksheet)

            Dim lastRow As Long = ws.Cells(ws.Rows.Count, "E").End(XlDirection.xlUp).Row

            ' Read all data into arrays for faster processing
            Dim assemblyMarks As Object(,) = ws.Range("E8:E" & lastRow).Value2
            Dim lengthValues As Object(,) = ws.Range("F8:F" & lastRow).Value2
            Dim qtyValues As Object(,) = ws.Range("G8:G" & lastRow).Value2
            Dim unitWeightValues As Object(,) = ws.Range("H8:H" & lastRow).Value2
            Dim totalWeightValues As Object(,) = ws.Range("I8:I" & lastRow).Value2
            Dim unitAreaValues As Object(,) = ws.Range("J8:J" & lastRow).Value2
            Dim totalAreaValues As Object(,) = ws.Range("K8:K" & lastRow).Value2
            Dim partsInAssemblyValues As Object(,) = ws.Range("L8:L" & lastRow).Value2

            Dim duplicateCount As Long = 0
            Dim serialNum As Long = 1
            Dim rowsToDelete As New List(Of Long)

            ' Dictionary to track processed assembly marks and their consolidated data
            Dim processedMarks As New Dictionary(Of String, Integer) ' Mark -> First occurrence row index

            For i As Long = 1 To lastRow - 7 ' Array is 1-based
                Dim assemblyMark As String = If(assemblyMarks(i, 1) Is Nothing, "", assemblyMarks(i, 1).ToString())

                If assemblyMark <> "" Then
                    Dim actualRowNum As Long = i + 7 ' Convert to actual Excel row

                    If processedMarks.ContainsKey(assemblyMark) Then
                        ' This is a duplicate - consolidate values with the first occurrence
                        Dim firstRowIndex As Integer = processedMarks(assemblyMark)
                        Dim firstActualRow As Long = firstRowIndex + 7

                        ' Aggregate values
                        Dim currentQty As Double = GetDoubleFromArray(qtyValues(firstRowIndex, 1))
                        Dim newQty As Double = GetDoubleFromArray(qtyValues(i, 1))
                        ws.Cells(firstActualRow, 7).Value = currentQty + newQty

                        Dim currentTotalWeight As Double = GetDoubleFromArray(totalWeightValues(firstRowIndex, 1))
                        Dim newTotalWeight As Double = GetDoubleFromArray(totalWeightValues(i, 1))
                        ws.Cells(firstActualRow, 9).Value = currentTotalWeight + newTotalWeight

                        Dim currentTotalArea As Double = GetDoubleFromArray(totalAreaValues(firstRowIndex, 1))
                        Dim newTotalArea As Double = GetDoubleFromArray(totalAreaValues(i, 1))
                        ws.Cells(firstActualRow, 11).Value = currentTotalArea + newTotalArea

                        ' Max values
                        Dim currentLength As Double = GetDoubleFromArray(lengthValues(firstRowIndex, 1))
                        Dim newLength As Double = GetDoubleFromArray(lengthValues(i, 1))
                        ws.Cells(firstActualRow, 6).Value = Math.Max(currentLength, newLength)

                        Dim currentUnitWeight As Double = GetDoubleFromArray(unitWeightValues(firstRowIndex, 1))
                        Dim newUnitWeight As Double = GetDoubleFromArray(unitWeightValues(i, 1))
                        ws.Cells(firstActualRow, 8).Value = Math.Max(currentUnitWeight, newUnitWeight)

                        Dim currentUnitArea As Double = GetDoubleFromArray(unitAreaValues(firstRowIndex, 1))
                        Dim newUnitArea As Double = GetDoubleFromArray(unitAreaValues(i, 1))
                        ws.Cells(firstActualRow, 10).Value = Math.Max(currentUnitArea, newUnitArea)

                        Dim currentPartsInAssembly As Double = GetDoubleFromArray(partsInAssemblyValues(firstRowIndex, 1))
                        Dim newPartsInAssembly As Double = GetDoubleFromArray(partsInAssemblyValues(i, 1))
                        ws.Cells(firstActualRow, 12).Value = Math.Max(currentPartsInAssembly, newPartsInAssembly)

                        ' Mark row for deletion
                        rowsToDelete.Add(actualRowNum)
                        duplicateCount += 1
                    Else
                        ' First occurrence of this assembly mark
                        processedMarks.Add(assemblyMark, i)
                    End If
                End If
            Next

            ' Delete duplicate rows (in reverse order to avoid index shifting)
            If rowsToDelete.Count > 0 Then
                rowsToDelete.Sort()
                For i As Integer = rowsToDelete.Count - 1 To 0 Step -1
                    ws.Rows(rowsToDelete(i)).Delete()
                Next
            End If

            ' Update serial numbers in one pass
            Dim newLastRow As Long = ws.Cells(ws.Rows.Count, "E").End(XlDirection.xlUp).Row
            If newLastRow >= 8 Then
                Dim serialRange As Range = ws.Range("A8:A" & newLastRow)
                Dim serialArray(newLastRow - 7, 0) As Object
                For i As Integer = 0 To newLastRow - 8
                    serialArray(i, 0) = i + 1
                Next
                serialRange.Value = serialArray
            End If

            ' Build message string
            If duplicateCount > 0 Then
                Return duplicateCount & " Duplicate Assembly Marks Removed & Merged!"
            End If

            Return ""
        End Function

        ' === Helper for numeric conversion from array values ===
        Private Function GetDoubleFromArray(val As Object) As Double
            If val Is Nothing OrElse IsDBNull(val) Then Return 0
            Dim str As String = val.ToString().Trim()
            If str = "" Then Return 0
            If IsNumeric(str) Then
                Return CDbl(str)
            End If
            Return 0
        End Function

        ' === OPTIMIZED: Toggle Filter for Backfilled Data ===
        Public Sub A3_Simple_Toggle_Filter()
            Try
                UnlockAllSheets()

                Dim xlApp As Application = CType(Marshal.GetActiveObject("Excel.Application"), Application)
                Dim ws As Worksheet = CType(xlApp.ActiveWorkbook.Sheets("Sheet1"), Worksheet)

                ' Check if filter is already applied
                If IsFilterApplied(ws) Then
                    RemoveAllFilters(ws)
                    MsgBox("Filter removed - All data visible!", MsgBoxStyle.Information)
                Else
                    ApplyColorBasedFilter(ws)
                End If

                LockAllSheets()
            Catch ex As Exception
                MsgBox("Error: " & ex.Message, MsgBoxStyle.Critical)
            End Try
        End Sub

        ' === OPTIMIZED: Check if filter is applied ===
        Private Function IsFilterApplied(ws As Worksheet) As Boolean
            Dim lastRow As Long = ws.Cells(ws.Rows.Count, "E").End(XlDirection.xlUp).Row
            If lastRow < 8 Then Return False

            ' Check first few and last few rows for hidden status (faster than checking all)
            Dim checkRows() As Long = {8, 9, 10, Math.Max(8, lastRow - 2), Math.Max(8, lastRow - 1), lastRow}

            For Each rowNum In checkRows
                If rowNum <= lastRow AndAlso ws.Rows(rowNum).Hidden Then
                    Return True
                End If
            Next

            Return False
        End Function

        ' === Remove all filters and show all rows ===
        Private Sub RemoveAllFilters(ws As Worksheet)
            ' Remove Excel AutoFilter if present
            If ws.AutoFilterMode Then
                ws.AutoFilterMode = False
            End If

            ' Unhide all rows using range operation
            Dim lastRow As Long = ws.Cells(ws.Rows.Count, "E").End(XlDirection.xlUp).Row
            If lastRow >= 8 Then
                ws.Range("A8:A" & lastRow).EntireRow.Hidden = False
            End If
        End Sub

        ' === OPTIMIZED: Apply filter based on text color ===
        Private Sub ApplyColorBasedFilter(ws As Worksheet)
            Try
                Dim lastRow As Long = ws.Cells(ws.Rows.Count, "E").End(XlDirection.xlUp).Row
                If lastRow < 8 Then
                    MsgBox("No data to filter!", MsgBoxStyle.Exclamation)
                    Exit Sub
                End If

                ' Remove any existing AutoFilter
                If ws.AutoFilterMode Then
                    ws.AutoFilterMode = False
                End If

                Dim visibleCount As Integer = 0
                Dim targetColor As Integer = RGB(192, 0, 0)

                ' Process in chunks for better performance
                Dim chunkSize As Integer = 100
                Dim currentRow As Long = 8

                While currentRow <= lastRow
                    Dim endRow As Long = Math.Min(currentRow + chunkSize - 1, lastRow)

                    For i As Long = currentRow To endRow
                        Dim cell As Range = ws.Cells(i, 5)

                        If cell.Font.Color = targetColor Then
                            ws.Rows(i).Hidden = False
                            visibleCount += 1
                        Else
                            ws.Rows(i).Hidden = True
                        End If
                    Next

                    currentRow = endRow + 1
                End While

                If visibleCount > 0 Then
                    MsgBox("Filter applied - " & visibleCount & " highlighted rows shown!" & vbCrLf &
                   "(Rows with red text in Assembly Mark column)", MsgBoxStyle.Information)
                Else
                    RemoveAllFilters(ws)
                    MsgBox("No highlighted rows found! Showing all data.", MsgBoxStyle.Information)
                End If

            Catch ex As Exception
                MsgBox("Color filter error: " & ex.Message, MsgBoxStyle.Critical)
                RemoveAllFilters(ws)
            End Try
        End Sub

        ' === Enhanced Drawing Check ===
        Public Function A3_Drawing_Check_Enhanced() As String
            Dim message As String = A3_Drawing_Check()

            If message <> "" Then
                message &= vbCrLf & "Tip: Use A3_Simple_Toggle_Filter() to view only highlighted rows!"
            End If

            Return message
        End Function

        ' === Helper for numeric conversion (original) ===
        Private Function GetDouble(val As Object) As Double
            If val Is Nothing OrElse IsDBNull(val) Then Return 0
            Dim str As String = val.ToString().Trim()
            If str = "" Then Return 0
            If IsNumeric(str) Then
                Return CDbl(str)
            End If
            Return 0
        End Function

        ' === DEBUG: Method to check text colors (optimized) ===
        Public Sub DebugTextColors()
            Try
                Dim xlApp As Application = CType(Marshal.GetActiveObject("Excel.Application"), Application)
                Dim ws As Worksheet = CType(xlApp.ActiveWorkbook.Sheets("Sheet1"), Worksheet)

                Dim lastRow As Long = ws.Cells(ws.Rows.Count, "E").End(XlDirection.xlUp).Row
                Dim redTextCount As Integer = 0
                Dim expectedColor As Integer = RGB(192, 0, 0)

                ' Check range of cells at once
                Dim checkRange As Range = ws.Range("E8:E" & Math.Min(15, lastRow))

                For Each cell As Range In checkRange
                    Dim fontColor As Integer = cell.Font.Color

                    If fontColor = expectedColor Then
                        redTextCount += 1
                    End If
                Next

                MsgBox("Found " & redTextCount & " rows with red text in first " & (Math.Min(15, lastRow) - 7) & " rows", MsgBoxStyle.Information)

            Catch ex As Exception
                MsgBox("Debug error: " & ex.Message, MsgBoxStyle.Critical)
            End Try
        End Sub

        ' === Sheet Unlock (optimized) ===
        Public Sub UnlockAllSheets()
            Dim xlApp As Application = CType(Marshal.GetActiveObject("Excel.Application"), Application)
            Dim currentSheet As Worksheet = CType(xlApp.ActiveSheet, Worksheet)

            ' Minimal performance settings - keep screen updating enabled
            xlApp.EnableEvents = False

            For Each ws As Worksheet In xlApp.ActiveWorkbook.Worksheets
                ws.Unprotect("2022")
            Next

            currentSheet.Activate()
            xlApp.EnableEvents = True
        End Sub

        ' === Sheet Lock (optimized) ===
        Public Sub LockAllSheets()
            Dim xlApp As Application = CType(Marshal.GetActiveObject("Excel.Application"), Application)
            Dim currentSheet As Worksheet = CType(xlApp.ActiveSheet, Worksheet)

            ' Minimal performance settings - keep screen updating enabled
            xlApp.EnableEvents = False

            For Each ws As Worksheet In xlApp.ActiveWorkbook.Worksheets
                ws.Protect("2022", UserInterfaceOnly:=True)
            Next

            currentSheet.Activate()
            xlApp.EnableEvents = True
        End Sub

    End Class

End Namespace