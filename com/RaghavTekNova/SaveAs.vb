Imports System
Imports System.IO
Imports System.Runtime.InteropServices
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Windows.Forms
Imports System.Collections.Generic
Imports Microsoft.VisualBasic
Imports System.Linq

<ComVisible(True)>
<ClassInterface(ClassInterfaceType.AutoDual)>
<ProgId("RaghavTekNova.SaveAs")>
Public Class SaveAs
    Implements IDisposable

    Private xlApp As Excel.Application = Nothing
    Private isAppCreatedInternally As Boolean = False

    Private Const BOQ_FOLDER As String = "BOQ"
    Private Const BREAKUP_FOLDER As String = "SaveAsBreakup"

#Region "Initialization and Cleanup"

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
    Public Sub SaveSheet1AsAssemblyList()
        SaveSheetAsFile("Sheet1", "Assembly List.xlsx", "A2:B5", True)
    End Sub

    <ComVisible(True)>
    Public Sub SaveSheet2AsBoltList()
        SaveSheet2AsBoltListSafe()
    End Sub

    <ComVisible(True)>
    Public Sub SaveSheet3AsBOQSummary()
        SaveSheetAsFile("Sheet3", "BOQ Summary List.xlsx", "A2:B5", False)
    End Sub

    <ComVisible(True)>
    Public Sub SaveSheet4AsPartList()
        SaveSheetAsFile("Sheet4", "Part List.xlsx", "A2:B5", True)
    End Sub

    <ComVisible(True)>
    Public Sub SaveSheet5AsMaterialSummary()
        SaveSheetAsFile("Sheet5", "Material Summary List.xlsx", "A2:B5", False)
    End Sub

    Public Sub SaveSheet10AsDrawingTracking()
        SaveSheetAsFile("Sheet10", "Drawing Tracking.xlsx", "A2:B5", True)
    End Sub

    <ComVisible(True)>
    Public Sub SaveCurrentSheet()
        InitializeExcelApplication()
        If xlApp Is Nothing OrElse xlApp.ActiveWorkbook Is Nothing Then
            MessageBox.Show("Excel is not available or no workbook is open.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        Dim originalWb As Excel.Workbook = Nothing
        Dim activeWs As Excel.Worksheet = Nothing
        Dim tempWb As Excel.Workbook = Nothing
        Dim tempWs As Excel.Worksheet = Nothing

        Try
            originalWb = xlApp.ActiveWorkbook
            Dim originalPath As String = If(originalWb.Path IsNot Nothing, originalWb.Path, String.Empty)
            If String.IsNullOrEmpty(originalPath) Then
                MessageBox.Show("Please save the workbook before using this feature.", "Workbook Not Saved", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            ' Safely get the active sheet
            Try
                activeWs = CType(xlApp.ActiveSheet, Excel.Worksheet)
            Catch ex As Exception
                MessageBox.Show("Could not access the active sheet. Please select a sheet and try again.", "Sheet Access Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End Try

            If activeWs Is Nothing Then
                MessageBox.Show("No active sheet found. Please select a sheet and try again.", "No Active Sheet", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            Dim sheetName As String = activeWs.Name
            If String.IsNullOrEmpty(sheetName) Then
                sheetName = "CurrentSheet"
            End If

            ' Check if the active sheet is Sheet1
            If sheetName.Equals("Sheet1", StringComparison.OrdinalIgnoreCase) Then
                MessageBox.Show("Use Save As button instead.", "Operation Not Allowed", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            ' Get file name from cells D2, D3, J3
            Dim d2Value As String = SafeGetCellValue(activeWs, "D2")
            Dim d3Value As String = SafeGetCellValue(activeWs, "D3")
            Dim j3Value As String = SafeGetCellValue(activeWs, "J3")

            ' Construct file name (use sheet name as fallback if all cells are empty)
            Dim fileNameParts As New List(Of String)
            If Not String.IsNullOrEmpty(d2Value) Then fileNameParts.Add(d2Value)
            If Not String.IsNullOrEmpty(d3Value) Then fileNameParts.Add(d3Value)
            If Not String.IsNullOrEmpty(j3Value) Then fileNameParts.Add(j3Value)

            Dim outputFileName As String
            If fileNameParts.Count > 0 Then
                outputFileName = CleanFileName(String.Join("_", fileNameParts)) & ".xlsx"
            Else
                outputFileName = CleanFileName(sheetName) & ".xlsx"
            End If

            xlApp.ScreenUpdating = False
            xlApp.DisplayAlerts = False

            originalWb.Activate()

            ' Handle Sheet2 specially (as per your existing logic)
            If sheetName.Equals("Sheet2", StringComparison.OrdinalIgnoreCase) Then
                tempWb = xlApp.Workbooks.Add()
                tempWs = CType(tempWb.Sheets(1), Excel.Worksheet)

                Dim newSheetName As String = If(outputFileName.Length <= 31, Path.GetFileNameWithoutExtension(outputFileName), Path.GetFileNameWithoutExtension(outputFileName).Substring(0, 31))
                tempWs.Name = newSheetName

                CopySheetContentManually(activeWs, tempWs)
            Else
                tempWb = xlApp.Workbooks.Add()
                activeWs.Copy(Before:=tempWb.Sheets(1))

                Dim defaultSheet As Excel.Worksheet = CType(tempWb.Sheets(2), Excel.Worksheet)
                defaultSheet.Delete()
                SafeReleaseComObject(defaultSheet)

                tempWs = CType(tempWb.Sheets(1), Excel.Worksheet)
            End If

            Dim folderPath As String = Path.Combine(originalPath, BREAKUP_FOLDER)
            Directory.CreateDirectory(folderPath)

            Dim fullFilePath As String = Path.Combine(folderPath, outputFileName)
            fullFilePath = HandleFileExists(fullFilePath)
            If String.IsNullOrEmpty(fullFilePath) Then
                tempWb.Close(False)
                Return
            End If

            ' 🔥 FIRST: Remove all form controls, shapes, and buttons
            SafeRemoveAllControlsAndShapes(tempWs)

            ' 🔓 SECOND: Unlock the worksheet (remove password "2022")
            SafeUnlockWorksheet(tempWs, "2022")

            ' 💾 THIRD: Save the workbook
            tempWb.SaveAs(fullFilePath, Excel.XlFileFormat.xlOpenXMLWorkbook)

            ' 🖼️ FOURTH: Insert logo if needed
            If Not ShouldSkipLogo(sheetName) Then
                SafeInsertCompanyLogo(tempWb, originalPath, "A2:B5")
                tempWb.Save()
            End If

            tempWb.Close(False)

            originalWb.Activate()
            activeWs.Select()

            MessageBox.Show($"Current sheet '{sheetName}' saved as '{Path.GetFileName(fullFilePath)}' in the '{BREAKUP_FOLDER}' folder.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            MessageBox.Show($"Error saving current sheet: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If xlApp IsNot Nothing Then
                Try
                    xlApp.ScreenUpdating = False
                    xlApp.DisplayAlerts = True
                Catch
                End Try
            End If

            SafeReleaseComObject(tempWs)
            SafeReleaseComObject(tempWb)

            If originalWb IsNot Nothing Then
                Try
                    originalWb.Activate()
                    If activeWs IsNot Nothing Then activeWs.Select()
                Catch
                End Try
            End If

            SafeReleaseComObject(activeWs)
            SafeReleaseComObject(originalWb)
        End Try
    End Sub

#End Region

#Region "Private Worker and Helper Methods"

    Private Sub SaveSheet2AsBoltListSafe()
        InitializeExcelApplication()
        If xlApp Is Nothing OrElse xlApp.ActiveWorkbook Is Nothing Then
            MessageBox.Show("Excel is not available or no workbook is open.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        Dim originalWb As Excel.Workbook = Nothing
        Dim originalWs As Excel.Worksheet = Nothing
        Dim tempWb As Excel.Workbook = Nothing
        Dim tempWs As Excel.Worksheet = Nothing

        Try
            originalWb = xlApp.ActiveWorkbook
            Dim originalPath As String = If(originalWb.Path IsNot Nothing, originalWb.Path, String.Empty)
            If String.IsNullOrEmpty(originalPath) Then
                MessageBox.Show("Please save the workbook before using this feature.", "Workbook Not Saved", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            originalWs = SafeGetWorksheet(originalWb, "Sheet2")
            If originalWs Is Nothing Then
                MessageBox.Show("Sheet2 was not found or cannot be accessed.", "Sheet Not Found", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            xlApp.ScreenUpdating = False
            xlApp.DisplayAlerts = False

            originalWb.Activate()

            tempWb = xlApp.Workbooks.Add()
            tempWs = CType(tempWb.Sheets(1), Excel.Worksheet)
            tempWs.Name = "Bolt List"

            CopySheetContentManually(originalWs, tempWs)

            Dim folderPath As String = Path.Combine(originalPath, BOQ_FOLDER)
            Directory.CreateDirectory(folderPath)

            Dim fullFilePath As String = Path.Combine(folderPath, "Bolt List.xlsx")
            fullFilePath = HandleFileExists(fullFilePath)
            If String.IsNullOrEmpty(fullFilePath) Then
                tempWb.Close(False)
                Return
            End If

            ' 🔥 FIRST: Remove all form controls, shapes, and buttons
            SafeRemoveAllControlsAndShapes(tempWs)

            ' 🔓 SECOND: Unlock the worksheet (remove password "2022")
            SafeUnlockWorksheet(tempWs, "2022")

            ' 💾 THIRD: Save the workbook
            tempWb.SaveAs(fullFilePath, Excel.XlFileFormat.xlOpenXMLWorkbook)

            ' 🖼️ FOURTH: Insert logo
            SafeInsertCompanyLogo(tempWb, originalPath, "A2:B5")
            tempWb.Save()

            tempWb.Close(False)

            originalWb.Activate()
            If originalWs IsNot Nothing Then originalWs.Select()

            MessageBox.Show($"Bolt List saved as '{Path.GetFileName(fullFilePath)}' in the '{BOQ_FOLDER}' folder." & vbCrLf & "📋 Only columns A to K were included.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            MessageBox.Show($"Error saving Bolt List: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If xlApp IsNot Nothing Then
                Try
                    xlApp.ScreenUpdating = False
                    xlApp.DisplayAlerts = True
                Catch
                End Try
            End If

            SafeReleaseComObject(tempWs)
            SafeReleaseComObject(tempWb)

            If originalWb IsNot Nothing Then
                Try
                    originalWb.Activate()
                    If originalWs IsNot Nothing Then originalWs.Select()
                Catch
                End Try
            End If

            SafeReleaseComObject(originalWs)
            SafeReleaseComObject(originalWb)
        End Try
    End Sub

    Private Sub CopySheetContentManually(sourceSheet As Excel.Worksheet, targetSheet As Excel.Worksheet)
        Try
            Dim lastRow As Integer = 1
            Dim lastCol As Integer = 1

            Try
                lastRow = sourceSheet.UsedRange.Rows.Count
                lastCol = sourceSheet.UsedRange.Columns.Count
            Catch
                lastRow = 100
                lastCol = 20
            End Try

            For row As Integer = 1 To lastRow
                For col As Integer = 1 To lastCol
                    Try
                        Dim sourceCell As Excel.Range = sourceSheet.Cells(row, col)
                        Dim targetCell As Excel.Range = targetSheet.Cells(row, col)

                        targetCell.Value = sourceCell.Value
                        targetCell.NumberFormat = sourceCell.NumberFormat

                        SafeReleaseComObject(sourceCell)
                        SafeReleaseComObject(targetCell)
                    Catch
                    End Try
                Next
            Next

            For col As Integer = 1 To lastCol
                Try
                    targetSheet.Columns(col).ColumnWidth = sourceSheet.Columns(col).ColumnWidth
                Catch
                End Try
            Next

            For row As Integer = 1 To lastRow
                Try
                    targetSheet.Rows(row).RowHeight = sourceSheet.Rows(row).RowHeight
                Catch
                End Try
            Next

        Catch ex As Exception
            MessageBox.Show($"Warning: Some content may not have copied correctly: {ex.Message}", "Partial Copy", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Try
    End Sub

    Private Sub SaveSheetAsFile(sheetName As String, outputFileName As String, logoRangeAddress As String, shouldInsertLogo As Boolean)
        InitializeExcelApplication()
        If xlApp Is Nothing OrElse xlApp.ActiveWorkbook Is Nothing Then
            MessageBox.Show("Excel is not available or no workbook is open.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        Dim originalWb As Excel.Workbook = Nothing
        Dim originalWs As Excel.Worksheet = Nothing
        Dim tempWb As Excel.Workbook = Nothing
        Dim tempWs As Excel.Worksheet = Nothing

        Try
            originalWb = xlApp.ActiveWorkbook
            Dim originalPath As String = If(originalWb.Path IsNot Nothing, originalWb.Path, String.Empty)
            If String.IsNullOrEmpty(originalPath) Then
                MessageBox.Show("Please save the workbook before using this feature.", "Workbook Not Saved", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            originalWs = SafeGetWorksheet(originalWb, sheetName)
            If originalWs Is Nothing Then
                MessageBox.Show($"The sheet '{sheetName}' was not found or cannot be accessed.", "Sheet Not Found", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            xlApp.ScreenUpdating = False
            xlApp.DisplayAlerts = False

            originalWb.Activate()

            tempWb = xlApp.Workbooks.Add()
            originalWs.Copy(Before:=tempWb.Sheets(1))

            Dim defaultSheet As Excel.Worksheet = CType(tempWb.Sheets(2), Excel.Worksheet)
            defaultSheet.Delete()

            tempWs = CType(tempWb.Sheets(1), Excel.Worksheet)

            Dim folderPath As String = Path.Combine(originalPath, BOQ_FOLDER)
            Directory.CreateDirectory(folderPath)

            Dim fullFilePath As String = Path.Combine(folderPath, outputFileName)
            fullFilePath = HandleFileExists(fullFilePath)
            If String.IsNullOrEmpty(fullFilePath) Then
                tempWb.Close(False)
                Return
            End If

            ' 🔥 FIRST: Remove all form controls, shapes, and buttons
            SafeRemoveAllControlsAndShapes(tempWs)

            ' 🔓 SECOND: Unlock the worksheet (remove password "2022")
            SafeUnlockWorksheet(tempWs, "2022")

            ' 💾 THIRD: Save the workbook
            tempWb.SaveAs(fullFilePath, Excel.XlFileFormat.xlOpenXMLWorkbook)

            ' 🖼️ FOURTH: Insert logo if needed
            If shouldInsertLogo Then
                SafeInsertCompanyLogo(tempWb, originalPath, logoRangeAddress)
                tempWb.Save()
            End If

            tempWb.Close(False)

            originalWb.Activate()
            originalWs.Select()

            MessageBox.Show($"File saved as '{Path.GetFileName(fullFilePath)}' in the '{BOQ_FOLDER}' folder.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            MessageBox.Show($"Error saving {sheetName}: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If xlApp IsNot Nothing Then
                Try
                    xlApp.ScreenUpdating = False
                    xlApp.DisplayAlerts = True
                Catch
                End Try
            End If

            SafeReleaseComObject(tempWs)
            SafeReleaseComObject(tempWb)

            If originalWb IsNot Nothing Then
                Try
                    originalWb.Activate()
                    If originalWs IsNot Nothing Then originalWs.Select()
                Catch
                End Try
            End If

            SafeReleaseComObject(originalWs)
            SafeReleaseComObject(originalWb)
        End Try
    End Sub

    ' 🆕 ENHANCED: Remove all controls, shapes AND ensure no hidden columns are saved
    Private Sub SafeRemoveAllControlsAndShapes(ws As Excel.Worksheet)
        If ws Is Nothing Then Return

        Try
            ' 🔥 Step 1: Remove all Shapes (including form controls, buttons, images, etc.)
            Dim shapes As Excel.Shapes = ws.Shapes
            Dim shapeCount As Integer = shapes.Count

            ' Remove shapes in reverse order to avoid index shifting issues
            For i As Integer = shapeCount To 1 Step -1
                Try
                    Dim shp As Excel.Shape = shapes.Item(i)

                    ' Clear any action assignments before deletion
                    If Not String.IsNullOrEmpty(shp.OnAction) Then
                        shp.OnAction = String.Empty
                    End If

                    ' Delete the shape
                    shp.Delete()
                    SafeReleaseComObject(shp)
                Catch ex As Exception
                    ' Continue even if individual shape deletion fails
                    Console.WriteLine($"Failed to delete shape {i}: {ex.Message}")
                End Try
            Next
            SafeReleaseComObject(shapes)

            ' 🔥 Step 2: Remove OLE Objects (embedded objects, ActiveX controls)
            Try
                Dim oleObjects As Excel.OLEObjects = ws.OLEObjects
                Dim oleCount As Integer = oleObjects.Count

                ' Remove OLE objects in reverse order
                For i As Integer = oleCount To 1 Step -1
                    Try
                        oleObjects.Item(i).Delete()
                    Catch ex As Exception
                        Console.WriteLine($"Failed to delete OLE object {i}: {ex.Message}")
                    End Try
                Next
                SafeReleaseComObject(oleObjects)
            Catch ex As Exception
                Console.WriteLine($"Failed to access OLE objects: {ex.Message}")
            End Try

            ' 🔥 Step 3: Clear any remaining OnAction properties from ranges
            Try
                Dim usedRange As Excel.Range = ws.UsedRange
                If usedRange IsNot Nothing Then
                    ' Remove any hyperlinks that might have actions
                    usedRange.Hyperlinks.Delete()
                End If
                SafeReleaseComObject(usedRange)
            Catch ex As Exception
                Console.WriteLine($"Failed to clear hyperlinks: {ex.Message}")
            End Try

            ' 🆕 Step 4: FORCIBLY DELETE columns L,M,N,O,P and beyond
            Try
                ' Clear and delete columns from L (12) onwards
                For col As Integer = 12 To 50  ' Clear L,M,N,O,P and more to be safe
                    Try
                        ws.Columns(col).Clear()
                        ws.Columns(col).Hidden = False
                    Catch
                        ' Continue if column clear fails
                    End Try
                Next

                ' Try to delete the entire range from column L onwards
                Try
                    Dim deleteRange As Excel.Range = ws.Range("L:XFD")  ' L to last Excel column
                    deleteRange.Clear()
                    SafeReleaseComObject(deleteRange)
                Catch
                    ' Deletion might fail, that's OK as we cleared the content
                End Try

            Catch ex As Exception
                Console.WriteLine($"Failed to clear columns L onwards: {ex.Message}")
            End Try

        Catch ex As Exception
            Console.WriteLine($"Error in SafeRemoveAllControlsAndShapes: {ex.Message}")
        End Try
    End Sub

    ' 🆕 NEW METHOD: Unlock worksheet with password
    Private Sub SafeUnlockWorksheet(ws As Excel.Worksheet, password As String)
        If ws Is Nothing Then Return

        Try
            ' Check if the worksheet is protected
            If ws.ProtectContents Or ws.ProtectDrawingObjects Or ws.ProtectScenarios Then
                ' Attempt to unprotect with the known password
                ws.Unprotect(password)

                ' Verify it's now unprotected
                If Not ws.ProtectContents Then
                    Console.WriteLine("Worksheet successfully unprotected.")
                End If
            End If
        Catch ex As Exception
            ' If unprotect fails, try without password (in case it was protected without password)
            Try
                ws.Unprotect()
                Console.WriteLine("Worksheet unprotected without password.")
            Catch ex2 As Exception
                Console.WriteLine($"Failed to unprotect worksheet: {ex2.Message}")
                ' Continue anyway - the sheet might not be protected or might have a different password
            End Try
        End Try
    End Sub

    Private Function SafeGetWorksheet(wb As Excel.Workbook, sheetName As String) As Excel.Worksheet
        Try
            Return CType(wb.Sheets(sheetName), Excel.Worksheet)
        Catch
            Return Nothing
        End Try
    End Function

    Private Function SafeGetCellValue(ws As Excel.Worksheet, cellAddress As String) As String
        Try
            Dim cellValue As Object = ws.Range(cellAddress).Value
            Return If(cellValue IsNot Nothing, cellValue.ToString(), String.Empty)
        Catch
            Return String.Empty
        End Try
    End Function

    Private Sub SafeInsertCompanyLogo(wb As Excel.Workbook, sourceFolderPath As String, logoRangeAddress As String)
        Dim ws As Excel.Worksheet = Nothing
        Dim logoRange As Excel.Range = Nothing
        Dim logoPic As Excel.Shape = Nothing

        Try
            If Not Directory.Exists(sourceFolderPath) Then Return

            Dim imageExtensions = {"*.jpg", "*.jpeg", "*.png", "*.bmp"}
            Dim imageFiles As New List(Of String)

            For Each ext In imageExtensions
                Try
                    imageFiles.AddRange(Directory.GetFiles(sourceFolderPath, ext))
                Catch
                End Try
            Next

            If imageFiles.Count = 0 Then Return

            Dim logoPath As String = If(imageFiles.Count > 1, ShowLogoSelectionDialog(imageFiles), imageFiles(0))
            If String.IsNullOrEmpty(logoPath) OrElse Not File.Exists(logoPath) Then Return

            ws = CType(wb.Sheets(1), Excel.Worksheet)
            logoRange = ws.Range(logoRangeAddress)
            Const margin As Double = 3.0

            logoPic = ws.Shapes.AddPicture(logoPath, False, True, 0, 0, -1, -1)

            Dim availableWidth = logoRange.Width - (2 * margin)
            Dim availableHeight = logoRange.Height - (2 * margin)

            With logoPic
                .LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoTrue
                If .Width > availableWidth Then .Width = availableWidth
                If .Height > availableHeight Then .Height = availableHeight
                .Left = logoRange.Left + margin + (availableWidth - .Width) / 2
                .Top = logoRange.Top + margin + (availableHeight - .Height) / 2
            End With

        Catch
        Finally
            SafeReleaseComObject(logoPic)
            SafeReleaseComObject(logoRange)
            SafeReleaseComObject(ws)
        End Try
    End Sub

    Private Function ShouldSkipLogo(sheetName As String) As Boolean
        Dim skipSheets() As String = {"Sheet3", "Sheet5", "BOQ Summary", "Material Summary"}
        Return skipSheets.Contains(sheetName)
    End Function

    Private Function HandleFileExists(filePath As String) As String
        If Not File.Exists(filePath) Then Return filePath

        Dim fileName = Path.GetFileName(filePath)
        Dim msg = $"The file '{fileName}' already exists." & vbCrLf & vbCrLf &
                  "Click YES to overwrite it." & vbCrLf &
                  "Click NO to save it with a new name." & vbCrLf &
                  "Click CANCEL to abort the operation."

        Dim result = MessageBox.Show(msg, "File Exists", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question)

        Select Case result
            Case DialogResult.Yes
                Try
                    File.Delete(filePath)
                    Return filePath
                Catch ex As Exception
                    MessageBox.Show($"Could not overwrite the file. Error: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Return String.Empty
                End Try
            Case DialogResult.No
                Return GetUniqueFilePath(Path.GetDirectoryName(filePath), fileName)
            Case DialogResult.Cancel, DialogResult.Abort
                Return String.Empty
            Case Else
                Return String.Empty
        End Select
    End Function

    Private Function GetUniqueFilePath(folderPath As String, fileName As String) As String
        Dim newPath = Path.Combine(folderPath, fileName)
        Dim nameWithoutExt = Path.GetFileNameWithoutExtension(fileName)
        Dim ext = Path.GetExtension(fileName)
        Dim i As Integer = 1

        While File.Exists(newPath)
            Dim newFileName = $"{nameWithoutExt} ({i}){ext}"
            newPath = Path.Combine(folderPath, newFileName)
            i += 1
        End While

        Return newPath
    End Function

    Private Function ShowLogoSelectionDialog(imageFiles As List(Of String)) As String
        Try
            Dim prompt As New System.Text.StringBuilder("Multiple logo images found. Please select one by number:")
            prompt.AppendLine().AppendLine()
            For i = 0 To imageFiles.Count - 1
                prompt.AppendLine($"{i + 1}: {Path.GetFileName(imageFiles(i))}")
            Next

            Dim input = Interaction.InputBox(prompt.ToString(), "Select Logo File", "1")

            Dim selection As Integer
            If Integer.TryParse(input, selection) AndAlso selection >= 1 AndAlso selection <= imageFiles.Count Then
                Return imageFiles(selection - 1)
            Else
                If Not String.IsNullOrEmpty(input) Then
                    MessageBox.Show("Invalid selection.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                End If
                Return String.Empty
            End If
        Catch
            Return String.Empty
        End Try
    End Function

    Private Function CleanFileName(fileName As String) As String
        Try
            Return String.Join("_", fileName.Split(Path.GetInvalidFileNameChars()))
        Catch
            Return "Saved_File.xlsx"
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