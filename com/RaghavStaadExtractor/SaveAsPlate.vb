Imports System.IO
Imports System.Runtime.InteropServices
Imports Microsoft.Office.Interop.Excel

<ComVisible(True)>
<Guid("A1234567-B89C-4DEF-0123-456789ABCDEF")>
Public Class SaveSheetsManager

    ' ====================================
    ' SECTION 1 : Save SHEET1 (Beam Output)
    ' ====================================
    Public Sub Save_Sheet1_BeamOutput()
        Dim excelApp As Application = Nothing
        Dim sourceWorkbook As Workbook = Nothing
        Dim sourceSheet As Worksheet = Nothing
        Dim copiedWorkbook As Workbook = Nothing
        Dim copiedSheet As Worksheet = Nothing
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
            excelApp.EnableEvents = False ' Disable events to prevent crashes

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

            ' Unprotect original temporarily (just in case)
            SafeUnprotectSheet(sourceSheet)

            ' Prepare folder path
            folderPath = Path.Combine(sourceWorkbook.Path, "PlateNovaBeamOutput")
            If Not Directory.Exists(folderPath) Then Directory.CreateDirectory(folderPath)

            ' Get base filename from E3
            Dim baseName As String = GetCellTextSafely(sourceSheet.Range("E3"))
            If String.IsNullOrWhiteSpace(baseName) Then baseName = "Unnamed"

            fileName = baseName & " BEAM OUTPUT.xlsx"
            fullFilePath = Path.Combine(folderPath, fileName)

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
                    Do While File.Exists(Path.Combine(folderPath, baseName & " BEAM OUTPUT (" & i & ").xlsx"))
                        i += 1
                    Loop
                    fileName = baseName & " BEAM OUTPUT (" & i & ").xlsx"
                    fullFilePath = Path.Combine(folderPath, fileName)
                Else
                    GoTo CleanUp
                End If
            End If

            ' Copy Sheet1 into new workbook with error handling
            Try
                sourceSheet.Copy()
                copiedWorkbook = excelApp.Workbooks(excelApp.Workbooks.Count)
                copiedSheet = CType(copiedWorkbook.Sheets(1), Worksheet)
            Catch ex As Exception
                MsgBox("Error copying sheet: " & ex.Message, MsgBoxStyle.Critical)
                GoTo CleanUp
            End Try

            ' === Modify Save As COPY only ===
            Try
                SafeUnprotectSheet(copiedSheet)

                ' Delete extra columns with error handling
                SafeDeleteRange(copiedSheet.Range("F:AJ"))
                SafeDeleteRange(copiedSheet.Range("M:N"))

                ' Delete all shapes safely
                SafeDeleteShapes(copiedSheet)

                ' === Prompt for company name ===
                Dim companyName As String = InputBox("Enter Company Name:", "Company Name", "")
                If Not String.IsNullOrWhiteSpace(companyName) Then
                    UpdateCompanyName(copiedSheet, companyName)
                End If

                'copiedSheet.Protect("2022") 'IF WANT SAVE AS FILE UNPROTECTED THEN IT'S NEED TO BE COMMENTED OUT
            Catch ex As Exception
                MsgBox("Error modifying copied sheet: " & ex.Message, MsgBoxStyle.Critical)
            End Try

            ' Save the copied workbook with proper cleanup
            Try
                copiedWorkbook.SaveAs(fullFilePath, XlFileFormat.xlOpenXMLWorkbook)
                copiedWorkbook.Close(False)
                copiedWorkbook = Nothing
            Catch ex As Exception
                MsgBox("Error saving the workbook: " & ex.Message, MsgBoxStyle.Critical)
            End Try

            MsgBox("Sheet1 saved successfully as '" & fileName & "' in folder:" & vbCrLf & folderPath, MsgBoxStyle.Information)

CleanUp:
            ' Ensure all objects are properly released
            If Not copiedSheet Is Nothing Then Marshal.ReleaseComObject(copiedSheet)
            If Not copiedWorkbook Is Nothing Then
                copiedWorkbook.Close(False)
                Marshal.ReleaseComObject(copiedWorkbook)
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
            ' Ensure COM objects are released
            If Not excelApp Is Nothing Then Marshal.ReleaseComObject(excelApp)
        End Try
    End Sub

    ' ==============================
    ' SECTION 2 : Save SHEET2 (Plate Summary)
    ' ==============================
    Public Sub Save_Sheet2_PlateSummary()
        Dim excelApp As Application = Nothing
        Dim sourceWorkbook As Workbook = Nothing
        Dim sourceSheet As Worksheet = Nothing
        Dim copiedWorkbook As Workbook = Nothing
        Dim copiedSheet As Worksheet = Nothing

        Try
            ' Get Excel instance with more reliable method
            excelApp = GetRunningExcel()
            If excelApp Is Nothing Then
                MsgBox("Could not connect to Excel application.", MsgBoxStyle.Critical)
                Exit Sub
            End If

            ' Critical settings to prevent crashes
            With excelApp
                .ScreenUpdating = False
                .DisplayAlerts = False
                .EnableEvents = False
                .Calculation = XlCalculation.xlCalculationManual
                .AskToUpdateLinks = False
            End With

            ' Get source workbook and sheet with additional checks
            sourceWorkbook = excelApp.ActiveWorkbook
            If sourceWorkbook Is Nothing Then
                MsgBox("No active workbook found.", MsgBoxStyle.Critical)
                GoTo CleanUp
            End If

            sourceSheet = TryCast(sourceWorkbook.Sheets("Sheet2"), Worksheet)
            If sourceSheet Is Nothing Then
                MsgBox("'Sheet2' does not exist in the active workbook.", MsgBoxStyle.Critical)
                GoTo CleanUp
            End If

            ' Create output directory
            Dim folderPath As String = Path.Combine(sourceWorkbook.Path, "PlateNovaPlateSummary")
            Try
                If Not Directory.Exists(folderPath) Then
                    Directory.CreateDirectory(folderPath)
                End If
            Catch ex As Exception
                MsgBox("Error creating output directory: " & ex.Message, MsgBoxStyle.Critical)
                GoTo CleanUp
            End Try

            ' Get C3 and C2 values for file naming
            Dim nameC3 As String = "Unnamed"
            Dim nameC2 As String = ""

            Try
                If Not sourceSheet.Range("C3").Value Is Nothing Then
                    nameC3 = sourceSheet.Range("C3").Value.ToString().Trim()
                    If String.IsNullOrWhiteSpace(nameC3) Then nameC3 = "Unnamed"
                End If
                If Not sourceSheet.Range("C2").Value Is Nothing Then
                    nameC2 = sourceSheet.Range("C2").Value.ToString().Trim()
                End If
            Catch
                ' Use defaults if error
            End Try

            ' Combine C3 + C2 for filename
            Dim fileName As String
            If String.IsNullOrWhiteSpace(nameC2) Then
                fileName = nameC3 & " MATERIAL SUMMARY LIST.xlsx"
            Else
                fileName = nameC3 & " " & nameC2 & " MATERIAL SUMMARY LIST.xlsx"
            End If

            Dim fullFilePath As String = Path.Combine(folderPath, fileName)
            Dim fileIndex As Integer = 1

            ' Handle file naming and overwrite logic
            If File.Exists(fullFilePath) Then
                Dim overwriteChoice As MsgBoxResult = MsgBox(
                "File '" & fileName & "' already exists. Overwrite?",
                MsgBoxStyle.YesNoCancel + MsgBoxStyle.Question)

                Select Case overwriteChoice
                    Case MsgBoxResult.Yes
                        Try
                            File.Delete(fullFilePath)
                        Catch ex As Exception
                            MsgBox("Could not delete existing file: " & ex.Message, MsgBoxStyle.Critical)
                            GoTo CleanUp
                        End Try
                    Case MsgBoxResult.No
                        Do
                            If String.IsNullOrWhiteSpace(nameC2) Then
                                fileName = nameC3 & " MATERIAL SUMMARY LIST (" & fileIndex & ").xlsx"
                            Else
                                fileName = nameC3 & " " & nameC2 & " MATERIAL SUMMARY LIST (" & fileIndex & ").xlsx"
                            End If
                            fullFilePath = Path.Combine(folderPath, fileName)
                            fileIndex += 1
                        Loop While File.Exists(fullFilePath)
                    Case MsgBoxResult.Cancel
                        GoTo CleanUp
                End Select
            End If

            ' Make a copy of the sheet with additional protection
            Try
                sourceSheet.Copy()
                System.Threading.Thread.Sleep(500) ' Small delay to allow copy operation to complete

                ' Get the copied workbook using index to avoid potential race conditions
                copiedWorkbook = excelApp.Workbooks(excelApp.Workbooks.Count)
                copiedSheet = CType(copiedWorkbook.Sheets(1), Worksheet)

                ' Remove shapes more safely
                SafeRemoveShapes(copiedSheet)

                ' Save with retry logic
                Dim saveAttempts As Integer = 0
                Dim maxAttempts As Integer = 3
                Dim savedSuccessfully As Boolean = False

                While saveAttempts < maxAttempts And Not savedSuccessfully
                    Try
                        copiedWorkbook.SaveAs(fullFilePath, XlFileFormat.xlOpenXMLWorkbook)
                        savedSuccessfully = True
                    Catch ex As Exception
                        saveAttempts += 1
                        If saveAttempts >= maxAttempts Then
                            MsgBox("Failed to save after " & maxAttempts & " attempts: " & ex.Message, MsgBoxStyle.Critical)
                            GoTo CleanUp
                        End If
                        System.Threading.Thread.Sleep(1000) ' Wait before retry
                    End Try
                End While

                ' Show success message
                MsgBox("Sheet2 saved successfully as '" & fileName & "' in:" & vbCrLf & folderPath, MsgBoxStyle.Information)

            Catch ex As Exception
                MsgBox("Error during save operation: " & ex.Message, MsgBoxStyle.Critical)
            End Try

CleanUp:
            ' Close and release copied workbook if it exists
            If Not copiedWorkbook Is Nothing Then
                Try
                    copiedWorkbook.Close(False)
                    Marshal.ReleaseComObject(copiedSheet)
                    Marshal.ReleaseComObject(copiedWorkbook)
                Catch
                    ' Ignore cleanup errors
                End Try
            End If

            ' Release source objects
            If Not sourceSheet Is Nothing Then Marshal.ReleaseComObject(sourceSheet)
            If Not sourceWorkbook Is Nothing Then Marshal.ReleaseComObject(sourceWorkbook)

            ' Restore Excel settings
            If Not excelApp Is Nothing Then
                Try
                    With excelApp
                        .ScreenUpdating = True
                        .DisplayAlerts = True
                        .EnableEvents = True
                        .Calculation = XlCalculation.xlCalculationAutomatic
                    End With
                    Marshal.ReleaseComObject(excelApp)
                Catch
                    ' Ignore restoration errors
                End Try
            End If

        Catch ex As Exception
            MsgBox("Unexpected error: " & ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub


    Private Function GetRunningExcel() As Application
        Try
            ' First try getting active instance
            Return Marshal.GetActiveObject("Excel.Application")
        Catch
            Try
                ' If no instance running, create new one
                Return New Application()
            Catch ex As Exception
                MsgBox("Could not start Excel: " & ex.Message, MsgBoxStyle.Critical)
                Return Nothing
            End Try
        End Try
    End Function

    ' ==============================
    ' IMPROVED HELPER FUNCTIONS
    ' ==============================
    Private Function GetExcelInstance() As Application
        Try
            Return CType(Marshal.GetActiveObject("Excel.Application"), Application)
        Catch ex As Exception
            MsgBox("Excel is not running or could not be accessed: " & ex.Message, MsgBoxStyle.Critical)
            Return Nothing
        End Try
    End Function

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

    Private Sub SafeDeleteRange(rng As Range)
        If rng Is Nothing Then Return
        Try
            rng.Delete()
        Catch
            ' Ignore deletion errors
        Finally
            If Not rng Is Nothing Then Marshal.ReleaseComObject(rng)
        End Try
    End Sub

    Private Sub SafeRemoveShapes(ws As Worksheet)
        If ws Is Nothing Then Exit Sub

        ' Remove all shapes (buttons, charts, text boxes, etc.)
        Try
            For i As Integer = ws.Shapes.Count To 1 Step -1
                Dim shp As Shape = Nothing
                Try
                    shp = ws.Shapes.Item(i)
                    shp.Delete()
                Catch
                    ' Ignore individual shape errors
                Finally
                    If shp IsNot Nothing Then Marshal.ReleaseComObject(shp)
                End Try
            Next
        Catch
            ' Ignore shape collection errors
        End Try

        ' Remove all ActiveX controls (e.g., CommandButtons, Checkboxes)
        Try
            For i As Integer = ws.OLEObjects.Count To 1 Step -1
                Dim ole As OLEObject = Nothing
                Try
                    ole = ws.OLEObjects.Item(i)
                    ole.Delete()
                Catch
                    ' Ignore individual OLE errors
                Finally
                    If ole IsNot Nothing Then Marshal.ReleaseComObject(ole)
                End Try
            Next
        Catch
            ' Ignore OLEObjects errors
        End Try
    End Sub

    Private Sub SafeDeleteShapes(ws As Worksheet)
        If ws Is Nothing Then Return
        Try
            For Each shp As Shape In ws.Shapes
                Try
                    shp.Delete()
                Catch
                    ' Ignore shape deletion errors
                Finally
                    Marshal.ReleaseComObject(shp)
                End Try
            Next
        Catch
            ' Ignore enumeration errors
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

    Private Sub UpdateCompanyName(ws As Worksheet, companyName As String)
        If ws Is Nothing Then Return
        Try
            Dim a1 As Range = ws.Range("A1")
            If a1.MergeCells Then
                Dim mergeArea As Range = a1.MergeArea
                Dim mergeAddress As String = mergeArea.Address
                mergeArea.UnMerge()
                ws.Range(mergeAddress).Value = companyName
                ws.Range(mergeAddress).Merge()
                With ws.Range(mergeAddress)
                    .HorizontalAlignment = XlHAlign.xlHAlignCenter
                    .VerticalAlignment = XlVAlign.xlVAlignCenter
                    .Font.Bold = True
                    .Font.Size = 14
                    .Locked = True
                End With
                Marshal.ReleaseComObject(mergeArea)
            Else
                a1.Value = companyName
                a1.HorizontalAlignment = XlHAlign.xlHAlignCenter
                a1.VerticalAlignment = XlVAlign.xlVAlignCenter
                a1.Font.Bold = True
                a1.Font.Size = 14
                a1.Locked = True
            End If
            Marshal.ReleaseComObject(a1)
        Catch ex As Exception
            ' Ignore formatting errors
        End Try
    End Sub
End Class