Imports System.Runtime.InteropServices
Imports Microsoft.Office.Interop.Excel
Imports System.IO
Imports Microsoft.VisualBasic ' For MsgBox

Namespace RaghavTekNova

    <ComVisible(True)>
    <Guid("E8D9F1B5-71A4-4D6E-ACD5-7A2F16E31234")>
    <ProgId("RaghavTekNova.GetAssemblyList")>
    Public Class GetAssemblyList
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
        ' === MAIN ENTRY POINT - ASSEMBLY LIST ===
        Public Sub A1_GET_PR_ASSEMBLY_LIST_VB()
            Dim excelApp As Application = Nothing
            Dim targetWB As Workbook = Nothing
            Dim targetSheet As Worksheet = Nothing
            Dim sourceWB As Workbook = Nothing
            Dim sourceSheet As Worksheet = Nothing

            Dim folderPath As String
            Dim filePath As String
            Dim fileNames As String() = {"PR_Assembly_list_Vba.xls", "PR_Assembly_list_Net_Vba.xls", "PR_Assembly_list_Gross_Vba.xls"}
            Dim tempFiles As New List(Of String)
            Dim selectedFile As String = ""
            Dim userChoice As Object

            Try
                excelApp = CType(Marshal.GetActiveObject("Excel.Application"), Application)
                targetWB = excelApp.ThisWorkbook
                folderPath = targetWB.Path & "\"

                ' Find available files
                For Each fname In fileNames
                    If File.Exists(Path.Combine(folderPath, fname)) Then
                        tempFiles.Add(fname)
                    End If
                Next

                If tempFiles.Count = 0 Then
                    MsgBox("No PR Assembly list file found in the folder.", vbExclamation)
                    Exit Sub
                ElseIf tempFiles.Count = 1 Then
                    selectedFile = tempFiles(0)
                Else
                    ' Ask user to choose
                    Dim displayList As String = "Multiple files found. Select one by entering the number:" & vbCrLf
                    For i As Integer = 0 To tempFiles.Count - 1
                        displayList &= (i + 1).ToString() & ": " & tempFiles(i) & vbCrLf
                    Next

                    userChoice = excelApp.InputBox(displayList, "Select Assembly File", Type:=1)

                    If Not IsNumeric(userChoice) OrElse CInt(userChoice) < 1 OrElse CInt(userChoice) > tempFiles.Count Then
                        MsgBox("Invalid selection. Operation canceled.", vbExclamation)
                        Exit Sub
                    End If

                    selectedFile = tempFiles(CInt(userChoice) - 1)
                End If

                filePath = Path.Combine(folderPath, selectedFile)

                ' Unlock all sheets
                UnlockAllSheets(excelApp, targetWB)

                ' Set target sheet (must be named "Sheet6")
                Try
                    targetSheet = targetWB.Sheets("Sheet6")
                Catch ex As Exception
                    MsgBox("Target sheet 'Sheet6' not found. Please check the codename in the VBA editor.", vbCritical)
                    Exit Sub
                End Try

                targetSheet.Cells.Clear()

                ' Copy from selected file
                excelApp.ScreenUpdating = False
                sourceWB = excelApp.Workbooks.Open(filePath, [ReadOnly]:=True)
                sourceSheet = sourceWB.Sheets(1)

                sourceSheet.UsedRange.Copy(targetSheet.Range("A1"))
                targetSheet.UsedRange.Font.Color = RGB(0, 0, 139)

                sourceWB.Close(SaveChanges:=False)
                excelApp.ScreenUpdating = True

                MsgBox("Assembly Data Flow from """ & selectedFile & """ aligned successfully!")

                ' Lock all sheets
                LockAllSheets(excelApp, targetWB)

            Catch ex As Exception
                MsgBox("💥 SYSTEM ERROR" & vbCrLf & vbCrLf &
                       "An unexpected error has occurred:" & vbCrLf &
                       ex.Message & vbCrLf & vbCrLf &
                       "Please contact support if this issue persists.",
                       vbCritical + vbSystemModal, "RaghavTekNova • Error Handler")
            End Try
        End Sub

        ' === BOLT LIST METHOD ===
        Public Sub B1_Get_PR_Bolt_List_VB()
            Dim excelApp As Application = Nothing
            Dim targetWB As Workbook = Nothing
            Dim targetSheet As Worksheet = Nothing
            Dim sourceWB As Workbook = Nothing
            Dim sourceSheet As Worksheet = Nothing

            Try
                excelApp = CType(Marshal.GetActiveObject("Excel.Application"), Application)
                targetWB = excelApp.ThisWorkbook

                Dim folderPath As String = targetWB.Path & "\"
                Dim fileName As String = "PR_Bolt_list_Vba.xls"
                Dim filePath As String = Path.Combine(folderPath, fileName)

                ' Check if file exists
                If Not File.Exists(filePath) Then
                    MsgBox("File '" & fileName & "' not found in the current folder.", vbExclamation)
                    Exit Sub
                End If

                ' Unlock all sheets
                UnlockAllSheets(excelApp, targetWB)

                ' Set target sheet (Sheet at index 7)
                targetSheet = targetWB.Sheets(7)
                targetSheet.Cells.Clear()

                ' Copy from source file (only columns A to G)
                excelApp.ScreenUpdating = False
                sourceWB = excelApp.Workbooks.Open(filePath, [ReadOnly]:=True)
                sourceSheet = sourceWB.Sheets(1)

                sourceSheet.Range("A:G").Copy(targetSheet.Range("A1"))

                sourceWB.Close(SaveChanges:=False)
                excelApp.ScreenUpdating = True

                MsgBox("Bolt Data Flow, Perfectly Aligned!")

                ' Lock all sheets
                LockAllSheets(excelApp, targetWB)

            Catch ex As Exception
                MsgBox("Error: " & ex.Message, vbCritical)
            End Try
        End Sub

        ' === ASSIGNED BOLT DATA METHOD ===
        Public Sub B2_Copy_Assigned_BoltData_VB()
            Dim excelApp As Application = Nothing
            Dim targetWB As Workbook = Nothing
            Dim sourceWB As Workbook = Nothing
            Dim targetSheet As Worksheet = Nothing
            Dim sourceSheet As Worksheet = Nothing
            Dim activeSheet As Worksheet = Nothing

            Try
                excelApp = CType(Marshal.GetActiveObject("Excel.Application"), Application)
                targetWB = excelApp.ThisWorkbook
                activeSheet = excelApp.ActiveSheet

                Dim folderPath As String = targetWB.Path & "\"
                Dim fileName As String = "Assigned Bolt.xlsm"
                Dim filePath As String = Path.Combine(folderPath, fileName)

                ' Check if file exists
                If Not File.Exists(filePath) Then
                    MsgBox("Source file not found: " & filePath, vbExclamation, "File Not Found")
                    Exit Sub
                End If

                ' Unlock all sheets
                UnlockAllSheets(excelApp, targetWB)

                ' Setup application settings
                excelApp.ScreenUpdating = False
                excelApp.DisplayAlerts = False
                excelApp.Calculation = XlCalculation.xlCalculationManual

                ' Set target sheet
                targetSheet = targetWB.Sheets("Sheet8")

                ' Clear existing data (A, B, C from row 2 onwards)
                targetSheet.Range("A2:C" & targetSheet.Rows.Count).ClearContents()

                ' Open source workbook
                sourceWB = excelApp.Workbooks.Open(filePath, [ReadOnly]:=True)
                sourceSheet = sourceWB.Sheets(1)

                ' Find last row in source data
                Dim inputLastRow As Long = sourceSheet.Cells(sourceSheet.Rows.Count, "A").End(XlDirection.xlUp).Row

                ' Check if there are values to copy
                If inputLastRow < 2 Then
                    MsgBox("Firstly assign values in the Assigned Bolt list Excel.", vbExclamation, "No Values Found")
                    sourceWB.Close(SaveChanges:=False)
                    Exit Sub
                End If

                ' Copy data (A, B, C columns only)
                sourceSheet.Range("A2:C" & inputLastRow).Copy()
                targetSheet.Range("A2").PasteSpecial(XlPasteType.xlPasteValues)

                sourceWB.Close(SaveChanges:=False)

                ' Restore application settings
                excelApp.Calculation = XlCalculation.xlCalculationAutomatic
                excelApp.DisplayAlerts = True
                excelApp.ScreenUpdating = True

                ' Reactivate original sheet
                activeSheet.Activate()

                MsgBox("Assigned Bolt Data Flow, Perfectly Aligned!", vbInformation, "Operation Complete")

                ' Lock all sheets
                LockAllSheets(excelApp, targetWB)

            Catch ex As Exception
                MsgBox("Error: " & ex.Message, vbCritical)
                ' Restore settings in case of error
                If excelApp IsNot Nothing Then
                    excelApp.Calculation = XlCalculation.xlCalculationAutomatic
                    excelApp.DisplayAlerts = True
                    excelApp.ScreenUpdating = True
                End If
            End Try
        End Sub

        ' === SUMMARY METHOD ===
        Public Sub SUMMARY_C1_VB()
            Dim excelApp As Application = Nothing
            Dim targetWB As Workbook = Nothing
            Dim wsDest As Worksheet = Nothing
            Dim wsSource As Worksheet = Nothing

            Try
                excelApp = CType(Marshal.GetActiveObject("Excel.Application"), Application)
                targetWB = excelApp.ThisWorkbook

                ' Unlock all sheets
                UnlockAllSheets(excelApp, targetWB)

                ' Set destination and source sheets
                wsDest = targetWB.Sheets("Sheet3")
                wsSource = targetWB.Sheets("Sheet1")

                ' Clear existing data and formatting
                wsDest.Cells.Clear()

                ' Initialize variables
                Dim rowNum As Long = 8
                Dim serialNum As Long = 1
                Dim sumC As Double = 0
                Dim sumD As Double = 0

                ' Setup headers and formatting
                With wsDest
                    .Range("A2").Value = "PROJECT NO:-"
                    .Range("A3").Value = "PROJECT NAME:-"
                    .Range("A4").Value = "CLIENT NAME:-"
                    .Range("A5").Value = "TOTAL:-"
                    .Range("A1:D1").Merge()
                    .Range("A1:D1").Value = "BOQ SUMMARY"
                    .Range("A1:D1").HorizontalAlignment = XlHAlign.xlHAlignCenter
                    .Range("A1:D1").Font.Size = 14
                    .Range("A1:H5").Font.Bold = True
                    .Range("A1:H5").Font.Size = 14
                    .Range("A7").Value = "SR.NO."
                    .Range("B7").Value = "CATEGORY"
                    .Range("C7").Value = "WEIGHT(KG)"
                    .Range("D7").Value = "AREA(SQM)"
                End With

                ' Format header row
                With wsDest.Range("A7:D7")
                    .Font.Bold = True
                    .HorizontalAlignment = XlHAlign.xlHAlignCenter
                    .VerticalAlignment = XlVAlign.xlVAlignCenter
                    .Interior.Color = RGB(173, 216, 230)
                    .Font.Size = 12
                End With

                ' Set background colors
                wsDest.Range("A1:D1").Interior.Color = RGB(248, 203, 173)
                wsDest.Range("A6").Interior.Color = RGB(248, 203, 173)

                ' Merge cells
                wsDest.Range("A2:B2").Merge()
                wsDest.Range("A3:B3").Merge()
                wsDest.Range("A4:B4").Merge()
                wsDest.Range("A5:B5").Merge()
                wsDest.Range("C2:D2").Merge()
                wsDest.Range("C3:D3").Merge()
                wsDest.Range("C4:D4").Merge()
                wsDest.Range("A6:D6").Merge()

                ' Set column widths
                wsDest.Columns("A").ColumnWidth = 8
                wsDest.Columns("B").ColumnWidth = 25
                wsDest.Columns("C").ColumnWidth = 25
                wsDest.Columns("D").ColumnWidth = 25

                ' Populate project information
                wsDest.Range("C2").Value = wsSource.Range("D2").Value
                wsDest.Range("C3").Value = wsSource.Range("D3").Value
                wsDest.Range("C4").Value = wsSource.Range("D4").Value

                ' Loop through worksheets from index 12 onwards
                For sheetIndex As Integer = 12 To targetWB.Sheets.Count
                    Dim ws As Worksheet = targetWB.Sheets(sheetIndex)

                    wsDest.Cells(rowNum, 1).Value = serialNum
                    wsDest.Cells(rowNum, 2).Value = ws.Range("J3").Value
                    wsDest.Cells(rowNum, 3).Value = ws.Range("J5").Value

                    Dim areaValue As Object = ws.Range("K5").Value
                    wsDest.Cells(rowNum, 4).Value = If(areaValue Is Nothing OrElse IsDBNull(areaValue) OrElse areaValue.ToString() = "", 0, areaValue)

                    ' Add to sums
                    sumC += CDbl(wsDest.Cells(rowNum, 3).Value)
                    sumD += CDbl(wsDest.Cells(rowNum, 4).Value)

                    rowNum += 1
                    serialNum += 1
                Next

                ' Add Nut bolt entry from Sheet2
                wsDest.Cells(rowNum, 1).Value = serialNum
                wsDest.Cells(rowNum, 2).Value = "Nut bolt"
                wsDest.Cells(rowNum, 3).Value = targetWB.Sheets("Sheet2").Range("H5").Value
                wsDest.Cells(rowNum, 4).Value = 0

                ' Add to sums
                sumC += CDbl(wsDest.Cells(rowNum, 3).Value)
                sumD += CDbl(wsDest.Cells(rowNum, 4).Value)

                ' Display totals
                wsDest.Range("C5").Value = sumC.ToString() & " (KG)"
                wsDest.Range("D5").Value = sumD.ToString() & " (SQM)"
                wsDest.Range("C5:D5").HorizontalAlignment = XlHAlign.xlHAlignLeft

                ' Apply formatting and borders
                With wsDest.Range("A8:D" & rowNum)
                    .HorizontalAlignment = XlHAlign.xlHAlignLeft
                    .VerticalAlignment = XlVAlign.xlVAlignCenter
                    .Borders.LineStyle = XlLineStyle.xlContinuous
                End With

                ' Apply borders to entire range
                With wsDest.Range("A1:D" & rowNum)
                    .Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                    .Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    .Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                    .Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    .Borders(XlBordersIndex.xlInsideVertical).LineStyle = XlLineStyle.xlContinuous
                    .Borders(XlBordersIndex.xlInsideHorizontal).LineStyle = XlLineStyle.xlContinuous
                End With

                MsgBox("Summary generated successfully!")

                ' Lock all sheets
                LockAllSheets(excelApp, targetWB)

            Catch ex As Exception
                MsgBox("Error " & ex.GetHashCode() & ": " & ex.Message, vbCritical, "Runtime Error")
            End Try
        End Sub

        ' === MATERIAL PART LIST METHOD ===
        Public Sub P1_GET_MATERIAL_PART_LIST_PR_TRIAL_VB()
            Dim excelApp As Application = Nothing
            Dim targetWB As Workbook = Nothing
            Dim targetSheet As Worksheet = Nothing
            Dim sourceWB As Workbook = Nothing
            Dim sourceSheet As Worksheet = Nothing

            Try
                excelApp = CType(Marshal.GetActiveObject("Excel.Application"), Application)
                targetWB = excelApp.ThisWorkbook

                Dim folderPath As String = targetWB.Path & "\"
                Dim fileNames As String() = {"PR_Material_Part_list_vba.xls", "PR_Material_Part_list_Net_vba.xls", "PR_Material_Part_list_Gross_vba.xls"}
                Dim tempFiles As New List(Of String)
                Dim selectedFile As String = ""
                Dim userChoice As Object

                ' Find available files
                For Each fname In fileNames
                    If File.Exists(Path.Combine(folderPath, fname)) Then
                        tempFiles.Add(fname)
                    End If
                Next

                If tempFiles.Count = 0 Then
                    MsgBox("No PR Material Part list file found in the folder.", vbExclamation)
                    Exit Sub
                ElseIf tempFiles.Count = 1 Then
                    selectedFile = tempFiles(0)
                Else
                    ' Ask user to choose
                    Dim displayList As String = "Multiple files found. Select one by entering the number:" & vbCrLf
                    For i As Integer = 0 To tempFiles.Count - 1
                        displayList &= (i + 1).ToString() & ": " & tempFiles(i) & vbCrLf
                    Next

                    userChoice = excelApp.InputBox(displayList, "Select Material Part List File", Type:=1)

                    If Not IsNumeric(userChoice) OrElse CInt(userChoice) < 1 OrElse CInt(userChoice) > tempFiles.Count Then
                        MsgBox("Invalid selection. Operation canceled.", vbExclamation)
                        Exit Sub
                    End If

                    selectedFile = tempFiles(CInt(userChoice) - 1)
                End If

                Dim filePath As String = Path.Combine(folderPath, selectedFile)

                ' Unlock all sheets
                UnlockAllSheets(excelApp, targetWB)

                ' Set target sheet
                Try
                    targetSheet = targetWB.Sheets("Sheet9")
                Catch ex As Exception
                    MsgBox("Target sheet 'Sheet9' not found. Please ensure the worksheet is named 'Sheet9'.", vbCritical)
                    Exit Sub
                End Try

                targetSheet.Cells.Clear()

                ' Copy from source file
                excelApp.ScreenUpdating = False
                sourceWB = excelApp.Workbooks.Open(filePath, [ReadOnly]:=True)
                sourceSheet = sourceWB.Sheets(1)

                sourceSheet.UsedRange.Copy(targetSheet.Range("A1"))
                targetSheet.UsedRange.Font.Color = RGB(0, 0, 139)

                sourceWB.Close(SaveChanges:=False)
                excelApp.ScreenUpdating = True

                MsgBox("Material list data from """ & selectedFile & """ aligned successfully!")

                ' Lock all sheets
                LockAllSheets(excelApp, targetWB)

            Catch ex As Exception
                MsgBox("Error: " & ex.Message, vbCritical)
            End Try
        End Sub

        ' === DRAWING TRACKING LIST METHOD ===
        Public Sub V1_GET_PR_Drawing_Tracking_list_VB()
            Dim excelApp As Application = Nothing
            Dim targetWB As Workbook = Nothing
            Dim targetSheet As Worksheet = Nothing
            Dim sourceWB As Workbook = Nothing
            Dim sourceSheet As Worksheet = Nothing

            Try
                excelApp = CType(Marshal.GetActiveObject("Excel.Application"), Application)
                targetWB = excelApp.ThisWorkbook

                Dim folderPath As String = targetWB.Path & "\"
                Dim fileNames As String() = {"PR_Drawing_Tracking_list_Vba.xls", "PR_Drawing_Tracking_list_Net_Vba.xls", "PR_Drawing_Tracking_list_Gross_Vba.xls"}
                Dim tempFiles As New List(Of String)
                Dim selectedFile As String = ""
                Dim userChoice As Object

                ' Find available files
                For Each fname In fileNames
                    If File.Exists(Path.Combine(folderPath, fname)) Then
                        tempFiles.Add(fname)
                    End If
                Next

                If tempFiles.Count = 0 Then
                    MsgBox("No Drawing Tracking list file found in the folder.", vbExclamation)
                    Exit Sub
                ElseIf tempFiles.Count = 1 Then
                    selectedFile = tempFiles(0)
                Else
                    ' Ask user to choose
                    Dim displayList As String = "Multiple files found. Select one by entering the number:" & vbCrLf
                    For i As Integer = 0 To tempFiles.Count - 1
                        displayList &= (i + 1).ToString() & ": " & tempFiles(i) & vbCrLf
                    Next

                    userChoice = excelApp.InputBox(displayList, "Select Drawing Tracking File", Type:=1)

                    If Not IsNumeric(userChoice) OrElse CInt(userChoice) < 1 OrElse CInt(userChoice) > tempFiles.Count Then
                        MsgBox("Invalid selection. Operation canceled.", vbExclamation)
                        Exit Sub
                    End If

                    selectedFile = tempFiles(CInt(userChoice) - 1)
                End If

                Dim filePath As String = Path.Combine(folderPath, selectedFile)

                ' Unlock all sheets
                UnlockAllSheets(excelApp, targetWB)

                ' Set target sheet (assuming Sheet11 exists - you may need to adjust this)
                Try
                    targetSheet = targetWB.Sheets(11) ' Using index 11 as per original VBA
                Catch ex As Exception
                    MsgBox("Target sheet at index 11 not found.", vbCritical)
                    Exit Sub
                End Try

                targetSheet.Cells.Clear()

                ' Copy from source file
                excelApp.ScreenUpdating = False
                sourceWB = excelApp.Workbooks.Open(filePath, [ReadOnly]:=True)
                sourceSheet = sourceWB.Sheets(1)

                sourceSheet.UsedRange.Copy(targetSheet.Range("A1"))
                targetSheet.UsedRange.Font.Color = RGB(0, 0, 139)

                sourceWB.Close(SaveChanges:=False)
                excelApp.ScreenUpdating = True

                MsgBox("Drawing data from """ & selectedFile & """ aligned successfully!")

                ' Lock all sheets
                LockAllSheets(excelApp, targetWB)

            Catch ex As Exception
                MsgBox("Error: " & ex.Message, vbCritical)
            End Try
        End Sub

        ' === Helper: Unlock all sheets ===
        Private Sub UnlockAllSheets(app As Application, wb As Workbook)
            Dim currentSheet As Worksheet = app.ActiveSheet
            Dim calcMode As XlCalculation = app.Calculation

            app.ScreenUpdating = False
            app.EnableEvents = False
            app.Calculation = XlCalculation.xlCalculationManual

            For Each ws As Worksheet In wb.Worksheets
                ws.Unprotect("2022")
            Next

            currentSheet.Activate()
            app.Calculation = calcMode
            app.EnableEvents = True
            app.ScreenUpdating = True
        End Sub

        ' === Helper: Lock all sheets ===
        Private Sub LockAllSheets(app As Application, wb As Workbook)
            Dim currentSheet As Worksheet = app.ActiveSheet
            Dim calcMode As XlCalculation = app.Calculation

            app.ScreenUpdating = False
            app.EnableEvents = False
            app.Calculation = XlCalculation.xlCalculationManual

            For Each ws As Worksheet In wb.Worksheets
                ws.Protect("2022", UserInterfaceOnly:=True)
            Next

            currentSheet.Activate()
            app.Calculation = calcMode
            app.EnableEvents = True
            app.ScreenUpdating = True
        End Sub

    End Class

End Namespace