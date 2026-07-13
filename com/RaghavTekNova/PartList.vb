Imports Microsoft.Office.Interop.Excel
Imports System.Runtime.InteropServices
Imports System.Text.RegularExpressions
Imports System.Collections.Generic

Namespace RaghavTekNova
    <ComVisible(True)>
    <Guid("E2C8A3D1-ABCD-4E6B-9A34-7D12AB123456")> ' Replace with a unique GUID
    <ProgId("RaghavTekNova.PartList")>
    Public Class PartList

        ' === Helper: Unlock all sheets ===
        Private Sub UnlockAllSheets(app As Application, wb As Workbook)
            Dim currentSheet As Worksheet = app.ActiveSheet
            Dim calcMode As XlCalculation = app.Calculation
            app.ScreenUpdating = False
            app.EnableEvents = False
            app.Calculation = XlCalculation.xlCalculationManual
            For Each ws As Worksheet In wb.Worksheets
                Try
                    ws.Unprotect("2022")
                Catch
                    ' Continue if sheet is not protected or password is different
                End Try
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
                Try
                    ws.Protect("2022", UserInterfaceOnly:=True)
                Catch
                    ' Continue if there's an issue protecting
                End Try
            Next
            currentSheet.Activate()
            app.Calculation = calcMode
            app.EnableEvents = True
            app.ScreenUpdating = True
        End Sub

        ' Helper function to safely convert to Double
        Private Function SafeToDouble(value As Object) As Double
            If value Is Nothing OrElse IsDBNull(value) Then
                Return 0
            End If

            Dim stringValue As String = value.ToString().Trim()
            If String.IsNullOrEmpty(stringValue) OrElse stringValue = "-" Then
                Return 0
            End If

            Dim result As Double
            If Double.TryParse(stringValue, result) Then
                Return result
            End If

            Return 0
        End Function

        ' Helper function to safely convert to String
        Private Function SafeToString(value As Object) As String
            If value Is Nothing OrElse IsDBNull(value) Then
                Return ""
            End If
            Return value.ToString().Trim()
        End Function

        ' === P2_RUN_MATERIAL_PART_LIST_OPERATIONS ===
        Public Sub P2_RUN_MATERIAL_PART_LIST_OPERATIONS()
            Dim xlApp As Application = Nothing
            Dim wb As Workbook = Nothing
            Dim wsTarget As Worksheet = Nothing
            Dim wsSource As Worksheet = Nothing
            Dim dataRange As Range = Nothing

            Try
                xlApp = CType(Marshal.GetActiveObject("Excel.Application"), Application)
                wb = xlApp.ActiveWorkbook

                ' Unlock all sheets at the beginning
                UnlockAllSheets(xlApp, wb)

                ' Turn off screen updating for the entire operation
                Dim originalScreenUpdating As Boolean = xlApp.ScreenUpdating
                Dim originalEnableEvents As Boolean = xlApp.EnableEvents
                Dim originalCalculation As XlCalculation = xlApp.Calculation
                xlApp.ScreenUpdating = False
                xlApp.EnableEvents = False
                xlApp.Calculation = XlCalculation.xlCalculationManual

                wsTarget = CType(wb.Sheets("Sheet4"), Worksheet)
                wsSource = CType(wb.Sheets("Sheet9"), Worksheet)

                ' Clear the entire content of the target sheet (Sheet4)
                wsTarget.Cells.Clear()

                Dim lastRow As Long = CLng(wsSource.Cells(wsSource.Rows.Count, "B").End(XlDirection.xlUp).Row)
                Dim targetStartRow As Long = 8

                ' Ensure there is sufficient data to copy
                If lastRow < 8 Then
                    MsgBox("Not enough data in the source sheet!")
                    GoTo Cleanup
                End If

                ' Copy project information from the source sheet (Sheet9) to target sheet (Sheet4)
                wsTarget.Range("D2").Value = wsSource.Range("E1").Value ' Project No.
                wsTarget.Range("D3").Value = wsSource.Range("E2").Value ' Project Name
                wsTarget.Range("D4").Value = wsSource.Range("E3").Value ' Client Name
                wsTarget.Range("D5").Value = wsSource.Range("E4").Value ' Date

                ' Add headers in the target sheet
                With wsTarget
                    .Range("C2").Value = "PROJECT NO:-"
                    .Range("C3").Value = "PROJECT NAME:-"
                    .Range("C4").Value = "CLIENT NAME:-"
                    .Range("C5").Value = "DATE:-"
                    .Range("G5").Value = "REVISION:-"
                    .Range("I2").Value = "TOTAL"
                    .Range("I3").Value = "PARTS:-"
                    .Range("I4").Value = "WEIGHT(KG):-"
                    .Range("I5").Value = "AREA(SQM):-"
                End With

                ' Format the project information text
                With wsTarget.Range("A2:J5")
                    .Font.Bold = True
                    .Font.Size = 13
                End With

                ' Align content in C2:J5
                With wsTarget.Range("C2:J5")
                    .Font.Bold = True
                    .Font.Size = 14
                End With

                ' Set light blue background color
                wsTarget.Range("A1:B1,C1:G1,H1:I1,J1").Interior.Color = RGB(248, 203, 173)
                wsTarget.Range("A7:J7").Interior.Color = RGB(173, 216, 230)
                wsTarget.Range("A6").Interior.Color = RGB(248, 203, 173)

                ' Merge and format cells for project headers and details
                wsTarget.Range("A2:B5").Merge()
                wsTarget.Range("D5:E5").Merge()
                wsTarget.Range("D2:H2").Merge()
                wsTarget.Range("D3:H3").Merge()
                wsTarget.Range("D4:H4").Merge()
                wsTarget.Range("I2:J2").Merge()

                With wsTarget.Range("A1:J1")
                    .Merge()
                    .Value = "BILL OF MATERIAL (PART LIST)"
                    .HorizontalAlignment = XlHAlign.xlHAlignCenter
                    .VerticalAlignment = XlVAlign.xlVAlignCenter
                    .Font.Bold = True
                    .Font.Size = 15
                End With

                With wsTarget.Range("I2:I5,J2:J5,K5")
                    .HorizontalAlignment = XlHAlign.xlHAlignLeft
                    .VerticalAlignment = XlVAlign.xlVAlignCenter
                    .Font.Bold = True
                    .Font.Size = 14
                End With

                ' Merge and leave row 6 blank
                With wsTarget.Range("A6:J6")
                    .Merge()
                    .Value = ""
                End With

                ' Write headers in row 7
                wsTarget.Range("A7").Value = "SR.NO."
                wsTarget.Range("B7").Value = "PART MARK"
                wsTarget.Range("C7").Value = "PROFILE"
                wsTarget.Range("D7").Value = "QTY."
                wsTarget.Range("E7").Value = "LENGTH"
                wsTarget.Range("F7").Value = "MATERIAL"
                wsTarget.Range("G7").Value = "UNIT WEIGHT"
                wsTarget.Range("H7").Value = "TOTAL WEIGHT"
                wsTarget.Range("I7").Value = "UNIT AREA"
                wsTarget.Range("J7").Value = "TOTAL AREA"

                ' Format header row
                With wsTarget.Range("A7:J7")
                    .Font.Bold = True
                    .HorizontalAlignment = XlHAlign.xlHAlignCenter
                    .VerticalAlignment = XlVAlign.xlVAlignCenter
                    .Font.Size = 12
                End With

                ' Adjust column widths
                wsTarget.Columns("A").ColumnWidth = 7
                wsTarget.Columns("B").ColumnWidth = 15
                wsTarget.Columns("C").ColumnWidth = 25
                wsTarget.Columns("D").ColumnWidth = 8
                wsTarget.Columns("E").ColumnWidth = 10
                wsTarget.Columns("F").ColumnWidth = 12
                wsTarget.Columns("G").ColumnWidth = 15
                wsTarget.Columns("H").ColumnWidth = 16
                wsTarget.Columns("I").ColumnWidth = 15
                wsTarget.Columns("J").ColumnWidth = 15

                ' Copy data from the source sheet
                dataRange = wsSource.Range("B8:M" & (lastRow - 1))
                dataRange.Copy(wsTarget.Cells(targetStartRow, 2))

                ' Read data into array for processing
                Dim dataArray As Object(,) = CType(wsTarget.Range(wsTarget.Cells(targetStartRow, 2), wsTarget.Cells(lastRow - 1, 11)).Value, Object(,))
                Dim rowCount As Integer = dataArray.GetUpperBound(0)
                Dim filteredData As New List(Of Object())()
                Dim totalWeight As Double = 0
                Dim totalQty As Double = 0
                Dim totalArea As Double = 0

                For rowIndex As Integer = 1 To rowCount
                    Dim totalWeightCell As Double = SafeToDouble(dataArray(rowIndex, 7)) ' H=7
                    If totalWeightCell <> 0 Then
                        Dim rowData(8) As Object ' B to J, 9 columns, index 0-8
                        For col As Integer = 1 To 9
                            rowData(col - 1) = dataArray(rowIndex, col)
                        Next
                        filteredData.Add(rowData)
                        totalWeight += totalWeightCell
                        totalQty += SafeToDouble(dataArray(rowIndex, 3)) ' D=3
                        totalArea += SafeToDouble(dataArray(rowIndex, 9)) ' J=9
                    End If
                Next

                ' Clear the data area
                wsTarget.Range(wsTarget.Cells(targetStartRow, 1), wsTarget.Cells(lastRow - 1, 10)).Clear()

                ' Write filtered data back with serial numbers
                Dim serialNum As Long = 1
                For Each rowData In filteredData
                    wsTarget.Cells(targetStartRow + serialNum - 1, 1).Value = serialNum
                    Dim outputRange As Range = wsTarget.Range(wsTarget.Cells(targetStartRow + serialNum - 1, 2), wsTarget.Cells(targetStartRow + serialNum - 1, 10))
                    outputRange.Value = rowData
                    serialNum += 1
                Next

                wsTarget.Range("J4").Value = totalWeight
                wsTarget.Range("J3").Value = totalQty
                wsTarget.Range("J5").Value = totalArea

                ' Apply borders to all data from A1 to J till last filled row
                Dim lastFilledRow As Long = targetStartRow + filteredData.Count - 1
                With wsTarget.Range("A1:J" & lastFilledRow).Borders
                    .LineStyle = XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = XlBorderWeight.xlThin
                End With

                ' Align content and set text color
                With wsTarget.Range("A8:J" & lastFilledRow)
                    .HorizontalAlignment = XlHAlign.xlHAlignLeft
                    .VerticalAlignment = XlVAlign.xlVAlignCenter
                    .Font.Color = RGB(0, 0, 139)
                End With

                ' Adjust row heights
                wsTarget.Rows("6:" & lastFilledRow).RowHeight = 18

                ' Merge the first empty row
                Dim firstEmptyRow As Long = lastFilledRow + 1
                wsTarget.Range(wsTarget.Cells(firstEmptyRow, 1), wsTarget.Cells(firstEmptyRow, 10)).Merge()
                wsTarget.Cells(firstEmptyRow, 1).Value = ""

                ' Freeze panes
                wsTarget.Activate()
                wsTarget.Cells(8, 1).Select()
                xlApp.ActiveWindow.FreezePanes = True

                MsgBox("Part List Created successfully.")

Cleanup:
                ' Restore settings
                xlApp.Calculation = originalCalculation
                xlApp.EnableEvents = originalEnableEvents
                xlApp.ScreenUpdating = originalScreenUpdating

                ' Lock all sheets back
                LockAllSheets(xlApp, wb)

            Catch ex As Exception
                MsgBox("An error occurred: " & ex.Message)
                ' Ensure sheets are locked even if error occurs
                Try
                    If xlApp IsNot Nothing And wb IsNot Nothing Then
                        LockAllSheets(xlApp, wb)
                    End If
                Catch
                    ' Ignore protection errors in catch block
                End Try
            Finally
                If dataRange IsNot Nothing Then Marshal.ReleaseComObject(dataRange)
                If wsTarget IsNot Nothing Then Marshal.ReleaseComObject(wsTarget)
                If wsSource IsNot Nothing Then Marshal.ReleaseComObject(wsSource)
                If wb IsNot Nothing Then Marshal.ReleaseComObject(wb)
                If xlApp IsNot Nothing Then Marshal.ReleaseComObject(xlApp)
            End Try
        End Sub

        ' === P3_MATERIAL_LIST_SUMMARY ===
        Public Sub P3_MATERIAL_LIST_SUMMARY()
            Dim xlApp As Application = Nothing
            Dim wb As Workbook = Nothing
            Dim wsSource As Worksheet = Nothing
            Dim wsDest As Worksheet = Nothing
            Dim wsTarget As Worksheet = Nothing

            Try
                xlApp = CType(Marshal.GetActiveObject("Excel.Application"), Application)
                wb = xlApp.ActiveWorkbook

                ' Unlock all sheets at the beginning
                UnlockAllSheets(xlApp, wb)

                ' Turn off screen updating for the entire operation
                Dim originalScreenUpdating As Boolean = xlApp.ScreenUpdating
                Dim originalEnableEvents As Boolean = xlApp.EnableEvents
                Dim originalCalculation As XlCalculation = xlApp.Calculation
                xlApp.ScreenUpdating = False
                xlApp.EnableEvents = False
                xlApp.Calculation = XlCalculation.xlCalculationManual

                wsSource = CType(wb.Sheets("Sheet4"), Worksheet)
                wsDest = CType(wb.Sheets("Sheet5"), Worksheet)
                wsTarget = CType(wb.Sheets("Sheet9"), Worksheet)

                ' Clear all existing data in Sheet5
                wsDest.Cells.Clear()

                Dim lastRow As Long = CLng(wsSource.Cells(wsSource.Rows.Count, "C").End(XlDirection.xlUp).Row)
                Dim outputRow As Long = 8

                ' Copy project information
                wsDest.Range("C2").Value = wsTarget.Range("E1").Value
                wsDest.Range("C3").Value = wsTarget.Range("E2").Value
                wsDest.Range("C4").Value = wsTarget.Range("E3").Value
                wsDest.Range("C5").Value = wsTarget.Range("E4").Value
                wsDest.Range("F5").Value = wsSource.Range("J4").Value

                ' Add headers
                With wsDest
                    .Range("A2").Value = "PROJECT NO:-"
                    .Range("A3").Value = "PROJECT NAME:-"
                    .Range("A4").Value = "CLIENT NAME:-"
                    .Range("A5").Value = "DATE:-"
                    .Range("E5").Value = "TOTAL:-"
                End With

                ' Format project information
                With wsDest.Range("A2:F5")
                    .Font.Bold = True
                    .Font.Size = 13.5
                End With

                ' Set background colors
                wsDest.Range("A1:B1,C1:F1").Interior.Color = RGB(248, 203, 173)
                wsDest.Range("A7:F7").Interior.Color = RGB(173, 216, 230)
                wsDest.Range("A6").Interior.Color = RGB(248, 203, 173)
                wsDest.Range("A1:F1").Merge()
                wsDest.Range("A1").Value = "MATERIAL SUMMARY"
                wsDest.Range("A1").HorizontalAlignment = XlHAlign.xlHAlignCenter
                wsDest.Range("A1").VerticalAlignment = XlVAlign.xlVAlignCenter
                wsDest.Range("A1").Font.Bold = True
                wsDest.Range("A1").Font.Size = 15

                ' Merge cells
                wsDest.Range("A2:B2").Merge()
                wsDest.Range("A3:B3").Merge()
                wsDest.Range("A4:B4").Merge()
                wsDest.Range("A5:B5").Merge()
                wsDest.Range("C2:F2").Merge()
                wsDest.Range("C3:F3").Merge()
                wsDest.Range("C4:F4").Merge()
                wsDest.Range("A6:F6").Merge()

                ' Write headers
                wsDest.Range("A7").Value = "SR.NO."
                wsDest.Range("B7").Value = "PROFILE"
                wsDest.Range("C7").Value = "LENGTH (MM)"
                wsDest.Range("D7").Value = "MATERIAL"
                wsDest.Range("E7").Value = "TOTAL WEIGHT"
                wsDest.Range("F7").Value = "TOTAL AREA"

                ' Format headers
                With wsDest.Range("A7:F7")
                    .Font.Bold = True
                    .HorizontalAlignment = XlHAlign.xlHAlignCenter
                    .VerticalAlignment = XlVAlign.xlVAlignCenter
                    .Font.Size = 12
                End With

                ' Adjust column widths
                wsDest.Columns("A").ColumnWidth = 8
                wsDest.Columns("B").ColumnWidth = 22
                wsDest.Columns("C").ColumnWidth = 15
                wsDest.Columns("D").ColumnWidth = 12
                wsDest.Columns("E").ColumnWidth = 16
                wsDest.Columns("F").ColumnWidth = 16

                ' Read data into array
                Dim dataArray As Object(,) = CType(wsSource.Range("B8:J" & lastRow).Value, Object(,)) ' B to J, columns 1 to 9
                Dim rowCount As Integer = dataArray.GetUpperBound(0)

                ' Use dictionary to accumulate sums
                Dim summaryDict As New Dictionary(Of String, Double())
                For rowIndex As Integer = 1 To rowCount
                    Dim profileName As String = SafeToString(dataArray(rowIndex, 2)) ' C=2 (B=1, C=2)
                    Dim materialName As String = SafeToString(dataArray(rowIndex, 5)) ' F=5
                    If profileName = "" Or materialName = "" Then Continue For

                    Dim key As String = profileName & "|" & materialName
                    Dim sums(2) As Double ' 0: length, 1: weight, 2: area

                    If summaryDict.ContainsKey(key) Then
                        sums = summaryDict(key)
                    End If

                    sums(0) += SafeToDouble(dataArray(rowIndex, 3)) * SafeToDouble(dataArray(rowIndex, 4)) ' D=3 * E=4
                    sums(1) += SafeToDouble(dataArray(rowIndex, 7)) ' H=7
                    sums(2) += SafeToDouble(dataArray(rowIndex, 9)) ' J=9

                    summaryDict(key) = sums
                Next

                ' Prepare output data
                Dim outputList As New List(Of Object())
                For Each kvp In summaryDict
                    Dim parts() As String = kvp.Key.Split("|"c)
                    Dim rowData(4) As Object ' B to F
                    rowData(0) = parts(0) ' B
                    rowData(1) = kvp.Value(0) ' C
                    rowData(2) = parts(1) ' D
                    rowData(3) = kvp.Value(1) ' E
                    rowData(4) = kvp.Value(2) ' F
                    outputList.Add(rowData)
                Next

                ' Sort the list by profile (B)
                outputList.Sort(Function(x, y) String.Compare(CStr(x(0)), CStr(y(0))))

                ' Write to sheet
                Dim serialNumber As Long = 1
                For Each rowData In outputList
                    wsDest.Cells(outputRow, "A").Value = serialNumber
                    Dim outputRange As Range = wsDest.Range(wsDest.Cells(outputRow, 2), wsDest.Cells(outputRow, 6))
                    outputRange.Value = rowData
                    outputRow += 1
                    serialNumber += 1
                Next

                ' Remove extra columns
                wsDest.Range("G:Z").EntireColumn.Delete()

                ' Apply borders
                With wsDest.UsedRange.Borders
                    .LineStyle = XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = XlBorderWeight.xlThin
                End With

                ' Set alignment and text color
                With wsDest.Range("A8:F" & (outputRow - 1))
                    .HorizontalAlignment = XlHAlign.xlHAlignLeft
                    .VerticalAlignment = XlVAlign.xlVAlignCenter
                    .Font.Color = RGB(0, 0, 139)
                End With

                MsgBox("Data summarized successfully!", MsgBoxStyle.Information)

                ' Restore settings
                xlApp.Calculation = originalCalculation
                xlApp.EnableEvents = originalEnableEvents
                xlApp.ScreenUpdating = originalScreenUpdating

                ' Lock all sheets back
                LockAllSheets(xlApp, wb)

            Catch ex As Exception
                MsgBox("An error occurred: " & ex.Message)
                ' Ensure sheets are locked even if error occurs
                Try
                    If xlApp IsNot Nothing And wb IsNot Nothing Then
                        LockAllSheets(xlApp, wb)
                    End If
                Catch
                    ' Ignore protection errors in catch block
                End Try
            Finally
                If wsSource IsNot Nothing Then Marshal.ReleaseComObject(wsSource)
                If wsDest IsNot Nothing Then Marshal.ReleaseComObject(wsDest)
                If wsTarget IsNot Nothing Then Marshal.ReleaseComObject(wsTarget)
                If wb IsNot Nothing Then Marshal.ReleaseComObject(wb)
                If xlApp IsNot Nothing Then Marshal.ReleaseComObject(xlApp)
            End Try
        End Sub

        ' === P4_SUMMERIZED_DATA === EXACT VBA LOGIC
        Public Sub P4_SUMMERIZED_DATA()
            Dim xlApp As Application = Nothing
            Dim wb As Workbook = Nothing
            Dim ws As Worksheet = Nothing

            Try
                xlApp = CType(Marshal.GetActiveObject("Excel.Application"), Application)
                wb = xlApp.ActiveWorkbook

                ' Unlock all sheets at the beginning
                UnlockAllSheets(xlApp, wb)

                ws = CType(wb.Sheets("Sheet5"), Worksheet)

                Dim lastRow As Long = CLng(ws.Cells(ws.Rows.Count, "B").End(XlDirection.xlUp).Row)
                Dim cell As Range
                Dim profile As String
                Dim num1 As Long, num2 As Long, minNum As Long
                Dim matches As Object
                Dim i As Integer, j As Long
                Dim serialNumber As Long
                Dim totalLength As Double, totalWeight As Double, totalArea As Double
                Dim material As String

                ' Profile keywords
                Dim profileKeywords() As String = {"PL", "PLT", "PLATE", "BL", "BPL", "FB", "FL",
                                                  "FLAT", "FLT", "FBL", "FPL", "L", "CLL", "BLL", "CHEQ_PLT"}

                ' --- First operation: Transform profile column ---
                For Each cell In ws.Range("B8:B" & lastRow)
                    profile = SafeToString(cell.Value)

                    ' Check if profile starts with any keyword
                    Dim isTargetProfile As Boolean = False
                    For i = 0 To profileKeywords.Length - 1
                        If profile.StartsWith(profileKeywords(i)) Then
                            isTargetProfile = True
                            Exit For
                        End If
                    Next

                    If isTargetProfile Then
                        If profile.StartsWith("CHEQ_PLT") Then
                            Dim re As Object = CreateObject("VBScript.RegExp")
                            re.Pattern = "\d+"
                            re.Global = True

                            If re.Test(profile) Then
                                matches = re.Execute(profile)
                                If matches.Count = 2 Then
                                    num1 = CLng(matches(0).Value)
                                    num2 = CLng(matches(1).Value)
                                    cell.Value = "CHEQ_PLT" & Math.Min(num1, num2)
                                End If
                            End If
                            Marshal.ReleaseComObject(re)
                        Else
                            Dim re As Object = CreateObject("VBScript.RegExp")
                            re.Pattern = "\d+"
                            re.Global = True

                            If re.Test(profile) Then
                                matches = re.Execute(profile)
                                If matches.Count = 2 Then
                                    num1 = CLng(matches(0).Value)
                                    num2 = CLng(matches(1).Value)
                                    cell.Value = "PL" & Math.Min(num1, num2)
                                ElseIf matches.Count > 2 Then
                                    minNum = CLng(matches(0).Value)
                                    For i = 0 To matches.Count - 1
                                        minNum = Math.Min(minNum, CLng(matches(i).Value))
                                    Next
                                    cell.Value = "PL" & minNum
                                End If
                            End If
                            Marshal.ReleaseComObject(re)
                        End If
                    End If
                Next

                ' --- Second operation: Consolidate and sort data with improved handling ---
                serialNumber = 1

                ' Use while loop to handle row deletion properly
                i = 8
                While i <= lastRow
                    profile = SafeToString(ws.Cells(i, "B").Value)
                    If profile = "" Then
                        Exit While
                    End If

                    material = SafeToString(ws.Cells(i, "D").Value)
                    totalLength = SafeToDouble(ws.Cells(i, "C").Value)
                    totalWeight = SafeToDouble(ws.Cells(i, "E").Value)
                    totalArea = SafeToDouble(ws.Cells(i, "F").Value)

                    ' Find and sum duplicates with improved string comparison
                    For j = lastRow To i + 1 Step -1
                        If SafeToString(ws.Cells(j, "B").Value) = profile And SafeToString(ws.Cells(j, "D").Value) = material Then
                            totalLength += SafeToDouble(ws.Cells(j, "C").Value)
                            totalWeight += SafeToDouble(ws.Cells(j, "E").Value)
                            totalArea += SafeToDouble(ws.Cells(j, "F").Value)

                            ws.Rows(j).Delete()
                            lastRow = lastRow - 1
                        End If
                    Next

                    If totalLength > 0 Or totalWeight > 0 Or totalArea > 0 Then
                        ws.Cells(i, "A").Value = serialNumber

                        If profile.StartsWith("PL") Or profile.StartsWith("CHEQ_PLT") Then
                            ws.Cells(i, "C").Value = "-"
                        ElseIf totalLength = 0 Then
                            ws.Cells(i, "C").Value = "-"
                        Else
                            ws.Cells(i, "C").Value = totalLength
                        End If

                        ws.Cells(i, "E").Value = totalWeight
                        ws.Cells(i, "F").Value = totalArea
                        serialNumber = serialNumber + 1
                        i = i + 1
                    Else
                        ws.Rows(i).Delete()
                        lastRow = lastRow - 1
                    End If
                End While

                ' Sort data
                With ws.Sort
                    .SortFields.Clear()
                    .SortFields.Add(Key:=ws.Range("B8:B" & lastRow),
                                   SortOn:=XlSortOn.xlSortOnValues,
                                   Order:=XlSortOrder.xlAscending,
                                   DataOption:=XlSortDataOption.xlSortNormal)
                    .SetRange(ws.Range("A8:F" & lastRow))
                    .Header = XlYesNoGuess.xlNo
                    .Apply()
                End With

                ' Reassign serial numbers
                serialNumber = 1
                For i = 8 To lastRow
                    If SafeToString(ws.Cells(i, "B").Value) <> "" Then
                        ws.Cells(i, "A").Value = serialNumber
                        serialNumber = serialNumber + 1
                    End If
                Next

                ' Lock all sheets back
                LockAllSheets(xlApp, wb)

                MsgBox("Data processing completed successfully!", MsgBoxStyle.Information)

            Catch ex As Exception
                MsgBox("An error occurred: " & ex.Message)
                ' Ensure sheets are locked even if error occurs
                Try
                    If xlApp IsNot Nothing And wb IsNot Nothing Then
                        LockAllSheets(xlApp, wb)
                    End If
                Catch
                    ' Ignore protection errors in catch block
                End Try
            Finally
                If ws IsNot Nothing Then Marshal.ReleaseComObject(ws)
                If wb IsNot Nothing Then Marshal.ReleaseComObject(wb)
                If xlApp IsNot Nothing Then Marshal.ReleaseComObject(xlApp)
            End Try
        End Sub

    End Class
End Namespace