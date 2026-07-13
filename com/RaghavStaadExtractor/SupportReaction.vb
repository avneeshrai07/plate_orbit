Option Strict Off
Option Explicit On

Imports System.Drawing
Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports Microsoft.Office.Interop.Excel
Imports ExcelApp = Microsoft.Office.Interop.Excel.Application

<ComVisible(True)>
<ClassInterface(ClassInterfaceType.AutoDual)>
<ProgId("RaghavStaadExtractor.SupportReaction")>
Public Class SupportReaction

    Private Const SheetPassword As String = "2022"

    ' Main public method to be called from VBA
    Public Sub ProcessSupportReaction()
        Dim xlApp As ExcelApp = Nothing
        Dim ws As Worksheet = Nothing
        Dim objOpenSTAAD As Object = Nothing

        Try
            ' Get running Excel instance
            xlApp = CType(Marshal.GetActiveObject("Excel.Application"), ExcelApp)
            ws = CType(xlApp.ActiveWorkbook.Sheets("Sheet1"), Worksheet)

            ' PERFORMANCE: Disable screen updating and calculations
            xlApp.ScreenUpdating = False
            xlApp.Calculation = XlCalculation.xlCalculationManual
            xlApp.EnableEvents = False

            UnlockSheetWithPassword(ws)

            ' Connect to STAAD
            Try
                objOpenSTAAD = Marshal.GetActiveObject("StaadPro.OpenSTAAD")
            Catch ex As Exception
                MessageBox.Show("Run StaadPro First", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Return
            End Try

            BackupAndClearSheet(ws)
            WriteJobInfo(ws, objOpenSTAAD)
            FormatHeaders(ws)

            Dim startLcase As Long = CLng(ws.Range("A3").Value)
            Dim endLcase As Long = CLng(ws.Range("B3").Value)

            If startLcase <= 0 OrElse endLcase <= 0 OrElse startLcase > endLcase Then
                MessageBox.Show("Invalid load case range in A3:B3", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Exit Sub
            End If

            Dim nodeCount As Long = objOpenSTAAD.Geometry.GetNoOfSelectedNodes()
            If nodeCount = 0 Then
                MessageBox.Show("No nodes selected in STAAD.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Exit Sub
            End If

            Dim selectedNodes(CInt(nodeCount - 1)) As Integer
            objOpenSTAAD.Geometry.GetSelectedNodes(selectedNodes)

            WriteSupportReactions(ws, objOpenSTAAD, selectedNodes, nodeCount, startLcase, endLcase)

            DrawBorders(ws)

            WriteSummary(ws, 4, "Fx")
            WriteSummary(ws, 5, "Fy")
            WriteSummary(ws, 6, "Fz")
            WriteSummary(ws, 7, "Mx")
            WriteSummary(ws, 8, "My")
            WriteSummary(ws, 9, "Mz")

            FormatSummaryBorders(ws)

            HideRepeatedNodeIDs(ws)
            CreateUniqueDropdownInI5(ws)

            LockSheetWithPassword(ws)

        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message & vbCrLf & vbCrLf & ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            ' PERFORMANCE: Re-enable screen updating and calculations
            If xlApp IsNot Nothing Then
                xlApp.ScreenUpdating = True
                xlApp.Calculation = XlCalculation.xlCalculationAutomatic
                xlApp.EnableEvents = True
            End If

            If objOpenSTAAD IsNot Nothing Then
                Try
                    Marshal.ReleaseComObject(objOpenSTAAD)
                Catch
                End Try
            End If
            If ws IsNot Nothing Then
                Try
                    Marshal.ReleaseComObject(ws)
                Catch
                End Try
            End If
            If xlApp IsNot Nothing Then
                Try
                    Marshal.ReleaseComObject(xlApp)
                Catch
                End Try
            End If
        End Try
    End Sub

    Private Sub UnlockSheetWithPassword(ws As Worksheet)
        Try
            ws.Unprotect(SheetPassword)
        Catch ex As Exception
            ' Sheet might not be protected - ignore
        End Try
    End Sub

    Private Sub LockSheetWithPassword(ws As Worksheet)
        ' Unlock all first
        ws.Cells.Locked = False

        ' Lock all except A3 and B3 as input cells
        ws.Range("A1:Z1000").Locked = True
        ws.Range("A3:B3,A9:B10,K1:T14,V1:AE14").Locked = False

        ' Now protect the sheet (UserInterfaceOnly so macros can edit)
        ws.Protect(SheetPassword, UserInterfaceOnly:=True)
    End Sub

    Private Sub BackupAndClearSheet(ws As Worksheet)
        Dim tempA3 As Object = ws.Range("A3").Value
        Dim tempB3 As Object = ws.Range("B3").Value
        ws.Cells.ClearContents()
        ws.Range("A3").Value = tempA3
        ws.Range("B3").Value = tempB3
    End Sub

    Private Sub WriteJobInfo(ws As Worksheet, objOpenSTAAD As Object)
        Dim jobName As String = ""
        Dim jobClient As String = ""
        Dim enggName As String = ""
        Dim eDate As String = ""
        Dim jobNumber As String = ""
        Dim revision As String = ""
        Dim part As String = ""
        Dim reference As String = ""
        Dim checkerName As String = ""
        Dim chDate As String = ""
        Dim approverName As String = ""
        Dim aDate As String = ""
        Dim comments As String = ""

        Try
            ' Direct COM call - VB.NET handles ByRef parameters automatically
            objOpenSTAAD.GetFullJobInfo(jobName, jobClient, enggName, eDate, jobNumber, revision,
                                       part, reference, checkerName, chDate, approverName, aDate, comments)
        Catch ex As Exception
            ' If GetFullJobInfo fails, continue with empty values
            Debug.WriteLine("GetFullJobInfo error: " & ex.Message)
        End Try

        ' PERFORMANCE: Use arrays for bulk data writing
        Dim currentDate As String = DateTime.Now.ToString("dd-MMM-yyyy")
        Dim currentTime As String = DateTime.Now.ToString("hh:mm tt")

        ' Assign values
        ws.Range("C2").Value = "JOB NAME:-"
        ws.Range("D2").Value = jobName
        ws.Range("C4").Value = "JOB NUMBER:-"
        ws.Range("D4").Value = jobNumber
        ws.Range("C6").Value = "ENGINEER NAME:-"
        ws.Range("D6").Value = enggName
        ws.Range("F6").Value = "CHECKER NAME:-"
        ws.Range("H6").Value = checkerName
        ws.Range("C8").Value = "APPROVER NAME:-"
        ws.Range("D8").Value = approverName
        ws.Range("G10").Value = "REVISION:-"
        ws.Range("H10").Value = revision
        ws.Range("C10").Value = "DATE:-"
        ws.Range("D10").Value = currentDate
        ws.Range("E10").Value = "TIME:-"
        ws.Range("F10").Value = currentTime

        ' Merge cells - do all at once where possible
        ws.Range("C2:C3").Merge()
        ws.Range("D2:I3").Merge()
        ws.Range("C4:C5").Merge()
        ws.Range("D4:I5").Merge()
        ws.Range("C6:C7").Merge()
        ws.Range("D6:E7").Merge()
        ws.Range("F6:G7").Merge()
        ws.Range("H6:I7").Merge()
        ws.Range("C8:C9").Merge()
        ws.Range("D8:E9").Merge()
        ws.Range("G10:G11").Merge()
        ws.Range("H10:H11").Merge()
        ws.Range("C10:C11").Merge()
        ws.Range("D10:D11").Merge()
        ws.Range("E10:E11").Merge()
        ws.Range("F10:F11").Merge()
        ws.Range("F8:G9").Merge()
        ws.Range("H8:I9").Merge()
        ws.Range("A2:B2").Merge()
        ws.Range("A4:B8").Merge()

        ' Format job info cells - apply to entire range at once
        With ws.Range("C2:I11")
            .Borders.LineStyle = XlLineStyle.xlContinuous
            .HorizontalAlignment = XlHAlign.xlHAlignLeft
            .VerticalAlignment = XlVAlign.xlVAlignCenter
            .Font.Size = 13
            .Font.Color = RGB(0, 0, 0)
            .Font.Bold = True
        End With

        ' Format input cells
        With ws.Range("A3:B3,A9:B9")
            .Borders.LineStyle = XlLineStyle.xlContinuous
            .HorizontalAlignment = XlHAlign.xlHAlignCenter
            .VerticalAlignment = XlVAlign.xlVAlignCenter
            .Font.Size = 13
            .Font.Color = RGB(0, 0, 0)
            .Font.Bold = True
            .Interior.Color = RGB(23, 233, 188)
        End With
    End Sub

    Private Sub FormatHeaders(ws As Worksheet)
        With ws.Range("A1:I1")
            .Merge()
            .Interior.Color = RGB(173, 216, 230)
            .HorizontalAlignment = XlHAlign.xlHAlignCenter
            .VerticalAlignment = XlVAlign.xlVAlignCenter
            .Font.Size = 12
            .Font.Bold = True
        End With

        With ws.Range("A12:I12")
            .Interior.Color = RGB(173, 216, 230)
            .HorizontalAlignment = XlHAlign.xlHAlignCenter
            .VerticalAlignment = XlVAlign.xlVAlignCenter
            .Font.Size = 12
            .Font.Bold = True
        End With
        ws.Range("K1").Value = "Global"

        With ws.Range("C2:I5")
            .HorizontalAlignment = XlHAlign.xlHAlignLeft
            .VerticalAlignment = XlVAlign.xlVAlignCenter
            .Font.Bold = True
        End With

        With ws.Range("A1:I6").Borders
            .LineStyle = XlLineStyle.xlContinuous
            .Weight = XlBorderWeight.xlThin
        End With

        ' PERFORMANCE: Batch write header values using arrays
        Dim headerLabels() As Object = {
            "Node ID", "Load Case", "Load Title", "Horizontal", "Vertical", "Horizontal"
        }
        ws.Range("A13:F13").Value = headerLabels

        ws.Range("G13").Value = "Moment"
        ws.Range("G13:I13").Merge()

        Dim detailHeaders() As Object = {
            "Fx (kN)", "Fy (kN)", "Fz (kN)", "Mx (kN-m)", "My (kN-m)", "Mz (kN-m)"
        }
        ws.Range("D14:I14").Value = detailHeaders

        ' Merge header cells
        ws.Range("A13:A14").Merge()
        ws.Range("B13:B14").Merge()
        ws.Range("C13:C14").Merge()

        ' FOR GLOBAL MINIMUM & MAXIMUM
        ws.Range("K1:K2").Merge()
        ws.Range("L1").Value = "Node ID"
        ws.Range("L1:L2").Merge()
        ws.Range("M1").Value = "Load Case"
        ws.Range("M1:M2").Merge()
        ws.Range("N1").Value = "Load Title"
        ws.Range("N1:N2").Merge()

        Dim globalHeaders1() As Object = {"Horizontal", "Vertical", "Horizontal"}
        ws.Range("O1:Q1").Value = globalHeaders1

        ws.Range("R1").Value = "Moment"
        ws.Range("R1:T1").Merge()

        Dim globalHeaders2() As Object = {
            "Fx (kN)", "Fy (kN)", "Fz (kN)", "Mx (kN-m)", "My (kN-m)", "Mz (kN-m)"
        }
        ws.Range("O2:T2").Value = globalHeaders2
        ws.Range("Z2:AE2").Value = globalHeaders2

        ' FOR LOCAL MINIMUM & MAXIMUM
        ws.Range("V1").Value = "Local"
        ws.Range("V1:V2").Merge()
        ws.Range("W1").Value = "Node ID"
        ws.Range("W1:W2").Merge()
        ws.Range("X1").Value = "Load Case"
        ws.Range("X1:X2").Merge()
        ws.Range("Y1").Value = "Load Title"
        ws.Range("Y1:Y2").Merge()

        Dim localHeaders1() As Object = {"Horizontal", "Vertical", "Horizontal"}
        ws.Range("Z1:AB1").Value = localHeaders1

        ws.Range("AC1").Value = "Moment"
        ws.Range("AC1:AE1").Merge()

        Dim localLabels() As Object = {
            "Fx(Max)", "Fx(min)", "Fy(Max)", "Fy(min)", "Fz(Max)", "Fz(min)",
            "Mx(Max)", "Mx(min)", "My(Max)", "My(min)", "Mz(Max)", "Mz(min)"
        }

        For i As Integer = 0 To localLabels.Length - 1
            ws.Cells(3 + i, 22).Value = localLabels(i) ' Column V = 22
        Next

        ' Format all headers at once
        With ws.Range("A4:A3,B3,A13:I14,K1:T1,V1:AE1")
            .Font.Bold = True
            .HorizontalAlignment = XlHAlign.xlHAlignCenter
            .VerticalAlignment = XlVAlign.xlVAlignCenter
        End With

        ' Set column widths
        ws.Columns("A:B").ColumnWidth = 12
        ws.Columns("D:I").ColumnWidth = 12
        ws.Columns("D:M").ColumnWidth = 12
        ws.Columns("O:T").ColumnWidth = 12
        ws.Columns("C").ColumnWidth = 40
        ws.Columns("N").ColumnWidth = 40
        ws.Columns("Y").ColumnWidth = 40

        ' HIDE COLUMN V TO AE
        ws.Columns("V:AE").EntireColumn.Hidden = True
    End Sub

    Private Sub DrawBorders(ws As Worksheet)
        Dim lastRow As Long = ws.Cells(ws.Rows.Count, "B").End(XlDirection.xlUp).Row
        If lastRow < 15 Then Return
        With ws.Range("A13:I" & lastRow).Borders
            .LineStyle = XlLineStyle.xlContinuous
            .Weight = XlBorderWeight.xlThin
        End With
    End Sub

    Private Function GetNodeFromRow(ws As Worksheet, searchRow As Long) As Object
        For i As Long = searchRow To 15 Step -1
            If IsNumeric(ws.Cells(i, 1).Value) Then
                Return ws.Cells(i, 1).Value
            End If
        Next
        Return "Not Found"
    End Function

    Private Sub WriteSummarySTAADStyle(ws As Worksheet)
        Dim lastRow As Long = ws.Cells(ws.Rows.Count, "B").End(XlDirection.xlUp).Row
        If lastRow < 15 Then Exit Sub

        ' PERFORMANCE: Read all reaction data into array once (columns D-I = Fx,Fy,Fz,Mx,My,Mz)
        Dim dataRange As Range = ws.Range(ws.Cells(15, 1), ws.Cells(lastRow, 9))
        Dim allData As Object = dataRange.Value

        ' Initialize min/max tracking for all 6 reactions
        Dim reactions() As String = {"Fx", "Fy", "Fz", "Mx", "My", "Mz"}
        Dim colIndices() As Integer = {4, 5, 6, 7, 8, 9}  ' D, E, F, G, H, I columns

        ' Track absolute max, positive max, and negative min for each reaction
        Dim absMaxVals(5) As Double
        Dim absMaxRows(5) As Long
        Dim posMaxVals(5) As Double
        Dim posMaxRows(5) As Long
        Dim negMinVals(5) As Double
        Dim negMinRows(5) As Long

        ' Initialize arrays
        For i As Integer = 0 To 5
            absMaxVals(i) = 0
            absMaxRows(i) = 0
            posMaxVals(i) = -1.0E+99
            posMaxRows(i) = 0
            negMinVals(i) = 1.0E+99
            negMinRows(i) = 0
        Next

        ' Single pass through all data - STAAD style analysis
        For row As Long = 1 To UBound(allData, 1)
            For i As Integer = 0 To 5
                Dim colOffset As Integer = colIndices(i) - 1  ' Adjust for array indexing
                If IsNumeric(allData(row, colOffset)) Then
                    Dim val As Double = CDbl(allData(row, colOffset))
                    Dim absVal As Double = Math.Abs(val)

                    ' Track absolute maximum (like STAAD's overall max)
                    If absVal > Math.Abs(absMaxVals(i)) Then
                        absMaxVals(i) = val
                        absMaxRows(i) = row + 14  ' Adjust for starting at row 15
                    End If

                    ' Track positive maximum
                    If val > 0 AndAlso val > posMaxVals(i) Then
                        posMaxVals(i) = val
                        posMaxRows(i) = row + 14
                    End If

                    ' Track negative minimum
                    If val < 0 AndAlso val < negMinVals(i) Then
                        negMinVals(i) = val
                        negMinRows(i) = row + 14
                    End If
                End If
            Next
        Next

        ' Write STAAD-style summary (Absolute Max and its opposite)
        For i As Integer = 0 To 5
            Dim baseRow As Long = 3 + (i * 2)  ' Rows 3,5,7,9,11,13

            If absMaxRows(i) > 0 Then
                Dim absMaxNode As Object = GetNodeFromRow(ws, absMaxRows(i))

                ' Determine if we found both positive and negative values
                Dim hasPositive As Boolean = (posMaxRows(i) > 0)
                Dim hasNegative As Boolean = (negMinRows(i) > 0)

                If hasPositive AndAlso hasNegative Then
                    ' STAAD Style: Show larger absolute value first, then opposite
                    Dim displayMaxRow As Long
                    Dim displayMinRow As Long
                    Dim displayMaxNode As Object
                    Dim displayMinNode As Object

                    If Math.Abs(posMaxVals(i)) >= Math.Abs(negMinVals(i)) Then
                        ' Positive is larger - show positive max, then negative min
                        displayMaxRow = posMaxRows(i)
                        displayMinRow = negMinRows(i)
                        displayMaxNode = GetNodeFromRow(ws, displayMaxRow)
                        displayMinNode = GetNodeFromRow(ws, displayMinRow)
                    Else
                        ' Negative is larger - show negative min first, then positive max
                        displayMaxRow = negMinRows(i)
                        displayMinRow = posMaxRows(i)
                        displayMaxNode = GetNodeFromRow(ws, displayMaxRow)
                        displayMinNode = GetNodeFromRow(ws, displayMinRow)
                    End If

                    ' Write Max (larger absolute value)
                    ws.Range("K" & baseRow).Value = reactions(i) & "(Max)"
                    ws.Range("L" & baseRow).Value = displayMaxNode
                    ws.Range("M" & baseRow & ":T" & baseRow).Value = ws.Range("B" & displayMaxRow & ":I" & displayMaxRow).Value

                    ' Write Min (opposite sign)
                    ws.Range("K" & (baseRow + 1)).Value = reactions(i) & "(Min)"
                    ws.Range("L" & (baseRow + 1)).Value = displayMinNode
                    ws.Range("M" & (baseRow + 1) & ":T" & (baseRow + 1)).Value = ws.Range("B" & displayMinRow & ":I" & displayMinRow).Value

                ElseIf hasPositive Then
                    ' Only positive values found
                    Dim displayMaxRow As Long = posMaxRows(i)
                    Dim displayMaxNode As Object = GetNodeFromRow(ws, displayMaxRow)

                    ws.Range("K" & baseRow).Value = reactions(i) & "(Max)"
                    ws.Range("L" & baseRow).Value = displayMaxNode
                    ws.Range("M" & baseRow & ":T" & baseRow).Value = ws.Range("B" & displayMaxRow & ":I" & displayMaxRow).Value

                    ws.Range("K" & (baseRow + 1)).Value = reactions(i) & "(Min)"
                    ws.Range("L" & (baseRow + 1)).Value = "-"
                    ws.Range("M" & (baseRow + 1) & ":T" & (baseRow + 1)).Value = 0

                ElseIf hasNegative Then
                    ' Only negative values found
                    Dim displayMinRow As Long = negMinRows(i)
                    Dim displayMinNode As Object = GetNodeFromRow(ws, displayMinRow)

                    ws.Range("K" & baseRow).Value = reactions(i) & "(Max)"
                    ws.Range("L" & baseRow).Value = "-"
                    ws.Range("M" & baseRow & ":T" & baseRow).Value = 0

                    ws.Range("K" & (baseRow + 1)).Value = reactions(i) & "(Min)"
                    ws.Range("L" & (baseRow + 1)).Value = displayMinNode
                    ws.Range("M" & (baseRow + 1) & ":T" & (baseRow + 1)).Value = ws.Range("B" & displayMinRow & ":I" & displayMinRow).Value
                Else
                    ' All zeros
                    ws.Range("K" & baseRow).Value = reactions(i) & "(Max)"
                    ws.Range("L" & baseRow).Value = "-"
                    ws.Range("M" & baseRow & ":T" & baseRow).Value = 0

                    ws.Range("K" & (baseRow + 1)).Value = reactions(i) & "(Min)"
                    ws.Range("L" & (baseRow + 1)).Value = "-"
                    ws.Range("M" & (baseRow + 1) & ":T" & (baseRow + 1)).Value = 0
                End If
            Else
                ' No data found
                ws.Range("K" & baseRow).Value = reactions(i) & "(Max)"
                ws.Range("L" & baseRow).Value = "-"
                ws.Range("M" & baseRow & ":T" & baseRow).Value = 0

                ws.Range("K" & (baseRow + 1)).Value = reactions(i) & "(Min)"
                ws.Range("L" & (baseRow + 1)).Value = "-"
                ws.Range("M" & (baseRow + 1) & ":T" & (baseRow + 1)).Value = 0
            End If
        Next
    End Sub

    ' Legacy method - kept for reference but not used
    Private Sub WriteSummary(ws As Worksheet, colIndex As Long, label As String)
        Dim maxVal As Double = -1.0E+99
        Dim minVal As Double = 1.0E+99
        Dim maxRow As Long = 0
        Dim minRow As Long = 0
        Dim maxNode As Object = Nothing
        Dim minNode As Object = Nothing

        Dim lastRow As Long = ws.Cells(ws.Rows.Count, "B").End(XlDirection.xlUp).Row
        If lastRow < 15 Then Exit Sub

        ' PERFORMANCE: Read entire column into array once
        Dim dataRange As Range = ws.Range(ws.Cells(15, colIndex), ws.Cells(lastRow, colIndex))
        Dim values As Object = dataRange.Value

        For i As Long = 1 To UBound(values, 1)
            If IsNumeric(values(i, 1)) Then
                Dim val As Double = CDbl(values(i, 1))
                If val > maxVal Then
                    maxVal = val
                    maxRow = i + 14  ' Adjust for starting at row 15
                End If
                If val < minVal Then
                    minVal = val
                    minRow = i + 14
                End If
            End If
        Next

        If maxRow = 0 Then Return

        maxNode = GetNodeFromRow(ws, maxRow)
        minNode = GetNodeFromRow(ws, minRow)

        Dim baseRow As Long = 0
        Select Case label
            Case "Fx"
                baseRow = 3
            Case "Fy"
                baseRow = 5
            Case "Fz"
                baseRow = 7
            Case "Mx"
                baseRow = 9
            Case "My"
                baseRow = 11
            Case "Mz"
                baseRow = 13
        End Select

        ws.Range("K" & baseRow).Value = label & "(Max)"
        ws.Range("L" & baseRow).Value = maxNode
        If maxVal = 0 Then
            ws.Range("M" & baseRow & ":T" & baseRow).Value = 0
        Else
            ws.Range("M" & baseRow & ":T" & baseRow).Value = ws.Range("B" & maxRow & ":I" & maxRow).Value
        End If

        ws.Range("K" & (baseRow + 1)).Value = label & "(Min)"
        ws.Range("L" & (baseRow + 1)).Value = minNode
        If minVal = 0 Then
            ws.Range("M" & (baseRow + 1) & ":T" & (baseRow + 1)).Value = 0
        Else
            ws.Range("M" & (baseRow + 1) & ":T" & (baseRow + 1)).Value = ws.Range("B" & minRow & ":I" & minRow).Value
        End If
    End Sub

    Private Sub FormatSummaryBorders(ws As Worksheet)
        With ws.Range("K1:T14").Borders
            .LineStyle = XlLineStyle.xlContinuous
            .Weight = XlBorderWeight.xlThin
        End With
    End Sub

    Private Sub HideRepeatedNodeIDs(ws As Worksheet)
        Dim lastRow As Long = ws.Cells(ws.Rows.Count, "B").End(XlDirection.xlUp).Row
        If lastRow < 15 Then Return

        ' PERFORMANCE: Read entire column into array
        Dim nodeRange As Range = ws.Range(ws.Cells(15, 1), ws.Cells(lastRow, 1))
        Dim nodeValues As Object = nodeRange.Value
        Dim currentNode As Object = ""

        ' Process and prepare font colors
        For i As Long = 1 To UBound(nodeValues, 1)
            If Not nodeValues(i, 1).Equals(currentNode) Then
                currentNode = nodeValues(i, 1)
                ws.Cells(14 + i, 1).Font.Color = ColorTranslator.ToOle(Color.Black)
            Else
                ws.Cells(14 + i, 1).Font.Color = RGB(255, 255, 255)
            End If
        Next
    End Sub

    Private Sub CreateUniqueDropdownInI5(ws As Worksheet)
        Dim lastRow As Long
        Dim outputRow As Long
        Dim dict As New Dictionary(Of String, Object)
        Dim helperRange As Range

        lastRow = ws.Cells(ws.Rows.Count, "A").End(XlDirection.xlUp).Row
        If lastRow < 15 Then Exit Sub

        ' PERFORMANCE: Read entire column into array
        Dim nodeRange As Range = ws.Range("A15:A" & lastRow)
        Dim nodeValues As Object = nodeRange.Value

        ' Collect unique values from array
        For i As Long = 1 To UBound(nodeValues, 1)
            If Not IsNothing(nodeValues(i, 1)) AndAlso Not String.IsNullOrEmpty(nodeValues(i, 1).ToString()) Then
                Dim cellValue As String = nodeValues(i, 1).ToString()
                If Not dict.ContainsKey(cellValue) Then
                    dict.Add(cellValue, Nothing)
                End If
            End If
        Next

        If dict.Count = 0 Then Return

        ' PERFORMANCE: Write unique values as array to helper column
        ws.Range(ws.Cells(1, 38), ws.Cells(1000, 38)).ClearContents()

        Dim uniqueArray(dict.Count - 1, 0) As Object
        outputRow = 0
        For Each key As String In dict.Keys
            uniqueArray(outputRow, 0) = key
            outputRow += 1
        Next

        ' Write array in one operation
        ws.Range(ws.Cells(1, 38), ws.Cells(dict.Count, 38)).Value = uniqueArray

        ' Set dropdown in I5
        helperRange = ws.Range(ws.Cells(1, 38), ws.Cells(dict.Count, 38))
        With ws.Range("I5").Validation
            .Delete()
            .Add(XlDVType.xlValidateList,
                 XlDVAlertStyle.xlValidAlertStop,
                 XlFormatConditionOperator.xlBetween,
                 "=" & helperRange.Address(False, False))
            .IgnoreBlank = True
            .InCellDropdown = True
            .ShowInput = True
            .ShowError = True
        End With

        ws.Columns(38).Hidden = True
    End Sub

    Private Sub WriteSupportReactions(ws As Worksheet, objOpenSTAAD As Object, selectedNodes() As Integer,
                                      nodeCount As Long, startLcase As Long, endLcase As Long)
        Dim totalRows As Long = nodeCount * (endLcase - startLcase + 1)

        ' PERFORMANCE: Pre-allocate 2D array for bulk writing
        Dim dataArray(totalRows - 1, 8) As Object
        Dim arrayIndex As Long = 0

        ' Cache load case titles once
        Dim loadTitles As New Dictionary(Of Long, String)
        For lc As Long = startLcase To endLcase
            Try
                loadTitles(lc) = objOpenSTAAD.Load.GetLoadCaseTitle(CInt(lc))
            Catch
                loadTitles(lc) = "LC " & lc.ToString()
            End Try
        Next

        ' Process all data into array
        For i As Integer = 0 To CInt(nodeCount) - 1
            Dim nodeId As Integer = selectedNodes(i)

            For lc As Long = startLcase To endLcase
                Dim dReactionArray(0 To 5) As Double
                Dim RetVal As Object = Nothing

                Try
                    ' Call GetSupportReactions
                    RetVal = objOpenSTAAD.Output.GetSupportReactions(nodeId, CInt(lc), dReactionArray)

                    ' Store in array
                    dataArray(arrayIndex, 0) = nodeId
                    dataArray(arrayIndex, 1) = lc
                    dataArray(arrayIndex, 2) = loadTitles(lc)

                    ' Process reaction values
                    For count As Integer = 0 To 5
                        Dim formattedValue As Double = dReactionArray(count)

                        If Math.Abs(formattedValue - Math.Truncate(formattedValue)) < 0.0000001 Then
                            dataArray(arrayIndex, count + 3) = Math.Truncate(formattedValue)
                        Else
                            dataArray(arrayIndex, count + 3) = Math.Round(formattedValue, 3)
                        End If
                    Next

                Catch ex As Exception
                    Debug.WriteLine($"Error for Node {nodeId}, LC {lc}: {ex.Message}")
                    ' Fill with defaults
                    dataArray(arrayIndex, 0) = nodeId
                    dataArray(arrayIndex, 1) = lc
                    dataArray(arrayIndex, 2) = loadTitles(lc)
                    For count As Integer = 3 To 8
                        dataArray(arrayIndex, count) = 0
                    Next
                Finally
                    arrayIndex += 1
                End Try
            Next
        Next

        ' PERFORMANCE: Write entire array to Excel in ONE operation
        If totalRows > 0 Then
            ws.Range(ws.Cells(15, 1), ws.Cells(14 + totalRows, 9)).Value = dataArray
        End If
    End Sub

    Private Function RGB(r As Integer, g As Integer, b As Integer) As Integer
        Return r + (g * 256) + (b * 65536)
    End Function

End Class