Imports System.Runtime.InteropServices
Imports Microsoft.Office.Interop.Excel
Imports Microsoft.VisualBasic
Imports System.Globalization

Namespace RaghavTekNova

    <ComVisible(True)>
    <Guid("D8A47D8E-9F11-4C7C-A77A-AB1234567890")> ' <-- replace with your own GUID
    <ProgId("RaghavTekNova.BoltList")>
    Public Class BoltList

        ' -------------------------
        ' Helper converters
        ' -------------------------
        Private Function ToDouble(o As Object) As Double
            Try
                If o Is Nothing Then Return 0
                ' Direct numeric types
                If TypeOf o Is Double OrElse TypeOf o Is Single OrElse TypeOf o Is Decimal Then
                    Return Convert.ToDouble(o)
                End If
                If TypeOf o Is Integer OrElse TypeOf o Is Long OrElse TypeOf o Is Short OrElse TypeOf o Is Byte Then
                    Return Convert.ToDouble(o)
                End If

                Dim s As String = o.ToString()
                If String.IsNullOrWhiteSpace(s) Then Return 0

                ' Clean common formatting (commas, NBSP)
                s = s.Trim().Replace(",", "").Replace(ChrW(160), "")

                Dim d As Double = 0
                ' Try invariant culture first (works for data like "12.5")
                If Double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, d) Then
                    Return d
                End If
                ' Try current culture as fallback (works for "12,5" etc)
                If Double.TryParse(s, NumberStyles.Any, CultureInfo.CurrentCulture, d) Then
                    Return d
                End If

                ' Last resort: remove any non-numeric chars except dot and minus
                Dim cleaned As New System.Text.StringBuilder()
                For Each ch As Char In s
                    If Char.IsDigit(ch) OrElse ch = "."c OrElse ch = "-"c Then cleaned.Append(ch)
                Next
                If Double.TryParse(cleaned.ToString(), NumberStyles.Any, CultureInfo.InvariantCulture, d) Then
                    Return d
                End If
            Catch ex As Exception
                ' swallow and return 0
            End Try
            Return 0
        End Function

        Private Function ToStr(o As Object) As String
            If o Is Nothing Then Return String.Empty
            Return o.ToString().Trim()
        End Function

        ' -------------------------
        ' B3_RUN_PR_BOLT_LIST_VBA
        ' -------------------------
        Public Sub Run_PR_Bolt_List()
            Try
                Dim app As Application = CType(Marshal.GetActiveObject("Excel.Application"), Application)
                Dim wb As Workbook = app.ActiveWorkbook
                Dim wsSource As Worksheet = wb.Sheets(2)
                Dim wsTarget As Worksheet = wb.Sheets(7) ' system-generated data
                Dim wsStandards As Worksheet = wb.Sheets(8) ' standards

                ' Unprotect sheets
                wsSource.Unprotect("2022")
                wsTarget.Unprotect("2022")
                wsStandards.Unprotect("2022")

                ' Clear previous content
                wsSource.Cells.Clear()

                ' Header / project info (same as original)
                With wsSource
                    .Range("C2").Value = "PROJECT NO:-"
                    .Range("C3").Value = "PROJECT NAME:-"
                    .Range("C4").Value = "CLIENT NAME:-"
                    .Range("C5").Value = "DATE:-"
                    .Range("E5").Value = "REVISION:-"
                    .Range("G5").Value = "TOTAL:-"
                End With

                With wsSource.Range("C2:C5,G5,E5,D2:D5,F5,H5")
                    .Font.Bold = True
                    .Font.Size = 14
                End With

                wsSource.Range("A1:H1").Interior.Color = RGB(248, 203, 173)
                wsSource.Range("A7:H7").Interior.Color = RGB(173, 216, 230)
                wsSource.Range("A6").Interior.Color = RGB(248, 203, 173)

                wsSource.Range("A2:B5").Merge()
                wsSource.Range("D2:G2").Merge()
                wsSource.Range("D3:G3").Merge()
                wsSource.Range("D4:G4").Merge()
                wsSource.Range("A6:H6").Merge()
                wsSource.Range("H2:H4").Merge()

                With wsSource.Range("A1:H1")
                    .Merge()
                    .Value = "BILL OF MATERIAL (BOLT LIST)"
                    .HorizontalAlignment = XlHAlign.xlHAlignCenter
                    .VerticalAlignment = XlVAlign.xlVAlignCenter
                    .Font.Bold = True
                    .Font.Size = 15
                End With

                ' Headers
                wsSource.Range("A7").Value = "SR.NO."
                wsSource.Range("B7").Value = "STANDARD"
                wsSource.Range("C7").Value = "BOLT DIA"
                wsSource.Range("D7").Value = "BOLT LENGTH"
                wsSource.Range("E7").Value = "QTY."
                wsSource.Range("F7").Value = "BOLT GRADE"
                wsSource.Range("G7").Value = "UNIT WEIGHT(Kg)"
                wsSource.Range("H7").Value = "WEIGHT(Kg)"

                With wsSource.Range("A7:H7")
                    .Font.Bold = True
                    .HorizontalAlignment = XlHAlign.xlHAlignCenter
                    .VerticalAlignment = XlVAlign.xlVAlignCenter
                    .Font.Size = 12
                End With

                wsSource.Columns("A").ColumnWidth = 8
                wsSource.Columns("B").ColumnWidth = 12
                wsSource.Columns("C").ColumnWidth = 20
                wsSource.Columns("D").ColumnWidth = 20
                wsSource.Columns("E").ColumnWidth = 12
                wsSource.Columns("F").ColumnWidth = 15
                wsSource.Columns("G").ColumnWidth = 20
                wsSource.Columns("H").ColumnWidth = 15

                ' Project values from system sheet
                wsSource.Range("D2").Value = wsTarget.Range("E1").Value
                wsSource.Range("D3").Value = wsTarget.Range("E2").Value
                wsSource.Range("D4").Value = wsTarget.Range("E3").Value
                wsSource.Range("D5").Value = wsTarget.Range("E4").Value
                wsSource.Range("F5").Value = wsTarget.Range("F4").Value
                wsSource.Range("H5").Value = wsTarget.Range("G4").Value

                ' Last rows
                Dim lastRow As Integer = wsTarget.Cells(wsTarget.Rows.Count, 1).End(XlDirection.xlUp).Row
                Dim lastStandardRow As Integer = wsStandards.Cells(wsStandards.Rows.Count, 1).End(XlDirection.xlUp).Row

                ' Loop through data rows (source system sheet starts at row 8 per original)
                For i As Integer = 8 To lastRow
                    Dim dia As Double = ToDouble(wsTarget.Cells(i, 1).Value)
                    Dim length As Double = ToDouble(wsTarget.Cells(i, 2).Value)
                    Dim qty As Double = ToDouble(wsTarget.Cells(i, 3).Value)
                    Dim grade As String = ToStr(wsTarget.Cells(i, 4).Value)
                    Dim matchFound As Boolean = False
                    Dim unitWeight As Double = 0

                    For j As Integer = 2 To lastStandardRow
                        Dim stdDia As Double = ToDouble(wsStandards.Cells(j, 1).Value)
                        Dim stdLength As Double = ToDouble(wsStandards.Cells(j, 2).Value)
                        unitWeight = ToDouble(wsStandards.Cells(j, 3).Value)

                        If dia = stdDia Then
                            If length <= stdLength Then
                                wsSource.Cells(i, 3).Value = dia
                                wsSource.Cells(i, 4).Value = stdLength
                                wsSource.Cells(i, 5).Value = qty
                                wsSource.Cells(i, 6).Value = grade
                                wsSource.Cells(i, 7).Value = Math.Round(unitWeight, 2)
                                matchFound = True
                                Exit For
                            End If
                        End If
                    Next

                    If Not matchFound Then
                        ' Use safe string display
                        MsgBox("BOLT " & dia.ToString(CultureInfo.InvariantCulture) & " X " & length.ToString(CultureInfo.InvariantCulture) & " not available in standard.", vbExclamation)
                    End If
                Next

                ' Serial numbers
                For i As Integer = 8 To lastRow
                    wsSource.Cells(i, 1).Value = i - 7
                Next

                With wsSource.Range("A8:H" & lastRow)
                    .HorizontalAlignment = XlHAlign.xlHAlignCenter
                    .VerticalAlignment = XlVAlign.xlVAlignCenter
                    .Font.Color = RGB(0, 0, 139)
                End With

                Dim mergeRow As Integer = lastRow + 1
                wsSource.Range("A" & mergeRow & ":H" & mergeRow).Merge()
                wsSource.Range("A" & mergeRow).Value = " "

                With wsSource.Range("A1:H" & mergeRow)
                    .Borders.LineStyle = XlLineStyle.xlContinuous
                End With

                ' Protect sheets back
                wsSource.Protect("2022")
                wsTarget.Protect("2022")
                wsStandards.Protect("2022")

                ' Activate the source sheet (Sheet2)
                wsSource.Activate()

            Catch ex As Exception
                ' Ensure sheets are protected even if error occurs
                Try
                    Dim app As Application = CType(Marshal.GetActiveObject("Excel.Application"), Application)
                    Dim wb As Workbook = app.ActiveWorkbook
                    wb.Sheets(2).Protect("2022")
                    wb.Sheets(7).Protect("2022")
                    wb.Sheets(8).Protect("2022")
                Catch
                    ' Ignore protection errors in cleanup
                End Try
                MsgBox("Run_PR_Bolt_List error: " & ex.Message, vbCritical)
            End Try
        End Sub

        ' -------------------------
        ' B4_Calculate_Total_Bolt_weight
        ' -------------------------
        Public Sub Calculate_Total_Bolt_weight()
            Try
                Dim app As Application = CType(Marshal.GetActiveObject("Excel.Application"), Application)
                Dim wb As Workbook = app.ActiveWorkbook
                Dim ws As Worksheet = wb.Sheets("Sheet2")

                ' Unprotect sheet
                ws.Unprotect("2022")

                Dim lastRow As Integer = ws.Cells(ws.Rows.Count, "E").End(XlDirection.xlUp).Row
                Dim totalSum As Double = 0

                For i As Integer = 8 To lastRow
                    Dim qty As Double = ToDouble(ws.Cells(i, "E").Value)
                    Dim unitWeight As Double = ToDouble(ws.Cells(i, "G").Value)
                    Dim totalWeight As Double = qty * unitWeight
                    ws.Cells(i, "H").Value = totalWeight

                    totalSum += totalWeight

                    Dim cellCValue As Double = ToDouble(ws.Range("C" & i).Value)
                    Dim cellDValue As String = ToStr(ws.Range("D" & i).Value)
                    Dim concatenatedValue As String

                    If cellCValue < 50 Then
                        concatenatedValue = "M-" & cellCValue.ToString(CultureInfo.InvariantCulture) & "X" & cellDValue
                    Else
                        concatenatedValue = "H-" & cellCValue.ToString(CultureInfo.InvariantCulture) & "X" & cellDValue
                    End If

                    ws.Range("B" & i).Value = concatenatedValue
                Next

                ws.Range("H5").Value = totalSum

                ' Protect sheet back
                ws.Protect("2022")

            Catch ex As Exception
                ' Ensure sheet is protected even if error occurs
                Try
                    Dim app As Application = CType(Marshal.GetActiveObject("Excel.Application"), Application)
                    Dim wb As Workbook = app.ActiveWorkbook
                    wb.Sheets("Sheet2").Protect("2022")
                Catch
                    ' Ignore protection errors in cleanup
                End Try
                MsgBox("Calculate_Total_Bolt_weight error: " & ex.Message, vbCritical)
            End Try
        End Sub

        ' -------------------------
        ' B5_Organize_Boltdata
        ' -------------------------
        Public Sub Organize_Boltdata()
            Try
                Dim app As Application = CType(Marshal.GetActiveObject("Excel.Application"), Application)
                Dim wb As Workbook = app.ActiveWorkbook
                Dim ws As Worksheet = app.ActiveSheet

                ' Unprotect sheet
                ws.Unprotect("2022")

                Dim lastRow As Integer = ws.Cells(ws.Rows.Count, "C").End(XlDirection.xlUp).Row
                Dim serialNum As Integer = 1

                For i As Integer = 8 To lastRow
                    If ToStr(ws.Cells(i, "C").Value) <> String.Empty Then
                        Dim boltDia As Double = ToDouble(ws.Cells(i, "C").Value)
                        Dim boltLength As Double = ToDouble(ws.Cells(i, "D").Value)
                        Dim qtySum As Double = ToDouble(ws.Cells(i, "E").Value)
                        Dim totalWeight As Double = ToDouble(ws.Cells(i, "G").Value) * qtySum

                        For j As Integer = i + 1 To lastRow
                            If ToDouble(ws.Cells(j, "C").Value) = boltDia AndAlso ToDouble(ws.Cells(j, "D").Value) = boltLength Then
                                qtySum += ToDouble(ws.Cells(j, "E").Value)
                                totalWeight += ToDouble(ws.Cells(j, "G").Value) * ToDouble(ws.Cells(j, "E").Value)
                                ws.Rows(j).Delete()
                                lastRow -= 1
                                j -= 1
                            End If
                        Next

                        ws.Cells(i, "E").Value = qtySum
                        ws.Cells(i, "H").Value = totalWeight
                    End If
                Next

                For i As Integer = 8 To lastRow
                    ws.Cells(i, "A").Value = serialNum
                    serialNum += 1
                Next

                ' Protect sheet back
                ws.Protect("2022")

                MsgBox("Bolt data organized successfully.", vbInformation)
            Catch ex As Exception
                ' Ensure sheet is protected even if error occurs
                Try
                    Dim app As Application = CType(Marshal.GetActiveObject("Excel.Application"), Application)
                    Dim wb As Workbook = app.ActiveWorkbook
                    app.ActiveSheet.Protect("2022")
                Catch
                    ' Ignore protection errors in cleanup
                End Try
                MsgBox("Organize_Boltdata error: " & ex.Message, vbCritical)
            End Try
        End Sub

    End Class
End Namespace