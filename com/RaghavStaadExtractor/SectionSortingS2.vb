Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Runtime.InteropServices
Imports System.Reflection
Imports Excel = Microsoft.Office.Interop.Excel
Imports VBIDE = Microsoft.Vbe.Interop

<ComVisible(True)>
<Guid("FC14406F-BD26-4667-B360-98D9A389B651")>
<ClassInterface(ClassInterfaceType.AutoDual)>
Public Class SectionSortingS2

    ' Constructor with license check
    Public Sub New()
        EnsureLicense() ' ✅ License check runs instantly
    End Sub

    ' License validation method
    Private Sub EnsureLicense()
        Dim validator As New ExecutionValidation()
        If Not validator.IsLicenceValid() Then
            Throw New UnauthorizedAccessException("License check failed.")
        End If
    End Sub

    Public Sub SummerizedToSheet2()
        Try
            Dim excelApp As Excel.Application = CType(Marshal.GetActiveObject("Excel.Application"), Excel.Application)
            Dim wb As Excel.Workbook = excelApp.ActiveWorkbook
            Dim wsSource As Excel.Worksheet = CType(wb.Sheets("Sheet1"), Excel.Worksheet)
            Dim wsDest As Excel.Worksheet = CType(wb.Sheets("Sheet2"), Excel.Worksheet)

            ' Unprotect Sheet2 and Workbook before clearing
            UnprotectSheetAndWorkbook(wsDest)

            wsDest.Cells.Clear()
            wsDest.Range("B8").Value = "THK./SECTION"
            wsDest.Range("C8").Value = "LENGTH (mm)"
            wsDest.Range("D8").Value = "WIDTH(mm)"
            wsDest.Range("E8").Value = "WEIGHT(Kg)"

            Dim lastRow As Integer = wsSource.Cells(wsSource.Rows.Count, "C").End(Excel.XlDirection.xlUp).Row
            Dim data(,) As Object = wsSource.Range("A9", $"AQ{lastRow}").Value2

            Dim output As New List(Of Object())()

            ' Process all rows (including 697) for WEB, TOP FLANGE, BOTTOM FLANGE
            For i As Integer = 1 To data.GetLength(0)
                Dim columnFValue As Double = ValObj(data(i, 6)) ' Column F is index 6
                Dim length As Double = ValObj(data(i, 3)) * 1000

                ' Check if this is property 697 with AK having value and AL being blank
                If columnFValue = 697 Then
                    Dim akValue As Object = data(i, 37) ' Column AK (index 37)
                    Dim alValue As Object = data(i, 38) ' Column AL (index 38)

                    ' If AK has value and AL is blank, use standard logic
                    If Not IsNothing(akValue) AndAlso Convert.ToString(akValue).Trim() <> "" AndAlso
                       (IsNothing(alValue) OrElse Convert.ToString(alValue).Trim() = "") Then

                        ' Use standard logic for this 697 row
                        AddIfValid(output, data(i, 25), length, data(i, 36), data(i, 40)) ' WEB
                        AddIfValid(output, data(i, 27), length, data(i, 26), data(i, 41)) ' TOP FLANGE
                        AddIfValid(output, data(i, 29), length, data(i, 28), data(i, 42)) ' BOTTOM FLANGE
                    Else
                        ' Use property 697 specific logic
                        AddIfValidFor697(output, data(i, 37), length, data(i, 40)) ' WEB: AK (index 37) with weight from AN (index 40)
                        AddIfValidFor697(output, data(i, 38), length, data(i, 41)) ' TOP FLANGE: AL (index 38) with weight from AO (index 41)
                        AddIfValidFor697(output, data(i, 39), length, data(i, 42)) ' BOTTOM FLANGE: AM (index 39) with weight from AP (index 42)
                    End If
                Else
                    ' Standard logic for non-697 rows
                    AddIfValid(output, data(i, 25), length, data(i, 36), data(i, 40)) ' WEB
                    AddIfValid(output, data(i, 27), length, data(i, 26), data(i, 41)) ' TOP FLANGE
                    AddIfValid(output, data(i, 29), length, data(i, 28), data(i, 42)) ' BOTTOM FLANGE
                End If
            Next

            ' Process rows where columns X-AJ are all zero
            For i As Integer = 1 To data.GetLength(0)
                Dim columnFValue As Double = ValObj(data(i, 6)) ' Column F is index 6
                Dim length As Double = ValObj(data(i, 3)) * 1000

                Dim allZero As Boolean = True
                For j As Integer = 24 To 36
                    If ValObj(data(i, j)) <> 0 Then
                        allZero = False
                        Exit For
                    End If
                Next

                If allZero Then
                    If columnFValue = 697 Then
                        Dim akValue As Object = data(i, 37) ' Column AK (index 37)
                        Dim alValue As Object = data(i, 38) ' Column AL (index 38)

                        ' If AK has value and AL is blank, use standard logic
                        If Not IsNothing(akValue) AndAlso Convert.ToString(akValue).Trim() <> "" AndAlso
                           (IsNothing(alValue) OrElse Convert.ToString(alValue).Trim() = "") Then

                            ' Use standard logic - add AK as section with weight from AQ
                            Dim section As Object = data(i, 37)
                            Dim weight As Object = data(i, 43)
                            output.Add(New Object() {section, length, "", weight})
                        Else
                            ' Only process AK for 697 if the flange columns (AK, AL, AM) are empty
                            If (IsNothing(data(i, 37)) OrElse Convert.ToString(data(i, 37)).Trim() = "") AndAlso
                               (IsNothing(data(i, 38)) OrElse Convert.ToString(data(i, 38)).Trim() = "") AndAlso
                               (IsNothing(data(i, 39)) OrElse Convert.ToString(data(i, 39)).Trim() = "") Then

                                Dim section As Object = data(i, 37) ' Column AK
                                Dim weight As Object = data(i, 43)  ' Column AQ

                                ' Handle 2x in section name for zero columns case
                                Dim sectionStr As String = Convert.ToString(section).Trim()
                                If sectionStr.StartsWith("2x", StringComparison.OrdinalIgnoreCase) Then
                                    Dim cleanSection As String = sectionStr.Substring(2)
                                    Dim parsedSection = ParseSectionName(cleanSection)
                                    Dim halfWeight As Double = ValObj(weight) / 2
                                    ' Add two rows with parsed thickness and width, half weight each
                                    output.Add(New Object() {parsedSection.Thickness, length, parsedSection.Width, halfWeight})
                                    output.Add(New Object() {parsedSection.Thickness, length, parsedSection.Width, halfWeight})
                                Else
                                    Dim parsedSection = ParseSectionName(Convert.ToString(section))
                                    output.Add(New Object() {parsedSection.Thickness, length, parsedSection.Width, weight})
                                End If
                            End If
                        End If
                    Else
                        ' Standard logic for non-697 rows
                        Dim section As Object = data(i, 37)
                        Dim weight As Object = data(i, 43)
                        output.Add(New Object() {section, length, "", weight})
                    End If
                End If
            Next

            ' Process rows where column F has values 613, 614, or 615
            For i As Integer = 1 To data.GetLength(0)
                Dim columnFValue As Double = ValObj(data(i, 6)) ' Column F is index 6

                If columnFValue = 613 OrElse columnFValue = 614 OrElse columnFValue = 615 Then
                    Dim section As Object = data(i, 37)  ' Column AK (index 37)
                    Dim length As Double = ValObj(data(i, 3)) * 1000  ' Column C * 1000
                    Dim weight As Object = data(i, 40)  ' Column AN (index 40)

                    ' Add to output if section name and weight are valid
                    If Not IsNothing(section) AndAlso Convert.ToString(section).Trim() <> "" AndAlso ValObj(weight) <> 0 Then
                        output.Add(New Object() {section, length, "", weight})
                    End If
                End If
            Next

            Dim sorted = output.OrderBy(Function(r) Convert.ToString(r(0))).ToList()

            Dim destRange As Excel.Range = wsDest.Range("B9", $"E{8 + sorted.Count}")
            Dim outData(sorted.Count - 1, 3) As Object

            For i As Integer = 0 To sorted.Count - 1
                For j As Integer = 0 To 3
                    outData(i, j) = sorted(i)(j)
                Next
            Next

            destRange.Value2 = outData
            wsDest.Range("A:E").EntireColumn.AutoFit()

        Catch
        End Try
    End Sub

    Private Sub UnprotectSheetAndWorkbook(ws As Excel.Worksheet)
        Dim wb As Excel.Workbook = CType(ws.Parent, Excel.Workbook)
        Try
            ws.Unprotect("2022")
        Catch
        End Try
        Try
            wb.Unprotect("2022")
        Catch
        End Try
    End Sub

    Private Sub AddIfValidFor697(list As List(Of Object()), thk As Object, length As Double, weight As Object)
        ' For property 697: Without 2x keep original weight, with 2x divide into 2 rows
        If Not IsNothing(thk) AndAlso Convert.ToString(thk).Trim() <> "" AndAlso ValObj(weight) <> 0 Then
            Dim thkStr As String = Convert.ToString(thk).Trim()
            Dim fullWeight As Double = ValObj(weight) ' Get full weight

            ' Check if section name starts with "2x"
            If thkStr.StartsWith("2x", StringComparison.OrdinalIgnoreCase) Then
                Dim cleanThk As String = thkStr.Substring(2) ' Remove "2x" prefix
                Dim parsedSection = ParseSectionName(cleanThk)
                Dim halfWeight As Double = fullWeight / 2 ' Divide by 2 for "2x" case

                ' Add two rows with parsed thickness and width, each with half weight
                list.Add(New Object() {parsedSection.Thickness, length, parsedSection.Width, halfWeight})
                list.Add(New Object() {parsedSection.Thickness, length, parsedSection.Width, halfWeight})
            Else
                ' Normal case - parse section name for single row with original weight (no division)
                Dim parsedSection = ParseSectionName(thkStr)
                list.Add(New Object() {parsedSection.Thickness, length, parsedSection.Width, fullWeight})
            End If
        End If
    End Sub

    Private Function ParseSectionName(sectionName As String) As (Thickness As String, Width As String)
        Dim section As String = Convert.ToString(sectionName).Trim()

        ' Check if section contains "x" (like "622x8")
        If section.Contains("x") Then
            Dim parts() As String = section.Split("x"c)
            If parts.Length = 2 Then
                ' Return: Thickness = second part (8), Width = first part (622)
                Return (parts(1).Trim(), parts(0).Trim())
            End If
        End If

        ' If no "x" found or invalid format, return original as thickness and empty width
        Return (section, "")
    End Function

    Private Sub AddIfValid(list As List(Of Object()), thk As Object, length As Double, width As Object, weight As Object)
        If ValObj(thk) <> 0 AndAlso ValObj(weight) <> 0 Then
            list.Add(New Object() {thk, length, width, weight})
        End If
    End Sub

    Private Function ValObj(v As Object) As Double
        Dim d As Double
        If Double.TryParse(Convert.ToString(v), d) Then
            Return d
        Else
            Return 0
        End If
    End Function
End Class