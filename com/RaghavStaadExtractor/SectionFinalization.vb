Imports System.Runtime.InteropServices
Imports Microsoft.Office.Interop.Excel

<ComVisible(True)>
<Guid("1B3A4D71-5A2D-4B7E-9A2F-1EFA2E4C0A55")>
<ProgId("RaghavStaadExtractor.SectionFinalization")>
Public Class SectionFinalization

    <ComVisible(True)>
    Public Sub FinalSummarySheet2()

        Dim excelApp As Application = Nothing
        Dim ws As Worksheet = Nothing
        Dim sheet1 As Worksheet = Nothing
        Dim dict As New Dictionary(Of String, Double())

        Dim lastRow As Long
        Dim rawStartRow As Long = 9
        Dim summaryStartRow As Long = 9
        Dim serial As Long = 1
        Dim lastSummaryRow As Long

        Try
            excelApp = CType(Marshal.GetActiveObject("Excel.Application"), Application)
            ws = CType(excelApp.ActiveWorkbook.Sheets("Sheet2"), Worksheet)
            sheet1 = CType(excelApp.ActiveWorkbook.Sheets("Sheet1"), Worksheet)

            ' === Prevent accidental re-summarization ===
            If ws.Range("A2").Value IsNot Nothing AndAlso Trim(CStr(ws.Range("A2").Value)) <> "" Then
                MsgBox("Data Already Summarized." & vbCrLf & "To Summarize Again, Press B1 in Sheet 2", MsgBoxStyle.Exclamation)
                Exit Sub
            End If

            ' === Clear header area ===
            ws.Range("A1:F7").ClearContents()

            ' === Copy header info from Sheet1 ===
            ws.Range("C2").Value = sheet1.Range("E2").Value
            ws.Range("C3").Value = sheet1.Range("E3").Value
            ws.Range("C4").Value = sheet1.Range("E4").Value
            ws.Range("C5").Value = sheet1.Range("E5").Value
            ws.Range("C6").Value = sheet1.Range("AO5").Value

            ' === Header formatting ===
            ws.Range("A8").Value = "SR.NO "
            ws.Range("A1").Value = "MATERIAL SUMMARY LIST"
            ws.Range("A7").Value = ""
            ws.Range("A2:B2").Merge()
            ws.Range("A3:B3").Merge()
            ws.Range("A4:B4").Merge()
            ws.Range("A5:B5").Merge()
            ws.Range("A6:B6").Merge()

            ws.Range("A2").Value = "JOB NAME:-"
            ws.Range("A3").Value = "JOB NUMBER:-"
            ws.Range("A4").Value = "ENGG. NAME:-"
            ws.Range("A5").Value = "PART:-"
            ws.Range("A6").Value = "REVISION:-"
            ws.Range("D6").Value = "DATE:-"
            ws.Range("E6").Value = Format(Now.Date, "dd-MMM-yyyy")
            ws.Range("E4").Value = "TOTAL WT:-"

            ws.Range("A1:F8").Borders.LineStyle = XlLineStyle.xlContinuous
            ws.Range("A1:F1,A7:F7").Merge()
            ws.Range("A1:E7").Font.Size = 11.5
            ws.Range("A1:F8").Font.Bold = True
            ws.Range("A1:F1,A7:F7").Interior.Color = RGB(248, 203, 173)
            ws.Range("A1,A7").HorizontalAlignment = XlHAlign.xlHAlignCenter

            ws.Range("C2:F2").Merge()
            ws.Range("C3:F3").Merge()
            ws.Range("C4:D4").Merge()
            ws.Range("C5:E5").Merge()

            lastRow = ws.Cells(ws.Rows.Count, "B").End(XlDirection.xlUp).Row

            ' === Summarize data ===
            For i As Long = rawStartRow To lastRow
                Dim prop As String = Trim(CStr(ws.Cells(i, "B").Value))
                If prop <> "" Then
                    Dim valC As Double = Val(ws.Cells(i, "C").Value)
                    Dim valD As Double = Val(ws.Cells(i, "D").Value)
                    Dim valE As Double = Val(ws.Cells(i, "E").Value)

                    If Not dict.ContainsKey(prop) Then
                        dict.Add(prop, New Double() {0, 0, 0})
                    End If

                    Dim arr As Double() = dict(prop)
                    arr(0) += valC
                    arr(1) += valD
                    arr(2) += valE
                    dict(prop) = arr
                End If
            Next

            ' === Clear old summary ===
            ws.Range("A" & summaryStartRow & ":G" & ws.Rows.Count).ClearContents()

            ' === Write new summary ===
            Dim rowIndex As Long = summaryStartRow

            For Each key As String In dict.Keys
                Dim arr As Double() = dict(key)
                Dim product As Double = arr(0) * arr(1) / 1000000

                ws.Cells(rowIndex, "A").Value = serial
                ws.Cells(rowIndex, "B").Value = key
                ws.Cells(rowIndex, "E").Value = arr(2)

                If product > 0 Then
                    ws.Cells(rowIndex, "C").Value = "-"
                    ws.Cells(rowIndex, "D").Value = "-"
                Else
                    ws.Cells(rowIndex, "C").Value = arr(0)
                    ws.Cells(rowIndex, "D").Value = "-"
                End If

                ws.Cells(rowIndex, "G").Value = Len(CStr(ws.Cells(rowIndex, "B").Value))

                rowIndex += 1
                serial += 1
            Next

            lastSummaryRow = rowIndex - 1

            ' === Sort by digit-length in Column G ===
            ws.Sort.SortFields.Clear()
            ws.Sort.SortFields.Add(ws.Range("G" & summaryStartRow & ":G" & lastSummaryRow),
                                   XlSortOn.xlSortOnValues,
                                   XlSortOrder.xlAscending)

            ws.Sort.SetRange(ws.Range("A" & summaryStartRow & ":G" & lastSummaryRow))
            ws.Sort.Header = XlYesNoGuess.xlNo
            ws.Sort.Apply()

            ' === Reassign SR.NO after sorting ===
            For i As Long = summaryStartRow To lastSummaryRow
                ws.Cells(i, "A").Value = i - summaryStartRow + 1
            Next

            ' === Remove helper column G ===
            ws.Range("G" & summaryStartRow & ":G" & lastSummaryRow).ClearContents()

            ' === Apply borders ===
            If lastSummaryRow >= summaryStartRow Then
                With ws.Range("A" & summaryStartRow & ":F" & lastSummaryRow)
                    .Borders.LineStyle = XlLineStyle.xlContinuous
                    .HorizontalAlignment = XlHAlign.xlHAlignLeft
                End With
            End If

            ' === Merge Column F ===
            ws.Range("F8:F" & lastSummaryRow).Merge()

            MsgBox("DATA SUMMERIZED.", MsgBoxStyle.Information)

        Catch ex As Exception
            MsgBox("Error: " & ex.Message, MsgBoxStyle.Critical)

        Finally
            dict.Clear()
            dict = Nothing
            ws = Nothing
            sheet1 = Nothing
            excelApp = Nothing
        End Try

    End Sub

End Class
