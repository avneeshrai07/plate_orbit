Sub ConvertWeightUnits()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim currentUnit As String
    Dim isToTon As Boolean

    ' Set worksheet
    Set ws = ThisWorkbook.Sheets("Sheet2")
    
    ' Get unit in E8
    currentUnit = Trim(UCase(ws.Range("E8").Value))
    
    ' Determine direction
    If currentUnit = "WEIGHT(KG)" Then
        isToTon = True
        ws.Range("E8").Value = "WEIGHT(Ton)"
        ws.Range("C8").Value = "LENGTH (M)"
        ws.Range("D8").Value = "WIDTH (M)"
    ElseIf currentUnit = "WEIGHT(TON)" Then
        isToTon = False
        ws.Range("E8").Value = "WEIGHT(Kg)"
        ws.Range("C8").Value = "LENGTH (mm)"
        ws.Range("D8").Value = "WIDTH (mm)"
    Else
        MsgBox "E8 must be WEIGHT(Kg) or WEIGHT(Ton)", vbExclamation
        Exit Sub
    End If

    ' Find last row (based on col C or E)
    lastRow = ws.Cells(ws.Rows.count, 3).End(xlUp).Row
    If ws.Cells(ws.Rows.count, 5).End(xlUp).Row > lastRow Then
        lastRow = ws.Cells(ws.Rows.count, 5).End(xlUp).Row
    End If
    If lastRow < 9 Then lastRow = 9

    ' Loop and process
    For i = 9 To lastRow
        ' Weight conversion in Column E
        If IsNumeric(ws.Cells(i, 5).Value) And Not IsEmpty(ws.Cells(i, 5).Value) Then
            If isToTon Then
                ' Store original KG in col Z
                ws.Cells(i, 26).Value = ws.Cells(i, 5).Value
                ' Show 3-decimal Ton
                ws.Cells(i, 5).Value = Round(ws.Cells(i, 5).Value / 1000, 3)
            Else
                ' Restore full KG from col Z if available
                If IsNumeric(ws.Cells(i, 26).Value) Then
                    ws.Cells(i, 5).Value = ws.Cells(i, 26).Value
                    ws.Cells(i, 26).ClearContents
                Else
                    ' fallback: multiply existing value if no backup
                    ws.Cells(i, 5).Value = ws.Cells(i, 5).Value * 1000
                End If
            End If
        End If

        ' Length conversion in Column C
        If IsNumeric(ws.Cells(i, 3).Value) And Not IsEmpty(ws.Cells(i, 3).Value) Then
            If isToTon Then
                ws.Cells(i, 3).Value = Round(ws.Cells(i, 3).Value / 1000, 3)
            Else
                ws.Cells(i, 3).Value = ws.Cells(i, 3).Value * 1000
            End If
        End If
    Next i

    'MsgBox "Conversion completed!", vbInformation
End Sub
