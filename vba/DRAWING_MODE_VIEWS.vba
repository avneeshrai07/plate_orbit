'============================================
'=========DIFFRENT DRAWING MODES=============
'============================================
Sub FrontAndRun()
    With ThisWorkbook.Sheets("Sheet1")
        .Range("AR2").Value = "FRONT"
        .Range("AS2").Value = "1"
    End With

    ' Call the macro named BeamDataFromStaad
    Call ByPassBeamDataFromStaad
    ' Call the macro named ExportToDXF
    Call ExportToDXF
End Sub

Sub BackAndRun()
    With ThisWorkbook.Sheets("Sheet1")
        .Range("AR2").Value = "BACK"
        .Range("AS2").Value = "2"
    End With

    ' Call the macro named BeamDataFromStaad
    Call ByPassBeamDataFromStaad
    ' Call the macro named ExportToDXF
    Call ExportToDXF
End Sub

Sub LeftAndRun()
     With ThisWorkbook.Sheets("Sheet1")
        .Range("AR2").Value = "LEFT"
        .Range("AS2").Value = "3"
    End With

    ' Call the macro named BeamDataFromStaad
    Call ByPassBeamDataFromStaad
    ' Call the macro named ExportToDXF
    Call ExportToDXF
End Sub

Sub RightAndRun()
       With ThisWorkbook.Sheets("Sheet1")
        .Range("AR2").Value = "RIGHT"
        .Range("AS2").Value = "4"
    End With

    ' Call the macro named BeamDataFromStaad
    Call ByPassBeamDataFromStaad
    ' Call the macro named ExportToDXF
    Call ExportToDXF
End Sub

Sub TopAndRun()
    With ThisWorkbook.Sheets("Sheet1")
        .Range("AR2").Value = "TOP PLAN"
        .Range("AS2").Value = "5"
    End With

    ' Call the macro named BeamDataFromStaad
    Call ByPassBeamDataFromStaad
    ' Call the macro named ExportToDXF
    Call ExportToDXF
End Sub
'============================================
