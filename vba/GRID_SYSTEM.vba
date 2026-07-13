Sub CallGridSystem()
    Dim obj As Object
    Set obj = CreateObject("RaghavStaadExtractor.GridSystem")
    obj.ShowGridInputForm
End Sub
Sub WriteX()
    Sheets("Sheet1").Range("AW2").Value = "X"
End Sub

Sub WriteZ()
    Sheets("Sheet1").Range("AW2").Value = "Z"
End Sub
Sub WriteG()
    Sheets("Sheet1").Range("AW2").Value = "G"
End Sub
