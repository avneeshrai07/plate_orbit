'=====================================================
'=====THIS CODE IS USED FOR CREATE DXF FILES =========
'=====================================================
Sub ExportToDXF()
    On Error Resume Next
    Dim exporter As Object
    Set exporter = CreateObject("RaghavStaadExtractor.DxfExporterRaghav")
    exporter.ExportBeamsToDxfWithPrompt ActiveSheet
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description, vbCritical
End Sub
'=====================================================
