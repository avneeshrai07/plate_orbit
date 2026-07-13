'===========================================================================
'=====THIS CODE IS USED FOR CREATE DXF FILES BY BYPASSING BOQ MODE =========
'===========================================================================
Sub ByPassBeamDataFromStaad()
    On Error Resume Next
    
    ' Create the COM object
    Dim extractor As Object
    Set extractor = CreateObject("RaghavStaadExtractor.StaadDataExtractor")
    
    If extractor Is Nothing Then
        MsgBox "Failed to create STAAD extractor object. Check if DLL is registered.", vbCritical
        Exit Sub
    End If
    
    ' Call the extraction method
    Dim result As String
    result = extractor.ExtractBeamDataFromStaad()
    
    'Call DrawingMode
    Call DrawingMode
    ' Show result
    'MsgBox result, vbInformation, "STAAD Extraction"
    
    ' Clean up
    Set extractor = Nothing
End Sub
'===========================================================================
Sub ByPassForAnalysis()
    Dim checkVal As String
    
    ' === Validation for AW2 ===
    checkVal = UCase(Trim(Range("AW2").Value))
    If checkVal <> "G" And checkVal <> "X" And checkVal <> "Z" Then
        MsgBox "Press Z,X or GRID button for Analysis Or Define Grids", vbExclamation, "Validation Failed"
        Exit Sub
    End If
    
    On Error Resume Next
    
    ' Create the COM object
    Dim extractor As Object
    Set extractor = CreateObject("RaghavStaadExtractor.StaadDataExtractor")
    
    If extractor Is Nothing Then
        MsgBox "Failed to create STAAD extractor object. Check if DLL is registered.", vbCritical
        Exit Sub
    End If
    
    ' Call the extraction method
    Dim result As String
    Dim analysisResult As String
    result = extractor.ExtractBeamDataFromStaad()
    
    'Call DrawingMode
    Call DrawingMode
    ' Show result
'MsgBox result, vbInformation, "STAAD Extraction_Extraction"
' Show analysis results if available
If extractor.HasFrameAnalysisResult() Then
    analysisResult = extractor.GetFrameAnalysisResult()
    MsgBox analysisResult, vbInformation, "STAAD Extraction_Analysis"
End If
    
    ' Clean up
    Set extractor = Nothing
End Sub

