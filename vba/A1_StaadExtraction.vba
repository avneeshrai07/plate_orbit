'==================================================================
'=====THIS CODE IS USED FOR EXTRACT STAAD DATA FROM STAAD =========
'==================================================================
Sub ExtractStaadDataR1()
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
    
    'Call section shorting for sheet 2
    Call SectionSortingS2
    'Call BoqMode
    Call BoqMode
       
    ' Show result
    'MsgBox result, vbInformation, "STAAD Extraction"
' Always show extraction results
MsgBox result, vbInformation, "STAAD Extraction_Extraction"

' Show analysis results if available
If extractor.HasFrameAnalysisResult() Then
    analysisResult = extractor.GetFrameAnalysisResult()
    MsgBox analysisResult, vbInformation, "STAAD Extraction_Analysis"
End If
    
    ' Clean up
    Set extractor = Nothing
End Sub
'==================================================================
