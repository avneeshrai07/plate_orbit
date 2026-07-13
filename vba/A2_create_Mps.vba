'====================================================
'=====THIS CODE IS USED FOR SECTION SORTING =========
'====================================================
Sub SectionSortingS2()

    Dim sorter As Object
    On Error GoTo ErrHandler
    ' Instantiate using ProgID
    Set sorter = CreateObject("RaghavStaadExtractor.SectionSortingS2")
    ' Call method
    sorter.SummerizedToSheet2
    'MsgBox "Sorting complete", vbInformation
    Exit Sub

ErrHandler:
    MsgBox "Error: " & Err.Description, vbCritical

End Sub

'=========================================================
'=====THIS CODE IS USED FOR SECTION FINALIZATION =========
'=========================================================
Sub SectionFinalization()
    On Error GoTo ErrHandler
    Dim obj As Object

    ' Create instance of the COM class
    Set obj = CreateObject("RaghavStaadExtractor.SectionFinalization")

    ' Call the method from the COM class
    obj.FinalSummarySheet2
    ' ? Call the local macro to switch sheet
    'Call SwitchToSheet2
    ' Cleanup
    Set obj = Nothing
    Exit Sub

ErrHandler:
    MsgBox "Error while calling class: " & Err.Description, vbCritical
End Sub
'=========================================================
