'==========================================================================
'=====THIS CODE IS USED FOR SWITCHING BETWEEN DIFFRENT SHEET MODES=========
'==========================================================================

Sub BoqMode()
    CreateObject("RaghavStaadExtractor.ColumnHider").ShowAAtoAQ_HideARtoAW
End Sub

'==========================================================================
'==========================================================================
Sub DrawingMode()
    CreateObject("RaghavStaadExtractor.ColumnHider").HideAAtoAQ_ShowARtoAW
End Sub
'==========================================================================
'==========================================================================
Sub NormalMode()
    CreateObject("RaghavStaadExtractor.ColumnHider").NormalViewExceptFAJ
End Sub

'==========================================================================
'==========THIS CODE IS USED FOR UNLOCK JOB INFORMATION ENTRY==============
'==========================================================================
Sub UnlockRange()
    Dim unlocker As Object
    Set unlocker = CreateObject("RaghavStaadExtractor.Unlock")
    unlocker.UnlockSheetRange
End Sub
'==========================================================================
