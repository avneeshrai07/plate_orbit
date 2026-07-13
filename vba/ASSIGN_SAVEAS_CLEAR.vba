'===========================================================
'=====THIS CODE IS USED FOR ASSIGN RAGHAV DATA BASE=========
'===========================================================
Sub AssignSectionDatabase()
    CreateObject("RaghavStaadExtractor.AssignSectionDatabase").CopySectionDatabaseToSheet3
End Sub

'=======================================================================
'=====THIS CODE IS USED FOR SAVE AS BEAM OUTPUT & PLATE SUMMARY=========
'=======================================================================

Sub SaveBeamOutput()
    CreateObject("RaghavStaadExtractor.SaveSheetsManager").Save_Sheet1_BeamOutput
End Sub

Sub SavePlateSummary()
    CreateObject("RaghavStaadExtractor.SaveSheetsManager").Save_Sheet2_PlateSummary
End Sub
'=================================================
'=====THIS CODE IS USED FOR CLEAR SHEETS =========
'=================================================
Sub Clear_Sheet1()
    Dim obj As Object
    Set obj = CreateObject("RaghavStaadExtractor.ClearSheet")
    obj.ClearAndResetSheet1
    Set obj = Nothing
End Sub

Sub Clear_Sheet2()
    Dim obj As Object
    Set obj = CreateObject("RaghavStaadExtractor.ClearSheet")
    obj.ClearAndResetSheet2
    Set obj = Nothing
End Sub

Sub Clear_AllSheets()
    Dim obj As Object
    Set obj = CreateObject("RaghavStaadExtractor.ClearSheet")
    obj.ClearAllSheets
    Set obj = Nothing
End Sub
'=======================================================================
