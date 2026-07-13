Sub OpenMyLinkedInProfile()
    Dim linkedInURL As String
    
    ' Your LinkedIn profile URL below
    linkedInURL = "https://www.linkedin.com/in/peeyushraghav15/"
    
    ' Opens the URL in the default web browser
    ThisWorkbook.FollowHyperlink Address:=linkedInURL, NewWindow:=True
End Sub

