Imports Microsoft.Office.Interop.Excel
Imports System.Runtime.InteropServices

<ComVisible(True)>
<Guid("B1234567-89CD-4ABC-8123-DEF456789ABC")>
<ProgId("RaghavStaadExtractor.ColumnHider")>
Public Class ColumnHider

    Private Const PASSWORD As String = "2022"
    Private Const SHEET_NAME As String = "Sheet1"

    ' Shared method to apply hiding/unhiding logic
    Private Sub ApplyVisibility(columnStates As Dictionary(Of String, Boolean))
        Dim excelApp As Application = Nothing
        Dim workbook As Workbook = Nothing
        Dim sheet As Worksheet = Nothing

        Try
            excelApp = Marshal.GetActiveObject("Excel.Application")
            workbook = excelApp.ActiveWorkbook
            sheet = CType(workbook.Sheets(SHEET_NAME), Worksheet)

            sheet.Unprotect(PASSWORD)

            ' Apply visibility states
            For Each pair In columnStates
                sheet.Range(pair.Key).EntireColumn.Hidden = pair.Value
            Next

            ' Always hide F:AJ regardless of operation
            sheet.Range("F:AJ").EntireColumn.Hidden = True
            sheet.Range("AY:BC").EntireColumn.Hidden = True

            sheet.Protect(PASSWORD, DrawingObjects:=True, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True)

        Catch ex As Exception
            MsgBox("Error modifying column visibility: " & ex.Message)
        Finally
            If sheet IsNot Nothing Then Marshal.ReleaseComObject(sheet)
            If workbook IsNot Nothing Then Marshal.ReleaseComObject(workbook)
            If excelApp IsNot Nothing Then Marshal.ReleaseComObject(excelApp)
            GC.Collect()
            GC.WaitForPendingFinalizers()
        End Try
    End Sub

    ' ✅ Operation 1: Hide AR:AW, show A:AQ (F:AJ always hidden)
    Public Sub ShowAAtoAQ_HideARtoAW()
        Dim columnsToSet As New Dictionary(Of String, Boolean) From {
            {"A:AQ", False},
            {"AR:AW", True}
        }
        ApplyVisibility(columnsToSet)
    End Sub

    ' ✅ Operation 2: Hide A:AQ, show AR:AW (F:AJ always hidden)
    Public Sub HideAAtoAQ_ShowARtoAW()
        Dim columnsToSet As New Dictionary(Of String, Boolean) From {
            {"A:AQ", True},
            {"AR:AW", False}
        }
        ApplyVisibility(columnsToSet)
    End Sub

    ' ✅ Operation 3: Unhide all columns except F:AJ
    Public Sub NormalViewExceptFAJ()
        Dim excelApp As Application = Nothing
        Dim workbook As Workbook = Nothing
        Dim sheet As Worksheet = Nothing

        Try
            excelApp = Marshal.GetActiveObject("Excel.Application")
            workbook = excelApp.ActiveWorkbook
            sheet = CType(workbook.Sheets(SHEET_NAME), Worksheet)

            sheet.Unprotect(PASSWORD)

            ' Unhide all columns first
            sheet.Cells.EntireColumn.Hidden = False

            ' Then hide only F:AJ
            sheet.Range("F:AJ").EntireColumn.Hidden = True
            sheet.Range("AY:BC").EntireColumn.Hidden = True

            sheet.Protect(PASSWORD, DrawingObjects:=True, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True)

        Catch ex As Exception
            MsgBox("Error restoring normal view: " & ex.Message)
        Finally
            If sheet IsNot Nothing Then Marshal.ReleaseComObject(sheet)
            If workbook IsNot Nothing Then Marshal.ReleaseComObject(workbook)
            If excelApp IsNot Nothing Then Marshal.ReleaseComObject(excelApp)
            GC.Collect()
            GC.WaitForPendingFinalizers()
        End Try
    End Sub


End Class
