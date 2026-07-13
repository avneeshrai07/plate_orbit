Imports Microsoft.Office.Interop.Excel
Imports System.IO
Imports System.Runtime.InteropServices

<ComVisible(True)>
<Guid("B3E62114-9A8A-4A61-B71F-3217288DB2E2")>
<ProgId("RaghavStaadExtractor.AssignSectionDatabase")>
Public Class AssignSectionDatabase

    Public Sub CopySectionDatabaseToSheet3()
        Dim excelApp As Application = Nothing
        Dim sourceWB As Workbook = Nothing
        Dim sourceSheet As Worksheet = Nothing
        Dim targetWB As Workbook = Nothing
        Dim targetSheet As Worksheet = Nothing

        Try
            excelApp = Marshal.GetActiveObject("Excel.Application")

            excelApp.ScreenUpdating = False

            targetWB = excelApp.ActiveWorkbook
            targetSheet = CType(targetWB.Sheets(3), Worksheet)
            targetSheet.Cells.Clear()

            Dim folderPath As String = targetWB.Path & "\"
            Dim fileName As String = "RAGHAV DATABASE.xlsm"
            Dim filePath As String = Path.Combine(folderPath, fileName)

            If Not File.Exists(filePath) Then
                MsgBox("File '" & fileName & "' not found in the current folder.", MsgBoxStyle.Exclamation)
                Exit Sub
            End If

            sourceWB = excelApp.Workbooks.Open(filePath, [ReadOnly]:=True)
            sourceSheet = CType(sourceWB.Sheets(1), Worksheet)

            Dim lastRow As Long = sourceSheet.Cells(sourceSheet.Rows.Count, 1).End(XlDirection.xlUp).Row
            Dim lastCol As Long = sourceSheet.Cells(5, sourceSheet.Columns.Count).End(XlDirection.xlToLeft).Column
            Dim dataRange As Range = sourceSheet.Range(sourceSheet.Cells(5, 1), sourceSheet.Cells(lastRow, lastCol))

            Dim destRange As Range = targetSheet.Range("A1").Resize(dataRange.Rows.Count, dataRange.Columns.Count)
            destRange.Value = dataRange.Value
            destRange.Font.Color = RGB(0, 0, 139)

            sourceWB.Close(False)

            MsgBox("Staad Section Database, Perfectly Aligned!", MsgBoxStyle.Information)

        Catch ex As Exception
            MsgBox("Error: " & ex.Message, MsgBoxStyle.Critical)

        Finally
            If sourceWB IsNot Nothing Then Marshal.ReleaseComObject(sourceWB)
            If sourceSheet IsNot Nothing Then Marshal.ReleaseComObject(sourceSheet)
            If targetSheet IsNot Nothing Then Marshal.ReleaseComObject(targetSheet)
            If targetWB IsNot Nothing Then Marshal.ReleaseComObject(targetWB)
            If excelApp IsNot Nothing Then excelApp.ScreenUpdating = True

            sourceSheet = Nothing
            sourceWB = Nothing
            targetSheet = Nothing
            targetWB = Nothing
            excelApp = Nothing

            GC.Collect()
            GC.WaitForPendingFinalizers()
        End Try

    End Sub

End Class
