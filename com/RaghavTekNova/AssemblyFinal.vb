Imports System.Runtime.InteropServices
Imports Microsoft.Office.Interop.Excel
Imports System.IO
Imports Microsoft.VisualBasic   ' ✅ For MsgBox

Namespace RaghavTekNova
    <ComVisible(True)>
    <Guid("C3F9E0A1-ABCD-4A22-9D21-8B92A1234567")>
    <ProgId("RaghavTekNova.AssemblyFinal")>
    Public Class AssemblyFinal

        Public Sub New()
            EnsureLicense()                     ' ✅ Existing licensing
            CheckExpirationDateAndBackdate()   ' 🔐 Expiry + backdate tamper check
        End Sub

        Public Sub A4_Finalized_Assembly_list()
            Dim xlApp As Application = Nothing
            Dim ws As Worksheet = Nothing
            Dim lastRow As Long
            Dim i As Long

            Try
                xlApp = CType(Marshal.GetActiveObject("Excel.Application"), Application)

                ' ✅ Always Sheet1
                ws = CType(xlApp.ActiveWorkbook.Sheets("Sheet1"), Worksheet)

                ' Unlock all sheets
                UnlockAllSheets(xlApp)

                ' Find last used row in column L
                lastRow = ws.Cells(ws.Rows.Count, "L").End(XlDirection.xlUp).Row

                ' Loop through rows 8 to lastRow - 1
                For i = 8 To lastRow - 1
                    Dim val As Object = ws.Cells(i, "L").Value
                    If val IsNot Nothing AndAlso IsNumeric(val) AndAlso CDbl(val) > 1 Then
                        ws.Cells(i, "D").ClearContents()
                    End If
                Next

                ' Hide columns
                ws.Columns("L").Hidden = True
                ws.Columns("M").Hidden = True

                ' ✅ Ensure Sheet1 stays active
                ws.Activate()

                ' ✅ Use VB.NET MsgBox (no error 1004)
                MsgBox("All Assemblies Profile Removed!", vbInformation, "Operation Complete")

                ' Lock sheets
                LockAllSheets(xlApp)

                ' ✅ Keep Sheet1 active
                ws.Activate()

            Catch ex As Exception
                MsgBox("Error: " & ex.Message, vbCritical, "Assembly Final")
            End Try
        End Sub
        Private Sub EnsureLicense()
            Dim validator As New ExecutionValidation()
            If Not validator.IsLicenceValid() Then
                Throw New UnauthorizedAccessException("License check failed.")
            End If
        End Sub

        Private Sub CheckExpirationDateAndBackdate()
            Dim expiryDate As New DateTime(2027, 4, 9)
            Dim today As DateTime = DateTime.Now.Date

            ' 🔒 Path to store last run date (user-specific and hidden)
            Dim filePath As String = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "RaghavStaad.last")

            ' Check if file exists
            If File.Exists(filePath) Then
                Dim lastRunText As String = File.ReadAllText(filePath).Trim()
                Dim lastRunDate As DateTime
                If DateTime.TryParse(lastRunText, lastRunDate) Then

                    ' ⛔ Backdate tampering detected
                    If today < lastRunDate Then
                        Throw New UnauthorizedAccessException("System date has been tampered. Please correct your clock.")
                    End If
                End If
            End If

            ' Save today's date as last run (overwrite or create)
            File.WriteAllText(filePath, today.ToString("yyyy-MM-dd"))

            ' 🔐 Expiry check
            If today > expiryDate Then
                Throw New UnauthorizedAccessException("Software expired. Please contact support for an updated version.")
            End If
        End Sub

        Private Sub UnlockAllSheets(xlApp As Application)
            Dim calcMode As XlCalculation = xlApp.Calculation
            xlApp.ScreenUpdating = False
            xlApp.EnableEvents = False
            xlApp.Calculation = XlCalculation.xlCalculationManual

            For Each ws As Worksheet In xlApp.ActiveWorkbook.Worksheets
                ws.Unprotect(Password:="2022")
            Next

            xlApp.Calculation = calcMode
            xlApp.EnableEvents = True
            xlApp.ScreenUpdating = True
        End Sub

        Private Sub LockAllSheets(xlApp As Application)
            Dim calcMode As XlCalculation = xlApp.Calculation
            xlApp.ScreenUpdating = False
            xlApp.EnableEvents = False
            xlApp.Calculation = XlCalculation.xlCalculationManual

            For Each ws As Worksheet In xlApp.ActiveWorkbook.Worksheets
                ws.Protect(Password:="2022", UserInterfaceOnly:=True)
            Next

            xlApp.Calculation = calcMode
            xlApp.EnableEvents = True
            xlApp.ScreenUpdating = True
        End Sub

    End Class
End Namespace