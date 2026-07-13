Imports System.Runtime.InteropServices
Imports Microsoft.Office.Interop.Excel

#Region "COM Interfaces"
' Main COM Interface
<ComVisible(True)>
<Guid("12345678-1234-1234-1234-123456789013")>
<InterfaceType(ComInterfaceType.InterfaceIsDual)>
Public Interface IGuidManager
    <DispId(1)>
    Sub DisplayGuid()

    <DispId(2)>
    Sub RemoveGuidDisplay()
End Interface

' COM Events Interface
<ComVisible(True)>
<Guid("12345678-1234-1234-1234-123456789014")>
<InterfaceType(ComInterfaceType.InterfaceIsIDispatch)>
Public Interface IGuidManagerEvents
    ' Event definitions can be added here if needed
End Interface
#End Region

#Region "Main COM Class"
' Main COM Class containing both methods
<ComVisible(True)>
<Guid("12345678-1234-1234-1234-123456789012")>
<ProgId("RaghavTekNova.GuidManager")>
<ClassInterface(ClassInterfaceType.None)>
<ComSourceInterfaces(GetType(IGuidManagerEvents))>
Public Class GuidManager
    Implements IGuidManager

    Public Sub New()
        MyBase.New()
    End Sub

    ''' <summary>
    ''' Display GUID - equivalent to A10_GET_GUID()
    ''' </summary>
    Public Sub DisplayGuid() Implements IGuidManager.DisplayGuid
        Dim xlApp As Application = Nothing
        Dim ws As Worksheet = Nothing
        Dim password As String = "2022"
        Dim lastRow As Long

        Try
            ' Get the active Excel application
            xlApp = Marshal.GetActiveObject("Excel.Application")
            ws = xlApp.ActiveWorkbook.Worksheets("Sheet1")

            ' Turn off screen updating for better performance
            xlApp.ScreenUpdating = False

            ' Unprotect the sheet
            ws.Unprotect(password)

            ' Unlock only Column O
            ws.Columns("O:O").Locked = False

            ' Unhide Column O
            ws.Columns("O:O").EntireColumn.Hidden = False

            ' Merge top 6 rows in Column O and set the background color to white
            ws.Range("O1:O6").Merge()
            ws.Range("O1").Interior.Color = RGB(255, 255, 255)

            ' Find the last used row in column A
            lastRow = ws.Cells(ws.Rows.Count, "A").End(XlDirection.xlUp).Row

            ' Hide Columns F to N starting from row 7 onward
            ws.Range("F7:N" & lastRow.ToString()).EntireColumn.Hidden = True

            ' Format cell O7 - Bold, Centered, with Background Color
            With ws.Range("O7")
                .Font.Bold = True
                .HorizontalAlignment = XlHAlign.xlHAlignCenter
                .VerticalAlignment = XlVAlign.xlVAlignCenter
                .Interior.Color = RGB(237, 138, 111) ' Yellow background - you can change this color
            End With

            ' Protect the sheet again, keeping unlocked cells accessible
            ws.Protect(Password:=password, UserInterfaceOnly:=True)

        Catch ex As Exception
            Throw New COMException("Error in DisplayGuid: " & ex.Message, ex)
        Finally
            ' Turn screen updating back on
            If xlApp IsNot Nothing Then
                xlApp.ScreenUpdating = True
            End If

            ' Clean up COM objects
            If ws IsNot Nothing Then Marshal.ReleaseComObject(ws)
            If xlApp IsNot Nothing Then Marshal.ReleaseComObject(xlApp)
        End Try
    End Sub

    ''' <summary>
    ''' Remove GUID display - equivalent to A11_GET_GUID_ANTI()
    ''' </summary>
    Public Sub RemoveGuidDisplay() Implements IGuidManager.RemoveGuidDisplay
        Dim xlApp As Application = Nothing
        Dim ws As Worksheet = Nothing
        Dim lastRow As Long

        Try
            ' Get the active Excel application
            xlApp = Marshal.GetActiveObject("Excel.Application")
            ws = xlApp.ActiveWorkbook.Worksheets("Sheet1")

            ' Turn off screen updating for better performance
            xlApp.ScreenUpdating = False

            ' Unmerge the top 6 rows in Column O
            ws.Range("O1:O6").UnMerge()
            ws.Range("O1:O6").Interior.ColorIndex = XlColorIndex.xlColorIndexNone

            ' Find the last used row in column A
            lastRow = ws.Cells(ws.Rows.Count, "A").End(XlDirection.xlUp).Row

            ' Unhide Columns F to N from row 7 onward
            ws.Range("F7:N" & lastRow.ToString()).EntireColumn.Hidden = False

            ' Keep Columns L to O hidden
            ws.Columns("L:O").EntireColumn.Hidden = True

            ' Show message
            xlApp.MessageBox("Assembly GUID Removed", Title:="Information")

        Catch ex As Exception
            Throw New COMException("Error in RemoveGuidDisplay: " & ex.Message, ex)
        Finally
            ' Turn screen updating back on
            If xlApp IsNot Nothing Then
                xlApp.ScreenUpdating = True
            End If

            ' Clean up COM objects
            If ws IsNot Nothing Then Marshal.ReleaseComObject(ws)
            If xlApp IsNot Nothing Then Marshal.ReleaseComObject(xlApp)
        End Try
    End Sub

    ''' <summary>
    ''' Helper function to convert RGB values (equivalent to VBA RGB function)
    ''' </summary>
    Private Function RGB(red As Integer, green As Integer, blue As Integer) As Integer
        Return red + (green * 256) + (blue * 65536)
    End Function

End Class
#End Region

#Region "Separate Public Classes for Easy Calling"
' Separate class for Display functionality - Easy to call
<ComVisible(True)>
<ProgId("RaghavTekNova.DisplayGuid")>
Public Class DisplayGuid
    Private guidManager As GuidManager

    Public Sub New()
        guidManager = New GuidManager()
    End Sub

    ''' <summary>
    ''' Execute Display GUID operation
    ''' </summary>
    Public Sub Execute()
        guidManager.DisplayGuid()
    End Sub

    Protected Overrides Sub Finalize()
        If guidManager IsNot Nothing Then
            Marshal.ReleaseComObject(guidManager)
            guidManager = Nothing
        End If
        MyBase.Finalize()
    End Sub
End Class

' Separate class for Remove functionality - Easy to call
<ComVisible(True)>
<ProgId("RaghavTekNova.RemoveGuid")>
Public Class RemoveGuid
    Private guidManager As GuidManager

    Public Sub New()
        guidManager = New GuidManager()
    End Sub

    ''' <summary>
    ''' Execute Remove GUID operation
    ''' </summary>
    Public Sub Execute()
        guidManager.RemoveGuidDisplay()
    End Sub

    Protected Overrides Sub Finalize()
        If guidManager IsNot Nothing Then
            Marshal.ReleaseComObject(guidManager)
            guidManager = Nothing
        End If
        MyBase.Finalize()
    End Sub
End Class
#End Region