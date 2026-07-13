Imports System
Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.Drawing
Imports System.IO
Imports System.Linq
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Windows.Forms
Imports Microsoft.VisualBasic

<ComVisible(True)>
<Guid("F9396D1F-8F3F-4A8B-9E5D-3B7D8C1A4B2E")>
<ProgId("RaghavStaadExtractor.StaadDataExtractor")>
Public Class StaadDataExtractor

    Public Structure BeamData
        Public BeamNo As Integer
        Public Length As Double
        Public SectionName As String
        Public NodeA As Integer
        Public NodeB As Integer
        Public XA As Double
        Public YA As Double
        Public ZA As Double
        Public XB As Double
        Public YB As Double
        Public ZB As Double
        Public F1 As Double
        Public F2 As Double
        Public F3 As Double
        Public F4 As Double
        Public F5 As Double
        Public F6 As Double
        Public F7 As Double
        Public Beta As Double
    End Structure

    Private Structure Sheet4GridData
        Public GridNames As List(Of String)
        Public AllCoordinates As List(Of Double)
        Public SpacingString As String
    End Structure

    Private Const PROP_DEPTH_START As Integer = 0
    Private Const PROP_WEB_THICKNESS As Integer = 1
    Private Const PROP_DEPTH_END As Integer = 2
    Private Const PROP_FLANGE_WIDTH As Integer = 3
    Private Const PROP_FLANGE_THICKNESS As Integer = 4
    Private Const PROP_BOTTOM_FLANGE_WIDTH As Integer = 5
    Private Const PROP_BOTTOM_FLANGE_THICKNESS As Integer = 6

    Private _lastFrameAnalysisResult As String
    Private _lastLogFilePath As String

    Public Sub New()
        _lastFrameAnalysisResult = ""
        _lastLogFilePath = ""
        EnsureLicense()
        CheckExpirationDateAndBackdate()
    End Sub

    <ComVisible(True)>
    Public Function ExtractBeamDataFromStaad() As String
        Dim excelApp As Object = Nothing
        Dim staadApp As Object = Nothing
        Dim result As String
        Try
            excelApp = GetOrCreateExcelApp()
            staadApp = GetStaadApplication()
            If staadApp Is Nothing Then
                result = "STAAD is not running or OpenSTAAD not available."
            Else
                Dim wb As Object = excelApp.ActiveWorkbook
                If wb Is Nothing Then
                    If excelApp.Workbooks.Count > 0 Then
                        wb = excelApp.Workbooks(1)
                    Else
                        wb = excelApp.Workbooks.Add()
                    End If
                End If
                Dim ws As Object = GetOrCreateWorksheet(wb, "Sheet1")
                excelApp.ScreenUpdating = False
                excelApp.Calculation = -4135 ' xlCalculationManual
                excelApp.EnableEvents = False
                excelApp.DisplayAlerts = False
                UnprotectSheetAndWorkbook(ws)
                ShowHiddenRows(ws)
                PreserveAndClearWorksheet(ws)
                WriteJobInfoToExcel(ws, staadApp)
                SetupExcelHeaders(ws)
                Dim text As String = ProcessBeamData(ws, staadApp)
                FormatExcelWorksheet(ws)
                ProcessSectionFarmingAndWeight(ws)
                CalculateTotalWeight(ws)
                _lastFrameAnalysisResult = ProcessFrameAnalysis(ws, staadApp)
                HideColumns(ws)
                ProtectSheetAndWorkbook(ws)
                excelApp.ScreenUpdating = True
                excelApp.Calculation = -4105 ' xlCalculationAutomatic
                excelApp.EnableEvents = True
                excelApp.DisplayAlerts = True
                wb.Sheets("Sheet1").Select()
                Dim activeWindow As Object = excelApp.ActiveWindow
                activeWindow.DisplayGridlines = False
                activeWindow.DisplayHeadings = False
                excelApp.DisplayFormulaBar = True
                activeWindow.DisplayWorkbookTabs = False
                excelApp.ExecuteExcel4Macro("SHOW.TOOLBAR(""Ribbon"", False)")
                activeWindow.ScrollRow = 1
                activeWindow.ScrollColumn = 1
                activeWindow.Zoom = 100
                Dim windowCount As Integer = CInt(wb.Windows.Count)
                For i As Integer = 1 To windowCount
                    Dim win As Object = wb.Windows(i)
                    If win IsNot Nothing Then
                        win.DisplayGridlines = False
                    End If
                Next
                AutoSaveWorkbook(wb)
                ws.Range("AW2").Value = ""
                result = text
            End If
        Catch ex As Exception
            result = "Error: " & ex.Message
        Finally
            If staadApp IsNot Nothing Then Marshal.ReleaseComObject(staadApp)
            If excelApp IsNot Nothing Then Marshal.ReleaseComObject(excelApp)
        End Try
        Return result
    End Function

    <ComVisible(True)>
    Public Function GetFrameAnalysisResult() As String
        Return _lastFrameAnalysisResult
    End Function

    <ComVisible(True)>
    Public Function GetCombinedResults() As String
        Dim text As String = ExtractBeamDataFromStaad()
        If String.IsNullOrEmpty(_lastFrameAnalysisResult) Then
            Return text
        End If
        Return text & vbCrLf & vbCrLf & "=== FRAME ANALYSIS RESULTS ===" & vbCrLf & _lastFrameAnalysisResult
    End Function

    <ComVisible(True)>
    Public Function HasFrameAnalysisResult() As Boolean
        Return Not String.IsNullOrEmpty(_lastFrameAnalysisResult)
    End Function

    Private Sub ShowHiddenRows(ws As Object)
        Try
            ws.Rows.Hidden = False
        Catch ex As Exception
            Debug.WriteLine("ClearAllFilters failed: " & ex.Message)
        End Try
    End Sub

    <ComVisible(True)>
    Public Sub ShowFrameAnalysisDialog()
        If Not String.IsNullOrEmpty(_lastFrameAnalysisResult) Then
            ShowCustomAnalysisDialog(_lastFrameAnalysisResult, _lastLogFilePath)
        End If
    End Sub

    <ComVisible(True)>
    Public Function HasLogFile() As Boolean
        Return Not String.IsNullOrEmpty(_lastLogFilePath) AndAlso File.Exists(_lastLogFilePath)
    End Function

    Private Sub ShowCustomAnalysisDialog(message As String, logFilePath As String)
        Try
            Dim form As New Form()
            form.Text = "Nova Frame Analysis Results"
            form.Size = New Size(600, 450)
            form.StartPosition = FormStartPosition.CenterScreen
            form.FormBorderStyle = FormBorderStyle.FixedDialog
            form.MaximizeBox = False
            form.MinimizeBox = False
            Dim textBox As New TextBox()
            textBox.Multiline = True
            textBox.ScrollBars = ScrollBars.Vertical
            textBox.ReadOnly = True
            textBox.Text = message
            textBox.Font = New Font("Consolas", 9.0F)
            textBox.Location = New Point(10, 10)
            textBox.Size = New Size(565, 350)
            Dim panel As New Panel()
            panel.Location = New Point(10, 370)
            panel.Size = New Size(565, 40)
            Dim button As New Button()
            button.Text = "OK"
            button.Size = New Size(80, 30)
            button.Location = New Point(475, 5)
            button.DialogResult = DialogResult.OK
            If Not String.IsNullOrEmpty(logFilePath) AndAlso File.Exists(logFilePath) Then
                Dim button2 As New Button()
                button2.Text = "Open Log File"
                button2.Size = New Size(120, 30)
                button2.Location = New Point(340, 5)
                AddHandler button2.Click, Sub(sender As Object, e As EventArgs)
                                               Try
                                                   Process.Start("explorer.exe", "/select,""" & logFilePath & """")
                                               Catch
                                                   Try
                                                       Process.Start("explorer.exe", Path.GetDirectoryName(logFilePath))
                                                   Catch
                                                       Interaction.MsgBox("Could not open log file location.", MsgBoxStyle.Exclamation)
                                                   End Try
                                               End Try
                                           End Sub
                panel.Controls.Add(button2)
            End If
            panel.Controls.Add(button)
            form.Controls.Add(textBox)
            form.Controls.Add(panel)
            form.AcceptButton = button
            form.ShowDialog()
            form.Dispose()
        Catch
            Interaction.MsgBox(message, MsgBoxStyle.Information, "Nova Frame Analysis Results")
        End Try
    End Sub

    Private Function ProcessFrameAnalysis(ws As Object, staadApp As Object) As String
        Dim result As String
        Try
            _lastLogFilePath = ""
            Dim text As String = ""
            If ws.Range("AW2").Value IsNot Nothing Then
                text = ws.Range("AW2").Value.ToString().Trim().ToUpper()
            End If
            If String.IsNullOrEmpty(text) Then
                result = ""
            ElseIf text <> "X" AndAlso text <> "Z" AndAlso text <> "G" Then
                result = $"Invalid prompt '{text}' in AW2. Please enter 'X', 'Z', or 'G' only."
            ElseIf text = "G" Then
                result = ProcessBothDirectionsForGridSilent(ws, staadApp)
            ElseIf Not AreGridsDefined(text, ws) Then
                result = "FOR FRAME ANALYSIS DEFINE GRIDS FIRST USING 'G' OPTION" & vbCrLf & vbCrLf &
                    $"Grid data for '{text}' direction is not found Nova Grid System." & vbCrLf & vbCrLf &
                    "Please follow these steps:" & vbCrLf &
                    "1. Enter 'G' in cell AW2" & vbCrLf &
                    "2. Run the analysis to generate grid system" & vbCrLf &
                    $"3. Then use '{text}' for frame analysis"
            Else
                Dim beamList As New List(Of BeamData)()
                Dim lastRow As Integer = CInt(ws.Cells(ws.Rows.Count, 1).End(-4162).Row)
                If lastRow < 9 Then
                    result = "No beam data found for frame analysis."
                Else
                    For i As Integer = 9 To lastRow
                        If ws.Cells(i, 2).Value Is Nothing Then Continue For
                        Dim item As BeamData = Nothing
                        item.BeamNo = CInt(Math.Round(Val(ws.Cells(i, 2).Value)))
                        item.Length = Val(ws.Cells(i, 3).Value)
                        If ws.Cells(i, 4).Value Is Nothing Then
                            item.SectionName = ""
                        Else
                            item.SectionName = ws.Cells(i, 4).Value.ToString()
                        End If
                        item.NodeA = CInt(Math.Round(Val(ws.Cells(i, 14).Value)))
                        item.NodeB = CInt(Math.Round(Val(ws.Cells(i, 15).Value)))
                        item.XA = Val(ws.Cells(i, 16).Value)
                        item.YA = Val(ws.Cells(i, 17).Value)
                        item.ZA = Val(ws.Cells(i, 18).Value)
                        item.XB = Val(ws.Cells(i, 19).Value)
                        item.YB = Val(ws.Cells(i, 20).Value)
                        item.ZB = Val(ws.Cells(i, 21).Value)
                        item.F1 = Val(ws.Cells(i, 7).Value)
                        item.F2 = Val(ws.Cells(i, 8).Value)
                        item.F3 = Val(ws.Cells(i, 9).Value)
                        item.F4 = Val(ws.Cells(i, 10).Value)
                        item.F5 = Val(ws.Cells(i, 11).Value)
                        item.F6 = Val(ws.Cells(i, 12).Value)
                        item.F7 = Val(ws.Cells(i, 13).Value)
                        item.Beta = Val(ws.Cells(i, 22).Value)
                        beamList.Add(item)
                    Next
                    If beamList.Count = 0 Then
                        result = "No valid beam data found for frame analysis."
                    Else
                        Dim filtered As List(Of BeamData) = FilterBeamsForDefinedGridsOnly(beamList, text, ws)
                        If filtered.Count = 0 Then
                            result = $"No beams found in predefined grids for {text} direction analysis. Check your grid definition using 'G' option first."
                        Else
                            Dim excludedCount As Integer = beamList.Count - filtered.Count
                            Dim excludedMsg As String = $"Analyzing only predefined grids: {filtered.Count} beams from defined grid coordinates. Excluded {excludedCount} beams outside grid system."
                            Dim analysisText As String = AnalyzeSimilarFrames(filtered, text, ws)
                            result = $"Grid-Restricted Frame Analysis for '{text}' direction:" & vbCrLf & analysisText & vbCrLf & vbCrLf & excludedMsg
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            result = $"Frame analysis error: {ex.Message}"
        End Try
        Return result
    End Function

    Private Function ProcessBothDirectionsForGridSilent(ws As Object, staadApp As Object) As String
        Dim result As String
        Try
            If HasExistingGridData(ws) Then
                Interaction.MsgBox("Nova coordinates must be cleared before Genrating new grids.", MsgBoxStyle.Exclamation, "Existing Grid Detected")
                result = ""
            Else
                Dim beamList As New List(Of BeamData)()
                Dim lastRow As Integer = CInt(ws.Cells(ws.Rows.Count, 1).End(-4162).Row)
                If lastRow < 9 Then
                    result = "No beam data found for grid generation."
                Else
                    For i As Integer = 9 To lastRow
                        If ws.Cells(i, 2).Value Is Nothing Then Continue For
                        Dim item As BeamData = Nothing
                        item.BeamNo = CInt(Math.Round(Val(ws.Cells(i, 2).Value)))
                        item.Length = Val(ws.Cells(i, 3).Value)
                        If ws.Cells(i, 4).Value Is Nothing Then
                            item.SectionName = ""
                        Else
                            item.SectionName = ws.Cells(i, 4).Value.ToString()
                        End If
                        item.NodeA = CInt(Math.Round(Val(ws.Cells(i, 14).Value)))
                        item.NodeB = CInt(Math.Round(Val(ws.Cells(i, 15).Value)))
                        item.XA = Val(ws.Cells(i, 16).Value)
                        item.YA = Val(ws.Cells(i, 17).Value)
                        item.ZA = Val(ws.Cells(i, 18).Value)
                        item.XB = Val(ws.Cells(i, 19).Value)
                        item.YB = Val(ws.Cells(i, 20).Value)
                        item.ZB = Val(ws.Cells(i, 21).Value)
                        item.F1 = Val(ws.Cells(i, 7).Value)
                        item.F2 = Val(ws.Cells(i, 8).Value)
                        item.F3 = Val(ws.Cells(i, 9).Value)
                        item.F4 = Val(ws.Cells(i, 10).Value)
                        item.F5 = Val(ws.Cells(i, 11).Value)
                        item.F6 = Val(ws.Cells(i, 12).Value)
                        item.F7 = Val(ws.Cells(i, 13).Value)
                        item.Beta = Val(ws.Cells(i, 22).Value)
                        beamList.Add(item)
                    Next
                    If beamList.Count = 0 Then
                        result = "No valid beam data found for grid generation."
                    Else
                        Dim summary As New List(Of String)()
                        Dim beamsX As List(Of BeamData) = FilterBeamsForFrameAnalysis(beamList, "X")
                        If beamsX.Count > 0 Then
                            Dim gridsX As Dictionary(Of String, List(Of BeamData)) = CreateGridSystem(beamsX, "X", ws)
                            summary.Add($"X Direction: {gridsX.Count} grids created")
                        Else
                            summary.Add("X Direction: No beams found in YZ plane")
                        End If
                        Dim beamsZ As List(Of BeamData) = FilterBeamsForFrameAnalysis(beamList, "Z")
                        If beamsZ.Count > 0 Then
                            Dim gridsZ As Dictionary(Of String, List(Of BeamData)) = CreateGridSystem(beamsZ, "Z", ws)
                            summary.Add($"Z Direction: {gridsZ.Count} grids created")
                        Else
                            summary.Add("Z Direction: No beams found in XY plane")
                        End If
                        Dim text As String = "NOVA GRID SYSTEM GENERATED SUCCESSFULLY!" & vbCrLf & vbCrLf &
                            "Grid data has been created in NovaGrid System:" & vbCrLf &
                            String.Join(vbCrLf, summary) & vbCrLf & vbCrLf &
                            "You can now perform frame analysis using 'X' or 'Z' options." & vbCrLf &
                            "Analysis will be restricted to these predefined grids only."
                        result = text
                    End If
                End If
            End If
        Catch ex As Exception
            result = $"Grid generation error: {ex.Message}"
        End Try
        Return result
    End Function

    Private Function HasExistingGridData(ws As Object) As Boolean
        Try
            Dim wb As Object = ws.Parent
            Dim sheet4 As Object
            Try
                sheet4 = wb.Sheets("Sheet4")
            Catch
                Return False
            End Try
            For Each addr As String In New String() {"C2", "C3", "C4"}
                If sheet4.Range(addr).Value IsNot Nothing Then
                    If Not String.IsNullOrEmpty(sheet4.Range(addr).Value.ToString().Trim()) Then
                        Return True
                    End If
                End If
            Next
            Return False
        Catch
            Return False
        End Try
    End Function

    Private Function AreGridsDefined(direction As String, ws As Object) As Boolean
        Try
            Dim wb As Object = ws.Parent
            Dim sheet4 As Object
            Try
                sheet4 = wb.Sheets("Sheet4")
            Catch
                Return False
            End Try
            Dim nameCell As String, coordCell As String
            If direction.ToUpper() = "X" Then
                nameCell = "B2" : coordCell = "C2"
            ElseIf direction.ToUpper() = "Z" Then
                nameCell = "B4" : coordCell = "C4"
            Else
                Return False
            End If
            Dim nameOk As Boolean = False
            Dim coordOk As Boolean = False
            If sheet4.Range(nameCell).Value IsNot Nothing Then
                If Not String.IsNullOrEmpty(sheet4.Range(nameCell).Value.ToString().Trim()) Then
                    nameOk = True
                End If
            End If
            If sheet4.Range(coordCell).Value IsNot Nothing Then
                If Not String.IsNullOrEmpty(sheet4.Range(coordCell).Value.ToString().Trim()) Then
                    coordOk = True
                End If
            End If
            Return nameOk AndAlso coordOk
        Catch
            Return False
        End Try
    End Function

    Private Function FilterBeamsForDefinedGridsOnly(beamDataList As List(Of BeamData), direction As String, ws As Object) As List(Of BeamData)
        Dim result As New List(Of BeamData)()
        Const tolerance As Double = 0.01
        Dim definedGridCoordinates As List(Of Double) = GetDefinedGridCoordinates(direction, ws)
        If definedGridCoordinates.Count = 0 Then Return result
        For Each beam As BeamData In beamDataList
            Dim isInGrid As Boolean = False
            Dim coord As Double = 0.0
            If direction = "Z" Then
                If Math.Abs(beam.ZA - beam.ZB) <= tolerance Then
                    coord = Math.Round((beam.ZA + beam.ZB) / 2.0, 2)
                    isInGrid = IsCoordinateInDefinedGrids(coord, definedGridCoordinates, tolerance)
                End If
            ElseIf direction = "X" AndAlso Math.Abs(beam.XA - beam.XB) <= tolerance Then
                coord = Math.Round((beam.XA + beam.XB) / 2.0, 2)
                isInGrid = IsCoordinateInDefinedGrids(coord, definedGridCoordinates, tolerance)
            End If
            If isInGrid Then result.Add(beam)
        Next
        Return result
    End Function

    Private Function GetDefinedGridCoordinates(direction As String, ws As Object) As List(Of Double)
        Dim result As New List(Of Double)()
        Try
            result = GetCompleteGridDataFromSheet4(direction, ws).AllCoordinates
        Catch
        End Try
        Return result
    End Function

    Private Function IsCoordinateInDefinedGrids(coordinate As Double, definedCoordinates As List(Of Double), tolerance As Double) As Boolean
        For Each definedCoordinate As Double In definedCoordinates
            If Math.Abs(coordinate - definedCoordinate) <= tolerance Then Return True
        Next
        Return False
    End Function

    Private Function FilterBeamsForFrameAnalysis(beamDataList As List(Of BeamData), direction As String) As List(Of BeamData)
        Dim result As New List(Of BeamData)()
        Const tolerance As Double = 0.01
        For Each beam As BeamData In beamDataList
            Dim keep As Boolean = False
            If direction = "Z" Then
                If Math.Abs(beam.ZA - beam.ZB) <= tolerance Then keep = True
            ElseIf direction = "X" AndAlso Math.Abs(beam.XA - beam.XB) <= tolerance Then
                keep = True
            End If
            If keep Then result.Add(beam)
        Next
        Return result
    End Function

    Private Function CreateGridSystem(beamDataList As List(Of BeamData), direction As String, ws As Object) As Dictionary(Of String, List(Of BeamData))
        Dim dictionary As New Dictionary(Of String, List(Of BeamData))()
        Const tolerance As Double = 0.01
        Dim sheet4Data As Sheet4GridData = GetCompleteGridDataFromSheet4(direction, ws)
        Dim byCoord As New Dictionary(Of Double, List(Of BeamData))()
        For Each beam As BeamData In beamDataList
            Dim coord As Double = 0.0
            If direction = "Z" Then
                coord = Math.Round((beam.ZA + beam.ZB) / 2.0, 2)
            ElseIf direction = "X" Then
                coord = Math.Round((beam.XA + beam.XB) / 2.0, 2)
            End If
            Dim matchedKey As Double = -999999.0
            For Each key As Double In byCoord.Keys
                If Math.Abs(key - coord) <= tolerance Then
                    matchedKey = key
                    Exit For
                End If
            Next
            If matchedKey = -999999.0 Then
                byCoord(coord) = New List(Of BeamData)()
                matchedKey = coord
            End If
            byCoord(matchedKey).Add(beam)
        Next
        Dim sortedKeys As List(Of Double) = byCoord.Keys.OrderBy(Function(x) x).ToList()
        If ShouldUpdateSheet4CoordinateSpacing(sortedKeys, sheet4Data.AllCoordinates, direction) Then
            WriteCoordinateSpacingToSheet4(sortedKeys, direction, ws)
        End If
        If sheet4Data.GridNames.Count > 0 AndAlso sheet4Data.AllCoordinates.Count > 0 Then
            For Each coord As Double In sortedKeys
                Dim matchIndex As Integer = FindBestCoordinateMatch(coord, sheet4Data.AllCoordinates, tolerance)
                If matchIndex >= 0 AndAlso matchIndex < sheet4Data.GridNames.Count Then
                    Dim gridName As String = sheet4Data.GridNames(matchIndex)
                    Dim key As String = $"{gridName} ({direction}={coord:F2})"
                    dictionary(key) = byCoord(coord)
                Else
                    Dim key As String = $"UNMATCHED-{coord:F2} ({direction}={coord:F2})"
                    dictionary(key) = byCoord(coord)
                End If
            Next
        Else
            Dim negatives As List(Of Double) = sortedKeys.Where(Function(c) c < -0.005).OrderByDescending(Function(c) c).ToList()
            Dim zeros As List(Of Double) = sortedKeys.Where(Function(c) Math.Abs(c) < 0.005).ToList()
            Dim positives As List(Of Double) = sortedKeys.Where(Function(c) c > 0.005).OrderBy(Function(c) c).ToList()
            Dim negIndex As Integer = 1
            Dim posIndex As Integer = 1
            For Each coord As Double In negatives
                Dim key As String = $"-{direction}{negIndex} ({direction}={coord:F2})"
                dictionary(key) = byCoord(coord)
                negIndex += 1
            Next
            For Each coord As Double In zeros
                Dim key As String = $"{direction}0 ({direction}={coord:F2})"
                dictionary(key) = byCoord(coord)
            Next
            For Each coord As Double In positives
                Dim key As String = $"{direction}{posIndex} ({direction}={coord:F2})"
                dictionary(key) = byCoord(coord)
                posIndex += 1
            Next
        End If
        Return dictionary
    End Function

    Private Function GetCompleteGridDataFromSheet4(direction As String, ws As Object) As Sheet4GridData
        Dim result As New Sheet4GridData With {
            .GridNames = New List(Of String)(),
            .AllCoordinates = New List(Of Double)(),
            .SpacingString = ""
        }
        Try
            Dim wb As Object = ws.Parent
            Dim sheet4 As Object = wb.Sheets("Sheet4")
            Dim nameCell As String, spacingCell As String
            If direction.ToUpper() = "X" Then
                nameCell = "B2" : spacingCell = "C2"
            ElseIf direction.ToUpper() = "Z" Then
                nameCell = "B4" : spacingCell = "C4"
            Else
                Return result
            End If
            If sheet4.Range(nameCell).Value IsNot Nothing Then
                Dim namesRaw As String = sheet4.Range(nameCell).Value.ToString().Trim()
                If Not String.IsNullOrEmpty(namesRaw) Then
                    Dim parts() As String = namesRaw.Split(New Char() {" "c, vbTab.Chars(0)}, StringSplitOptions.RemoveEmptyEntries)
                    For Each part As String In parts
                        Dim trimmed As String = part.Trim()
                        If Not String.IsNullOrEmpty(trimmed) Then result.GridNames.Add(trimmed)
                    Next
                End If
            End If
            If sheet4.Range(spacingCell).Value IsNot Nothing Then
                result.SpacingString = sheet4.Range(spacingCell).Value.ToString().Trim()
                If Not String.IsNullOrEmpty(result.SpacingString) Then
                    result.AllCoordinates = ParseCoordinateSpacing(result.SpacingString)
                End If
            End If
        Catch
        End Try
        Return result
    End Function

    Private Function ParseCoordinateSpacing(spacingString As String) As List(Of Double)
        Dim list As New List(Of Double)()
        Try
            Dim tokens() As String = spacingString.Split(New Char() {" "c}, StringSplitOptions.RemoveEmptyEntries)
            If tokens.Length = 0 Then Return list
            Dim running As Double = Double.Parse(tokens(0)) / 1000.0
            list.Add(Math.Round(running, 3))
            For i As Integer = 1 To tokens.Length - 1
                Dim token As String = tokens(i).Trim()
                If token.Contains("*") Then
                    Dim pieces() As String = token.Split("*"c)
                    If pieces.Length = 2 Then
                        Dim repeatCount As Integer = Integer.Parse(pieces(0))
                        Dim step As Double = Double.Parse(pieces(1)) / 1000.0
                        For j As Integer = 1 To repeatCount
                            running += step
                            list.Add(Math.Round(running, 3))
                        Next
                    End If
                Else
                    Dim step As Double = Double.Parse(token) / 1000.0
                    running += step
                    list.Add(Math.Round(running, 3))
                End If
            Next
        Catch
            list.Clear()
        End Try
        Return list
    End Function

    Private Function FindBestCoordinateMatch(targetCoord As Double, sheet4Coordinates As List(Of Double), tolerance As Double) As Integer
        Dim result As Integer = -1
        Dim best As Double = Double.MaxValue
        For i As Integer = 0 To sheet4Coordinates.Count - 1
            Dim diff As Double = Math.Abs(targetCoord - sheet4Coordinates(i))
            If diff <= tolerance AndAlso diff < best Then
                best = diff
                result = i
            End If
        Next
        Return result
    End Function

    Private Function ShouldUpdateSheet4CoordinateSpacing(analysisCoords As List(Of Double), sheet4Coords As List(Of Double), direction As String) As Boolean
        If sheet4Coords.Count = 0 Then Return True
        If analysisCoords.Count >= sheet4Coords.Count Then Return True
        Return False
    End Function

    Private Sub WriteCoordinateSpacingToSheet4(sortedCoordinates As List(Of Double), direction As String, ws As Object)
        Try
            Dim wb As Object = ws.Parent
            Dim sheet4 As Object
            Try
                sheet4 = wb.Sheets("Sheet4")
            Catch
                Return
            End Try
            Dim spacingCell As String
            If direction.ToUpper() = "X" Then
                spacingCell = "C2"
            ElseIf direction.ToUpper() = "Z" Then
                spacingCell = "C4"
            Else
                Return
            End If
            If sheet4.Range(spacingCell).Value IsNot Nothing Then
                If Not String.IsNullOrEmpty(sheet4.Range(spacingCell).Value.ToString().Trim()) Then
                    Dim existingSpacing As String = sheet4.Range(spacingCell).Value.ToString().Trim()
                    Dim existingCoords As List(Of Double) = ParseCoordinateSpacing(existingSpacing)
                    If sortedCoordinates.Count < existingCoords.Count Then Return
                End If
            End If
            If sortedCoordinates.Count < 2 Then Return
            Dim diffsMm As New List(Of Integer)()
            For i As Integer = 1 To sortedCoordinates.Count - 1
                Dim diff As Double = sortedCoordinates(i) - sortedCoordinates(i - 1)
                diffsMm.Add(CInt(Math.Round(Math.Abs(diff) * 1000.0)))
            Next
            Dim parts As New List(Of String)()
            parts.Add(CInt(Math.Round(sortedCoordinates(0) * 1000.0)).ToString())
            Dim j As Integer = 0
            While j < diffsMm.Count
                Dim value As Integer = diffsMm(j)
                Dim runLength As Integer = 1
                While j + runLength < diffsMm.Count AndAlso diffsMm(j + runLength) = value
                    runLength += 1
                End While
                If runLength = 1 Then
                    parts.Add(value.ToString())
                Else
                    parts.Add($"{runLength}*{value}")
                End If
                j += runLength
            End While
            Dim spacingText As String = String.Join(" ", parts)
            sheet4.Range(spacingCell).Value = spacingText
        Catch
        End Try
    End Sub

    Private Function AnalyzeSimilarFrames(beamDataList As List(Of BeamData), direction As String, ws As Object) As String
        Dim allGrids As Dictionary(Of String, List(Of BeamData)) = CreateGridSystem(beamDataList, direction, ws)
        Dim eligibleGrids As Dictionary(Of String, List(Of BeamData)) = allGrids.Where(Function(g) g.Value.Count >= 2).ToDictionary(Function(k) k.Key, Function(v) v.Value)
        If eligibleGrids.Count < 2 Then
            Return "Insufficient grids for comparison analysis. Need at least 2 grids with 2+ beams each."
        End If
        Dim similarGroups As New List(Of List(Of KeyValuePair(Of String, List(Of BeamData))))()
        Dim visited As New HashSet(Of String)()
        For Each entry As KeyValuePair(Of String, List(Of BeamData)) In eligibleGrids
            If visited.Contains(entry.Key) Then Continue For
            Dim group As New List(Of KeyValuePair(Of String, List(Of BeamData)))()
            group.Add(entry)
            visited.Add(entry.Key)
            For Each other As KeyValuePair(Of String, List(Of BeamData)) In eligibleGrids
                If Not visited.Contains(other.Key) AndAlso AreFramesSimilar(entry.Value, other.Value, direction) Then
                    group.Add(other)
                    visited.Add(other.Key)
                End If
            Next
            If group.Count > 1 Then similarGroups.Add(group)
        Next
        If direction <> "G" Then
            Dim content As String = GenerateEnhancedFrameAnalysisLog(similarGroups, allGrids, direction, ws)
            _lastLogFilePath = SaveLogFile(content, direction, ws)
            Return $"Frame analysis completed. Similar sets: {similarGroups.Count}. Log saved to file."
        End If
        _lastLogFilePath = ""
        Return $"Grid system created. Similar sets: {similarGroups.Count}"
    End Function

    Private Function GetGridNamingInfoForResult(direction As String, ws As Object) As String
        Dim result As String
        Try
            Dim wb As Object = ws.Parent
            Dim sheet4 As Object = wb.Sheets("Sheet4")
            Dim gridNameText As String = ""
            If direction.ToUpper() = "X" Then
                If sheet4.Range("B2").Value IsNot Nothing Then
                    gridNameText = sheet4.Range("B2").Value.ToString().Trim()
                End If
            ElseIf direction.ToUpper() = "Z" Then
                If sheet4.Range("B4").Value IsNot Nothing Then
                    gridNameText = sheet4.Range("B4").Value.ToString().Trim()
                End If
            End If
            If String.IsNullOrEmpty(gridNameText) Then
                result = "default naming (Nova's Grid System is empty)"
            ElseIf GetCompleteGridDataFromSheet4(direction, ws).AllCoordinates.Count <= 0 Then
                result = $"From Nova's Grid System: '{gridNameText}'"
            Else
                result = $"From Nova's Grid System: '{gridNameText}' (Coordinate-matched)"
            End If
        Catch
            result = "default naming (Nova's Grid System not available)"
        End Try
        Return result
    End Function

    Private Function AreFramesSimilar(frame1 As List(Of BeamData), frame2 As List(Of BeamData), direction As String) As Boolean
        If frame1.Count <> frame2.Count Then Return False
        Dim list1 As New List(Of BeamData)(frame1)
        Dim list2 As New List(Of BeamData)(frame2)
        Dim origin1a As Double, origin1b As Double, origin2a As Double, origin2b As Double
        If direction = "Z" Then
            origin1a = list1.Min(Function(b) Math.Min(b.XA, b.XB))
            origin1b = list1.Min(Function(b) Math.Min(b.YA, b.YB))
            origin2a = list2.Min(Function(b) Math.Min(b.XA, b.XB))
            origin2b = list2.Min(Function(b) Math.Min(b.YA, b.YB))
        Else
            If direction <> "X" Then Return False
            origin1a = list1.Min(Function(b) Math.Min(b.YA, b.YB))
            origin1b = list1.Min(Function(b) Math.Min(b.ZA, b.ZB))
            origin2a = list2.Min(Function(b) Math.Min(b.YA, b.YB))
            origin2b = list2.Min(Function(b) Math.Min(b.ZA, b.ZB))
        End If
        For i As Integer = 0 To list1.Count - 1
            Dim value As BeamData = list1(i)
            If direction = "Z" Then
                value.XA -= origin1a : value.XB -= origin1a
                value.YA -= origin1b : value.YB -= origin1b
            Else
                value.YA -= origin1a : value.YB -= origin1a
                value.ZA -= origin1b : value.ZB -= origin1b
            End If
            list1(i) = value
        Next
        For i As Integer = 0 To list2.Count - 1
            Dim value As BeamData = list2(i)
            If direction = "Z" Then
                value.XA -= origin2a : value.XB -= origin2a
                value.YA -= origin2b : value.YB -= origin2b
            Else
                value.YA -= origin2a : value.YB -= origin2a
                value.ZA -= origin2b : value.ZB -= origin2b
            End If
            list2(i) = value
        Next
        Dim ordered1 As List(Of BeamData), ordered2 As List(Of BeamData)
        If direction = "Z" Then
            ordered1 = list1.OrderBy(Function(b) Math.Min(b.XA, b.XB)).ThenBy(Function(b) Math.Min(b.YA, b.YB)).ToList()
            ordered2 = list2.OrderBy(Function(b) Math.Min(b.XA, b.XB)).ThenBy(Function(b) Math.Min(b.YA, b.YB)).ToList()
        Else
            ordered1 = list1.OrderBy(Function(b) Math.Min(b.YA, b.YB)).ThenBy(Function(b) Math.Min(b.ZA, b.ZB)).ToList()
            ordered2 = list2.OrderBy(Function(b) Math.Min(b.YA, b.YB)).ThenBy(Function(b) Math.Min(b.ZA, b.ZB)).ToList()
        End If
        For i As Integer = 0 To ordered1.Count - 1
            Dim beam As BeamData = ordered1(i)
            Dim beam2 As BeamData = ordered2(i)
            If Not AreSectionsSimilar(beam, beam2) Then Return False
            If Math.Abs(beam.Beta - beam2.Beta) > 0.5 Then Return False
            If Not AreBeamPositionsSimilar(beam, beam2, direction) Then Return False
            If Math.Abs(beam.Length - beam2.Length) > 0.01 Then Return False
            Dim span1 As Double, span2 As Double
            If direction = "Z" Then
                span1 = beam.ZB - beam.ZA
                span2 = beam2.ZB - beam2.ZA
            Else
                span1 = beam.XB - beam.XA
                span2 = beam2.XB - beam2.XA
            End If
            If Math.Abs(span1 - span2) > 0.01 Then Return False
        Next
        Return True
    End Function

    Private Function AreBeamPositionsSimilar(beam1 As BeamData, beam2 As BeamData, direction As String) As Boolean
        Const tolerance As Double = 0.01
        If direction = "Z" Then
            Return Math.Abs(beam1.XA - beam2.XA) <= tolerance AndAlso Math.Abs(beam1.YA - beam2.YA) <= tolerance AndAlso Math.Abs(beam1.XB - beam2.XB) <= tolerance AndAlso Math.Abs(beam1.YB - beam2.YB) <= tolerance
        End If
        If direction = "X" Then
            Return Math.Abs(beam1.YA - beam2.YA) <= tolerance AndAlso Math.Abs(beam1.ZA - beam2.ZA) <= tolerance AndAlso Math.Abs(beam1.YB - beam2.YB) <= tolerance AndAlso Math.Abs(beam1.ZB - beam2.ZB) <= tolerance
        End If
        Return False
    End Function

    Private Function AreSectionsSimilar(beam1 As BeamData, beam2 As BeamData) As Boolean
        Const tolerance As Double = 0.001
        If beam1.SectionName <> beam2.SectionName Then Return False
        If Math.Abs(beam1.F1 - beam2.F1) > tolerance Then Return False
        If Math.Abs(beam1.F2 - beam2.F2) > tolerance Then Return False
        If Math.Abs(beam1.F3 - beam2.F3) > tolerance Then Return False
        If Math.Abs(beam1.F4 - beam2.F4) > tolerance Then Return False
        If Math.Abs(beam1.F5 - beam2.F5) > tolerance Then Return False
        If Math.Abs(beam1.F6 - beam2.F6) > tolerance Then Return False
        If Math.Abs(beam1.F7 - beam2.F7) > tolerance Then Return False
        Return True
    End Function

    Private Function GenerateEnhancedFrameAnalysisLog(similarGridGroups As List(Of List(Of KeyValuePair(Of String, List(Of BeamData)))), allGridGroups As Dictionary(Of String, List(Of BeamData)), direction As String, Optional ws As Object = Nothing) As String
        Dim sb As New StringBuilder()
        sb.AppendLine("🏗️ NOVA'S STAAD FRAME SIMILARITY ANALYSIS 🏗️")
        sb.AppendLine("📐 Advanced Grid Reference System 📐")
        sb.AppendLine()
        sb.AppendLine("=" & New String("="c, 80))
        sb.AppendLine()
        sb.AppendLine("📊 ANALYSIS METADATA")
        sb.AppendLine()
        sb.AppendLine("📅 Analysis Date: " & DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        sb.AppendLine("🧭 Analysis Direction: " & direction)
        sb.AppendLine("📏 Total Grids Analyzed: " & allGridGroups.Count)
        sb.AppendLine("🎯 Similar Groups Found: " & similarGridGroups.Count)
        If ws IsNot Nothing Then
            Dim namingInfo As String = GetGridNamingInfoForResult(direction, ws)
            sb.AppendLine("📋 Grid Naming Source: " & namingInfo)
            Try
                Dim wb As Object = ws.Parent
                Dim sheet4 As Object = wb.Sheets("Sheet4")
                Dim spacingCell As String = If(direction.ToUpper() = "X", "C2", "C4")
                If sheet4.Range(spacingCell).Value IsNot Nothing Then
                    sb.AppendLine("📐 Coordinate Spacing: " & String.Format("{0} ({1})", sheet4.Range(spacingCell).Value, spacingCell))
                End If
            Catch
                sb.AppendLine("📐 Coordinate Spacing: Sheet4 not accessible")
            End Try
        End If
        sb.AppendLine()
        sb.AppendLine("=" & New String("="c, 80))
        sb.AppendLine()
        sb.AppendLine("🎯 SIMILAR GRID GROUPS SUMMARY")
        sb.AppendLine()
        sb.AppendLine("Group #" & New String(" "c, 10) & "➤ Grid Count" & New String(" "c, 10) & "➤ Grid Names")
        sb.AppendLine(New String("-"c, 80))
        For i As Integer = 0 To similarGridGroups.Count - 1
            Dim group As List(Of KeyValuePair(Of String, List(Of BeamData))) = similarGridGroups(i)
            Dim namesJoined As String = String.Join(", ", group.Select(Function(g) g.Key.Split("("c)(0).Trim()))
            If namesJoined.Length > 45 Then
                namesJoined = namesJoined.Substring(0, 42) & "..."
            End If
            sb.AppendLine($"Group {(i + 1).ToString().PadLeft(2)}" & New String(" "c, 12) & "➤ " & group.Count.ToString().PadLeft(6) & New String(" "c, 12) & "➤ " & namesJoined)
        Next
        sb.AppendLine()
        sb.AppendLine("=" & New String("="c, 80))
        sb.AppendLine()
        sb.AppendLine("🗂️ GRID SYSTEM OVERVIEW")
        sb.AppendLine()
        sb.AppendLine("Grid Name" & New String(" "c, 30) & "Beam Count" & New String(" "c, 10) & "Coordinate")
        sb.AppendLine(New String("-"c, 80))
        For Each item As KeyValuePair(Of String, List(Of BeamData)) In allGridGroups.OrderBy(Function(g) ExtractCoordinateFromGridName(g.Key))
            Dim gridDisplayName As String = item.Key.Split("("c)(0).Trim()
            Dim coordText As String = ExtractCoordinateFromGridName(item.Key).ToString("F1")
            sb.AppendLine(gridDisplayName.PadRight(35) & item.Value.Count.ToString().PadLeft(15) & coordText.PadLeft(15))
        Next
        sb.AppendLine()
        sb.AppendLine("=" & New String("="c, 80))
        sb.AppendLine()
        sb.AppendLine("📋 DETAILED GROUP ANALYSIS")
        sb.AppendLine()
        For groupIndex As Integer = 0 To similarGridGroups.Count - 1
            Dim group As List(Of KeyValuePair(Of String, List(Of BeamData))) = similarGridGroups(groupIndex)
            sb.AppendLine($"🏗️ SIMILAR GRID GROUP #{groupIndex + 1}")
            sb.AppendLine()
            sb.AppendLine($"📊 Similar Grids Count: {group.Count}")
            sb.AppendLine($"🔧 Beams per Grid: {group(0).Value.Count}")
            sb.AppendLine()
            sb.AppendLine("🗂️ GRIDS IN THIS GROUP:")
            For Each item As KeyValuePair(Of String, List(Of BeamData)) In group
                sb.AppendLine($"   ➤ {item.Key}")
            Next
            sb.AppendLine()
            For gridIndex As Integer = 0 To group.Count - 1
                Dim kv As KeyValuePair(Of String, List(Of BeamData)) = group(gridIndex)
                sb.AppendLine($"📐 {kv.Key} DETAILS:")
                sb.AppendLine()
                For Each beam As BeamData In kv.Value.OrderBy(Function(b) b.BeamNo)
                    sb.AppendLine($"   🔹 Beam {beam.BeamNo.ToString().PadLeft(3)}: {beam.SectionName}")
                    sb.AppendLine($"      🔗 Nodes: {beam.NodeA} ➜ {beam.NodeB}")
                    sb.AppendLine($"      📍 Coord A: ({beam.XA:F3}, {beam.YA:F3}, {beam.ZA:F3})")
                    sb.AppendLine($"      📍 Coord B: ({beam.XB:F3}, {beam.YB:F3}, {beam.ZB:F3})")
                    sb.AppendLine($"      📏 Length: {beam.Length:F3}m")
                    sb.AppendLine($"      ⚙️ Section Props: F1={beam.F1:F3} F2={beam.F2:F3} F3={beam.F3:F3}")
                    sb.AppendLine($"         F4={beam.F4:F3} F5={beam.F5:F3} F6={beam.F6:F3} F7={beam.F7:F3}")
                    sb.AppendLine()
                Next
                If gridIndex < group.Count - 1 Then
                    sb.AppendLine(New String("-"c, 50))
                    sb.AppendLine()
                End If
            Next
            If groupIndex < similarGridGroups.Count - 1 Then
                sb.AppendLine("=" & New String("="c, 80))
                sb.AppendLine()
            End If
        Next
        sb.AppendLine("=" & New String("="c, 80))
        sb.AppendLine()
        sb.AppendLine("📈 COMPREHENSIVE ANALYSIS SUMMARY")
        sb.AppendLine()
        Dim totalSimilarGrids As Integer = 0
        Dim totalBeamsInSimilarGrids As Integer = 0
        For Each group As List(Of KeyValuePair(Of String, List(Of BeamData))) In similarGridGroups
            totalSimilarGrids += group.Count
            For Each item As KeyValuePair(Of String, List(Of BeamData)) In group
                totalBeamsInSimilarGrids += item.Value.Count
            Next
        Next
        sb.AppendLine($"Total Grids with Similar Patterns: {totalSimilarGrids}")
        sb.AppendLine($"Total Beams in Similar Grids: {totalBeamsInSimilarGrids}")
        If totalSimilarGrids > 0 Then
            Dim avg As String = (CDbl(totalBeamsInSimilarGrids) / CDbl(totalSimilarGrids)).ToString("F1")
            sb.AppendLine($"Average Beams per Similar Grid: {avg}")
        Else
            sb.AppendLine("Average Beams per Similar Grid: 0")
        End If
        Dim efficiency As String = If(allGridGroups.Count > 0, (CDbl(totalSimilarGrids) / CDbl(allGridGroups.Count) * 100.0).ToString("F1"), "0")
        sb.AppendLine($"Similarity Efficiency: {efficiency}%")
        sb.AppendLine()
        sb.AppendLine("=" & New String("="c, 80))
        sb.AppendLine()
        sb.AppendLine("*** Report Generated by PlateNova's RaghavStaadDataExtractor ***")
        sb.AppendLine("Enhanced with Nova Grid Intelligence System")
        sb.AppendLine(String.Format("Generated on: {0}", DateTime.Now.ToString("dddd, MMMM dd, yyyy 'at' HH:mm:ss")))
        Return sb.ToString()
    End Function

    Private Function ExtractCoordinateFromGridName(gridName As String) As Double
        Try
            Dim startIndex As Integer = gridName.IndexOf("=") + 1
            Dim endIndex As Integer = gridName.IndexOf(")")
            If startIndex > 0 AndAlso endIndex > startIndex Then
                Dim s As String = gridName.Substring(startIndex, endIndex - startIndex)
                Return Double.Parse(s)
            End If
        Catch
        End Try
        Return 0.0
    End Function

    Private Function SaveLogFile(content As String, direction As String, ws As Object) As String
        Dim result As String
        Try
            Dim fileName As String = String.Format("Frame_Analysis_{0}_Direction_{1}.txt", direction, DateTime.Now.ToString("yyyyMMdd_HHmmss"))
            Dim wb As Object = ws.Parent
            Dim baseFolder As String
            If String.IsNullOrEmpty(CStr(wb.Path)) Then
                baseFolder = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
            Else
                baseFolder = CStr(wb.Path)
            End If
            Dim folder As String = Path.Combine(baseFolder, "Frame Analysis")
            If Not Directory.Exists(folder) Then Directory.CreateDirectory(folder)
            Dim fullPath As String = Path.Combine(folder, fileName)
            File.WriteAllText(fullPath, content)
            result = fullPath
        Catch
            Try
                Dim fileName As String = String.Format("Frame_Analysis_{0}_Direction_{1}.txt", direction, DateTime.Now.ToString("yyyyMMdd_HHmmss"))
                Dim desktopFolder As String = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
                Dim folder As String = Path.Combine(desktopFolder, "Frame Analysis")
                If Not Directory.Exists(folder) Then Directory.CreateDirectory(folder)
                Dim fullPath As String = Path.Combine(folder, fileName)
                File.WriteAllText(fullPath, content)
                result = fullPath
            Catch ex2 As Exception
                result = $"Error saving log file: {ex2.Message}"
            End Try
        End Try
        Return result
    End Function

    Private Sub AutoSaveWorkbook(workbook As Object, Optional savePath As String = "")
        Try
            If workbook Is Nothing Then Return
            If String.IsNullOrEmpty(savePath) Then
                If CStr(workbook.Path) <> "" Then
                    workbook.Save()
                    Return
                End If
                Dim folder As String = Environment.GetFolderPath(Environment.SpecialFolder.Personal)
                savePath = Path.Combine(folder, "Staad_Extract_" & DateAndTime.Now.ToString("yyyyMMdd_HHmmss") & ".xlsx")
                workbook.SaveAs(savePath)
            Else
                workbook.SaveAs(savePath)
            End If
        Catch
        End Try
    End Sub

    Private Sub EnsureLicense()
        Dim executionValidation As New ExecutionValidation()
        If Not executionValidation.IsLicenceValid() Then
            Throw New UnauthorizedAccessException("License check failed.")
        End If
    End Sub

    Private Sub CheckExpirationDateAndBackdate()
        Dim expiryDate As New DateTime(2027, 4, 9)
        Dim today As DateTime = DateTime.Now.Date
        Dim stampPath As String = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "RaghavStaad.last")
        If File.Exists(stampPath) Then
            Dim s As String = File.ReadAllText(stampPath).Trim()
            Dim lastRun As DateTime = DateTime.MinValue
            If DateTime.TryParse(s, lastRun) Then
                Dim lastRunExact As DateTime = DateTime.Parse(s)
                If DateTime.Compare(today, lastRunExact) < 0 Then
                    Throw New UnauthorizedAccessException("System date has been tampered. Please correct your clock.")
                End If
            End If
        End If
        File.WriteAllText(stampPath, today.ToString("yyyy-MM-dd"))
        If DateTime.Compare(today, expiryDate) > 0 Then
            Throw New UnauthorizedAccessException("Software expired. Please contact support for an updated version.")
        End If
    End Sub

    Private Function GetOrCreateExcelApp() As Object
        Dim result As Object
        Try
            result = Marshal.GetActiveObject("Excel.Application")
        Catch
            Dim app As Object = Interaction.CreateObject("Excel.Application")
            app.Visible = True
            result = app
        End Try
        Return result
    End Function

    Private Function GetStaadApplication() As Object
        Dim result As Object
        Try
            result = Marshal.GetActiveObject("StaadPro.OpenSTAAD")
        Catch
            result = Nothing
        End Try
        Return result
    End Function

    Private Function GetOrCreateWorksheet(workbook As Object, sheetName As String) As Object
        Dim result As Object
        Try
            result = workbook.Sheets(sheetName)
        Catch
            Dim newSheet As Object = workbook.Sheets.Add()
            newSheet.Name = sheetName
            result = newSheet
        End Try
        Return result
    End Function

    Private Sub PreserveAndClearWorksheet(ws As Object)
        Dim as2 As Object = ws.Range("AS2").Value
        Dim as3 As Object = ws.Range("AS3").Value
        Dim as4 As Object = ws.Range("AS4").Value
        Dim as5 As Object = ws.Range("AS5").Value
        Dim ar2 As Object = ws.Range("AR2").Value
        Dim ar7 As Object = ws.Range("AR7").Value
        Dim as7 As Object = ws.Range("AS7").Value
        Dim aw2 As Object = ws.Range("AW2").Value

        ws.Cells.Clear()

        ws.Range("AS2").Value = as2
        ws.Range("AS3").Value = as3
        ws.Range("AS4").Value = as4
        ws.Range("AS5").Value = as5
        ws.Range("AR2").Value = ar2
        ws.Range("AR7").Value = ar7
        ws.Range("AS7").Value = as7
        ws.Range("AW2").Value = aw2

        Dim dateCell As Object = ws.Range("AM4")
        dateCell.Value = DateTime.Now.ToString("dd-MM-yyyy")
        dateCell.Font.Bold = True
        dateCell.HorizontalAlignment = -4108 ' xlCenter
        dateCell.VerticalAlignment = -4108
        dateCell.EntireColumn.ColumnWidth = 12

        Dim aw2Cell As Object = ws.Range("AW2")
        If ws.Range("AW2").Value Is Nothing Then
            aw2Cell.Value = ""
        End If
        aw2Cell.Font.Bold = True
        aw2Cell.HorizontalAlignment = -4108
        aw2Cell.VerticalAlignment = -4108
        aw2Cell.Interior.Color = Color.FromArgb(146, 208, 80)
        aw2Cell.Font.Color = Color.FromArgb(0, 0, 0)
        aw2Cell.Borders.LineStyle = 1

        Dim aw1Cell As Object = ws.Range("AW1")
        aw1Cell.Value = "Frame Analysis:"
        aw1Cell.Font.Bold = True
        aw1Cell.Font.Color = Color.FromArgb(0, 32, 96)
        aw1Cell.HorizontalAlignment = -4131 ' xlLeft
        aw1Cell.VerticalAlignment = -4108
    End Sub

    Private Sub WriteJobInfoToExcel(ws As Object, staadApp As Object)
        ' GetFullJobInfo fills these 13 ByRef params positionally; only the ones actually
        ' used below are named meaningfully (matches original param usage exactly).
        Dim p0_jobName As String = ""
        Dim p1_unused As String = ""
        Dim p2_engineerName As String = ""
        Dim p3_unused As String = ""
        Dim p4_jobNumber As String = ""
        Dim p5_revision As String = ""
        Dim p6_partName As String = ""
        Dim p7_unused As String = ""
        Dim p8_unused As String = ""
        Dim p9_unused As String = ""
        Dim p10_approverName As String = ""
        Dim p11_unused As String = ""
        Dim p12_unused As String = ""

        Dim taglines() As String = {
            "BUILT WITH PURPOSE — RAGHAV", "INNOVATION IN EVERY CELL — RAGHAV", "CRAFTED BY RAGHAV, POWERED BY CODE", "BRINGING LOGIC TO LIFE — RAGHAV", "DESIGNED FOR EXCELLENCE — RAGHAV", "ENGINEERED BY RAGHAV WITH PRECISION", "CREATIVE CODE BY RAGHAV", "WHERE IDEAS MEET EXECUTION — RAGHAV", "LINES OF CODE, LINES OF THOUGHT — RAGHAV", "BY RAGHAV — FOR PERFECTION",
            "CRAFTED WITH PASSION, WRITTEN IN LOGIC — RAGHAV", "THE ART OF AUTOMATION — RAGHAV", "VISION TO VALUE — BY RAGHAV", "ENGINEERING BEYOND LIMITS — RAGHAV", "WHERE LOGIC MEETS AESTHETICS — RAGHAV", "EVERY CELL TELLS A STORY — RAGHAV", "STRUCTURE, STYLE, STABILITY — RAGHAV", "TECH WITH A HUMAN TOUCH — RAGHAV", "SOLUTIONS THAT SPEAK — RAGHAV"
        }
        Dim rnd As New Random()
        Dim tagline As String = taglines(rnd.Next(taglines.Length))

        staadApp.GetFullJobInfo(p0_jobName, p1_unused, p2_engineerName, p3_unused, p4_jobNumber, p5_revision, p6_partName, p7_unused, p8_unused, p9_unused, p10_approverName, p11_unused, p12_unused)

        Dim instance As Object = ws
        instance.Range("C2").Value = "JOB NAME:-"
        instance.Range("C2:D2").Merge()
        instance.Range("E2").Value = p0_jobName
        instance.Range("E2:AQ2").Merge()
        instance.Range("C3").Value = "JOB NUMBER:-"
        instance.Range("C3:D3").Merge()
        instance.Range("E3").Value = p4_jobNumber
        instance.Range("E3:AQ3").Merge()
        instance.Range("C4").Value = "ENGINEER NAME:-"
        instance.Range("C4:D4").Merge()
        instance.Range("E4").Value = p2_engineerName
        instance.Range("AN4").Value = "APPROVER NAME:-"
        instance.Range("AN4:AO4").Merge()
        instance.Range("AP4").Value = p10_approverName
        instance.Range("AP4:AQ4").Merge()
        instance.Range("AN5").Value = "REVISION:-"
        instance.Range("AO5").Value = p5_revision
        instance.Range("AL4").Value = "DATE:-"
        instance.Range("C5").Value = "PART:-"
        instance.Range("C5:D5").Merge()
        instance.Range("E5").Value = p6_partName
        instance.Range("E5:AM5").Merge()
        instance.Range("AR3").Value = "PLACEMENT:-"
        instance.Range("AR4").Value = "TEXT HEIGHT:-"
        instance.Range("AR5").Value = "TEXT COLOR:-"
        instance.Range("A1:AS7").Font.Bold = True
        instance.Range("A1:AS1,A6:AR6").Interior.Color = Color.FromArgb(248, 203, 173)
        instance.Range("AS1,AQ5").Interior.Color = Color.FromArgb(146, 205, 220)
        instance.Range("A1:AQ5").Font.Color = Color.FromArgb(0, 32, 96)
        instance.Range("AR2:AS2").Font.Color = Color.FromArgb(192, 0, 0)
        instance.Range("A1:AQ6,AR2:AS6").Borders.LineStyle = 1
        instance.Range("A1:AQ1,A6:AQ6,AR6:AS6").Merge()
        instance.Range("A2:B5").Merge()
        instance.Range("A1:AQ6").Font.Size = 12
        instance.Range("A1").Value = tagline
        instance.Range("A1").Font.Color = Color.FromArgb(192, 0, 0)
        instance.Range("A1,AN3,AN4,AN5,AQ3,AQ4,AR7,AS7").HorizontalAlignment = -4108
        instance.Range("E3,K5").HorizontalAlignment = -4131
        instance.Range("AR1").Value = "CREATED BY:-"
        instance.Range("AS1").Value = "PEEYUSH"
    End Sub

    Private Sub SetupExcelHeaders(ws As Object)
        Dim obj As Object = ws
        obj.Range("A7").Value = "S.No"
        obj.Range("B7").Value = "BEAM No"
        obj.Range("C7").Value = "LENGTH(M)"
        obj.Range("D7").Value = "SECTION"
        obj.Range("E7").Value = "PROPERTY"
        obj.Range("F7").Value = "TYPE"
        obj.Range("G7").Value = "F1(DEPTH START)"
        obj.Range("H7").Value = "F2(WEB THK.)"
        obj.Range("I7").Value = "F3(DEPTH END)"
        obj.Range("J7").Value = "F4(WIDTH Tf)"
        obj.Range("K7").Value = "F5(THK. Tf)"
        obj.Range("L7").Value = "F6(WIDTH Bf)"
        obj.Range("M7").Value = "F7(THK.Tb)"
        obj.Range("N7").Value = "Node A"
        obj.Range("O7").Value = "Node B"
        obj.Range("P7").Value = "COORDINATE A"
        obj.Range("S7").Value = "COORDINATE B"
        obj.Range("V7").Value = "Beta Angle (°)"
        obj.Range("AK7").Value = "SECTION FARMING"
        obj.Range("AN7").Value = "WEIGHT(Kg)"
        obj.Range("P8").Value = "X"
        obj.Range("Q8").Value = "Y"
        obj.Range("R8").Value = "Z"
        obj.Range("S8").Value = "X"
        obj.Range("T8").Value = "Y"
        obj.Range("U8").Value = "Z"
        obj.Range("AK8").Value = "WEB"
        obj.Range("AL8").Value = "TOP FLANGE"
        obj.Range("AM8").Value = "BOTT.FLANGE"
        obj.Range("AN8").Value = "WEB(Kg)"
        obj.Range("AO8").Value = "TOP FLANGE(Kg)"
        obj.Range("AP8").Value = "BOTT.FLANGE(Kg)"
        obj.Range("AQ8").Value = "TOTAL(Kg)"
        obj.Range("P7:R7").Merge()
        obj.Range("S7:U7").Merge()
        obj.Range("V7:V8").Merge()
        obj.Range("AK7:AM7").Merge()
        obj.Range("AN7:AQ7").Merge()

        Dim headerRange As Object = obj.Range("AK8:AQ8")
        headerRange.Font.Bold = True
        headerRange.HorizontalAlignment = -4108
        headerRange.VerticalAlignment = -4108
        headerRange.WrapText = True

        For col As Integer = 1 To 15
            Dim mergeRange As Object = obj.Range(obj.Cells(7, col), obj.Cells(8, col))
            mergeRange.Merge()
            mergeRange.WrapText = True
            mergeRange.HorizontalAlignment = -4108
            mergeRange.VerticalAlignment = -4108
        Next

        Dim allHeaders As Object = obj.Range("A7:AQ8")
        allHeaders.Font.Bold = True
        allHeaders.HorizontalAlignment = -4108
        allHeaders.VerticalAlignment = -4108

        obj.Range("7:8").Font.Color = RGB(0, 32, 96)
    End Sub

    Private Function ProcessBeamData(ws As Object, staadApp As Object) As String
        Dim beamCount As Integer = CInt(staadApp.Geometry.GetNoOfSelectedBeams())
        If beamCount = 0 Then
            Return "No members selected in STAAD."
        End If

        Dim beamNumbers(beamCount - 1) As Integer
        staadApp.Geometry.GetSelectedBeams(beamNumbers)

        Dim propValueCount As Integer = CInt(staadApp.Property.GetCountofSectionPropertyValuesEx())
        Dim propValues(propValueCount - 1) As Double
        Dim outputData(beamCount - 1, 21) As Object
        Dim writtenCount As Integer = 0
        Dim maxPropLength As Integer = 0

        For Each beamNo As Integer In beamNumbers
            Try
                Dim length As Double = CDbl(staadApp.Geometry.GetBeamLength(beamNo))
                Dim sectionName As String = CStr(staadApp.Property.GetBeamSectionName(beamNo))
                Dim refNo As Integer = CInt(staadApp.Property.GetBeamSectionPropertyRefNo(beamNo))
                Dim typeNo As Integer = CInt(staadApp.Property.GetBeamSectionPropertyTypeNo(beamNo))
                Dim beta As Double = CDbl(staadApp.Property.GetBetaAngle(beamNo))

                Dim nodeA As Integer = 0, nodeB As Integer = 0
                staadApp.Geometry.GetMemberIncidence(beamNo, nodeA, nodeB)

                Dim xa As Double = 0, ya As Double = 0, za As Double = 0
                staadApp.Geometry.GetNodeCoordinates(nodeA, xa, ya, za)

                Dim xb As Double = 0, yb As Double = 0, zb As Double = 0
                staadApp.Geometry.GetNodeCoordinates(nodeB, xb, yb, zb)

                staadApp.Property.GetSectionPropertyValuesEx(refNo, typeNo, propValues)

                If propValues.Length > maxPropLength Then maxPropLength = propValues.Length

                outputData(writtenCount, 0) = writtenCount + 1
                outputData(writtenCount, 1) = beamNo
                outputData(writtenCount, 2) = Math.Round(length, 3)
                outputData(writtenCount, 3) = sectionName
                outputData(writtenCount, 4) = "R" & refNo.ToString()
                outputData(writtenCount, 5) = typeNo

                Dim upperF As Integer = Math.Min(6, propValues.Length - 1)
                For j As Integer = 0 To upperF
                    If typeNo = 680 Then
                        outputData(writtenCount, 6 + j) = Math.Round(propValues(j), 3)
                    Else
                        outputData(writtenCount, 6 + j) = Math.Round(propValues(j), 6)
                    End If
                Next

                outputData(writtenCount, 13) = nodeA
                outputData(writtenCount, 14) = nodeB
                outputData(writtenCount, 15) = Math.Round(xa, 3)
                outputData(writtenCount, 16) = Math.Round(ya, 3)
                outputData(writtenCount, 17) = Math.Round(za, 3)
                outputData(writtenCount, 18) = Math.Round(xb, 3)
                outputData(writtenCount, 19) = Math.Round(yb, 3)
                outputData(writtenCount, 20) = Math.Round(zb, 3)
                outputData(writtenCount, 21) = Math.Round(beta, 2)

                If propValues.Length > 7 Then
                    Dim upperExtra As Integer = Math.Min(11, propValues.Length - 1)
                    For k As Integer = 7 To upperExtra
                        Dim col As Integer = 51 + (k - 7)
                        If typeNo = 680 Then
                            ws.Cells(9 + writtenCount, col).Value = Math.Round(propValues(k), 3)
                        Else
                            ws.Cells(9 + writtenCount, col).Value = Math.Round(propValues(k), 6)
                        End If
                    Next
                End If

                writtenCount += 1
            Catch ex As Exception
                Debug.WriteLine($"Error processing beam {beamNo}: {ex.Message}")
            End Try
        Next

        If writtenCount > 0 Then
            Dim targetRange As Object = ws.Range(ws.Cells(9, 1), ws.Cells(8 + writtenCount, 22))
            targetRange.Value = outputData
            For l As Integer = 9 To 8 + writtenCount
                If Val(ws.Cells(l, 6).Value) <> 680.0 Then Continue For
                If Val(ws.Cells(l, 12).Value) <> 0.0 Then Continue For
                If Val(ws.Cells(l, 13).Value) = 0.0 Then
                    ws.Cells(l, 12).Value = ws.Cells(l, 10).Value
                    ws.Cells(l, 13).Value = ws.Cells(l, 11).Value
                End If
            Next
        End If

        Return $"{writtenCount} of {beamCount} beams processed successfully."
    End Function

    Private Sub FormatExcelWorksheet(ws As Object)
        Dim columnWidths() As Double = {
            6.0, 10.0, 12.0, 15.0, 10.0, 8.0, 12.0, 12.0, 12.0, 12.0,
            12.0, 12.0, 12.0, 10.0, 10.0, 10.0, 10.0, 10.0, 10.0, 10.0,
            10.0, 10.0, 10.0, 14.0, 14.0, 14.0, 14.0, 14.0, 14.0, 14.0,
            14.0, 14.0, 14.0, 14.0, 14.0, 14.0, 14.0, 14.0, 14.0
        }
        For i As Integer = 0 To columnWidths.Length - 1
            ws.Columns(i + 1).ColumnWidth = columnWidths(i)
        Next

        Dim lastRow As Integer = CInt(ws.Cells(ws.Rows.Count, 1).End(-4162).Row)
        If lastRow >= 9 Then
            Dim borderRange As Object = ws.Range(ws.Cells(7, 1), ws.Cells(lastRow, 22))
            Dim borders As Object = borderRange.Borders
            borders.LineStyle = 1
            borders.Weight = 2
        End If

        Dim as2to6 As Object = ws.Range("AS2:AS6")
        as2to6.HorizontalAlignment = -4108
        as2to6.VerticalAlignment = -4108
        ws.Columns("AS").AutoFit()

        Dim arBorders As Object = ws.Range("AR1:AS6").Borders
        arBorders.LineStyle = 1
        arBorders.Weight = 2
    End Sub

    Private Sub HideColumns(ws As Object)
        ws.Range("A:E").EntireColumn.Hidden = False
        ws.Range("F:AJ").EntireColumn.Hidden = True
        ws.Range("AK:AW").EntireColumn.Hidden = False
        ws.Range("AY:BC").EntireColumn.Hidden = True
    End Sub

    Private Sub CalculateTotalWeight(ws As Object)
        Dim lastRow As Integer = CInt(ws.Cells(ws.Rows.Count, "AQ").End(-4162).Row)
        Dim total As Double = 0.0
        If lastRow >= 9 Then
            total = CDbl(ws.Application.WorksheetFunction.Sum(ws.Range("AQ9:AQ" & lastRow.ToString())))
        End If
        ws.Range("AQ5").Value = Math.Round(total, 2)
        ws.Range("AP5").Value = "TOTAL WEIGHT:"
        ws.Range("I5:K5").Font.Bold = True
    End Sub

    Private Sub ProcessSectionFarmingAndWeight(ws As Object)
        Dim lastRow As Integer = CInt(ws.Cells(ws.Rows.Count, 7).End(-4162).Row)
        Dim clearRange As Object = ws.Range(ws.Cells(7, 24), ws.Cells(ws.Rows.Count, 50))
        clearRange.ClearContents()

        Dim obj As Object = ws
        obj.Cells(7, 24).Value = "NODE A"
        obj.Cells(7, 30).Value = "NODE B"
        obj.Range("AK7").Value = "SECTION FARMING"
        obj.Range("AN7").Value = "WEIGHT(Kg)"
        obj.Range("X8").Value = "WEB WIDTH"
        obj.Range("Y8").Value = "WEB THK."
        obj.Range("Z8").Value = "Tf WIDTH"
        obj.Range("AA8").Value = "Tf THK."
        obj.Range("AB8").Value = "Bf WIDTH"
        obj.Range("AC8").Value = "Bf THK."
        obj.Range("AD8").Value = "WEB WIDTH"
        obj.Range("AE8").Value = "WEB THK."
        obj.Range("AF8").Value = "Tf WIDTH"
        obj.Range("AG8").Value = "Tf THK."
        obj.Range("AH8").Value = "Bf WIDTH"
        obj.Range("AI8").Value = "Bf THK."
        obj.Range("AJ7").Value = "AVG.WEB"
        obj.Range("AK8").Value = "WEB"
        obj.Range("AL8").Value = "TOP FLANGE"
        obj.Range("AM8").Value = "BOTT.FLANGE"
        obj.Range("AN8").Value = "WEB(Kg)"
        obj.Range("AO8").Value = "TOP FLANGE(Kg)"
        obj.Range("AP8").Value = "BOTT.FLANGE(Kg)"
        obj.Range("AQ8").Value = "TOTAL(Kg)"
        obj.Range("E4:AK4").Merge()
        obj.Range("X7:AC7").Merge()
        obj.Range("AD7:AI7").Merge()
        obj.Range("AJ7:AJ8").Merge()
        obj.Range("AK7:AM7").Merge()
        obj.Range("AN7:AQ7").Merge()

        For i As Integer = 9 To lastRow
            Dim lengthM As Double = Val(ws.Cells(i, 3).Value)
            Dim sectionType As String = CStr(ws.Cells(i, 4).Value).Trim().ToUpper()
            Dim typeNo As Integer = CInt(Math.Round(Val(ws.Cells(i, 6).Value)))

            If typeNo = 671 OrElse typeNo = 672 Then
                ws.Cells(i, 43).Value = 0
                ws.Cells(i, 37).Value = ws.Cells(i, 4).Value
                ws.Cells(i, 38).Value = ""
                ws.Cells(i, 39).Value = ""
            Else
                If typeNo = 680 Then
                    If Val(ws.Cells(i, 13).Value) = 0.0 Then
                        ws.Cells(i, 13).Value = ws.Cells(i, 11).Value
                    End If
                    If Val(ws.Cells(i, 12).Value) = 0.0 Then
                        ws.Cells(i, 12).Value = ws.Cells(i, 10).Value
                    End If
                End If

                If typeNo = 697 Then
                    Dim f2 As Double = Val(ws.Cells(i, 8).Value)
                    Dim f3 As Double = Val(ws.Cells(i, 9).Value)
                    Dim f4 As Double = Val(ws.Cells(i, 10).Value)
                    Dim f5 As Double = Val(ws.Cells(i, 11).Value)
                    If f2 = f4 Then
                        Dim f1 As Double = Val(ws.Cells(i, 7).Value)
                        ws.Cells(i, 43).Value = Math.Round(lengthM * f1 * 7850.0, 3)
                        ws.Cells(i, 37).Value = "UPT"
                        ws.Cells(i, 38).Value = ""
                        ws.Cells(i, 39).Value = ""
                    Else
                        Dim webWidthMm As Integer = CInt(Math.Round(1000.0 * f2 - 2000.0 * f5))
                        Dim webThkMm As Integer = CInt(Math.Round(1000.0 * f3))
                        Dim flangeWidthMm As Integer = CInt(Math.Round(1000.0 * f4))
                        Dim flangeThkMm As Integer = CInt(Math.Round(1000.0 * f5))
                        ws.Cells(i, 37).Value = "2x" & webWidthMm.ToString() & "x" & webThkMm.ToString()
                        ws.Cells(i, 38).Value = "2x" & flangeWidthMm.ToString() & "x" & flangeThkMm.ToString()
                        ws.Cells(i, 39).Value = "2x" & flangeWidthMm.ToString() & "x" & flangeThkMm.ToString()
                        Dim topFlangeWeight As Double = Math.Round(2.0 * f4 * f5 * lengthM * 7.85 * 1000.0, 3)
                        ws.Cells(i, 41).Value = topFlangeWeight
                        Dim bottomFlangeWeight As Double = Math.Round(2.0 * f4 * f5 * lengthM * 7.85 * 1000.0, 3)
                        ws.Cells(i, 42).Value = bottomFlangeWeight
                        Dim f1Total As Double = Val(ws.Cells(i, 7).Value)
                        Dim totalWeight As Double = Math.Round(lengthM * f1Total * 7850.0, 3)
                        ws.Cells(i, 43).Value = totalWeight
                        Dim webWeight As Double = Math.Round(totalWeight - (topFlangeWeight + bottomFlangeWeight), 3)
                        ws.Cells(i, 40).Value = webWeight
                    End If
                ElseIf typeNo >= 690 AndAlso typeNo <= 699 AndAlso typeNo <> 697 Then
                    Dim f1 As Double = Val(ws.Cells(i, 7).Value)
                    ws.Cells(i, 43).Value = Math.Round(lengthM * f1 * 7850.0, 3)
                    ws.Cells(i, 37).Value = "UPT"
                    ws.Cells(i, 38).Value = ""
                    ws.Cells(i, 39).Value = ""
                ElseIf typeNo = 613 OrElse typeNo = 614 OrElse typeNo = 615 Then
                    Dim props(7) As Double
                    For colOffset As Integer = 1 To 7
                        props(colOffset) = Val(ws.Cells(i, 6 + colOffset).Value)
                    Next
                    For col As Integer = 24 To 35
                        ws.Cells(i, col).Value = 0
                    Next
                    ws.Cells(i, 36).Value = 0
                    ws.Cells(i, 37).Value = ws.Cells(i, 4).Value

                    Dim f5 As Double = Val(ws.Cells(i, 52).Value)
                    Dim f6 As Double = Val(ws.Cells(i, 53).Value)
                    Dim f7 As Double = Val(ws.Cells(i, 54).Value)
                    Dim f8 As Double = Val(ws.Cells(i, 55).Value)

                    Select Case typeNo
                        Case 613
                            ws.Cells(i, 26).Value = Math.Round(f5 * 1000.0)
                            ws.Cells(i, 27).Value = Math.Round(f6 * 1000.0)
                            ws.Cells(i, 28).Value = 0
                            ws.Cells(i, 29).Value = 0
                            ws.Cells(i, 32).Value = Math.Round(f5 * 1000.0)
                            ws.Cells(i, 33).Value = Math.Round(f6 * 1000.0)
                            ws.Cells(i, 34).Value = 0
                            ws.Cells(i, 35).Value = 0
                        Case 614
                            ws.Cells(i, 26).Value = 0
                            ws.Cells(i, 27).Value = 0
                            ws.Cells(i, 28).Value = Math.Round(f5 * 1000.0)
                            ws.Cells(i, 29).Value = Math.Round(f6 * 1000.0)
                            ws.Cells(i, 32).Value = 0
                            ws.Cells(i, 33).Value = 0
                            ws.Cells(i, 34).Value = Math.Round(f5 * 1000.0)
                            ws.Cells(i, 35).Value = Math.Round(f6 * 1000.0)
                        Case 615
                            Dim webW As Double = f5, webT As Double = f6, flangeW As Double = f7, flangeT As Double = f8
                            If (webW = 0.0 OrElse webT = 0.0) AndAlso flangeW > 0.0 AndAlso flangeT > 0.0 Then
                                webW = flangeW
                                webT = flangeT
                            End If
                            If (flangeW = 0.0 OrElse flangeT = 0.0) AndAlso webW > 0.0 AndAlso webT > 0.0 Then
                                flangeW = webW
                                flangeT = webT
                            End If
                            ws.Cells(i, 26).Value = Math.Round(webW * 1000.0)
                            ws.Cells(i, 27).Value = Math.Round(webT * 1000.0)
                            ws.Cells(i, 28).Value = Math.Round(flangeW * 1000.0)
                            ws.Cells(i, 29).Value = Math.Round(flangeT * 1000.0)
                            ws.Cells(i, 32).Value = Math.Round(webW * 1000.0)
                            ws.Cells(i, 33).Value = Math.Round(webT * 1000.0)
                            ws.Cells(i, 34).Value = Math.Round(flangeW * 1000.0)
                            ws.Cells(i, 35).Value = Math.Round(flangeT * 1000.0)
                    End Select

                    Select Case typeNo
                        Case 613
                            ws.Cells(i, 38).Value = CStr(ws.Cells(i, 26).Value) & "x" & CStr(ws.Cells(i, 27).Value)
                            ws.Cells(i, 39).Value = ""
                        Case 614
                            ws.Cells(i, 38).Value = ""
                            ws.Cells(i, 39).Value = CStr(ws.Cells(i, 28).Value) & "x" & CStr(ws.Cells(i, 29).Value)
                        Case 615
                            ws.Cells(i, 38).Value = CStr(ws.Cells(i, 26).Value) & "x" & CStr(ws.Cells(i, 27).Value)
                            ws.Cells(i, 39).Value = CStr(ws.Cells(i, 28).Value) & "x" & CStr(ws.Cells(i, 29).Value)
                    End Select

                    Dim webWeight2 As Double = 0.0
                    Dim flangeWeight2 As Double = 0.0
                    Select Case typeNo
                        Case 613
                            Dim partial As Double = lengthM * Val(ws.Cells(i, 26).Value)
                            webWeight2 = Math.Round(partial * Val(ws.Cells(i, 27).Value) * 7.85 / 1000.0, 3)
                        Case 614
                            Dim partial As Double = lengthM * Val(ws.Cells(i, 28).Value)
                            flangeWeight2 = Math.Round(partial * Val(ws.Cells(i, 29).Value) * 7.85 / 1000.0, 3)
                        Case 615
                            Dim partialWeb As Double = lengthM * Val(ws.Cells(i, 26).Value)
                            webWeight2 = Math.Round(partialWeb * Val(ws.Cells(i, 27).Value) * 7.85 / 1000.0, 3)
                            Dim partialFlange As Double = lengthM * Val(ws.Cells(i, 28).Value)
                            flangeWeight2 = Math.Round(partialFlange * Val(ws.Cells(i, 29).Value) * 7.85 / 1000.0, 3)
                    End Select

                    Dim totalWeight2 As Double = Math.Round(lengthM * Val(ws.Cells(i, 7).Value) * 7850.0, 3)
                    Dim remainderWeight As Double = Math.Round(totalWeight2 - webWeight2 - flangeWeight2, 3)
                    ws.Cells(i, 40).Value = remainderWeight
                    ws.Cells(i, 41).Value = webWeight2
                    ws.Cells(i, 42).Value = flangeWeight2
                    ws.Cells(i, 43).Value = totalWeight2
                Else
                    Dim props(7) As Double
                    For colOffset As Integer = 1 To 7
                        props(colOffset) = Val(ws.Cells(i, 6 + colOffset).Value)
                    Next
                    If typeNo = 680 Then
                        ws.Cells(i, 24).Value = Math.Round((props(1) - props(5) - props(7)) * 1000.0)
                        ws.Cells(i, 25).Value = Math.Round(props(2) * 1000.0)
                        ws.Cells(i, 26).Value = Math.Round(props(4) * 1000.0)
                        ws.Cells(i, 27).Value = Math.Round(props(5) * 1000.0)
                        ws.Cells(i, 28).Value = Math.Round(props(6) * 1000.0)
                        ws.Cells(i, 29).Value = Math.Round(props(7) * 1000.0)
                        ws.Cells(i, 30).Value = Math.Round((props(3) - props(5) - props(7)) * 1000.0)
                        ws.Cells(i, 31).Value = Math.Round(props(2) * 1000.0)
                        ws.Cells(i, 32).Value = Math.Round(props(4) * 1000.0)
                        ws.Cells(i, 33).Value = Math.Round(props(5) * 1000.0)
                        ws.Cells(i, 34).Value = Math.Round(props(6) * 1000.0)
                        ws.Cells(i, 35).Value = Math.Round(props(7) * 1000.0)
                    Else
                        For col As Integer = 24 To 35
                            ws.Cells(i, col).Value = 0
                        Next
                    End If

                    Dim webWidth As Double = Val(ws.Cells(i, 24).Value)
                    Dim bottomWebWidth As Double = Val(ws.Cells(i, 30).Value)
                    Dim webThk As Double = Val(ws.Cells(i, 25).Value)
                    ws.Cells(i, 36).Value = Math.Round((webWidth + bottomWebWidth) / 2.0)

                    If webWidth = 0.0 OrElse bottomWebWidth = 0.0 OrElse webThk = 0.0 Then
                        ws.Cells(i, 37).Value = ws.Cells(i, 4).Value
                        ws.Cells(i, 38).Value = ""
                        ws.Cells(i, 39).Value = ""
                    Else
                        If webWidth = bottomWebWidth Then
                            ws.Cells(i, 37).Value = CStr(webWidth) & "x" & CStr(webThk)
                        Else
                            ws.Cells(i, 37).Value = CStr(webWidth) & "-" & CStr(bottomWebWidth) & "x" & CStr(webThk)
                        End If
                        ws.Cells(i, 38).Value = CStr(ws.Cells(i, 26).Value) & "x" & CStr(ws.Cells(i, 27).Value)
                        ws.Cells(i, 39).Value = CStr(ws.Cells(i, 28).Value) & "x" & CStr(ws.Cells(i, 29).Value)
                    End If

                    Dim avgWebWeight As Double = Math.Round(lengthM * Val(ws.Cells(i, 36).Value) * webThk * 7.85 / 1000.0, 3)
                    Dim topFlangePartial As Double = lengthM * Val(ws.Cells(i, 26).Value)
                    Dim topFlangeWeight2 As Double = Math.Round(topFlangePartial * Val(ws.Cells(i, 27).Value) * 7.85 / 1000.0, 3)
                    Dim bottomFlangePartial As Double = lengthM * Val(ws.Cells(i, 28).Value)
                    Dim bottomFlangeWeight2 As Double = Math.Round(bottomFlangePartial * Val(ws.Cells(i, 29).Value) * 7.85 / 1000.0, 3)
                    ws.Cells(i, 40).Value = avgWebWeight
                    ws.Cells(i, 41).Value = topFlangeWeight2
                    ws.Cells(i, 42).Value = bottomFlangeWeight2
                    ws.Cells(i, 43).Value = Math.Round(avgWebWeight + topFlangeWeight2 + bottomFlangeWeight2, 3)
                End If
            End If

            If Val(ws.Cells(i, 24).Value) = 0.0 Then
                If Val(ws.Cells(i, 43).Value) = 0.0 AndAlso typeNo <> 671 AndAlso typeNo <> 672 Then
                    Dim fallbackWeight As Double = lengthM * Val(ws.Cells(i, 7).Value) * 7850.0
                    ws.Cells(i, 43).Value = Math.Round(fallbackWeight, 3)
                End If
            End If

            If sectionType = "TUBE" OrElse sectionType = "USER DEFINED TUBE" Then
                Dim thk As Double = Val(ws.Cells(i, 11).Value)
                Dim outerA As Double = Val(ws.Cells(i, 12).Value)
                Dim outerB As Double = Val(ws.Cells(i, 13).Value)
                Dim depthMm As Double = (outerA + outerB - thk * 2.0) * 2.0 * 1000.0
                Dim thkMm As Double = thk * 1000.0
                ws.Cells(i, 36).Value = Math.Round(depthMm)
                ws.Cells(i, 25).Value = Math.Round(thkMm)
                Dim tubeWeight As Double = Math.Round(lengthM * depthMm * thkMm * 7.85 / 1000.0, 3)
                ws.Cells(i, 40).Value = tubeWeight
                ws.Cells(i, 43).Value = tubeWeight
                ws.Cells(i, 37).Value = "TUBE " & CStr(Math.Round(outerA * 1000.0)) & "x" & CStr(Math.Round(outerB * 1000.0)) & "x" & CStr(Math.Round(thk * 1000.0, 1))
            End If

            If sectionType <> "PIPE" AndAlso sectionType <> "USER DEFINED PIPE" Then Continue For

            Dim outerDia As Double = Val(ws.Cells(i, 11).Value)
            Dim innerDia As Double = Val(ws.Cells(i, 12).Value)
            Dim area As Double = Math.PI / 4.0 * (outerDia * outerDia - innerDia * innerDia)
            Dim pipeWeight As Double = Math.Round(lengthM * area * 7850.0, 3)
            ws.Cells(i, 43).Value = pipeWeight
            If innerDia = 0.0 Then
                ws.Cells(i, 37).Value = "ROD " & CStr(Math.Round(outerDia * 1000.0))
            Else
                ws.Cells(i, 37).Value = "PIPE " & CStr(Math.Round(outerDia * 1000.0)) & "*" & CStr(Math.Round((outerDia - innerDia) * 1000.0 / 2.0))
            End If
            ws.Cells(i, 38).Value = ""
            ws.Cells(i, 39).Value = ""
        Next

        ws.Range("X7:AQ" & lastRow.ToString()).Borders.LineStyle = 1
        ws.Range("X:AQ").Columns.AutoFit()
    End Sub

    Private Sub UnprotectSheetAndWorkbook(ws As Object)
        On Error Resume Next
        Dim wb As Object = ws.Parent
        ws.Unprotect("2022")
        wb.Unprotect("2022")
    End Sub

    Private Sub ProtectSheetAndWorkbook(ws As Object)
        On Error Resume Next
        Dim wb As Object = ws.Parent
        ws.Unprotect("2022")
        wb.Unprotect("2022")
        ws.Cells.Locked = True
        ws.Range("AS2:AS5,AR2,AW2").Locked = False
        ws.Protect("2022", DrawingObjects:=True, Contents:=True, Scenarios:=True,
                   AllowFormattingCells:=False, AllowFormattingColumns:=False, AllowFormattingRows:=False,
                   AllowInsertingColumns:=False, AllowInsertingRows:=False, AllowDeletingColumns:=False,
                   AllowDeletingRows:=False, AllowSorting:=False, AllowFiltering:=False, AllowUsingPivotTables:=False)
        wb.Protect("2022", Structure:=True, Windows:=False)
        If Not wb.ProtectStructure Then
            Throw New Exception("Workbook structure protection failed — cannot lock sheet add/delete.")
        End If
    End Sub

End Class
