Option Strict Off

Imports System
Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.Globalization
Imports System.IO
Imports System.Linq
Imports System.Runtime.InteropServices
Imports System.Text.RegularExpressions
Imports System.Windows.Forms
Imports Microsoft.Office.Interop.Excel
Imports netDxf
Imports netDxf.Entities
Imports netDxf.Header
Imports netDxf.Tables

<ComVisible(True)>
<Guid("BDF3BADA-1234-4567-9ABC-ABCDEF012345")>
<ProgId("RaghavStaadExtractor.DxfExporterRaghav")>
Public Class DxfExporterRaghav

    Private Const SCALING_FACTOR As Double = 1000.0

    Private ReadOnly memberColors() As AciColor
    Private ReadOnly tubeColors() As AciColor
    Private gridMemberSignatures As New Dictionary(Of String, String)
    Private gridColorMap As New Dictionary(Of String, AciColor)
    Private masterGridColors() As AciColor
    Private magentaColor As AciColor
    Private gridColorIndex As Integer

    Public Sub New()
        memberColors = New AciColor() {
            AciColor.FromCadIndex(1), AciColor.FromCadIndex(2), AciColor.FromCadIndex(3), AciColor.FromCadIndex(4),
            AciColor.FromCadIndex(5), AciColor.FromCadIndex(6), AciColor.FromCadIndex(7), AciColor.FromCadIndex(10),
            AciColor.FromCadIndex(11), AciColor.FromCadIndex(20), AciColor.FromCadIndex(21), AciColor.FromCadIndex(30),
            AciColor.FromCadIndex(31), AciColor.FromCadIndex(40), AciColor.FromCadIndex(41), AciColor.FromCadIndex(50),
            AciColor.FromCadIndex(51), AciColor.FromCadIndex(60), AciColor.FromCadIndex(61), AciColor.FromCadIndex(70),
            AciColor.FromCadIndex(71), AciColor.FromCadIndex(80), AciColor.FromCadIndex(81), AciColor.FromCadIndex(90),
            AciColor.FromCadIndex(91), AciColor.FromCadIndex(100), AciColor.FromCadIndex(101), AciColor.FromCadIndex(110),
            AciColor.FromCadIndex(111), AciColor.FromCadIndex(120), AciColor.FromCadIndex(121), AciColor.FromCadIndex(130),
            AciColor.FromCadIndex(131), AciColor.FromCadIndex(140), AciColor.FromCadIndex(141), AciColor.FromCadIndex(150),
            AciColor.FromCadIndex(151), AciColor.FromCadIndex(160), AciColor.FromCadIndex(161), AciColor.FromCadIndex(170),
            AciColor.FromCadIndex(171), AciColor.FromCadIndex(180), AciColor.FromCadIndex(181), AciColor.FromCadIndex(190),
            AciColor.FromCadIndex(191), AciColor.FromCadIndex(200), AciColor.FromCadIndex(201), AciColor.FromCadIndex(210),
            AciColor.FromCadIndex(211), AciColor.FromCadIndex(220), AciColor.FromCadIndex(221), AciColor.FromCadIndex(230),
            AciColor.FromCadIndex(231), AciColor.FromCadIndex(240), AciColor.FromCadIndex(241)
        }
        tubeColors = New AciColor() {AciColor.Yellow, AciColor.Cyan, AciColor.Magenta, AciColor.Green, AciColor.Red}
        gridMemberSignatures = New Dictionary(Of String, String)()
        gridColorMap = New Dictionary(Of String, AciColor)()
        masterGridColors = New AciColor() {
            AciColor.Red, AciColor.Blue, AciColor.Green, AciColor.Yellow, AciColor.Cyan,
            AciColor.FromCadIndex(30), AciColor.FromCadIndex(40), AciColor.FromCadIndex(50)
        }
        magentaColor = AciColor.Magenta
        gridColorIndex = 0
    End Sub

    ' Main entry point: reads beam geometry/section data from the active sheet and writes a DXF file
    Public Sub ExportBeamsToDxfWithPrompt(excelSheet As Object)
        Try
            If excelSheet Is Nothing Then
                MessageBox.Show("Excel sheet object is not valid.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Hand)
                Return
            End If
            If Not HasValidBeamData(excelSheet) Then
                MessageBox.Show("Pick a member in STAAD, then continue.", "No Beam Data Found", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Return
            End If

            Dim rotationValue As String = GetSafeCellValueAsString(excelSheet, 2, 45)
            If String.IsNullOrEmpty(rotationValue) Then rotationValue = "1"

            Dim textHeight As Double = GetSafeCellValue(excelSheet, 4, 45)
            If textHeight <= 0.0 Then textHeight = 150.0

            Dim viewMode As Integer = CInt(Math.Round(GetSafeCellValue(excelSheet, 5, 45)))
            Dim defaultFileName As String = GetSafeCellValueAsString(excelSheet, 2, 44).Trim()
            If String.IsNullOrWhiteSpace(defaultFileName) Then defaultFileName = "BeamsOutput"

            Dim dxfBaseName As String = ShowProfessionalInputDialog("Enter DXF File Name (without extension):", "Save DXF File", defaultFileName)
            If String.IsNullOrWhiteSpace(dxfBaseName) Then Return

            Dim workbookPath As String = CStr(CType(excelSheet, Worksheet).Parent.FullName)
            Dim workbookDir As String = Path.GetDirectoryName(workbookPath)
            Dim drawingsDir As String = Path.Combine(workbookDir, "PlateNovaDrawings")
            Directory.CreateDirectory(drawingsDir)
            Dim dxfFilePath As String = Path.Combine(drawingsDir, dxfBaseName & ".dxf")
            Dim dupIndex As Integer = 1
            While File.Exists(dxfFilePath)
                dxfFilePath = Path.Combine(drawingsDir, dxfBaseName & "_" & CStr(dupIndex) & ".dxf")
                dupIndex += 1
            End While

            Dim dxfDocument As New DxfDocument(DxfVersion.AutoCad2010)
            Dim beamsLayer As New Layer("Beams")
            dxfDocument.Layers.Add(beamsLayer)
            Dim leaderLinesLayer As New Layer("LeaderLines")
            dxfDocument.Layers.Add(leaderLinesLayer)
            Dim beamTextLayer As New Layer("BeamText")
            dxfDocument.Layers.Add(beamTextLayer)

            Dim gridWasDrawn As Boolean = AddGridToDxf(dxfDocument, excelSheet, rotationValue)
            Dim leaderDirectionAngle As Double? = GetLeaderDirectionAngle(excelSheet)
            Dim lastRow As Integer = CInt(CType(excelSheet, Worksheet).Cells(CType(excelSheet, Worksheet).Rows.Count, 1).End(XlDirection.xlUp).Row)

            Dim useColorLayers As Boolean = False
            Dim useColorLayersNoLength As Boolean = False
            Dim singleLineMode As Boolean = False
            Dim viewModeValue As Double = GetSafeCellValue(excelSheet, 5, 45)
            If viewModeValue = 100.0 Then
                useColorLayers = True
            ElseIf viewModeValue = 101.0 Then
                useColorLayers = True
                useColorLayersNoLength = True
            ElseIf viewModeValue = 200.0 Then
                useColorLayers = True
                singleLineMode = True
            End If

            Dim colorLayerMap As Dictionary(Of String, Layer) = Nothing
            If useColorLayers Then
                colorLayerMap = CreateColorBasedLayers(dxfDocument, excelSheet, lastRow)
            End If

            For i As Integer = 9 To lastRow
                Try
                    Dim x1 As Double = GetSafeCellValue(excelSheet, i, 16) * 1000.0
                    Dim y1 As Double = GetSafeCellValue(excelSheet, i, 17) * 1000.0
                    Dim z1 As Double = GetSafeCellValue(excelSheet, i, 18) * 1000.0
                    Dim x2 As Double = GetSafeCellValue(excelSheet, i, 19) * 1000.0
                    Dim y2 As Double = GetSafeCellValue(excelSheet, i, 20) * 1000.0
                    Dim z2 As Double = GetSafeCellValue(excelSheet, i, 21) * 1000.0
                    If x1 = 0.0 AndAlso y1 = 0.0 AndAlso z1 = 0.0 AndAlso x2 = 0.0 AndAlso y2 = 0.0 AndAlso z2 = 0.0 Then
                        Continue For
                    End If

                    Dim startPoint As Vector3 = ApplyRotation(New Vector3(x1, y1, z1), rotationValue)
                    Dim endPoint As Vector3 = ApplyRotation(New Vector3(x2, y2, z2), rotationValue)
                    Dim betaAngle As Double = GetSafeCellValue(excelSheet, i, 22)
                    Dim memberLayer As Layer = beamsLayer
                    Dim lineColor As AciColor
                    Dim sectionColor As AciColor

                    If useColorLayers AndAlso colorLayerMap IsNot Nothing Then
                        memberLayer = GetBeamLayerByParameters(excelSheet, i, colorLayerMap, beamsLayer)
                        lineColor = memberLayer.Color
                        sectionColor = lineColor
                    Else
                        lineColor = memberColors((i - 9) Mod memberColors.Length)
                        sectionColor = tubeColors((i - 9) Mod tubeColors.Length)
                    End If

                    Dim beamLine As New Line(startPoint, endPoint) With {.Layer = memberLayer, .Color = lineColor}
                    dxfDocument.Entities.Add(beamLine)

                    If Not singleLineMode Then
                        Dim sectionType As Double = GetSafeCellValue(excelSheet, i, 6)
                        If sectionType = 655.0 OrElse sectionType = 660.0 Then
                            Dim outerDiameter As Double = GetSafeCellValue(excelSheet, i, 8) * 1000.0
                            Dim thickness As Double = GetSafeCellValue(excelSheet, i, 9) * 1000.0
                            DrawPipeSection(dxfDocument, startPoint, endPoint, outerDiameter, thickness, sectionColor, betaAngle, rotationValue, memberLayer)
                        ElseIf sectionType = 651.0 Then
                            Dim width As Double = GetSafeCellValue(excelSheet, i, 12) * 1000.0
                            Dim thickness As Double = GetSafeCellValue(excelSheet, i, 11) * 1000.0
                            DrawTubeSection(dxfDocument, startPoint, endPoint, width, thickness, sectionColor, betaAngle, rotationValue, memberLayer)
                        ElseIf sectionType = 654.0 OrElse sectionType = 650.0 Then
                            Dim width As Double = GetSafeCellValue(excelSheet, i, 8) * 1000.0
                            Dim height As Double = GetSafeCellValue(excelSheet, i, 9) * 1000.0
                            Dim thickness As Double = GetSafeCellValue(excelSheet, i, 10) * 1000.0
                            DrawHollowRectangleTube(dxfDocument, startPoint, endPoint, width, height, thickness, sectionColor, betaAngle, rotationValue, memberLayer)
                        ElseIf sectionType = 672.0 Then
                            Dim width As Double = GetSafeCellValue(excelSheet, i, 11) * 1000.0
                            Dim height As Double = GetSafeCellValue(excelSheet, i, 12) * 1000.0
                            Dim thickness As Double = 0.0
                            DrawHollowRectangleTube(dxfDocument, startPoint, endPoint, width, height, thickness, lineColor, betaAngle, rotationValue, memberLayer)
                        ElseIf sectionType = 697.0 Then
                            Dim d1 As Double = GetSafeCellValue(excelSheet, i, 8) * 1000.0
                            Dim d2 As Double = GetSafeCellValue(excelSheet, i, 9) * 1000.0
                            Dim d3 As Double = GetSafeCellValue(excelSheet, i, 10) * 1000.0
                            Dim d4 As Double = GetSafeCellValue(excelSheet, i, 11) * 1000.0
                            If Math.Abs(d1 - d3) > 0.001 Then
                                Dim webThickness As Double = d1 - 2.0 * d4
                                DrawTaperedISection(dxfDocument, startPoint, endPoint, webThickness, d2, d3, d4, d3, d4, webThickness, d2, d3, d4, d3, d4, lineColor, betaAngle, rotationValue, memberLayer)
                                Dim betaAngle2 As Double = betaAngle + 90.0
                                DrawTaperedISection(dxfDocument, startPoint, endPoint, webThickness, d2, d3, d4, d3, d4, webThickness, d2, d3, d4, d3, d4, lineColor, betaAngle2, rotationValue, memberLayer)
                            End If
                        ElseIf sectionType = 680.0 Then
                            DrawTaperedISection(dxfDocument, startPoint, endPoint,
                                GetSafeCellValue(excelSheet, i, 24), GetSafeCellValue(excelSheet, i, 25), GetSafeCellValue(excelSheet, i, 26), GetSafeCellValue(excelSheet, i, 27),
                                GetSafeCellValue(excelSheet, i, 28), GetSafeCellValue(excelSheet, i, 29), GetSafeCellValue(excelSheet, i, 30), GetSafeCellValue(excelSheet, i, 31),
                                GetSafeCellValue(excelSheet, i, 32), GetSafeCellValue(excelSheet, i, 33), GetSafeCellValue(excelSheet, i, 34), GetSafeCellValue(excelSheet, i, 35),
                                lineColor, betaAngle, rotationValue, memberLayer)
                        ElseIf sectionType = 630.0 Then
                            Dim d1 As Double = GetSafeCellValue(excelSheet, i, 8) * 1000.0
                            Dim d2 As Double = GetSafeCellValue(excelSheet, i, 11) * 1000.0
                            Dim d3 As Double = GetSafeCellValue(excelSheet, i, 9) * 1000.0
                            Dim d4 As Double = GetSafeCellValue(excelSheet, i, 10) * 1000.0
                            Dim webThickness As Double = d1 - 2.0 * d4
                            DrawTaperedChannelSection(dxfDocument, startPoint, endPoint, webThickness, d2, d3, d4, d3, d4, webThickness, d2, d3, d4, d3, d4, lineColor, betaAngle, rotationValue, memberLayer)
                        ElseIf sectionType = 610.0 Then
                            Dim d1 As Double = GetSafeCellValue(excelSheet, i, 8) * 1000.0
                            Dim d2 As Double = GetSafeCellValue(excelSheet, i, 11) * 1000.0
                            Dim d3 As Double = GetSafeCellValue(excelSheet, i, 9) * 1000.0
                            Dim d4 As Double = GetSafeCellValue(excelSheet, i, 10) * 1000.0
                            Dim d5 As Double = GetSafeCellValue(excelSheet, i, 9) * 1000.0
                            Dim d6 As Double = GetSafeCellValue(excelSheet, i, 10) * 1000.0
                            Dim webThickness As Double = d1 - 2.0 * d4
                            DrawTaperedISection(dxfDocument, startPoint, endPoint, webThickness, d2, d3, d4, d5, d6, webThickness, d2, d3, d4, d5, d6, lineColor, betaAngle, rotationValue, memberLayer)
                        ElseIf sectionType = 613.0 Then
                            Dim d1 As Double = GetSafeCellValue(excelSheet, i, 8) * 1000.0
                            Dim d2 As Double = GetSafeCellValue(excelSheet, i, 11) * 1000.0
                            Dim d3 As Double = GetSafeCellValue(excelSheet, i, 9) * 1000.0
                            Dim d4 As Double = GetSafeCellValue(excelSheet, i, 10) * 1000.0
                            Dim d5 As Double = GetSafeCellValue(excelSheet, i, 9) * 1000.0
                            Dim d6 As Double = GetSafeCellValue(excelSheet, i, 10) * 1000.0
                            Dim extraTopWidth As Double = GetSafeCellValue(excelSheet, i, 52) * 1000.0
                            Dim extraTopThickness As Double = GetSafeCellValue(excelSheet, i, 53) * 1000.0
                            Dim webThickness As Double = d1 - 2.0 * d4
                            DrawTaperedISection(dxfDocument, startPoint, endPoint, webThickness, d2, d3, d4, d5, d6, webThickness, d2, d3, d4, d5, d6, lineColor, betaAngle, rotationValue, memberLayer)
                            If extraTopWidth > 0.0 AndAlso extraTopThickness > 0.0 Then
                                DrawExtraTopFlange(dxfDocument, startPoint, endPoint, extraTopWidth, extraTopThickness, webThickness, d4, lineColor, betaAngle, rotationValue, memberLayer)
                            End If
                        ElseIf sectionType = 614.0 Then
                            Dim d1 As Double = GetSafeCellValue(excelSheet, i, 8) * 1000.0
                            Dim d2 As Double = GetSafeCellValue(excelSheet, i, 11) * 1000.0
                            Dim d3 As Double = GetSafeCellValue(excelSheet, i, 9) * 1000.0
                            Dim d4 As Double = GetSafeCellValue(excelSheet, i, 10) * 1000.0
                            Dim d5 As Double = GetSafeCellValue(excelSheet, i, 9) * 1000.0
                            Dim d6 As Double = GetSafeCellValue(excelSheet, i, 10) * 1000.0
                            Dim extraBotWidth As Double = GetSafeCellValue(excelSheet, i, 52) * 1000.0
                            Dim extraBotThickness As Double = GetSafeCellValue(excelSheet, i, 53) * 1000.0
                            Dim webThickness As Double = d1 - 2.0 * d4
                            DrawTaperedISection(dxfDocument, startPoint, endPoint, webThickness, d2, d3, d4, d5, d6, webThickness, d2, d3, d4, d5, d6, lineColor, betaAngle, rotationValue, memberLayer)
                            If extraBotWidth > 0.0 AndAlso extraBotThickness > 0.0 Then
                                DrawExtraBottomFlange(dxfDocument, startPoint, endPoint, extraBotWidth, extraBotThickness, webThickness, d6, lineColor, betaAngle, rotationValue, memberLayer)
                            End If
                        ElseIf sectionType = 615.0 Then
                            Dim d1 As Double = GetSafeCellValue(excelSheet, i, 8) * 1000.0
                            Dim d2 As Double = GetSafeCellValue(excelSheet, i, 11) * 1000.0
                            Dim d3 As Double = GetSafeCellValue(excelSheet, i, 9) * 1000.0
                            Dim d4 As Double = GetSafeCellValue(excelSheet, i, 10) * 1000.0
                            Dim d5 As Double = GetSafeCellValue(excelSheet, i, 9) * 1000.0
                            Dim d6 As Double = GetSafeCellValue(excelSheet, i, 10) * 1000.0
                            Dim extraTopWidth As Double = GetSafeCellValue(excelSheet, i, 52) * 1000.0
                            Dim extraTopThickness As Double = GetSafeCellValue(excelSheet, i, 53) * 1000.0
                            Dim extraBotWidth As Double = GetSafeCellValue(excelSheet, i, 54) * 1000.0
                            Dim extraBotThickness As Double = GetSafeCellValue(excelSheet, i, 55) * 1000.0
                            Dim finalTopWidth, finalTopThickness, finalBotWidth, finalBotThickness As Double
                            If extraBotWidth <= 0.0 OrElse extraBotThickness <= 0.0 Then
                                finalTopWidth = extraTopWidth : finalTopThickness = extraTopThickness
                                finalBotWidth = extraTopWidth : finalBotThickness = extraTopThickness
                            ElseIf extraTopWidth <= 0.0 OrElse extraTopThickness <= 0.0 Then
                                finalTopWidth = extraBotWidth : finalTopThickness = extraBotThickness
                                finalBotWidth = extraBotWidth : finalBotThickness = extraBotThickness
                            Else
                                finalTopWidth = extraTopWidth : finalTopThickness = extraTopThickness
                                finalBotWidth = extraBotWidth : finalBotThickness = extraBotThickness
                            End If
                            Dim webThickness As Double = d1 - 2.0 * d4
                            DrawTaperedISection(dxfDocument, startPoint, endPoint, webThickness, d2, d3, d4, d5, d6, webThickness, d2, d3, d4, d5, d6, lineColor, betaAngle, rotationValue, memberLayer)
                            If finalTopWidth > 0.0 AndAlso finalTopThickness > 0.0 Then
                                DrawExtraTopFlange(dxfDocument, startPoint, endPoint, finalTopWidth, finalTopThickness, webThickness, d4, lineColor, betaAngle, rotationValue, memberLayer)
                            End If
                            If finalBotWidth > 0.0 AndAlso finalBotThickness > 0.0 Then
                                DrawExtraBottomFlange(dxfDocument, startPoint, endPoint, finalBotWidth, finalBotThickness, webThickness, d6, lineColor, betaAngle, rotationValue, memberLayer)
                            End If
                        End If
                    End If

                    If leaderDirectionAngle.HasValue Then
                        Dim angle As Double = leaderDirectionAngle.Value
                        Dim midPoint As New Vector3((startPoint.X + endPoint.X) / 2.0, (startPoint.Y + endPoint.Y) / 2.0, (startPoint.Z + endPoint.Z) / 2.0)
                        Dim viewModeInt As Integer = CInt(Math.Round(GetSafeCellValue(excelSheet, 2, 45)))
                        Dim sizeA, sizeB As Double
                        If New Integer() {3, 4, 5}.Contains(viewModeInt) Then
                            sizeA = GetSafeCellValue(excelSheet, i, 10) * 1000.0
                            sizeB = GetSafeCellValue(excelSheet, i, 12) * 1000.0
                        Else
                            sizeA = GetSafeCellValue(excelSheet, i, 7) * 1000.0
                            sizeB = GetSafeCellValue(excelSheet, i, 9) * 1000.0
                        End If
                        Dim maxSize As Double = Math.Max(sizeA, sizeB)
                        Dim baseLeaderLength As Double = If(singleLineMode, 2.0 * textHeight, 5.5 * textHeight)
                        Dim leaderLength As Double = baseLeaderLength + If(singleLineMode, 0.0, maxSize / 1.2)
                        Dim leaderBend As New Vector3(midPoint.X + leaderLength * Math.Cos(angle), midPoint.Y + leaderLength * Math.Sin(angle), midPoint.Z)
                        Dim leaderTailLength As Double = 100.0
                        Dim leaderEnd As New Vector3(leaderBend.X + leaderTailLength, leaderBend.Y, leaderBend.Z)
                        Dim leaderColor As AciColor = If(Not useColorLayers, If(viewMode = 0, lineColor, AciColor.FromCadIndex(CShort(viewMode))), lineColor)

                        Dim leaderLine1 As New Line(midPoint, leaderBend) With {.Layer = If(useColorLayers, memberLayer, leaderLinesLayer), .Color = leaderColor}
                        dxfDocument.Entities.Add(leaderLine1)
                        Dim leaderLine2 As New Line(leaderBend, leaderEnd) With {.Layer = If(useColorLayers, memberLayer, leaderLinesLayer), .Color = leaderColor}
                        dxfDocument.Entities.Add(leaderLine2)

                        Dim webText As String = GetSafeCellValueAsString(excelSheet, i, 37)
                        Dim tfText As String = GetSafeCellValueAsString(excelSheet, i, 38)
                        Dim bfText As String = GetSafeCellValueAsString(excelSheet, i, 39)
                        Dim beamLength As Double = GetSafeCellValue(excelSheet, i, 3) * 1000.0
                        Dim sectionTypeForText As Double = GetSafeCellValue(excelSheet, i, 6)
                        Dim sizeText As String
                        If useColorLayersNoLength Then
                            Dim columnEValue As String = GetSafeCellValueAsString(excelSheet, i, 5).Trim()
                            sizeText = BuildSizeTextWithoutLength(webText, tfText, bfText, sectionTypeForText, columnEValue)
                        Else
                            sizeText = BuildSizeText(webText, tfText, bfText, beamLength, sectionTypeForText)
                        End If

                        If Not String.IsNullOrWhiteSpace(sizeText) Then
                            Dim textDrop As Double = -(3.0 * textHeight + 10.0)
                            Dim textPosition As New Vector3(leaderEnd.X, leaderEnd.Y + textDrop, leaderEnd.Z)
                            Dim mText As New MText(sizeText.Replace(vbCrLf, "\P"), textPosition, textHeight, 0.0) With {
                                .Layer = If(useColorLayers, memberLayer, beamTextLayer),
                                .Color = leaderColor,
                                .AttachmentPoint = MTextAttachmentPoint.BottomLeft
                            }
                            dxfDocument.Entities.Add(mText)
                        End If
                    End If
                Catch ex As Exception
                    ' Skip this beam row and continue with the rest
                End Try
            Next

            dxfDocument.Save(dxfFilePath)
            Dim successMessage As String = String.Format("DXF file saved successfully:{0}{1}", vbCrLf, dxfFilePath)
            If gridWasDrawn Then
                successMessage = String.Format("DXF file Created successfully:{0}{1}", vbCrLf, dxfFilePath)
            End If
            If singleLineMode Then
                successMessage = String.Format("Single Line DXF file Created successfully:{0}{1}", vbCrLf, dxfFilePath)
            End If
            ShowSuccessMessageWithOpen(successMessage, dxfFilePath, drawingsDir)
        Catch ex As Exception
            MessageBox.Show($"Error exporting to DXF: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Hand)
        End Try
    End Sub

    Private Sub DrawExtraTopFlange(dxf As DxfDocument, startPoint As Vector3, endPoint As Vector3, extraFlangeWidth As Double, extraFlangeThickness As Double, webHeight As Double, baseTopFlangeThickness As Double, color As AciColor, betaAngle As Double, rotationValue As String, targetLayer As Layer)
        Dim axis As Vector3 = endPoint - startPoint
        Dim axisNorm As Vector3 = axis
        axisNorm.Normalize()
        Const zThreshold As Double = 0.99
        Dim localY, localZ As Vector3
        If Math.Abs(axisNorm.Z) > zThreshold Then
            localY = New Vector3(1.0, 0.0, 0.0)
            localZ = Vector3.CrossProduct(axisNorm, localY)
            localZ.Normalize()
            localY = Vector3.CrossProduct(localZ, axisNorm)
            localY.Normalize()
        Else
            localZ = Vector3.CrossProduct(New Vector3(0.0, 0.0, 1.0), axisNorm)
            Dim mag As Double = Math.Sqrt(localZ.X * localZ.X + localZ.Y * localZ.Y + localZ.Z * localZ.Z)
            If mag < 0.001 Then
                localZ = New Vector3(0.0, 1.0, 0.0)
            Else
                localZ.Normalize()
            End If
            localY = Vector3.CrossProduct(localZ, axisNorm)
            localY.Normalize()
            localZ = Vector3.CrossProduct(axisNorm, localY)
            localZ.Normalize()
        End If
        If rotationValue.Trim() = "5" Then
            localY = RotateVectorAroundAxis(localY, axisNorm, -90.0)
            localZ = RotateVectorAroundAxis(localZ, axisNorm, -90.0)
        ElseIf rotationValue.Trim() = "3" Then
            If Math.Abs(axisNorm.X) > 0.7 OrElse Math.Abs(axisNorm.Z) > 0.7 Then
                localY = RotateVectorAroundAxis(localY, axisNorm, -90.0)
                localZ = RotateVectorAroundAxis(localZ, axisNorm, -90.0)
            End If
            localY = RotateVectorAroundAxis(localY, axisNorm, -90.0)
            localZ = RotateVectorAroundAxis(localZ, axisNorm, -90.0)
        ElseIf rotationValue.Trim() = "4" Then
            If Math.Abs(axisNorm.X) > 0.7 OrElse Math.Abs(axisNorm.Z) > 0.7 Then
                localY = RotateVectorAroundAxis(localY, axisNorm, -90.0)
                localZ = RotateVectorAroundAxis(localZ, axisNorm, -90.0)
            End If
            localY = RotateVectorAroundAxis(localY, axisNorm, -90.0)
            localZ = RotateVectorAroundAxis(localZ, axisNorm, -90.0)
        End If
        If Math.Abs(betaAngle) > 0.001 Then
            localY = RotateVectorAroundAxis(localY, axisNorm, betaAngle)
            localZ = RotateVectorAroundAxis(localZ, axisNorm, betaAngle)
        End If
        Dim pointAt = Function(offsetY As Double, offsetZ As Double, basePoint As Vector3) basePoint + localY * offsetY + localZ * offsetZ
        Dim halfWebHeight As Double = webHeight / 2.0
        Dim flangeInnerZ As Double = halfWebHeight + baseTopFlangeThickness
        Dim flangeOuterZ As Double = flangeInnerZ + extraFlangeThickness
        Dim p1 As Vector3 = pointAt(-extraFlangeWidth / 2.0, flangeInnerZ, startPoint)
        Dim p2 As Vector3 = pointAt(extraFlangeWidth / 2.0, flangeInnerZ, startPoint)
        Dim p3 As Vector3 = pointAt(-extraFlangeWidth / 2.0, flangeOuterZ, startPoint)
        Dim p4 As Vector3 = pointAt(extraFlangeWidth / 2.0, flangeOuterZ, startPoint)
        Dim p5 As Vector3 = pointAt(-extraFlangeWidth / 2.0, flangeInnerZ, endPoint)
        Dim p6 As Vector3 = pointAt(extraFlangeWidth / 2.0, flangeInnerZ, endPoint)
        Dim p7 As Vector3 = pointAt(-extraFlangeWidth / 2.0, flangeOuterZ, endPoint)
        Dim p8 As Vector3 = pointAt(extraFlangeWidth / 2.0, flangeOuterZ, endPoint)
        dxf.Entities.Add(New Line(p1, p5) With {.Color = color, .Layer = targetLayer})
        dxf.Entities.Add(New Line(p2, p6) With {.Color = color, .Layer = targetLayer})
        dxf.Entities.Add(New Line(p3, p7) With {.Color = color, .Layer = targetLayer})
        dxf.Entities.Add(New Line(p4, p8) With {.Color = color, .Layer = targetLayer})
        dxf.Entities.Add(New Line(p1, p2) With {.Color = color, .Layer = targetLayer})
        dxf.Entities.Add(New Line(p2, p4) With {.Color = color, .Layer = targetLayer})
        dxf.Entities.Add(New Line(p4, p3) With {.Color = color, .Layer = targetLayer})
        dxf.Entities.Add(New Line(p3, p1) With {.Color = color, .Layer = targetLayer})
        dxf.Entities.Add(New Line(p5, p6) With {.Color = color, .Layer = targetLayer})
        dxf.Entities.Add(New Line(p6, p8) With {.Color = color, .Layer = targetLayer})
        dxf.Entities.Add(New Line(p8, p7) With {.Color = color, .Layer = targetLayer})
        dxf.Entities.Add(New Line(p7, p5) With {.Color = color, .Layer = targetLayer})
    End Sub

    Private Sub DrawExtraBottomFlange(dxf As DxfDocument, startPoint As Vector3, endPoint As Vector3, extraFlangeWidth As Double, extraFlangeThickness As Double, webHeight As Double, baseBottomFlangeThickness As Double, color As AciColor, betaAngle As Double, rotationValue As String, targetLayer As Layer)
        Dim axis As Vector3 = endPoint - startPoint
        Dim axisNorm As Vector3 = axis
        axisNorm.Normalize()
        Const zThreshold As Double = 0.99
        Dim localY, localZ As Vector3
        If Math.Abs(axisNorm.Z) > zThreshold Then
            localY = New Vector3(1.0, 0.0, 0.0)
            localZ = Vector3.CrossProduct(axisNorm, localY)
            localZ.Normalize()
            localY = Vector3.CrossProduct(localZ, axisNorm)
            localY.Normalize()
        Else
            localZ = Vector3.CrossProduct(New Vector3(0.0, 0.0, 1.0), axisNorm)
            Dim mag As Double = Math.Sqrt(localZ.X * localZ.X + localZ.Y * localZ.Y + localZ.Z * localZ.Z)
            If mag < 0.001 Then
                localZ = New Vector3(0.0, 1.0, 0.0)
            Else
                localZ.Normalize()
            End If
            localY = Vector3.CrossProduct(localZ, axisNorm)
            localY.Normalize()
            localZ = Vector3.CrossProduct(axisNorm, localY)
            localZ.Normalize()
        End If
        If rotationValue.Trim() = "5" Then
            localY = RotateVectorAroundAxis(localY, axisNorm, -90.0)
            localZ = RotateVectorAroundAxis(localZ, axisNorm, -90.0)
        ElseIf rotationValue.Trim() = "3" Then
            If Math.Abs(axisNorm.X) > 0.7 OrElse Math.Abs(axisNorm.Z) > 0.7 Then
                localY = RotateVectorAroundAxis(localY, axisNorm, -90.0)
                localZ = RotateVectorAroundAxis(localZ, axisNorm, -90.0)
            End If
            localY = RotateVectorAroundAxis(localY, axisNorm, -90.0)
            localZ = RotateVectorAroundAxis(localZ, axisNorm, -90.0)
        ElseIf rotationValue.Trim() = "4" Then
            If Math.Abs(axisNorm.X) > 0.7 OrElse Math.Abs(axisNorm.Z) > 0.7 Then
                localY = RotateVectorAroundAxis(localY, axisNorm, -90.0)
                localZ = RotateVectorAroundAxis(localZ, axisNorm, -90.0)
            End If
            localY = RotateVectorAroundAxis(localY, axisNorm, -90.0)
            localZ = RotateVectorAroundAxis(localZ, axisNorm, -90.0)
        End If
        If Math.Abs(betaAngle) > 0.001 Then
            localY = RotateVectorAroundAxis(localY, axisNorm, betaAngle)
            localZ = RotateVectorAroundAxis(localZ, axisNorm, betaAngle)
        End If
        Dim pointAt = Function(offsetY As Double, offsetZ As Double, basePoint As Vector3) basePoint + localY * offsetY + localZ * offsetZ
        Dim halfWebHeight As Double = -webHeight / 2.0
        Dim flangeInnerZ As Double = halfWebHeight - baseBottomFlangeThickness
        Dim flangeOuterZ As Double = flangeInnerZ - extraFlangeThickness
        Dim p1 As Vector3 = pointAt(-extraFlangeWidth / 2.0, flangeInnerZ, startPoint)
        Dim p2 As Vector3 = pointAt(extraFlangeWidth / 2.0, flangeInnerZ, startPoint)
        Dim p3 As Vector3 = pointAt(-extraFlangeWidth / 2.0, flangeOuterZ, startPoint)
        Dim p4 As Vector3 = pointAt(extraFlangeWidth / 2.0, flangeOuterZ, startPoint)
        Dim p5 As Vector3 = pointAt(-extraFlangeWidth / 2.0, flangeInnerZ, endPoint)
        Dim p6 As Vector3 = pointAt(extraFlangeWidth / 2.0, flangeInnerZ, endPoint)
        Dim p7 As Vector3 = pointAt(-extraFlangeWidth / 2.0, flangeOuterZ, endPoint)
        Dim p8 As Vector3 = pointAt(extraFlangeWidth / 2.0, flangeOuterZ, endPoint)
        dxf.Entities.Add(New Line(p1, p5) With {.Color = color, .Layer = targetLayer})
        dxf.Entities.Add(New Line(p2, p6) With {.Color = color, .Layer = targetLayer})
        dxf.Entities.Add(New Line(p3, p7) With {.Color = color, .Layer = targetLayer})
        dxf.Entities.Add(New Line(p4, p8) With {.Color = color, .Layer = targetLayer})
        dxf.Entities.Add(New Line(p1, p2) With {.Color = color, .Layer = targetLayer})
        dxf.Entities.Add(New Line(p2, p4) With {.Color = color, .Layer = targetLayer})
        dxf.Entities.Add(New Line(p4, p3) With {.Color = color, .Layer = targetLayer})
        dxf.Entities.Add(New Line(p3, p1) With {.Color = color, .Layer = targetLayer})
        dxf.Entities.Add(New Line(p5, p6) With {.Color = color, .Layer = targetLayer})
        dxf.Entities.Add(New Line(p6, p8) With {.Color = color, .Layer = targetLayer})
        dxf.Entities.Add(New Line(p8, p7) With {.Color = color, .Layer = targetLayer})
        dxf.Entities.Add(New Line(p7, p5) With {.Color = color, .Layer = targetLayer})
    End Sub

    ' Custom-styled borderless input dialog (rounded corners, draggable, hover effects on OK/Cancel).
    ' NOTE: the decompiler could not recover the body of a compiler-generated closure class backing five
    ' cosmetic handlers (Form.Load, and MouseEnter/MouseLeave on both buttons). The hover colors below are
    ' reconstructed to match the FlatAppearance.MouseOverBackColor/MouseDownBackColor values already set on
    ' each button (the standard pattern for this kind of manual hover effect); Form.Load re-focuses the input
    ' box. Everything else in this method is a faithful decompile.
    Private Function ShowProfessionalInputDialog(prompt As String, title As String, defaultValue As String) As String
        Dim form As New Form()
        Dim panel As New Panel()
        Dim label As New Label()
        Dim label2 As New Label()
        Dim textBox As New TextBox()
        Dim button As New Button()
        Dim button2 As New Button()
        Dim panel2 As New Panel()
        Dim result As String
        Try
            form.Text = String.Empty
            form.Size = New Size(480, 240)
            form.StartPosition = FormStartPosition.CenterParent
            form.FormBorderStyle = FormBorderStyle.None
            form.MaximizeBox = False
            form.MinimizeBox = False
            form.BackColor = Color.FromArgb(240, 240, 242)
            form.ShowInTaskbar = False
            form.TopMost = True

            AddHandler form.Load, Sub(a0, a1)
                                       textBox.SelectAll()
                                       textBox.Focus()
                                   End Sub
            AddHandler form.Paint, Sub(sender, e)
                                        e.Graphics.SmoothingMode = SmoothingMode.AntiAlias
                                        Dim rect As New Rectangle(2, 2, form.Width - 2, form.Height - 2)
                                        Using path As New GraphicsPath()
                                            Dim r As Integer = 12
                                            path.AddArc(rect.X, rect.Y, r, r, 180F, 90F)
                                            path.AddArc(rect.X + rect.Width - r, rect.Y, r, r, 270F, 90F)
                                            path.AddArc(rect.X + rect.Width - r, rect.Y + rect.Height - r, r, r, 0F, 90F)
                                            path.AddArc(rect.X, rect.Y + rect.Height - r, r, r, 90F, 90F)
                                            path.CloseFigure()
                                            Using brush As New SolidBrush(Color.FromArgb(30, 0, 0, 0))
                                                e.Graphics.FillPath(brush, path)
                                            End Using
                                        End Using
                                    End Sub

            panel.Location = New Point(0, 0)
            panel.Size = New Size(480, 60)
            panel.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
            panel.BackColor = Color.FromArgb(0, 120, 215)
            panel.Padding = New Padding(20, 15, 20, 5)
            AddHandler panel.Paint, Sub(sender, e)
                                         Dim rect As New Rectangle(0, 0, panel.Width, panel.Height)
                                         Using brush As New LinearGradientBrush(rect, Color.FromArgb(16, 137, 230), Color.FromArgb(0, 120, 215), LinearGradientMode.Vertical)
                                             e.Graphics.FillRectangle(brush, rect)
                                         End Using
                                     End Sub

            label.Text = title
            label.Font = New Font("Segoe UI", 12.0F, FontStyle.Bold)
            label.ForeColor = Color.White
            label.AutoSize = True
            label.Location = New Point(20, 18)
            label.BackColor = Color.Transparent

            panel2.Location = New Point(0, 60)
            panel2.Size = New Size(480, 180)
            panel2.Anchor = AnchorStyles.Top Or AnchorStyles.Bottom Or AnchorStyles.Left Or AnchorStyles.Right
            panel2.Padding = New Padding(25, 20, 25, 20)
            panel2.BackColor = Color.FromArgb(240, 240, 242)

            label2.Text = prompt
            label2.Font = New Font("Segoe UI", 10.0F, FontStyle.Regular)
            label2.ForeColor = Color.FromArgb(50, 50, 50)
            label2.AutoSize = False
            label2.Size = New Size(410, 30)
            label2.Location = New Point(25, 25)
            label2.TextAlign = ContentAlignment.MiddleLeft
            label2.BackColor = Color.Transparent

            textBox.Font = New Font("Segoe UI", 11.0F, FontStyle.Regular)
            textBox.Size = New Size(410, 32)
            textBox.Location = New Point(25, 65)
            textBox.Text = defaultValue
            textBox.BorderStyle = BorderStyle.FixedSingle
            textBox.BackColor = Color.White
            textBox.ForeColor = Color.FromArgb(40, 40, 40)
            textBox.SelectAll()
            textBox.Focus()
            AddHandler textBox.Paint, Sub(sender, e)
                                           Dim rect As New Rectangle(0, 0, textBox.Width - 1, textBox.Height - 1)
                                           Using pen As New Pen(Color.FromArgb(180, 180, 180), 1.0F)
                                               e.Graphics.SmoothingMode = SmoothingMode.AntiAlias
                                               Using path As New GraphicsPath()
                                                   Dim r As Integer = 4
                                                   path.AddArc(rect.X, rect.Y, r, r, 180F, 90F)
                                                   path.AddArc(rect.X + rect.Width - r, rect.Y, r, r, 270F, 90F)
                                                   path.AddArc(rect.X + rect.Width - r, rect.Y + rect.Height - r, r, r, 0F, 90F)
                                                   path.AddArc(rect.X, rect.Y + rect.Height - r, r, r, 90F, 90F)
                                                   path.CloseFigure()
                                                   e.Graphics.DrawPath(pen, path)
                                               End Using
                                           End Using
                                       End Sub

            button.Text = "Create DXF"
            button.Font = New Font("Segoe UI", 9.0F, FontStyle.Bold)
            button.Size = New Size(110, 38)
            button.Location = New Point(215, 115)
            button.BackColor = Color.FromArgb(0, 120, 215)
            button.ForeColor = Color.White
            button.FlatStyle = FlatStyle.Flat
            button.FlatAppearance.BorderSize = 0
            button.FlatAppearance.MouseOverBackColor = Color.FromArgb(16, 137, 230)
            button.FlatAppearance.MouseDownBackColor = Color.FromArgb(0, 103, 184)
            button.UseVisualStyleBackColor = False
            button.DialogResult = DialogResult.OK
            AddHandler button.Paint, Sub(sender, e)
                                          Using path As New GraphicsPath()
                                              Dim rect As New Rectangle(0, 0, button.Width, button.Height)
                                              Dim r As Integer = 6
                                              path.AddArc(rect.X, rect.Y, r, r, 180F, 90F)
                                              path.AddArc(rect.X + rect.Width - r, rect.Y, r, r, 270F, 90F)
                                              path.AddArc(rect.X + rect.Width - r, rect.Y + rect.Height - r, r, r, 0F, 90F)
                                              path.AddArc(rect.X, rect.Y + rect.Height - r, r, r, 90F, 90F)
                                              path.CloseFigure()
                                              button.Region = New Region(path)
                                          End Using
                                      End Sub
            AddHandler button.MouseEnter, Sub(a0, a1) button.BackColor = Color.FromArgb(16, 137, 230)
            AddHandler button.MouseLeave, Sub(a0, a1) button.BackColor = Color.FromArgb(0, 120, 215)

            button2.Text = "Cancel"
            button2.Font = New Font("Segoe UI", 9.0F, FontStyle.Regular)
            button2.Size = New Size(110, 38)
            button2.Location = New Point(335, 115)
            button2.BackColor = Color.FromArgb(235, 235, 240)
            button2.ForeColor = Color.FromArgb(60, 60, 60)
            button2.FlatStyle = FlatStyle.Flat
            button2.FlatAppearance.BorderSize = 1
            button2.FlatAppearance.BorderColor = Color.FromArgb(180, 180, 185)
            button2.FlatAppearance.MouseOverBackColor = Color.FromArgb(225, 225, 230)
            button2.FlatAppearance.MouseDownBackColor = Color.FromArgb(215, 215, 220)
            button2.UseVisualStyleBackColor = False
            button2.DialogResult = DialogResult.Cancel
            AddHandler button2.Paint, Sub(sender, e)
                                           Using path As New GraphicsPath()
                                               Dim rect As New Rectangle(0, 0, button2.Width, button2.Height)
                                               Dim r As Integer = 6
                                               path.AddArc(rect.X, rect.Y, r, r, 180F, 90F)
                                               path.AddArc(rect.X + rect.Width - r, rect.Y, r, r, 270F, 90F)
                                               path.AddArc(rect.X + rect.Width - r, rect.Y + rect.Height - r, r, r, 0F, 90F)
                                               path.AddArc(rect.X, rect.Y + rect.Height - r, r, r, 90F, 90F)
                                               path.CloseFigure()
                                               button2.Region = New Region(path)
                                           End Using
                                       End Sub
            AddHandler button2.MouseEnter, Sub(a0, a1) button2.BackColor = Color.FromArgb(225, 225, 230)
            AddHandler button2.MouseLeave, Sub(a0, a1) button2.BackColor = Color.FromArgb(235, 235, 240)

            AddHandler textBox.KeyDown, Sub(sender, e)
                                             If e.KeyCode = Keys.Return Then
                                                 button.PerformClick()
                                             ElseIf e.KeyCode = Keys.Escape Then
                                                 button2.PerformClick()
                                             End If
                                         End Sub

            Dim dragging As Boolean = False
            Dim dragStart As Point = Nothing
            AddHandler panel.MouseDown, Sub(sender, e)
                                             dragging = True
                                             dragStart = Cursor.Position
                                         End Sub
            AddHandler panel.MouseMove, Sub(sender, e)
                                             If dragging Then
                                                 Dim pos As Point = Cursor.Position
                                                 form.Location = New Point(form.Location.X + (pos.X - dragStart.X), form.Location.Y + (pos.Y - dragStart.Y))
                                                 dragStart = pos
                                             End If
                                         End Sub
            AddHandler panel.MouseUp, Sub(sender, e) dragging = False
            AddHandler label.MouseDown, Sub(sender, e)
                                             dragging = True
                                             dragStart = Cursor.Position
                                         End Sub
            AddHandler label.MouseMove, Sub(sender, e)
                                             If dragging Then
                                                 Dim pos As Point = Cursor.Position
                                                 form.Location = New Point(form.Location.X + (pos.X - dragStart.X), form.Location.Y + (pos.Y - dragStart.Y))
                                                 dragStart = pos
                                             End If
                                         End Sub
            AddHandler label.MouseUp, Sub(sender, e) dragging = False

            panel.Controls.Add(label)
            panel2.Controls.AddRange(New Control() {label2, textBox, button, button2})
            form.Controls.AddRange(New Control() {panel, panel2})
            label2.TabIndex = 0
            textBox.TabIndex = 1
            button.TabIndex = 2
            button2.TabIndex = 3
            form.AcceptButton = button
            form.CancelButton = button2

            Dim dialogResult As DialogResult = form.ShowDialog()
            result = If(dialogResult <> DialogResult.OK OrElse String.IsNullOrWhiteSpace(textBox.Text.Trim()), String.Empty, textBox.Text.Trim())
        Catch ex As Exception
            MessageBox.Show($"Error creating custom dialog: {ex.Message}", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            result = Interaction.InputBox(prompt, title, defaultValue)
        Finally
            If form IsNot Nothing Then form.Dispose()
        End Try
        Return result
    End Function

    Private Sub ShowSuccessMessageWithOpen(successMessage As String, filePath As String, folderPath As String)
        Dim form As New Form()
        form.Text = ChrW(&HD83D) & ChrW(&HDD8C) & ChrW(&HFE0F) & " Designed with Passion by Peeyush Raghav"
        form.Size = New Size(650, 320)
        form.StartPosition = FormStartPosition.CenterScreen
        form.FormBorderStyle = FormBorderStyle.FixedDialog
        form.MaximizeBox = False
        form.MinimizeBox = False
        form.BackColor = Color.White
        form.Font = New Font("Segoe UI", 9.0F, FontStyle.Regular)

        Dim panel As New Panel() With {.Size = New Size(650, 60), .Location = New Point(0, 0), .BackColor = Color.FromArgb(45, 52, 67)}
        Dim label As New Label() With {
            .Text = "" & ChrW(&H2713), .Font = New Font("Segoe UI", 24.0F, FontStyle.Bold), .ForeColor = Color.White,
            .Size = New Size(40, 40), .Location = New Point(20, 10), .TextAlign = ContentAlignment.MiddleCenter
        }
        Dim label2 As New Label() With {
            .Text = "Export Completed Successfully", .Font = New Font("Segoe UI", 14.0F, FontStyle.Bold), .ForeColor = Color.White,
            .Size = New Size(500, 30), .Location = New Point(70, 15), .TextAlign = ContentAlignment.MiddleLeft
        }
        Dim label3 As New Label() With {
            .Text = successMessage, .Font = New Font("Segoe UI", 10.0F, FontStyle.Regular), .ForeColor = Color.FromArgb(64, 64, 64),
            .Size = New Size(610, 80), .Location = New Point(20, 80), .AutoSize = False, .TextAlign = ContentAlignment.TopLeft
        }
        Dim panel2 As New Panel() With {.Size = New Size(610, 1), .Location = New Point(20, 190), .BackColor = Color.FromArgb(220, 220, 220)}
        Dim button As New Button() With {
            .Text = "OK", .Size = New Size(120, 40), .Location = New Point(120, 220), .DialogResult = DialogResult.OK,
            .BackColor = Color.FromArgb(240, 240, 240), .ForeColor = Color.FromArgb(64, 64, 64), .Font = New Font("Segoe UI", 9.0F, FontStyle.Regular),
            .FlatStyle = FlatStyle.Flat
        }
        button.FlatAppearance.BorderSize = 1
        button.FlatAppearance.BorderColor = Color.FromArgb(180, 180, 180)
        button.FlatAppearance.MouseOverBackColor = Color.FromArgb(230, 230, 230)

        Dim button2 As New Button() With {
            .Text = "SHOW IN FOLDER", .Size = New Size(150, 40), .Location = New Point(260, 220),
            .BackColor = Color.FromArgb(0, 122, 204), .ForeColor = Color.White, .Font = New Font("Segoe UI", 9.0F, FontStyle.Bold),
            .FlatStyle = FlatStyle.Flat
        }
        button2.FlatAppearance.BorderSize = 0
        button2.FlatAppearance.MouseOverBackColor = Color.FromArgb(0, 100, 180)

        Dim button3 As New Button() With {
            .Text = "OPEN IN CAD", .Size = New Size(130, 40), .Location = New Point(430, 220),
            .BackColor = Color.FromArgb(40, 167, 69), .ForeColor = Color.White, .Font = New Font("Segoe UI", 9.0F, FontStyle.Bold),
            .FlatStyle = FlatStyle.Flat
        }
        button3.FlatAppearance.BorderSize = 0
        button3.FlatAppearance.MouseOverBackColor = Color.FromArgb(33, 136, 56)

        AddHandler button2.Click, Sub(sender, e)
                                       Try
                                           Process.Start("explorer.exe", $"/select,""{filePath}""")
                                       Catch ex As Exception
                                           Try
                                               Process.Start("explorer.exe", folderPath)
                                           Catch
                                               MessageBox.Show("Could not open file location.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                                           End Try
                                       End Try
                                       form.Close()
                                   End Sub
        AddHandler button3.Click, Sub(sender, e)
                                       Try
                                           Process.Start(filePath)
                                       Catch ex As Exception
                                           Try
                                               Process.Start(New ProcessStartInfo() With {.FileName = filePath, .UseShellExecute = True, .Verb = "open"})
                                           Catch
                                               MessageBox.Show("Could not open DXF file. Please ensure you have a CAD application installed.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                                           End Try
                                       End Try
                                       form.Close()
                                   End Sub

        panel.Controls.Add(label)
        panel.Controls.Add(label2)
        form.Controls.Add(panel)
        form.Controls.Add(label3)
        form.Controls.Add(panel2)
        form.Controls.Add(button)
        form.Controls.Add(button2)
        form.Controls.Add(button3)
        form.ShowDialog()
        form.Dispose()
    End Sub

    Private Function BuildSizeTextWithoutLength(web As String, tf As String, bf As String, columnFValue As Double, columnEValue As String) As String
        web = If(web?.Replace("PL", "").Trim(), "")
        tf = If(tf?.Replace("PL", "").Trim(), "")
        bf = If(bf?.Replace("PL", "").Trim(), "")
        Dim list As New List(Of String)()
        If columnFValue = 680.0 Then
            If Not String.IsNullOrEmpty(web) Then list.Add($"W-{web}")
            If Not String.IsNullOrEmpty(tf) AndAlso Not String.IsNullOrEmpty(bf) Then
                list.Add(If(tf = bf, $"F-{tf}", $"TF-{tf}\PBF-{bf}"))
            ElseIf Not String.IsNullOrEmpty(tf) Then
                list.Add($"F-{tf}")
            ElseIf Not String.IsNullOrEmpty(bf) Then
                list.Add($"F-{bf}")
            End If
        Else
            If Not String.IsNullOrEmpty(web) Then list.Add(web)
            If Not String.IsNullOrEmpty(tf) AndAlso Not String.IsNullOrEmpty(bf) Then
                list.Add(If(tf = bf, tf, $"{tf}\P{bf}"))
            ElseIf Not String.IsNullOrEmpty(tf) Then
                list.Add(tf)
            ElseIf Not String.IsNullOrEmpty(bf) Then
                list.Add(bf)
            End If
        End If
        If Not String.IsNullOrEmpty(columnEValue) Then list.Add($"P-{columnEValue}")
        Return String.Join("\P", list)
    End Function

    Private Function HasValidBeamData(excelSheet As Object) As Boolean
        Try
            Dim ws As Worksheet = CType(excelSheet, Worksheet)
            Dim lastRow As Integer = CInt(ws.Cells(ws.Rows.Count, 1).End(XlDirection.xlUp).Row)
            For i As Integer = 9 To lastRow
                Dim v1 As Double = GetSafeCellValue(excelSheet, i, 16)
                Dim v2 As Double = GetSafeCellValue(excelSheet, i, 17)
                Dim v3 As Double = GetSafeCellValue(excelSheet, i, 18)
                Dim v4 As Double = GetSafeCellValue(excelSheet, i, 19)
                Dim v5 As Double = GetSafeCellValue(excelSheet, i, 20)
                Dim v6 As Double = GetSafeCellValue(excelSheet, i, 21)
                If v1 <> 0.0 OrElse v2 <> 0.0 OrElse v3 <> 0.0 OrElse v4 <> 0.0 OrElse v5 <> 0.0 OrElse v6 <> 0.0 Then
                    Return True
                End If
            Next
            Return False
        Catch ex As Exception
            Return False
        End Try
    End Function

    Private Function AddGridToDxf(dxfDoc As DxfDocument, excelSheet As Object, rotationValue As String) As Boolean
        Try
            Dim worksheet As Worksheet = CType(excelSheet, Worksheet)
            Dim workbook As Workbook = CType(worksheet.Parent, Workbook)
            Dim gridSheet As Worksheet
            Try
                gridSheet = CType(workbook.Worksheets("Sheet4"), Worksheet)
            Catch ex As Exception
                Return False
            End Try

            Dim xText As String = GetSafeCellValueAsString(gridSheet, 2, 3).Trim()
            Dim yText As String = GetSafeCellValueAsString(gridSheet, 3, 3).Trim()
            Dim zText As String = GetSafeCellValueAsString(gridSheet, 4, 3).Trim()
            Dim includesZeroX As Boolean = False
            Dim includesZeroY As Boolean = False
            Dim includesZeroZ As Boolean = False
            Dim xOffsets As Double() = If(String.IsNullOrWhiteSpace(xText), New Double() {}, ParseOffsets(xText, includesZeroX))
            Dim yOffsets As Double() = If(String.IsNullOrWhiteSpace(yText), New Double() {}, ParseOffsets(yText, includesZeroY))
            Dim zOffsets As Double() = If(String.IsNullOrWhiteSpace(zText), New Double() {}, ParseOffsets(zText, includesZeroZ))

            If xOffsets.Length = 0 AndAlso yOffsets.Length = 0 AndAlso zOffsets.Length = 0 Then
                Return False
            End If

            Dim xPositions As Double() = If(xOffsets.Length > 0, CalculatePositions(xOffsets), New Double() {})
            Dim yPositions As Double() = If(yOffsets.Length > 0, CalculatePositions(yOffsets), New Double() {})
            Dim zPositions As Double() = If(zOffsets.Length > 0, CalculatePositions(zOffsets), New Double() {})

            Dim textHeight As Double = 300.0
            Dim textHeightText As String = GetSafeCellValueAsString(gridSheet, 2, 5).Trim()
            If Not String.IsNullOrWhiteSpace(textHeightText) Then Double.TryParse(textHeightText, textHeight)

            Dim xMargin As Double = 1000.0
            Dim yMargin As Double = 1000.0
            Dim zMargin As Double = 1000.0
            Dim xMarginText As String = GetSafeCellValueAsString(gridSheet, 2, 4).Trim()
            If Not String.IsNullOrWhiteSpace(xMarginText) Then Double.TryParse(xMarginText, xMargin)
            Dim yMarginText As String = GetSafeCellValueAsString(gridSheet, 3, 4).Trim()
            If Not String.IsNullOrWhiteSpace(yMarginText) Then Double.TryParse(yMarginText, yMargin)
            Dim zMarginText As String = GetSafeCellValueAsString(gridSheet, 4, 4).Trim()
            If Not String.IsNullOrWhiteSpace(zMarginText) Then Double.TryParse(zMarginText, zMargin)

            Dim xNames As String() = If(String.IsNullOrWhiteSpace(GetSafeCellValueAsString(gridSheet, 2, 2)), New String() {}, ParseNames(GetSafeCellValueAsString(gridSheet, 2, 2)))
            Dim yNames As String() = If(String.IsNullOrWhiteSpace(GetSafeCellValueAsString(gridSheet, 3, 2)), New String() {}, ParseNames(GetSafeCellValueAsString(gridSheet, 3, 2)))
            Dim zNames As String() = If(String.IsNullOrWhiteSpace(GetSafeCellValueAsString(gridSheet, 4, 2)), New String() {}, ParseNames(GetSafeCellValueAsString(gridSheet, 4, 2)))

            Dim xSource As Double() = If(xPositions.Length > 0, xPositions, New Double(0) {})
            Dim ySource As Double() = If(yPositions.Length > 0, yPositions, New Double(0) {})
            Dim zSource As Double() = If(zPositions.Length > 0, zPositions, New Double(0) {})
            Dim xMin As Double = If(xPositions.Length > 0, xSource.Min(), 0.0)
            Dim xMax As Double = If(xPositions.Length > 0, xSource.Max(), 0.0)
            Dim yMin As Double = If(yPositions.Length > 0, ySource.Min(), 0.0)
            Dim yMax As Double = If(yPositions.Length > 0, ySource.Max(), 0.0)
            Dim zMin As Double = If(zPositions.Length > 0, zSource.Min(), 0.0)
            Dim zMax As Double = If(zPositions.Length > 0, zSource.Max(), 0.0)
            Dim xFixed As Double = If(includesZeroX OrElse xPositions.Length = 0, 0.0, xMin)
            Dim yFixed As Double = If(includesZeroY OrElse yPositions.Length = 0, 0.0, yMin)
            Dim zFixed As Double = If(includesZeroZ OrElse zPositions.Length = 0, 0.0, zMin)

            Dim gridColor As AciColor = AciColor.FromCadIndex(6)
            Dim circleRadius As Double = textHeight * 0.8

            If xPositions.Length > 0 Then
                Dim layerX As New Layer("GRID_X") With {.Color = gridColor}
                dxfDoc.Layers.Add(layerX)
                For i As Integer = 0 To xPositions.Length - 1
                    Dim xVal As Double = xPositions(i)
                    Dim label As String = If(i < xNames.Length, xNames(i), $"X{i + 1}")
                    If yPositions.Length > 0 Then
                        Dim sp As Vector3 = ApplyRotation(New Vector3(xVal, yMin - xMargin + circleRadius, zFixed), rotationValue)
                        Dim ep As Vector3 = ApplyRotation(New Vector3(xVal, yMax + xMargin - circleRadius, zFixed), rotationValue)
                        dxfDoc.Entities.Add(New Line(sp, ep) With {.Layer = layerX})
                        Dim c1 As Vector3 = ApplyRotation(New Vector3(xVal, yMin - xMargin, zFixed), rotationValue)
                        Dim c2 As Vector3 = ApplyRotation(New Vector3(xVal, yMax + xMargin, zFixed), rotationValue)
                        dxfDoc.Entities.Add(New Text(label, c1, textHeight) With {.Layer = layerX, .Alignment = TextAlignment.MiddleCenter})
                        dxfDoc.Entities.Add(New Circle(c1, circleRadius) With {.Layer = layerX})
                        dxfDoc.Entities.Add(New Text(label, c2, textHeight) With {.Layer = layerX, .Alignment = TextAlignment.MiddleCenter})
                        dxfDoc.Entities.Add(New Circle(c2, circleRadius) With {.Layer = layerX})
                    End If
                    If zPositions.Length > 0 Then
                        Dim sp As Vector3 = ApplyRotation(New Vector3(xVal, yFixed, zMin - xMargin + circleRadius), rotationValue)
                        Dim ep As Vector3 = ApplyRotation(New Vector3(xVal, yFixed, zMax + xMargin - circleRadius), rotationValue)
                        dxfDoc.Entities.Add(New Line(sp, ep) With {.Layer = layerX})
                        Dim c1 As Vector3 = ApplyRotation(New Vector3(xVal, yFixed, zMin - xMargin), rotationValue)
                        Dim c2 As Vector3 = ApplyRotation(New Vector3(xVal, yFixed, zMax + xMargin), rotationValue)
                        dxfDoc.Entities.Add(New Text(label, c1, textHeight) With {.Layer = layerX, .Alignment = TextAlignment.MiddleCenter})
                        dxfDoc.Entities.Add(New Circle(c1, circleRadius) With {.Layer = layerX})
                        dxfDoc.Entities.Add(New Text(label, c2, textHeight) With {.Layer = layerX, .Alignment = TextAlignment.MiddleCenter})
                        dxfDoc.Entities.Add(New Circle(c2, circleRadius) With {.Layer = layerX})
                    End If
                Next
                If yPositions.Length > 0 AndAlso zPositions.Length > 0 Then
                    For j As Integer = 0 To zPositions.Length - 1
                        Dim zVal As Double = zPositions(j)
                        Dim label As String = If(j < zNames.Length, zNames(j), $"Z{j + 1}")
                        Dim sp As Vector3 = ApplyRotation(New Vector3(xMin, yMin - xMargin + circleRadius, zVal), rotationValue)
                        Dim ep As Vector3 = ApplyRotation(New Vector3(xMin, yMax + xMargin - circleRadius, zVal), rotationValue)
                        dxfDoc.Entities.Add(New Line(sp, ep) With {.Layer = layerX})
                        Dim c1 As Vector3 = ApplyRotation(New Vector3(xMin, yMin - xMargin, zVal), rotationValue)
                        Dim c2 As Vector3 = ApplyRotation(New Vector3(xMin, yMax + xMargin, zVal), rotationValue)
                        dxfDoc.Entities.Add(New Text(label, c1, textHeight) With {.Layer = layerX, .Alignment = TextAlignment.MiddleCenter})
                        dxfDoc.Entities.Add(New Circle(c1, circleRadius) With {.Layer = layerX})
                        dxfDoc.Entities.Add(New Text(label, c2, textHeight) With {.Layer = layerX, .Alignment = TextAlignment.MiddleCenter})
                        dxfDoc.Entities.Add(New Circle(c2, circleRadius) With {.Layer = layerX})
                    Next
                    For k As Integer = 0 To yPositions.Length - 1
                        Dim yVal As Double = yPositions(k)
                        Dim label As String = If(k < yNames.Length, yNames(k), $"Y{k + 1}")
                        Dim sp As Vector3 = ApplyRotation(New Vector3(xMin, yVal, zMin - xMargin + circleRadius), rotationValue)
                        Dim ep As Vector3 = ApplyRotation(New Vector3(xMin, yVal, zMax + xMargin - circleRadius), rotationValue)
                        dxfDoc.Entities.Add(New Line(sp, ep) With {.Layer = layerX})
                        Dim c1 As Vector3 = ApplyRotation(New Vector3(xMin, yVal, zMin - xMargin), rotationValue)
                        Dim c2 As Vector3 = ApplyRotation(New Vector3(xMin, yVal, zMax + xMargin), rotationValue)
                        dxfDoc.Entities.Add(New Text(label, c1, textHeight) With {.Layer = layerX, .Alignment = TextAlignment.MiddleCenter})
                        dxfDoc.Entities.Add(New Circle(c1, circleRadius) With {.Layer = layerX})
                        dxfDoc.Entities.Add(New Text(label, c2, textHeight) With {.Layer = layerX, .Alignment = TextAlignment.MiddleCenter})
                        dxfDoc.Entities.Add(New Circle(c2, circleRadius) With {.Layer = layerX})
                    Next
                End If
                If yPositions.Length > 0 AndAlso zPositions.Length > 0 Then
                    For j As Integer = 0 To zPositions.Length - 1
                        Dim zVal As Double = zPositions(j)
                        Dim label As String = If(j < zNames.Length, zNames(j), $"Z{j + 1}")
                        Dim sp As Vector3 = ApplyRotation(New Vector3(xMax, yMin - xMargin + circleRadius, zVal), rotationValue)
                        Dim ep As Vector3 = ApplyRotation(New Vector3(xMax, yMax + xMargin - circleRadius, zVal), rotationValue)
                        dxfDoc.Entities.Add(New Line(sp, ep) With {.Layer = layerX})
                        Dim c1 As Vector3 = ApplyRotation(New Vector3(xMax, yMin - xMargin, zVal), rotationValue)
                        Dim c2 As Vector3 = ApplyRotation(New Vector3(xMax, yMax + xMargin, zVal), rotationValue)
                        dxfDoc.Entities.Add(New Text(label, c1, textHeight) With {.Layer = layerX, .Alignment = TextAlignment.MiddleCenter})
                        dxfDoc.Entities.Add(New Circle(c1, circleRadius) With {.Layer = layerX})
                        dxfDoc.Entities.Add(New Text(label, c2, textHeight) With {.Layer = layerX, .Alignment = TextAlignment.MiddleCenter})
                        dxfDoc.Entities.Add(New Circle(c2, circleRadius) With {.Layer = layerX})
                    Next
                    For k As Integer = 0 To yPositions.Length - 1
                        Dim yVal As Double = yPositions(k)
                        Dim label As String = If(k < yNames.Length, yNames(k), $"Y{k + 1}")
                        Dim sp As Vector3 = ApplyRotation(New Vector3(xMax, yVal, zMin - xMargin + circleRadius), rotationValue)
                        Dim ep As Vector3 = ApplyRotation(New Vector3(xMax, yVal, zMax + xMargin - circleRadius), rotationValue)
                        dxfDoc.Entities.Add(New Line(sp, ep) With {.Layer = layerX})
                        Dim c1 As Vector3 = ApplyRotation(New Vector3(xMax, yVal, zMin - xMargin), rotationValue)
                        Dim c2 As Vector3 = ApplyRotation(New Vector3(xMax, yVal, zMax + xMargin), rotationValue)
                        dxfDoc.Entities.Add(New Text(label, c1, textHeight) With {.Layer = layerX, .Alignment = TextAlignment.MiddleCenter})
                        dxfDoc.Entities.Add(New Circle(c1, circleRadius) With {.Layer = layerX})
                        dxfDoc.Entities.Add(New Text(label, c2, textHeight) With {.Layer = layerX, .Alignment = TextAlignment.MiddleCenter})
                        dxfDoc.Entities.Add(New Circle(c2, circleRadius) With {.Layer = layerX})
                    Next
                End If
            End If

            If yPositions.Length > 0 Then
                Dim layerY As New Layer("GRID_Y") With {.Color = gridColor}
                dxfDoc.Layers.Add(layerY)
                For n As Integer = 0 To yPositions.Length - 1
                    Dim yVal As Double = yPositions(n)
                    Dim label As String = If(n < yNames.Length, yNames(n), $"Y{n + 1}")
                    Dim r As Double = circleRadius
                    If xPositions.Length > 0 Then
                        Dim sp As Vector3 = ApplyRotation(New Vector3(xMin - yMargin + r, yVal, zFixed), rotationValue)
                        Dim ep As Vector3 = ApplyRotation(New Vector3(xMax + yMargin - r, yVal, zFixed), rotationValue)
                        dxfDoc.Entities.Add(New Line(sp, ep) With {.Layer = layerY})
                        Dim c1 As Vector3 = ApplyRotation(New Vector3(xMin - yMargin, yVal, zFixed), rotationValue)
                        Dim c2 As Vector3 = ApplyRotation(New Vector3(xMax + yMargin, yVal, zFixed), rotationValue)
                        dxfDoc.Entities.Add(New Text(label, c1, textHeight) With {.Layer = layerY, .Alignment = TextAlignment.MiddleCenter})
                        dxfDoc.Entities.Add(New Circle(c1, circleRadius) With {.Layer = layerY})
                        dxfDoc.Entities.Add(New Text(label, c2, textHeight) With {.Layer = layerY, .Alignment = TextAlignment.MiddleCenter})
                        dxfDoc.Entities.Add(New Circle(c2, circleRadius) With {.Layer = layerY})
                    End If
                    If zPositions.Length > 0 Then
                        Dim sp As Vector3 = ApplyRotation(New Vector3(xFixed, yVal, zMin - yMargin + r), rotationValue)
                        Dim ep As Vector3 = ApplyRotation(New Vector3(xFixed, yVal, zMax + yMargin - r), rotationValue)
                        dxfDoc.Entities.Add(New Line(sp, ep) With {.Layer = layerY})
                        Dim c1 As Vector3 = ApplyRotation(New Vector3(xFixed, yVal, zMin - yMargin), rotationValue)
                        Dim c2 As Vector3 = ApplyRotation(New Vector3(xFixed, yVal, zMax + yMargin), rotationValue)
                        dxfDoc.Entities.Add(New Text(label, c1, textHeight) With {.Layer = layerY, .Alignment = TextAlignment.MiddleCenter})
                        dxfDoc.Entities.Add(New Circle(c1, circleRadius) With {.Layer = layerY})
                        dxfDoc.Entities.Add(New Text(label, c2, textHeight) With {.Layer = layerY, .Alignment = TextAlignment.MiddleCenter})
                        dxfDoc.Entities.Add(New Circle(c2, circleRadius) With {.Layer = layerY})
                    End If
                Next
                If xPositions.Length > 0 AndAlso zPositions.Length > 0 Then
                    For idx As Integer = 0 To zPositions.Length - 1
                        Dim zVal As Double = zPositions(idx)
                        Dim label As String = If(idx < zNames.Length, zNames(idx), $"Z{idx + 1}")
                        Dim sp As Vector3 = ApplyRotation(New Vector3(xMin - yMargin + circleRadius, yMin, zVal), rotationValue)
                        Dim ep As Vector3 = ApplyRotation(New Vector3(xMax + yMargin - circleRadius, yMin, zVal), rotationValue)
                        dxfDoc.Entities.Add(New Line(sp, ep) With {.Layer = layerY})
                        Dim c1 As Vector3 = ApplyRotation(New Vector3(xMin - yMargin, yMin, zVal), rotationValue)
                        Dim c2 As Vector3 = ApplyRotation(New Vector3(xMax + yMargin, yMin, zVal), rotationValue)
                        dxfDoc.Entities.Add(New Text(label, c1, textHeight) With {.Layer = layerY, .Alignment = TextAlignment.MiddleCenter})
                        dxfDoc.Entities.Add(New Circle(c1, circleRadius) With {.Layer = layerY})
                        dxfDoc.Entities.Add(New Text(label, c2, textHeight) With {.Layer = layerY, .Alignment = TextAlignment.MiddleCenter})
                        dxfDoc.Entities.Add(New Circle(c2, circleRadius) With {.Layer = layerY})
                    Next
                    For idx As Integer = 0 To xPositions.Length - 1
                        Dim xVal As Double = xPositions(idx)
                        Dim label As String = If(idx < xNames.Length, xNames(idx), $"X{idx + 1}")
                        Dim sp As Vector3 = ApplyRotation(New Vector3(xVal, yMin, zMin - yMargin + circleRadius), rotationValue)
                        Dim ep As Vector3 = ApplyRotation(New Vector3(xVal, yMin, zMax + yMargin - circleRadius), rotationValue)
                        dxfDoc.Entities.Add(New Line(sp, ep) With {.Layer = layerY})
                        Dim c1 As Vector3 = ApplyRotation(New Vector3(xVal, yMin, zMin - yMargin), rotationValue)
                        Dim c2 As Vector3 = ApplyRotation(New Vector3(xVal, yMin, zMax + yMargin), rotationValue)
                        dxfDoc.Entities.Add(New Text(label, c1, textHeight) With {.Layer = layerY, .Alignment = TextAlignment.MiddleCenter})
                        dxfDoc.Entities.Add(New Circle(c1, circleRadius) With {.Layer = layerY})
                        dxfDoc.Entities.Add(New Text(label, c2, textHeight) With {.Layer = layerY, .Alignment = TextAlignment.MiddleCenter})
                        dxfDoc.Entities.Add(New Circle(c2, circleRadius) With {.Layer = layerY})
                    Next
                End If
                If xPositions.Length > 0 AndAlso zPositions.Length > 0 Then
                    For idx As Integer = 0 To zPositions.Length - 1
                        Dim zVal As Double = zPositions(idx)
                        Dim label As String = If(idx < zNames.Length, zNames(idx), $"Z{idx + 1}")
                        Dim sp As Vector3 = ApplyRotation(New Vector3(xMin - yMargin + circleRadius, yMax, zVal), rotationValue)
                        Dim ep As Vector3 = ApplyRotation(New Vector3(xMax + yMargin - circleRadius, yMax, zVal), rotationValue)
                        dxfDoc.Entities.Add(New Line(sp, ep) With {.Layer = layerY})
                        Dim c1 As Vector3 = ApplyRotation(New Vector3(xMin - yMargin, yMax, zVal), rotationValue)
                        Dim c2 As Vector3 = ApplyRotation(New Vector3(xMax + yMargin, yMax, zVal), rotationValue)
                        dxfDoc.Entities.Add(New Text(label, c1, textHeight) With {.Layer = layerY, .Alignment = TextAlignment.MiddleCenter})
                        dxfDoc.Entities.Add(New Circle(c1, circleRadius) With {.Layer = layerY})
                        dxfDoc.Entities.Add(New Text(label, c2, textHeight) With {.Layer = layerY, .Alignment = TextAlignment.MiddleCenter})
                        dxfDoc.Entities.Add(New Circle(c2, circleRadius) With {.Layer = layerY})
                    Next
                    For idx As Integer = 0 To xPositions.Length - 1
                        Dim xVal As Double = xPositions(idx)
                        Dim label As String = If(idx < xNames.Length, xNames(idx), $"X{idx + 1}")
                        Dim sp As Vector3 = ApplyRotation(New Vector3(xVal, yMax, zMin - yMargin + circleRadius), rotationValue)
                        Dim ep As Vector3 = ApplyRotation(New Vector3(xVal, yMax, zMax + yMargin - circleRadius), rotationValue)
                        dxfDoc.Entities.Add(New Line(sp, ep) With {.Layer = layerY})
                        Dim c1 As Vector3 = ApplyRotation(New Vector3(xVal, yMax, zMin - yMargin), rotationValue)
                        Dim c2 As Vector3 = ApplyRotation(New Vector3(xVal, yMax, zMax + yMargin), rotationValue)
                        dxfDoc.Entities.Add(New Text(label, c1, textHeight) With {.Layer = layerY, .Alignment = TextAlignment.MiddleCenter})
                        dxfDoc.Entities.Add(New Circle(c1, circleRadius) With {.Layer = layerY})
                        dxfDoc.Entities.Add(New Text(label, c2, textHeight) With {.Layer = layerY, .Alignment = TextAlignment.MiddleCenter})
                        dxfDoc.Entities.Add(New Circle(c2, circleRadius) With {.Layer = layerY})
                    Next
                End If
            End If

            If zPositions.Length > 0 Then
                Dim layerZ As New Layer("GRID_Z") With {.Color = gridColor}
                dxfDoc.Layers.Add(layerZ)
                For idx As Integer = 0 To zPositions.Length - 1
                    Dim zVal As Double = zPositions(idx)
                    Dim label As String = If(idx < zNames.Length, zNames(idx), $"Z{idx + 1}")
                    Dim r As Double = circleRadius
                    If xPositions.Length > 0 Then
                        Dim sp As Vector3 = ApplyRotation(New Vector3(xMin - zMargin + r, yFixed, zVal), rotationValue)
                        Dim ep As Vector3 = ApplyRotation(New Vector3(xMax + zMargin - r, yFixed, zVal), rotationValue)
                        dxfDoc.Entities.Add(New Line(sp, ep) With {.Layer = layerZ})
                        Dim c1 As Vector3 = ApplyRotation(New Vector3(xMin - zMargin, yFixed, zVal), rotationValue)
                        Dim c2 As Vector3 = ApplyRotation(New Vector3(xMax + zMargin, yFixed, zVal), rotationValue)
                        dxfDoc.Entities.Add(New Text(label, c1, textHeight) With {.Layer = layerZ, .Alignment = TextAlignment.MiddleCenter})
                        dxfDoc.Entities.Add(New Circle(c1, circleRadius) With {.Layer = layerZ})
                        dxfDoc.Entities.Add(New Text(label, c2, textHeight) With {.Layer = layerZ, .Alignment = TextAlignment.MiddleCenter})
                        dxfDoc.Entities.Add(New Circle(c2, circleRadius) With {.Layer = layerZ})
                    End If
                    If yPositions.Length > 0 Then
                        Dim sp As Vector3 = ApplyRotation(New Vector3(xFixed, yMin - zMargin + r, zVal), rotationValue)
                        Dim ep As Vector3 = ApplyRotation(New Vector3(xFixed, yMax + zMargin - r, zVal), rotationValue)
                        dxfDoc.Entities.Add(New Line(sp, ep) With {.Layer = layerZ})
                        Dim c1 As Vector3 = ApplyRotation(New Vector3(xFixed, yMin - zMargin, zVal), rotationValue)
                        Dim c2 As Vector3 = ApplyRotation(New Vector3(xFixed, yMax + zMargin, zVal), rotationValue)
                        dxfDoc.Entities.Add(New Text(label, c1, textHeight) With {.Layer = layerZ, .Alignment = TextAlignment.MiddleCenter})
                        dxfDoc.Entities.Add(New Circle(c1, circleRadius) With {.Layer = layerZ})
                        dxfDoc.Entities.Add(New Text(label, c2, textHeight) With {.Layer = layerZ, .Alignment = TextAlignment.MiddleCenter})
                        dxfDoc.Entities.Add(New Circle(c2, circleRadius) With {.Layer = layerZ})
                    End If
                Next
                If xPositions.Length > 0 AndAlso yPositions.Length > 0 Then
                    For idx As Integer = 0 To yPositions.Length - 1
                        Dim yVal As Double = yPositions(idx)
                        Dim label As String = If(idx < yNames.Length, yNames(idx), $"Y{idx + 1}")
                        Dim sp As Vector3 = ApplyRotation(New Vector3(xMin - zMargin + circleRadius, yVal, zMin), rotationValue)
                        Dim ep As Vector3 = ApplyRotation(New Vector3(xMax + zMargin - circleRadius, yVal, zMin), rotationValue)
                        dxfDoc.Entities.Add(New Line(sp, ep) With {.Layer = layerZ})
                        Dim c1 As Vector3 = ApplyRotation(New Vector3(xMin - zMargin, yVal, zMin), rotationValue)
                        Dim c2 As Vector3 = ApplyRotation(New Vector3(xMax + zMargin, yVal, zMin), rotationValue)
                        dxfDoc.Entities.Add(New Text(label, c1, textHeight) With {.Layer = layerZ, .Alignment = TextAlignment.MiddleCenter})
                        dxfDoc.Entities.Add(New Circle(c1, circleRadius) With {.Layer = layerZ})
                        dxfDoc.Entities.Add(New Text(label, c2, textHeight) With {.Layer = layerZ, .Alignment = TextAlignment.MiddleCenter})
                        dxfDoc.Entities.Add(New Circle(c2, circleRadius) With {.Layer = layerZ})
                    Next
                    For idx As Integer = 0 To xPositions.Length - 1
                        Dim xVal As Double = xPositions(idx)
                        Dim label As String = If(idx < xNames.Length, xNames(idx), $"X{idx + 1}")
                        Dim sp As Vector3 = ApplyRotation(New Vector3(xVal, yMin - zMargin + circleRadius, zMin), rotationValue)
                        Dim ep As Vector3 = ApplyRotation(New Vector3(xVal, yMax + zMargin - circleRadius, zMin), rotationValue)
                        dxfDoc.Entities.Add(New Line(sp, ep) With {.Layer = layerZ})
                        Dim c1 As Vector3 = ApplyRotation(New Vector3(xVal, yMin - zMargin, zMin), rotationValue)
                        Dim c2 As Vector3 = ApplyRotation(New Vector3(xVal, yMax + zMargin, zMin), rotationValue)
                        dxfDoc.Entities.Add(New Text(label, c1, textHeight) With {.Layer = layerZ, .Alignment = TextAlignment.MiddleCenter})
                        dxfDoc.Entities.Add(New Circle(c1, circleRadius) With {.Layer = layerZ})
                        dxfDoc.Entities.Add(New Text(label, c2, textHeight) With {.Layer = layerZ, .Alignment = TextAlignment.MiddleCenter})
                        dxfDoc.Entities.Add(New Circle(c2, circleRadius) With {.Layer = layerZ})
                    Next
                End If
                If xPositions.Length > 0 AndAlso yPositions.Length > 0 Then
                    For idx As Integer = 0 To yPositions.Length - 1
                        Dim yVal As Double = yPositions(idx)
                        Dim label As String = If(idx < yNames.Length, yNames(idx), $"Y{idx + 1}")
                        Dim sp As Vector3 = ApplyRotation(New Vector3(xMin - zMargin + circleRadius, yVal, zMax), rotationValue)
                        Dim ep As Vector3 = ApplyRotation(New Vector3(xMax + zMargin - circleRadius, yVal, zMax), rotationValue)
                        dxfDoc.Entities.Add(New Line(sp, ep) With {.Layer = layerZ})
                        Dim c1 As Vector3 = ApplyRotation(New Vector3(xMin - zMargin, yVal, zMax), rotationValue)
                        Dim c2 As Vector3 = ApplyRotation(New Vector3(xMax + zMargin, yVal, zMax), rotationValue)
                        dxfDoc.Entities.Add(New Text(label, c1, textHeight) With {.Layer = layerZ, .Alignment = TextAlignment.MiddleCenter})
                        dxfDoc.Entities.Add(New Circle(c1, circleRadius) With {.Layer = layerZ})
                        dxfDoc.Entities.Add(New Text(label, c2, textHeight) With {.Layer = layerZ, .Alignment = TextAlignment.MiddleCenter})
                        dxfDoc.Entities.Add(New Circle(c2, circleRadius) With {.Layer = layerZ})
                    Next
                    For idx As Integer = 0 To xPositions.Length - 1
                        Dim xVal As Double = xPositions(idx)
                        Dim label As String = If(idx < xNames.Length, xNames(idx), $"X{idx + 1}")
                        Dim sp As Vector3 = ApplyRotation(New Vector3(xVal, yMin - zMargin + circleRadius, zMax), rotationValue)
                        Dim ep As Vector3 = ApplyRotation(New Vector3(xVal, yMax + zMargin - circleRadius, zMax), rotationValue)
                        dxfDoc.Entities.Add(New Line(sp, ep) With {.Layer = layerZ})
                        Dim c1 As Vector3 = ApplyRotation(New Vector3(xVal, yMin - zMargin, zMax), rotationValue)
                        Dim c2 As Vector3 = ApplyRotation(New Vector3(xVal, yMax + zMargin, zMax), rotationValue)
                        dxfDoc.Entities.Add(New Text(label, c1, textHeight) With {.Layer = layerZ, .Alignment = TextAlignment.MiddleCenter})
                        dxfDoc.Entities.Add(New Circle(c1, circleRadius) With {.Layer = layerZ})
                        dxfDoc.Entities.Add(New Text(label, c2, textHeight) With {.Layer = layerZ, .Alignment = TextAlignment.MiddleCenter})
                        dxfDoc.Entities.Add(New Circle(c2, circleRadius) With {.Layer = layerZ})
                    Next
                End If
            End If

            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    Private Function ParseOffsets(offsetString As String, ByRef includesZero As Boolean) As Double()
        includesZero = False
        Try
            If String.IsNullOrWhiteSpace(offsetString) Then Return New Double() {}
            Dim input As String = offsetString.Replace(","c, " "c).Replace(";"c, " "c).Replace(vbTab, " ").Trim()
            Dim tokens As String() = Regex.Split(input, "\s+")
            Dim list As New List(Of Double)()
            For Each token As String In tokens
                If String.IsNullOrWhiteSpace(token) Then Continue For
                If token.Contains("*") Then
                    Dim parts As String() = token.Split("*"c)
                    Dim count As Integer
                    Dim value As Double
                    If parts.Length <> 2 OrElse Not Integer.TryParse(parts(0).Trim(), count) OrElse count <= 0 OrElse
                       (Not Double.TryParse(parts(1).Trim(), NumberStyles.Float Or NumberStyles.AllowThousands, CultureInfo.InvariantCulture, value) AndAlso
                        Not Double.TryParse(parts(1).Trim(), NumberStyles.Float Or NumberStyles.AllowThousands, CultureInfo.CurrentCulture, value)) Then
                        Continue For
                    End If
                    For j As Integer = 1 To count
                        list.Add(value)
                        If Math.Abs(value) < 0.000000001 Then includesZero = True
                    Next
                Else
                    Dim parsed As Double
                    If Double.TryParse(token, NumberStyles.Float Or NumberStyles.AllowThousands, CultureInfo.InvariantCulture, parsed) OrElse
                       Double.TryParse(token, NumberStyles.Float Or NumberStyles.AllowThousands, CultureInfo.CurrentCulture, parsed) Then
                        list.Add(parsed)
                        If Math.Abs(parsed) < 0.000000001 Then includesZero = True
                    End If
                End If
            Next
            Return list.ToArray()
        Catch ex As Exception
            Return New Double() {}
        End Try
    End Function

    Private Function ParseNames(nameString As String) As String()
        Try
            If String.IsNullOrWhiteSpace(nameString) Then Return New String() {}
            Dim source As String() = nameString.Split(New Char() {" "c}, StringSplitOptions.RemoveEmptyEntries)
            Return source.Select(Function(p) p.Trim()).Where(Function(n) n <> "").ToArray()
        Catch ex As Exception
            Return New String() {}
        End Try
    End Function

    Private Function CalculatePositions(offsets As Double()) As Double()
        Try
            If offsets.Length = 0 Then Return New Double() {}
            Dim list As New List(Of Double)()
            Dim running As Double = offsets(0)
            list.Add(running)
            For i As Integer = 1 To offsets.Length - 1
                running += offsets(i)
                list.Add(running)
            Next
            Return list.Distinct().OrderBy(Function(x) x).ToArray()
        Catch ex As Exception
            Return New Double() {}
        End Try
    End Function

    Private Function GetLeaderDirectionAngle(sheet As Object) As Double?
        Select Case GetSafeCellValueAsString(sheet, 3, 45).Trim().ToUpper()
            Case "D"
                Return -Math.PI / 4.0
            Case "U"
                Return Math.PI / 4.0
            Case "N"
                Return Nothing
            Case Else
                Return Math.PI / 4.0
        End Select
    End Function

    Private Function ApplyRotation(point As Vector3, rotationValue As String) As Vector3
        Select Case If(String.IsNullOrWhiteSpace(rotationValue), "1", rotationValue.Trim())
            Case "2"
                Return New Vector3(-point.X, point.Y, -point.Z)
            Case "3"
                Return New Vector3(point.Z, point.Y, -point.X)
            Case "4"
                Return New Vector3(-point.Z, point.Y, point.X)
            Case "5"
                Return New Vector3(point.X, -point.Z, point.Y)
            Case Else
                Return point
        End Select
    End Function

    Private Function BuildSizeText(web As String, tf As String, bf As String, beamLength As Double, columnFValue As Double) As String
        web = If(web?.Replace("PL", "").Trim(), "")
        tf = If(tf?.Replace("PL", "").Trim(), "")
        bf = If(bf?.Replace("PL", "").Trim(), "")
        Dim list As New List(Of String)()
        If columnFValue = 680.0 Then
            If Not String.IsNullOrEmpty(web) Then list.Add($"W-{web}")
            If Not String.IsNullOrEmpty(tf) AndAlso Not String.IsNullOrEmpty(bf) Then
                list.Add(If(tf = bf, $"F-{tf}", $"TF-{tf}\PBF-{bf}"))
            ElseIf Not String.IsNullOrEmpty(tf) Then
                list.Add($"F-{tf}")
            ElseIf Not String.IsNullOrEmpty(bf) Then
                list.Add($"F-{bf}")
            End If
        Else
            If Not String.IsNullOrEmpty(web) Then list.Add(web)
            If Not String.IsNullOrEmpty(tf) AndAlso Not String.IsNullOrEmpty(bf) Then
                list.Add(If(tf = bf, tf, $"{tf}\P{bf}"))
            ElseIf Not String.IsNullOrEmpty(tf) Then
                list.Add(tf)
            ElseIf Not String.IsNullOrEmpty(bf) Then
                list.Add(bf)
            End If
        End If
        If beamLength > 0.0 Then list.Add($"L-{CInt(Math.Round(beamLength))}")
        Return String.Join("\P", list)
    End Function

    Private Function GetSafeCellValue(excelSheet As Object, row As Integer, col As Integer) As Double
        Try
            Dim cellValue As Object = CType(excelSheet, Worksheet).Cells(row, col).Value
            If cellValue Is Nothing Then Return 0.0
            Dim result As Double
            Return If(Double.TryParse(cellValue.ToString(), result), result, 0.0)
        Catch ex As Exception
            Return 0.0
        End Try
    End Function

    Private Function GetSafeCellValueAsString(excelSheet As Object, row As Integer, col As Integer) As String
        Try
            Dim cellValue As Object = CType(excelSheet, Worksheet).Cells(row, col).Value
            Return If(cellValue?.ToString(), "")
        Catch ex As Exception
            Return ""
        End Try
    End Function

    Private Function CreateColorBasedLayers(dxfDoc As DxfDocument, excelSheet As Object, lastRow As Integer) As Dictionary(Of String, Layer)
        Dim dictionary As New Dictionary(Of String, Layer)()
        Dim colorIndex As Integer = 1
        For i As Integer = 9 To lastRow
            Try
                Dim v1 As Double = GetSafeCellValue(excelSheet, i, 16)
                Dim v2 As Double = GetSafeCellValue(excelSheet, i, 17)
                Dim v3 As Double = GetSafeCellValue(excelSheet, i, 18)
                Dim v4 As Double = GetSafeCellValue(excelSheet, i, 19)
                Dim v5 As Double = GetSafeCellValue(excelSheet, i, 20)
                Dim v6 As Double = GetSafeCellValue(excelSheet, i, 21)
                If v1 = 0.0 AndAlso v2 = 0.0 AndAlso v3 = 0.0 AndAlso v4 = 0.0 AndAlso v5 = 0.0 AndAlso v6 = 0.0 Then
                    Continue For
                End If
                Dim signature As String = ""
                For col As Integer = 7 To 13
                    signature &= GetSafeCellValueAsString(excelSheet, i, col) & "|"
                Next
                If signature.EndsWith("|") Then signature = signature.Substring(0, signature.Length - 1)
                If Not dictionary.ContainsKey(signature) Then
                    Dim color As AciColor = memberColors((colorIndex - 1) Mod memberColors.Length)
                    Dim name As String = CreateLayerNameFromBeamText(excelSheet, i)
                    Dim layer As Layer
                    If dxfDoc.Layers.Contains(name) Then
                        layer = dxfDoc.Layers(name)
                    Else
                        layer = New Layer(name)
                        layer.Color = color
                        dxfDoc.Layers.Add(layer)
                    End If
                    dictionary.Add(signature, layer)
                    colorIndex += 1
                End If
            Catch ex As Exception
                ' Skip this row and continue
            End Try
        Next
        Return dictionary
    End Function

    Private Function GetBeamLayerByParameters(excelSheet As Object, rowIndex As Integer, parameterLayerMap As Dictionary(Of String, Layer), defaultLayer As Layer) As Layer
        Try
            Dim signature As String = ""
            For col As Integer = 7 To 13
                signature &= GetSafeCellValueAsString(excelSheet, rowIndex, col) & "|"
            Next
            If signature.EndsWith("|") Then signature = signature.Substring(0, signature.Length - 1)
            Return If(Not parameterLayerMap.ContainsKey(signature), defaultLayer, parameterLayerMap(signature))
        Catch ex As Exception
            Return defaultLayer
        End Try
    End Function

    Private Function CreateLayerNameFromBeamText(excelSheet As Object, rowIndex As Integer) As String
        Try
            Dim web As String = GetSafeCellValueAsString(excelSheet, rowIndex, 37)
            Dim tf As String = GetSafeCellValueAsString(excelSheet, rowIndex, 38)
            Dim bf As String = GetSafeCellValueAsString(excelSheet, rowIndex, 39)
            Dim sectionType As Double = GetSafeCellValue(excelSheet, rowIndex, 6)
            web = If(web?.Replace("PL", "").Trim(), "")
            tf = If(tf?.Replace("PL", "").Trim(), "")
            bf = If(bf?.Replace("PL", "").Trim(), "")
            Dim list As New List(Of String)()
            If sectionType = 680.0 Then
                If Not String.IsNullOrEmpty(web) Then list.Add($"W-{web}")
                If Not String.IsNullOrEmpty(tf) AndAlso Not String.IsNullOrEmpty(bf) Then
                    list.Add(If(tf = bf, $"F-{tf}", $"TF-{tf} BF-{bf}"))
                ElseIf Not String.IsNullOrEmpty(tf) Then
                    list.Add($"F-{tf}")
                ElseIf Not String.IsNullOrEmpty(bf) Then
                    list.Add($"F-{bf}")
                End If
            Else
                If Not String.IsNullOrEmpty(web) Then list.Add(web)
                If Not String.IsNullOrEmpty(tf) AndAlso Not String.IsNullOrEmpty(bf) Then
                    list.Add(If(tf = bf, tf, $"{tf}x{bf}"))
                ElseIf Not String.IsNullOrEmpty(tf) Then
                    list.Add(tf)
                ElseIf Not String.IsNullOrEmpty(bf) Then
                    list.Add(bf)
                End If
            End If
            Dim text As String = String.Join(" ", list.Take(2))
            If String.IsNullOrWhiteSpace(text) Then text = "Beam_" & rowIndex
            text = text.Replace("/", "_").Replace("\", "_").Replace(":", "_").
                Replace("*", "_").Replace("?", "_").Replace("""", "_").
                Replace("<", "_").Replace(">", "_").Replace("|", "_")
            Return text
        Catch ex As Exception
            Return "Beam_" & rowIndex
        End Try
    End Function

    Private Function RotateVectorAroundAxis(vector As Vector3, axis As Vector3, angleInDegrees As Double) As Vector3
        If Math.Abs(angleInDegrees) < 0.001 Then Return vector
        Dim radians As Double = angleInDegrees * Math.PI / 180.0
        Dim cosA As Double = Math.Cos(radians)
        Dim sinA As Double = Math.Sin(radians)
        Dim axisNorm As Vector3 = axis
        axisNorm.Normalize()
        Dim dot As Double = Vector3.DotProduct(vector, axisNorm)
        Dim cross As Vector3 = Vector3.CrossProduct(axisNorm, vector)
        Return vector * cosA + cross * sinA + axisNorm * dot * (1.0 - cosA)
    End Function

    Private Sub DrawTaperedISection(dxf As DxfDocument, startPoint As Vector3, endPoint As Vector3, webHeightStart As Double, webThicknessStart As Double, topFlangeWidthStart As Double, topFlangeThicknessStart As Double, bottomFlangeWidthStart As Double, bottomFlangeThicknessStart As Double, webHeightEnd As Double, webThicknessEnd As Double, topFlangeWidthEnd As Double, topFlangeThicknessEnd As Double, bottomFlangeWidthEnd As Double, bottomFlangeThicknessEnd As Double, color As AciColor, betaAngle As Double, rotationValue As String, targetLayer As Layer)
        Dim axis As Vector3 = endPoint - startPoint
        Dim axisNorm As Vector3 = axis
        axisNorm.Normalize()
        Const zThreshold As Double = 0.99
        Dim localY, localZ As Vector3
        If Math.Abs(axisNorm.Z) > zThreshold Then
            localY = New Vector3(1.0, 0.0, 0.0)
            localZ = Vector3.CrossProduct(axisNorm, localY)
            localZ.Normalize()
            localY = Vector3.CrossProduct(localZ, axisNorm)
            localY.Normalize()
        Else
            localZ = Vector3.CrossProduct(New Vector3(0.0, 0.0, 1.0), axisNorm)
            Dim mag As Double = Math.Sqrt(localZ.X * localZ.X + localZ.Y * localZ.Y + localZ.Z * localZ.Z)
            If mag < 0.001 Then
                localZ = New Vector3(0.0, 1.0, 0.0)
            Else
                localZ.Normalize()
            End If
            localY = Vector3.CrossProduct(localZ, axisNorm)
            localY.Normalize()
            localZ = Vector3.CrossProduct(axisNorm, localY)
            localZ.Normalize()
        End If
        If rotationValue.Trim() = "5" Then
            localY = RotateVectorAroundAxis(localY, axisNorm, -90.0)
            localZ = RotateVectorAroundAxis(localZ, axisNorm, -90.0)
        End If
        If rotationValue.Trim() = "3" Then
            If Math.Abs(axisNorm.X) > 0.7 OrElse Math.Abs(axisNorm.Z) > 0.7 Then
                localY = RotateVectorAroundAxis(localY, axisNorm, -90.0)
                localZ = RotateVectorAroundAxis(localZ, axisNorm, -90.0)
            End If
            localY = RotateVectorAroundAxis(localY, axisNorm, -90.0)
            localZ = RotateVectorAroundAxis(localZ, axisNorm, -90.0)
        End If
        If rotationValue.Trim() = "4" Then
            If Math.Abs(axisNorm.X) > 0.7 OrElse Math.Abs(axisNorm.Z) > 0.7 Then
                localY = RotateVectorAroundAxis(localY, axisNorm, -90.0)
                localZ = RotateVectorAroundAxis(localZ, axisNorm, -90.0)
            End If
            localY = RotateVectorAroundAxis(localY, axisNorm, -90.0)
            localZ = RotateVectorAroundAxis(localZ, axisNorm, -90.0)
        End If
        If Math.Abs(betaAngle) > 0.001 Then
            localY = RotateVectorAroundAxis(localY, axisNorm, betaAngle)
            localZ = RotateVectorAroundAxis(localZ, axisNorm, betaAngle)
        End If
        Dim pointAt = Function(offsetY As Double, offsetZ As Double, basePoint As Vector3) basePoint + localY * offsetY + localZ * offsetZ

        Dim webSP1 As Vector3 = pointAt(-webThicknessStart / 2.0, webHeightStart / 2.0, startPoint)
        Dim webSP2 As Vector3 = pointAt(webThicknessStart / 2.0, webHeightStart / 2.0, startPoint)
        Dim webSP3 As Vector3 = pointAt(-webThicknessStart / 2.0, -webHeightStart / 2.0, startPoint)
        Dim webSP4 As Vector3 = pointAt(webThicknessStart / 2.0, -webHeightStart / 2.0, startPoint)
        Dim webEP1 As Vector3 = pointAt(-webThicknessEnd / 2.0, webHeightEnd / 2.0, endPoint)
        Dim webEP2 As Vector3 = pointAt(webThicknessEnd / 2.0, webHeightEnd / 2.0, endPoint)
        Dim webEP3 As Vector3 = pointAt(-webThicknessEnd / 2.0, -webHeightEnd / 2.0, endPoint)
        Dim webEP4 As Vector3 = pointAt(webThicknessEnd / 2.0, -webHeightEnd / 2.0, endPoint)
        dxf.Entities.Add(New Line(webSP1, webEP1) With {.Color = color, .Layer = targetLayer})
        dxf.Entities.Add(New Line(webSP2, webEP2) With {.Color = color, .Layer = targetLayer})
        dxf.Entities.Add(New Line(webSP3, webEP3) With {.Color = color, .Layer = targetLayer})
        dxf.Entities.Add(New Line(webSP4, webEP4) With {.Color = color, .Layer = targetLayer})

        Dim tfSP1 As Vector3 = pointAt(-topFlangeWidthStart / 2.0, webHeightStart / 2.0 + topFlangeThicknessStart, startPoint)
        Dim tfSP2 As Vector3 = pointAt(topFlangeWidthStart / 2.0, webHeightStart / 2.0 + topFlangeThicknessStart, startPoint)
        Dim tfSP3 As Vector3 = pointAt(-topFlangeWidthStart / 2.0, webHeightStart / 2.0, startPoint)
        Dim tfSP4 As Vector3 = pointAt(topFlangeWidthStart / 2.0, webHeightStart / 2.0, startPoint)
        Dim tfEP1 As Vector3 = pointAt(-topFlangeWidthEnd / 2.0, webHeightEnd / 2.0 + topFlangeThicknessEnd, endPoint)
        Dim tfEP2 As Vector3 = pointAt(topFlangeWidthEnd / 2.0, webHeightEnd / 2.0 + topFlangeThicknessEnd, endPoint)
        Dim tfEP3 As Vector3 = pointAt(-topFlangeWidthEnd / 2.0, webHeightEnd / 2.0, endPoint)
        Dim tfEP4 As Vector3 = pointAt(topFlangeWidthEnd / 2.0, webHeightEnd / 2.0, endPoint)
        dxf.Entities.Add(New Line(tfSP1, tfEP1) With {.Color = color, .Layer = targetLayer})
        dxf.Entities.Add(New Line(tfSP2, tfEP2) With {.Color = color, .Layer = targetLayer})
        dxf.Entities.Add(New Line(tfSP3, tfEP3) With {.Color = color, .Layer = targetLayer})
        dxf.Entities.Add(New Line(tfSP4, tfEP4) With {.Color = color, .Layer = targetLayer})

        Dim bfSP1 As Vector3 = pointAt(-bottomFlangeWidthStart / 2.0, -webHeightStart / 2.0 - bottomFlangeThicknessStart, startPoint)
        Dim bfSP2 As Vector3 = pointAt(bottomFlangeWidthStart / 2.0, -webHeightStart / 2.0 - bottomFlangeThicknessStart, startPoint)
        Dim bfSP3 As Vector3 = pointAt(-bottomFlangeWidthStart / 2.0, -webHeightStart / 2.0, startPoint)
        Dim bfSP4 As Vector3 = pointAt(bottomFlangeWidthStart / 2.0, -webHeightStart / 2.0, startPoint)
        Dim bfEP1 As Vector3 = pointAt(-bottomFlangeWidthEnd / 2.0, -webHeightEnd / 2.0 - bottomFlangeThicknessEnd, endPoint)
        Dim bfEP2 As Vector3 = pointAt(bottomFlangeWidthEnd / 2.0, -webHeightEnd / 2.0 - bottomFlangeThicknessEnd, endPoint)
        Dim bfEP3 As Vector3 = pointAt(-bottomFlangeWidthEnd / 2.0, -webHeightEnd / 2.0, endPoint)
        Dim bfEP4 As Vector3 = pointAt(bottomFlangeWidthEnd / 2.0, -webHeightEnd / 2.0, endPoint)
        dxf.Entities.Add(New Line(bfSP1, bfEP1) With {.Color = color, .Layer = targetLayer})
        dxf.Entities.Add(New Line(bfSP2, bfEP2) With {.Color = color, .Layer = targetLayer})
        dxf.Entities.Add(New Line(bfSP3, bfEP3) With {.Color = color, .Layer = targetLayer})
        dxf.Entities.Add(New Line(bfSP4, bfEP4) With {.Color = color, .Layer = targetLayer})

        DrawIBeamFace(dxf, startPoint, localY, localZ, webHeightStart, webThicknessStart, topFlangeWidthStart, topFlangeThicknessStart, bottomFlangeWidthStart, bottomFlangeThicknessStart, color, targetLayer)
        DrawIBeamFace(dxf, endPoint, localY, localZ, webHeightEnd, webThicknessEnd, topFlangeWidthEnd, topFlangeThicknessEnd, bottomFlangeWidthEnd, bottomFlangeThicknessEnd, color, targetLayer)
    End Sub

    Private Sub DrawIBeamFace(dxf As DxfDocument, centerPoint As Vector3, localY As Vector3, localZ As Vector3, webHeight As Double, webThickness As Double, topFlangeWidth As Double, topFlangeThickness As Double, bottomFlangeWidth As Double, bottomFlangeThickness As Double, color As AciColor, targetLayer As Layer)
        Dim pointAt = Function(offsetY As Double, offsetZ As Double) centerPoint + localY * offsetY + localZ * offsetZ
        Dim webLeft As Double = -webThickness / 2.0
        Dim webRight As Double = webThickness / 2.0
        Dim webTop As Double = webHeight / 2.0
        Dim webBottom As Double = -webHeight / 2.0
        Dim tfLeft As Double = -topFlangeWidth / 2.0
        Dim tfRight As Double = topFlangeWidth / 2.0
        Dim tfOuterZ As Double = webTop + topFlangeThickness
        Dim tfInnerZ As Double = webTop
        Dim bfLeft As Double = -bottomFlangeWidth / 2.0
        Dim bfRight As Double = bottomFlangeWidth / 2.0
        Dim bfInnerZ As Double = webBottom
        Dim bfOuterZ As Double = webBottom - bottomFlangeThickness
        dxf.Entities.Add(New Line(pointAt(tfLeft, tfOuterZ), pointAt(tfRight, tfOuterZ)) With {.Color = color, .Layer = targetLayer})
        dxf.Entities.Add(New Line(pointAt(tfRight, tfOuterZ), pointAt(tfRight, tfInnerZ)) With {.Color = color, .Layer = targetLayer})
        dxf.Entities.Add(New Line(pointAt(tfRight, tfInnerZ), pointAt(webRight, tfInnerZ)) With {.Color = color, .Layer = targetLayer})
        dxf.Entities.Add(New Line(pointAt(webRight, tfInnerZ), pointAt(webRight, bfInnerZ)) With {.Color = color, .Layer = targetLayer})
        dxf.Entities.Add(New Line(pointAt(webRight, bfInnerZ), pointAt(bfRight, bfInnerZ)) With {.Color = color, .Layer = targetLayer})
        dxf.Entities.Add(New Line(pointAt(bfRight, bfInnerZ), pointAt(bfRight, bfOuterZ)) With {.Color = color, .Layer = targetLayer})
        dxf.Entities.Add(New Line(pointAt(bfRight, bfOuterZ), pointAt(bfLeft, bfOuterZ)) With {.Color = color, .Layer = targetLayer})
        dxf.Entities.Add(New Line(pointAt(bfLeft, bfOuterZ), pointAt(bfLeft, bfInnerZ)) With {.Color = color, .Layer = targetLayer})
        dxf.Entities.Add(New Line(pointAt(bfLeft, bfInnerZ), pointAt(webLeft, bfInnerZ)) With {.Color = color, .Layer = targetLayer})
        dxf.Entities.Add(New Line(pointAt(webLeft, bfInnerZ), pointAt(webLeft, tfInnerZ)) With {.Color = color, .Layer = targetLayer})
        dxf.Entities.Add(New Line(pointAt(webLeft, tfInnerZ), pointAt(tfLeft, tfInnerZ)) With {.Color = color, .Layer = targetLayer})
        dxf.Entities.Add(New Line(pointAt(tfLeft, tfInnerZ), pointAt(tfLeft, tfOuterZ)) With {.Color = color, .Layer = targetLayer})
    End Sub

    Private Sub DrawTaperedChannelSection(dxf As DxfDocument, startPoint As Vector3, endPoint As Vector3, webHeightStart As Double, webThicknessStart As Double, topFlangeWidthStart As Double, topFlangeThicknessStart As Double, bottomFlangeWidthStart As Double, bottomFlangeThicknessStart As Double, webHeightEnd As Double, webThicknessEnd As Double, topFlangeWidthEnd As Double, topFlangeThicknessEnd As Double, bottomFlangeWidthEnd As Double, bottomFlangeThicknessEnd As Double, color As AciColor, betaAngle As Double, rotationValue As String, targetLayer As Layer)
        Dim axis As Vector3 = endPoint - startPoint
        Dim axisNorm As Vector3 = axis
        axisNorm.Normalize()
        Const zThreshold As Double = 0.99
        Dim localY, localZ As Vector3
        If Math.Abs(axisNorm.Z) > zThreshold Then
            localY = New Vector3(1.0, 0.0, 0.0)
            localZ = Vector3.CrossProduct(axisNorm, localY)
            localZ.Normalize()
            localY = Vector3.CrossProduct(localZ, axisNorm)
            localY.Normalize()
        Else
            localZ = Vector3.CrossProduct(New Vector3(0.0, 0.0, 1.0), axisNorm)
            Dim mag As Double = Math.Sqrt(localZ.X * localZ.X + localZ.Y * localZ.Y + localZ.Z * localZ.Z)
            If mag < 0.001 Then
                localZ = New Vector3(0.0, 1.0, 0.0)
            Else
                localZ.Normalize()
            End If
            localY = Vector3.CrossProduct(localZ, axisNorm)
            localY.Normalize()
            localZ = Vector3.CrossProduct(axisNorm, localY)
            localZ.Normalize()
        End If
        If rotationValue.Trim() = "5" Then
            localY = RotateVectorAroundAxis(localY, axisNorm, -90.0)
            localZ = RotateVectorAroundAxis(localZ, axisNorm, -90.0)
        End If
        If rotationValue.Trim() = "3" Then
            If Math.Abs(axisNorm.X) > 0.7 OrElse Math.Abs(axisNorm.Z) > 0.7 Then
                localY = RotateVectorAroundAxis(localY, axisNorm, -90.0)
                localZ = RotateVectorAroundAxis(localZ, axisNorm, -90.0)
            End If
            localY = RotateVectorAroundAxis(localY, axisNorm, -90.0)
            localZ = RotateVectorAroundAxis(localZ, axisNorm, -90.0)
        End If
        If rotationValue.Trim() = "4" Then
            If Math.Abs(axisNorm.X) > 0.7 OrElse Math.Abs(axisNorm.Z) > 0.7 Then
                localY = RotateVectorAroundAxis(localY, axisNorm, -90.0)
                localZ = RotateVectorAroundAxis(localZ, axisNorm, -90.0)
            End If
            localY = RotateVectorAroundAxis(localY, axisNorm, -90.0)
            localZ = RotateVectorAroundAxis(localZ, axisNorm, -90.0)
        End If
        If Math.Abs(betaAngle) > 0.001 Then
            localY = RotateVectorAroundAxis(localY, axisNorm, betaAngle)
            localZ = RotateVectorAroundAxis(localZ, axisNorm, betaAngle)
        End If
        Dim webCenterYStart As Double = webThicknessStart / 2.0
        Dim webCenterYEnd As Double = webThicknessEnd / 2.0
        Dim pointAt = Function(offsetY As Double, offsetZ As Double, basePoint As Vector3, center As Double) basePoint + localY * (offsetY - center) + localZ * offsetZ

        Dim webSP1 As Vector3 = pointAt(0.0, webHeightStart / 2.0, startPoint, webCenterYStart)
        Dim webSP2 As Vector3 = pointAt(webThicknessStart, webHeightStart / 2.0, startPoint, webCenterYStart)
        Dim webSP3 As Vector3 = pointAt(0.0, -webHeightStart / 2.0, startPoint, webCenterYStart)
        Dim webSP4 As Vector3 = pointAt(webThicknessStart, -webHeightStart / 2.0, startPoint, webCenterYStart)
        Dim webEP1 As Vector3 = pointAt(0.0, webHeightEnd / 2.0, endPoint, webCenterYEnd)
        Dim webEP2 As Vector3 = pointAt(webThicknessEnd, webHeightEnd / 2.0, endPoint, webCenterYEnd)
        Dim webEP3 As Vector3 = pointAt(0.0, -webHeightEnd / 2.0, endPoint, webCenterYEnd)
        Dim webEP4 As Vector3 = pointAt(webThicknessEnd, -webHeightEnd / 2.0, endPoint, webCenterYEnd)
        dxf.Entities.Add(New Line(webSP1, webEP1) With {.Color = color, .Layer = targetLayer})
        dxf.Entities.Add(New Line(webSP2, webEP2) With {.Color = color, .Layer = targetLayer})
        dxf.Entities.Add(New Line(webSP3, webEP3) With {.Color = color, .Layer = targetLayer})
        dxf.Entities.Add(New Line(webSP4, webEP4) With {.Color = color, .Layer = targetLayer})

        Dim tfSP1 As Vector3 = pointAt(0.0, webHeightStart / 2.0 + topFlangeThicknessStart, startPoint, webCenterYStart)
        Dim tfSP2 As Vector3 = pointAt(topFlangeWidthStart, webHeightStart / 2.0 + topFlangeThicknessStart, startPoint, webCenterYStart)
        Dim tfSP3 As Vector3 = pointAt(0.0, webHeightStart / 2.0, startPoint, webCenterYStart)
        Dim tfSP4 As Vector3 = pointAt(topFlangeWidthStart, webHeightStart / 2.0, startPoint, webCenterYStart)
        Dim tfEP1 As Vector3 = pointAt(0.0, webHeightEnd / 2.0 + topFlangeThicknessEnd, endPoint, webCenterYEnd)
        Dim tfEP2 As Vector3 = pointAt(topFlangeWidthEnd, webHeightEnd / 2.0 + topFlangeThicknessEnd, endPoint, webCenterYEnd)
        Dim tfEP3 As Vector3 = pointAt(0.0, webHeightEnd / 2.0, endPoint, webCenterYEnd)
        Dim tfEP4 As Vector3 = pointAt(topFlangeWidthEnd, webHeightEnd / 2.0, endPoint, webCenterYEnd)
        dxf.Entities.Add(New Line(tfSP1, tfEP1) With {.Color = color, .Layer = targetLayer})
        dxf.Entities.Add(New Line(tfSP2, tfEP2) With {.Color = color, .Layer = targetLayer})
        dxf.Entities.Add(New Line(tfSP3, tfEP3) With {.Color = color, .Layer = targetLayer})
        dxf.Entities.Add(New Line(tfSP4, tfEP4) With {.Color = color, .Layer = targetLayer})

        Dim bfSP1 As Vector3 = pointAt(0.0, -webHeightStart / 2.0 - bottomFlangeThicknessStart, startPoint, webCenterYStart)
        Dim bfSP2 As Vector3 = pointAt(bottomFlangeWidthStart, -webHeightStart / 2.0 - bottomFlangeThicknessStart, startPoint, webCenterYStart)
        Dim bfSP3 As Vector3 = pointAt(0.0, -webHeightStart / 2.0, startPoint, webCenterYStart)
        Dim bfSP4 As Vector3 = pointAt(bottomFlangeWidthStart, -webHeightStart / 2.0, startPoint, webCenterYStart)
        Dim bfEP1 As Vector3 = pointAt(0.0, -webHeightEnd / 2.0 - bottomFlangeThicknessEnd, endPoint, webCenterYEnd)
        Dim bfEP2 As Vector3 = pointAt(bottomFlangeWidthEnd, -webHeightEnd / 2.0 - bottomFlangeThicknessEnd, endPoint, webCenterYEnd)
        Dim bfEP3 As Vector3 = pointAt(0.0, -webHeightEnd / 2.0, endPoint, webCenterYEnd)
        Dim bfEP4 As Vector3 = pointAt(bottomFlangeWidthEnd, -webHeightEnd / 2.0, endPoint, webCenterYEnd)
        dxf.Entities.Add(New Line(bfSP1, bfEP1) With {.Color = color, .Layer = targetLayer})
        dxf.Entities.Add(New Line(bfSP2, bfEP2) With {.Color = color, .Layer = targetLayer})
        dxf.Entities.Add(New Line(bfSP3, bfEP3) With {.Color = color, .Layer = targetLayer})
        dxf.Entities.Add(New Line(bfSP4, bfEP4) With {.Color = color, .Layer = targetLayer})

        DrawChannelFace(dxf, startPoint, localY, localZ, webHeightStart, webThicknessStart, topFlangeWidthStart, topFlangeThicknessStart, bottomFlangeWidthStart, bottomFlangeThicknessStart, color, targetLayer, webCenterYStart)
        DrawChannelFace(dxf, endPoint, localY, localZ, webHeightEnd, webThicknessEnd, topFlangeWidthEnd, topFlangeThicknessEnd, bottomFlangeWidthEnd, bottomFlangeThicknessEnd, color, targetLayer, webCenterYEnd)
    End Sub

    Private Sub DrawChannelFace(dxf As DxfDocument, centerPoint As Vector3, localY As Vector3, localZ As Vector3, webHeight As Double, webThickness As Double, topFlangeWidth As Double, topFlangeThickness As Double, bottomFlangeWidth As Double, bottomFlangeThickness As Double, color As AciColor, targetLayer As Layer, webCenterY As Double)
        Dim pointAt = Function(offsetY As Double, offsetZ As Double) centerPoint + localY * (offsetY - webCenterY) + localZ * offsetZ
        Dim left As Double = 0
        Dim webTop As Double = webHeight / 2.0
        Dim webBottom As Double = -webHeight / 2.0
        Dim innerLeft As Double = 0
        Dim tfOuterZ As Double = webTop + topFlangeThickness
        Dim tfInnerZ As Double = webTop
        Dim innerLeft2 As Double = 0
        Dim bfInnerZ As Double = webBottom
        Dim bfOuterZ As Double = webBottom - bottomFlangeThickness
        dxf.Entities.Add(New Line(pointAt(innerLeft, tfOuterZ), pointAt(topFlangeWidth, tfOuterZ)) With {.Color = color, .Layer = targetLayer})
        dxf.Entities.Add(New Line(pointAt(topFlangeWidth, tfOuterZ), pointAt(topFlangeWidth, tfInnerZ)) With {.Color = color, .Layer = targetLayer})
        dxf.Entities.Add(New Line(pointAt(topFlangeWidth, tfInnerZ), pointAt(webThickness, tfInnerZ)) With {.Color = color, .Layer = targetLayer})
        dxf.Entities.Add(New Line(pointAt(webThickness, tfInnerZ), pointAt(webThickness, bfInnerZ)) With {.Color = color, .Layer = targetLayer})
        dxf.Entities.Add(New Line(pointAt(webThickness, bfInnerZ), pointAt(bottomFlangeWidth, bfInnerZ)) With {.Color = color, .Layer = targetLayer})
        dxf.Entities.Add(New Line(pointAt(bottomFlangeWidth, bfInnerZ), pointAt(bottomFlangeWidth, bfOuterZ)) With {.Color = color, .Layer = targetLayer})
        dxf.Entities.Add(New Line(pointAt(bottomFlangeWidth, bfOuterZ), pointAt(innerLeft2, bfOuterZ)) With {.Color = color, .Layer = targetLayer})
        dxf.Entities.Add(New Line(pointAt(innerLeft2, bfOuterZ), pointAt(innerLeft2, bfInnerZ)) With {.Color = color, .Layer = targetLayer})
        dxf.Entities.Add(New Line(pointAt(innerLeft2, bfInnerZ), pointAt(left, bfInnerZ)) With {.Color = color, .Layer = targetLayer})
        dxf.Entities.Add(New Line(pointAt(left, bfInnerZ), pointAt(left, tfInnerZ)) With {.Color = color, .Layer = targetLayer})
        dxf.Entities.Add(New Line(pointAt(left, tfInnerZ), pointAt(innerLeft, tfInnerZ)) With {.Color = color, .Layer = targetLayer})
        dxf.Entities.Add(New Line(pointAt(innerLeft, tfInnerZ), pointAt(innerLeft, tfOuterZ)) With {.Color = color, .Layer = targetLayer})
    End Sub

    Private Sub DrawHollowRectangleTube(dxf As DxfDocument, startPoint As Vector3, endPoint As Vector3, width As Double, height As Double, thickness As Double, color As AciColor, betaAngle As Double, rotationValue As String, targetLayer As Layer)
        Dim axis As Vector3 = endPoint - startPoint
        Dim axisNorm As Vector3 = axis
        axisNorm.Normalize()
        Const zThreshold As Double = 0.99
        Dim localZ, localY As Vector3
        If Math.Abs(axisNorm.Z) > zThreshold Then
            localY = Vector3.CrossProduct(New Vector3(1.0, 0.0, 0.0), axisNorm)
            localY.Normalize()
            localZ = Vector3.CrossProduct(localY, axisNorm)
            localZ.Normalize()
        Else
            localY = Vector3.CrossProduct(New Vector3(0.0, 0.0, 1.0), axisNorm)
            Dim mag As Double = Math.Sqrt(localY.X * localY.X + localY.Y * localY.Y + localY.Z * localY.Z)
            If mag < 0.001 Then
                localY = New Vector3(0.0, 1.0, 0.0)
            Else
                localY.Normalize()
            End If
            localZ = Vector3.CrossProduct(localY, axisNorm)
            localZ.Normalize()
            localY = Vector3.CrossProduct(axisNorm, localZ)
            localY.Normalize()
        End If
        If rotationValue.Trim() = "5" Then
            localZ = RotateVectorAroundAxis(localZ, axisNorm, -90.0)
            localY = RotateVectorAroundAxis(localY, axisNorm, -90.0)
        End If
        If rotationValue.Trim() = "3" Then
            If Math.Abs(axisNorm.X) > 0.7 OrElse Math.Abs(axisNorm.Z) > 0.7 Then
                localZ = RotateVectorAroundAxis(localZ, axisNorm, -90.0)
                localY = RotateVectorAroundAxis(localY, axisNorm, -90.0)
            End If
            localZ = RotateVectorAroundAxis(localZ, axisNorm, -90.0)
            localY = RotateVectorAroundAxis(localY, axisNorm, -90.0)
        End If
        If rotationValue.Trim() = "4" Then
            If Math.Abs(axisNorm.X) > 0.7 OrElse Math.Abs(axisNorm.Z) > 0.7 Then
                localZ = RotateVectorAroundAxis(localZ, axisNorm, -90.0)
                localY = RotateVectorAroundAxis(localY, axisNorm, -90.0)
            End If
            localZ = RotateVectorAroundAxis(localZ, axisNorm, -90.0)
            localY = RotateVectorAroundAxis(localY, axisNorm, -90.0)
        End If
        If Math.Abs(betaAngle) > 0.001 Then
            localZ = RotateVectorAroundAxis(localZ, axisNorm, betaAngle)
            localY = RotateVectorAroundAxis(localY, axisNorm, betaAngle)
        End If
        Dim halfWidth As Double = width / 2.0
        Dim halfHeight As Double = height / 2.0
        Dim innerHalfWidth As Double = halfWidth - thickness
        Dim innerHalfHeight As Double = halfHeight - thickness
        DrawRectangularTube3D(dxf, startPoint, endPoint, localZ, localY, halfWidth, halfHeight, color, targetLayer)
        If innerHalfWidth > 0.0 AndAlso innerHalfHeight > 0.0 Then
            DrawRectangularTube3D(dxf, startPoint, endPoint, localZ, localY, innerHalfWidth, innerHalfHeight, color, targetLayer)
        End If
    End Sub

    Private Sub DrawRectangularTube3D(dxf As DxfDocument, startPoint As Vector3, endPoint As Vector3, localY As Vector3, localZ As Vector3, halfWidth As Double, halfHeight As Double, color As AciColor, targetLayer As Layer)
        Dim pointAt = Function(offsetY As Double, offsetZ As Double, basePoint As Vector3) basePoint + localY * offsetY + localZ * offsetZ
        Dim p1 As Vector3 = pointAt(halfHeight, halfWidth, startPoint)
        Dim p2 As Vector3 = pointAt(-halfHeight, halfWidth, startPoint)
        Dim p3 As Vector3 = pointAt(-halfHeight, -halfWidth, startPoint)
        Dim p4 As Vector3 = pointAt(halfHeight, -halfWidth, startPoint)
        Dim p5 As Vector3 = pointAt(halfHeight, halfWidth, endPoint)
        Dim p6 As Vector3 = pointAt(-halfHeight, halfWidth, endPoint)
        Dim p7 As Vector3 = pointAt(-halfHeight, -halfWidth, endPoint)
        Dim p8 As Vector3 = pointAt(halfHeight, -halfWidth, endPoint)
        DrawLine3D(dxf, p1, p5, color, targetLayer)
        DrawLine3D(dxf, p2, p6, color, targetLayer)
        DrawLine3D(dxf, p3, p7, color, targetLayer)
        DrawLine3D(dxf, p4, p8, color, targetLayer)
        DrawLine3D(dxf, p1, p2, color, targetLayer)
        DrawLine3D(dxf, p2, p3, color, targetLayer)
        DrawLine3D(dxf, p3, p4, color, targetLayer)
        DrawLine3D(dxf, p4, p1, color, targetLayer)
        DrawLine3D(dxf, p5, p6, color, targetLayer)
        DrawLine3D(dxf, p6, p7, color, targetLayer)
        DrawLine3D(dxf, p7, p8, color, targetLayer)
        DrawLine3D(dxf, p8, p5, color, targetLayer)
    End Sub

    Private Sub DrawTubeSection(dxf As DxfDocument, startPoint As Vector3, endPoint As Vector3, width As Double, thickness As Double, color As AciColor, betaAngle As Double, rotationValue As String, targetLayer As Layer)
        Dim axis As Vector3 = endPoint - startPoint
        Dim axisNorm As Vector3 = axis
        axisNorm.Normalize()
        Const zThreshold As Double = 0.99
        Dim localZ, localY As Vector3
        If Math.Abs(axisNorm.Z) > zThreshold Then
            localY = Vector3.CrossProduct(New Vector3(1.0, 0.0, 0.0), axisNorm)
            localY.Normalize()
            localZ = Vector3.CrossProduct(localY, axisNorm)
            localZ.Normalize()
        Else
            localY = Vector3.CrossProduct(New Vector3(0.0, 0.0, 1.0), axisNorm)
            Dim mag As Double = Math.Sqrt(localY.X * localY.X + localY.Y * localY.Y + localY.Z * localY.Z)
            If mag < 0.001 Then
                localY = New Vector3(0.0, 1.0, 0.0)
            Else
                localY.Normalize()
            End If
            localZ = Vector3.CrossProduct(localY, axisNorm)
            localZ.Normalize()
            localY = Vector3.CrossProduct(axisNorm, localZ)
            localY.Normalize()
        End If
        If rotationValue.Trim() = "5" Then
            localZ = RotateVectorAroundAxis(localZ, axisNorm, -90.0)
            localY = RotateVectorAroundAxis(localY, axisNorm, -90.0)
        End If
        If rotationValue.Trim() = "3" Then
            If Math.Abs(axisNorm.X) > 0.7 OrElse Math.Abs(axisNorm.Z) > 0.7 Then
                localZ = RotateVectorAroundAxis(localZ, axisNorm, -90.0)
                localY = RotateVectorAroundAxis(localY, axisNorm, -90.0)
            End If
            localZ = RotateVectorAroundAxis(localZ, axisNorm, -90.0)
            localY = RotateVectorAroundAxis(localY, axisNorm, -90.0)
        End If
        If rotationValue.Trim() = "4" Then
            If Math.Abs(axisNorm.X) > 0.7 OrElse Math.Abs(axisNorm.Z) > 0.7 Then
                localZ = RotateVectorAroundAxis(localZ, axisNorm, -90.0)
                localY = RotateVectorAroundAxis(localY, axisNorm, -90.0)
            End If
            localZ = RotateVectorAroundAxis(localZ, axisNorm, -90.0)
            localY = RotateVectorAroundAxis(localY, axisNorm, -90.0)
        End If
        If Math.Abs(betaAngle) > 0.001 Then
            localZ = RotateVectorAroundAxis(localZ, axisNorm, betaAngle)
            localY = RotateVectorAroundAxis(localY, axisNorm, betaAngle)
        End If
        Dim halfWidth As Double = width / 2.0
        Dim innerHalfWidth As Double = halfWidth - thickness
        DrawSquareTube3D(dxf, startPoint, endPoint, localZ, localY, halfWidth, color, targetLayer)
        If innerHalfWidth > 0.0 Then
            DrawSquareTube3D(dxf, startPoint, endPoint, localZ, localY, innerHalfWidth, color, targetLayer)
        End If
    End Sub

    Private Sub DrawSquareTube3D(dxf As DxfDocument, startPoint As Vector3, endPoint As Vector3, localY As Vector3, localZ As Vector3, halfWidth As Double, color As AciColor, targetLayer As Layer)
        Dim pointAt = Function(offsetY As Double, offsetZ As Double, basePoint As Vector3) basePoint + localY * offsetY + localZ * offsetZ
        Dim p1 As Vector3 = pointAt(halfWidth, halfWidth, startPoint)
        Dim p2 As Vector3 = pointAt(-halfWidth, halfWidth, startPoint)
        Dim p3 As Vector3 = pointAt(-halfWidth, -halfWidth, startPoint)
        Dim p4 As Vector3 = pointAt(halfWidth, -halfWidth, startPoint)
        Dim p5 As Vector3 = pointAt(halfWidth, halfWidth, endPoint)
        Dim p6 As Vector3 = pointAt(-halfWidth, halfWidth, endPoint)
        Dim p7 As Vector3 = pointAt(-halfWidth, -halfWidth, endPoint)
        Dim p8 As Vector3 = pointAt(halfWidth, -halfWidth, endPoint)
        DrawLine3D(dxf, p1, p5, color, targetLayer)
        DrawLine3D(dxf, p2, p6, color, targetLayer)
        DrawLine3D(dxf, p3, p7, color, targetLayer)
        DrawLine3D(dxf, p4, p8, color, targetLayer)
        DrawLine3D(dxf, p1, p2, color, targetLayer)
        DrawLine3D(dxf, p2, p3, color, targetLayer)
        DrawLine3D(dxf, p3, p4, color, targetLayer)
        DrawLine3D(dxf, p4, p1, color, targetLayer)
        DrawLine3D(dxf, p5, p6, color, targetLayer)
        DrawLine3D(dxf, p6, p7, color, targetLayer)
        DrawLine3D(dxf, p7, p8, color, targetLayer)
        DrawLine3D(dxf, p8, p5, color, targetLayer)
    End Sub

    Private Sub DrawPipeSection(dxf As DxfDocument, startPoint As Vector3, endPoint As Vector3, outerDiameter As Double, thickness As Double, color As AciColor, betaAngle As Double, rotationValue As String, targetLayer As Layer)
        If outerDiameter <= 0.0 OrElse thickness <= 0.0 Then Return
        Dim innerDiameter As Double = outerDiameter - 2.0 * thickness
        If innerDiameter <= 0.0 Then
            DrawSolidPipeSection(dxf, startPoint, endPoint, outerDiameter, color, betaAngle, rotationValue, targetLayer)
            Return
        End If
        DrawSolidPipeSection(dxf, startPoint, endPoint, outerDiameter, color, betaAngle, rotationValue, targetLayer)
        DrawSolidPipeSection(dxf, startPoint, endPoint, innerDiameter, color, betaAngle, rotationValue, targetLayer)
    End Sub

    Private Sub DrawSolidPipeSection(dxf As DxfDocument, startPoint As Vector3, endPoint As Vector3, diameter As Double, color As AciColor, betaAngle As Double, rotationValue As String, targetLayer As Layer)
        Dim axis As Vector3 = endPoint - startPoint
        Dim axisNorm As Vector3 = axis
        axisNorm.Normalize()
        Dim localY As Vector3 = Vector3.CrossProduct(New Vector3(0.0, 0.0, 1.0), axisNorm)
        If localY.Modulus() < 0.001 Then
            localY = New Vector3(0.0, 1.0, 0.0)
        Else
            localY.Normalize()
        End If
        Dim localZ As Vector3 = Vector3.CrossProduct(localY, axisNorm)
        localZ.Normalize()
        If Math.Abs(betaAngle) > 0.001 Then
            localZ = RotateVectorAroundAxis(localZ, axisNorm, betaAngle)
            localY = RotateVectorAroundAxis(localY, axisNorm, betaAngle)
        End If
        Dim segments As Integer = 12
        Dim radius As Double = diameter / 2.0
        Dim angleStep As Double = Math.PI * 2.0 / CDbl(segments)
        DrawCircle3D(dxf, startPoint, localZ, localY, radius, segments, color, targetLayer)
        DrawCircle3D(dxf, endPoint, localZ, localY, radius, segments, color, targetLayer)
        For i As Integer = 0 To segments - 1
            Dim angle As Double = CDbl(i) * angleStep
            Dim offY As Double = radius * Math.Cos(angle)
            Dim offZ As Double = radius * Math.Sin(angle)
            Dim sp As Vector3 = startPoint + localZ * offY + localY * offZ
            Dim ep As Vector3 = endPoint + localZ * offY + localY * offZ
            dxf.Entities.Add(New Line(sp, ep) With {.Color = color, .Layer = targetLayer})
        Next
    End Sub

    Private Sub DrawCircle3D(dxf As DxfDocument, center As Vector3, axisY As Vector3, axisZ As Vector3, radius As Double, segments As Integer, color As AciColor, targetLayer As Layer)
        Dim angleStep As Double = Math.PI * 2.0 / CDbl(segments)
        Dim haveFirst As Boolean = False
        Dim previousPoint As Vector3 = Nothing
        Dim firstPoint As Vector3 = Nothing
        For i As Integer = 0 To segments
            Dim angle As Double = CDbl(i) * angleStep
            Dim offY As Double = radius * Math.Cos(angle)
            Dim offZ As Double = radius * Math.Sin(angle)
            Dim point As Vector3 = center + axisY * offY + axisZ * offZ
            If i > 0 Then
                dxf.Entities.Add(New Line(previousPoint, point) With {.Color = color, .Layer = targetLayer})
            Else
                firstPoint = point
                haveFirst = True
            End If
            previousPoint = point
        Next
        If haveFirst Then
            dxf.Entities.Add(New Line(previousPoint, firstPoint) With {.Color = color, .Layer = targetLayer})
        End If
    End Sub

    Private Sub DrawLine3D(dxf As DxfDocument, p1 As Vector3, p2 As Vector3, color As AciColor, layer As Layer)
        dxf.Entities.Add(New Line(p1, p2) With {.Color = color, .Layer = layer})
    End Sub

End Class
