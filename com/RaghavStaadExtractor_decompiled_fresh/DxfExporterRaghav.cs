// ===== RaghavStaadExtractor.DxfExporterRaghav =====

// Decompiled member-by-member from RaghavStaadExtractor.dll



private const double SCALING_FACTOR = 1000.0;


using netDxf;

private readonly AciColor[] memberColors;


using netDxf;

private readonly AciColor[] tubeColors;


using System.Collections.Generic;

private Dictionary<string, string> gridMemberSignatures;


using System.Collections.Generic;
using netDxf;

private Dictionary<string, AciColor> gridColorMap;


using netDxf;

private AciColor[] masterGridColors;


using netDxf;

private AciColor magentaColor;


private int gridColorIndex;


using System.Collections.Generic;
using netDxf;

public DxfExporterRaghav()
{
	memberColors = new AciColor[55]
	{
		AciColor.FromCadIndex(1),
		AciColor.FromCadIndex(2),
		AciColor.FromCadIndex(3),
		AciColor.FromCadIndex(4),
		AciColor.FromCadIndex(5),
		AciColor.FromCadIndex(6),
		AciColor.FromCadIndex(7),
		AciColor.FromCadIndex(10),
		AciColor.FromCadIndex(11),
		AciColor.FromCadIndex(20),
		AciColor.FromCadIndex(21),
		AciColor.FromCadIndex(30),
		AciColor.FromCadIndex(31),
		AciColor.FromCadIndex(40),
		AciColor.FromCadIndex(41),
		AciColor.FromCadIndex(50),
		AciColor.FromCadIndex(51),
		AciColor.FromCadIndex(60),
		AciColor.FromCadIndex(61),
		AciColor.FromCadIndex(70),
		AciColor.FromCadIndex(71),
		AciColor.FromCadIndex(80),
		AciColor.FromCadIndex(81),
		AciColor.FromCadIndex(90),
		AciColor.FromCadIndex(91),
		AciColor.FromCadIndex(100),
		AciColor.FromCadIndex(101),
		AciColor.FromCadIndex(110),
		AciColor.FromCadIndex(111),
		AciColor.FromCadIndex(120),
		AciColor.FromCadIndex(121),
		AciColor.FromCadIndex(130),
		AciColor.FromCadIndex(131),
		AciColor.FromCadIndex(140),
		AciColor.FromCadIndex(141),
		AciColor.FromCadIndex(150),
		AciColor.FromCadIndex(151),
		AciColor.FromCadIndex(160),
		AciColor.FromCadIndex(161),
		AciColor.FromCadIndex(170),
		AciColor.FromCadIndex(171),
		AciColor.FromCadIndex(180),
		AciColor.FromCadIndex(181),
		AciColor.FromCadIndex(190),
		AciColor.FromCadIndex(191),
		AciColor.FromCadIndex(200),
		AciColor.FromCadIndex(201),
		AciColor.FromCadIndex(210),
		AciColor.FromCadIndex(211),
		AciColor.FromCadIndex(220),
		AciColor.FromCadIndex(221),
		AciColor.FromCadIndex(230),
		AciColor.FromCadIndex(231),
		AciColor.FromCadIndex(240),
		AciColor.FromCadIndex(241)
	};
	tubeColors = new AciColor[5]
	{
		AciColor.Yellow,
		AciColor.Cyan,
		AciColor.Magenta,
		AciColor.Green,
		AciColor.Red
	};
	gridMemberSignatures = new Dictionary<string, string>();
	gridColorMap = new Dictionary<string, AciColor>();
	masterGridColors = new AciColor[8]
	{
		AciColor.Red,
		AciColor.Blue,
		AciColor.Green,
		AciColor.Yellow,
		AciColor.Cyan,
		AciColor.FromCadIndex(30),
		AciColor.FromCadIndex(40),
		AciColor.FromCadIndex(50)
	};
	magentaColor = AciColor.Magenta;
	gridColorIndex = 0;
}


using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;
using netDxf;
using netDxf.Entities;
using netDxf.Header;
using netDxf.Tables;

public void ExportBeamsToDxfWithPrompt(object excelSheet)
{
	checked
	{
		try
		{
			if (excelSheet == null)
			{
				MessageBox.Show("Excel sheet object is not valid.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Hand);
				return;
			}
			if (!HasValidBeamData(RuntimeHelpers.GetObjectValue(excelSheet)))
			{
				MessageBox.Show("Pick a member in STAAD, then continue.", "No Beam Data Found", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
				return;
			}
			string text = GetSafeCellValueAsString(RuntimeHelpers.GetObjectValue(excelSheet), 2, 45);
			if (string.IsNullOrEmpty(text))
			{
				text = "1";
			}
			double num = GetSafeCellValue(RuntimeHelpers.GetObjectValue(excelSheet), 4, 45);
			if (num <= 0.0)
			{
				num = 150.0;
			}
			int num2 = (int)Math.Round(GetSafeCellValue(RuntimeHelpers.GetObjectValue(excelSheet), 5, 45));
			string text2 = GetSafeCellValueAsString(RuntimeHelpers.GetObjectValue(excelSheet), 2, 44).Trim();
			if (string.IsNullOrWhiteSpace(text2))
			{
				text2 = "BeamsOutput";
			}
			string text3 = ShowProfessionalInputDialog("Enter DXF File Name (without extension):", "Save DXF File", text2);
			if (string.IsNullOrWhiteSpace(text3))
			{
				return;
			}
			string path = Conversions.ToString(NewLateBinding.LateGet(((Worksheet)excelSheet).Parent, null, "FullName", new object[0], null, null, null));
			string directoryName = Path.GetDirectoryName(path);
			string text4 = Path.Combine(directoryName, "PlateNovaDrawings");
			Directory.CreateDirectory(text4);
			string text5 = Path.Combine(text4, text3 + ".dxf");
			int num3 = 1;
			while (File.Exists(text5))
			{
				text5 = Path.Combine(text4, text3 + "_" + Conversions.ToString(num3) + ".dxf");
				num3++;
			}
			DxfDocument dxfDocument = new DxfDocument(DxfVersion.AutoCad2010);
			Layer layer = new Layer("Beams");
			dxfDocument.Layers.Add(layer);
			Layer layer2 = new Layer("LeaderLines");
			dxfDocument.Layers.Add(layer2);
			Layer layer3 = new Layer("BeamText");
			dxfDocument.Layers.Add(layer3);
			bool flag = AddGridToDxf(dxfDocument, RuntimeHelpers.GetObjectValue(excelSheet), text);
			double? leaderDirectionAngle = GetLeaderDirectionAngle(RuntimeHelpers.GetObjectValue(excelSheet));
			int num4 = Conversions.ToInteger(NewLateBinding.LateGet(NewLateBinding.LateGet(((Worksheet)excelSheet).Cells[((Worksheet)excelSheet).Rows.Count, 1], null, "End", new object[1] { XlDirection.xlUp }, null, null, null), null, "Row", new object[0], null, null, null));
			bool flag2 = false;
			bool flag3 = false;
			bool flag4 = false;
			double safeCellValue = GetSafeCellValue(RuntimeHelpers.GetObjectValue(excelSheet), 5, 45);
			if (safeCellValue == 100.0)
			{
				flag2 = true;
			}
			else if (safeCellValue == 101.0)
			{
				flag2 = true;
				flag3 = true;
			}
			else if (safeCellValue == 200.0)
			{
				flag2 = true;
				flag4 = true;
			}
			Dictionary<string, Layer> dictionary = null;
			if (flag2)
			{
				dictionary = CreateColorBasedLayers(dxfDocument, RuntimeHelpers.GetObjectValue(excelSheet), num4);
			}
			int num5 = num4;
			for (int i = 9; i <= num5; i++)
			{
				try
				{
					double num6 = GetSafeCellValue(RuntimeHelpers.GetObjectValue(excelSheet), i, 16) * 1000.0;
					double num7 = GetSafeCellValue(RuntimeHelpers.GetObjectValue(excelSheet), i, 17) * 1000.0;
					double num8 = GetSafeCellValue(RuntimeHelpers.GetObjectValue(excelSheet), i, 18) * 1000.0;
					double num9 = GetSafeCellValue(RuntimeHelpers.GetObjectValue(excelSheet), i, 19) * 1000.0;
					double num10 = GetSafeCellValue(RuntimeHelpers.GetObjectValue(excelSheet), i, 20) * 1000.0;
					double num11 = GetSafeCellValue(RuntimeHelpers.GetObjectValue(excelSheet), i, 21) * 1000.0;
					if (num6 == 0.0 && num7 == 0.0 && num8 == 0.0 && num9 == 0.0 && num10 == 0.0 && num11 == 0.0)
					{
						continue;
					}
					Vector3 startPoint = ApplyRotation(new Vector3(num6, num7, num8), text);
					Vector3 endPoint = ApplyRotation(new Vector3(num9, num10, num11), text);
					double safeCellValue2 = GetSafeCellValue(RuntimeHelpers.GetObjectValue(excelSheet), i, 22);
					Layer layer4 = layer;
					AciColor aciColor;
					unchecked
					{
						AciColor color;
						if (flag2 && dictionary != null)
						{
							layer4 = GetBeamLayerByParameters(RuntimeHelpers.GetObjectValue(excelSheet), i, dictionary, layer);
							aciColor = layer4.Color;
							color = aciColor;
						}
						else
						{
							int num12 = checked(i - 9) % memberColors.Length;
							aciColor = memberColors[num12];
							int num13 = checked(i - 9) % tubeColors.Length;
							color = tubeColors[num13];
						}
						Line line = new Line(startPoint, endPoint);
						line.Layer = layer4;
						line.Color = aciColor;
						dxfDocument.Entities.Add(line);
						if (!flag4)
						{
							double safeCellValue3 = GetSafeCellValue(RuntimeHelpers.GetObjectValue(excelSheet), i, 6);
							if (safeCellValue3 == 655.0 || safeCellValue3 == 660.0)
							{
								double outerDiameter = GetSafeCellValue(RuntimeHelpers.GetObjectValue(excelSheet), i, 8) * 1000.0;
								double thickness = GetSafeCellValue(RuntimeHelpers.GetObjectValue(excelSheet), i, 9) * 1000.0;
								DrawPipeSection(dxfDocument, startPoint, endPoint, outerDiameter, thickness, color, safeCellValue2, text, layer4);
							}
							else if (safeCellValue3 == 651.0)
							{
								double width = GetSafeCellValue(RuntimeHelpers.GetObjectValue(excelSheet), i, 12) * 1000.0;
								double thickness2 = GetSafeCellValue(RuntimeHelpers.GetObjectValue(excelSheet), i, 11) * 1000.0;
								DrawTubeSection(dxfDocument, startPoint, endPoint, width, thickness2, color, safeCellValue2, text, layer4);
							}
							else if (safeCellValue3 == 654.0 || safeCellValue3 == 650.0)
							{
								double width2 = GetSafeCellValue(RuntimeHelpers.GetObjectValue(excelSheet), i, 8) * 1000.0;
								double height = GetSafeCellValue(RuntimeHelpers.GetObjectValue(excelSheet), i, 9) * 1000.0;
								double thickness3 = GetSafeCellValue(RuntimeHelpers.GetObjectValue(excelSheet), i, 10) * 1000.0;
								DrawHollowRectangleTube(dxfDocument, startPoint, endPoint, width2, height, thickness3, color, safeCellValue2, text, layer4);
							}
							else if (safeCellValue3 == 672.0)
							{
								double width3 = GetSafeCellValue(RuntimeHelpers.GetObjectValue(excelSheet), i, 11) * 1000.0;
								double height2 = GetSafeCellValue(RuntimeHelpers.GetObjectValue(excelSheet), i, 12) * 1000.0;
								double thickness4 = 0.0;
								DrawHollowRectangleTube(dxfDocument, startPoint, endPoint, width3, height2, thickness4, aciColor, safeCellValue2, text, layer4);
							}
							else if (safeCellValue3 == 697.0)
							{
								double num14 = GetSafeCellValue(RuntimeHelpers.GetObjectValue(excelSheet), i, 8) * 1000.0;
								double num15 = GetSafeCellValue(RuntimeHelpers.GetObjectValue(excelSheet), i, 9) * 1000.0;
								double num16 = GetSafeCellValue(RuntimeHelpers.GetObjectValue(excelSheet), i, 10) * 1000.0;
								double num17 = GetSafeCellValue(RuntimeHelpers.GetObjectValue(excelSheet), i, 11) * 1000.0;
								if (Math.Abs(num14 - num16) > 0.001)
								{
									double num18 = num14 - 2.0 * num17;
									DrawTaperedISection(dxfDocument, startPoint, endPoint, num18, num15, num16, num17, num16, num17, num18, num15, num16, num17, num16, num17, aciColor, safeCellValue2, text, layer4);
									double betaAngle = safeCellValue2 + 90.0;
									DrawTaperedISection(dxfDocument, startPoint, endPoint, num18, num15, num16, num17, num16, num17, num18, num15, num16, num17, num16, num17, aciColor, betaAngle, text, layer4);
								}
							}
							else if (safeCellValue3 == 680.0)
							{
								DrawTaperedISection(dxfDocument, startPoint, endPoint, GetSafeCellValue(RuntimeHelpers.GetObjectValue(excelSheet), i, 24), GetSafeCellValue(RuntimeHelpers.GetObjectValue(excelSheet), i, 25), GetSafeCellValue(RuntimeHelpers.GetObjectValue(excelSheet), i, 26), GetSafeCellValue(RuntimeHelpers.GetObjectValue(excelSheet), i, 27), GetSafeCellValue(RuntimeHelpers.GetObjectValue(excelSheet), i, 28), GetSafeCellValue(RuntimeHelpers.GetObjectValue(excelSheet), i, 29), GetSafeCellValue(RuntimeHelpers.GetObjectValue(excelSheet), i, 30), GetSafeCellValue(RuntimeHelpers.GetObjectValue(excelSheet), i, 31), GetSafeCellValue(RuntimeHelpers.GetObjectValue(excelSheet), i, 32), GetSafeCellValue(RuntimeHelpers.GetObjectValue(excelSheet), i, 33), GetSafeCellValue(RuntimeHelpers.GetObjectValue(excelSheet), i, 34), GetSafeCellValue(RuntimeHelpers.GetObjectValue(excelSheet), i, 35), aciColor, safeCellValue2, text, layer4);
							}
							else if (safeCellValue3 == 630.0)
							{
								double num19 = GetSafeCellValue(RuntimeHelpers.GetObjectValue(excelSheet), i, 8) * 1000.0;
								double num20 = GetSafeCellValue(RuntimeHelpers.GetObjectValue(excelSheet), i, 11) * 1000.0;
								double num21 = GetSafeCellValue(RuntimeHelpers.GetObjectValue(excelSheet), i, 9) * 1000.0;
								double num22 = GetSafeCellValue(RuntimeHelpers.GetObjectValue(excelSheet), i, 10) * 1000.0;
								double num23 = num19 - 2.0 * num22;
								DrawTaperedChannelSection(dxfDocument, startPoint, endPoint, num23, num20, num21, num22, num21, num22, num23, num20, num21, num22, num21, num22, aciColor, safeCellValue2, text, layer4);
							}
							else if (safeCellValue3 == 610.0)
							{
								double num24 = GetSafeCellValue(RuntimeHelpers.GetObjectValue(excelSheet), i, 8) * 1000.0;
								double num25 = GetSafeCellValue(RuntimeHelpers.GetObjectValue(excelSheet), i, 11) * 1000.0;
								double num26 = GetSafeCellValue(RuntimeHelpers.GetObjectValue(excelSheet), i, 9) * 1000.0;
								double num27 = GetSafeCellValue(RuntimeHelpers.GetObjectValue(excelSheet), i, 10) * 1000.0;
								double num28 = GetSafeCellValue(RuntimeHelpers.GetObjectValue(excelSheet), i, 9) * 1000.0;
								double num29 = GetSafeCellValue(RuntimeHelpers.GetObjectValue(excelSheet), i, 10) * 1000.0;
								double num30 = num24 - 2.0 * num27;
								DrawTaperedISection(dxfDocument, startPoint, endPoint, num30, num25, num26, num27, num28, num29, num30, num25, num26, num27, num28, num29, aciColor, safeCellValue2, text, layer4);
							}
							else if (safeCellValue3 == 613.0)
							{
								double num31 = GetSafeCellValue(RuntimeHelpers.GetObjectValue(excelSheet), i, 8) * 1000.0;
								double num32 = GetSafeCellValue(RuntimeHelpers.GetObjectValue(excelSheet), i, 11) * 1000.0;
								double num33 = GetSafeCellValue(RuntimeHelpers.GetObjectValue(excelSheet), i, 9) * 1000.0;
								double num34 = GetSafeCellValue(RuntimeHelpers.GetObjectValue(excelSheet), i, 10) * 1000.0;
								double num35 = GetSafeCellValue(RuntimeHelpers.GetObjectValue(excelSheet), i, 9) * 1000.0;
								double num36 = GetSafeCellValue(RuntimeHelpers.GetObjectValue(excelSheet), i, 10) * 1000.0;
								double num37 = GetSafeCellValue(RuntimeHelpers.GetObjectValue(excelSheet), i, 52) * 1000.0;
								double num38 = GetSafeCellValue(RuntimeHelpers.GetObjectValue(excelSheet), i, 53) * 1000.0;
								double num39 = num31 - 2.0 * num34;
								DrawTaperedISection(dxfDocument, startPoint, endPoint, num39, num32, num33, num34, num35, num36, num39, num32, num33, num34, num35, num36, aciColor, safeCellValue2, text, layer4);
								if (num37 > 0.0 && num38 > 0.0)
								{
									DrawExtraTopFlange(dxfDocument, startPoint, endPoint, num37, num38, num39, num34, aciColor, safeCellValue2, text, layer4);
								}
							}
							else if (safeCellValue3 == 614.0)
							{
								double num40 = GetSafeCellValue(RuntimeHelpers.GetObjectValue(excelSheet), i, 8) * 1000.0;
								double num41 = GetSafeCellValue(RuntimeHelpers.GetObjectValue(excelSheet), i, 11) * 1000.0;
								double num42 = GetSafeCellValue(RuntimeHelpers.GetObjectValue(excelSheet), i, 9) * 1000.0;
								double num43 = GetSafeCellValue(RuntimeHelpers.GetObjectValue(excelSheet), i, 10) * 1000.0;
								double num44 = GetSafeCellValue(RuntimeHelpers.GetObjectValue(excelSheet), i, 9) * 1000.0;
								double num45 = GetSafeCellValue(RuntimeHelpers.GetObjectValue(excelSheet), i, 10) * 1000.0;
								double num46 = GetSafeCellValue(RuntimeHelpers.GetObjectValue(excelSheet), i, 52) * 1000.0;
								double num47 = GetSafeCellValue(RuntimeHelpers.GetObjectValue(excelSheet), i, 53) * 1000.0;
								double num48 = num40 - 2.0 * num43;
								DrawTaperedISection(dxfDocument, startPoint, endPoint, num48, num41, num42, num43, num44, num45, num48, num41, num42, num43, num44, num45, aciColor, safeCellValue2, text, layer4);
								if (num46 > 0.0 && num47 > 0.0)
								{
									DrawExtraBottomFlange(dxfDocument, startPoint, endPoint, num46, num47, num48, num45, aciColor, safeCellValue2, text, layer4);
								}
							}
							else if (safeCellValue3 == 615.0)
							{
								double num49 = GetSafeCellValue(RuntimeHelpers.GetObjectValue(excelSheet), i, 8) * 1000.0;
								double num50 = GetSafeCellValue(RuntimeHelpers.GetObjectValue(excelSheet), i, 11) * 1000.0;
								double num51 = GetSafeCellValue(RuntimeHelpers.GetObjectValue(excelSheet), i, 9) * 1000.0;
								double num52 = GetSafeCellValue(RuntimeHelpers.GetObjectValue(excelSheet), i, 10) * 1000.0;
								double num53 = GetSafeCellValue(RuntimeHelpers.GetObjectValue(excelSheet), i, 9) * 1000.0;
								double num54 = GetSafeCellValue(RuntimeHelpers.GetObjectValue(excelSheet), i, 10) * 1000.0;
								double num55 = GetSafeCellValue(RuntimeHelpers.GetObjectValue(excelSheet), i, 52) * 1000.0;
								double num56 = GetSafeCellValue(RuntimeHelpers.GetObjectValue(excelSheet), i, 53) * 1000.0;
								double num57 = GetSafeCellValue(RuntimeHelpers.GetObjectValue(excelSheet), i, 54) * 1000.0;
								double num58 = GetSafeCellValue(RuntimeHelpers.GetObjectValue(excelSheet), i, 55) * 1000.0;
								double num59;
								double num60;
								double num61;
								double num62;
								if (num57 <= 0.0 || num58 <= 0.0)
								{
									num59 = num55;
									num60 = num56;
									num61 = num55;
									num62 = num56;
								}
								else if (num55 <= 0.0 || num56 <= 0.0)
								{
									num59 = num57;
									num60 = num58;
									num61 = num57;
									num62 = num58;
								}
								else
								{
									num59 = num55;
									num60 = num56;
									num61 = num57;
									num62 = num58;
								}
								double num63 = num49 - 2.0 * num52;
								DrawTaperedISection(dxfDocument, startPoint, endPoint, num63, num50, num51, num52, num53, num54, num63, num50, num51, num52, num53, num54, aciColor, safeCellValue2, text, layer4);
								if (num59 > 0.0 && num60 > 0.0)
								{
									DrawExtraTopFlange(dxfDocument, startPoint, endPoint, num59, num60, num63, num52, aciColor, safeCellValue2, text, layer4);
								}
								if (num61 > 0.0 && num62 > 0.0)
								{
									DrawExtraBottomFlange(dxfDocument, startPoint, endPoint, num61, num62, num63, num54, aciColor, safeCellValue2, text, layer4);
								}
							}
						}
					}
					if (leaderDirectionAngle.HasValue)
					{
						double value = leaderDirectionAngle.Value;
						Vector3 startPoint2 = new Vector3((startPoint.X + endPoint.X) / 2.0, (startPoint.Y + endPoint.Y) / 2.0, (startPoint.Z + endPoint.Z) / 2.0);
						int value2 = (int)Math.Round(GetSafeCellValue(RuntimeHelpers.GetObjectValue(excelSheet), 2, 45));
						double val;
						double val2;
						if (new int[3] { 3, 4, 5 }.Contains(value2))
						{
							val = GetSafeCellValue(RuntimeHelpers.GetObjectValue(excelSheet), i, 10) * 1000.0;
							val2 = GetSafeCellValue(RuntimeHelpers.GetObjectValue(excelSheet), i, 12) * 1000.0;
						}
						else
						{
							val = GetSafeCellValue(RuntimeHelpers.GetObjectValue(excelSheet), i, 7) * 1000.0;
							val2 = GetSafeCellValue(RuntimeHelpers.GetObjectValue(excelSheet), i, 9) * 1000.0;
						}
						double num64 = Math.Max(val, val2);
						double num65 = (flag4 ? (2.0 * num) : (5.5 * num));
						double num66 = num65 + (flag4 ? 0.0 : (num64 / 1.2));
						Vector3 vector = new Vector3(startPoint2.X + num66 * Math.Cos(value), startPoint2.Y + num66 * Math.Sin(value), startPoint2.Z);
						double num67 = 100.0;
						Vector3 endPoint2 = new Vector3(vector.X + num67, vector.Y, vector.Z);
						AciColor color2 = ((!flag2) ? ((num2 == 0) ? aciColor : AciColor.FromCadIndex((short)num2)) : aciColor);
						Line line2 = new Line(startPoint2, vector);
						line2.Layer = (flag2 ? layer4 : layer2);
						line2.Color = color2;
						dxfDocument.Entities.Add(line2);
						Line line3 = new Line(vector, endPoint2);
						line3.Layer = (flag2 ? layer4 : layer2);
						line3.Color = color2;
						dxfDocument.Entities.Add(line3);
						string safeCellValueAsString = GetSafeCellValueAsString(RuntimeHelpers.GetObjectValue(excelSheet), i, 37);
						string safeCellValueAsString2 = GetSafeCellValueAsString(RuntimeHelpers.GetObjectValue(excelSheet), i, 38);
						string safeCellValueAsString3 = GetSafeCellValueAsString(RuntimeHelpers.GetObjectValue(excelSheet), i, 39);
						double beamLength = GetSafeCellValue(RuntimeHelpers.GetObjectValue(excelSheet), i, 3) * 1000.0;
						double safeCellValue4 = GetSafeCellValue(RuntimeHelpers.GetObjectValue(excelSheet), i, 6);
						string text6;
						if (flag3)
						{
							string columnEValue = GetSafeCellValueAsString(RuntimeHelpers.GetObjectValue(excelSheet), i, 5).Trim();
							text6 = BuildSizeTextWithoutLength(safeCellValueAsString, safeCellValueAsString2, safeCellValueAsString3, safeCellValue4, columnEValue);
						}
						else
						{
							text6 = BuildSizeText(safeCellValueAsString, safeCellValueAsString2, safeCellValueAsString3, beamLength, safeCellValue4);
						}
						if (!string.IsNullOrWhiteSpace(text6))
						{
							double num68 = 0.0 - (3.0 * num + 10.0);
							Vector3 position = new Vector3(endPoint2.X, endPoint2.Y + num68, endPoint2.Z);
							MText mText = new MText(text6.Replace("\r\n", "\\P"), position, num, 0.0);
							mText.Layer = (flag2 ? layer4 : layer3);
							mText.Color = color2;
							mText.AttachmentPoint = MTextAttachmentPoint.BottomLeft;
							dxfDocument.Entities.Add(mText);
						}
					}
				}
				catch (Exception projectError)
				{
					ProjectData.SetProjectError(projectError);
					ProjectData.ClearProjectError();
				}
			}
			dxfDocument.Save(text5);
			string successMessage = string.Format("DXF file saved successfully:{0}{1}", "\r\n", text5);
			if (flag)
			{
				successMessage = string.Format("DXF file Created successfully:{0}{1}", "\r\n", text5);
			}
			if (flag4)
			{
				successMessage = string.Format("Single Line DXF file Created successfully:{0}{1}", "\r\n", text5);
			}
			ShowSuccessMessageWithOpen(successMessage, text5, text4);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			MessageBox.Show($"Error exporting to DXF: {ex2.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Hand);
			ProjectData.ClearProjectError();
		}
	}
}


using System;
using Microsoft.VisualBasic.CompilerServices;
using netDxf;
using netDxf.Entities;
using netDxf.Tables;

private void DrawExtraTopFlange(DxfDocument dxf, Vector3 startPoint, Vector3 endPoint, double extraFlangeWidth, double extraFlangeThickness, double webHeight, double baseTopFlangeThickness, AciColor color, double betaAngle, string rotationValue, Layer targetLayer)
{
	Vector3 vector = endPoint - startPoint;
	Vector3 vector2 = vector;
	vector2.Normalize();
	double num = 0.99;
	Vector3 vector3;
	Vector3 vector4;
	if (Math.Abs(vector2.Z) > num)
	{
		vector3 = new Vector3(1.0, 0.0, 0.0);
		vector4 = Vector3.CrossProduct(vector2, vector3);
		vector4.Normalize();
		vector3 = Vector3.CrossProduct(vector4, vector2);
		vector3.Normalize();
	}
	else
	{
		vector4 = Vector3.CrossProduct(v: new Vector3(0.0, 0.0, 1.0), u: vector2);
		double num2 = Math.Sqrt(vector4.X * vector4.X + vector4.Y * vector4.Y + vector4.Z * vector4.Z);
		if (num2 < 0.001)
		{
			vector4 = new Vector3(0.0, 1.0, 0.0);
		}
		else
		{
			vector4.Normalize();
		}
		vector3 = Vector3.CrossProduct(vector4, vector2);
		vector3.Normalize();
		vector4 = Vector3.CrossProduct(vector2, vector3);
		vector4.Normalize();
	}
	if (Operators.CompareString(rotationValue.Trim(), "5", TextCompare: false) == 0)
	{
		vector3 = RotateVectorAroundAxis(vector3, vector2, -90.0);
		vector4 = RotateVectorAroundAxis(vector4, vector2, -90.0);
	}
	else if (Operators.CompareString(rotationValue.Trim(), "3", TextCompare: false) == 0)
	{
		double num3 = Math.Abs(vector2.X);
		double num4 = Math.Abs(vector2.Y);
		double num5 = Math.Abs(vector2.Z);
		if (num3 > 0.7 || num5 > 0.7)
		{
			vector3 = RotateVectorAroundAxis(vector3, vector2, -90.0);
			vector4 = RotateVectorAroundAxis(vector4, vector2, -90.0);
		}
		vector3 = RotateVectorAroundAxis(vector3, vector2, -90.0);
		vector4 = RotateVectorAroundAxis(vector4, vector2, -90.0);
	}
	else if (Operators.CompareString(rotationValue.Trim(), "4", TextCompare: false) == 0)
	{
		double num6 = Math.Abs(vector2.X);
		double num7 = Math.Abs(vector2.Y);
		double num8 = Math.Abs(vector2.Z);
		if (num6 > 0.7 || num8 > 0.7)
		{
			vector3 = RotateVectorAroundAxis(vector3, vector2, -90.0);
			vector4 = RotateVectorAroundAxis(vector4, vector2, -90.0);
		}
		vector3 = RotateVectorAroundAxis(vector3, vector2, -90.0);
		vector4 = RotateVectorAroundAxis(vector4, vector2, -90.0);
	}
	if (Math.Abs(betaAngle) > 0.001)
	{
		vector3 = RotateVectorAroundAxis(vector3, vector2, betaAngle);
		vector4 = RotateVectorAroundAxis(vector4, vector2, betaAngle);
	}
	VB$AnonymousDelegate_0<double, double, Vector3, Vector3> vB$AnonymousDelegate_ = (double offsetY, double offsetZ, Vector3 basePoint) => basePoint + vector3 * offsetY + vector4 * offsetZ;
	double num9 = webHeight / 2.0;
	double num10 = num9 + baseTopFlangeThickness;
	double num11 = num10;
	double offsetZ2 = num11 + extraFlangeThickness;
	Vector3 vector5 = vB$AnonymousDelegate_((0.0 - extraFlangeWidth) / 2.0, num11, startPoint);
	Vector3 vector6 = vB$AnonymousDelegate_(extraFlangeWidth / 2.0, num11, startPoint);
	Vector3 vector7 = vB$AnonymousDelegate_((0.0 - extraFlangeWidth) / 2.0, offsetZ2, startPoint);
	Vector3 vector8 = vB$AnonymousDelegate_(extraFlangeWidth / 2.0, offsetZ2, startPoint);
	Vector3 vector9 = vB$AnonymousDelegate_((0.0 - extraFlangeWidth) / 2.0, num11, endPoint);
	Vector3 vector10 = vB$AnonymousDelegate_(extraFlangeWidth / 2.0, num11, endPoint);
	Vector3 vector11 = vB$AnonymousDelegate_((0.0 - extraFlangeWidth) / 2.0, offsetZ2, endPoint);
	Vector3 vector12 = vB$AnonymousDelegate_(extraFlangeWidth / 2.0, offsetZ2, endPoint);
	dxf.Entities.Add(new Line(vector5, vector9)
	{
		Color = color,
		Layer = targetLayer
	});
	dxf.Entities.Add(new Line(vector6, vector10)
	{
		Color = color,
		Layer = targetLayer
	});
	dxf.Entities.Add(new Line(vector7, vector11)
	{
		Color = color,
		Layer = targetLayer
	});
	dxf.Entities.Add(new Line(vector8, vector12)
	{
		Color = color,
		Layer = targetLayer
	});
	dxf.Entities.Add(new Line(vector5, vector6)
	{
		Color = color,
		Layer = targetLayer
	});
	dxf.Entities.Add(new Line(vector6, vector8)
	{
		Color = color,
		Layer = targetLayer
	});
	dxf.Entities.Add(new Line(vector8, vector7)
	{
		Color = color,
		Layer = targetLayer
	});
	dxf.Entities.Add(new Line(vector7, vector5)
	{
		Color = color,
		Layer = targetLayer
	});
	dxf.Entities.Add(new Line(vector9, vector10)
	{
		Color = color,
		Layer = targetLayer
	});
	dxf.Entities.Add(new Line(vector10, vector12)
	{
		Color = color,
		Layer = targetLayer
	});
	dxf.Entities.Add(new Line(vector12, vector11)
	{
		Color = color,
		Layer = targetLayer
	});
	dxf.Entities.Add(new Line(vector11, vector9)
	{
		Color = color,
		Layer = targetLayer
	});
}


using System;
using Microsoft.VisualBasic.CompilerServices;
using netDxf;
using netDxf.Entities;
using netDxf.Tables;

private void DrawExtraBottomFlange(DxfDocument dxf, Vector3 startPoint, Vector3 endPoint, double extraFlangeWidth, double extraFlangeThickness, double webHeight, double baseBottomFlangeThickness, AciColor color, double betaAngle, string rotationValue, Layer targetLayer)
{
	Vector3 vector = endPoint - startPoint;
	Vector3 vector2 = vector;
	vector2.Normalize();
	double num = 0.99;
	Vector3 vector3;
	Vector3 vector4;
	if (Math.Abs(vector2.Z) > num)
	{
		vector3 = new Vector3(1.0, 0.0, 0.0);
		vector4 = Vector3.CrossProduct(vector2, vector3);
		vector4.Normalize();
		vector3 = Vector3.CrossProduct(vector4, vector2);
		vector3.Normalize();
	}
	else
	{
		vector4 = Vector3.CrossProduct(v: new Vector3(0.0, 0.0, 1.0), u: vector2);
		double num2 = Math.Sqrt(vector4.X * vector4.X + vector4.Y * vector4.Y + vector4.Z * vector4.Z);
		if (num2 < 0.001)
		{
			vector4 = new Vector3(0.0, 1.0, 0.0);
		}
		else
		{
			vector4.Normalize();
		}
		vector3 = Vector3.CrossProduct(vector4, vector2);
		vector3.Normalize();
		vector4 = Vector3.CrossProduct(vector2, vector3);
		vector4.Normalize();
	}
	if (Operators.CompareString(rotationValue.Trim(), "5", TextCompare: false) == 0)
	{
		vector3 = RotateVectorAroundAxis(vector3, vector2, -90.0);
		vector4 = RotateVectorAroundAxis(vector4, vector2, -90.0);
	}
	else if (Operators.CompareString(rotationValue.Trim(), "3", TextCompare: false) == 0)
	{
		double num3 = Math.Abs(vector2.X);
		double num4 = Math.Abs(vector2.Y);
		double num5 = Math.Abs(vector2.Z);
		if (num3 > 0.7 || num5 > 0.7)
		{
			vector3 = RotateVectorAroundAxis(vector3, vector2, -90.0);
			vector4 = RotateVectorAroundAxis(vector4, vector2, -90.0);
		}
		vector3 = RotateVectorAroundAxis(vector3, vector2, -90.0);
		vector4 = RotateVectorAroundAxis(vector4, vector2, -90.0);
	}
	else if (Operators.CompareString(rotationValue.Trim(), "4", TextCompare: false) == 0)
	{
		double num6 = Math.Abs(vector2.X);
		double num7 = Math.Abs(vector2.Y);
		double num8 = Math.Abs(vector2.Z);
		if (num6 > 0.7 || num8 > 0.7)
		{
			vector3 = RotateVectorAroundAxis(vector3, vector2, -90.0);
			vector4 = RotateVectorAroundAxis(vector4, vector2, -90.0);
		}
		vector3 = RotateVectorAroundAxis(vector3, vector2, -90.0);
		vector4 = RotateVectorAroundAxis(vector4, vector2, -90.0);
	}
	if (Math.Abs(betaAngle) > 0.001)
	{
		vector3 = RotateVectorAroundAxis(vector3, vector2, betaAngle);
		vector4 = RotateVectorAroundAxis(vector4, vector2, betaAngle);
	}
	VB$AnonymousDelegate_0<double, double, Vector3, Vector3> vB$AnonymousDelegate_ = (double offsetY, double offsetZ, Vector3 basePoint) => basePoint + vector3 * offsetY + vector4 * offsetZ;
	double num9 = (0.0 - webHeight) / 2.0;
	double num10 = num9 - baseBottomFlangeThickness;
	double num11 = num10;
	double offsetZ2 = num11 - extraFlangeThickness;
	Vector3 vector5 = vB$AnonymousDelegate_((0.0 - extraFlangeWidth) / 2.0, num11, startPoint);
	Vector3 vector6 = vB$AnonymousDelegate_(extraFlangeWidth / 2.0, num11, startPoint);
	Vector3 vector7 = vB$AnonymousDelegate_((0.0 - extraFlangeWidth) / 2.0, offsetZ2, startPoint);
	Vector3 vector8 = vB$AnonymousDelegate_(extraFlangeWidth / 2.0, offsetZ2, startPoint);
	Vector3 vector9 = vB$AnonymousDelegate_((0.0 - extraFlangeWidth) / 2.0, num11, endPoint);
	Vector3 vector10 = vB$AnonymousDelegate_(extraFlangeWidth / 2.0, num11, endPoint);
	Vector3 vector11 = vB$AnonymousDelegate_((0.0 - extraFlangeWidth) / 2.0, offsetZ2, endPoint);
	Vector3 vector12 = vB$AnonymousDelegate_(extraFlangeWidth / 2.0, offsetZ2, endPoint);
	dxf.Entities.Add(new Line(vector5, vector9)
	{
		Color = color,
		Layer = targetLayer
	});
	dxf.Entities.Add(new Line(vector6, vector10)
	{
		Color = color,
		Layer = targetLayer
	});
	dxf.Entities.Add(new Line(vector7, vector11)
	{
		Color = color,
		Layer = targetLayer
	});
	dxf.Entities.Add(new Line(vector8, vector12)
	{
		Color = color,
		Layer = targetLayer
	});
	dxf.Entities.Add(new Line(vector5, vector6)
	{
		Color = color,
		Layer = targetLayer
	});
	dxf.Entities.Add(new Line(vector6, vector8)
	{
		Color = color,
		Layer = targetLayer
	});
	dxf.Entities.Add(new Line(vector8, vector7)
	{
		Color = color,
		Layer = targetLayer
	});
	dxf.Entities.Add(new Line(vector7, vector5)
	{
		Color = color,
		Layer = targetLayer
	});
	dxf.Entities.Add(new Line(vector9, vector10)
	{
		Color = color,
		Layer = targetLayer
	});
	dxf.Entities.Add(new Line(vector10, vector12)
	{
		Color = color,
		Layer = targetLayer
	});
	dxf.Entities.Add(new Line(vector12, vector11)
	{
		Color = color,
		Layer = targetLayer
	});
	dxf.Entities.Add(new Line(vector11, vector9)
	{
		Color = color,
		Layer = targetLayer
	});
}


using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

private string ShowProfessionalInputDialog(string prompt, string title, string defaultValue)
{
	Form form = new Form();
	Panel panel = new Panel();
	Label label = new Label();
	Label label2 = new Label();
	TextBox textBox = new TextBox();
	Button button = new Button();
	Button button2 = new Button();
	Panel panel2 = new Panel();
	checked
	{
		try
		{
			Form form2 = form;
			form2.Text = string.Empty;
			form2.Size = new Size(480, 240);
			form2.StartPosition = FormStartPosition.CenterParent;
			form2.FormBorderStyle = FormBorderStyle.None;
			form2.MaximizeBox = false;
			form2.MinimizeBox = false;
			form2.BackColor = Color.FromArgb(240, 240, 242);
			form2.ShowInTaskbar = false;
			form2.TopMost = true;
			form2 = null;
			_Closure$__12-0 closure$__12-;
			form.Load += delegate
			{
				closure$__12-._Lambda$__0();
			};
			form.Paint += delegate(object sender, PaintEventArgs e)
			{
				e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
				Rectangle rectangle4 = new Rectangle(2, 2, form.Width - 2, form.Height - 2);
				using GraphicsPath graphicsPath4 = new GraphicsPath();
				int num4 = 12;
				graphicsPath4.AddArc(rectangle4.X, rectangle4.Y, num4, num4, 180f, 90f);
				graphicsPath4.AddArc(rectangle4.X + rectangle4.Width - num4, rectangle4.Y, num4, num4, 270f, 90f);
				graphicsPath4.AddArc(rectangle4.X + rectangle4.Width - num4, rectangle4.Y + rectangle4.Height - num4, num4, num4, 0f, 90f);
				graphicsPath4.AddArc(rectangle4.X, rectangle4.Y + rectangle4.Height - num4, num4, num4, 90f, 90f);
				graphicsPath4.CloseFigure();
				using SolidBrush brush2 = new SolidBrush(Color.FromArgb(30, 0, 0, 0));
				e.Graphics.FillPath(brush2, graphicsPath4);
			};
			Panel panel3 = panel;
			panel3.Location = new Point(0, 0);
			panel3.Size = new Size(480, 60);
			panel3.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
			panel3.BackColor = Color.FromArgb(0, 120, 215);
			panel3.Padding = new Padding(20, 15, 20, 5);
			panel3 = null;
			panel.Paint += delegate(object sender, PaintEventArgs e)
			{
				Rectangle rect = new Rectangle(0, 0, panel.Width, panel.Height);
				using LinearGradientBrush brush = new LinearGradientBrush(rect, Color.FromArgb(16, 137, 230), Color.FromArgb(0, 120, 215), LinearGradientMode.Vertical);
				e.Graphics.FillRectangle(brush, rect);
			};
			Label label3 = label;
			label3.Text = title;
			label3.Font = new Font("Segoe UI", 12f, FontStyle.Bold);
			label3.ForeColor = Color.White;
			label3.AutoSize = true;
			label3.Location = new Point(20, 18);
			label3.BackColor = Color.Transparent;
			label3 = null;
			Panel panel4 = panel2;
			panel4.Location = new Point(0, 60);
			panel4.Size = new Size(480, 180);
			panel4.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
			panel4.Padding = new Padding(25, 20, 25, 20);
			panel4.BackColor = Color.FromArgb(240, 240, 242);
			panel4 = null;
			Label label4 = label2;
			label4.Text = prompt;
			label4.Font = new Font("Segoe UI", 10f, FontStyle.Regular);
			label4.ForeColor = Color.FromArgb(50, 50, 50);
			label4.AutoSize = false;
			label4.Size = new Size(410, 30);
			label4.Location = new Point(25, 25);
			label4.TextAlign = ContentAlignment.MiddleLeft;
			label4.BackColor = Color.Transparent;
			label4 = null;
			TextBox textBox2 = textBox;
			textBox2.Font = new Font("Segoe UI", 11f, FontStyle.Regular);
			textBox2.Size = new Size(410, 32);
			textBox2.Location = new Point(25, 65);
			textBox2.Text = defaultValue;
			textBox2.BorderStyle = BorderStyle.FixedSingle;
			textBox2.BackColor = Color.White;
			textBox2.ForeColor = Color.FromArgb(40, 40, 40);
			textBox2.SelectAll();
			textBox2.Focus();
			textBox2 = null;
			textBox.Paint += delegate(object sender, PaintEventArgs e)
			{
				Rectangle rectangle3 = new Rectangle(0, 0, textBox.Width - 1, textBox.Height - 1);
				using Pen pen = new Pen(Color.FromArgb(180, 180, 180), 1f);
				e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
				using GraphicsPath graphicsPath3 = new GraphicsPath();
				int num3 = 4;
				graphicsPath3.AddArc(rectangle3.X, rectangle3.Y, num3, num3, 180f, 90f);
				graphicsPath3.AddArc(rectangle3.X + rectangle3.Width - num3, rectangle3.Y, num3, num3, 270f, 90f);
				graphicsPath3.AddArc(rectangle3.X + rectangle3.Width - num3, rectangle3.Y + rectangle3.Height - num3, num3, num3, 0f, 90f);
				graphicsPath3.AddArc(rectangle3.X, rectangle3.Y + rectangle3.Height - num3, num3, num3, 90f, 90f);
				graphicsPath3.CloseFigure();
				e.Graphics.DrawPath(pen, graphicsPath3);
			};
			Button button3 = button;
			button3.Text = "Create DXF";
			button3.Font = new Font("Segoe UI", 9f, FontStyle.Bold);
			button3.Size = new Size(110, 38);
			button3.Location = new Point(215, 115);
			button3.BackColor = Color.FromArgb(0, 120, 215);
			button3.ForeColor = Color.White;
			button3.FlatStyle = FlatStyle.Flat;
			button3.FlatAppearance.BorderSize = 0;
			button3.FlatAppearance.MouseOverBackColor = Color.FromArgb(16, 137, 230);
			button3.FlatAppearance.MouseDownBackColor = Color.FromArgb(0, 103, 184);
			button3.UseVisualStyleBackColor = false;
			button3.DialogResult = DialogResult.OK;
			button3 = null;
			button.Paint += delegate
			{
				using GraphicsPath graphicsPath2 = new GraphicsPath();
				Rectangle rectangle2 = new Rectangle(0, 0, button.Width, button.Height);
				int num2 = 6;
				graphicsPath2.AddArc(rectangle2.X, rectangle2.Y, num2, num2, 180f, 90f);
				graphicsPath2.AddArc(rectangle2.X + rectangle2.Width - num2, rectangle2.Y, num2, num2, 270f, 90f);
				graphicsPath2.AddArc(rectangle2.X + rectangle2.Width - num2, rectangle2.Y + rectangle2.Height - num2, num2, num2, 0f, 90f);
				graphicsPath2.AddArc(rectangle2.X, rectangle2.Y + rectangle2.Height - num2, num2, num2, 90f, 90f);
				graphicsPath2.CloseFigure();
				button.Region = new Region(graphicsPath2);
			};
			Button button4 = button2;
			button4.Text = "Cancel";
			button4.Font = new Font("Segoe UI", 9f, FontStyle.Regular);
			button4.Size = new Size(110, 38);
			button4.Location = new Point(335, 115);
			button4.BackColor = Color.FromArgb(235, 235, 240);
			button4.ForeColor = Color.FromArgb(60, 60, 60);
			button4.FlatStyle = FlatStyle.Flat;
			button4.FlatAppearance.BorderSize = 1;
			button4.FlatAppearance.BorderColor = Color.FromArgb(180, 180, 185);
			button4.FlatAppearance.MouseOverBackColor = Color.FromArgb(225, 225, 230);
			button4.FlatAppearance.MouseDownBackColor = Color.FromArgb(215, 215, 220);
			button4.UseVisualStyleBackColor = false;
			button4.DialogResult = DialogResult.Cancel;
			button4 = null;
			button2.Paint += delegate
			{
				using GraphicsPath graphicsPath = new GraphicsPath();
				Rectangle rectangle = new Rectangle(0, 0, button2.Width, button2.Height);
				int num = 6;
				graphicsPath.AddArc(rectangle.X, rectangle.Y, num, num, 180f, 90f);
				graphicsPath.AddArc(rectangle.X + rectangle.Width - num, rectangle.Y, num, num, 270f, 90f);
				graphicsPath.AddArc(rectangle.X + rectangle.Width - num, rectangle.Y + rectangle.Height - num, num, num, 0f, 90f);
				graphicsPath.AddArc(rectangle.X, rectangle.Y + rectangle.Height - num, num, num, 90f, 90f);
				graphicsPath.CloseFigure();
				button2.Region = new Region(graphicsPath);
			};
			textBox.KeyDown += delegate(object sender, KeyEventArgs e)
			{
				if (e.KeyCode == Keys.Return)
				{
					button.PerformClick();
				}
				else if (e.KeyCode == Keys.Escape)
				{
					button2.PerformClick();
				}
			};
			button.MouseEnter += delegate
			{
				closure$__12-._Lambda$__7();
			};
			button.MouseLeave += delegate
			{
				closure$__12-._Lambda$__8();
			};
			button2.MouseEnter += delegate
			{
				closure$__12-._Lambda$__9();
			};
			button2.MouseLeave += delegate
			{
				closure$__12-._Lambda$__10();
			};
			bool flag = false;
			Point point = default(Point);
			panel.MouseDown += delegate
			{
				flag = true;
				point = Cursor.Position;
			};
			panel.MouseMove += delegate
			{
				if (flag)
				{
					Point position2 = Cursor.Position;
					form.Location = new Point(form.Location.X + (position2.X - point.X), form.Location.Y + (position2.Y - point.Y));
					point = position2;
				}
			};
			panel.MouseUp += delegate
			{
				flag = false;
			};
			label.MouseDown += delegate
			{
				flag = true;
				point = Cursor.Position;
			};
			label.MouseMove += delegate
			{
				if (flag)
				{
					Point position = Cursor.Position;
					form.Location = new Point(form.Location.X + (position.X - point.X), form.Location.Y + (position.Y - point.Y));
					point = position;
				}
			};
			label.MouseUp += delegate
			{
				flag = false;
			};
			panel.Controls.Add(label);
			panel2.Controls.AddRange(new Control[4] { label2, textBox, button, button2 });
			form.Controls.AddRange(new Control[2] { panel, panel2 });
			label2.TabIndex = 0;
			textBox.TabIndex = 1;
			button.TabIndex = 2;
			button2.TabIndex = 3;
			form.AcceptButton = button;
			form.CancelButton = button2;
			DialogResult dialogResult = form.ShowDialog();
			if (dialogResult == DialogResult.OK && !string.IsNullOrWhiteSpace(textBox.Text.Trim()))
			{
				return textBox.Text.Trim();
			}
			return string.Empty;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			MessageBox.Show($"Error creating custom dialog: {ex2.Message}", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
			string result = Interaction.InputBox(prompt, title, defaultValue);
			ProjectData.ClearProjectError();
			return result;
		}
		finally
		{
			if (form != null)
			{
				form.Dispose();
			}
		}
	}
}


using System;
using System.Diagnostics;
using System.Drawing;
using System.Windows.Forms;
using Microsoft.VisualBasic.CompilerServices;

private void ShowSuccessMessageWithOpen(string successMessage, string filePath, string folderPath)
{
	Form form = new Form();
	form.Text = "\ud83d\udd8c\ufe0f Designed with Passion by Peeyush Raghav";
	form.Size = new Size(650, 320);
	form.StartPosition = FormStartPosition.CenterScreen;
	form.FormBorderStyle = FormBorderStyle.FixedDialog;
	form.MaximizeBox = false;
	form.MinimizeBox = false;
	form.BackColor = Color.White;
	form.Font = new Font("Segoe UI", 9f, FontStyle.Regular);
	Panel panel = new Panel();
	panel.Size = new Size(650, 60);
	panel.Location = new Point(0, 0);
	panel.BackColor = Color.FromArgb(45, 52, 67);
	Label label = new Label();
	label.Text = "✓";
	label.Font = new Font("Segoe UI", 24f, FontStyle.Bold);
	label.ForeColor = Color.White;
	label.Size = new Size(40, 40);
	label.Location = new Point(20, 10);
	label.TextAlign = ContentAlignment.MiddleCenter;
	Label label2 = new Label();
	label2.Text = "Export Completed Successfully";
	label2.Font = new Font("Segoe UI", 14f, FontStyle.Bold);
	label2.ForeColor = Color.White;
	label2.Size = new Size(500, 30);
	label2.Location = new Point(70, 15);
	label2.TextAlign = ContentAlignment.MiddleLeft;
	Label label3 = new Label();
	label3.Text = successMessage;
	label3.Font = new Font("Segoe UI", 10f, FontStyle.Regular);
	label3.ForeColor = Color.FromArgb(64, 64, 64);
	label3.Size = new Size(610, 80);
	label3.Location = new Point(20, 80);
	label3.AutoSize = false;
	label3.TextAlign = ContentAlignment.TopLeft;
	Panel panel2 = new Panel();
	panel2.Size = new Size(610, 1);
	panel2.Location = new Point(20, 190);
	panel2.BackColor = Color.FromArgb(220, 220, 220);
	Button button = new Button();
	button.Text = "OK";
	button.Size = new Size(120, 40);
	button.Location = new Point(120, 220);
	button.DialogResult = DialogResult.OK;
	button.BackColor = Color.FromArgb(240, 240, 240);
	button.ForeColor = Color.FromArgb(64, 64, 64);
	button.Font = new Font("Segoe UI", 9f, FontStyle.Regular);
	button.FlatStyle = FlatStyle.Flat;
	button.FlatAppearance.BorderSize = 1;
	button.FlatAppearance.BorderColor = Color.FromArgb(180, 180, 180);
	button.FlatAppearance.MouseOverBackColor = Color.FromArgb(230, 230, 230);
	Button button2 = new Button();
	button2.Text = "SHOW IN FOLDER";
	button2.Size = new Size(150, 40);
	button2.Location = new Point(260, 220);
	button2.BackColor = Color.FromArgb(0, 122, 204);
	button2.ForeColor = Color.White;
	button2.Font = new Font("Segoe UI", 9f, FontStyle.Bold);
	button2.FlatStyle = FlatStyle.Flat;
	button2.FlatAppearance.BorderSize = 0;
	button2.FlatAppearance.MouseOverBackColor = Color.FromArgb(0, 100, 180);
	Button button3 = new Button();
	button3.Text = "OPEN IN CAD";
	button3.Size = new Size(130, 40);
	button3.Location = new Point(430, 220);
	button3.BackColor = Color.FromArgb(40, 167, 69);
	button3.ForeColor = Color.White;
	button3.Font = new Font("Segoe UI", 9f, FontStyle.Bold);
	button3.FlatStyle = FlatStyle.Flat;
	button3.FlatAppearance.BorderSize = 0;
	button3.FlatAppearance.MouseOverBackColor = Color.FromArgb(33, 136, 56);
	button2.Click += delegate
	{
		try
		{
			Process.Start("explorer.exe", $"/select,\"{filePath}\"");
		}
		catch (Exception ex3)
		{
			ProjectData.SetProjectError(ex3);
			Exception ex4 = ex3;
			try
			{
				Process.Start("explorer.exe", folderPath);
			}
			catch (Exception projectError2)
			{
				ProjectData.SetProjectError(projectError2);
				MessageBox.Show("Could not open file location.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
				ProjectData.ClearProjectError();
			}
			ProjectData.ClearProjectError();
		}
		form.Close();
	};
	button3.Click += delegate
	{
		try
		{
			Process.Start(filePath);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			try
			{
				Process.Start(new ProcessStartInfo
				{
					FileName = filePath,
					UseShellExecute = true,
					Verb = "open"
				});
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				MessageBox.Show("Could not open DXF file. Please ensure you have a CAD application installed.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
				ProjectData.ClearProjectError();
			}
			ProjectData.ClearProjectError();
		}
		form.Close();
	};
	panel.Controls.Add(label);
	panel.Controls.Add(label2);
	form.Controls.Add(panel);
	form.Controls.Add(label3);
	form.Controls.Add(panel2);
	form.Controls.Add(button);
	form.Controls.Add(button2);
	form.Controls.Add(button3);
	form.ShowDialog();
	form.Dispose();
}


using System.Collections.Generic;
using Microsoft.VisualBasic.CompilerServices;

private string BuildSizeTextWithoutLength(string web, string tf, string bf, double columnFValue, string columnEValue)
{
	web = web?.Replace("PL", "").Trim() ?? "";
	tf = tf?.Replace("PL", "").Trim() ?? "";
	bf = bf?.Replace("PL", "").Trim() ?? "";
	List<string> list = new List<string>();
	if (columnFValue == 680.0)
	{
		if (!string.IsNullOrEmpty(web))
		{
			list.Add($"W-{web}");
		}
		if (!string.IsNullOrEmpty(tf) && !string.IsNullOrEmpty(bf))
		{
			list.Add((Operators.CompareString(tf, bf, TextCompare: false) == 0) ? $"F-{tf}" : $"TF-{tf}\\PBF-{bf}");
		}
		else if (!string.IsNullOrEmpty(tf))
		{
			list.Add($"F-{tf}");
		}
		else if (!string.IsNullOrEmpty(bf))
		{
			list.Add($"F-{bf}");
		}
	}
	else
	{
		if (!string.IsNullOrEmpty(web))
		{
			list.Add(web);
		}
		if (!string.IsNullOrEmpty(tf) && !string.IsNullOrEmpty(bf))
		{
			list.Add((Operators.CompareString(tf, bf, TextCompare: false) == 0) ? tf : $"{tf}\\P{bf}");
		}
		else if (!string.IsNullOrEmpty(tf))
		{
			list.Add(tf);
		}
		else if (!string.IsNullOrEmpty(bf))
		{
			list.Add(bf);
		}
	}
	if (!string.IsNullOrEmpty(columnEValue))
	{
		list.Add($"P-{columnEValue}");
	}
	return string.Join("\\P", list);
}


using System;
using System.Runtime.CompilerServices;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

private bool HasValidBeamData(object excelSheet)
{
	try
	{
		int num = Conversions.ToInteger(NewLateBinding.LateGet(NewLateBinding.LateGet(((Worksheet)excelSheet).Cells[((Worksheet)excelSheet).Rows.Count, 1], null, "End", new object[1] { XlDirection.xlUp }, null, null, null), null, "Row", new object[0], null, null, null));
		int num2 = num;
		for (int i = 9; i <= num2; i = checked(i + 1))
		{
			double safeCellValue = GetSafeCellValue(RuntimeHelpers.GetObjectValue(excelSheet), i, 16);
			double safeCellValue2 = GetSafeCellValue(RuntimeHelpers.GetObjectValue(excelSheet), i, 17);
			double safeCellValue3 = GetSafeCellValue(RuntimeHelpers.GetObjectValue(excelSheet), i, 18);
			double safeCellValue4 = GetSafeCellValue(RuntimeHelpers.GetObjectValue(excelSheet), i, 19);
			double safeCellValue5 = GetSafeCellValue(RuntimeHelpers.GetObjectValue(excelSheet), i, 20);
			double safeCellValue6 = GetSafeCellValue(RuntimeHelpers.GetObjectValue(excelSheet), i, 21);
			if (safeCellValue != 0.0 || safeCellValue2 != 0.0 || safeCellValue3 != 0.0 || safeCellValue4 != 0.0 || safeCellValue5 != 0.0 || safeCellValue6 != 0.0)
			{
				return true;
			}
		}
		return false;
	}
	catch (Exception projectError)
	{
		ProjectData.SetProjectError(projectError);
		bool result = false;
		ProjectData.ClearProjectError();
		return result;
	}
}


using System;
using System.Linq;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;
using netDxf;
using netDxf.Entities;
using netDxf.Tables;

private bool AddGridToDxf(DxfDocument dxfDoc, object excelSheet, string rotationValue)
{
	checked
	{
		try
		{
			Worksheet worksheet = (Worksheet)excelSheet;
			Workbook workbook = (Workbook)worksheet.Parent;
			Worksheet worksheet2 = null;
			try
			{
				worksheet2 = (Worksheet)workbook.Worksheets["Sheet4"];
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				bool result = false;
				ProjectData.ClearProjectError();
				return result;
			}
			string text = GetSafeCellValueAsString(worksheet2, 2, 3).Trim();
			string text2 = GetSafeCellValueAsString(worksheet2, 3, 3).Trim();
			string text3 = GetSafeCellValueAsString(worksheet2, 4, 3).Trim();
			bool includesZero = false;
			bool includesZero2 = false;
			bool includesZero3 = false;
			double[] array = (string.IsNullOrWhiteSpace(text) ? new double[0] : ParseOffsets(text, ref includesZero));
			double[] array2 = (string.IsNullOrWhiteSpace(text2) ? new double[0] : ParseOffsets(text2, ref includesZero2));
			double[] array3 = (string.IsNullOrWhiteSpace(text3) ? new double[0] : ParseOffsets(text3, ref includesZero3));
			if (array.Length == 0 && array2.Length == 0 && array3.Length == 0)
			{
				return false;
			}
			double[] array4 = ((array.Length > 0) ? CalculatePositions(array) : new double[0]);
			double[] array5 = ((array2.Length > 0) ? CalculatePositions(array2) : new double[0]);
			double[] array6 = ((array3.Length > 0) ? CalculatePositions(array3) : new double[0]);
			double result2 = 300.0;
			string text4 = GetSafeCellValueAsString(worksheet2, 2, 5).Trim();
			if (!string.IsNullOrWhiteSpace(text4))
			{
				double.TryParse(text4, out result2);
			}
			double result3 = 1000.0;
			double result4 = 1000.0;
			double result5 = 1000.0;
			string text5 = GetSafeCellValueAsString(worksheet2, 2, 4).Trim();
			if (!string.IsNullOrWhiteSpace(text5))
			{
				double.TryParse(text5, out result3);
			}
			string text6 = GetSafeCellValueAsString(worksheet2, 3, 4).Trim();
			if (!string.IsNullOrWhiteSpace(text6))
			{
				double.TryParse(text6, out result4);
			}
			string text7 = GetSafeCellValueAsString(worksheet2, 4, 4).Trim();
			if (!string.IsNullOrWhiteSpace(text7))
			{
				double.TryParse(text7, out result5);
			}
			string[] array7 = (string.IsNullOrWhiteSpace(GetSafeCellValueAsString(worksheet2, 2, 2)) ? new string[0] : ParseNames(GetSafeCellValueAsString(worksheet2, 2, 2)));
			string[] array8 = (string.IsNullOrWhiteSpace(GetSafeCellValueAsString(worksheet2, 3, 2)) ? new string[0] : ParseNames(GetSafeCellValueAsString(worksheet2, 3, 2)));
			string[] array9 = (string.IsNullOrWhiteSpace(GetSafeCellValueAsString(worksheet2, 4, 2)) ? new string[0] : ParseNames(GetSafeCellValueAsString(worksheet2, 4, 2)));
			double[] source = ((array4.Length > 0) ? array4 : new double[1]);
			double[] source2 = ((array5.Length > 0) ? array5 : new double[1]);
			double[] source3 = ((array6.Length > 0) ? array6 : new double[1]);
			double num = ((array4.Length > 0) ? source.Min() : 0.0);
			double num2 = ((array4.Length > 0) ? source.Max() : 0.0);
			double num3 = ((array5.Length > 0) ? source2.Min() : 0.0);
			double num4 = ((array5.Length > 0) ? source2.Max() : 0.0);
			double num5 = ((array6.Length > 0) ? source3.Min() : 0.0);
			double num6 = ((array6.Length > 0) ? source3.Max() : 0.0);
			double x = ((includesZero || array4.Length == 0) ? 0.0 : num);
			double y = ((includesZero2 || array5.Length == 0) ? 0.0 : num3);
			double z = ((includesZero3 || array6.Length == 0) ? 0.0 : num5);
			AciColor color = AciColor.FromCadIndex(6);
			double num7 = result2 * 0.8;
			if (array4.Length > 0)
			{
				Layer layer = new Layer("GRID_X")
				{
					Color = color
				};
				dxfDoc.Layers.Add(layer);
				int num8 = array4.Length - 1;
				for (int i = 0; i <= num8; i++)
				{
					double x2 = array4[i];
					string text8 = ((i < array7.Length) ? array7[i] : $"X{i + 1}");
					if (array5.Length > 0)
					{
						Vector3 startPoint = ApplyRotation(new Vector3(x2, num3 - result3 + num7, z), rotationValue);
						Vector3 endPoint = ApplyRotation(new Vector3(x2, num4 + result3 - num7, z), rotationValue);
						dxfDoc.Entities.Add(new Line(startPoint, endPoint)
						{
							Layer = layer
						});
						Vector3 vector = ApplyRotation(new Vector3(x2, num3 - result3, z), rotationValue);
						Vector3 vector2 = ApplyRotation(new Vector3(x2, num4 + result3, z), rotationValue);
						dxfDoc.Entities.Add(new Text(text8, vector, result2)
						{
							Layer = layer,
							Alignment = TextAlignment.MiddleCenter
						});
						dxfDoc.Entities.Add(new Circle(vector, num7)
						{
							Layer = layer
						});
						dxfDoc.Entities.Add(new Text(text8, vector2, result2)
						{
							Layer = layer,
							Alignment = TextAlignment.MiddleCenter
						});
						dxfDoc.Entities.Add(new Circle(vector2, num7)
						{
							Layer = layer
						});
					}
					if (array6.Length > 0)
					{
						Vector3 startPoint2 = ApplyRotation(new Vector3(x2, y, num5 - result3 + num7), rotationValue);
						Vector3 endPoint2 = ApplyRotation(new Vector3(x2, y, num6 + result3 - num7), rotationValue);
						dxfDoc.Entities.Add(new Line(startPoint2, endPoint2)
						{
							Layer = layer
						});
						Vector3 vector3 = ApplyRotation(new Vector3(x2, y, num5 - result3), rotationValue);
						Vector3 vector4 = ApplyRotation(new Vector3(x2, y, num6 + result3), rotationValue);
						dxfDoc.Entities.Add(new Text(text8, vector3, result2)
						{
							Layer = layer,
							Alignment = TextAlignment.MiddleCenter
						});
						dxfDoc.Entities.Add(new Circle(vector3, num7)
						{
							Layer = layer
						});
						dxfDoc.Entities.Add(new Text(text8, vector4, result2)
						{
							Layer = layer,
							Alignment = TextAlignment.MiddleCenter
						});
						dxfDoc.Entities.Add(new Circle(vector4, num7)
						{
							Layer = layer
						});
					}
				}
				if (array5.Length > 0 && array6.Length > 0)
				{
					int num9 = array6.Length - 1;
					for (int j = 0; j <= num9; j++)
					{
						double z2 = array6[j];
						string text9 = ((j < array9.Length) ? array9[j] : $"Z{j + 1}");
						Vector3 startPoint3 = ApplyRotation(new Vector3(num, num3 - result3 + num7, z2), rotationValue);
						Vector3 endPoint3 = ApplyRotation(new Vector3(num, num4 + result3 - num7, z2), rotationValue);
						dxfDoc.Entities.Add(new Line(startPoint3, endPoint3)
						{
							Layer = layer
						});
						Vector3 vector5 = ApplyRotation(new Vector3(num, num3 - result3, z2), rotationValue);
						Vector3 vector6 = ApplyRotation(new Vector3(num, num4 + result3, z2), rotationValue);
						dxfDoc.Entities.Add(new Text(text9, vector5, result2)
						{
							Layer = layer,
							Alignment = TextAlignment.MiddleCenter
						});
						dxfDoc.Entities.Add(new Circle(vector5, num7)
						{
							Layer = layer
						});
						dxfDoc.Entities.Add(new Text(text9, vector6, result2)
						{
							Layer = layer,
							Alignment = TextAlignment.MiddleCenter
						});
						dxfDoc.Entities.Add(new Circle(vector6, num7)
						{
							Layer = layer
						});
					}
					int num10 = array5.Length - 1;
					for (int k = 0; k <= num10; k++)
					{
						double y2 = array5[k];
						string text10 = ((k < array8.Length) ? array8[k] : $"Y{k + 1}");
						Vector3 startPoint4 = ApplyRotation(new Vector3(num, y2, num5 - result3 + num7), rotationValue);
						Vector3 endPoint4 = ApplyRotation(new Vector3(num, y2, num6 + result3 - num7), rotationValue);
						dxfDoc.Entities.Add(new Line(startPoint4, endPoint4)
						{
							Layer = layer
						});
						Vector3 vector7 = ApplyRotation(new Vector3(num, y2, num5 - result3), rotationValue);
						Vector3 vector8 = ApplyRotation(new Vector3(num, y2, num6 + result3), rotationValue);
						dxfDoc.Entities.Add(new Text(text10, vector7, result2)
						{
							Layer = layer,
							Alignment = TextAlignment.MiddleCenter
						});
						dxfDoc.Entities.Add(new Circle(vector7, num7)
						{
							Layer = layer
						});
						dxfDoc.Entities.Add(new Text(text10, vector8, result2)
						{
							Layer = layer,
							Alignment = TextAlignment.MiddleCenter
						});
						dxfDoc.Entities.Add(new Circle(vector8, num7)
						{
							Layer = layer
						});
					}
				}
				if (array5.Length > 0 && array6.Length > 0)
				{
					int num11 = array6.Length - 1;
					for (int l = 0; l <= num11; l++)
					{
						double z3 = array6[l];
						string text11 = ((l < array9.Length) ? array9[l] : $"Z{l + 1}");
						Vector3 startPoint5 = ApplyRotation(new Vector3(num2, num3 - result3 + num7, z3), rotationValue);
						Vector3 endPoint5 = ApplyRotation(new Vector3(num2, num4 + result3 - num7, z3), rotationValue);
						dxfDoc.Entities.Add(new Line(startPoint5, endPoint5)
						{
							Layer = layer
						});
						Vector3 vector9 = ApplyRotation(new Vector3(num2, num3 - result3, z3), rotationValue);
						Vector3 vector10 = ApplyRotation(new Vector3(num2, num4 + result3, z3), rotationValue);
						dxfDoc.Entities.Add(new Text(text11, vector9, result2)
						{
							Layer = layer,
							Alignment = TextAlignment.MiddleCenter
						});
						dxfDoc.Entities.Add(new Circle(vector9, num7)
						{
							Layer = layer
						});
						dxfDoc.Entities.Add(new Text(text11, vector10, result2)
						{
							Layer = layer,
							Alignment = TextAlignment.MiddleCenter
						});
						dxfDoc.Entities.Add(new Circle(vector10, num7)
						{
							Layer = layer
						});
					}
					int num12 = array5.Length - 1;
					for (int m = 0; m <= num12; m++)
					{
						double y3 = array5[m];
						string text12 = ((m < array8.Length) ? array8[m] : $"Y{m + 1}");
						Vector3 startPoint6 = ApplyRotation(new Vector3(num2, y3, num5 - result3 + num7), rotationValue);
						Vector3 endPoint6 = ApplyRotation(new Vector3(num2, y3, num6 + result3 - num7), rotationValue);
						dxfDoc.Entities.Add(new Line(startPoint6, endPoint6)
						{
							Layer = layer
						});
						Vector3 vector11 = ApplyRotation(new Vector3(num2, y3, num5 - result3), rotationValue);
						Vector3 vector12 = ApplyRotation(new Vector3(num2, y3, num6 + result3), rotationValue);
						dxfDoc.Entities.Add(new Text(text12, vector11, result2)
						{
							Layer = layer,
							Alignment = TextAlignment.MiddleCenter
						});
						dxfDoc.Entities.Add(new Circle(vector11, num7)
						{
							Layer = layer
						});
						dxfDoc.Entities.Add(new Text(text12, vector12, result2)
						{
							Layer = layer,
							Alignment = TextAlignment.MiddleCenter
						});
						dxfDoc.Entities.Add(new Circle(vector12, num7)
						{
							Layer = layer
						});
					}
				}
			}
			if (array5.Length > 0)
			{
				Layer layer2 = new Layer("GRID_Y")
				{
					Color = color
				};
				dxfDoc.Layers.Add(layer2);
				int num13 = array5.Length - 1;
				for (int n = 0; n <= num13; n++)
				{
					double y4 = array5[n];
					string text13 = ((n < array8.Length) ? array8[n] : $"Y{n + 1}");
					double num14 = num7;
					if (array4.Length > 0)
					{
						Vector3 startPoint7 = ApplyRotation(new Vector3(num - result4 + num14, y4, z), rotationValue);
						Vector3 endPoint7 = ApplyRotation(new Vector3(num2 + result4 - num14, y4, z), rotationValue);
						dxfDoc.Entities.Add(new Line(startPoint7, endPoint7)
						{
							Layer = layer2
						});
						Vector3 vector13 = ApplyRotation(new Vector3(num - result4, y4, z), rotationValue);
						Vector3 vector14 = ApplyRotation(new Vector3(num2 + result4, y4, z), rotationValue);
						dxfDoc.Entities.Add(new Text(text13, vector13, result2)
						{
							Layer = layer2,
							Alignment = TextAlignment.MiddleCenter
						});
						dxfDoc.Entities.Add(new Circle(vector13, num7)
						{
							Layer = layer2
						});
						dxfDoc.Entities.Add(new Text(text13, vector14, result2)
						{
							Layer = layer2,
							Alignment = TextAlignment.MiddleCenter
						});
						dxfDoc.Entities.Add(new Circle(vector14, num7)
						{
							Layer = layer2
						});
					}
					if (array6.Length > 0)
					{
						Vector3 startPoint8 = ApplyRotation(new Vector3(x, y4, num5 - result4 + num14), rotationValue);
						Vector3 endPoint8 = ApplyRotation(new Vector3(x, y4, num6 + result4 - num14), rotationValue);
						dxfDoc.Entities.Add(new Line(startPoint8, endPoint8)
						{
							Layer = layer2
						});
						Vector3 vector15 = ApplyRotation(new Vector3(x, y4, num5 - result4), rotationValue);
						Vector3 vector16 = ApplyRotation(new Vector3(x, y4, num6 + result4), rotationValue);
						dxfDoc.Entities.Add(new Text(text13, vector15, result2)
						{
							Layer = layer2,
							Alignment = TextAlignment.MiddleCenter
						});
						dxfDoc.Entities.Add(new Circle(vector15, num7)
						{
							Layer = layer2
						});
						dxfDoc.Entities.Add(new Text(text13, vector16, result2)
						{
							Layer = layer2,
							Alignment = TextAlignment.MiddleCenter
						});
						dxfDoc.Entities.Add(new Circle(vector16, num7)
						{
							Layer = layer2
						});
					}
				}
				if (array4.Length > 0 && array6.Length > 0)
				{
					int num15 = array6.Length - 1;
					for (int num16 = 0; num16 <= num15; num16++)
					{
						double z4 = array6[num16];
						string text14 = ((num16 < array9.Length) ? array9[num16] : $"Z{num16 + 1}");
						Vector3 startPoint9 = ApplyRotation(new Vector3(num - result4 + num7, num3, z4), rotationValue);
						Vector3 endPoint9 = ApplyRotation(new Vector3(num2 + result4 - num7, num3, z4), rotationValue);
						dxfDoc.Entities.Add(new Line(startPoint9, endPoint9)
						{
							Layer = layer2
						});
						Vector3 vector17 = ApplyRotation(new Vector3(num - result4, num3, z4), rotationValue);
						Vector3 vector18 = ApplyRotation(new Vector3(num2 + result4, num3, z4), rotationValue);
						dxfDoc.Entities.Add(new Text(text14, vector17, result2)
						{
							Layer = layer2,
							Alignment = TextAlignment.MiddleCenter
						});
						dxfDoc.Entities.Add(new Circle(vector17, num7)
						{
							Layer = layer2
						});
						dxfDoc.Entities.Add(new Text(text14, vector18, result2)
						{
							Layer = layer2,
							Alignment = TextAlignment.MiddleCenter
						});
						dxfDoc.Entities.Add(new Circle(vector18, num7)
						{
							Layer = layer2
						});
					}
					int num17 = array4.Length - 1;
					for (int num18 = 0; num18 <= num17; num18++)
					{
						double x3 = array4[num18];
						string text15 = ((num18 < array7.Length) ? array7[num18] : $"X{num18 + 1}");
						Vector3 startPoint10 = ApplyRotation(new Vector3(x3, num3, num5 - result4 + num7), rotationValue);
						Vector3 endPoint10 = ApplyRotation(new Vector3(x3, num3, num6 + result4 - num7), rotationValue);
						dxfDoc.Entities.Add(new Line(startPoint10, endPoint10)
						{
							Layer = layer2
						});
						Vector3 vector19 = ApplyRotation(new Vector3(x3, num3, num5 - result4), rotationValue);
						Vector3 vector20 = ApplyRotation(new Vector3(x3, num3, num6 + result4), rotationValue);
						dxfDoc.Entities.Add(new Text(text15, vector19, result2)
						{
							Layer = layer2,
							Alignment = TextAlignment.MiddleCenter
						});
						dxfDoc.Entities.Add(new Circle(vector19, num7)
						{
							Layer = layer2
						});
						dxfDoc.Entities.Add(new Text(text15, vector20, result2)
						{
							Layer = layer2,
							Alignment = TextAlignment.MiddleCenter
						});
						dxfDoc.Entities.Add(new Circle(vector20, num7)
						{
							Layer = layer2
						});
					}
				}
				if (array4.Length > 0 && array6.Length > 0)
				{
					int num19 = array6.Length - 1;
					for (int num20 = 0; num20 <= num19; num20++)
					{
						double z5 = array6[num20];
						string text16 = ((num20 < array9.Length) ? array9[num20] : $"Z{num20 + 1}");
						Vector3 startPoint11 = ApplyRotation(new Vector3(num - result4 + num7, num4, z5), rotationValue);
						Vector3 endPoint11 = ApplyRotation(new Vector3(num2 + result4 - num7, num4, z5), rotationValue);
						dxfDoc.Entities.Add(new Line(startPoint11, endPoint11)
						{
							Layer = layer2
						});
						Vector3 vector21 = ApplyRotation(new Vector3(num - result4, num4, z5), rotationValue);
						Vector3 vector22 = ApplyRotation(new Vector3(num2 + result4, num4, z5), rotationValue);
						dxfDoc.Entities.Add(new Text(text16, vector21, result2)
						{
							Layer = layer2,
							Alignment = TextAlignment.MiddleCenter
						});
						dxfDoc.Entities.Add(new Circle(vector21, num7)
						{
							Layer = layer2
						});
						dxfDoc.Entities.Add(new Text(text16, vector22, result2)
						{
							Layer = layer2,
							Alignment = TextAlignment.MiddleCenter
						});
						dxfDoc.Entities.Add(new Circle(vector22, num7)
						{
							Layer = layer2
						});
					}
					int num21 = array4.Length - 1;
					for (int num22 = 0; num22 <= num21; num22++)
					{
						double x4 = array4[num22];
						string text17 = ((num22 < array7.Length) ? array7[num22] : $"X{num22 + 1}");
						Vector3 startPoint12 = ApplyRotation(new Vector3(x4, num4, num5 - result4 + num7), rotationValue);
						Vector3 endPoint12 = ApplyRotation(new Vector3(x4, num4, num6 + result4 - num7), rotationValue);
						dxfDoc.Entities.Add(new Line(startPoint12, endPoint12)
						{
							Layer = layer2
						});
						Vector3 vector23 = ApplyRotation(new Vector3(x4, num4, num5 - result4), rotationValue);
						Vector3 vector24 = ApplyRotation(new Vector3(x4, num4, num6 + result4), rotationValue);
						dxfDoc.Entities.Add(new Text(text17, vector23, result2)
						{
							Layer = layer2,
							Alignment = TextAlignment.MiddleCenter
						});
						dxfDoc.Entities.Add(new Circle(vector23, num7)
						{
							Layer = layer2
						});
						dxfDoc.Entities.Add(new Text(text17, vector24, result2)
						{
							Layer = layer2,
							Alignment = TextAlignment.MiddleCenter
						});
						dxfDoc.Entities.Add(new Circle(vector24, num7)
						{
							Layer = layer2
						});
					}
				}
			}
			if (array6.Length > 0)
			{
				Layer layer3 = new Layer("GRID_Z")
				{
					Color = color
				};
				dxfDoc.Layers.Add(layer3);
				int num23 = array6.Length - 1;
				for (int num24 = 0; num24 <= num23; num24++)
				{
					double z6 = array6[num24];
					string text18 = ((num24 < array9.Length) ? array9[num24] : $"Z{num24 + 1}");
					double num25 = num7;
					if (array4.Length > 0)
					{
						Vector3 startPoint13 = ApplyRotation(new Vector3(num - result5 + num25, y, z6), rotationValue);
						Vector3 endPoint13 = ApplyRotation(new Vector3(num2 + result5 - num25, y, z6), rotationValue);
						dxfDoc.Entities.Add(new Line(startPoint13, endPoint13)
						{
							Layer = layer3
						});
						Vector3 vector25 = ApplyRotation(new Vector3(num - result5, y, z6), rotationValue);
						Vector3 vector26 = ApplyRotation(new Vector3(num2 + result5, y, z6), rotationValue);
						dxfDoc.Entities.Add(new Text(text18, vector25, result2)
						{
							Layer = layer3,
							Alignment = TextAlignment.MiddleCenter
						});
						dxfDoc.Entities.Add(new Circle(vector25, num7)
						{
							Layer = layer3
						});
						dxfDoc.Entities.Add(new Text(text18, vector26, result2)
						{
							Layer = layer3,
							Alignment = TextAlignment.MiddleCenter
						});
						dxfDoc.Entities.Add(new Circle(vector26, num7)
						{
							Layer = layer3
						});
					}
					if (array5.Length > 0)
					{
						Vector3 startPoint14 = ApplyRotation(new Vector3(x, num3 - result5 + num25, z6), rotationValue);
						Vector3 endPoint14 = ApplyRotation(new Vector3(x, num4 + result5 - num25, z6), rotationValue);
						dxfDoc.Entities.Add(new Line(startPoint14, endPoint14)
						{
							Layer = layer3
						});
						Vector3 vector27 = ApplyRotation(new Vector3(x, num3 - result5, z6), rotationValue);
						Vector3 vector28 = ApplyRotation(new Vector3(x, num4 + result5, z6), rotationValue);
						dxfDoc.Entities.Add(new Text(text18, vector27, result2)
						{
							Layer = layer3,
							Alignment = TextAlignment.MiddleCenter
						});
						dxfDoc.Entities.Add(new Circle(vector27, num7)
						{
							Layer = layer3
						});
						dxfDoc.Entities.Add(new Text(text18, vector28, result2)
						{
							Layer = layer3,
							Alignment = TextAlignment.MiddleCenter
						});
						dxfDoc.Entities.Add(new Circle(vector28, num7)
						{
							Layer = layer3
						});
					}
				}
				if (array4.Length > 0 && array5.Length > 0)
				{
					int num26 = array5.Length - 1;
					for (int num27 = 0; num27 <= num26; num27++)
					{
						double y5 = array5[num27];
						string text19 = ((num27 < array8.Length) ? array8[num27] : $"Y{num27 + 1}");
						Vector3 startPoint15 = ApplyRotation(new Vector3(num - result5 + num7, y5, num5), rotationValue);
						Vector3 endPoint15 = ApplyRotation(new Vector3(num2 + result5 - num7, y5, num5), rotationValue);
						dxfDoc.Entities.Add(new Line(startPoint15, endPoint15)
						{
							Layer = layer3
						});
						Vector3 vector29 = ApplyRotation(new Vector3(num - result5, y5, num5), rotationValue);
						Vector3 vector30 = ApplyRotation(new Vector3(num2 + result5, y5, num5), rotationValue);
						dxfDoc.Entities.Add(new Text(text19, vector29, result2)
						{
							Layer = layer3,
							Alignment = TextAlignment.MiddleCenter
						});
						dxfDoc.Entities.Add(new Circle(vector29, num7)
						{
							Layer = layer3
						});
						dxfDoc.Entities.Add(new Text(text19, vector30, result2)
						{
							Layer = layer3,
							Alignment = TextAlignment.MiddleCenter
						});
						dxfDoc.Entities.Add(new Circle(vector30, num7)
						{
							Layer = layer3
						});
					}
					int num28 = array4.Length - 1;
					for (int num29 = 0; num29 <= num28; num29++)
					{
						double x5 = array4[num29];
						string text20 = ((num29 < array7.Length) ? array7[num29] : $"X{num29 + 1}");
						Vector3 startPoint16 = ApplyRotation(new Vector3(x5, num3 - result5 + num7, num5), rotationValue);
						Vector3 endPoint16 = ApplyRotation(new Vector3(x5, num4 + result5 - num7, num5), rotationValue);
						dxfDoc.Entities.Add(new Line(startPoint16, endPoint16)
						{
							Layer = layer3
						});
						Vector3 vector31 = ApplyRotation(new Vector3(x5, num3 - result5, num5), rotationValue);
						Vector3 vector32 = ApplyRotation(new Vector3(x5, num4 + result5, num5), rotationValue);
						dxfDoc.Entities.Add(new Text(text20, vector31, result2)
						{
							Layer = layer3,
							Alignment = TextAlignment.MiddleCenter
						});
						dxfDoc.Entities.Add(new Circle(vector31, num7)
						{
							Layer = layer3
						});
						dxfDoc.Entities.Add(new Text(text20, vector32, result2)
						{
							Layer = layer3,
							Alignment = TextAlignment.MiddleCenter
						});
						dxfDoc.Entities.Add(new Circle(vector32, num7)
						{
							Layer = layer3
						});
					}
				}
				if (array4.Length > 0 && array5.Length > 0)
				{
					int num30 = array5.Length - 1;
					for (int num31 = 0; num31 <= num30; num31++)
					{
						double y6 = array5[num31];
						string text21 = ((num31 < array8.Length) ? array8[num31] : $"Y{num31 + 1}");
						Vector3 startPoint17 = ApplyRotation(new Vector3(num - result5 + num7, y6, num6), rotationValue);
						Vector3 endPoint17 = ApplyRotation(new Vector3(num2 + result5 - num7, y6, num6), rotationValue);
						dxfDoc.Entities.Add(new Line(startPoint17, endPoint17)
						{
							Layer = layer3
						});
						Vector3 vector33 = ApplyRotation(new Vector3(num - result5, y6, num6), rotationValue);
						Vector3 vector34 = ApplyRotation(new Vector3(num2 + result5, y6, num6), rotationValue);
						dxfDoc.Entities.Add(new Text(text21, vector33, result2)
						{
							Layer = layer3,
							Alignment = TextAlignment.MiddleCenter
						});
						dxfDoc.Entities.Add(new Circle(vector33, num7)
						{
							Layer = layer3
						});
						dxfDoc.Entities.Add(new Text(text21, vector34, result2)
						{
							Layer = layer3,
							Alignment = TextAlignment.MiddleCenter
						});
						dxfDoc.Entities.Add(new Circle(vector34, num7)
						{
							Layer = layer3
						});
					}
					int num32 = array4.Length - 1;
					for (int num33 = 0; num33 <= num32; num33++)
					{
						double x6 = array4[num33];
						string text22 = ((num33 < array7.Length) ? array7[num33] : $"X{num33 + 1}");
						Vector3 startPoint18 = ApplyRotation(new Vector3(x6, num3 - result5 + num7, num6), rotationValue);
						Vector3 endPoint18 = ApplyRotation(new Vector3(x6, num4 + result5 - num7, num6), rotationValue);
						dxfDoc.Entities.Add(new Line(startPoint18, endPoint18)
						{
							Layer = layer3
						});
						Vector3 vector35 = ApplyRotation(new Vector3(x6, num3 - result5, num6), rotationValue);
						Vector3 vector36 = ApplyRotation(new Vector3(x6, num4 + result5, num6), rotationValue);
						dxfDoc.Entities.Add(new Text(text22, vector35, result2)
						{
							Layer = layer3,
							Alignment = TextAlignment.MiddleCenter
						});
						dxfDoc.Entities.Add(new Circle(vector35, num7)
						{
							Layer = layer3
						});
						dxfDoc.Entities.Add(new Text(text22, vector36, result2)
						{
							Layer = layer3,
							Alignment = TextAlignment.MiddleCenter
						});
						dxfDoc.Entities.Add(new Circle(vector36, num7)
						{
							Layer = layer3
						});
					}
				}
			}
			return true;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			bool result = false;
			ProjectData.ClearProjectError();
			return result;
		}
	}
}


using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text.RegularExpressions;
using Microsoft.VisualBasic.CompilerServices;

private double[] ParseOffsets(string offsetString, ref bool includesZero)
{
	includesZero = false;
	try
	{
		if (string.IsNullOrWhiteSpace(offsetString))
		{
			return new double[0];
		}
		string input = offsetString.Replace(',', ' ').Replace(';', ' ').Replace('\t', ' ')
			.Trim();
		string[] array = Regex.Split(input, "\\s+");
		List<double> list = new List<double>();
		string[] array2 = array;
		foreach (string text in array2)
		{
			if (string.IsNullOrWhiteSpace(text))
			{
				continue;
			}
			double result3;
			if (text.Contains("*"))
			{
				string[] array3 = text.Split('*');
				if (array3.Length != 2 || !int.TryParse(array3[0].Trim(), out var result) || result <= 0 || (!double.TryParse(array3[1].Trim(), NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out var result2) && !double.TryParse(array3[1].Trim(), NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.CurrentCulture, out result2)))
				{
					continue;
				}
				int num = result;
				for (int j = 1; j <= num; j = checked(j + 1))
				{
					list.Add(result2);
					if (Math.Abs(result2) < 1E-09)
					{
						includesZero = true;
					}
				}
			}
			else if (double.TryParse(text, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out result3) || double.TryParse(text, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.CurrentCulture, out result3))
			{
				list.Add(result3);
				if (Math.Abs(result3) < 1E-09)
				{
					includesZero = true;
				}
			}
		}
		return list.ToArray();
	}
	catch (Exception projectError)
	{
		ProjectData.SetProjectError(projectError);
		double[] result4 = new double[0];
		ProjectData.ClearProjectError();
		return result4;
	}
}


using System;
using System.Linq;
using Microsoft.VisualBasic.CompilerServices;

private string[] ParseNames(string nameString)
{
	try
	{
		if (string.IsNullOrWhiteSpace(nameString))
		{
			return new string[0];
		}
		string[] source = nameString.Split(new char[1] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
		return (from p in source
			select p.Trim() into n
			where Operators.CompareString(n, "", TextCompare: false) != 0
			select n).ToArray();
	}
	catch (Exception projectError)
	{
		ProjectData.SetProjectError(projectError);
		string[] result = new string[0];
		ProjectData.ClearProjectError();
		return result;
	}
}


using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualBasic.CompilerServices;

private double[] CalculatePositions(double[] offsets)
{
	checked
	{
		try
		{
			if (offsets.Length == 0)
			{
				return new double[0];
			}
			List<double> list = new List<double>();
			double num = offsets[0];
			list.Add(num);
			int num2 = offsets.Length - 1;
			for (int i = 1; i <= num2; i++)
			{
				num += offsets[i];
				list.Add(num);
			}
			return (from x in list.Distinct()
				orderby x
				select x).ToArray();
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			double[] result = new double[0];
			ProjectData.ClearProjectError();
			return result;
		}
	}
}


using System;
using System.Runtime.CompilerServices;

private double? GetLeaderDirectionAngle(object sheet)
{
	return GetSafeCellValueAsString(RuntimeHelpers.GetObjectValue(sheet), 3, 45).Trim().ToUpper() switch
	{
		"D" => -Math.PI / 4.0, 
		"U" => Math.PI / 4.0, 
		"N" => null, 
		_ => Math.PI / 4.0, 
	};
}


using netDxf;

private Vector3 ApplyRotation(Vector3 point, string rotationValue)
{
	Vector3 result;
	switch (string.IsNullOrWhiteSpace(rotationValue) ? "1" : rotationValue.Trim())
	{
	case "2":
		result = new Vector3(0.0 - point.X, point.Y, 0.0 - point.Z);
		break;
	case "3":
		result = new Vector3(point.Z, point.Y, 0.0 - point.X);
		break;
	case "4":
		result = new Vector3(0.0 - point.Z, point.Y, point.X);
		break;
	case "5":
		result = new Vector3(point.X, 0.0 - point.Z, point.Y);
		break;
	default:
		return point;
	}
	return result;
}


using System;
using System.Collections.Generic;
using Microsoft.VisualBasic.CompilerServices;

private string BuildSizeText(string web, string tf, string bf, double beamLength, double columnFValue)
{
	web = web?.Replace("PL", "").Trim() ?? "";
	tf = tf?.Replace("PL", "").Trim() ?? "";
	bf = bf?.Replace("PL", "").Trim() ?? "";
	List<string> list = new List<string>();
	if (columnFValue == 680.0)
	{
		if (!string.IsNullOrEmpty(web))
		{
			list.Add($"W-{web}");
		}
		if (!string.IsNullOrEmpty(tf) && !string.IsNullOrEmpty(bf))
		{
			list.Add((Operators.CompareString(tf, bf, TextCompare: false) == 0) ? $"F-{tf}" : $"TF-{tf}\\PBF-{bf}");
		}
		else if (!string.IsNullOrEmpty(tf))
		{
			list.Add($"F-{tf}");
		}
		else if (!string.IsNullOrEmpty(bf))
		{
			list.Add($"F-{bf}");
		}
	}
	else
	{
		if (!string.IsNullOrEmpty(web))
		{
			list.Add(web);
		}
		if (!string.IsNullOrEmpty(tf) && !string.IsNullOrEmpty(bf))
		{
			list.Add((Operators.CompareString(tf, bf, TextCompare: false) == 0) ? tf : $"{tf}\\P{bf}");
		}
		else if (!string.IsNullOrEmpty(tf))
		{
			list.Add(tf);
		}
		else if (!string.IsNullOrEmpty(bf))
		{
			list.Add(bf);
		}
	}
	if (beamLength > 0.0)
	{
		list.Add($"L-{checked((int)Math.Round(beamLength))}");
	}
	return string.Join("\\P", list);
}


using System;
using System.Runtime.CompilerServices;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

private double GetSafeCellValue(object excelSheet, int row, int col)
{
	try
	{
		object objectValue = RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(((Worksheet)excelSheet).Cells[row, col], null, "Value", new object[0], null, null, null));
		if (objectValue == null)
		{
			return 0.0;
		}
		double result;
		return double.TryParse(objectValue.ToString(), out result) ? result : 0.0;
	}
	catch (Exception projectError)
	{
		ProjectData.SetProjectError(projectError);
		double result2 = 0.0;
		ProjectData.ClearProjectError();
		return result2;
	}
}


using System;
using System.Runtime.CompilerServices;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

private string GetSafeCellValueAsString(object excelSheet, int row, int col)
{
	try
	{
		return RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(((Worksheet)excelSheet).Cells[row, col], null, "Value", new object[0], null, null, null))?.ToString() ?? "";
	}
	catch (Exception projectError)
	{
		ProjectData.SetProjectError(projectError);
		string result = "";
		ProjectData.ClearProjectError();
		return result;
	}
}


using System;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using Microsoft.VisualBasic.CompilerServices;
using netDxf;
using netDxf.Tables;

private Dictionary<string, Layer> CreateColorBasedLayers(DxfDocument dxfDoc, object excelSheet, int lastRow)
{
	Dictionary<string, Layer> dictionary = new Dictionary<string, Layer>();
	int num = 1;
	checked
	{
		for (int i = 9; i <= lastRow; i++)
		{
			try
			{
				double safeCellValue = GetSafeCellValue(RuntimeHelpers.GetObjectValue(excelSheet), i, 16);
				double safeCellValue2 = GetSafeCellValue(RuntimeHelpers.GetObjectValue(excelSheet), i, 17);
				double safeCellValue3 = GetSafeCellValue(RuntimeHelpers.GetObjectValue(excelSheet), i, 18);
				double safeCellValue4 = GetSafeCellValue(RuntimeHelpers.GetObjectValue(excelSheet), i, 19);
				double safeCellValue5 = GetSafeCellValue(RuntimeHelpers.GetObjectValue(excelSheet), i, 20);
				double safeCellValue6 = GetSafeCellValue(RuntimeHelpers.GetObjectValue(excelSheet), i, 21);
				if (safeCellValue == 0.0 && safeCellValue2 == 0.0 && safeCellValue3 == 0.0 && safeCellValue4 == 0.0 && safeCellValue5 == 0.0 && safeCellValue6 == 0.0)
				{
					continue;
				}
				string text = "";
				int num2 = 7;
				do
				{
					string safeCellValueAsString = GetSafeCellValueAsString(RuntimeHelpers.GetObjectValue(excelSheet), i, num2);
					text = text + safeCellValueAsString + "|";
					num2++;
				}
				while (num2 <= 13);
				if (text.EndsWith("|"))
				{
					text = text.Substring(0, text.Length - 1);
				}
				if (!dictionary.ContainsKey(text))
				{
					unchecked
					{
						int num3 = checked(num - 1) % memberColors.Length;
						AciColor color = memberColors[num3];
						string name = CreateLayerNameFromBeamText(RuntimeHelpers.GetObjectValue(excelSheet), i);
						Layer layer = null;
						if (dxfDoc.Layers.Contains(name))
						{
							layer = dxfDoc.Layers[name];
						}
						else
						{
							layer = new Layer(name);
							layer.Color = color;
							dxfDoc.Layers.Add(layer);
						}
						dictionary.Add(text, layer);
					}
					num++;
				}
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				ProjectData.ClearProjectError();
			}
		}
		return dictionary;
	}
}


using System;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using Microsoft.VisualBasic.CompilerServices;
using netDxf.Tables;

private Layer GetBeamLayerByParameters(object excelSheet, int rowIndex, Dictionary<string, Layer> parameterLayerMap, Layer defaultLayer)
{
	checked
	{
		try
		{
			string text = "";
			int num = 7;
			do
			{
				string safeCellValueAsString = GetSafeCellValueAsString(RuntimeHelpers.GetObjectValue(excelSheet), rowIndex, num);
				text = text + safeCellValueAsString + "|";
				num++;
			}
			while (num <= 13);
			if (text.EndsWith("|"))
			{
				text = text.Substring(0, text.Length - 1);
			}
			if (parameterLayerMap.ContainsKey(text))
			{
				return parameterLayerMap[text];
			}
			return defaultLayer;
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
			return defaultLayer;
		}
	}
}


using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using Microsoft.VisualBasic.CompilerServices;

private string CreateLayerNameFromBeamText(object excelSheet, int rowIndex)
{
	try
	{
		string safeCellValueAsString = GetSafeCellValueAsString(RuntimeHelpers.GetObjectValue(excelSheet), rowIndex, 37);
		string safeCellValueAsString2 = GetSafeCellValueAsString(RuntimeHelpers.GetObjectValue(excelSheet), rowIndex, 38);
		string safeCellValueAsString3 = GetSafeCellValueAsString(RuntimeHelpers.GetObjectValue(excelSheet), rowIndex, 39);
		double safeCellValue = GetSafeCellValue(RuntimeHelpers.GetObjectValue(excelSheet), rowIndex, 6);
		safeCellValueAsString = safeCellValueAsString?.Replace("PL", "").Trim() ?? "";
		safeCellValueAsString2 = safeCellValueAsString2?.Replace("PL", "").Trim() ?? "";
		safeCellValueAsString3 = safeCellValueAsString3?.Replace("PL", "").Trim() ?? "";
		List<string> list = new List<string>();
		if (safeCellValue == 680.0)
		{
			if (!string.IsNullOrEmpty(safeCellValueAsString))
			{
				list.Add($"W-{safeCellValueAsString}");
			}
			if (!string.IsNullOrEmpty(safeCellValueAsString2) && !string.IsNullOrEmpty(safeCellValueAsString3))
			{
				list.Add((Operators.CompareString(safeCellValueAsString2, safeCellValueAsString3, TextCompare: false) == 0) ? $"F-{safeCellValueAsString2}" : $"TF-{safeCellValueAsString2} BF-{safeCellValueAsString3}");
			}
			else if (!string.IsNullOrEmpty(safeCellValueAsString2))
			{
				list.Add($"F-{safeCellValueAsString2}");
			}
			else if (!string.IsNullOrEmpty(safeCellValueAsString3))
			{
				list.Add($"F-{safeCellValueAsString3}");
			}
		}
		else
		{
			if (!string.IsNullOrEmpty(safeCellValueAsString))
			{
				list.Add(safeCellValueAsString);
			}
			if (!string.IsNullOrEmpty(safeCellValueAsString2) && !string.IsNullOrEmpty(safeCellValueAsString3))
			{
				list.Add((Operators.CompareString(safeCellValueAsString2, safeCellValueAsString3, TextCompare: false) == 0) ? safeCellValueAsString2 : $"{safeCellValueAsString2}x{safeCellValueAsString3}");
			}
			else if (!string.IsNullOrEmpty(safeCellValueAsString2))
			{
				list.Add(safeCellValueAsString2);
			}
			else if (!string.IsNullOrEmpty(safeCellValueAsString3))
			{
				list.Add(safeCellValueAsString3);
			}
		}
		string text = string.Join(" ", list.Take(2));
		if (string.IsNullOrWhiteSpace(text))
		{
			text = "Beam_" + rowIndex;
		}
		return text.Replace("/", "_").Replace("\\", "_").Replace(":", "_")
			.Replace("*", "_")
			.Replace("?", "_")
			.Replace("\"", "_")
			.Replace("<", "_")
			.Replace(">", "_")
			.Replace("|", "_");
	}
	catch (Exception projectError)
	{
		ProjectData.SetProjectError(projectError);
		string result = "Beam_" + rowIndex;
		ProjectData.ClearProjectError();
		return result;
	}
}


using System;
using netDxf;

private Vector3 RotateVectorAroundAxis(Vector3 vector, Vector3 axis, double angleInDegrees)
{
	if (Math.Abs(angleInDegrees) < 0.001)
	{
		return vector;
	}
	double num = angleInDegrees * Math.PI / 180.0;
	double num2 = Math.Cos(num);
	double num3 = Math.Sin(num);
	Vector3 vector2 = axis;
	vector2.Normalize();
	double num4 = Vector3.DotProduct(vector, vector2);
	Vector3 vector3 = Vector3.CrossProduct(vector2, vector);
	return vector * num2 + vector3 * num3 + vector2 * num4 * (1.0 - num2);
}


using System;
using Microsoft.VisualBasic.CompilerServices;
using netDxf;
using netDxf.Entities;
using netDxf.Tables;

private void DrawTaperedISection(DxfDocument dxf, Vector3 startPoint, Vector3 endPoint, double webHeightStart, double webThicknessStart, double topFlangeWidthStart, double topFlangeThicknessStart, double bottomFlangeWidthStart, double bottomFlangeThicknessStart, double webHeightEnd, double webThicknessEnd, double topFlangeWidthEnd, double topFlangeThicknessEnd, double bottomFlangeWidthEnd, double bottomFlangeThicknessEnd, AciColor color, double betaAngle, string rotationValue, Layer targetLayer)
{
	Vector3 vector = endPoint - startPoint;
	Vector3 vector2 = vector;
	vector2.Normalize();
	double num = 0.99;
	Vector3 vector3;
	Vector3 vector4;
	if (Math.Abs(vector2.Z) > num)
	{
		vector3 = new Vector3(1.0, 0.0, 0.0);
		vector4 = Vector3.CrossProduct(vector2, vector3);
		vector4.Normalize();
		vector3 = Vector3.CrossProduct(vector4, vector2);
		vector3.Normalize();
	}
	else
	{
		vector4 = Vector3.CrossProduct(v: new Vector3(0.0, 0.0, 1.0), u: vector2);
		double num2 = Math.Sqrt(vector4.X * vector4.X + vector4.Y * vector4.Y + vector4.Z * vector4.Z);
		if (num2 < 0.001)
		{
			vector4 = new Vector3(0.0, 1.0, 0.0);
		}
		else
		{
			vector4.Normalize();
		}
		vector3 = Vector3.CrossProduct(vector4, vector2);
		vector3.Normalize();
		vector4 = Vector3.CrossProduct(vector2, vector3);
		vector4.Normalize();
	}
	if (Operators.CompareString(rotationValue.Trim(), "5", TextCompare: false) == 0)
	{
		vector3 = RotateVectorAroundAxis(vector3, vector2, -90.0);
		vector4 = RotateVectorAroundAxis(vector4, vector2, -90.0);
	}
	if (Operators.CompareString(rotationValue.Trim(), "3", TextCompare: false) == 0)
	{
		double num3 = Math.Abs(vector2.X);
		double num4 = Math.Abs(vector2.Y);
		double num5 = Math.Abs(vector2.Z);
		if (num3 > 0.7 || num5 > 0.7)
		{
			vector3 = RotateVectorAroundAxis(vector3, vector2, -90.0);
			vector4 = RotateVectorAroundAxis(vector4, vector2, -90.0);
		}
		vector3 = RotateVectorAroundAxis(vector3, vector2, -90.0);
		vector4 = RotateVectorAroundAxis(vector4, vector2, -90.0);
	}
	if (Operators.CompareString(rotationValue.Trim(), "4", TextCompare: false) == 0)
	{
		double num6 = Math.Abs(vector2.X);
		double num7 = Math.Abs(vector2.Y);
		double num8 = Math.Abs(vector2.Z);
		if (num6 > 0.7 || num8 > 0.7)
		{
			vector3 = RotateVectorAroundAxis(vector3, vector2, -90.0);
			vector4 = RotateVectorAroundAxis(vector4, vector2, -90.0);
		}
		vector3 = RotateVectorAroundAxis(vector3, vector2, -90.0);
		vector4 = RotateVectorAroundAxis(vector4, vector2, -90.0);
	}
	if (Math.Abs(betaAngle) > 0.001)
	{
		vector3 = RotateVectorAroundAxis(vector3, vector2, betaAngle);
		vector4 = RotateVectorAroundAxis(vector4, vector2, betaAngle);
	}
	VB$AnonymousDelegate_0<double, double, Vector3, Vector3> vB$AnonymousDelegate_ = (double offsetY, double offsetZ, Vector3 basePoint) => basePoint + vector3 * offsetY + vector4 * offsetZ;
	Vector3 startPoint2 = vB$AnonymousDelegate_((0.0 - webThicknessStart) / 2.0, webHeightStart / 2.0, startPoint);
	Vector3 startPoint3 = vB$AnonymousDelegate_(webThicknessStart / 2.0, webHeightStart / 2.0, startPoint);
	Vector3 startPoint4 = vB$AnonymousDelegate_((0.0 - webThicknessStart) / 2.0, (0.0 - webHeightStart) / 2.0, startPoint);
	Vector3 startPoint5 = vB$AnonymousDelegate_(webThicknessStart / 2.0, (0.0 - webHeightStart) / 2.0, startPoint);
	Vector3 endPoint2 = vB$AnonymousDelegate_((0.0 - webThicknessEnd) / 2.0, webHeightEnd / 2.0, endPoint);
	Vector3 endPoint3 = vB$AnonymousDelegate_(webThicknessEnd / 2.0, webHeightEnd / 2.0, endPoint);
	Vector3 endPoint4 = vB$AnonymousDelegate_((0.0 - webThicknessEnd) / 2.0, (0.0 - webHeightEnd) / 2.0, endPoint);
	Vector3 endPoint5 = vB$AnonymousDelegate_(webThicknessEnd / 2.0, (0.0 - webHeightEnd) / 2.0, endPoint);
	dxf.Entities.Add(new Line(startPoint2, endPoint2)
	{
		Color = color,
		Layer = targetLayer
	});
	dxf.Entities.Add(new Line(startPoint3, endPoint3)
	{
		Color = color,
		Layer = targetLayer
	});
	dxf.Entities.Add(new Line(startPoint4, endPoint4)
	{
		Color = color,
		Layer = targetLayer
	});
	dxf.Entities.Add(new Line(startPoint5, endPoint5)
	{
		Color = color,
		Layer = targetLayer
	});
	Vector3 startPoint6 = vB$AnonymousDelegate_((0.0 - topFlangeWidthStart) / 2.0, webHeightStart / 2.0 + topFlangeThicknessStart, startPoint);
	Vector3 startPoint7 = vB$AnonymousDelegate_(topFlangeWidthStart / 2.0, webHeightStart / 2.0 + topFlangeThicknessStart, startPoint);
	Vector3 startPoint8 = vB$AnonymousDelegate_((0.0 - topFlangeWidthStart) / 2.0, webHeightStart / 2.0, startPoint);
	Vector3 startPoint9 = vB$AnonymousDelegate_(topFlangeWidthStart / 2.0, webHeightStart / 2.0, startPoint);
	Vector3 endPoint6 = vB$AnonymousDelegate_((0.0 - topFlangeWidthEnd) / 2.0, webHeightEnd / 2.0 + topFlangeThicknessEnd, endPoint);
	Vector3 endPoint7 = vB$AnonymousDelegate_(topFlangeWidthEnd / 2.0, webHeightEnd / 2.0 + topFlangeThicknessEnd, endPoint);
	Vector3 endPoint8 = vB$AnonymousDelegate_((0.0 - topFlangeWidthEnd) / 2.0, webHeightEnd / 2.0, endPoint);
	Vector3 endPoint9 = vB$AnonymousDelegate_(topFlangeWidthEnd / 2.0, webHeightEnd / 2.0, endPoint);
	dxf.Entities.Add(new Line(startPoint6, endPoint6)
	{
		Color = color,
		Layer = targetLayer
	});
	dxf.Entities.Add(new Line(startPoint7, endPoint7)
	{
		Color = color,
		Layer = targetLayer
	});
	dxf.Entities.Add(new Line(startPoint8, endPoint8)
	{
		Color = color,
		Layer = targetLayer
	});
	dxf.Entities.Add(new Line(startPoint9, endPoint9)
	{
		Color = color,
		Layer = targetLayer
	});
	Vector3 startPoint10 = vB$AnonymousDelegate_((0.0 - bottomFlangeWidthStart) / 2.0, (0.0 - webHeightStart) / 2.0 - bottomFlangeThicknessStart, startPoint);
	Vector3 startPoint11 = vB$AnonymousDelegate_(bottomFlangeWidthStart / 2.0, (0.0 - webHeightStart) / 2.0 - bottomFlangeThicknessStart, startPoint);
	Vector3 startPoint12 = vB$AnonymousDelegate_((0.0 - bottomFlangeWidthStart) / 2.0, (0.0 - webHeightStart) / 2.0, startPoint);
	Vector3 startPoint13 = vB$AnonymousDelegate_(bottomFlangeWidthStart / 2.0, (0.0 - webHeightStart) / 2.0, startPoint);
	Vector3 endPoint10 = vB$AnonymousDelegate_((0.0 - bottomFlangeWidthEnd) / 2.0, (0.0 - webHeightEnd) / 2.0 - bottomFlangeThicknessEnd, endPoint);
	Vector3 endPoint11 = vB$AnonymousDelegate_(bottomFlangeWidthEnd / 2.0, (0.0 - webHeightEnd) / 2.0 - bottomFlangeThicknessEnd, endPoint);
	Vector3 endPoint12 = vB$AnonymousDelegate_((0.0 - bottomFlangeWidthEnd) / 2.0, (0.0 - webHeightEnd) / 2.0, endPoint);
	Vector3 endPoint13 = vB$AnonymousDelegate_(bottomFlangeWidthEnd / 2.0, (0.0 - webHeightEnd) / 2.0, endPoint);
	dxf.Entities.Add(new Line(startPoint10, endPoint10)
	{
		Color = color,
		Layer = targetLayer
	});
	dxf.Entities.Add(new Line(startPoint11, endPoint11)
	{
		Color = color,
		Layer = targetLayer
	});
	dxf.Entities.Add(new Line(startPoint12, endPoint12)
	{
		Color = color,
		Layer = targetLayer
	});
	dxf.Entities.Add(new Line(startPoint13, endPoint13)
	{
		Color = color,
		Layer = targetLayer
	});
	DrawIBeamFace(dxf, startPoint, vector3, vector4, webHeightStart, webThicknessStart, topFlangeWidthStart, topFlangeThicknessStart, bottomFlangeWidthStart, bottomFlangeThicknessStart, color, targetLayer);
	DrawIBeamFace(dxf, endPoint, vector3, vector4, webHeightEnd, webThicknessEnd, topFlangeWidthEnd, topFlangeThicknessEnd, bottomFlangeWidthEnd, bottomFlangeThicknessEnd, color, targetLayer);
}


using netDxf;
using netDxf.Entities;
using netDxf.Tables;

private void DrawIBeamFace(DxfDocument dxf, Vector3 centerPoint, Vector3 localY, Vector3 localZ, double webHeight, double webThickness, double topFlangeWidth, double topFlangeThickness, double bottomFlangeWidth, double bottomFlangeThickness, AciColor color, Layer targetLayer)
{
	VB$AnonymousDelegate_1<double, double, Vector3> vB$AnonymousDelegate_ = (double offsetY, double offsetZ) => centerPoint + localY * offsetY + localZ * offsetZ;
	double offsetY2 = (0.0 - webThickness) / 2.0;
	double offsetY3 = webThickness / 2.0;
	double num = webHeight / 2.0;
	double num2 = (0.0 - webHeight) / 2.0;
	double offsetY4 = (0.0 - topFlangeWidth) / 2.0;
	double offsetY5 = topFlangeWidth / 2.0;
	double offsetZ2 = num + topFlangeThickness;
	double offsetZ3 = num;
	double offsetY6 = (0.0 - bottomFlangeWidth) / 2.0;
	double offsetY7 = bottomFlangeWidth / 2.0;
	double offsetZ4 = num2;
	double offsetZ5 = num2 - bottomFlangeThickness;
	dxf.Entities.Add(new Line(vB$AnonymousDelegate_(offsetY4, offsetZ2), vB$AnonymousDelegate_(offsetY5, offsetZ2))
	{
		Color = color,
		Layer = targetLayer
	});
	dxf.Entities.Add(new Line(vB$AnonymousDelegate_(offsetY5, offsetZ2), vB$AnonymousDelegate_(offsetY5, offsetZ3))
	{
		Color = color,
		Layer = targetLayer
	});
	dxf.Entities.Add(new Line(vB$AnonymousDelegate_(offsetY5, offsetZ3), vB$AnonymousDelegate_(offsetY3, offsetZ3))
	{
		Color = color,
		Layer = targetLayer
	});
	dxf.Entities.Add(new Line(vB$AnonymousDelegate_(offsetY3, offsetZ3), vB$AnonymousDelegate_(offsetY3, offsetZ4))
	{
		Color = color,
		Layer = targetLayer
	});
	dxf.Entities.Add(new Line(vB$AnonymousDelegate_(offsetY3, offsetZ4), vB$AnonymousDelegate_(offsetY7, offsetZ4))
	{
		Color = color,
		Layer = targetLayer
	});
	dxf.Entities.Add(new Line(vB$AnonymousDelegate_(offsetY7, offsetZ4), vB$AnonymousDelegate_(offsetY7, offsetZ5))
	{
		Color = color,
		Layer = targetLayer
	});
	dxf.Entities.Add(new Line(vB$AnonymousDelegate_(offsetY7, offsetZ5), vB$AnonymousDelegate_(offsetY6, offsetZ5))
	{
		Color = color,
		Layer = targetLayer
	});
	dxf.Entities.Add(new Line(vB$AnonymousDelegate_(offsetY6, offsetZ5), vB$AnonymousDelegate_(offsetY6, offsetZ4))
	{
		Color = color,
		Layer = targetLayer
	});
	dxf.Entities.Add(new Line(vB$AnonymousDelegate_(offsetY6, offsetZ4), vB$AnonymousDelegate_(offsetY2, offsetZ4))
	{
		Color = color,
		Layer = targetLayer
	});
	dxf.Entities.Add(new Line(vB$AnonymousDelegate_(offsetY2, offsetZ4), vB$AnonymousDelegate_(offsetY2, offsetZ3))
	{
		Color = color,
		Layer = targetLayer
	});
	dxf.Entities.Add(new Line(vB$AnonymousDelegate_(offsetY2, offsetZ3), vB$AnonymousDelegate_(offsetY4, offsetZ3))
	{
		Color = color,
		Layer = targetLayer
	});
	dxf.Entities.Add(new Line(vB$AnonymousDelegate_(offsetY4, offsetZ3), vB$AnonymousDelegate_(offsetY4, offsetZ2))
	{
		Color = color,
		Layer = targetLayer
	});
}


using System;
using Microsoft.VisualBasic.CompilerServices;
using netDxf;
using netDxf.Entities;
using netDxf.Tables;

private void DrawTaperedChannelSection(DxfDocument dxf, Vector3 startPoint, Vector3 endPoint, double webHeightStart, double webThicknessStart, double topFlangeWidthStart, double topFlangeThicknessStart, double bottomFlangeWidthStart, double bottomFlangeThicknessStart, double webHeightEnd, double webThicknessEnd, double topFlangeWidthEnd, double topFlangeThicknessEnd, double bottomFlangeWidthEnd, double bottomFlangeThicknessEnd, AciColor color, double betaAngle, string rotationValue, Layer targetLayer)
{
	Vector3 vector = endPoint - startPoint;
	Vector3 vector2 = vector;
	vector2.Normalize();
	double num = 0.99;
	Vector3 vector3;
	Vector3 vector4;
	if (Math.Abs(vector2.Z) > num)
	{
		vector3 = new Vector3(1.0, 0.0, 0.0);
		vector4 = Vector3.CrossProduct(vector2, vector3);
		vector4.Normalize();
		vector3 = Vector3.CrossProduct(vector4, vector2);
		vector3.Normalize();
	}
	else
	{
		vector4 = Vector3.CrossProduct(v: new Vector3(0.0, 0.0, 1.0), u: vector2);
		double num2 = Math.Sqrt(vector4.X * vector4.X + vector4.Y * vector4.Y + vector4.Z * vector4.Z);
		if (num2 < 0.001)
		{
			vector4 = new Vector3(0.0, 1.0, 0.0);
		}
		else
		{
			vector4.Normalize();
		}
		vector3 = Vector3.CrossProduct(vector4, vector2);
		vector3.Normalize();
		vector4 = Vector3.CrossProduct(vector2, vector3);
		vector4.Normalize();
	}
	if (Operators.CompareString(rotationValue.Trim(), "5", TextCompare: false) == 0)
	{
		vector3 = RotateVectorAroundAxis(vector3, vector2, -90.0);
		vector4 = RotateVectorAroundAxis(vector4, vector2, -90.0);
	}
	if (Operators.CompareString(rotationValue.Trim(), "3", TextCompare: false) == 0)
	{
		double num3 = Math.Abs(vector2.X);
		double num4 = Math.Abs(vector2.Y);
		double num5 = Math.Abs(vector2.Z);
		if (num3 > 0.7 || num5 > 0.7)
		{
			vector3 = RotateVectorAroundAxis(vector3, vector2, -90.0);
			vector4 = RotateVectorAroundAxis(vector4, vector2, -90.0);
		}
		vector3 = RotateVectorAroundAxis(vector3, vector2, -90.0);
		vector4 = RotateVectorAroundAxis(vector4, vector2, -90.0);
	}
	if (Operators.CompareString(rotationValue.Trim(), "4", TextCompare: false) == 0)
	{
		double num6 = Math.Abs(vector2.X);
		double num7 = Math.Abs(vector2.Y);
		double num8 = Math.Abs(vector2.Z);
		if (num6 > 0.7 || num8 > 0.7)
		{
			vector3 = RotateVectorAroundAxis(vector3, vector2, -90.0);
			vector4 = RotateVectorAroundAxis(vector4, vector2, -90.0);
		}
		vector3 = RotateVectorAroundAxis(vector3, vector2, -90.0);
		vector4 = RotateVectorAroundAxis(vector4, vector2, -90.0);
	}
	if (Math.Abs(betaAngle) > 0.001)
	{
		vector3 = RotateVectorAroundAxis(vector3, vector2, betaAngle);
		vector4 = RotateVectorAroundAxis(vector4, vector2, betaAngle);
	}
	double webCenterY2 = webThicknessStart / 2.0;
	double webCenterY3 = webThicknessEnd / 2.0;
	VB$AnonymousDelegate_2<double, double, Vector3, double, Vector3> vB$AnonymousDelegate_ = (double offsetY, double offsetZ, Vector3 basePoint, double webCenterY) => basePoint + vector3 * (offsetY - webCenterY) + vector4 * offsetZ;
	Vector3 startPoint2 = vB$AnonymousDelegate_(0.0, webHeightStart / 2.0, startPoint, webCenterY2);
	Vector3 startPoint3 = vB$AnonymousDelegate_(webThicknessStart, webHeightStart / 2.0, startPoint, webCenterY2);
	Vector3 startPoint4 = vB$AnonymousDelegate_(0.0, (0.0 - webHeightStart) / 2.0, startPoint, webCenterY2);
	Vector3 startPoint5 = vB$AnonymousDelegate_(webThicknessStart, (0.0 - webHeightStart) / 2.0, startPoint, webCenterY2);
	Vector3 endPoint2 = vB$AnonymousDelegate_(0.0, webHeightEnd / 2.0, endPoint, webCenterY3);
	Vector3 endPoint3 = vB$AnonymousDelegate_(webThicknessEnd, webHeightEnd / 2.0, endPoint, webCenterY3);
	Vector3 endPoint4 = vB$AnonymousDelegate_(0.0, (0.0 - webHeightEnd) / 2.0, endPoint, webCenterY3);
	Vector3 endPoint5 = vB$AnonymousDelegate_(webThicknessEnd, (0.0 - webHeightEnd) / 2.0, endPoint, webCenterY3);
	dxf.Entities.Add(new Line(startPoint2, endPoint2)
	{
		Color = color,
		Layer = targetLayer
	});
	dxf.Entities.Add(new Line(startPoint3, endPoint3)
	{
		Color = color,
		Layer = targetLayer
	});
	dxf.Entities.Add(new Line(startPoint4, endPoint4)
	{
		Color = color,
		Layer = targetLayer
	});
	dxf.Entities.Add(new Line(startPoint5, endPoint5)
	{
		Color = color,
		Layer = targetLayer
	});
	Vector3 startPoint6 = vB$AnonymousDelegate_(0.0, webHeightStart / 2.0 + topFlangeThicknessStart, startPoint, webCenterY2);
	Vector3 startPoint7 = vB$AnonymousDelegate_(topFlangeWidthStart, webHeightStart / 2.0 + topFlangeThicknessStart, startPoint, webCenterY2);
	Vector3 startPoint8 = vB$AnonymousDelegate_(0.0, webHeightStart / 2.0, startPoint, webCenterY2);
	Vector3 startPoint9 = vB$AnonymousDelegate_(topFlangeWidthStart, webHeightStart / 2.0, startPoint, webCenterY2);
	Vector3 endPoint6 = vB$AnonymousDelegate_(0.0, webHeightEnd / 2.0 + topFlangeThicknessEnd, endPoint, webCenterY3);
	Vector3 endPoint7 = vB$AnonymousDelegate_(topFlangeWidthEnd, webHeightEnd / 2.0 + topFlangeThicknessEnd, endPoint, webCenterY3);
	Vector3 endPoint8 = vB$AnonymousDelegate_(0.0, webHeightEnd / 2.0, endPoint, webCenterY3);
	Vector3 endPoint9 = vB$AnonymousDelegate_(topFlangeWidthEnd, webHeightEnd / 2.0, endPoint, webCenterY3);
	dxf.Entities.Add(new Line(startPoint6, endPoint6)
	{
		Color = color,
		Layer = targetLayer
	});
	dxf.Entities.Add(new Line(startPoint7, endPoint7)
	{
		Color = color,
		Layer = targetLayer
	});
	dxf.Entities.Add(new Line(startPoint8, endPoint8)
	{
		Color = color,
		Layer = targetLayer
	});
	dxf.Entities.Add(new Line(startPoint9, endPoint9)
	{
		Color = color,
		Layer = targetLayer
	});
	Vector3 startPoint10 = vB$AnonymousDelegate_(0.0, (0.0 - webHeightStart) / 2.0 - bottomFlangeThicknessStart, startPoint, webCenterY2);
	Vector3 startPoint11 = vB$AnonymousDelegate_(bottomFlangeWidthStart, (0.0 - webHeightStart) / 2.0 - bottomFlangeThicknessStart, startPoint, webCenterY2);
	Vector3 startPoint12 = vB$AnonymousDelegate_(0.0, (0.0 - webHeightStart) / 2.0, startPoint, webCenterY2);
	Vector3 startPoint13 = vB$AnonymousDelegate_(bottomFlangeWidthStart, (0.0 - webHeightStart) / 2.0, startPoint, webCenterY2);
	Vector3 endPoint10 = vB$AnonymousDelegate_(0.0, (0.0 - webHeightEnd) / 2.0 - bottomFlangeThicknessEnd, endPoint, webCenterY3);
	Vector3 endPoint11 = vB$AnonymousDelegate_(bottomFlangeWidthEnd, (0.0 - webHeightEnd) / 2.0 - bottomFlangeThicknessEnd, endPoint, webCenterY3);
	Vector3 endPoint12 = vB$AnonymousDelegate_(0.0, (0.0 - webHeightEnd) / 2.0, endPoint, webCenterY3);
	Vector3 endPoint13 = vB$AnonymousDelegate_(bottomFlangeWidthEnd, (0.0 - webHeightEnd) / 2.0, endPoint, webCenterY3);
	dxf.Entities.Add(new Line(startPoint10, endPoint10)
	{
		Color = color,
		Layer = targetLayer
	});
	dxf.Entities.Add(new Line(startPoint11, endPoint11)
	{
		Color = color,
		Layer = targetLayer
	});
	dxf.Entities.Add(new Line(startPoint12, endPoint12)
	{
		Color = color,
		Layer = targetLayer
	});
	dxf.Entities.Add(new Line(startPoint13, endPoint13)
	{
		Color = color,
		Layer = targetLayer
	});
	DrawChannelFace(dxf, startPoint, vector3, vector4, webHeightStart, webThicknessStart, topFlangeWidthStart, topFlangeThicknessStart, bottomFlangeWidthStart, bottomFlangeThicknessStart, color, targetLayer, webCenterY2);
	DrawChannelFace(dxf, endPoint, vector3, vector4, webHeightEnd, webThicknessEnd, topFlangeWidthEnd, topFlangeThicknessEnd, bottomFlangeWidthEnd, bottomFlangeThicknessEnd, color, targetLayer, webCenterY3);
}


using netDxf;
using netDxf.Entities;
using netDxf.Tables;

private void DrawChannelFace(DxfDocument dxf, Vector3 centerPoint, Vector3 localY, Vector3 localZ, double webHeight, double webThickness, double topFlangeWidth, double topFlangeThickness, double bottomFlangeWidth, double bottomFlangeThickness, AciColor color, Layer targetLayer, double webCenterY)
{
	VB$AnonymousDelegate_1<double, double, Vector3> vB$AnonymousDelegate_ = (double offsetY, double offsetZ) => centerPoint + localY * (offsetY - webCenterY) + localZ * offsetZ;
	int num = 0;
	double num2 = webHeight / 2.0;
	double num3 = (0.0 - webHeight) / 2.0;
	int num4 = 0;
	double offsetZ2 = num2 + topFlangeThickness;
	double offsetZ3 = num2;
	int num5 = 0;
	double offsetZ4 = num3;
	double offsetZ5 = num3 - bottomFlangeThickness;
	dxf.Entities.Add(new Line(vB$AnonymousDelegate_(num4, offsetZ2), vB$AnonymousDelegate_(topFlangeWidth, offsetZ2))
	{
		Color = color,
		Layer = targetLayer
	});
	dxf.Entities.Add(new Line(vB$AnonymousDelegate_(topFlangeWidth, offsetZ2), vB$AnonymousDelegate_(topFlangeWidth, offsetZ3))
	{
		Color = color,
		Layer = targetLayer
	});
	dxf.Entities.Add(new Line(vB$AnonymousDelegate_(topFlangeWidth, offsetZ3), vB$AnonymousDelegate_(webThickness, offsetZ3))
	{
		Color = color,
		Layer = targetLayer
	});
	dxf.Entities.Add(new Line(vB$AnonymousDelegate_(webThickness, offsetZ3), vB$AnonymousDelegate_(webThickness, offsetZ4))
	{
		Color = color,
		Layer = targetLayer
	});
	dxf.Entities.Add(new Line(vB$AnonymousDelegate_(webThickness, offsetZ4), vB$AnonymousDelegate_(bottomFlangeWidth, offsetZ4))
	{
		Color = color,
		Layer = targetLayer
	});
	dxf.Entities.Add(new Line(vB$AnonymousDelegate_(bottomFlangeWidth, offsetZ4), vB$AnonymousDelegate_(bottomFlangeWidth, offsetZ5))
	{
		Color = color,
		Layer = targetLayer
	});
	dxf.Entities.Add(new Line(vB$AnonymousDelegate_(bottomFlangeWidth, offsetZ5), vB$AnonymousDelegate_(num5, offsetZ5))
	{
		Color = color,
		Layer = targetLayer
	});
	dxf.Entities.Add(new Line(vB$AnonymousDelegate_(num5, offsetZ5), vB$AnonymousDelegate_(num5, offsetZ4))
	{
		Color = color,
		Layer = targetLayer
	});
	dxf.Entities.Add(new Line(vB$AnonymousDelegate_(num5, offsetZ4), vB$AnonymousDelegate_(num, offsetZ4))
	{
		Color = color,
		Layer = targetLayer
	});
	dxf.Entities.Add(new Line(vB$AnonymousDelegate_(num, offsetZ4), vB$AnonymousDelegate_(num, offsetZ3))
	{
		Color = color,
		Layer = targetLayer
	});
	dxf.Entities.Add(new Line(vB$AnonymousDelegate_(num, offsetZ3), vB$AnonymousDelegate_(num4, offsetZ3))
	{
		Color = color,
		Layer = targetLayer
	});
	dxf.Entities.Add(new Line(vB$AnonymousDelegate_(num4, offsetZ3), vB$AnonymousDelegate_(num4, offsetZ2))
	{
		Color = color,
		Layer = targetLayer
	});
}


using System;
using Microsoft.VisualBasic.CompilerServices;
using netDxf;
using netDxf.Tables;

private void DrawHollowRectangleTube(DxfDocument dxf, Vector3 startPoint, Vector3 endPoint, double width, double height, double thickness, AciColor color, double betaAngle, string rotationValue, Layer targetLayer)
{
	Vector3 vector = endPoint - startPoint;
	Vector3 vector2 = vector;
	vector2.Normalize();
	double num = 0.99;
	Vector3 vector4;
	Vector3 vector3;
	if (Math.Abs(vector2.Z) > num)
	{
		vector3 = Vector3.CrossProduct(v: new Vector3(1.0, 0.0, 0.0), u: vector2);
		vector3.Normalize();
		vector4 = Vector3.CrossProduct(vector3, vector2);
		vector4.Normalize();
	}
	else
	{
		vector3 = Vector3.CrossProduct(v: new Vector3(0.0, 0.0, 1.0), u: vector2);
		double num2 = Math.Sqrt(vector3.X * vector3.X + vector3.Y * vector3.Y + vector3.Z * vector3.Z);
		if (num2 < 0.001)
		{
			vector3 = new Vector3(0.0, 1.0, 0.0);
		}
		else
		{
			vector3.Normalize();
		}
		vector4 = Vector3.CrossProduct(vector3, vector2);
		vector4.Normalize();
		vector3 = Vector3.CrossProduct(vector2, vector4);
		vector3.Normalize();
	}
	if (Operators.CompareString(rotationValue.Trim(), "5", TextCompare: false) == 0)
	{
		vector4 = RotateVectorAroundAxis(vector4, vector2, -90.0);
		vector3 = RotateVectorAroundAxis(vector3, vector2, -90.0);
	}
	if (Operators.CompareString(rotationValue.Trim(), "3", TextCompare: false) == 0)
	{
		double num3 = Math.Abs(vector2.X);
		double num4 = Math.Abs(vector2.Y);
		double num5 = Math.Abs(vector2.Z);
		if (num3 > 0.7 || num5 > 0.7)
		{
			vector4 = RotateVectorAroundAxis(vector4, vector2, -90.0);
			vector3 = RotateVectorAroundAxis(vector3, vector2, -90.0);
		}
		vector4 = RotateVectorAroundAxis(vector4, vector2, -90.0);
		vector3 = RotateVectorAroundAxis(vector3, vector2, -90.0);
	}
	if (Operators.CompareString(rotationValue.Trim(), "4", TextCompare: false) == 0)
	{
		double num6 = Math.Abs(vector2.X);
		double num7 = Math.Abs(vector2.Y);
		double num8 = Math.Abs(vector2.Z);
		if (num6 > 0.7 || num8 > 0.7)
		{
			vector4 = RotateVectorAroundAxis(vector4, vector2, -90.0);
			vector3 = RotateVectorAroundAxis(vector3, vector2, -90.0);
		}
		vector4 = RotateVectorAroundAxis(vector4, vector2, -90.0);
		vector3 = RotateVectorAroundAxis(vector3, vector2, -90.0);
	}
	if (Math.Abs(betaAngle) > 0.001)
	{
		vector4 = RotateVectorAroundAxis(vector4, vector2, betaAngle);
		vector3 = RotateVectorAroundAxis(vector3, vector2, betaAngle);
	}
	double num9 = width / 2.0;
	double num10 = height / 2.0;
	double num11 = num9 - thickness;
	double num12 = num10 - thickness;
	DrawRectangularTube3D(dxf, startPoint, endPoint, vector4, vector3, num9, num10, color, targetLayer);
	if (num11 > 0.0 && num12 > 0.0)
	{
		DrawRectangularTube3D(dxf, startPoint, endPoint, vector4, vector3, num11, num12, color, targetLayer);
	}
}


using netDxf;
using netDxf.Tables;

private void DrawRectangularTube3D(DxfDocument dxf, Vector3 startPoint, Vector3 endPoint, Vector3 localY, Vector3 localZ, double halfWidth, double halfHeight, AciColor color, Layer targetLayer)
{
	VB$AnonymousDelegate_0<double, double, Vector3, Vector3> vB$AnonymousDelegate_ = (double offsetY, double offsetZ, Vector3 basePoint) => basePoint + localY * offsetY + localZ * offsetZ;
	Vector3 vector = vB$AnonymousDelegate_(halfHeight, halfWidth, startPoint);
	Vector3 vector2 = vB$AnonymousDelegate_(0.0 - halfHeight, halfWidth, startPoint);
	Vector3 vector3 = vB$AnonymousDelegate_(0.0 - halfHeight, 0.0 - halfWidth, startPoint);
	Vector3 vector4 = vB$AnonymousDelegate_(halfHeight, 0.0 - halfWidth, startPoint);
	Vector3 vector5 = vB$AnonymousDelegate_(halfHeight, halfWidth, endPoint);
	Vector3 vector6 = vB$AnonymousDelegate_(0.0 - halfHeight, halfWidth, endPoint);
	Vector3 vector7 = vB$AnonymousDelegate_(0.0 - halfHeight, 0.0 - halfWidth, endPoint);
	Vector3 vector8 = vB$AnonymousDelegate_(halfHeight, 0.0 - halfWidth, endPoint);
	DrawLine3D(dxf, vector, vector5, color, targetLayer);
	DrawLine3D(dxf, vector2, vector6, color, targetLayer);
	DrawLine3D(dxf, vector3, vector7, color, targetLayer);
	DrawLine3D(dxf, vector4, vector8, color, targetLayer);
	DrawLine3D(dxf, vector, vector2, color, targetLayer);
	DrawLine3D(dxf, vector2, vector3, color, targetLayer);
	DrawLine3D(dxf, vector3, vector4, color, targetLayer);
	DrawLine3D(dxf, vector4, vector, color, targetLayer);
	DrawLine3D(dxf, vector5, vector6, color, targetLayer);
	DrawLine3D(dxf, vector6, vector7, color, targetLayer);
	DrawLine3D(dxf, vector7, vector8, color, targetLayer);
	DrawLine3D(dxf, vector8, vector5, color, targetLayer);
}


using System;
using Microsoft.VisualBasic.CompilerServices;
using netDxf;
using netDxf.Tables;

private void DrawTubeSection(DxfDocument dxf, Vector3 startPoint, Vector3 endPoint, double width, double thickness, AciColor color, double betaAngle, string rotationValue, Layer targetLayer)
{
	Vector3 vector = endPoint - startPoint;
	Vector3 vector2 = vector;
	vector2.Normalize();
	double num = 0.99;
	Vector3 vector4;
	Vector3 vector3;
	if (Math.Abs(vector2.Z) > num)
	{
		vector3 = Vector3.CrossProduct(v: new Vector3(1.0, 0.0, 0.0), u: vector2);
		vector3.Normalize();
		vector4 = Vector3.CrossProduct(vector3, vector2);
		vector4.Normalize();
	}
	else
	{
		vector3 = Vector3.CrossProduct(v: new Vector3(0.0, 0.0, 1.0), u: vector2);
		double num2 = Math.Sqrt(vector3.X * vector3.X + vector3.Y * vector3.Y + vector3.Z * vector3.Z);
		if (num2 < 0.001)
		{
			vector3 = new Vector3(0.0, 1.0, 0.0);
		}
		else
		{
			vector3.Normalize();
		}
		vector4 = Vector3.CrossProduct(vector3, vector2);
		vector4.Normalize();
		vector3 = Vector3.CrossProduct(vector2, vector4);
		vector3.Normalize();
	}
	if (Operators.CompareString(rotationValue.Trim(), "5", TextCompare: false) == 0)
	{
		vector4 = RotateVectorAroundAxis(vector4, vector2, -90.0);
		vector3 = RotateVectorAroundAxis(vector3, vector2, -90.0);
	}
	if (Operators.CompareString(rotationValue.Trim(), "3", TextCompare: false) == 0)
	{
		double num3 = Math.Abs(vector2.X);
		double num4 = Math.Abs(vector2.Y);
		double num5 = Math.Abs(vector2.Z);
		if (num3 > 0.7 || num5 > 0.7)
		{
			vector4 = RotateVectorAroundAxis(vector4, vector2, -90.0);
			vector3 = RotateVectorAroundAxis(vector3, vector2, -90.0);
		}
		vector4 = RotateVectorAroundAxis(vector4, vector2, -90.0);
		vector3 = RotateVectorAroundAxis(vector3, vector2, -90.0);
	}
	if (Operators.CompareString(rotationValue.Trim(), "4", TextCompare: false) == 0)
	{
		double num6 = Math.Abs(vector2.X);
		double num7 = Math.Abs(vector2.Y);
		double num8 = Math.Abs(vector2.Z);
		if (num6 > 0.7 || num8 > 0.7)
		{
			vector4 = RotateVectorAroundAxis(vector4, vector2, -90.0);
			vector3 = RotateVectorAroundAxis(vector3, vector2, -90.0);
		}
		vector4 = RotateVectorAroundAxis(vector4, vector2, -90.0);
		vector3 = RotateVectorAroundAxis(vector3, vector2, -90.0);
	}
	if (Math.Abs(betaAngle) > 0.001)
	{
		vector4 = RotateVectorAroundAxis(vector4, vector2, betaAngle);
		vector3 = RotateVectorAroundAxis(vector3, vector2, betaAngle);
	}
	double num9 = width / 2.0;
	double num10 = num9 - thickness;
	DrawSquareTube3D(dxf, startPoint, endPoint, vector4, vector3, num9, color, targetLayer);
	if (num10 > 0.0)
	{
		DrawSquareTube3D(dxf, startPoint, endPoint, vector4, vector3, num10, color, targetLayer);
	}
}


using netDxf;
using netDxf.Tables;

private void DrawSquareTube3D(DxfDocument dxf, Vector3 startPoint, Vector3 endPoint, Vector3 localY, Vector3 localZ, double halfWidth, AciColor color, Layer targetLayer)
{
	VB$AnonymousDelegate_0<double, double, Vector3, Vector3> vB$AnonymousDelegate_ = (double offsetY, double offsetZ, Vector3 basePoint) => basePoint + localY * offsetY + localZ * offsetZ;
	Vector3 vector = vB$AnonymousDelegate_(halfWidth, halfWidth, startPoint);
	Vector3 vector2 = vB$AnonymousDelegate_(0.0 - halfWidth, halfWidth, startPoint);
	Vector3 vector3 = vB$AnonymousDelegate_(0.0 - halfWidth, 0.0 - halfWidth, startPoint);
	Vector3 vector4 = vB$AnonymousDelegate_(halfWidth, 0.0 - halfWidth, startPoint);
	Vector3 vector5 = vB$AnonymousDelegate_(halfWidth, halfWidth, endPoint);
	Vector3 vector6 = vB$AnonymousDelegate_(0.0 - halfWidth, halfWidth, endPoint);
	Vector3 vector7 = vB$AnonymousDelegate_(0.0 - halfWidth, 0.0 - halfWidth, endPoint);
	Vector3 vector8 = vB$AnonymousDelegate_(halfWidth, 0.0 - halfWidth, endPoint);
	DrawLine3D(dxf, vector, vector5, color, targetLayer);
	DrawLine3D(dxf, vector2, vector6, color, targetLayer);
	DrawLine3D(dxf, vector3, vector7, color, targetLayer);
	DrawLine3D(dxf, vector4, vector8, color, targetLayer);
	DrawLine3D(dxf, vector, vector2, color, targetLayer);
	DrawLine3D(dxf, vector2, vector3, color, targetLayer);
	DrawLine3D(dxf, vector3, vector4, color, targetLayer);
	DrawLine3D(dxf, vector4, vector, color, targetLayer);
	DrawLine3D(dxf, vector5, vector6, color, targetLayer);
	DrawLine3D(dxf, vector6, vector7, color, targetLayer);
	DrawLine3D(dxf, vector7, vector8, color, targetLayer);
	DrawLine3D(dxf, vector8, vector5, color, targetLayer);
}


using netDxf;
using netDxf.Tables;

private void DrawPipeSection(DxfDocument dxf, Vector3 startPoint, Vector3 endPoint, double outerDiameter, double thickness, AciColor color, double betaAngle, string rotationValue, Layer targetLayer)
{
	if (!(outerDiameter <= 0.0) && !(thickness <= 0.0))
	{
		double num = outerDiameter - 2.0 * thickness;
		if (num <= 0.0)
		{
			DrawSolidPipeSection(dxf, startPoint, endPoint, outerDiameter, color, betaAngle, rotationValue, targetLayer);
			return;
		}
		DrawSolidPipeSection(dxf, startPoint, endPoint, outerDiameter, color, betaAngle, rotationValue, targetLayer);
		DrawSolidPipeSection(dxf, startPoint, endPoint, num, color, betaAngle, rotationValue, targetLayer);
	}
}


using System;
using netDxf;
using netDxf.Entities;
using netDxf.Tables;

private void DrawSolidPipeSection(DxfDocument dxf, Vector3 startPoint, Vector3 endPoint, double diameter, AciColor color, double betaAngle, string rotationValue, Layer targetLayer)
{
	Vector3 vector = endPoint - startPoint;
	Vector3 vector2 = vector;
	vector2.Normalize();
	Vector3 vector3 = Vector3.CrossProduct(v: new Vector3(0.0, 0.0, 1.0), u: vector2);
	if (vector3.Modulus() < 0.001)
	{
		vector3 = new Vector3(0.0, 1.0, 0.0);
	}
	else
	{
		vector3.Normalize();
	}
	Vector3 vector4 = Vector3.CrossProduct(vector3, vector2);
	vector4.Normalize();
	if (Math.Abs(betaAngle) > 0.001)
	{
		vector4 = RotateVectorAroundAxis(vector4, vector2, betaAngle);
		vector3 = RotateVectorAroundAxis(vector3, vector2, betaAngle);
	}
	int num = 12;
	double num2 = diameter / 2.0;
	double num3 = Math.PI * 2.0 / (double)num;
	DrawCircle3D(dxf, startPoint, vector4, vector3, num2, num, color, targetLayer);
	DrawCircle3D(dxf, endPoint, vector4, vector3, num2, num, color, targetLayer);
	checked
	{
		int num4 = num - 1;
		for (int i = 0; i <= num4; i++)
		{
			double num5 = (double)i * num3;
			double num6 = num2 * Math.Cos(num5);
			double num7 = num2 * Math.Sin(num5);
			Vector3 startPoint2 = startPoint + vector4 * num6 + vector3 * num7;
			Vector3 endPoint2 = endPoint + vector4 * num6 + vector3 * num7;
			dxf.Entities.Add(new Line(startPoint2, endPoint2)
			{
				Color = color,
				Layer = targetLayer
			});
		}
	}
}


using System;
using netDxf;
using netDxf.Entities;
using netDxf.Tables;

private void DrawCircle3D(DxfDocument dxf, Vector3 center, Vector3 axisY, Vector3 axisZ, double radius, int segments, AciColor color, Layer targetLayer)
{
	double num = Math.PI * 2.0 / (double)segments;
	bool flag = false;
	Vector3 startPoint = default(Vector3);
	Vector3 endPoint = default(Vector3);
	for (int i = 0; i <= segments; i = checked(i + 1))
	{
		double num2 = (double)i * num;
		double num3 = radius * Math.Cos(num2);
		double num4 = radius * Math.Sin(num2);
		Vector3 vector = center + axisY * num3 + axisZ * num4;
		if (i > 0)
		{
			Line line = new Line(startPoint, vector);
			line.Color = color;
			line.Layer = targetLayer;
			dxf.Entities.Add(line);
		}
		else
		{
			endPoint = vector;
			flag = true;
		}
		startPoint = vector;
	}
	if (flag)
	{
		Line line2 = new Line(startPoint, endPoint);
		line2.Color = color;
		line2.Layer = targetLayer;
		dxf.Entities.Add(line2);
	}
}


using netDxf;
using netDxf.Entities;
using netDxf.Tables;

private void DrawLine3D(DxfDocument dxf, Vector3 p1, Vector3 p2, AciColor color, Layer layer)
{
	Line entity = new Line(p1, p2)
	{
		Color = color,
		Layer = layer
	};
	dxf.Entities.Add(entity);
}
