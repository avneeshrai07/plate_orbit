#define DEBUG
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace RaghavStaadExtractor;

[ComVisible(true)]
[Guid("F9396D1F-8F3F-4A8B-9E5D-3B7D8C1A4B2E")]
[ProgId("RaghavStaadExtractor.StaadDataExtractor")]
public class StaadDataExtractor
{
	public struct BeamData
	{
		public int BeamNo;

		public double Length;

		public string SectionName;

		public int NodeA;

		public int NodeB;

		public double XA;

		public double YA;

		public double ZA;

		public double XB;

		public double YB;

		public double ZB;

		public double F1;

		public double F2;

		public double F3;

		public double F4;

		public double F5;

		public double F6;

		public double F7;

		public double Beta;
	}

	private struct Sheet4GridData
	{
		public List<string> GridNames;

		public List<double> AllCoordinates;

		public string SpacingString;
	}

	private const int PROP_DEPTH_START = 0;

	private const int PROP_WEB_THICKNESS = 1;

	private const int PROP_DEPTH_END = 2;

	private const int PROP_FLANGE_WIDTH = 3;

	private const int PROP_FLANGE_THICKNESS = 4;

	private const int PROP_BOTTOM_FLANGE_WIDTH = 5;

	private const int PROP_BOTTOM_FLANGE_THICKNESS = 6;

	private string _lastFrameAnalysisResult;

	private string _lastLogFilePath;

	public StaadDataExtractor()
	{
		_lastFrameAnalysisResult = "";
		_lastLogFilePath = "";
		EnsureLicense();
		CheckExpirationDateAndBackdate();
	}

	[ComVisible(true)]
	public string ExtractBeamDataFromStaad()
	{
		object obj = null;
		object obj2 = null;
		string result;
		try
		{
			obj = RuntimeHelpers.GetObjectValue(GetOrCreateExcelApp());
			obj2 = RuntimeHelpers.GetObjectValue(GetStaadApplication());
			if (obj2 == null)
			{
				result = "STAAD is not running or OpenSTAAD not available.";
			}
			else
			{
				object obj3 = RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(obj, null, "ActiveWorkbook", new object[0], null, null, null));
				if (obj3 == null)
				{
					obj3 = ((!Operators.ConditionalCompareObjectGreater(NewLateBinding.LateGet(NewLateBinding.LateGet(obj, null, "Workbooks", new object[0], null, null, null), null, "Count", new object[0], null, null, null), 0, TextCompare: false)) ? RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(NewLateBinding.LateGet(obj, null, "Workbooks", new object[0], null, null, null), null, "Add", new object[0], null, null, null)) : RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(obj, null, "Workbooks", new object[1] { 1 }, null, null, null)));
				}
				object objectValue = RuntimeHelpers.GetObjectValue(GetOrCreateWorksheet(RuntimeHelpers.GetObjectValue(obj3), "Sheet1"));
				NewLateBinding.LateSet(obj, null, "ScreenUpdating", new object[1] { false }, null, null);
				NewLateBinding.LateSet(obj, null, "Calculation", new object[1] { -4135 }, null, null);
				NewLateBinding.LateSet(obj, null, "EnableEvents", new object[1] { false }, null, null);
				NewLateBinding.LateSet(obj, null, "DisplayAlerts", new object[1] { false }, null, null);
				UnprotectSheetAndWorkbook(RuntimeHelpers.GetObjectValue(objectValue));
				ShowHiddenRows(RuntimeHelpers.GetObjectValue(objectValue));
				PreserveAndClearWorksheet(RuntimeHelpers.GetObjectValue(objectValue));
				WriteJobInfoToExcel(RuntimeHelpers.GetObjectValue(objectValue), RuntimeHelpers.GetObjectValue(obj2));
				SetupExcelHeaders(RuntimeHelpers.GetObjectValue(objectValue));
				string text = ProcessBeamData(RuntimeHelpers.GetObjectValue(objectValue), RuntimeHelpers.GetObjectValue(obj2));
				FormatExcelWorksheet(RuntimeHelpers.GetObjectValue(objectValue));
				ProcessSectionFarmingAndWeight(RuntimeHelpers.GetObjectValue(objectValue));
				CalculateTotalWeight(RuntimeHelpers.GetObjectValue(objectValue));
				_lastFrameAnalysisResult = ProcessFrameAnalysis(RuntimeHelpers.GetObjectValue(objectValue), RuntimeHelpers.GetObjectValue(obj2));
				HideColumns(RuntimeHelpers.GetObjectValue(objectValue));
				ProtectSheetAndWorkbook(RuntimeHelpers.GetObjectValue(objectValue));
				NewLateBinding.LateSet(obj, null, "ScreenUpdating", new object[1] { true }, null, null);
				NewLateBinding.LateSet(obj, null, "Calculation", new object[1] { -4105 }, null, null);
				NewLateBinding.LateSet(obj, null, "EnableEvents", new object[1] { true }, null, null);
				NewLateBinding.LateSet(obj, null, "DisplayAlerts", new object[1] { true }, null, null);
				NewLateBinding.LateCall(NewLateBinding.LateGet(obj3, null, "Sheets", new object[1] { "Sheet1" }, null, null, null), null, "Select", new object[0], null, null, null, IgnoreReturn: true);
				object objectValue2 = RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(obj, null, "ActiveWindow", new object[0], null, null, null));
				NewLateBinding.LateSet(objectValue2, null, "DisplayGridlines", new object[1] { false }, null, null);
				NewLateBinding.LateSet(objectValue2, null, "DisplayHeadings", new object[1] { false }, null, null);
				NewLateBinding.LateSet(obj, null, "DisplayFormulaBar", new object[1] { true }, null, null);
				NewLateBinding.LateSet(objectValue2, null, "DisplayWorkbookTabs", new object[1] { false }, null, null);
				NewLateBinding.LateCall(obj, null, "ExecuteExcel4Macro", new object[1] { "SHOW.TOOLBAR(\"Ribbon\", False)" }, null, null, null, IgnoreReturn: true);
				NewLateBinding.LateSet(objectValue2, null, "ScrollRow", new object[1] { 1 }, null, null);
				NewLateBinding.LateSet(objectValue2, null, "ScrollColumn", new object[1] { 1 }, null, null);
				NewLateBinding.LateSet(objectValue2, null, "Zoom", new object[1] { 100 }, null, null);
				int num = Conversions.ToInteger(NewLateBinding.LateGet(NewLateBinding.LateGet(obj3, null, "Windows", new object[0], null, null, null), null, "Count", new object[0], null, null, null));
				for (int i = 1; i <= num; i = checked(i + 1))
				{
					object[] array;
					bool[] array2;
					object obj4 = NewLateBinding.LateGet(obj3, null, "Windows", array = new object[1] { i }, null, null, array2 = new bool[1] { true });
					if (array2[0])
					{
						i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
					}
					object objectValue3 = RuntimeHelpers.GetObjectValue(obj4);
					if (objectValue3 != null)
					{
						NewLateBinding.LateSet(objectValue3, null, "DisplayGridlines", new object[1] { false }, null, null);
					}
				}
				AutoSaveWorkbook(RuntimeHelpers.GetObjectValue(obj3));
				NewLateBinding.LateSetComplex(NewLateBinding.LateGet(objectValue, null, "Range", new object[1] { "AW2" }, null, null, null), null, "Value", new object[1] { "" }, null, null, OptimisticSet: false, RValueBase: true);
				result = text;
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			result = "Error: " + ex2.Message;
			ProjectData.ClearProjectError();
		}
		finally
		{
			if (obj2 != null)
			{
				Marshal.ReleaseComObject(RuntimeHelpers.GetObjectValue(obj2));
			}
			if (obj != null)
			{
				Marshal.ReleaseComObject(RuntimeHelpers.GetObjectValue(obj));
			}
		}
		return result;
	}

	[ComVisible(true)]
	public string GetFrameAnalysisResult()
	{
		return _lastFrameAnalysisResult;
	}

	[ComVisible(true)]
	public string GetCombinedResults()
	{
		string text = ExtractBeamDataFromStaad();
		string lastFrameAnalysisResult = _lastFrameAnalysisResult;
		if (string.IsNullOrEmpty(lastFrameAnalysisResult))
		{
			return text;
		}
		return text + "\r\n\r\n=== FRAME ANALYSIS RESULTS ===\r\n" + lastFrameAnalysisResult;
	}

	[ComVisible(true)]
	public bool HasFrameAnalysisResult()
	{
		return !string.IsNullOrEmpty(_lastFrameAnalysisResult);
	}

	private void ShowHiddenRows(object ws)
	{
		try
		{
			NewLateBinding.LateSetComplex(NewLateBinding.LateGet(ws, null, "Rows", new object[0], null, null, null), null, "Hidden", new object[1] { false }, null, null, OptimisticSet: false, RValueBase: true);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			Debug.WriteLine("ClearAllFilters failed: " + ex2.Message);
			ProjectData.ClearProjectError();
		}
	}

	[ComVisible(true)]
	public void ShowFrameAnalysisDialog()
	{
		if (!string.IsNullOrEmpty(_lastFrameAnalysisResult))
		{
			ShowCustomAnalysisDialog(_lastFrameAnalysisResult, _lastLogFilePath);
		}
	}

	[ComVisible(true)]
	public bool HasLogFile()
	{
		return !string.IsNullOrEmpty(_lastLogFilePath) && File.Exists(_lastLogFilePath);
	}

	private void ShowCustomAnalysisDialog(string message, string logFilePath)
	{
		try
		{
			Form form = new Form();
			form.Text = "Nova Frame Analysis Results";
			form.Size = new Size(600, 450);
			form.StartPosition = FormStartPosition.CenterScreen;
			form.FormBorderStyle = FormBorderStyle.FixedDialog;
			form.MaximizeBox = false;
			form.MinimizeBox = false;
			TextBox textBox = new TextBox();
			textBox.Multiline = true;
			textBox.ScrollBars = ScrollBars.Vertical;
			textBox.ReadOnly = true;
			textBox.Text = message;
			textBox.Font = new Font("Consolas", 9f);
			textBox.Location = new Point(10, 10);
			textBox.Size = new Size(565, 350);
			Panel panel = new Panel();
			panel.Location = new Point(10, 370);
			panel.Size = new Size(565, 40);
			Button button = new Button();
			button.Text = "OK";
			button.Size = new Size(80, 30);
			button.Location = new Point(475, 5);
			button.DialogResult = DialogResult.OK;
			if (!string.IsNullOrEmpty(logFilePath) && File.Exists(logFilePath))
			{
				Button button2 = new Button();
				button2.Text = "Open Log File";
				button2.Size = new Size(120, 30);
				button2.Location = new Point(340, 5);
				button2.Click += [SpecialName] (object sender, EventArgs e) =>
				{
					try
					{
						Process.Start("explorer.exe", "/select,\"" + logFilePath + "\"");
					}
					catch (Exception ex3)
					{
						ProjectData.SetProjectError(ex3);
						Exception ex4 = ex3;
						try
						{
							Process.Start("explorer.exe", Path.GetDirectoryName(logFilePath));
						}
						catch (Exception projectError)
						{
							ProjectData.SetProjectError(projectError);
							Interaction.MsgBox("Could not open log file location.", MsgBoxStyle.Exclamation);
							ProjectData.ClearProjectError();
						}
						ProjectData.ClearProjectError();
					}
				};
				panel.Controls.Add(button2);
			}
			panel.Controls.Add(button);
			form.Controls.Add(textBox);
			form.Controls.Add(panel);
			form.AcceptButton = button;
			form.ShowDialog();
			form.Dispose();
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			Interaction.MsgBox(message, MsgBoxStyle.Information, "Nova Frame Analysis Results");
			ProjectData.ClearProjectError();
		}
	}

	private string ProcessFrameAnalysis(object ws, object staadApp)
	{
		checked
		{
			string result;
			try
			{
				_lastLogFilePath = "";
				string text = "";
				if (NewLateBinding.LateGet(NewLateBinding.LateGet(ws, null, "Range", new object[1] { "AW2" }, null, null, null), null, "Value", new object[0], null, null, null) != null)
				{
					text = NewLateBinding.LateGet(NewLateBinding.LateGet(ws, null, "Range", new object[1] { "AW2" }, null, null, null), null, "Value", new object[0], null, null, null).ToString().Trim()
						.ToUpper();
				}
				if (string.IsNullOrEmpty(text))
				{
					result = "";
				}
				else if ((Operators.CompareString(text, "X", TextCompare: false) != 0) & (Operators.CompareString(text, "Z", TextCompare: false) != 0) & (Operators.CompareString(text, "G", TextCompare: false) != 0))
				{
					result = $"Invalid prompt '{text}' in AW2. Please enter 'X', 'Z', or 'G' only.";
				}
				else if (Operators.CompareString(text, "G", TextCompare: false) == 0)
				{
					result = ProcessBothDirectionsForGridSilent(RuntimeHelpers.GetObjectValue(ws), RuntimeHelpers.GetObjectValue(staadApp));
				}
				else if (!AreGridsDefined(text, RuntimeHelpers.GetObjectValue(ws)))
				{
					result = string.Format("FOR FRAME ANALYSIS DEFINE GRIDS FIRST USING 'G' OPTION{0}{1}Grid data for '{2}' direction is not found Nova Grid System.{3}{4}Please follow these steps:{5}1. Enter 'G' in cell AW2{6}2. Run the analysis to generate grid system{7}3. Then use '{8}' for frame analysis", "\r\n", "\r\n", text, "\r\n", "\r\n", "\r\n", "\r\n", "\r\n", text);
				}
				else
				{
					List<BeamData> list = new List<BeamData>();
					object instance;
					object[] obj = new object[2]
					{
						NewLateBinding.LateGet(instance = NewLateBinding.LateGet(ws, null, "Rows", new object[0], null, null, null), null, "Count", new object[0], null, null, null),
						1
					};
					object[] array = obj;
					bool[] obj2 = new bool[2] { true, false };
					bool[] array2 = obj2;
					object instance2 = NewLateBinding.LateGet(ws, null, "Cells", obj, null, null, obj2);
					if (array2[0])
					{
						NewLateBinding.LateSetComplex(instance, null, "Count", new object[1] { array[0] }, null, null, OptimisticSet: true, RValueBase: true);
					}
					int num = Conversions.ToInteger(NewLateBinding.LateGet(NewLateBinding.LateGet(instance2, null, "End", new object[1] { -4162 }, null, null, null), null, "Row", new object[0], null, null, null));
					if (num < 9)
					{
						result = "No beam data found for frame analysis.";
					}
					else
					{
						int num2 = num;
						for (int i = 9; i <= num2; i++)
						{
							instance = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 2 }, null, null, array2 = new bool[2] { true, false });
							if (array2[0])
							{
								i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
							}
							if (NewLateBinding.LateGet(instance, null, "Value", new object[0], null, null, null) == null)
							{
								continue;
							}
							BeamData item = default(BeamData);
							object[] obj3 = new object[2] { i, 2 };
							array = obj3;
							bool[] obj4 = new bool[2] { true, false };
							array2 = obj4;
							instance = NewLateBinding.LateGet(ws, null, "Cells", obj3, null, null, obj4);
							if (array2[0])
							{
								i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
							}
							item.BeamNo = (int)Math.Round(Conversion.Val(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(instance, null, "Value", new object[0], null, null, null))));
							instance = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 3 }, null, null, array2 = new bool[2] { true, false });
							if (array2[0])
							{
								i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
							}
							item.Length = Conversion.Val(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(instance, null, "Value", new object[0], null, null, null)));
							instance = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 4 }, null, null, array2 = new bool[2] { true, false });
							if (array2[0])
							{
								i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
							}
							object sectionName;
							if (NewLateBinding.LateGet(instance, null, "Value", new object[0], null, null, null) == null)
							{
								sectionName = "";
							}
							else
							{
								object[] obj5 = new object[2] { i, 4 };
								array = obj5;
								bool[] obj6 = new bool[2] { true, false };
								array2 = obj6;
								instance = NewLateBinding.LateGet(ws, null, "Cells", obj5, null, null, obj6);
								if (array2[0])
								{
									i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
								}
								sectionName = NewLateBinding.LateGet(instance, null, "Value", new object[0], null, null, null).ToString();
							}
							item.SectionName = (string)sectionName;
							instance = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 14 }, null, null, array2 = new bool[2] { true, false });
							if (array2[0])
							{
								i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
							}
							item.NodeA = (int)Math.Round(Conversion.Val(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(instance, null, "Value", new object[0], null, null, null))));
							instance = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 15 }, null, null, array2 = new bool[2] { true, false });
							if (array2[0])
							{
								i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
							}
							item.NodeB = (int)Math.Round(Conversion.Val(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(instance, null, "Value", new object[0], null, null, null))));
							instance = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 16 }, null, null, array2 = new bool[2] { true, false });
							if (array2[0])
							{
								i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
							}
							item.XA = Conversion.Val(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(instance, null, "Value", new object[0], null, null, null)));
							instance = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 17 }, null, null, array2 = new bool[2] { true, false });
							if (array2[0])
							{
								i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
							}
							item.YA = Conversion.Val(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(instance, null, "Value", new object[0], null, null, null)));
							instance = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 18 }, null, null, array2 = new bool[2] { true, false });
							if (array2[0])
							{
								i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
							}
							item.ZA = Conversion.Val(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(instance, null, "Value", new object[0], null, null, null)));
							instance = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 19 }, null, null, array2 = new bool[2] { true, false });
							if (array2[0])
							{
								i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
							}
							item.XB = Conversion.Val(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(instance, null, "Value", new object[0], null, null, null)));
							instance = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 20 }, null, null, array2 = new bool[2] { true, false });
							if (array2[0])
							{
								i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
							}
							item.YB = Conversion.Val(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(instance, null, "Value", new object[0], null, null, null)));
							instance = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 21 }, null, null, array2 = new bool[2] { true, false });
							if (array2[0])
							{
								i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
							}
							item.ZB = Conversion.Val(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(instance, null, "Value", new object[0], null, null, null)));
							instance = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 7 }, null, null, array2 = new bool[2] { true, false });
							if (array2[0])
							{
								i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
							}
							item.F1 = Conversion.Val(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(instance, null, "Value", new object[0], null, null, null)));
							instance = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 8 }, null, null, array2 = new bool[2] { true, false });
							if (array2[0])
							{
								i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
							}
							item.F2 = Conversion.Val(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(instance, null, "Value", new object[0], null, null, null)));
							instance = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 9 }, null, null, array2 = new bool[2] { true, false });
							if (array2[0])
							{
								i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
							}
							item.F3 = Conversion.Val(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(instance, null, "Value", new object[0], null, null, null)));
							instance = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 10 }, null, null, array2 = new bool[2] { true, false });
							if (array2[0])
							{
								i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
							}
							item.F4 = Conversion.Val(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(instance, null, "Value", new object[0], null, null, null)));
							instance = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 11 }, null, null, array2 = new bool[2] { true, false });
							if (array2[0])
							{
								i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
							}
							item.F5 = Conversion.Val(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(instance, null, "Value", new object[0], null, null, null)));
							instance = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 12 }, null, null, array2 = new bool[2] { true, false });
							if (array2[0])
							{
								i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
							}
							item.F6 = Conversion.Val(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(instance, null, "Value", new object[0], null, null, null)));
							instance = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 13 }, null, null, array2 = new bool[2] { true, false });
							if (array2[0])
							{
								i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
							}
							item.F7 = Conversion.Val(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(instance, null, "Value", new object[0], null, null, null)));
							instance = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 22 }, null, null, array2 = new bool[2] { true, false });
							if (array2[0])
							{
								i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
							}
							item.Beta = Conversion.Val(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(instance, null, "Value", new object[0], null, null, null)));
							list.Add(item);
						}
						if (list.Count == 0)
						{
							result = "No valid beam data found for frame analysis.";
						}
						else
						{
							List<BeamData> list2 = FilterBeamsForDefinedGridsOnly(list, text, RuntimeHelpers.GetObjectValue(ws));
							if (list2.Count == 0)
							{
								result = $"No beams found in predefined grids for {text} direction analysis. Check your grid definition using 'G' option first.";
							}
							else
							{
								int num3 = list.Count - list2.Count;
								string text2 = $"Analyzing only predefined grids: {list2.Count} beams from defined grid coordinates. Excluded {num3} beams outside grid system.";
								string text3 = AnalyzeSimilarFrames(list2, text, RuntimeHelpers.GetObjectValue(ws));
								result = string.Format("Grid-Restricted Frame Analysis for '{0}' direction:{1}{2}{3}{4}{5}", text, "\r\n", text3, "\r\n", "\r\n", text2);
							}
						}
					}
				}
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				result = $"Frame analysis error: {ex2.Message}";
				ProjectData.ClearProjectError();
			}
			return result;
		}
	}

	private string ProcessBothDirectionsForGridSilent(object ws, object staadApp)
	{
		checked
		{
			string result;
			try
			{
				if (HasExistingGridData(RuntimeHelpers.GetObjectValue(ws)))
				{
					Interaction.MsgBox("Nova coordinates must be cleared before Genrating new grids.", MsgBoxStyle.Exclamation, "Existing Grid Detected");
					result = "";
				}
				else
				{
					List<BeamData> list = new List<BeamData>();
					object instance;
					object[] obj = new object[2]
					{
						NewLateBinding.LateGet(instance = NewLateBinding.LateGet(ws, null, "Rows", new object[0], null, null, null), null, "Count", new object[0], null, null, null),
						1
					};
					object[] array = obj;
					bool[] obj2 = new bool[2] { true, false };
					bool[] array2 = obj2;
					object instance2 = NewLateBinding.LateGet(ws, null, "Cells", obj, null, null, obj2);
					if (array2[0])
					{
						NewLateBinding.LateSetComplex(instance, null, "Count", new object[1] { array[0] }, null, null, OptimisticSet: true, RValueBase: true);
					}
					int num = Conversions.ToInteger(NewLateBinding.LateGet(NewLateBinding.LateGet(instance2, null, "End", new object[1] { -4162 }, null, null, null), null, "Row", new object[0], null, null, null));
					if (num < 9)
					{
						result = "No beam data found for grid generation.";
					}
					else
					{
						int num2 = num;
						for (int i = 9; i <= num2; i++)
						{
							instance = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 2 }, null, null, array2 = new bool[2] { true, false });
							if (array2[0])
							{
								i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
							}
							if (NewLateBinding.LateGet(instance, null, "Value", new object[0], null, null, null) == null)
							{
								continue;
							}
							BeamData item = default(BeamData);
							object[] obj3 = new object[2] { i, 2 };
							array = obj3;
							bool[] obj4 = new bool[2] { true, false };
							array2 = obj4;
							instance = NewLateBinding.LateGet(ws, null, "Cells", obj3, null, null, obj4);
							if (array2[0])
							{
								i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
							}
							item.BeamNo = (int)Math.Round(Conversion.Val(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(instance, null, "Value", new object[0], null, null, null))));
							instance = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 3 }, null, null, array2 = new bool[2] { true, false });
							if (array2[0])
							{
								i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
							}
							item.Length = Conversion.Val(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(instance, null, "Value", new object[0], null, null, null)));
							instance = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 4 }, null, null, array2 = new bool[2] { true, false });
							if (array2[0])
							{
								i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
							}
							object sectionName;
							if (NewLateBinding.LateGet(instance, null, "Value", new object[0], null, null, null) == null)
							{
								sectionName = "";
							}
							else
							{
								object[] obj5 = new object[2] { i, 4 };
								array = obj5;
								bool[] obj6 = new bool[2] { true, false };
								array2 = obj6;
								instance = NewLateBinding.LateGet(ws, null, "Cells", obj5, null, null, obj6);
								if (array2[0])
								{
									i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
								}
								sectionName = NewLateBinding.LateGet(instance, null, "Value", new object[0], null, null, null).ToString();
							}
							item.SectionName = (string)sectionName;
							instance = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 14 }, null, null, array2 = new bool[2] { true, false });
							if (array2[0])
							{
								i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
							}
							item.NodeA = (int)Math.Round(Conversion.Val(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(instance, null, "Value", new object[0], null, null, null))));
							instance = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 15 }, null, null, array2 = new bool[2] { true, false });
							if (array2[0])
							{
								i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
							}
							item.NodeB = (int)Math.Round(Conversion.Val(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(instance, null, "Value", new object[0], null, null, null))));
							instance = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 16 }, null, null, array2 = new bool[2] { true, false });
							if (array2[0])
							{
								i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
							}
							item.XA = Conversion.Val(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(instance, null, "Value", new object[0], null, null, null)));
							instance = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 17 }, null, null, array2 = new bool[2] { true, false });
							if (array2[0])
							{
								i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
							}
							item.YA = Conversion.Val(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(instance, null, "Value", new object[0], null, null, null)));
							instance = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 18 }, null, null, array2 = new bool[2] { true, false });
							if (array2[0])
							{
								i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
							}
							item.ZA = Conversion.Val(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(instance, null, "Value", new object[0], null, null, null)));
							instance = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 19 }, null, null, array2 = new bool[2] { true, false });
							if (array2[0])
							{
								i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
							}
							item.XB = Conversion.Val(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(instance, null, "Value", new object[0], null, null, null)));
							instance = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 20 }, null, null, array2 = new bool[2] { true, false });
							if (array2[0])
							{
								i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
							}
							item.YB = Conversion.Val(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(instance, null, "Value", new object[0], null, null, null)));
							instance = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 21 }, null, null, array2 = new bool[2] { true, false });
							if (array2[0])
							{
								i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
							}
							item.ZB = Conversion.Val(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(instance, null, "Value", new object[0], null, null, null)));
							instance = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 7 }, null, null, array2 = new bool[2] { true, false });
							if (array2[0])
							{
								i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
							}
							item.F1 = Conversion.Val(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(instance, null, "Value", new object[0], null, null, null)));
							instance = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 8 }, null, null, array2 = new bool[2] { true, false });
							if (array2[0])
							{
								i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
							}
							item.F2 = Conversion.Val(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(instance, null, "Value", new object[0], null, null, null)));
							instance = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 9 }, null, null, array2 = new bool[2] { true, false });
							if (array2[0])
							{
								i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
							}
							item.F3 = Conversion.Val(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(instance, null, "Value", new object[0], null, null, null)));
							instance = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 10 }, null, null, array2 = new bool[2] { true, false });
							if (array2[0])
							{
								i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
							}
							item.F4 = Conversion.Val(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(instance, null, "Value", new object[0], null, null, null)));
							instance = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 11 }, null, null, array2 = new bool[2] { true, false });
							if (array2[0])
							{
								i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
							}
							item.F5 = Conversion.Val(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(instance, null, "Value", new object[0], null, null, null)));
							instance = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 12 }, null, null, array2 = new bool[2] { true, false });
							if (array2[0])
							{
								i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
							}
							item.F6 = Conversion.Val(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(instance, null, "Value", new object[0], null, null, null)));
							instance = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 13 }, null, null, array2 = new bool[2] { true, false });
							if (array2[0])
							{
								i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
							}
							item.F7 = Conversion.Val(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(instance, null, "Value", new object[0], null, null, null)));
							instance = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 22 }, null, null, array2 = new bool[2] { true, false });
							if (array2[0])
							{
								i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
							}
							item.Beta = Conversion.Val(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(instance, null, "Value", new object[0], null, null, null)));
							list.Add(item);
						}
						if (list.Count == 0)
						{
							result = "No valid beam data found for grid generation.";
						}
						else
						{
							List<string> list2 = new List<string>();
							List<BeamData> list3 = FilterBeamsForFrameAnalysis(list, "X");
							if (list3.Count > 0)
							{
								Dictionary<string, List<BeamData>> dictionary = CreateGridSystem(list3, "X", RuntimeHelpers.GetObjectValue(ws));
								list2.Add($"X Direction: {dictionary.Count} grids created");
							}
							else
							{
								list2.Add("X Direction: No beams found in YZ plane");
							}
							List<BeamData> list4 = FilterBeamsForFrameAnalysis(list, "Z");
							if (list4.Count > 0)
							{
								Dictionary<string, List<BeamData>> dictionary2 = CreateGridSystem(list4, "Z", RuntimeHelpers.GetObjectValue(ws));
								list2.Add($"Z Direction: {dictionary2.Count} grids created");
							}
							else
							{
								list2.Add("Z Direction: No beams found in XY plane");
							}
							string text = string.Format("NOVA GRID SYSTEM GENERATED SUCCESSFULLY!{0}{1}Grid data has been created in NovaGrid System:{2}", "\r\n", "\r\n", "\r\n") + string.Join("\r\n", list2) + string.Format("{0}{1}", "\r\n", "\r\n") + string.Format("You can now perform frame analysis using 'X' or 'Z' options.{0}", "\r\n") + "Analysis will be restricted to these predefined grids only.";
							result = text;
						}
					}
				}
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				result = $"Grid generation error: {ex2.Message}";
				ProjectData.ClearProjectError();
			}
			return result;
		}
	}

	private bool HasExistingGridData(object ws)
	{
		bool result;
		try
		{
			object objectValue = RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(ws, null, "Parent", new object[0], null, null, null));
			object obj = null;
			try
			{
				obj = RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue, null, "Sheets", new object[1] { "Sheet4" }, null, null, null));
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				result = false;
				ProjectData.ClearProjectError();
				goto end_IL_0001;
			}
			string[] array = new string[3] { "C2", "C3", "C4" };
			string[] array2 = array;
			int num = 0;
			while (true)
			{
				if (num < array2.Length)
				{
					string text = array2[num];
					object[] array3;
					bool[] array4;
					object instance = NewLateBinding.LateGet(obj, null, "Range", array3 = new object[1] { text }, null, null, array4 = new bool[1] { true });
					if (array4[0])
					{
						text = (string)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array3[0]), typeof(string));
					}
					if (NewLateBinding.LateGet(instance, null, "Value", new object[0], null, null, null) != null)
					{
						object instance2 = obj;
						object[] obj2 = new object[1] { text };
						array3 = obj2;
						bool[] obj3 = new bool[1] { true };
						array4 = obj3;
						instance = NewLateBinding.LateGet(instance2, null, "Range", obj2, null, null, obj3);
						if (array4[0])
						{
							text = (string)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array3[0]), typeof(string));
						}
						if (!string.IsNullOrEmpty(NewLateBinding.LateGet(instance, null, "Value", new object[0], null, null, null).ToString().Trim()))
						{
							result = true;
							break;
						}
					}
					num = checked(num + 1);
					continue;
				}
				result = false;
				break;
			}
			end_IL_0001:;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			result = false;
			ProjectData.ClearProjectError();
		}
		return result;
	}

	private bool AreGridsDefined(string direction, object ws)
	{
		bool result;
		try
		{
			object objectValue = RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(ws, null, "Parent", new object[0], null, null, null));
			object obj = null;
			try
			{
				obj = RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue, null, "Sheets", new object[1] { "Sheet4" }, null, null, null));
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				result = false;
				ProjectData.ClearProjectError();
				goto end_IL_0001;
			}
			string text = "";
			string text2 = "";
			if (Operators.CompareString(direction.ToUpper(), "X", TextCompare: false) == 0)
			{
				text = "B2";
				text2 = "C2";
				goto IL_00c1;
			}
			if (Operators.CompareString(direction.ToUpper(), "Z", TextCompare: false) == 0)
			{
				text = "B4";
				text2 = "C4";
				goto IL_00c1;
			}
			result = false;
			goto end_IL_0001;
			IL_00c1:
			bool flag = false;
			bool flag2 = false;
			object instance = obj;
			object[] obj2 = new object[1] { text };
			object[] array = obj2;
			bool[] obj3 = new bool[1] { true };
			bool[] array2 = obj3;
			object instance2 = NewLateBinding.LateGet(instance, null, "Range", obj2, null, null, obj3);
			if (array2[0])
			{
				text = (string)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(string));
			}
			if (NewLateBinding.LateGet(instance2, null, "Value", new object[0], null, null, null) != null)
			{
				object instance3 = obj;
				object[] obj4 = new object[1] { text };
				array = obj4;
				bool[] obj5 = new bool[1] { true };
				array2 = obj5;
				instance2 = NewLateBinding.LateGet(instance3, null, "Range", obj4, null, null, obj5);
				if (array2[0])
				{
					text = (string)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(string));
				}
				if (!string.IsNullOrEmpty(NewLateBinding.LateGet(instance2, null, "Value", new object[0], null, null, null).ToString().Trim()))
				{
					flag = true;
				}
			}
			instance2 = NewLateBinding.LateGet(obj, null, "Range", array = new object[1] { text2 }, null, null, array2 = new bool[1] { true });
			if (array2[0])
			{
				text2 = (string)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(string));
			}
			if (NewLateBinding.LateGet(instance2, null, "Value", new object[0], null, null, null) != null)
			{
				object instance4 = obj;
				object[] obj6 = new object[1] { text2 };
				array = obj6;
				bool[] obj7 = new bool[1] { true };
				array2 = obj7;
				instance2 = NewLateBinding.LateGet(instance4, null, "Range", obj6, null, null, obj7);
				if (array2[0])
				{
					text2 = (string)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(string));
				}
				if (!string.IsNullOrEmpty(NewLateBinding.LateGet(instance2, null, "Value", new object[0], null, null, null).ToString().Trim()))
				{
					flag2 = true;
				}
			}
			result = flag && flag2;
			end_IL_0001:;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			result = false;
			ProjectData.ClearProjectError();
		}
		return result;
	}

	private List<BeamData> FilterBeamsForDefinedGridsOnly(List<BeamData> beamDataList, string direction, object ws)
	{
		List<BeamData> list = new List<BeamData>();
		double num = 0.01;
		List<double> definedGridCoordinates = GetDefinedGridCoordinates(direction, RuntimeHelpers.GetObjectValue(ws));
		if (definedGridCoordinates.Count == 0)
		{
			return list;
		}
		foreach (BeamData beamData in beamDataList)
		{
			bool flag = false;
			double num2 = 0.0;
			if (Operators.CompareString(direction, "Z", TextCompare: false) == 0)
			{
				if (Math.Abs(beamData.ZA - beamData.ZB) <= num)
				{
					num2 = Math.Round((beamData.ZA + beamData.ZB) / 2.0, 2);
					flag = IsCoordinateInDefinedGrids(num2, definedGridCoordinates, num);
				}
			}
			else if (Operators.CompareString(direction, "X", TextCompare: false) == 0 && Math.Abs(beamData.XA - beamData.XB) <= num)
			{
				num2 = Math.Round((beamData.XA + beamData.XB) / 2.0, 2);
				flag = IsCoordinateInDefinedGrids(num2, definedGridCoordinates, num);
			}
			if (flag)
			{
				list.Add(beamData);
			}
		}
		return list;
	}

	private List<double> GetDefinedGridCoordinates(string direction, object ws)
	{
		List<double> result = new List<double>();
		try
		{
			result = GetCompleteGridDataFromSheet4(direction, RuntimeHelpers.GetObjectValue(ws)).AllCoordinates;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		return result;
	}

	private bool IsCoordinateInDefinedGrids(double coordinate, List<double> definedCoordinates, double tolerance)
	{
		foreach (double definedCoordinate in definedCoordinates)
		{
			double num = definedCoordinate;
			if (Math.Abs(coordinate - num) <= tolerance)
			{
				return true;
			}
		}
		return false;
	}

	private List<BeamData> FilterBeamsForFrameAnalysis(List<BeamData> beamDataList, string direction)
	{
		List<BeamData> list = new List<BeamData>();
		double num = 0.01;
		foreach (BeamData beamData in beamDataList)
		{
			bool flag = false;
			if (Operators.CompareString(direction, "Z", TextCompare: false) == 0)
			{
				if (Math.Abs(beamData.ZA - beamData.ZB) <= num)
				{
					flag = true;
				}
			}
			else if (Operators.CompareString(direction, "X", TextCompare: false) == 0 && Math.Abs(beamData.XA - beamData.XB) <= num)
			{
				flag = true;
			}
			if (flag)
			{
				list.Add(beamData);
			}
		}
		return list;
	}

	private Dictionary<string, List<BeamData>> CreateGridSystem(List<BeamData> beamDataList, string direction, object ws)
	{
		Dictionary<string, List<BeamData>> dictionary = new Dictionary<string, List<BeamData>>();
		double num = 0.01;
		Sheet4GridData completeGridDataFromSheet = GetCompleteGridDataFromSheet4(direction, RuntimeHelpers.GetObjectValue(ws));
		Dictionary<double, List<BeamData>> dictionary2 = new Dictionary<double, List<BeamData>>();
		foreach (BeamData beamData in beamDataList)
		{
			double num2 = 0.0;
			if (Operators.CompareString(direction, "Z", TextCompare: false) == 0)
			{
				num2 = Math.Round((beamData.ZA + beamData.ZB) / 2.0, 2);
			}
			else if (Operators.CompareString(direction, "X", TextCompare: false) == 0)
			{
				num2 = Math.Round((beamData.XA + beamData.XB) / 2.0, 2);
			}
			double num3 = -999999.0;
			foreach (double key6 in dictionary2.Keys)
			{
				double num4 = key6;
				if (Math.Abs(num4 - num2) <= num)
				{
					num3 = num4;
					break;
				}
			}
			if (num3 == -999999.0)
			{
				dictionary2[num2] = new List<BeamData>();
				num3 = num2;
			}
			dictionary2[num3].Add(beamData);
		}
		List<double> list = dictionary2.Keys.OrderBy([SpecialName] (double x) => x).ToList();
		if (ShouldUpdateSheet4CoordinateSpacing(list, completeGridDataFromSheet.AllCoordinates, direction))
		{
			WriteCoordinateSpacingToSheet4(list, direction, RuntimeHelpers.GetObjectValue(ws));
		}
		checked
		{
			if (completeGridDataFromSheet.GridNames.Count > 0 && completeGridDataFromSheet.AllCoordinates.Count > 0)
			{
				foreach (double item in list)
				{
					double num5 = item;
					int num6 = FindBestCoordinateMatch(num5, completeGridDataFromSheet.AllCoordinates, num);
					if (num6 >= 0 && num6 < completeGridDataFromSheet.GridNames.Count)
					{
						string arg = completeGridDataFromSheet.GridNames[num6];
						string key = $"{arg} ({direction}={num5:F2})";
						dictionary[key] = dictionary2[num5];
					}
					else
					{
						string key2 = $"UNMATCHED-{num5:F2} ({direction}={num5:F2})";
						dictionary[key2] = dictionary2[num5];
					}
				}
			}
			else
			{
				List<double> list2 = (from c in list
					where c < -0.005
					orderby c descending
					select c).ToList();
				List<double> list3 = list.Where([SpecialName] (double c) => Math.Abs(c) < 0.005).ToList();
				List<double> list4 = (from c in list
					where c > 0.005
					orderby c
					select c).ToList();
				int num7 = 1;
				int num8 = 1;
				foreach (double item2 in list2)
				{
					double num9 = item2;
					string key3 = $"-{direction}{num7} ({direction}={num9:F2})";
					dictionary[key3] = dictionary2[num9];
					num7++;
				}
				foreach (double item3 in list3)
				{
					double num10 = item3;
					string key4 = $"{direction}0 ({direction}={num10:F2})";
					dictionary[key4] = dictionary2[num10];
				}
				foreach (double item4 in list4)
				{
					double num11 = item4;
					string key5 = $"{direction}{num8} ({direction}={num11:F2})";
					dictionary[key5] = dictionary2[num11];
					num8++;
				}
			}
			return dictionary;
		}
	}

	private Sheet4GridData GetCompleteGridDataFromSheet4(string direction, object ws)
	{
		Sheet4GridData result = new Sheet4GridData
		{
			GridNames = new List<string>(),
			AllCoordinates = new List<double>(),
			SpacingString = ""
		};
		try
		{
			object objectValue = RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(ws, null, "Parent", new object[0], null, null, null));
			object objectValue2 = RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue, null, "Sheets", new object[1] { "Sheet4" }, null, null, null));
			string text = "";
			string text2 = "";
			if (Operators.CompareString(direction.ToUpper(), "X", TextCompare: false) == 0)
			{
				text = "B2";
				text2 = "C2";
			}
			else
			{
				if (Operators.CompareString(direction.ToUpper(), "Z", TextCompare: false) != 0)
				{
					return result;
				}
				text = "B4";
				text2 = "C4";
			}
			object[] array;
			bool[] array2;
			object instance = NewLateBinding.LateGet(objectValue2, null, "Range", array = new object[1] { text }, null, null, array2 = new bool[1] { true });
			if (array2[0])
			{
				text = (string)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(string));
			}
			if (NewLateBinding.LateGet(instance, null, "Value", new object[0], null, null, null) != null)
			{
				object[] obj = new object[1] { text };
				array = obj;
				bool[] obj2 = new bool[1] { true };
				array2 = obj2;
				instance = NewLateBinding.LateGet(objectValue2, null, "Range", obj, null, null, obj2);
				if (array2[0])
				{
					text = (string)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(string));
				}
				string text3 = NewLateBinding.LateGet(instance, null, "Value", new object[0], null, null, null).ToString().Trim();
				if (!string.IsNullOrEmpty(text3))
				{
					string[] array3 = text3.Split(new char[2] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries);
					string[] array4 = array3;
					foreach (string text4 in array4)
					{
						string text5 = text4.Trim();
						if (!string.IsNullOrEmpty(text5))
						{
							result.GridNames.Add(text5);
						}
					}
				}
			}
			instance = NewLateBinding.LateGet(objectValue2, null, "Range", array = new object[1] { text2 }, null, null, array2 = new bool[1] { true });
			if (array2[0])
			{
				text2 = (string)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(string));
			}
			if (NewLateBinding.LateGet(instance, null, "Value", new object[0], null, null, null) != null)
			{
				object[] obj3 = new object[1] { text2 };
				array = obj3;
				bool[] obj4 = new bool[1] { true };
				array2 = obj4;
				instance = NewLateBinding.LateGet(objectValue2, null, "Range", obj3, null, null, obj4);
				if (array2[0])
				{
					text2 = (string)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(string));
				}
				result.SpacingString = NewLateBinding.LateGet(instance, null, "Value", new object[0], null, null, null).ToString().Trim();
				if (!string.IsNullOrEmpty(result.SpacingString))
				{
					result.AllCoordinates = ParseCoordinateSpacing(result.SpacingString);
				}
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		return result;
	}

	private List<double> ParseCoordinateSpacing(string spacingString)
	{
		List<double> list = new List<double>();
		checked
		{
			try
			{
				string[] array = spacingString.Split(new char[1] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
				if (array.Length == 0)
				{
					return list;
				}
				double num = double.Parse(array[0]) / 1000.0;
				list.Add(Math.Round(num, 3));
				int num2 = array.Length - 1;
				for (int i = 1; i <= num2; i++)
				{
					string text = array[i].Trim();
					if (text.Contains("*"))
					{
						string[] array2 = text.Split('*');
						if (array2.Length == 2)
						{
							int num3 = int.Parse(array2[0]);
							double num4 = double.Parse(array2[1]) / 1000.0;
							int num5 = num3;
							for (int j = 1; j <= num5; j++)
							{
								num += num4;
								list.Add(Math.Round(num, 3));
							}
						}
					}
					else
					{
						double num6 = double.Parse(text) / 1000.0;
						num += num6;
						list.Add(Math.Round(num, 3));
					}
				}
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				list.Clear();
				ProjectData.ClearProjectError();
			}
			return list;
		}
	}

	private int FindBestCoordinateMatch(double targetCoord, List<double> sheet4Coordinates, double tolerance)
	{
		int result = -1;
		double num = double.MaxValue;
		checked
		{
			int num2 = sheet4Coordinates.Count - 1;
			for (int i = 0; i <= num2; i++)
			{
				double num3 = Math.Abs(targetCoord - sheet4Coordinates[i]);
				if (num3 <= tolerance && num3 < num)
				{
					num = num3;
					result = i;
				}
			}
			return result;
		}
	}

	private bool ShouldUpdateSheet4CoordinateSpacing(List<double> analysisCoords, List<double> sheet4Coords, string direction)
	{
		if (sheet4Coords.Count == 0)
		{
			return true;
		}
		if (analysisCoords.Count >= sheet4Coords.Count)
		{
			return true;
		}
		return false;
	}

	private void WriteCoordinateSpacingToSheet4(List<double> sortedCoordinates, string direction, object ws)
	{
		checked
		{
			try
			{
				object objectValue = RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(ws, null, "Parent", new object[0], null, null, null));
				object obj = null;
				try
				{
					obj = RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue, null, "Sheets", new object[1] { "Sheet4" }, null, null, null));
				}
				catch (Exception projectError)
				{
					ProjectData.SetProjectError(projectError);
					ProjectData.ClearProjectError();
					return;
				}
				string text = "";
				if (Operators.CompareString(direction.ToUpper(), "X", TextCompare: false) == 0)
				{
					text = "C2";
				}
				else
				{
					if (Operators.CompareString(direction.ToUpper(), "Z", TextCompare: false) != 0)
					{
						return;
					}
					text = "C4";
				}
				object[] array;
				bool[] array2;
				object instance = NewLateBinding.LateGet(obj, null, "Range", array = new object[1] { text }, null, null, array2 = new bool[1] { true });
				if (array2[0])
				{
					text = (string)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(string));
				}
				if (NewLateBinding.LateGet(instance, null, "Value", new object[0], null, null, null) != null)
				{
					object instance2 = obj;
					object[] obj2 = new object[1] { text };
					array = obj2;
					bool[] obj3 = new bool[1] { true };
					array2 = obj3;
					instance = NewLateBinding.LateGet(instance2, null, "Range", obj2, null, null, obj3);
					if (array2[0])
					{
						text = (string)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(string));
					}
					if (!string.IsNullOrEmpty(NewLateBinding.LateGet(instance, null, "Value", new object[0], null, null, null).ToString().Trim()))
					{
						object instance3 = obj;
						object[] obj4 = new object[1] { text };
						array = obj4;
						bool[] obj5 = new bool[1] { true };
						array2 = obj5;
						instance = NewLateBinding.LateGet(instance3, null, "Range", obj4, null, null, obj5);
						if (array2[0])
						{
							text = (string)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(string));
						}
						string spacingString = NewLateBinding.LateGet(instance, null, "Value", new object[0], null, null, null).ToString().Trim();
						List<double> list = ParseCoordinateSpacing(spacingString);
						if (sortedCoordinates.Count < list.Count)
						{
							return;
						}
					}
				}
				if (sortedCoordinates.Count < 2)
				{
					return;
				}
				List<int> list2 = new List<int>();
				int num = sortedCoordinates.Count - 1;
				for (int i = 1; i <= num; i++)
				{
					double value = sortedCoordinates[i] - sortedCoordinates[i - 1];
					int item = (int)Math.Round(Math.Abs(value) * 1000.0);
					list2.Add(item);
				}
				List<string> list3 = new List<string>();
				list3.Add(((int)Math.Round(sortedCoordinates[0] * 1000.0)).ToString());
				int k;
				for (int j = 0; j < list2.Count; j += k)
				{
					int num2 = list2[j];
					for (k = 1; j + k < list2.Count && list2[j + k] == num2; k++)
					{
					}
					if (k == 1)
					{
						list3.Add(num2.ToString());
					}
					else
					{
						list3.Add($"{k}*{num2}");
					}
				}
				string text2 = string.Join(" ", list3);
				object instance4 = obj;
				object[] obj6 = new object[1] { text };
				array = obj6;
				bool[] obj7 = new bool[1] { true };
				array2 = obj7;
				instance = NewLateBinding.LateGet(instance4, null, "Range", obj6, null, null, obj7);
				if (array2[0])
				{
					text = (string)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(string));
				}
				NewLateBinding.LateSetComplex(instance, null, "Value", new object[1] { text2 }, null, null, OptimisticSet: false, RValueBase: true);
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
		}
	}

	private string AnalyzeSimilarFrames(List<BeamData> beamDataList, string direction, object ws)
	{
		Dictionary<string, List<BeamData>> dictionary = CreateGridSystem(beamDataList, direction, RuntimeHelpers.GetObjectValue(ws));
		Dictionary<string, List<BeamData>> dictionary2 = dictionary.Where([SpecialName] (KeyValuePair<string, List<BeamData>> g) => g.Value.Count >= 2).ToDictionary([SpecialName] (KeyValuePair<string, List<BeamData>> k) => k.Key, [SpecialName] (KeyValuePair<string, List<BeamData>> v) => v.Value);
		if (dictionary2.Count < 2)
		{
			return "Insufficient grids for comparison analysis. Need at least 2 grids with 2+ beams each.";
		}
		List<List<KeyValuePair<string, List<BeamData>>>> list = new List<List<KeyValuePair<string, List<BeamData>>>>();
		HashSet<string> hashSet = new HashSet<string>();
		foreach (KeyValuePair<string, List<BeamData>> item in dictionary2)
		{
			if (hashSet.Contains(item.Key))
			{
				continue;
			}
			List<KeyValuePair<string, List<BeamData>>> list2 = new List<KeyValuePair<string, List<BeamData>>>();
			list2.Add(item);
			hashSet.Add(item.Key);
			foreach (KeyValuePair<string, List<BeamData>> item2 in dictionary2)
			{
				if (!hashSet.Contains(item2.Key) && AreFramesSimilar(item.Value, item2.Value, direction))
				{
					list2.Add(item2);
					hashSet.Add(item2.Key);
				}
			}
			if (list2.Count > 1)
			{
				list.Add(list2);
			}
		}
		if (Operators.CompareString(direction, "G", TextCompare: false) != 0)
		{
			string content = GenerateEnhancedFrameAnalysisLog(list, dictionary, direction, RuntimeHelpers.GetObjectValue(ws));
			string lastLogFilePath = SaveLogFile(content, direction, RuntimeHelpers.GetObjectValue(ws));
			_lastLogFilePath = lastLogFilePath;
			return $"Frame analysis completed. Similar sets: {list.Count}. Log saved to file.";
		}
		_lastLogFilePath = "";
		return $"Grid system created. Similar sets: {list.Count}";
	}

	private string GetGridNamingInfoForResult(string direction, object ws)
	{
		string result;
		try
		{
			object objectValue = RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(ws, null, "Parent", new object[0], null, null, null));
			object objectValue2 = RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue, null, "Sheets", new object[1] { "Sheet4" }, null, null, null));
			string text = "";
			string text2 = "";
			if (Operators.CompareString(direction.ToUpper(), "X", TextCompare: false) == 0)
			{
				text2 = "B2";
				if (NewLateBinding.LateGet(NewLateBinding.LateGet(objectValue2, null, "Range", new object[1] { "B2" }, null, null, null), null, "Value", new object[0], null, null, null) != null)
				{
					text = NewLateBinding.LateGet(NewLateBinding.LateGet(objectValue2, null, "Range", new object[1] { "B2" }, null, null, null), null, "Value", new object[0], null, null, null).ToString().Trim();
				}
			}
			else if (Operators.CompareString(direction.ToUpper(), "Z", TextCompare: false) == 0)
			{
				text2 = "B4";
				if (NewLateBinding.LateGet(NewLateBinding.LateGet(objectValue2, null, "Range", new object[1] { "B4" }, null, null, null), null, "Value", new object[0], null, null, null) != null)
				{
					text = NewLateBinding.LateGet(NewLateBinding.LateGet(objectValue2, null, "Range", new object[1] { "B4" }, null, null, null), null, "Value", new object[0], null, null, null).ToString().Trim();
				}
			}
			result = (string.IsNullOrEmpty(text) ? "default naming (Nova's Grid System is empty)" : ((GetCompleteGridDataFromSheet4(direction, RuntimeHelpers.GetObjectValue(ws)).AllCoordinates.Count <= 0) ? $"From Nova's Grid System: '{text}'" : $"From Nova's Grid System: '{text}' (Coordinate-matched)"));
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			result = "default naming (Nova's Grid System not available)";
			ProjectData.ClearProjectError();
		}
		return result;
	}

	private bool AreFramesSimilar(List<BeamData> frame1, List<BeamData> frame2, string direction)
	{
		if (frame1.Count != frame2.Count)
		{
			return false;
		}
		List<BeamData> list = new List<BeamData>(frame1);
		List<BeamData> list2 = new List<BeamData>(frame2);
		double num = 0.0;
		double num2 = 0.0;
		double num3 = 0.0;
		double num4 = 0.0;
		if (Operators.CompareString(direction, "Z", TextCompare: false) == 0)
		{
			num = list.Min([SpecialName] (BeamData b) => Math.Min(b.XA, b.XB));
			num2 = list.Min([SpecialName] (BeamData b) => Math.Min(b.YA, b.YB));
			num3 = list2.Min([SpecialName] (BeamData b) => Math.Min(b.XA, b.XB));
			num4 = list2.Min([SpecialName] (BeamData b) => Math.Min(b.YA, b.YB));
		}
		else
		{
			if (Operators.CompareString(direction, "X", TextCompare: false) != 0)
			{
				return false;
			}
			num = list.Min([SpecialName] (BeamData b) => Math.Min(b.YA, b.YB));
			num2 = list.Min([SpecialName] (BeamData b) => Math.Min(b.ZA, b.ZB));
			num3 = list2.Min([SpecialName] (BeamData b) => Math.Min(b.YA, b.YB));
			num4 = list2.Min([SpecialName] (BeamData b) => Math.Min(b.ZA, b.ZB));
		}
		checked
		{
			int num5 = list.Count - 1;
			for (int num6 = 0; num6 <= num5; num6++)
			{
				BeamData value = list[num6];
				if (Operators.CompareString(direction, "Z", TextCompare: false) == 0)
				{
					value.XA -= num;
					value.XB -= num;
					value.YA -= num2;
					value.YB -= num2;
				}
				else
				{
					value.YA -= num;
					value.YB -= num;
					value.ZA -= num2;
					value.ZB -= num2;
				}
				list[num6] = value;
			}
			int num7 = list2.Count - 1;
			for (int num8 = 0; num8 <= num7; num8++)
			{
				BeamData value2 = list2[num8];
				if (Operators.CompareString(direction, "Z", TextCompare: false) == 0)
				{
					value2.XA -= num3;
					value2.XB -= num3;
					value2.YA -= num4;
					value2.YB -= num4;
				}
				else
				{
					value2.YA -= num3;
					value2.YB -= num3;
					value2.ZA -= num4;
					value2.ZB -= num4;
				}
				list2[num8] = value2;
			}
			List<BeamData> list3;
			List<BeamData> list4;
			if (Operators.CompareString(direction, "Z", TextCompare: false) == 0)
			{
				list3 = (from b in list
					orderby Math.Min(b.XA, b.XB), Math.Min(b.YA, b.YB)
					select b).ToList();
				list4 = (from b in list2
					orderby Math.Min(b.XA, b.XB), Math.Min(b.YA, b.YB)
					select b).ToList();
			}
			else
			{
				list3 = (from b in list
					orderby Math.Min(b.YA, b.YB), Math.Min(b.ZA, b.ZB)
					select b).ToList();
				list4 = (from b in list2
					orderby Math.Min(b.YA, b.YB), Math.Min(b.ZA, b.ZB)
					select b).ToList();
			}
			int num9 = list3.Count - 1;
			for (int num10 = 0; num10 <= num9; num10++)
			{
				BeamData beam = list3[num10];
				BeamData beam2 = list4[num10];
				if (!AreSectionsSimilar(beam, beam2))
				{
					return false;
				}
				if (Math.Abs(beam.Beta - beam2.Beta) > 0.5)
				{
					return false;
				}
				if (!AreBeamPositionsSimilar(beam, beam2, direction))
				{
					return false;
				}
				if (Math.Abs(beam.Length - beam2.Length) > 0.01)
				{
					return false;
				}
				double num11;
				double num12;
				if (Operators.CompareString(direction, "Z", TextCompare: false) == 0)
				{
					num11 = beam.ZB - beam.ZA;
					num12 = beam2.ZB - beam2.ZA;
				}
				else
				{
					num11 = beam.XB - beam.XA;
					num12 = beam2.XB - beam2.XA;
				}
				if (Math.Abs(num11 - num12) > 0.01)
				{
					return false;
				}
			}
			return true;
		}
	}

	private bool AreBeamPositionsSimilar(BeamData beam1, BeamData beam2, string direction)
	{
		double num = 0.01;
		if (Operators.CompareString(direction, "Z", TextCompare: false) == 0)
		{
			return Math.Abs(beam1.XA - beam2.XA) <= num && Math.Abs(beam1.YA - beam2.YA) <= num && Math.Abs(beam1.XB - beam2.XB) <= num && Math.Abs(beam1.YB - beam2.YB) <= num;
		}
		if (Operators.CompareString(direction, "X", TextCompare: false) == 0)
		{
			return Math.Abs(beam1.YA - beam2.YA) <= num && Math.Abs(beam1.ZA - beam2.ZA) <= num && Math.Abs(beam1.YB - beam2.YB) <= num && Math.Abs(beam1.ZB - beam2.ZB) <= num;
		}
		return false;
	}

	private bool AreSectionsSimilar(BeamData beam1, BeamData beam2)
	{
		double num = 0.001;
		if (Operators.CompareString(beam1.SectionName, beam2.SectionName, TextCompare: false) != 0)
		{
			return false;
		}
		if (Math.Abs(beam1.F1 - beam2.F1) > num)
		{
			return false;
		}
		if (Math.Abs(beam1.F2 - beam2.F2) > num)
		{
			return false;
		}
		if (Math.Abs(beam1.F3 - beam2.F3) > num)
		{
			return false;
		}
		if (Math.Abs(beam1.F4 - beam2.F4) > num)
		{
			return false;
		}
		if (Math.Abs(beam1.F5 - beam2.F5) > num)
		{
			return false;
		}
		if (Math.Abs(beam1.F6 - beam2.F6) > num)
		{
			return false;
		}
		if (Math.Abs(beam1.F7 - beam2.F7) > num)
		{
			return false;
		}
		return true;
	}

	private string GenerateEnhancedFrameAnalysisLog(List<List<KeyValuePair<string, List<BeamData>>>> similarGridGroups, Dictionary<string, List<BeamData>> allGridGroups, string direction, object ws = null)
	{
		StringBuilder stringBuilder = new StringBuilder();
		stringBuilder.AppendLine("\ud83c\udfd7\ufe0f NOVA'S STAAD FRAME SIMILARITY ANALYSIS \ud83c\udfd7\ufe0f");
		stringBuilder.AppendLine("\ud83d\udcd0 Advanced Grid Reference System \ud83d\udcd0");
		stringBuilder.AppendLine();
		stringBuilder.AppendLine("=" + new string('=', 80));
		stringBuilder.AppendLine();
		stringBuilder.AppendLine("\ud83d\udcca ANALYSIS METADATA");
		stringBuilder.AppendLine();
		stringBuilder.AppendLine("\ud83d\udcc5 Analysis Date: " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
		stringBuilder.AppendLine("\ud83e\udded Analysis Direction: " + direction);
		stringBuilder.AppendLine("\ud83d\udccf Total Grids Analyzed: " + allGridGroups.Count);
		stringBuilder.AppendLine("\ud83c\udfaf Similar Groups Found: " + similarGridGroups.Count);
		if (ws != null)
		{
			string gridNamingInfoForResult = GetGridNamingInfoForResult(direction, RuntimeHelpers.GetObjectValue(ws));
			stringBuilder.AppendLine("\ud83d\udccb Grid Naming Source: " + gridNamingInfoForResult);
			try
			{
				object objectValue = RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(ws, null, "Parent", new object[0], null, null, null));
				object objectValue2 = RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue, null, "Sheets", new object[1] { "Sheet4" }, null, null, null));
				string text = ((Operators.CompareString(direction.ToUpper(), "X", TextCompare: false) == 0) ? "C2" : "C4");
				object[] array;
				bool[] array2;
				object instance = NewLateBinding.LateGet(objectValue2, null, "Range", array = new object[1] { text }, null, null, array2 = new bool[1] { true });
				if (array2[0])
				{
					text = (string)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(string));
				}
				if (NewLateBinding.LateGet(instance, null, "Value", new object[0], null, null, null) != null)
				{
					object[] obj = new object[1] { text };
					array = obj;
					bool[] obj2 = new bool[1] { true };
					array2 = obj2;
					instance = NewLateBinding.LateGet(objectValue2, null, "Range", obj, null, null, obj2);
					if (array2[0])
					{
						text = (string)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(string));
					}
					stringBuilder.AppendLine("\ud83d\udcd0 Coordinate Spacing: " + string.Format("{0} ({1})", RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(instance, null, "Value", new object[0], null, null, null)), text));
				}
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				stringBuilder.AppendLine("\ud83d\udcd0 Coordinate Spacing: Sheet4 not accessible");
				ProjectData.ClearProjectError();
			}
		}
		stringBuilder.AppendLine();
		stringBuilder.AppendLine("=" + new string('=', 80));
		stringBuilder.AppendLine();
		stringBuilder.AppendLine("\ud83c\udfaf SIMILAR GRID GROUPS SUMMARY");
		stringBuilder.AppendLine();
		stringBuilder.AppendLine("Group #" + new string(' ', 10) + "➤ Grid Count" + new string(' ', 10) + "➤ Grid Names");
		stringBuilder.AppendLine(new string('-', 80));
		checked
		{
			int num = similarGridGroups.Count - 1;
			for (int i = 0; i <= num; i++)
			{
				List<KeyValuePair<string, List<BeamData>>> list = similarGridGroups[i];
				string text2 = string.Join(", ", list.Select([SpecialName] (KeyValuePair<string, List<BeamData>> g) => g.Key.Split('(')[0].Trim()));
				if (text2.Length > 45)
				{
					text2 = text2.Substring(0, 42) + "...";
				}
				stringBuilder.AppendLine($"Group {(i + 1).ToString().PadLeft(2)}" + new string(' ', 12) + "➤ " + list.Count.ToString().PadLeft(6) + new string(' ', 12) + "➤ " + text2);
			}
			stringBuilder.AppendLine();
			stringBuilder.AppendLine("=" + new string('=', 80));
			stringBuilder.AppendLine();
			stringBuilder.AppendLine("\ud83d\uddc2\ufe0f GRID SYSTEM OVERVIEW");
			stringBuilder.AppendLine();
			stringBuilder.AppendLine("Grid Name" + new string(' ', 30) + "Beam Count" + new string(' ', 10) + "Coordinate");
			stringBuilder.AppendLine(new string('-', 80));
			foreach (KeyValuePair<string, List<BeamData>> item in allGridGroups.OrderBy([SpecialName] (KeyValuePair<string, List<BeamData>> g) => ExtractCoordinateFromGridName(g.Key)))
			{
				string text3 = item.Key.Split('(')[0].Trim();
				string text4 = ExtractCoordinateFromGridName(item.Key).ToString("F1");
				stringBuilder.AppendLine(text3.PadRight(35) + item.Value.Count.ToString().PadLeft(15) + text4.PadLeft(15));
			}
			stringBuilder.AppendLine();
			stringBuilder.AppendLine("=" + new string('=', 80));
			stringBuilder.AppendLine();
			stringBuilder.AppendLine("\ud83d\udccb DETAILED GROUP ANALYSIS");
			stringBuilder.AppendLine();
			int num2 = similarGridGroups.Count - 1;
			for (int num3 = 0; num3 <= num2; num3++)
			{
				List<KeyValuePair<string, List<BeamData>>> list2 = similarGridGroups[num3];
				stringBuilder.AppendLine($"\ud83c\udfd7\ufe0f SIMILAR GRID GROUP #{num3 + 1}");
				stringBuilder.AppendLine();
				stringBuilder.AppendLine($"\ud83d\udcca Similar Grids Count: {list2.Count}");
				stringBuilder.AppendLine($"\ud83d\udd27 Beams per Grid: {list2[0].Value.Count}");
				stringBuilder.AppendLine();
				stringBuilder.AppendLine("\ud83d\uddc2\ufe0f GRIDS IN THIS GROUP:");
				foreach (KeyValuePair<string, List<BeamData>> item2 in list2)
				{
					stringBuilder.AppendLine($"   ➤ {item2.Key}");
				}
				stringBuilder.AppendLine();
				int num4 = list2.Count - 1;
				for (int num5 = 0; num5 <= num4; num5++)
				{
					KeyValuePair<string, List<BeamData>> keyValuePair = list2[num5];
					stringBuilder.AppendLine($"\ud83d\udcd0 {keyValuePair.Key} DETAILS:");
					stringBuilder.AppendLine();
					foreach (BeamData item3 in keyValuePair.Value.OrderBy([SpecialName] (BeamData b) => b.BeamNo))
					{
						stringBuilder.AppendLine($"   \ud83d\udd39 Beam {item3.BeamNo.ToString().PadLeft(3)}: {item3.SectionName}");
						stringBuilder.AppendLine($"      \ud83d\udd17 Nodes: {item3.NodeA} ➜ {item3.NodeB}");
						stringBuilder.AppendLine($"      \ud83d\udccd Coord A: ({item3.XA:F3}, {item3.YA:F3}, {item3.ZA:F3})");
						stringBuilder.AppendLine($"      \ud83d\udccd Coord B: ({item3.XB:F3}, {item3.YB:F3}, {item3.ZB:F3})");
						stringBuilder.AppendLine($"      \ud83d\udccf Length: {item3.Length:F3}m");
						stringBuilder.AppendLine($"      ⚙\ufe0f Section Props: F1={item3.F1:F3} F2={item3.F2:F3} F3={item3.F3:F3}");
						stringBuilder.AppendLine($"         F4={item3.F4:F3} F5={item3.F5:F3} F6={item3.F6:F3} F7={item3.F7:F3}");
						stringBuilder.AppendLine();
					}
					if (num5 < list2.Count - 1)
					{
						stringBuilder.AppendLine(new string('-', 50));
						stringBuilder.AppendLine();
					}
				}
				if (num3 < similarGridGroups.Count - 1)
				{
					stringBuilder.AppendLine("=" + new string('=', 80));
					stringBuilder.AppendLine();
				}
			}
			stringBuilder.AppendLine("=" + new string('=', 80));
			stringBuilder.AppendLine();
			stringBuilder.AppendLine("\ud83d\udcc8 COMPREHENSIVE ANALYSIS SUMMARY");
			stringBuilder.AppendLine();
			int num6 = 0;
			int num7 = 0;
			foreach (List<KeyValuePair<string, List<BeamData>>> similarGridGroup in similarGridGroups)
			{
				num6 += similarGridGroup.Count;
				foreach (KeyValuePair<string, List<BeamData>> item4 in similarGridGroup)
				{
					num7 += item4.Value.Count;
				}
			}
			stringBuilder.AppendLine($"Total Grids with Similar Patterns: {num6}");
			stringBuilder.AppendLine($"Total Beams in Similar Grids: {num7}");
			if (num6 > 0)
			{
				string arg = ((double)num7 / (double)num6).ToString("F1");
				stringBuilder.AppendLine($"Average Beams per Similar Grid: {arg}");
			}
			else
			{
				stringBuilder.AppendLine("Average Beams per Similar Grid: 0");
			}
			string arg2 = ((allGridGroups.Count > 0) ? ((double)num6 / (double)allGridGroups.Count * 100.0).ToString("F1") : "0");
			stringBuilder.AppendLine($"Similarity Efficiency: {arg2}%");
			stringBuilder.AppendLine();
			stringBuilder.AppendLine("=" + new string('=', 80));
			stringBuilder.AppendLine();
			stringBuilder.AppendLine("*** Report Generated by PlateNova's RaghavStaadDataExtractor ***");
			stringBuilder.AppendLine("Enhanced with Nova Grid Intelligence System");
			stringBuilder.AppendLine(string.Format("Generated on: {0}", DateTime.Now.ToString("dddd, MMMM dd, yyyy 'at' HH:mm:ss")));
			return stringBuilder.ToString();
		}
	}

	private double ExtractCoordinateFromGridName(string gridName)
	{
		try
		{
			int num = checked(gridName.IndexOf("=") + 1);
			int num2 = gridName.IndexOf(")");
			if (num > 0 && num2 > num)
			{
				string s = gridName.Substring(num, checked(num2 - num));
				return double.Parse(s);
			}
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
		return 0.0;
	}

	private string SaveLogFile(string content, string direction, object ws)
	{
		string result;
		try
		{
			string path = string.Format("Frame_Analysis_{0}_Direction_{1}.txt", direction, DateTime.Now.ToString("yyyyMMdd_HHmmss"));
			object objectValue = RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(ws, null, "Parent", new object[0], null, null, null));
			string text = "";
			text = ((!Operators.ConditionalCompareObjectNotEqual(NewLateBinding.LateGet(objectValue, null, "Path", new object[0], null, null, null), "", TextCompare: false)) ? Environment.GetFolderPath(Environment.SpecialFolder.Desktop) : Conversions.ToString(NewLateBinding.LateGet(objectValue, null, "Path", new object[0], null, null, null)));
			string text2 = Path.Combine(text, "Frame Analysis");
			if (!Directory.Exists(text2))
			{
				Directory.CreateDirectory(text2);
			}
			string text3 = Path.Combine(text2, path);
			File.WriteAllText(text3, content);
			result = text3;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			try
			{
				string path2 = string.Format("Frame_Analysis_{0}_Direction_{1}.txt", direction, DateTime.Now.ToString("yyyyMMdd_HHmmss"));
				string folderPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
				string text4 = Path.Combine(folderPath, "Frame Analysis");
				if (!Directory.Exists(text4))
				{
					Directory.CreateDirectory(text4);
				}
				string text5 = Path.Combine(text4, path2);
				File.WriteAllText(text5, content);
				result = text5;
				ProjectData.ClearProjectError();
			}
			catch (Exception ex3)
			{
				ProjectData.SetProjectError(ex3);
				Exception ex4 = ex3;
				result = $"Error saving log file: {ex4.Message}";
				ProjectData.ClearProjectError();
			}
		}
		return result;
	}

	private void AutoSaveWorkbook(object workbook, string savePath = "")
	{
		try
		{
			if (workbook == null)
			{
				return;
			}
			if (string.IsNullOrEmpty(savePath))
			{
				if (Operators.ConditionalCompareObjectNotEqual(NewLateBinding.LateGet(workbook, null, "Path", new object[0], null, null, null), "", TextCompare: false))
				{
					NewLateBinding.LateCall(workbook, null, "Save", new object[0], null, null, null, IgnoreReturn: true);
					return;
				}
				string folderPath = Environment.GetFolderPath(Environment.SpecialFolder.Personal);
				savePath = Path.Combine(folderPath, "Staad_Extract_" + DateAndTime.Now.ToString("yyyyMMdd_HHmmss") + ".xlsx");
				object[] obj = new object[1] { savePath };
				object[] array = obj;
				bool[] obj2 = new bool[1] { true };
				bool[] array2 = obj2;
				NewLateBinding.LateCall(workbook, null, "SaveAs", obj, null, null, obj2, IgnoreReturn: true);
				if (array2[0])
				{
					savePath = (string)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(string));
				}
			}
			else
			{
				object[] obj3 = new object[1] { savePath };
				object[] array = obj3;
				bool[] obj4 = new bool[1] { true };
				bool[] array2 = obj4;
				NewLateBinding.LateCall(workbook, null, "SaveAs", obj3, null, null, obj4, IgnoreReturn: true);
				if (array2[0])
				{
					savePath = (string)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(string));
				}
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	private void EnsureLicense()
	{
		ExecutionValidation executionValidation = new ExecutionValidation();
		if (!executionValidation.IsLicenceValid())
		{
			throw new UnauthorizedAccessException("License check failed.");
		}
	}

	private void CheckExpirationDateAndBackdate()
	{
		DateTime t = new DateTime(2027, 4, 9);
		DateTime date = DateTime.Now.Date;
		string path = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "RaghavStaad.last");
		if (File.Exists(path))
		{
			string s = File.ReadAllText(path).Trim();
			DateTime result = DateTime.MinValue;
			if (DateTime.TryParse(s, out result))
			{
				DateTime t2 = DateTime.Parse(s);
				if (DateTime.Compare(date, t2) < 0)
				{
					throw new UnauthorizedAccessException("System date has been tampered. Please correct your clock.");
				}
			}
		}
		File.WriteAllText(path, date.ToString("yyyy-MM-dd"));
		if (DateTime.Compare(date, t) > 0)
		{
			throw new UnauthorizedAccessException("Software expired. Please contact support for an updated version.");
		}
	}

	private object GetOrCreateExcelApp()
	{
		object result;
		try
		{
			result = Marshal.GetActiveObject("Excel.Application");
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			object objectValue = RuntimeHelpers.GetObjectValue(Interaction.CreateObject("Excel.Application"));
			NewLateBinding.LateSet(objectValue, null, "Visible", new object[1] { true }, null, null);
			result = objectValue;
			ProjectData.ClearProjectError();
		}
		return result;
	}

	private object GetStaadApplication()
	{
		object result;
		try
		{
			result = Marshal.GetActiveObject("StaadPro.OpenSTAAD");
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			result = null;
			ProjectData.ClearProjectError();
		}
		return result;
	}

	private object GetOrCreateWorksheet(object workbook, string sheetName)
	{
		object result;
		try
		{
			object[] obj = new object[1] { sheetName };
			object[] array = obj;
			bool[] obj2 = new bool[1] { true };
			bool[] array2 = obj2;
			object obj3 = NewLateBinding.LateGet(workbook, null, "Sheets", obj, null, null, obj2);
			if (array2[0])
			{
				sheetName = (string)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(string));
			}
			result = obj3;
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			object objectValue = RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(NewLateBinding.LateGet(workbook, null, "Sheets", new object[0], null, null, null), null, "Add", new object[0], null, null, null));
			NewLateBinding.LateSet(objectValue, null, "Name", new object[1] { sheetName }, null, null);
			result = objectValue;
			ProjectData.ClearProjectError();
		}
		return result;
	}

	private void PreserveAndClearWorksheet(object ws)
	{
		object[] array = new object[9]
		{
			null,
			RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(NewLateBinding.LateGet(ws, null, "Range", new object[1] { "AS2" }, null, null, null), null, "Value", new object[0], null, null, null)),
			RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(NewLateBinding.LateGet(ws, null, "Range", new object[1] { "AS3" }, null, null, null), null, "Value", new object[0], null, null, null)),
			RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(NewLateBinding.LateGet(ws, null, "Range", new object[1] { "AS4" }, null, null, null), null, "Value", new object[0], null, null, null)),
			RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(NewLateBinding.LateGet(ws, null, "Range", new object[1] { "AS5" }, null, null, null), null, "Value", new object[0], null, null, null)),
			RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(NewLateBinding.LateGet(ws, null, "Range", new object[1] { "AR2" }, null, null, null), null, "Value", new object[0], null, null, null)),
			RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(NewLateBinding.LateGet(ws, null, "Range", new object[1] { "AR7" }, null, null, null), null, "Value", new object[0], null, null, null)),
			RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(NewLateBinding.LateGet(ws, null, "Range", new object[1] { "AS7" }, null, null, null), null, "Value", new object[0], null, null, null)),
			RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(NewLateBinding.LateGet(ws, null, "Range", new object[1] { "AW2" }, null, null, null), null, "Value", new object[0], null, null, null))
		};
		NewLateBinding.LateCall(NewLateBinding.LateGet(ws, null, "Cells", new object[0], null, null, null), null, "Clear", new object[0], null, null, null, IgnoreReturn: true);
		NewLateBinding.LateSetComplex(NewLateBinding.LateGet(ws, null, "Range", new object[1] { "AS2" }, null, null, null), null, "Value", new object[1] { array[1] }, null, null, OptimisticSet: false, RValueBase: true);
		NewLateBinding.LateSetComplex(NewLateBinding.LateGet(ws, null, "Range", new object[1] { "AS3" }, null, null, null), null, "Value", new object[1] { array[2] }, null, null, OptimisticSet: false, RValueBase: true);
		NewLateBinding.LateSetComplex(NewLateBinding.LateGet(ws, null, "Range", new object[1] { "AS4" }, null, null, null), null, "Value", new object[1] { array[3] }, null, null, OptimisticSet: false, RValueBase: true);
		NewLateBinding.LateSetComplex(NewLateBinding.LateGet(ws, null, "Range", new object[1] { "AS5" }, null, null, null), null, "Value", new object[1] { array[4] }, null, null, OptimisticSet: false, RValueBase: true);
		NewLateBinding.LateSetComplex(NewLateBinding.LateGet(ws, null, "Range", new object[1] { "AR2" }, null, null, null), null, "Value", new object[1] { array[5] }, null, null, OptimisticSet: false, RValueBase: true);
		NewLateBinding.LateSetComplex(NewLateBinding.LateGet(ws, null, "Range", new object[1] { "AR7" }, null, null, null), null, "Value", new object[1] { array[6] }, null, null, OptimisticSet: false, RValueBase: true);
		NewLateBinding.LateSetComplex(NewLateBinding.LateGet(ws, null, "Range", new object[1] { "AS7" }, null, null, null), null, "Value", new object[1] { array[7] }, null, null, OptimisticSet: false, RValueBase: true);
		NewLateBinding.LateSetComplex(NewLateBinding.LateGet(ws, null, "Range", new object[1] { "AW2" }, null, null, null), null, "Value", new object[1] { array[8] }, null, null, OptimisticSet: false, RValueBase: true);
		object instance = NewLateBinding.LateGet(ws, null, "Range", new object[1] { "AM4" }, null, null, null);
		NewLateBinding.LateSetComplex(instance, null, "Value", new object[1] { DateTime.Now.ToString("dd-MM-yyyy") }, null, null, OptimisticSet: false, RValueBase: true);
		NewLateBinding.LateSetComplex(NewLateBinding.LateGet(instance, null, "Font", new object[0], null, null, null), null, "Bold", new object[1] { true }, null, null, OptimisticSet: false, RValueBase: true);
		NewLateBinding.LateSetComplex(instance, null, "HorizontalAlignment", new object[1] { -4108 }, null, null, OptimisticSet: false, RValueBase: true);
		NewLateBinding.LateSetComplex(instance, null, "VerticalAlignment", new object[1] { -4108 }, null, null, OptimisticSet: false, RValueBase: true);
		NewLateBinding.LateSetComplex(NewLateBinding.LateGet(instance, null, "EntireColumn", new object[0], null, null, null), null, "ColumnWidth", new object[1] { 12 }, null, null, OptimisticSet: false, RValueBase: true);
		instance = null;
		object instance2 = NewLateBinding.LateGet(ws, null, "Range", new object[1] { "AW2" }, null, null, null);
		if (NewLateBinding.LateGet(NewLateBinding.LateGet(ws, null, "Range", new object[1] { "AW2" }, null, null, null), null, "Value", new object[0], null, null, null) == null)
		{
			NewLateBinding.LateSetComplex(instance2, null, "Value", new object[1] { "" }, null, null, OptimisticSet: false, RValueBase: true);
		}
		NewLateBinding.LateSetComplex(NewLateBinding.LateGet(instance2, null, "Font", new object[0], null, null, null), null, "Bold", new object[1] { true }, null, null, OptimisticSet: false, RValueBase: true);
		NewLateBinding.LateSetComplex(instance2, null, "HorizontalAlignment", new object[1] { -4108 }, null, null, OptimisticSet: false, RValueBase: true);
		NewLateBinding.LateSetComplex(instance2, null, "VerticalAlignment", new object[1] { -4108 }, null, null, OptimisticSet: false, RValueBase: true);
		NewLateBinding.LateSetComplex(NewLateBinding.LateGet(instance2, null, "Interior", new object[0], null, null, null), null, "Color", new object[1] { Color.FromArgb(146, 208, 80) }, null, null, OptimisticSet: false, RValueBase: true);
		NewLateBinding.LateSetComplex(NewLateBinding.LateGet(instance2, null, "Font", new object[0], null, null, null), null, "Color", new object[1] { Color.FromArgb(0, 0, 0) }, null, null, OptimisticSet: false, RValueBase: true);
		NewLateBinding.LateSetComplex(NewLateBinding.LateGet(instance2, null, "Borders", new object[0], null, null, null), null, "LineStyle", new object[1] { 1 }, null, null, OptimisticSet: false, RValueBase: true);
		instance2 = null;
		object instance3 = NewLateBinding.LateGet(ws, null, "Range", new object[1] { "AW1" }, null, null, null);
		NewLateBinding.LateSetComplex(instance3, null, "Value", new object[1] { "Frame Analysis:" }, null, null, OptimisticSet: false, RValueBase: true);
		NewLateBinding.LateSetComplex(NewLateBinding.LateGet(instance3, null, "Font", new object[0], null, null, null), null, "Bold", new object[1] { true }, null, null, OptimisticSet: false, RValueBase: true);
		NewLateBinding.LateSetComplex(NewLateBinding.LateGet(instance3, null, "Font", new object[0], null, null, null), null, "Color", new object[1] { Color.FromArgb(0, 32, 96) }, null, null, OptimisticSet: false, RValueBase: true);
		NewLateBinding.LateSetComplex(instance3, null, "HorizontalAlignment", new object[1] { -4131 }, null, null, OptimisticSet: false, RValueBase: true);
		NewLateBinding.LateSetComplex(instance3, null, "VerticalAlignment", new object[1] { -4108 }, null, null, OptimisticSet: false, RValueBase: true);
		instance3 = null;
	}

	private void WriteJobInfoToExcel(object ws, object staadApp)
	{
		string text = "";
		string text2 = "";
		string text3 = "";
		string text4 = "";
		string text5 = "";
		string text6 = "";
		string text7 = "";
		string text8 = "";
		string text9 = "";
		string text10 = "";
		string text11 = "";
		string text12 = "";
		string text13 = "";
		string[] array = new string[19]
		{
			"BUILT WITH PURPOSE — RAGHAV", "INNOVATION IN EVERY CELL — RAGHAV", "CRAFTED BY RAGHAV, POWERED BY CODE", "BRINGING LOGIC TO LIFE — RAGHAV", "DESIGNED FOR EXCELLENCE — RAGHAV", "ENGINEERED BY RAGHAV WITH PRECISION", "CREATIVE CODE BY RAGHAV", "WHERE IDEAS MEET EXECUTION — RAGHAV", "LINES OF CODE, LINES OF THOUGHT — RAGHAV", "BY RAGHAV — FOR PERFECTION",
			"CRAFTED WITH PASSION, WRITTEN IN LOGIC — RAGHAV", "THE ART OF AUTOMATION — RAGHAV", "VISION TO VALUE — BY RAGHAV", "ENGINEERING BEYOND LIMITS — RAGHAV", "WHERE LOGIC MEETS AESTHETICS — RAGHAV", "EVERY CELL TELLS A STORY — RAGHAV", "STRUCTURE, STYLE, STABILITY — RAGHAV", "TECH WITH A HUMAN TOUCH — RAGHAV", "SOLUTIONS THAT SPEAK — RAGHAV"
		};
		Random random = new Random();
		string text14 = array[random.Next(array.Length)];
		object[] obj = new object[13]
		{
			text, text2, text3, text4, text5, text6, text7, text8, text9, text10,
			text11, text12, text13
		};
		object[] array2 = obj;
		bool[] obj2 = new bool[13]
		{
			true, true, true, true, true, true, true, true, true, true,
			true, true, true
		};
		bool[] array3 = obj2;
		NewLateBinding.LateCall(staadApp, null, "GetFullJobInfo", obj, null, null, obj2, IgnoreReturn: true);
		if (array3[0])
		{
			text = (string)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array2[0]), typeof(string));
		}
		if (array3[1])
		{
			text2 = (string)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array2[1]), typeof(string));
		}
		if (array3[2])
		{
			text3 = (string)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array2[2]), typeof(string));
		}
		if (array3[3])
		{
			text4 = (string)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array2[3]), typeof(string));
		}
		if (array3[4])
		{
			text5 = (string)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array2[4]), typeof(string));
		}
		if (array3[5])
		{
			text6 = (string)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array2[5]), typeof(string));
		}
		if (array3[6])
		{
			text7 = (string)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array2[6]), typeof(string));
		}
		if (array3[7])
		{
			text8 = (string)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array2[7]), typeof(string));
		}
		if (array3[8])
		{
			text9 = (string)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array2[8]), typeof(string));
		}
		if (array3[9])
		{
			text10 = (string)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array2[9]), typeof(string));
		}
		if (array3[10])
		{
			text11 = (string)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array2[10]), typeof(string));
		}
		if (array3[11])
		{
			text12 = (string)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array2[11]), typeof(string));
		}
		if (array3[12])
		{
			text13 = (string)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array2[12]), typeof(string));
		}
		object instance = ws;
		NewLateBinding.LateSetComplex(NewLateBinding.LateGet(instance, null, "Range", new object[1] { "C2" }, null, null, null), null, "Value", new object[1] { "JOB NAME:-" }, null, null, OptimisticSet: false, RValueBase: true);
		NewLateBinding.LateCall(NewLateBinding.LateGet(instance, null, "Range", new object[1] { "C2:D2" }, null, null, null), null, "Merge", new object[0], null, null, null, IgnoreReturn: true);
		NewLateBinding.LateSetComplex(NewLateBinding.LateGet(instance, null, "Range", new object[1] { "E2" }, null, null, null), null, "Value", new object[1] { text }, null, null, OptimisticSet: false, RValueBase: true);
		NewLateBinding.LateCall(NewLateBinding.LateGet(instance, null, "Range", new object[1] { "E2:AQ2" }, null, null, null), null, "Merge", new object[0], null, null, null, IgnoreReturn: true);
		NewLateBinding.LateSetComplex(NewLateBinding.LateGet(instance, null, "Range", new object[1] { "C3" }, null, null, null), null, "Value", new object[1] { "JOB NUMBER:-" }, null, null, OptimisticSet: false, RValueBase: true);
		NewLateBinding.LateCall(NewLateBinding.LateGet(instance, null, "Range", new object[1] { "C3:D3" }, null, null, null), null, "Merge", new object[0], null, null, null, IgnoreReturn: true);
		NewLateBinding.LateSetComplex(NewLateBinding.LateGet(instance, null, "Range", new object[1] { "E3" }, null, null, null), null, "Value", new object[1] { text5 }, null, null, OptimisticSet: false, RValueBase: true);
		NewLateBinding.LateCall(NewLateBinding.LateGet(instance, null, "Range", new object[1] { "E3:AQ3" }, null, null, null), null, "Merge", new object[0], null, null, null, IgnoreReturn: true);
		NewLateBinding.LateSetComplex(NewLateBinding.LateGet(instance, null, "Range", new object[1] { "C4" }, null, null, null), null, "Value", new object[1] { "ENGINEER NAME:-" }, null, null, OptimisticSet: false, RValueBase: true);
		NewLateBinding.LateCall(NewLateBinding.LateGet(instance, null, "Range", new object[1] { "C4:D4" }, null, null, null), null, "Merge", new object[0], null, null, null, IgnoreReturn: true);
		NewLateBinding.LateSetComplex(NewLateBinding.LateGet(instance, null, "Range", new object[1] { "E4" }, null, null, null), null, "Value", new object[1] { text3 }, null, null, OptimisticSet: false, RValueBase: true);
		NewLateBinding.LateSetComplex(NewLateBinding.LateGet(instance, null, "Range", new object[1] { "AN4" }, null, null, null), null, "Value", new object[1] { "APPROVER NAME:-" }, null, null, OptimisticSet: false, RValueBase: true);
		NewLateBinding.LateCall(NewLateBinding.LateGet(instance, null, "Range", new object[1] { "AN4:AO4" }, null, null, null), null, "Merge", new object[0], null, null, null, IgnoreReturn: true);
		NewLateBinding.LateSetComplex(NewLateBinding.LateGet(instance, null, "Range", new object[1] { "AP4" }, null, null, null), null, "Value", new object[1] { text11 }, null, null, OptimisticSet: false, RValueBase: true);
		NewLateBinding.LateCall(NewLateBinding.LateGet(instance, null, "Range", new object[1] { "AP4:AQ4" }, null, null, null), null, "Merge", new object[0], null, null, null, IgnoreReturn: true);
		NewLateBinding.LateSetComplex(NewLateBinding.LateGet(instance, null, "Range", new object[1] { "AN5" }, null, null, null), null, "Value", new object[1] { "REVISION:-" }, null, null, OptimisticSet: false, RValueBase: true);
		NewLateBinding.LateSetComplex(NewLateBinding.LateGet(instance, null, "Range", new object[1] { "AO5" }, null, null, null), null, "Value", new object[1] { text6 }, null, null, OptimisticSet: false, RValueBase: true);
		NewLateBinding.LateSetComplex(NewLateBinding.LateGet(instance, null, "Range", new object[1] { "AL4" }, null, null, null), null, "Value", new object[1] { "DATE:-" }, null, null, OptimisticSet: false, RValueBase: true);
		NewLateBinding.LateSetComplex(NewLateBinding.LateGet(instance, null, "Range", new object[1] { "C5" }, null, null, null), null, "Value", new object[1] { "PART:-" }, null, null, OptimisticSet: false, RValueBase: true);
		NewLateBinding.LateCall(NewLateBinding.LateGet(instance, null, "Range", new object[1] { "C5:D5" }, null, null, null), null, "Merge", new object[0], null, null, null, IgnoreReturn: true);
		NewLateBinding.LateSetComplex(NewLateBinding.LateGet(instance, null, "Range", new object[1] { "E5" }, null, null, null), null, "Value", new object[1] { text7 }, null, null, OptimisticSet: false, RValueBase: true);
		NewLateBinding.LateCall(NewLateBinding.LateGet(instance, null, "Range", new object[1] { "E5:AM5" }, null, null, null), null, "Merge", new object[0], null, null, null, IgnoreReturn: true);
		NewLateBinding.LateSetComplex(NewLateBinding.LateGet(instance, null, "Range", new object[1] { "AR3" }, null, null, null), null, "Value", new object[1] { "PLACEMENT:-" }, null, null, OptimisticSet: false, RValueBase: true);
		NewLateBinding.LateSetComplex(NewLateBinding.LateGet(instance, null, "Range", new object[1] { "AR4" }, null, null, null), null, "Value", new object[1] { "TEXT HEIGHT:-" }, null, null, OptimisticSet: false, RValueBase: true);
		NewLateBinding.LateSetComplex(NewLateBinding.LateGet(instance, null, "Range", new object[1] { "AR5" }, null, null, null), null, "Value", new object[1] { "TEXT COLOR:-" }, null, null, OptimisticSet: false, RValueBase: true);
		NewLateBinding.LateSetComplex(NewLateBinding.LateGet(NewLateBinding.LateGet(instance, null, "Range", new object[1] { "A1:AS7" }, null, null, null), null, "Font", new object[0], null, null, null), null, "Bold", new object[1] { true }, null, null, OptimisticSet: false, RValueBase: true);
		NewLateBinding.LateSetComplex(NewLateBinding.LateGet(NewLateBinding.LateGet(instance, null, "Range", new object[1] { "A1:AS1,A6:AR6" }, null, null, null), null, "Interior", new object[0], null, null, null), null, "Color", new object[1] { Color.FromArgb(248, 203, 173) }, null, null, OptimisticSet: false, RValueBase: true);
		NewLateBinding.LateSetComplex(NewLateBinding.LateGet(NewLateBinding.LateGet(instance, null, "Range", new object[1] { "AS1,AQ5" }, null, null, null), null, "Interior", new object[0], null, null, null), null, "Color", new object[1] { Color.FromArgb(146, 205, 220) }, null, null, OptimisticSet: false, RValueBase: true);
		NewLateBinding.LateSetComplex(NewLateBinding.LateGet(NewLateBinding.LateGet(instance, null, "Range", new object[1] { "A1:AQ5" }, null, null, null), null, "Font", new object[0], null, null, null), null, "Color", new object[1] { Color.FromArgb(0, 32, 96) }, null, null, OptimisticSet: false, RValueBase: true);
		NewLateBinding.LateSetComplex(NewLateBinding.LateGet(NewLateBinding.LateGet(instance, null, "Range", new object[1] { "AR2:AS2" }, null, null, null), null, "Font", new object[0], null, null, null), null, "Color", new object[1] { Color.FromArgb(192, 0, 0) }, null, null, OptimisticSet: false, RValueBase: true);
		NewLateBinding.LateSetComplex(NewLateBinding.LateGet(NewLateBinding.LateGet(instance, null, "Range", new object[1] { "A1:AQ6,AR2:AS6" }, null, null, null), null, "Borders", new object[0], null, null, null), null, "LineStyle", new object[1] { 1 }, null, null, OptimisticSet: false, RValueBase: true);
		NewLateBinding.LateCall(NewLateBinding.LateGet(instance, null, "Range", new object[1] { "A1:AQ1,A6:AQ6,AR6:AS6" }, null, null, null), null, "Merge", new object[0], null, null, null, IgnoreReturn: true);
		NewLateBinding.LateCall(NewLateBinding.LateGet(instance, null, "Range", new object[1] { "A2:B5" }, null, null, null), null, "Merge", new object[0], null, null, null, IgnoreReturn: true);
		NewLateBinding.LateSetComplex(NewLateBinding.LateGet(NewLateBinding.LateGet(instance, null, "Range", new object[1] { "A1:AQ6" }, null, null, null), null, "Font", new object[0], null, null, null), null, "Size", new object[1] { 12 }, null, null, OptimisticSet: false, RValueBase: true);
		NewLateBinding.LateSetComplex(NewLateBinding.LateGet(instance, null, "Range", new object[1] { "A1" }, null, null, null), null, "Value", new object[1] { text14 }, null, null, OptimisticSet: false, RValueBase: true);
		NewLateBinding.LateSetComplex(NewLateBinding.LateGet(NewLateBinding.LateGet(instance, null, "Range", new object[1] { "A1" }, null, null, null), null, "Font", new object[0], null, null, null), null, "Color", new object[1] { Color.FromArgb(192, 0, 0) }, null, null, OptimisticSet: false, RValueBase: true);
		NewLateBinding.LateSetComplex(NewLateBinding.LateGet(instance, null, "Range", new object[1] { "A1,AN3,AN4,AN5,AQ3,AQ4,AR7,AS7" }, null, null, null), null, "HorizontalAlignment", new object[1] { -4108 }, null, null, OptimisticSet: false, RValueBase: true);
		NewLateBinding.LateSetComplex(NewLateBinding.LateGet(instance, null, "Range", new object[1] { "E3,K5" }, null, null, null), null, "HorizontalAlignment", new object[1] { -4131 }, null, null, OptimisticSet: false, RValueBase: true);
		NewLateBinding.LateSetComplex(NewLateBinding.LateGet(instance, null, "Range", new object[1] { "AR1" }, null, null, null), null, "Value", new object[1] { "CREATED BY:-" }, null, null, OptimisticSet: false, RValueBase: true);
		NewLateBinding.LateSetComplex(NewLateBinding.LateGet(instance, null, "Range", new object[1] { "AS1" }, null, null, null), null, "Value", new object[1] { "PEEYUSH" }, null, null, OptimisticSet: false, RValueBase: true);
		instance = null;
	}

	private void SetupExcelHeaders(object ws)
	{
		object obj = ws;
		NewLateBinding.LateSetComplex(NewLateBinding.LateGet(obj, null, "Range", new object[1] { "A7" }, null, null, null), null, "Value", new object[1] { "S.No" }, null, null, OptimisticSet: false, RValueBase: true);
		NewLateBinding.LateSetComplex(NewLateBinding.LateGet(obj, null, "Range", new object[1] { "B7" }, null, null, null), null, "Value", new object[1] { "BEAM No" }, null, null, OptimisticSet: false, RValueBase: true);
		NewLateBinding.LateSetComplex(NewLateBinding.LateGet(obj, null, "Range", new object[1] { "C7" }, null, null, null), null, "Value", new object[1] { "LENGTH(M)" }, null, null, OptimisticSet: false, RValueBase: true);
		NewLateBinding.LateSetComplex(NewLateBinding.LateGet(obj, null, "Range", new object[1] { "D7" }, null, null, null), null, "Value", new object[1] { "SECTION" }, null, null, OptimisticSet: false, RValueBase: true);
		NewLateBinding.LateSetComplex(NewLateBinding.LateGet(obj, null, "Range", new object[1] { "E7" }, null, null, null), null, "Value", new object[1] { "PROPERTY" }, null, null, OptimisticSet: false, RValueBase: true);
		NewLateBinding.LateSetComplex(NewLateBinding.LateGet(obj, null, "Range", new object[1] { "F7" }, null, null, null), null, "Value", new object[1] { "TYPE" }, null, null, OptimisticSet: false, RValueBase: true);
		NewLateBinding.LateSetComplex(NewLateBinding.LateGet(obj, null, "Range", new object[1] { "G7" }, null, null, null), null, "Value", new object[1] { "F1(DEPTH START)" }, null, null, OptimisticSet: false, RValueBase: true);
		NewLateBinding.LateSetComplex(NewLateBinding.LateGet(obj, null, "Range", new object[1] { "H7" }, null, null, null), null, "Value", new object[1] { "F2(WEB THK.)" }, null, null, OptimisticSet: false, RValueBase: true);
		NewLateBinding.LateSetComplex(NewLateBinding.LateGet(obj, null, "Range", new object[1] { "I7" }, null, null, null), null, "Value", new object[1] { "F3(DEPTH END)" }, null, null, OptimisticSet: false, RValueBase: true);
		NewLateBinding.LateSetComplex(NewLateBinding.LateGet(obj, null, "Range", new object[1] { "J7" }, null, null, null), null, "Value", new object[1] { "F4(WIDTH Tf)" }, null, null, OptimisticSet: false, RValueBase: true);
		NewLateBinding.LateSetComplex(NewLateBinding.LateGet(obj, null, "Range", new object[1] { "K7" }, null, null, null), null, "Value", new object[1] { "F5(THK. Tf)" }, null, null, OptimisticSet: false, RValueBase: true);
		NewLateBinding.LateSetComplex(NewLateBinding.LateGet(obj, null, "Range", new object[1] { "L7" }, null, null, null), null, "Value", new object[1] { "F6(WIDTH Bf)" }, null, null, OptimisticSet: false, RValueBase: true);
		NewLateBinding.LateSetComplex(NewLateBinding.LateGet(obj, null, "Range", new object[1] { "M7" }, null, null, null), null, "Value", new object[1] { "F7(THK.Tb)" }, null, null, OptimisticSet: false, RValueBase: true);
		NewLateBinding.LateSetComplex(NewLateBinding.LateGet(obj, null, "Range", new object[1] { "N7" }, null, null, null), null, "Value", new object[1] { "Node A" }, null, null, OptimisticSet: false, RValueBase: true);
		NewLateBinding.LateSetComplex(NewLateBinding.LateGet(obj, null, "Range", new object[1] { "O7" }, null, null, null), null, "Value", new object[1] { "Node B" }, null, null, OptimisticSet: false, RValueBase: true);
		NewLateBinding.LateSetComplex(NewLateBinding.LateGet(obj, null, "Range", new object[1] { "P7" }, null, null, null), null, "Value", new object[1] { "COORDINATE A" }, null, null, OptimisticSet: false, RValueBase: true);
		NewLateBinding.LateSetComplex(NewLateBinding.LateGet(obj, null, "Range", new object[1] { "S7" }, null, null, null), null, "Value", new object[1] { "COORDINATE B" }, null, null, OptimisticSet: false, RValueBase: true);
		NewLateBinding.LateSetComplex(NewLateBinding.LateGet(obj, null, "Range", new object[1] { "V7" }, null, null, null), null, "Value", new object[1] { "Beta Angle (°)" }, null, null, OptimisticSet: false, RValueBase: true);
		NewLateBinding.LateSetComplex(NewLateBinding.LateGet(obj, null, "Range", new object[1] { "AK7" }, null, null, null), null, "Value", new object[1] { "SECTION FARMING" }, null, null, OptimisticSet: false, RValueBase: true);
		NewLateBinding.LateSetComplex(NewLateBinding.LateGet(obj, null, "Range", new object[1] { "AN7" }, null, null, null), null, "Value", new object[1] { "WEIGHT(Kg)" }, null, null, OptimisticSet: false, RValueBase: true);
		NewLateBinding.LateSetComplex(NewLateBinding.LateGet(obj, null, "Range", new object[1] { "P8" }, null, null, null), null, "Value", new object[1] { "X" }, null, null, OptimisticSet: false, RValueBase: true);
		NewLateBinding.LateSetComplex(NewLateBinding.LateGet(obj, null, "Range", new object[1] { "Q8" }, null, null, null), null, "Value", new object[1] { "Y" }, null, null, OptimisticSet: false, RValueBase: true);
		NewLateBinding.LateSetComplex(NewLateBinding.LateGet(obj, null, "Range", new object[1] { "R8" }, null, null, null), null, "Value", new object[1] { "Z" }, null, null, OptimisticSet: false, RValueBase: true);
		NewLateBinding.LateSetComplex(NewLateBinding.LateGet(obj, null, "Range", new object[1] { "S8" }, null, null, null), null, "Value", new object[1] { "X" }, null, null, OptimisticSet: false, RValueBase: true);
		NewLateBinding.LateSetComplex(NewLateBinding.LateGet(obj, null, "Range", new object[1] { "T8" }, null, null, null), null, "Value", new object[1] { "Y" }, null, null, OptimisticSet: false, RValueBase: true);
		NewLateBinding.LateSetComplex(NewLateBinding.LateGet(obj, null, "Range", new object[1] { "U8" }, null, null, null), null, "Value", new object[1] { "Z" }, null, null, OptimisticSet: false, RValueBase: true);
		NewLateBinding.LateSetComplex(NewLateBinding.LateGet(obj, null, "Range", new object[1] { "AK8" }, null, null, null), null, "Value", new object[1] { "WEB" }, null, null, OptimisticSet: false, RValueBase: true);
		NewLateBinding.LateSetComplex(NewLateBinding.LateGet(obj, null, "Range", new object[1] { "AL8" }, null, null, null), null, "Value", new object[1] { "TOP FLANGE" }, null, null, OptimisticSet: false, RValueBase: true);
		NewLateBinding.LateSetComplex(NewLateBinding.LateGet(obj, null, "Range", new object[1] { "AM8" }, null, null, null), null, "Value", new object[1] { "BOTT.FLANGE" }, null, null, OptimisticSet: false, RValueBase: true);
		NewLateBinding.LateSetComplex(NewLateBinding.LateGet(obj, null, "Range", new object[1] { "AN8" }, null, null, null), null, "Value", new object[1] { "WEB(Kg)" }, null, null, OptimisticSet: false, RValueBase: true);
		NewLateBinding.LateSetComplex(NewLateBinding.LateGet(obj, null, "Range", new object[1] { "AO8" }, null, null, null), null, "Value", new object[1] { "TOP FLANGE(Kg)" }, null, null, OptimisticSet: false, RValueBase: true);
		NewLateBinding.LateSetComplex(NewLateBinding.LateGet(obj, null, "Range", new object[1] { "AP8" }, null, null, null), null, "Value", new object[1] { "BOTT.FLANGE(Kg)" }, null, null, OptimisticSet: false, RValueBase: true);
		NewLateBinding.LateSetComplex(NewLateBinding.LateGet(obj, null, "Range", new object[1] { "AQ8" }, null, null, null), null, "Value", new object[1] { "TOTAL(Kg)" }, null, null, OptimisticSet: false, RValueBase: true);
		NewLateBinding.LateCall(NewLateBinding.LateGet(obj, null, "Range", new object[1] { "P7:R7" }, null, null, null), null, "Merge", new object[0], null, null, null, IgnoreReturn: true);
		NewLateBinding.LateCall(NewLateBinding.LateGet(obj, null, "Range", new object[1] { "S7:U7" }, null, null, null), null, "Merge", new object[0], null, null, null, IgnoreReturn: true);
		NewLateBinding.LateCall(NewLateBinding.LateGet(obj, null, "Range", new object[1] { "V7:V8" }, null, null, null), null, "Merge", new object[0], null, null, null, IgnoreReturn: true);
		NewLateBinding.LateCall(NewLateBinding.LateGet(obj, null, "Range", new object[1] { "AK7:AM7" }, null, null, null), null, "Merge", new object[0], null, null, null, IgnoreReturn: true);
		NewLateBinding.LateCall(NewLateBinding.LateGet(obj, null, "Range", new object[1] { "AN7:AQ7" }, null, null, null), null, "Merge", new object[0], null, null, null, IgnoreReturn: true);
		object instance = NewLateBinding.LateGet(obj, null, "Range", new object[1] { "AK8:AQ8" }, null, null, null);
		NewLateBinding.LateSetComplex(NewLateBinding.LateGet(instance, null, "Font", new object[0], null, null, null), null, "Bold", new object[1] { true }, null, null, OptimisticSet: false, RValueBase: true);
		NewLateBinding.LateSetComplex(instance, null, "HorizontalAlignment", new object[1] { -4108 }, null, null, OptimisticSet: false, RValueBase: true);
		NewLateBinding.LateSetComplex(instance, null, "VerticalAlignment", new object[1] { -4108 }, null, null, OptimisticSet: false, RValueBase: true);
		NewLateBinding.LateSetComplex(instance, null, "WrapText", new object[1] { true }, null, null, OptimisticSet: false, RValueBase: true);
		instance = null;
		int num = 1;
		do
		{
			object instance2 = obj;
			object[] array = new object[2];
			object instance3 = obj;
			object[] array2 = new object[2];
			object obj2 = (array2[0] = 7);
			int num2;
			array2[1] = (num2 = num);
			object[] array3 = array2;
			bool[] obj3 = new bool[2] { false, true };
			bool[] array4 = obj3;
			object obj4 = NewLateBinding.LateGet(instance3, null, "Cells", array2, null, null, obj3);
			if (array4[1])
			{
				num = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array3[1]), typeof(int));
			}
			array[0] = obj4;
			object instance4 = obj;
			object[] array5 = new object[2];
			object obj5 = (array5[0] = 8);
			int num3;
			array5[1] = (num3 = num);
			array3 = array5;
			obj4 = NewLateBinding.LateGet(instance4, null, "Cells", array5, null, null, array4 = new bool[2] { false, true });
			if (array4[1])
			{
				num = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array3[1]), typeof(int));
			}
			array[1] = obj4;
			object[] array6 = array;
			bool[] array7;
			object obj6 = NewLateBinding.LateGet(instance2, null, "Range", array, null, null, array7 = new bool[2] { true, true });
			if (array7[0])
			{
				NewLateBinding.LateSetComplex(instance3, null, "Cells", new object[3]
				{
					obj2,
					num2,
					array6[0]
				}, null, null, OptimisticSet: true, RValueBase: false);
			}
			if (array7[1])
			{
				NewLateBinding.LateSetComplex(instance4, null, "Cells", new object[3]
				{
					obj5,
					num3,
					array6[1]
				}, null, null, OptimisticSet: true, RValueBase: false);
			}
			object instance5 = obj6;
			NewLateBinding.LateCall(instance5, null, "Merge", new object[0], null, null, null, IgnoreReturn: true);
			NewLateBinding.LateSetComplex(instance5, null, "WrapText", new object[1] { true }, null, null, OptimisticSet: false, RValueBase: true);
			NewLateBinding.LateSetComplex(instance5, null, "HorizontalAlignment", new object[1] { -4108 }, null, null, OptimisticSet: false, RValueBase: true);
			NewLateBinding.LateSetComplex(instance5, null, "VerticalAlignment", new object[1] { -4108 }, null, null, OptimisticSet: false, RValueBase: true);
			instance5 = null;
			num = checked(num + 1);
		}
		while (num <= 15);
		object instance6 = NewLateBinding.LateGet(obj, null, "Range", new object[1] { "A7:AQ8" }, null, null, null);
		NewLateBinding.LateSetComplex(NewLateBinding.LateGet(instance6, null, "Font", new object[0], null, null, null), null, "Bold", new object[1] { true }, null, null, OptimisticSet: false, RValueBase: true);
		NewLateBinding.LateSetComplex(instance6, null, "HorizontalAlignment", new object[1] { -4108 }, null, null, OptimisticSet: false, RValueBase: true);
		NewLateBinding.LateSetComplex(instance6, null, "VerticalAlignment", new object[1] { -4108 }, null, null, OptimisticSet: false, RValueBase: true);
		instance6 = null;
		NewLateBinding.LateSetComplex(NewLateBinding.LateGet(NewLateBinding.LateGet(obj, null, "Range", new object[1] { "7:8" }, null, null, null), null, "Font", new object[0], null, null, null), null, "Color", new object[1] { Information.RGB(0, 32, 96) }, null, null, OptimisticSet: false, RValueBase: true);
		obj = null;
	}

	private string ProcessBeamData(object ws, object staadApp)
	{
		int num = Conversions.ToInteger(NewLateBinding.LateGet(NewLateBinding.LateGet(staadApp, null, "Geometry", new object[0], null, null, null), null, "GetNoOfSelectedBeams", new object[0], null, null, null));
		if (num == 0)
		{
			return "No members selected in STAAD.";
		}
		checked
		{
			int[] array = new int[num - 1 + 1];
			object[] array2;
			bool[] array3;
			NewLateBinding.LateCall(NewLateBinding.LateGet(staadApp, null, "Geometry", new object[0], null, null, null), null, "GetSelectedBeams", array2 = new object[1] { array }, null, null, array3 = new bool[1] { true }, IgnoreReturn: true);
			if (array3[0])
			{
				array = (int[])Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array2[0]), typeof(int[]));
			}
			int num2 = Conversions.ToInteger(NewLateBinding.LateGet(NewLateBinding.LateGet(staadApp, null, "Property", new object[0], null, null, null), null, "GetCountofSectionPropertyValuesEx", new object[0], null, null, null));
			double[] array4 = new double[num2 - 1 + 1];
			object[,] array5 = new object[num - 1 + 1, 22];
			int num3 = 0;
			int num4 = 0;
			int num5 = 0;
			int[] array6 = array;
			int num9 = default(int);
			int num10 = default(int);
			double num11 = default(double);
			double num12 = default(double);
			double num13 = default(double);
			double num14 = default(double);
			double num15 = default(double);
			double num16 = default(double);
			for (int i = 0; i < array6.Length; i++)
			{
				int num6 = array6[i];
				try
				{
					object instance = NewLateBinding.LateGet(staadApp, null, "Geometry", new object[0], null, null, null);
					object[] obj = new object[1] { num6 };
					array2 = obj;
					bool[] obj2 = new bool[1] { true };
					array3 = obj2;
					object value = NewLateBinding.LateGet(instance, null, "GetBeamLength", obj, null, null, obj2);
					if (array3[0])
					{
						num6 = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array2[0]), typeof(int));
					}
					double value2 = Conversions.ToDouble(value);
					value = NewLateBinding.LateGet(NewLateBinding.LateGet(staadApp, null, "Property", new object[0], null, null, null), null, "GetBeamSectionName", array2 = new object[1] { num6 }, null, null, array3 = new bool[1] { true });
					if (array3[0])
					{
						num6 = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array2[0]), typeof(int));
					}
					string text = Conversions.ToString(value);
					value = NewLateBinding.LateGet(NewLateBinding.LateGet(staadApp, null, "Property", new object[0], null, null, null), null, "GetBeamSectionPropertyRefNo", array2 = new object[1] { num6 }, null, null, array3 = new bool[1] { true });
					if (array3[0])
					{
						num6 = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array2[0]), typeof(int));
					}
					int num7 = Conversions.ToInteger(value);
					value = NewLateBinding.LateGet(NewLateBinding.LateGet(staadApp, null, "Property", new object[0], null, null, null), null, "GetBeamSectionPropertyTypeNo", array2 = new object[1] { num6 }, null, null, array3 = new bool[1] { true });
					if (array3[0])
					{
						num6 = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array2[0]), typeof(int));
					}
					int num8 = Conversions.ToInteger(value);
					value = NewLateBinding.LateGet(NewLateBinding.LateGet(staadApp, null, "Property", new object[0], null, null, null), null, "GetBetaAngle", array2 = new object[1] { num6 }, null, null, array3 = new bool[1] { true });
					if (array3[0])
					{
						num6 = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array2[0]), typeof(int));
					}
					double value3 = Conversions.ToDouble(value);
					NewLateBinding.LateCall(NewLateBinding.LateGet(staadApp, null, "Geometry", new object[0], null, null, null), null, "GetMemberIncidence", array2 = new object[3] { num6, num9, num10 }, null, null, array3 = new bool[3] { true, true, true }, IgnoreReturn: true);
					if (array3[0])
					{
						num6 = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array2[0]), typeof(int));
					}
					if (array3[1])
					{
						num9 = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array2[1]), typeof(int));
					}
					if (array3[2])
					{
						num10 = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array2[2]), typeof(int));
					}
					NewLateBinding.LateCall(NewLateBinding.LateGet(staadApp, null, "Geometry", new object[0], null, null, null), null, "GetNodeCoordinates", array2 = new object[4] { num9, num11, num12, num13 }, null, null, array3 = new bool[4] { true, true, true, true }, IgnoreReturn: true);
					if (array3[0])
					{
						num9 = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array2[0]), typeof(int));
					}
					if (array3[1])
					{
						num11 = (double)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array2[1]), typeof(double));
					}
					if (array3[2])
					{
						num12 = (double)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array2[2]), typeof(double));
					}
					if (array3[3])
					{
						num13 = (double)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array2[3]), typeof(double));
					}
					NewLateBinding.LateCall(NewLateBinding.LateGet(staadApp, null, "Geometry", new object[0], null, null, null), null, "GetNodeCoordinates", array2 = new object[4] { num10, num14, num15, num16 }, null, null, array3 = new bool[4] { true, true, true, true }, IgnoreReturn: true);
					if (array3[0])
					{
						num10 = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array2[0]), typeof(int));
					}
					if (array3[1])
					{
						num14 = (double)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array2[1]), typeof(double));
					}
					if (array3[2])
					{
						num15 = (double)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array2[2]), typeof(double));
					}
					if (array3[3])
					{
						num16 = (double)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array2[3]), typeof(double));
					}
					NewLateBinding.LateCall(NewLateBinding.LateGet(staadApp, null, "Property", new object[0], null, null, null), null, "GetSectionPropertyValuesEx", array2 = new object[3] { num7, num8, array4 }, null, null, array3 = new bool[3] { true, true, true }, IgnoreReturn: true);
					if (array3[0])
					{
						num7 = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array2[0]), typeof(int));
					}
					if (array3[1])
					{
						num8 = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array2[1]), typeof(int));
					}
					if (array3[2])
					{
						array4 = (double[])Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array2[2]), typeof(double[]));
					}
					if (array4.Length > num5)
					{
						num5 = array4.Length;
					}
					array5[num3, 0] = num3 + 1;
					array5[num3, 1] = num6;
					array5[num3, 2] = Math.Round(value2, 3);
					array5[num3, 3] = text;
					array5[num3, 4] = "R" + Conversions.ToString(num7);
					array5[num3, 5] = num8;
					int num17 = Math.Min(6, array4.Length - 1);
					for (int j = 0; j <= num17; j++)
					{
						if (num8 == 680)
						{
							array5[num3, 6 + j] = Math.Round(array4[j], 3);
						}
						else
						{
							array5[num3, 6 + j] = Math.Round(array4[j], 6);
						}
					}
					array5[num3, 13] = num9;
					array5[num3, 14] = num10;
					array5[num3, 15] = Math.Round(num11, 3);
					array5[num3, 16] = Math.Round(num12, 3);
					array5[num3, 17] = Math.Round(num13, 3);
					array5[num3, 18] = Math.Round(num14, 3);
					array5[num3, 19] = Math.Round(num15, 3);
					array5[num3, 20] = Math.Round(num16, 3);
					array5[num3, 21] = Math.Round(value3, 2);
					if (array4.Length > 7)
					{
						int num18 = Math.Min(11, array4.Length - 1);
						for (int k = 7; k <= num18; k++)
						{
							int num19 = 51 + (k - 7);
							if (num8 == 680)
							{
								object[] obj3 = new object[2]
								{
									9 + num3,
									num19
								};
								array2 = obj3;
								bool[] obj4 = new bool[2] { false, true };
								array3 = obj4;
								value = NewLateBinding.LateGet(ws, null, "Cells", obj3, null, null, obj4);
								if (array3[1])
								{
									num19 = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array2[1]), typeof(int));
								}
								NewLateBinding.LateSetComplex(value, null, "Value", new object[1] { Math.Round(array4[k], 3) }, null, null, OptimisticSet: false, RValueBase: true);
							}
							else
							{
								object[] obj5 = new object[2]
								{
									9 + num3,
									num19
								};
								array2 = obj5;
								bool[] obj6 = new bool[2] { false, true };
								array3 = obj6;
								value = NewLateBinding.LateGet(ws, null, "Cells", obj5, null, null, obj6);
								if (array3[1])
								{
									num19 = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array2[1]), typeof(int));
								}
								NewLateBinding.LateSetComplex(value, null, "Value", new object[1] { Math.Round(array4[k], 6) }, null, null, OptimisticSet: false, RValueBase: true);
							}
						}
					}
					num3++;
					num4++;
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					Debug.WriteLine($"Error processing beam {num6}: {ex2.Message}");
					ProjectData.ClearProjectError();
				}
			}
			if (num4 > 0)
			{
				object[] array7 = new object[2];
				object value = ws;
				object instance2 = value;
				object[] array8 = new object[2];
				object obj7 = (array8[0] = 9);
				object obj8 = (array8[1] = 1);
				array7[0] = NewLateBinding.LateGet(instance2, null, "Cells", array8, null, null, null);
				object[] array9 = new object[2];
				object obj9 = (array9[0] = 8 + num4);
				object obj10 = (array9[1] = 22);
				array7[1] = NewLateBinding.LateGet(ws, null, "Cells", array9, null, null, null);
				array2 = array7;
				bool[] obj11 = new bool[2] { true, true };
				array3 = obj11;
				object instance3 = NewLateBinding.LateGet(ws, null, "Range", array7, null, null, obj11);
				if (array3[0])
				{
					NewLateBinding.LateSetComplex(value, null, "Cells", new object[3]
					{
						obj7,
						obj8,
						array2[0]
					}, null, null, OptimisticSet: true, RValueBase: false);
				}
				if (array3[1])
				{
					NewLateBinding.LateSetComplex(ws, null, "Cells", new object[3]
					{
						obj9,
						obj10,
						array2[1]
					}, null, null, OptimisticSet: true, RValueBase: false);
				}
				NewLateBinding.LateSetComplex(instance3, null, "Value", new object[1] { array5 }, null, null, OptimisticSet: false, RValueBase: true);
				int num20 = 8 + num4;
				for (int l = 9; l <= num20; l++)
				{
					obj10 = NewLateBinding.LateGet(ws, null, "Cells", array2 = new object[2] { l, 6 }, null, null, array3 = new bool[2] { true, false });
					if (array3[0])
					{
						l = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array2[0]), typeof(int));
					}
					if (Conversion.Val(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(obj10, null, "Value", new object[0], null, null, null))) != 680.0)
					{
						continue;
					}
					object[] obj12 = new object[2] { l, 12 };
					array2 = obj12;
					bool[] obj13 = new bool[2] { true, false };
					array3 = obj13;
					obj10 = NewLateBinding.LateGet(ws, null, "Cells", obj12, null, null, obj13);
					if (array3[0])
					{
						l = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array2[0]), typeof(int));
					}
					if (Conversion.Val(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(obj10, null, "Value", new object[0], null, null, null))) != 0.0)
					{
						continue;
					}
					object[] obj14 = new object[2] { l, 13 };
					array2 = obj14;
					bool[] obj15 = new bool[2] { true, false };
					array3 = obj15;
					obj10 = NewLateBinding.LateGet(ws, null, "Cells", obj14, null, null, obj15);
					if (array3[0])
					{
						l = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array2[0]), typeof(int));
					}
					if (Conversion.Val(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(obj10, null, "Value", new object[0], null, null, null))) == 0.0)
					{
						object[] obj16 = new object[2] { l, 12 };
						array2 = obj16;
						bool[] obj17 = new bool[2] { true, false };
						array3 = obj17;
						obj10 = NewLateBinding.LateGet(ws, null, "Cells", obj16, null, null, obj17);
						if (array3[0])
						{
							l = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array2[0]), typeof(int));
						}
						object instance4 = obj10;
						object[] array10 = new object[1];
						obj10 = NewLateBinding.LateGet(ws, null, "Cells", array2 = new object[2] { l, 10 }, null, null, array3 = new bool[2] { true, false });
						if (array3[0])
						{
							l = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array2[0]), typeof(int));
						}
						array10[0] = NewLateBinding.LateGet(obj10, null, "Value", new object[0], null, null, null);
						NewLateBinding.LateSetComplex(instance4, null, "Value", array10, null, null, OptimisticSet: false, RValueBase: true);
						obj10 = NewLateBinding.LateGet(ws, null, "Cells", array2 = new object[2] { l, 13 }, null, null, array3 = new bool[2] { true, false });
						if (array3[0])
						{
							l = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array2[0]), typeof(int));
						}
						object instance5 = obj10;
						object[] array11 = new object[1];
						obj10 = NewLateBinding.LateGet(ws, null, "Cells", array2 = new object[2] { l, 11 }, null, null, array3 = new bool[2] { true, false });
						if (array3[0])
						{
							l = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array2[0]), typeof(int));
						}
						array11[0] = NewLateBinding.LateGet(obj10, null, "Value", new object[0], null, null, null);
						NewLateBinding.LateSetComplex(instance5, null, "Value", array11, null, null, OptimisticSet: false, RValueBase: true);
					}
				}
			}
			return $"{num4} of {num} beams processed successfully.";
		}
	}

	private void FormatExcelWorksheet(object ws)
	{
		double[] array = new double[39]
		{
			6.0, 10.0, 12.0, 15.0, 10.0, 8.0, 12.0, 12.0, 12.0, 12.0,
			12.0, 12.0, 12.0, 10.0, 10.0, 10.0, 10.0, 10.0, 10.0, 10.0,
			10.0, 10.0, 10.0, 14.0, 14.0, 14.0, 14.0, 14.0, 14.0, 14.0,
			14.0, 14.0, 14.0, 14.0, 14.0, 14.0, 14.0, 14.0, 14.0
		};
		checked
		{
			int num = array.Length - 1;
			for (int i = 0; i <= num; i++)
			{
				NewLateBinding.LateSetComplex(NewLateBinding.LateGet(ws, null, "Columns", new object[1] { i + 1 }, null, null, null), null, "ColumnWidth", new object[1] { array[i] }, null, null, OptimisticSet: false, RValueBase: true);
			}
			object instance;
			object[] obj = new object[2]
			{
				NewLateBinding.LateGet(instance = NewLateBinding.LateGet(ws, null, "Rows", new object[0], null, null, null), null, "Count", new object[0], null, null, null),
				1
			};
			object[] array2 = obj;
			bool[] obj2 = new bool[2] { true, false };
			bool[] array3 = obj2;
			object instance2 = NewLateBinding.LateGet(ws, null, "Cells", obj, null, null, obj2);
			if (array3[0])
			{
				NewLateBinding.LateSetComplex(instance, null, "Count", new object[1] { array2[0] }, null, null, OptimisticSet: true, RValueBase: true);
			}
			int num2 = Conversions.ToInteger(NewLateBinding.LateGet(NewLateBinding.LateGet(instance2, null, "End", new object[1] { -4162 }, null, null, null), null, "Row", new object[0], null, null, null));
			if (num2 >= 9)
			{
				object[] array4 = new object[2];
				instance = ws;
				object instance3 = instance;
				object[] array5 = new object[2];
				instance2 = (array5[0] = 7);
				object obj3 = (array5[1] = 1);
				array4[0] = NewLateBinding.LateGet(instance3, null, "Cells", array5, null, null, null);
				int num3;
				object[] obj4 = new object[2]
				{
					num3 = num2,
					null
				};
				object obj5 = (obj4[1] = 22);
				object[] array6 = obj4;
				bool[] obj6 = new bool[2] { true, false };
				bool[] array7 = obj6;
				object obj7 = NewLateBinding.LateGet(ws, null, "Cells", obj4, null, null, obj6);
				if (array7[0])
				{
					num2 = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array6[0]), typeof(int));
				}
				array4[1] = obj7;
				array2 = array4;
				object instance4 = NewLateBinding.LateGet(ws, null, "Range", array4, null, null, array3 = new bool[2] { true, true });
				if (array3[0])
				{
					NewLateBinding.LateSetComplex(instance, null, "Cells", new object[3]
					{
						instance2,
						obj3,
						array2[0]
					}, null, null, OptimisticSet: true, RValueBase: false);
				}
				if (array3[1])
				{
					NewLateBinding.LateSetComplex(ws, null, "Cells", new object[3]
					{
						num3,
						obj5,
						array2[1]
					}, null, null, OptimisticSet: true, RValueBase: false);
				}
				object instance5 = NewLateBinding.LateGet(instance4, null, "Borders", new object[0], null, null, null);
				NewLateBinding.LateSetComplex(instance5, null, "LineStyle", new object[1] { 1 }, null, null, OptimisticSet: false, RValueBase: true);
				NewLateBinding.LateSetComplex(instance5, null, "Weight", new object[1] { 2 }, null, null, OptimisticSet: false, RValueBase: true);
				instance5 = null;
			}
			object instance6 = NewLateBinding.LateGet(ws, null, "Range", new object[1] { "AS2:AS6" }, null, null, null);
			NewLateBinding.LateSetComplex(instance6, null, "HorizontalAlignment", new object[1] { -4108 }, null, null, OptimisticSet: false, RValueBase: true);
			NewLateBinding.LateSetComplex(instance6, null, "VerticalAlignment", new object[1] { -4108 }, null, null, OptimisticSet: false, RValueBase: true);
			instance6 = null;
			NewLateBinding.LateCall(NewLateBinding.LateGet(ws, null, "Columns", new object[1] { "AS" }, null, null, null), null, "AutoFit", new object[0], null, null, null, IgnoreReturn: true);
			object instance7 = NewLateBinding.LateGet(NewLateBinding.LateGet(ws, null, "Range", new object[1] { "AR1:AS6" }, null, null, null), null, "Borders", new object[0], null, null, null);
			NewLateBinding.LateSetComplex(instance7, null, "LineStyle", new object[1] { 1 }, null, null, OptimisticSet: false, RValueBase: true);
			NewLateBinding.LateSetComplex(instance7, null, "Weight", new object[1] { 2 }, null, null, OptimisticSet: false, RValueBase: true);
			instance7 = null;
		}
	}

	private void HideColumns(object ws)
	{
		NewLateBinding.LateSetComplex(NewLateBinding.LateGet(NewLateBinding.LateGet(ws, null, "Range", new object[1] { "A:E" }, null, null, null), null, "EntireColumn", new object[0], null, null, null), null, "Hidden", new object[1] { false }, null, null, OptimisticSet: false, RValueBase: true);
		NewLateBinding.LateSetComplex(NewLateBinding.LateGet(NewLateBinding.LateGet(ws, null, "Range", new object[1] { "F:AJ" }, null, null, null), null, "EntireColumn", new object[0], null, null, null), null, "Hidden", new object[1] { true }, null, null, OptimisticSet: false, RValueBase: true);
		NewLateBinding.LateSetComplex(NewLateBinding.LateGet(NewLateBinding.LateGet(ws, null, "Range", new object[1] { "AK:AW" }, null, null, null), null, "EntireColumn", new object[0], null, null, null), null, "Hidden", new object[1] { false }, null, null, OptimisticSet: false, RValueBase: true);
		NewLateBinding.LateSetComplex(NewLateBinding.LateGet(NewLateBinding.LateGet(ws, null, "Range", new object[1] { "AY:BC" }, null, null, null), null, "EntireColumn", new object[0], null, null, null), null, "Hidden", new object[1] { true }, null, null, OptimisticSet: false, RValueBase: true);
	}

	private void CalculateTotalWeight(object ws)
	{
		object instance;
		object[] obj = new object[2]
		{
			NewLateBinding.LateGet(instance = NewLateBinding.LateGet(ws, null, "Rows", new object[0], null, null, null), null, "Count", new object[0], null, null, null),
			"AQ"
		};
		object[] array = obj;
		bool[] obj2 = new bool[2] { true, false };
		bool[] array2 = obj2;
		object instance2 = NewLateBinding.LateGet(ws, null, "Cells", obj, null, null, obj2);
		if (array2[0])
		{
			NewLateBinding.LateSetComplex(instance, null, "Count", new object[1] { array[0] }, null, null, OptimisticSet: true, RValueBase: true);
		}
		int num = Conversions.ToInteger(NewLateBinding.LateGet(NewLateBinding.LateGet(instance2, null, "End", new object[1] { -4162 }, null, null, null), null, "Row", new object[0], null, null, null));
		double value = 0.0;
		if (num >= 9)
		{
			object instance3 = NewLateBinding.LateGet(NewLateBinding.LateGet(ws, null, "Application", new object[0], null, null, null), null, "WorksheetFunction", new object[0], null, null, null);
			object[] array3 = new object[1];
			instance = ws;
			object instance4 = instance;
			object[] array4 = new object[1];
			instance2 = (array4[0] = "AQ9:AQ" + Conversions.ToString(num));
			array3[0] = NewLateBinding.LateGet(instance4, null, "Range", array4, null, null, null);
			array = array3;
			bool[] obj3 = new bool[1] { true };
			array2 = obj3;
			object value2 = NewLateBinding.LateGet(instance3, null, "Sum", array3, null, null, obj3);
			if (array2[0])
			{
				NewLateBinding.LateSetComplex(instance, null, "Range", new object[2]
				{
					instance2,
					array[0]
				}, null, null, OptimisticSet: true, RValueBase: false);
			}
			value = Conversions.ToDouble(value2);
		}
		object instance5 = ws;
		NewLateBinding.LateSetComplex(NewLateBinding.LateGet(instance5, null, "Range", new object[1] { "AQ5" }, null, null, null), null, "Value", new object[1] { Math.Round(value, 2) }, null, null, OptimisticSet: false, RValueBase: true);
		NewLateBinding.LateSetComplex(NewLateBinding.LateGet(instance5, null, "Range", new object[1] { "AP5" }, null, null, null), null, "Value", new object[1] { "TOTAL WEIGHT:" }, null, null, OptimisticSet: false, RValueBase: true);
		NewLateBinding.LateSetComplex(NewLateBinding.LateGet(NewLateBinding.LateGet(instance5, null, "Range", new object[1] { "I5:K5" }, null, null, null), null, "Font", new object[0], null, null, null), null, "Bold", new object[1] { true }, null, null, OptimisticSet: false, RValueBase: true);
		instance5 = null;
	}

	private void ProcessSectionFarmingAndWeight(object ws)
	{
		object instance;
		object[] obj = new object[2]
		{
			NewLateBinding.LateGet(instance = NewLateBinding.LateGet(ws, null, "Rows", new object[0], null, null, null), null, "Count", new object[0], null, null, null),
			7
		};
		object[] array = obj;
		bool[] obj2 = new bool[2] { true, false };
		bool[] array2 = obj2;
		object instance2 = NewLateBinding.LateGet(ws, null, "Cells", obj, null, null, obj2);
		if (array2[0])
		{
			NewLateBinding.LateSetComplex(instance, null, "Count", new object[1] { array[0] }, null, null, OptimisticSet: true, RValueBase: true);
		}
		int num = Conversions.ToInteger(NewLateBinding.LateGet(NewLateBinding.LateGet(instance2, null, "End", new object[1] { -4162 }, null, null, null), null, "Row", new object[0], null, null, null));
		object[] array3 = new object[2];
		instance = ws;
		object instance3 = instance;
		object[] array4 = new object[2];
		instance2 = (array4[0] = 7);
		object obj3 = (array4[1] = 24);
		array3[0] = NewLateBinding.LateGet(instance3, null, "Cells", array4, null, null, null);
		object[] array5 = new object[2];
		object instance4;
		object obj4 = (array5[0] = NewLateBinding.LateGet(instance4 = NewLateBinding.LateGet(ws, null, "Rows", new object[0], null, null, null), null, "Count", new object[0], null, null, null));
		object obj5 = (array5[1] = 50);
		object[] array6 = array5;
		bool[] array7;
		object obj6 = NewLateBinding.LateGet(ws, null, "Cells", array5, null, null, array7 = new bool[2] { true, false });
		if (array7[0])
		{
			NewLateBinding.LateSetComplex(instance4, null, "Count", new object[1] { array6[0] }, null, null, OptimisticSet: true, RValueBase: true);
		}
		array3[1] = obj6;
		array = array3;
		object instance5 = NewLateBinding.LateGet(ws, null, "Range", array3, null, null, array2 = new bool[2] { true, true });
		if (array2[0])
		{
			NewLateBinding.LateSetComplex(instance, null, "Cells", new object[3]
			{
				instance2,
				obj3,
				array[0]
			}, null, null, OptimisticSet: true, RValueBase: false);
		}
		if (array2[1])
		{
			NewLateBinding.LateSetComplex(ws, null, "Cells", new object[3]
			{
				obj4,
				obj5,
				array[1]
			}, null, null, OptimisticSet: true, RValueBase: false);
		}
		NewLateBinding.LateCall(instance5, null, "ClearContents", new object[0], null, null, null, IgnoreReturn: true);
		object instance6 = ws;
		NewLateBinding.LateSetComplex(NewLateBinding.LateGet(instance6, null, "Cells", new object[2] { 7, 24 }, null, null, null), null, "Value", new object[1] { "NODE A" }, null, null, OptimisticSet: false, RValueBase: true);
		NewLateBinding.LateSetComplex(NewLateBinding.LateGet(instance6, null, "Cells", new object[2] { 7, 30 }, null, null, null), null, "Value", new object[1] { "NODE B" }, null, null, OptimisticSet: false, RValueBase: true);
		NewLateBinding.LateSetComplex(NewLateBinding.LateGet(instance6, null, "Range", new object[1] { "AK7" }, null, null, null), null, "Value", new object[1] { "SECTION FARMING" }, null, null, OptimisticSet: false, RValueBase: true);
		NewLateBinding.LateSetComplex(NewLateBinding.LateGet(instance6, null, "Range", new object[1] { "AN7" }, null, null, null), null, "Value", new object[1] { "WEIGHT(Kg)" }, null, null, OptimisticSet: false, RValueBase: true);
		NewLateBinding.LateSetComplex(NewLateBinding.LateGet(instance6, null, "Range", new object[1] { "X8" }, null, null, null), null, "Value", new object[1] { "WEB WIDTH" }, null, null, OptimisticSet: false, RValueBase: true);
		NewLateBinding.LateSetComplex(NewLateBinding.LateGet(instance6, null, "Range", new object[1] { "Y8" }, null, null, null), null, "Value", new object[1] { "WEB THK." }, null, null, OptimisticSet: false, RValueBase: true);
		NewLateBinding.LateSetComplex(NewLateBinding.LateGet(instance6, null, "Range", new object[1] { "Z8" }, null, null, null), null, "Value", new object[1] { "Tf WIDTH" }, null, null, OptimisticSet: false, RValueBase: true);
		NewLateBinding.LateSetComplex(NewLateBinding.LateGet(instance6, null, "Range", new object[1] { "AA8" }, null, null, null), null, "Value", new object[1] { "Tf THK." }, null, null, OptimisticSet: false, RValueBase: true);
		NewLateBinding.LateSetComplex(NewLateBinding.LateGet(instance6, null, "Range", new object[1] { "AB8" }, null, null, null), null, "Value", new object[1] { "Bf WIDTH" }, null, null, OptimisticSet: false, RValueBase: true);
		NewLateBinding.LateSetComplex(NewLateBinding.LateGet(instance6, null, "Range", new object[1] { "AC8" }, null, null, null), null, "Value", new object[1] { "Bf THK." }, null, null, OptimisticSet: false, RValueBase: true);
		NewLateBinding.LateSetComplex(NewLateBinding.LateGet(instance6, null, "Range", new object[1] { "AD8" }, null, null, null), null, "Value", new object[1] { "WEB WIDTH" }, null, null, OptimisticSet: false, RValueBase: true);
		NewLateBinding.LateSetComplex(NewLateBinding.LateGet(instance6, null, "Range", new object[1] { "AE8" }, null, null, null), null, "Value", new object[1] { "WEB THK." }, null, null, OptimisticSet: false, RValueBase: true);
		NewLateBinding.LateSetComplex(NewLateBinding.LateGet(instance6, null, "Range", new object[1] { "AF8" }, null, null, null), null, "Value", new object[1] { "Tf WIDTH" }, null, null, OptimisticSet: false, RValueBase: true);
		NewLateBinding.LateSetComplex(NewLateBinding.LateGet(instance6, null, "Range", new object[1] { "AG8" }, null, null, null), null, "Value", new object[1] { "Tf THK." }, null, null, OptimisticSet: false, RValueBase: true);
		NewLateBinding.LateSetComplex(NewLateBinding.LateGet(instance6, null, "Range", new object[1] { "AH8" }, null, null, null), null, "Value", new object[1] { "Bf WIDTH" }, null, null, OptimisticSet: false, RValueBase: true);
		NewLateBinding.LateSetComplex(NewLateBinding.LateGet(instance6, null, "Range", new object[1] { "AI8" }, null, null, null), null, "Value", new object[1] { "Bf THK." }, null, null, OptimisticSet: false, RValueBase: true);
		NewLateBinding.LateSetComplex(NewLateBinding.LateGet(instance6, null, "Range", new object[1] { "AJ7" }, null, null, null), null, "Value", new object[1] { "AVG.WEB" }, null, null, OptimisticSet: false, RValueBase: true);
		NewLateBinding.LateSetComplex(NewLateBinding.LateGet(instance6, null, "Range", new object[1] { "AK8" }, null, null, null), null, "Value", new object[1] { "WEB" }, null, null, OptimisticSet: false, RValueBase: true);
		NewLateBinding.LateSetComplex(NewLateBinding.LateGet(instance6, null, "Range", new object[1] { "AL8" }, null, null, null), null, "Value", new object[1] { "TOP FLANGE" }, null, null, OptimisticSet: false, RValueBase: true);
		NewLateBinding.LateSetComplex(NewLateBinding.LateGet(instance6, null, "Range", new object[1] { "AM8" }, null, null, null), null, "Value", new object[1] { "BOTT.FLANGE" }, null, null, OptimisticSet: false, RValueBase: true);
		NewLateBinding.LateSetComplex(NewLateBinding.LateGet(instance6, null, "Range", new object[1] { "AN8" }, null, null, null), null, "Value", new object[1] { "WEB(Kg)" }, null, null, OptimisticSet: false, RValueBase: true);
		NewLateBinding.LateSetComplex(NewLateBinding.LateGet(instance6, null, "Range", new object[1] { "AO8" }, null, null, null), null, "Value", new object[1] { "TOP FLANGE(Kg)" }, null, null, OptimisticSet: false, RValueBase: true);
		NewLateBinding.LateSetComplex(NewLateBinding.LateGet(instance6, null, "Range", new object[1] { "AP8" }, null, null, null), null, "Value", new object[1] { "BOTT.FLANGE(Kg)" }, null, null, OptimisticSet: false, RValueBase: true);
		NewLateBinding.LateSetComplex(NewLateBinding.LateGet(instance6, null, "Range", new object[1] { "AQ8" }, null, null, null), null, "Value", new object[1] { "TOTAL(Kg)" }, null, null, OptimisticSet: false, RValueBase: true);
		NewLateBinding.LateCall(NewLateBinding.LateGet(instance6, null, "Range", new object[1] { "E4:AK4" }, null, null, null), null, "Merge", new object[0], null, null, null, IgnoreReturn: true);
		NewLateBinding.LateCall(NewLateBinding.LateGet(instance6, null, "Range", new object[1] { "X7:AC7" }, null, null, null), null, "Merge", new object[0], null, null, null, IgnoreReturn: true);
		NewLateBinding.LateCall(NewLateBinding.LateGet(instance6, null, "Range", new object[1] { "AD7:AI7" }, null, null, null), null, "Merge", new object[0], null, null, null, IgnoreReturn: true);
		NewLateBinding.LateCall(NewLateBinding.LateGet(instance6, null, "Range", new object[1] { "AJ7:AJ8" }, null, null, null), null, "Merge", new object[0], null, null, null, IgnoreReturn: true);
		NewLateBinding.LateCall(NewLateBinding.LateGet(instance6, null, "Range", new object[1] { "AK7:AM7" }, null, null, null), null, "Merge", new object[0], null, null, null, IgnoreReturn: true);
		NewLateBinding.LateCall(NewLateBinding.LateGet(instance6, null, "Range", new object[1] { "AN7:AQ7" }, null, null, null), null, "Merge", new object[0], null, null, null, IgnoreReturn: true);
		instance6 = null;
		int num2 = num;
		checked
		{
			for (int i = 9; i <= num2; i++)
			{
				obj5 = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 3 }, null, null, array2 = new bool[2] { true, false });
				if (array2[0])
				{
					i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
				}
				double num3 = Conversion.Val(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(obj5, null, "Value", new object[0], null, null, null)));
				obj5 = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 4 }, null, null, array2 = new bool[2] { true, false });
				if (array2[0])
				{
					i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
				}
				string left = Conversions.ToString(NewLateBinding.LateGet(obj5, null, "Value", new object[0], null, null, null)).Trim().ToUpper();
				obj5 = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 6 }, null, null, array2 = new bool[2] { true, false });
				if (array2[0])
				{
					i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
				}
				int num4 = (int)Math.Round(Conversion.Val(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(obj5, null, "Value", new object[0], null, null, null))));
				if (unchecked(num4 == 671 || num4 == 672))
				{
					object[] obj7 = new object[2] { i, 43 };
					array = obj7;
					bool[] obj8 = new bool[2] { true, false };
					array2 = obj8;
					obj5 = NewLateBinding.LateGet(ws, null, "Cells", obj7, null, null, obj8);
					if (array2[0])
					{
						i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
					}
					NewLateBinding.LateSetComplex(obj5, null, "Value", new object[1] { 0 }, null, null, OptimisticSet: false, RValueBase: true);
					obj5 = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 37 }, null, null, array2 = new bool[2] { true, false });
					if (array2[0])
					{
						i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
					}
					object instance7 = obj5;
					object[] array8 = new object[1];
					obj5 = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 4 }, null, null, array2 = new bool[2] { true, false });
					if (array2[0])
					{
						i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
					}
					array8[0] = NewLateBinding.LateGet(obj5, null, "Value", new object[0], null, null, null);
					NewLateBinding.LateSetComplex(instance7, null, "Value", array8, null, null, OptimisticSet: false, RValueBase: true);
					obj5 = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 38 }, null, null, array2 = new bool[2] { true, false });
					if (array2[0])
					{
						i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
					}
					NewLateBinding.LateSetComplex(obj5, null, "Value", new object[1] { "" }, null, null, OptimisticSet: false, RValueBase: true);
					obj5 = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 39 }, null, null, array2 = new bool[2] { true, false });
					if (array2[0])
					{
						i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
					}
					NewLateBinding.LateSetComplex(obj5, null, "Value", new object[1] { "" }, null, null, OptimisticSet: false, RValueBase: true);
				}
				else
				{
					if (num4 == 680)
					{
						object[] obj9 = new object[2] { i, 13 };
						array = obj9;
						bool[] obj10 = new bool[2] { true, false };
						array2 = obj10;
						obj5 = NewLateBinding.LateGet(ws, null, "Cells", obj9, null, null, obj10);
						if (array2[0])
						{
							i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
						}
						if (Conversion.Val(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(obj5, null, "Value", new object[0], null, null, null))) == 0.0)
						{
							object[] obj11 = new object[2] { i, 13 };
							array = obj11;
							bool[] obj12 = new bool[2] { true, false };
							array2 = obj12;
							obj5 = NewLateBinding.LateGet(ws, null, "Cells", obj11, null, null, obj12);
							if (array2[0])
							{
								i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
							}
							object instance8 = obj5;
							object[] array9 = new object[1];
							obj5 = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 11 }, null, null, array2 = new bool[2] { true, false });
							if (array2[0])
							{
								i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
							}
							array9[0] = NewLateBinding.LateGet(obj5, null, "Value", new object[0], null, null, null);
							NewLateBinding.LateSetComplex(instance8, null, "Value", array9, null, null, OptimisticSet: false, RValueBase: true);
						}
						obj5 = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 12 }, null, null, array2 = new bool[2] { true, false });
						if (array2[0])
						{
							i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
						}
						if (Conversion.Val(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(obj5, null, "Value", new object[0], null, null, null))) == 0.0)
						{
							object[] obj13 = new object[2] { i, 12 };
							array = obj13;
							bool[] obj14 = new bool[2] { true, false };
							array2 = obj14;
							obj5 = NewLateBinding.LateGet(ws, null, "Cells", obj13, null, null, obj14);
							if (array2[0])
							{
								i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
							}
							object instance9 = obj5;
							object[] array10 = new object[1];
							obj5 = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 10 }, null, null, array2 = new bool[2] { true, false });
							if (array2[0])
							{
								i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
							}
							array10[0] = NewLateBinding.LateGet(obj5, null, "Value", new object[0], null, null, null);
							NewLateBinding.LateSetComplex(instance9, null, "Value", array10, null, null, OptimisticSet: false, RValueBase: true);
						}
					}
					if (num4 == 697)
					{
						object[] obj15 = new object[2] { i, 8 };
						array = obj15;
						bool[] obj16 = new bool[2] { true, false };
						array2 = obj16;
						obj5 = NewLateBinding.LateGet(ws, null, "Cells", obj15, null, null, obj16);
						if (array2[0])
						{
							i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
						}
						double num5 = Conversion.Val(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(obj5, null, "Value", new object[0], null, null, null)));
						obj5 = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 9 }, null, null, array2 = new bool[2] { true, false });
						if (array2[0])
						{
							i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
						}
						double num6 = Conversion.Val(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(obj5, null, "Value", new object[0], null, null, null)));
						obj5 = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 10 }, null, null, array2 = new bool[2] { true, false });
						if (array2[0])
						{
							i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
						}
						double num7 = Conversion.Val(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(obj5, null, "Value", new object[0], null, null, null)));
						obj5 = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 11 }, null, null, array2 = new bool[2] { true, false });
						if (array2[0])
						{
							i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
						}
						double num8 = Conversion.Val(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(obj5, null, "Value", new object[0], null, null, null)));
						if (num5 == num7)
						{
							object[] obj17 = new object[2] { i, 7 };
							array = obj17;
							bool[] obj18 = new bool[2] { true, false };
							array2 = obj18;
							obj5 = NewLateBinding.LateGet(ws, null, "Cells", obj17, null, null, obj18);
							if (array2[0])
							{
								i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
							}
							double num9 = Conversion.Val(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(obj5, null, "Value", new object[0], null, null, null)));
							obj5 = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 43 }, null, null, array2 = new bool[2] { true, false });
							if (array2[0])
							{
								i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
							}
							NewLateBinding.LateSetComplex(obj5, null, "Value", new object[1] { Math.Round(num3 * num9 * 7850.0, 3) }, null, null, OptimisticSet: false, RValueBase: true);
							obj5 = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 37 }, null, null, array2 = new bool[2] { true, false });
							if (array2[0])
							{
								i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
							}
							NewLateBinding.LateSetComplex(obj5, null, "Value", new object[1] { "UPT" }, null, null, OptimisticSet: false, RValueBase: true);
							obj5 = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 38 }, null, null, array2 = new bool[2] { true, false });
							if (array2[0])
							{
								i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
							}
							NewLateBinding.LateSetComplex(obj5, null, "Value", new object[1] { "" }, null, null, OptimisticSet: false, RValueBase: true);
							obj5 = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 39 }, null, null, array2 = new bool[2] { true, false });
							if (array2[0])
							{
								i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
							}
							NewLateBinding.LateSetComplex(obj5, null, "Value", new object[1] { "" }, null, null, OptimisticSet: false, RValueBase: true);
						}
						else
						{
							int num10 = (int)Math.Round(1000.0 * num5 - 2000.0 * num8);
							int num11 = (int)Math.Round(1000.0 * num6);
							int num12 = (int)Math.Round(1000.0 * num7);
							int num13 = (int)Math.Round(1000.0 * num8);
							object[] obj19 = new object[2] { i, 37 };
							array = obj19;
							bool[] obj20 = new bool[2] { true, false };
							array2 = obj20;
							obj5 = NewLateBinding.LateGet(ws, null, "Cells", obj19, null, null, obj20);
							if (array2[0])
							{
								i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
							}
							NewLateBinding.LateSetComplex(obj5, null, "Value", new object[1] { "2x" + Conversions.ToString(num10) + "x" + Conversions.ToString(num11) }, null, null, OptimisticSet: false, RValueBase: true);
							obj5 = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 38 }, null, null, array2 = new bool[2] { true, false });
							if (array2[0])
							{
								i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
							}
							NewLateBinding.LateSetComplex(obj5, null, "Value", new object[1] { "2x" + Conversions.ToString(num12) + "x" + Conversions.ToString(num13) }, null, null, OptimisticSet: false, RValueBase: true);
							obj5 = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 39 }, null, null, array2 = new bool[2] { true, false });
							if (array2[0])
							{
								i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
							}
							NewLateBinding.LateSetComplex(obj5, null, "Value", new object[1] { "2x" + Conversions.ToString(num12) + "x" + Conversions.ToString(num13) }, null, null, OptimisticSet: false, RValueBase: true);
							double num14 = Math.Round(2.0 * num7 * num8 * num3 * 7.85 * 1000.0, 3);
							obj5 = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 41 }, null, null, array2 = new bool[2] { true, false });
							if (array2[0])
							{
								i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
							}
							NewLateBinding.LateSetComplex(obj5, null, "Value", new object[1] { num14 }, null, null, OptimisticSet: false, RValueBase: true);
							double num15 = Math.Round(2.0 * num7 * num8 * num3 * 7.85 * 1000.0, 3);
							obj5 = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 42 }, null, null, array2 = new bool[2] { true, false });
							if (array2[0])
							{
								i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
							}
							NewLateBinding.LateSetComplex(obj5, null, "Value", new object[1] { num15 }, null, null, OptimisticSet: false, RValueBase: true);
							obj5 = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 7 }, null, null, array2 = new bool[2] { true, false });
							if (array2[0])
							{
								i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
							}
							double num16 = Conversion.Val(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(obj5, null, "Value", new object[0], null, null, null)));
							double num17 = Math.Round(num3 * num16 * 7850.0, 3);
							obj5 = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 43 }, null, null, array2 = new bool[2] { true, false });
							if (array2[0])
							{
								i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
							}
							NewLateBinding.LateSetComplex(obj5, null, "Value", new object[1] { num17 }, null, null, OptimisticSet: false, RValueBase: true);
							double num18 = Math.Round(num17 - (num14 + num15), 3);
							obj5 = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 40 }, null, null, array2 = new bool[2] { true, false });
							if (array2[0])
							{
								i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
							}
							NewLateBinding.LateSetComplex(obj5, null, "Value", new object[1] { num18 }, null, null, OptimisticSet: false, RValueBase: true);
						}
					}
					else if (num4 >= 690 && num4 <= 699 && num4 != 697)
					{
						obj5 = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 7 }, null, null, array2 = new bool[2] { true, false });
						if (array2[0])
						{
							i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
						}
						double num19 = Conversion.Val(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(obj5, null, "Value", new object[0], null, null, null)));
						obj5 = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 43 }, null, null, array2 = new bool[2] { true, false });
						if (array2[0])
						{
							i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
						}
						NewLateBinding.LateSetComplex(obj5, null, "Value", new object[1] { Math.Round(num3 * num19 * 7850.0, 3) }, null, null, OptimisticSet: false, RValueBase: true);
						obj5 = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 37 }, null, null, array2 = new bool[2] { true, false });
						if (array2[0])
						{
							i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
						}
						NewLateBinding.LateSetComplex(obj5, null, "Value", new object[1] { "UPT" }, null, null, OptimisticSet: false, RValueBase: true);
						obj5 = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 38 }, null, null, array2 = new bool[2] { true, false });
						if (array2[0])
						{
							i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
						}
						NewLateBinding.LateSetComplex(obj5, null, "Value", new object[1] { "" }, null, null, OptimisticSet: false, RValueBase: true);
						obj5 = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 39 }, null, null, array2 = new bool[2] { true, false });
						if (array2[0])
						{
							i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
						}
						NewLateBinding.LateSetComplex(obj5, null, "Value", new object[1] { "" }, null, null, OptimisticSet: false, RValueBase: true);
					}
					else if (unchecked(num4 == 613 || num4 == 614 || num4 == 615))
					{
						double[] array11 = new double[8];
						int num20 = 1;
						do
						{
							int num21 = num20;
							object[] obj21 = new object[2]
							{
								i,
								6 + num20
							};
							array = obj21;
							bool[] obj22 = new bool[2] { true, false };
							array2 = obj22;
							obj5 = NewLateBinding.LateGet(ws, null, "Cells", obj21, null, null, obj22);
							if (array2[0])
							{
								i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
							}
							array11[num21] = Conversion.Val(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(obj5, null, "Value", new object[0], null, null, null)));
							num20++;
						}
						while (num20 <= 7);
						int num22 = 24;
						do
						{
							object[] obj23 = new object[2] { i, num22 };
							array = obj23;
							bool[] obj24 = new bool[2] { true, true };
							array2 = obj24;
							obj5 = NewLateBinding.LateGet(ws, null, "Cells", obj23, null, null, obj24);
							if (array2[0])
							{
								i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
							}
							if (array2[1])
							{
								num22 = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[1]), typeof(int));
							}
							NewLateBinding.LateSetComplex(obj5, null, "Value", new object[1] { 0 }, null, null, OptimisticSet: false, RValueBase: true);
							num22++;
						}
						while (num22 <= 35);
						object[] obj25 = new object[2] { i, 36 };
						array = obj25;
						bool[] obj26 = new bool[2] { true, false };
						array2 = obj26;
						obj5 = NewLateBinding.LateGet(ws, null, "Cells", obj25, null, null, obj26);
						if (array2[0])
						{
							i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
						}
						NewLateBinding.LateSetComplex(obj5, null, "Value", new object[1] { 0 }, null, null, OptimisticSet: false, RValueBase: true);
						obj5 = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 37 }, null, null, array2 = new bool[2] { true, false });
						if (array2[0])
						{
							i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
						}
						object instance10 = obj5;
						object[] array12 = new object[1];
						obj5 = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 4 }, null, null, array2 = new bool[2] { true, false });
						if (array2[0])
						{
							i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
						}
						array12[0] = NewLateBinding.LateGet(obj5, null, "Value", new object[0], null, null, null);
						NewLateBinding.LateSetComplex(instance10, null, "Value", array12, null, null, OptimisticSet: false, RValueBase: true);
						obj5 = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 52 }, null, null, array2 = new bool[2] { true, false });
						if (array2[0])
						{
							i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
						}
						double num23 = Conversion.Val(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(obj5, null, "Value", new object[0], null, null, null)));
						obj5 = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 53 }, null, null, array2 = new bool[2] { true, false });
						if (array2[0])
						{
							i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
						}
						double num24 = Conversion.Val(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(obj5, null, "Value", new object[0], null, null, null)));
						obj5 = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 54 }, null, null, array2 = new bool[2] { true, false });
						if (array2[0])
						{
							i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
						}
						double num25 = Conversion.Val(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(obj5, null, "Value", new object[0], null, null, null)));
						obj5 = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 55 }, null, null, array2 = new bool[2] { true, false });
						if (array2[0])
						{
							i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
						}
						double num26 = Conversion.Val(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(obj5, null, "Value", new object[0], null, null, null)));
						unchecked
						{
							switch (num4)
							{
							case 613:
							{
								object[] obj29 = new object[2] { i, 26 };
								array = obj29;
								bool[] obj30 = new bool[2] { true, false };
								array2 = obj30;
								obj5 = NewLateBinding.LateGet(ws, null, "Cells", obj29, null, null, obj30);
								if (array2[0])
								{
									i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
								}
								NewLateBinding.LateSetComplex(obj5, null, "Value", new object[1] { Math.Round(num23 * 1000.0) }, null, null, OptimisticSet: false, RValueBase: true);
								obj5 = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 27 }, null, null, array2 = new bool[2] { true, false });
								if (array2[0])
								{
									i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
								}
								NewLateBinding.LateSetComplex(obj5, null, "Value", new object[1] { Math.Round(num24 * 1000.0) }, null, null, OptimisticSet: false, RValueBase: true);
								obj5 = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 28 }, null, null, array2 = new bool[2] { true, false });
								if (array2[0])
								{
									i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
								}
								NewLateBinding.LateSetComplex(obj5, null, "Value", new object[1] { 0 }, null, null, OptimisticSet: false, RValueBase: true);
								obj5 = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 29 }, null, null, array2 = new bool[2] { true, false });
								if (array2[0])
								{
									i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
								}
								NewLateBinding.LateSetComplex(obj5, null, "Value", new object[1] { 0 }, null, null, OptimisticSet: false, RValueBase: true);
								obj5 = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 32 }, null, null, array2 = new bool[2] { true, false });
								if (array2[0])
								{
									i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
								}
								NewLateBinding.LateSetComplex(obj5, null, "Value", new object[1] { Math.Round(num23 * 1000.0) }, null, null, OptimisticSet: false, RValueBase: true);
								obj5 = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 33 }, null, null, array2 = new bool[2] { true, false });
								if (array2[0])
								{
									i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
								}
								NewLateBinding.LateSetComplex(obj5, null, "Value", new object[1] { Math.Round(num24 * 1000.0) }, null, null, OptimisticSet: false, RValueBase: true);
								obj5 = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 34 }, null, null, array2 = new bool[2] { true, false });
								if (array2[0])
								{
									i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
								}
								NewLateBinding.LateSetComplex(obj5, null, "Value", new object[1] { 0 }, null, null, OptimisticSet: false, RValueBase: true);
								obj5 = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 35 }, null, null, array2 = new bool[2] { true, false });
								if (array2[0])
								{
									i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
								}
								NewLateBinding.LateSetComplex(obj5, null, "Value", new object[1] { 0 }, null, null, OptimisticSet: false, RValueBase: true);
								break;
							}
							case 614:
							{
								object[] obj27 = new object[2] { i, 26 };
								array = obj27;
								bool[] obj28 = new bool[2] { true, false };
								array2 = obj28;
								obj5 = NewLateBinding.LateGet(ws, null, "Cells", obj27, null, null, obj28);
								if (array2[0])
								{
									i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
								}
								NewLateBinding.LateSetComplex(obj5, null, "Value", new object[1] { 0 }, null, null, OptimisticSet: false, RValueBase: true);
								obj5 = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 27 }, null, null, array2 = new bool[2] { true, false });
								if (array2[0])
								{
									i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
								}
								NewLateBinding.LateSetComplex(obj5, null, "Value", new object[1] { 0 }, null, null, OptimisticSet: false, RValueBase: true);
								obj5 = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 28 }, null, null, array2 = new bool[2] { true, false });
								if (array2[0])
								{
									i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
								}
								NewLateBinding.LateSetComplex(obj5, null, "Value", new object[1] { Math.Round(num23 * 1000.0) }, null, null, OptimisticSet: false, RValueBase: true);
								obj5 = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 29 }, null, null, array2 = new bool[2] { true, false });
								if (array2[0])
								{
									i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
								}
								NewLateBinding.LateSetComplex(obj5, null, "Value", new object[1] { Math.Round(num24 * 1000.0) }, null, null, OptimisticSet: false, RValueBase: true);
								obj5 = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 32 }, null, null, array2 = new bool[2] { true, false });
								if (array2[0])
								{
									i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
								}
								NewLateBinding.LateSetComplex(obj5, null, "Value", new object[1] { 0 }, null, null, OptimisticSet: false, RValueBase: true);
								obj5 = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 33 }, null, null, array2 = new bool[2] { true, false });
								if (array2[0])
								{
									i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
								}
								NewLateBinding.LateSetComplex(obj5, null, "Value", new object[1] { 0 }, null, null, OptimisticSet: false, RValueBase: true);
								obj5 = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 34 }, null, null, array2 = new bool[2] { true, false });
								if (array2[0])
								{
									i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
								}
								NewLateBinding.LateSetComplex(obj5, null, "Value", new object[1] { Math.Round(num23 * 1000.0) }, null, null, OptimisticSet: false, RValueBase: true);
								obj5 = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 35 }, null, null, array2 = new bool[2] { true, false });
								if (array2[0])
								{
									i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
								}
								NewLateBinding.LateSetComplex(obj5, null, "Value", new object[1] { Math.Round(num24 * 1000.0) }, null, null, OptimisticSet: false, RValueBase: true);
								break;
							}
							case 615:
							{
								double num27 = num23;
								double num28 = num24;
								double num29 = num25;
								double num30 = num26;
								if ((num27 == 0.0 || num28 == 0.0) && num29 > 0.0 && num30 > 0.0)
								{
									num27 = num29;
									num28 = num30;
								}
								if ((num29 == 0.0 || num30 == 0.0) && num27 > 0.0 && num28 > 0.0)
								{
									num29 = num27;
									num30 = num28;
								}
								obj5 = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 26 }, null, null, array2 = new bool[2] { true, false });
								if (array2[0])
								{
									i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
								}
								NewLateBinding.LateSetComplex(obj5, null, "Value", new object[1] { Math.Round(num27 * 1000.0) }, null, null, OptimisticSet: false, RValueBase: true);
								obj5 = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 27 }, null, null, array2 = new bool[2] { true, false });
								if (array2[0])
								{
									i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
								}
								NewLateBinding.LateSetComplex(obj5, null, "Value", new object[1] { Math.Round(num28 * 1000.0) }, null, null, OptimisticSet: false, RValueBase: true);
								obj5 = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 28 }, null, null, array2 = new bool[2] { true, false });
								if (array2[0])
								{
									i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
								}
								NewLateBinding.LateSetComplex(obj5, null, "Value", new object[1] { Math.Round(num29 * 1000.0) }, null, null, OptimisticSet: false, RValueBase: true);
								obj5 = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 29 }, null, null, array2 = new bool[2] { true, false });
								if (array2[0])
								{
									i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
								}
								NewLateBinding.LateSetComplex(obj5, null, "Value", new object[1] { Math.Round(num30 * 1000.0) }, null, null, OptimisticSet: false, RValueBase: true);
								obj5 = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 32 }, null, null, array2 = new bool[2] { true, false });
								if (array2[0])
								{
									i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
								}
								NewLateBinding.LateSetComplex(obj5, null, "Value", new object[1] { Math.Round(num27 * 1000.0) }, null, null, OptimisticSet: false, RValueBase: true);
								obj5 = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 33 }, null, null, array2 = new bool[2] { true, false });
								if (array2[0])
								{
									i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
								}
								NewLateBinding.LateSetComplex(obj5, null, "Value", new object[1] { Math.Round(num28 * 1000.0) }, null, null, OptimisticSet: false, RValueBase: true);
								obj5 = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 34 }, null, null, array2 = new bool[2] { true, false });
								if (array2[0])
								{
									i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
								}
								NewLateBinding.LateSetComplex(obj5, null, "Value", new object[1] { Math.Round(num29 * 1000.0) }, null, null, OptimisticSet: false, RValueBase: true);
								obj5 = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 35 }, null, null, array2 = new bool[2] { true, false });
								if (array2[0])
								{
									i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
								}
								NewLateBinding.LateSetComplex(obj5, null, "Value", new object[1] { Math.Round(num30 * 1000.0) }, null, null, OptimisticSet: false, RValueBase: true);
								break;
							}
							}
							switch (num4)
							{
							case 613:
							{
								object[] obj35 = new object[2] { i, 38 };
								array = obj35;
								bool[] obj36 = new bool[2] { true, false };
								array2 = obj36;
								obj5 = NewLateBinding.LateGet(ws, null, "Cells", obj35, null, null, obj36);
								if (array2[0])
								{
									i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
								}
								object instance14 = obj5;
								object[] array16 = new object[1];
								obj5 = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 26 }, null, null, array2 = new bool[2] { true, false });
								if (array2[0])
								{
									i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
								}
								object left5 = Operators.ConcatenateObject(NewLateBinding.LateGet(obj5, null, "Value", new object[0], null, null, null), "x");
								obj5 = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 27 }, null, null, array2 = new bool[2] { true, false });
								if (array2[0])
								{
									i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
								}
								array16[0] = Operators.ConcatenateObject(left5, NewLateBinding.LateGet(obj5, null, "Value", new object[0], null, null, null));
								NewLateBinding.LateSetComplex(instance14, null, "Value", array16, null, null, OptimisticSet: false, RValueBase: true);
								obj5 = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 39 }, null, null, array2 = new bool[2] { true, false });
								if (array2[0])
								{
									i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
								}
								NewLateBinding.LateSetComplex(obj5, null, "Value", new object[1] { "" }, null, null, OptimisticSet: false, RValueBase: true);
								break;
							}
							case 614:
							{
								object[] obj33 = new object[2] { i, 38 };
								array = obj33;
								bool[] obj34 = new bool[2] { true, false };
								array2 = obj34;
								obj5 = NewLateBinding.LateGet(ws, null, "Cells", obj33, null, null, obj34);
								if (array2[0])
								{
									i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
								}
								NewLateBinding.LateSetComplex(obj5, null, "Value", new object[1] { "" }, null, null, OptimisticSet: false, RValueBase: true);
								obj5 = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 39 }, null, null, array2 = new bool[2] { true, false });
								if (array2[0])
								{
									i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
								}
								object instance13 = obj5;
								object[] array15 = new object[1];
								obj5 = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 28 }, null, null, array2 = new bool[2] { true, false });
								if (array2[0])
								{
									i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
								}
								object left4 = Operators.ConcatenateObject(NewLateBinding.LateGet(obj5, null, "Value", new object[0], null, null, null), "x");
								obj5 = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 29 }, null, null, array2 = new bool[2] { true, false });
								if (array2[0])
								{
									i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
								}
								array15[0] = Operators.ConcatenateObject(left4, NewLateBinding.LateGet(obj5, null, "Value", new object[0], null, null, null));
								NewLateBinding.LateSetComplex(instance13, null, "Value", array15, null, null, OptimisticSet: false, RValueBase: true);
								break;
							}
							case 615:
							{
								object[] obj31 = new object[2] { i, 38 };
								array = obj31;
								bool[] obj32 = new bool[2] { true, false };
								array2 = obj32;
								obj5 = NewLateBinding.LateGet(ws, null, "Cells", obj31, null, null, obj32);
								if (array2[0])
								{
									i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
								}
								object instance11 = obj5;
								object[] array13 = new object[1];
								obj5 = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 26 }, null, null, array2 = new bool[2] { true, false });
								if (array2[0])
								{
									i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
								}
								object left2 = Operators.ConcatenateObject(NewLateBinding.LateGet(obj5, null, "Value", new object[0], null, null, null), "x");
								obj5 = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 27 }, null, null, array2 = new bool[2] { true, false });
								if (array2[0])
								{
									i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
								}
								array13[0] = Operators.ConcatenateObject(left2, NewLateBinding.LateGet(obj5, null, "Value", new object[0], null, null, null));
								NewLateBinding.LateSetComplex(instance11, null, "Value", array13, null, null, OptimisticSet: false, RValueBase: true);
								obj5 = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 39 }, null, null, array2 = new bool[2] { true, false });
								if (array2[0])
								{
									i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
								}
								object instance12 = obj5;
								object[] array14 = new object[1];
								obj5 = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 28 }, null, null, array2 = new bool[2] { true, false });
								if (array2[0])
								{
									i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
								}
								object left3 = Operators.ConcatenateObject(NewLateBinding.LateGet(obj5, null, "Value", new object[0], null, null, null), "x");
								obj5 = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 29 }, null, null, array2 = new bool[2] { true, false });
								if (array2[0])
								{
									i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
								}
								array14[0] = Operators.ConcatenateObject(left3, NewLateBinding.LateGet(obj5, null, "Value", new object[0], null, null, null));
								NewLateBinding.LateSetComplex(instance12, null, "Value", array14, null, null, OptimisticSet: false, RValueBase: true);
								break;
							}
							}
							double num31 = 0.0;
							double num32 = 0.0;
							switch (num4)
							{
							case 613:
							{
								object[] obj41 = new object[2] { i, 26 };
								array = obj41;
								bool[] obj42 = new bool[2] { true, false };
								array2 = obj42;
								obj5 = NewLateBinding.LateGet(ws, null, "Cells", obj41, null, null, obj42);
								if (array2[0])
								{
									i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
								}
								double num36 = num3 * Conversion.Val(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(obj5, null, "Value", new object[0], null, null, null)));
								obj5 = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 27 }, null, null, array2 = new bool[2] { true, false });
								if (array2[0])
								{
									i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
								}
								num31 = Math.Round(num36 * Conversion.Val(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(obj5, null, "Value", new object[0], null, null, null))) * 7.85 / 1000.0, 3);
								break;
							}
							case 614:
							{
								object[] obj39 = new object[2] { i, 28 };
								array = obj39;
								bool[] obj40 = new bool[2] { true, false };
								array2 = obj40;
								obj5 = NewLateBinding.LateGet(ws, null, "Cells", obj39, null, null, obj40);
								if (array2[0])
								{
									i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
								}
								double num35 = num3 * Conversion.Val(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(obj5, null, "Value", new object[0], null, null, null)));
								obj5 = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 29 }, null, null, array2 = new bool[2] { true, false });
								if (array2[0])
								{
									i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
								}
								num32 = Math.Round(num35 * Conversion.Val(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(obj5, null, "Value", new object[0], null, null, null))) * 7.85 / 1000.0, 3);
								break;
							}
							case 615:
							{
								object[] obj37 = new object[2] { i, 26 };
								array = obj37;
								bool[] obj38 = new bool[2] { true, false };
								array2 = obj38;
								obj5 = NewLateBinding.LateGet(ws, null, "Cells", obj37, null, null, obj38);
								if (array2[0])
								{
									i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
								}
								double num33 = num3 * Conversion.Val(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(obj5, null, "Value", new object[0], null, null, null)));
								obj5 = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 27 }, null, null, array2 = new bool[2] { true, false });
								if (array2[0])
								{
									i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
								}
								num31 = Math.Round(num33 * Conversion.Val(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(obj5, null, "Value", new object[0], null, null, null))) * 7.85 / 1000.0, 3);
								obj5 = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 28 }, null, null, array2 = new bool[2] { true, false });
								if (array2[0])
								{
									i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
								}
								double num34 = num3 * Conversion.Val(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(obj5, null, "Value", new object[0], null, null, null)));
								obj5 = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 29 }, null, null, array2 = new bool[2] { true, false });
								if (array2[0])
								{
									i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
								}
								num32 = Math.Round(num34 * Conversion.Val(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(obj5, null, "Value", new object[0], null, null, null))) * 7.85 / 1000.0, 3);
								break;
							}
							}
							obj5 = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 7 }, null, null, array2 = new bool[2] { true, false });
							if (array2[0])
							{
								i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
							}
							double num37 = Math.Round(num3 * Conversion.Val(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(obj5, null, "Value", new object[0], null, null, null))) * 7850.0, 3);
							double num38 = Math.Round(num37 - num31 - num32, 3);
							obj5 = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 40 }, null, null, array2 = new bool[2] { true, false });
							if (array2[0])
							{
								i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
							}
							NewLateBinding.LateSetComplex(obj5, null, "Value", new object[1] { num38 }, null, null, OptimisticSet: false, RValueBase: true);
							obj5 = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 41 }, null, null, array2 = new bool[2] { true, false });
							if (array2[0])
							{
								i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
							}
							NewLateBinding.LateSetComplex(obj5, null, "Value", new object[1] { num31 }, null, null, OptimisticSet: false, RValueBase: true);
							obj5 = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 42 }, null, null, array2 = new bool[2] { true, false });
							if (array2[0])
							{
								i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
							}
							NewLateBinding.LateSetComplex(obj5, null, "Value", new object[1] { num32 }, null, null, OptimisticSet: false, RValueBase: true);
							obj5 = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 43 }, null, null, array2 = new bool[2] { true, false });
							if (array2[0])
							{
								i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
							}
							NewLateBinding.LateSetComplex(obj5, null, "Value", new object[1] { num37 }, null, null, OptimisticSet: false, RValueBase: true);
						}
					}
					else
					{
						double[] array17 = new double[8];
						int num39 = 1;
						do
						{
							int num40 = num39;
							object[] obj43 = new object[2]
							{
								i,
								6 + num39
							};
							array = obj43;
							bool[] obj44 = new bool[2] { true, false };
							array2 = obj44;
							obj5 = NewLateBinding.LateGet(ws, null, "Cells", obj43, null, null, obj44);
							if (array2[0])
							{
								i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
							}
							array17[num40] = Conversion.Val(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(obj5, null, "Value", new object[0], null, null, null)));
							num39++;
						}
						while (num39 <= 7);
						if (num4 == 680)
						{
							object[] obj45 = new object[2] { i, 24 };
							array = obj45;
							bool[] obj46 = new bool[2] { true, false };
							array2 = obj46;
							obj5 = NewLateBinding.LateGet(ws, null, "Cells", obj45, null, null, obj46);
							if (array2[0])
							{
								i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
							}
							NewLateBinding.LateSetComplex(obj5, null, "Value", new object[1] { Math.Round((array17[1] - array17[5] - array17[7]) * 1000.0) }, null, null, OptimisticSet: false, RValueBase: true);
							obj5 = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 25 }, null, null, array2 = new bool[2] { true, false });
							if (array2[0])
							{
								i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
							}
							NewLateBinding.LateSetComplex(obj5, null, "Value", new object[1] { Math.Round(array17[2] * 1000.0) }, null, null, OptimisticSet: false, RValueBase: true);
							obj5 = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 26 }, null, null, array2 = new bool[2] { true, false });
							if (array2[0])
							{
								i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
							}
							NewLateBinding.LateSetComplex(obj5, null, "Value", new object[1] { Math.Round(array17[4] * 1000.0) }, null, null, OptimisticSet: false, RValueBase: true);
							obj5 = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 27 }, null, null, array2 = new bool[2] { true, false });
							if (array2[0])
							{
								i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
							}
							NewLateBinding.LateSetComplex(obj5, null, "Value", new object[1] { Math.Round(array17[5] * 1000.0) }, null, null, OptimisticSet: false, RValueBase: true);
							obj5 = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 28 }, null, null, array2 = new bool[2] { true, false });
							if (array2[0])
							{
								i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
							}
							NewLateBinding.LateSetComplex(obj5, null, "Value", new object[1] { Math.Round(array17[6] * 1000.0) }, null, null, OptimisticSet: false, RValueBase: true);
							obj5 = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 29 }, null, null, array2 = new bool[2] { true, false });
							if (array2[0])
							{
								i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
							}
							NewLateBinding.LateSetComplex(obj5, null, "Value", new object[1] { Math.Round(array17[7] * 1000.0) }, null, null, OptimisticSet: false, RValueBase: true);
							obj5 = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 30 }, null, null, array2 = new bool[2] { true, false });
							if (array2[0])
							{
								i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
							}
							NewLateBinding.LateSetComplex(obj5, null, "Value", new object[1] { Math.Round((array17[3] - array17[5] - array17[7]) * 1000.0) }, null, null, OptimisticSet: false, RValueBase: true);
							obj5 = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 31 }, null, null, array2 = new bool[2] { true, false });
							if (array2[0])
							{
								i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
							}
							NewLateBinding.LateSetComplex(obj5, null, "Value", new object[1] { Math.Round(array17[2] * 1000.0) }, null, null, OptimisticSet: false, RValueBase: true);
							obj5 = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 32 }, null, null, array2 = new bool[2] { true, false });
							if (array2[0])
							{
								i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
							}
							NewLateBinding.LateSetComplex(obj5, null, "Value", new object[1] { Math.Round(array17[4] * 1000.0) }, null, null, OptimisticSet: false, RValueBase: true);
							obj5 = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 33 }, null, null, array2 = new bool[2] { true, false });
							if (array2[0])
							{
								i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
							}
							NewLateBinding.LateSetComplex(obj5, null, "Value", new object[1] { Math.Round(array17[5] * 1000.0) }, null, null, OptimisticSet: false, RValueBase: true);
							obj5 = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 34 }, null, null, array2 = new bool[2] { true, false });
							if (array2[0])
							{
								i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
							}
							NewLateBinding.LateSetComplex(obj5, null, "Value", new object[1] { Math.Round(array17[6] * 1000.0) }, null, null, OptimisticSet: false, RValueBase: true);
							obj5 = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 35 }, null, null, array2 = new bool[2] { true, false });
							if (array2[0])
							{
								i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
							}
							NewLateBinding.LateSetComplex(obj5, null, "Value", new object[1] { Math.Round(array17[7] * 1000.0) }, null, null, OptimisticSet: false, RValueBase: true);
						}
						else
						{
							int num41 = 24;
							do
							{
								object[] obj47 = new object[2] { i, num41 };
								array = obj47;
								bool[] obj48 = new bool[2] { true, true };
								array2 = obj48;
								obj5 = NewLateBinding.LateGet(ws, null, "Cells", obj47, null, null, obj48);
								if (array2[0])
								{
									i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
								}
								if (array2[1])
								{
									num41 = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[1]), typeof(int));
								}
								NewLateBinding.LateSetComplex(obj5, null, "Value", new object[1] { 0 }, null, null, OptimisticSet: false, RValueBase: true);
								num41++;
							}
							while (num41 <= 35);
						}
						obj5 = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 24 }, null, null, array2 = new bool[2] { true, false });
						if (array2[0])
						{
							i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
						}
						double num42 = Conversion.Val(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(obj5, null, "Value", new object[0], null, null, null)));
						obj5 = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 30 }, null, null, array2 = new bool[2] { true, false });
						if (array2[0])
						{
							i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
						}
						double num43 = Conversion.Val(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(obj5, null, "Value", new object[0], null, null, null)));
						obj5 = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 25 }, null, null, array2 = new bool[2] { true, false });
						if (array2[0])
						{
							i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
						}
						double num44 = Conversion.Val(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(obj5, null, "Value", new object[0], null, null, null)));
						obj5 = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 36 }, null, null, array2 = new bool[2] { true, false });
						if (array2[0])
						{
							i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
						}
						NewLateBinding.LateSetComplex(obj5, null, "Value", new object[1] { Math.Round((num42 + num43) / 2.0) }, null, null, OptimisticSet: false, RValueBase: true);
						if (unchecked(num42 == 0.0 || num43 == 0.0 || num44 == 0.0))
						{
							object[] obj49 = new object[2] { i, 37 };
							array = obj49;
							bool[] obj50 = new bool[2] { true, false };
							array2 = obj50;
							obj5 = NewLateBinding.LateGet(ws, null, "Cells", obj49, null, null, obj50);
							if (array2[0])
							{
								i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
							}
							object instance15 = obj5;
							object[] array18 = new object[1];
							obj5 = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 4 }, null, null, array2 = new bool[2] { true, false });
							if (array2[0])
							{
								i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
							}
							array18[0] = NewLateBinding.LateGet(obj5, null, "Value", new object[0], null, null, null);
							NewLateBinding.LateSetComplex(instance15, null, "Value", array18, null, null, OptimisticSet: false, RValueBase: true);
							obj5 = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 38 }, null, null, array2 = new bool[2] { true, false });
							if (array2[0])
							{
								i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
							}
							NewLateBinding.LateSetComplex(obj5, null, "Value", new object[1] { "" }, null, null, OptimisticSet: false, RValueBase: true);
							obj5 = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 39 }, null, null, array2 = new bool[2] { true, false });
							if (array2[0])
							{
								i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
							}
							NewLateBinding.LateSetComplex(obj5, null, "Value", new object[1] { "" }, null, null, OptimisticSet: false, RValueBase: true);
						}
						else
						{
							if (num42 == num43)
							{
								object[] obj51 = new object[2] { i, 37 };
								array = obj51;
								bool[] obj52 = new bool[2] { true, false };
								array2 = obj52;
								obj5 = NewLateBinding.LateGet(ws, null, "Cells", obj51, null, null, obj52);
								if (array2[0])
								{
									i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
								}
								NewLateBinding.LateSetComplex(obj5, null, "Value", new object[1] { Conversions.ToString(num42) + "x" + Conversions.ToString(num44) }, null, null, OptimisticSet: false, RValueBase: true);
							}
							else
							{
								object[] obj53 = new object[2] { i, 37 };
								array = obj53;
								bool[] obj54 = new bool[2] { true, false };
								array2 = obj54;
								obj5 = NewLateBinding.LateGet(ws, null, "Cells", obj53, null, null, obj54);
								if (array2[0])
								{
									i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
								}
								NewLateBinding.LateSetComplex(obj5, null, "Value", new object[1] { Conversions.ToString(num42) + "-" + Conversions.ToString(num43) + "x" + Conversions.ToString(num44) }, null, null, OptimisticSet: false, RValueBase: true);
							}
							obj5 = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 38 }, null, null, array2 = new bool[2] { true, false });
							if (array2[0])
							{
								i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
							}
							object instance16 = obj5;
							object[] array19 = new object[1];
							obj5 = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 26 }, null, null, array2 = new bool[2] { true, false });
							if (array2[0])
							{
								i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
							}
							object left6 = Operators.ConcatenateObject(NewLateBinding.LateGet(obj5, null, "Value", new object[0], null, null, null), "x");
							obj5 = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 27 }, null, null, array2 = new bool[2] { true, false });
							if (array2[0])
							{
								i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
							}
							array19[0] = Operators.ConcatenateObject(left6, NewLateBinding.LateGet(obj5, null, "Value", new object[0], null, null, null));
							NewLateBinding.LateSetComplex(instance16, null, "Value", array19, null, null, OptimisticSet: false, RValueBase: true);
							obj5 = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 39 }, null, null, array2 = new bool[2] { true, false });
							if (array2[0])
							{
								i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
							}
							object instance17 = obj5;
							object[] array20 = new object[1];
							obj5 = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 28 }, null, null, array2 = new bool[2] { true, false });
							if (array2[0])
							{
								i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
							}
							object left7 = Operators.ConcatenateObject(NewLateBinding.LateGet(obj5, null, "Value", new object[0], null, null, null), "x");
							obj5 = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 29 }, null, null, array2 = new bool[2] { true, false });
							if (array2[0])
							{
								i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
							}
							array20[0] = Operators.ConcatenateObject(left7, NewLateBinding.LateGet(obj5, null, "Value", new object[0], null, null, null));
							NewLateBinding.LateSetComplex(instance17, null, "Value", array20, null, null, OptimisticSet: false, RValueBase: true);
						}
						obj5 = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 36 }, null, null, array2 = new bool[2] { true, false });
						if (array2[0])
						{
							i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
						}
						double num45 = Math.Round(num3 * Conversion.Val(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(obj5, null, "Value", new object[0], null, null, null))) * num44 * 7.85 / 1000.0, 3);
						obj5 = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 26 }, null, null, array2 = new bool[2] { true, false });
						if (array2[0])
						{
							i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
						}
						double num46 = num3 * Conversion.Val(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(obj5, null, "Value", new object[0], null, null, null)));
						obj5 = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 27 }, null, null, array2 = new bool[2] { true, false });
						if (array2[0])
						{
							i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
						}
						double num47 = Math.Round(num46 * Conversion.Val(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(obj5, null, "Value", new object[0], null, null, null))) * 7.85 / 1000.0, 3);
						obj5 = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 28 }, null, null, array2 = new bool[2] { true, false });
						if (array2[0])
						{
							i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
						}
						double num48 = num3 * Conversion.Val(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(obj5, null, "Value", new object[0], null, null, null)));
						obj5 = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 29 }, null, null, array2 = new bool[2] { true, false });
						if (array2[0])
						{
							i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
						}
						double num49 = Math.Round(num48 * Conversion.Val(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(obj5, null, "Value", new object[0], null, null, null))) * 7.85 / 1000.0, 3);
						obj5 = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 40 }, null, null, array2 = new bool[2] { true, false });
						if (array2[0])
						{
							i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
						}
						NewLateBinding.LateSetComplex(obj5, null, "Value", new object[1] { num45 }, null, null, OptimisticSet: false, RValueBase: true);
						obj5 = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 41 }, null, null, array2 = new bool[2] { true, false });
						if (array2[0])
						{
							i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
						}
						NewLateBinding.LateSetComplex(obj5, null, "Value", new object[1] { num47 }, null, null, OptimisticSet: false, RValueBase: true);
						obj5 = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 42 }, null, null, array2 = new bool[2] { true, false });
						if (array2[0])
						{
							i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
						}
						NewLateBinding.LateSetComplex(obj5, null, "Value", new object[1] { num49 }, null, null, OptimisticSet: false, RValueBase: true);
						obj5 = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 43 }, null, null, array2 = new bool[2] { true, false });
						if (array2[0])
						{
							i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
						}
						NewLateBinding.LateSetComplex(obj5, null, "Value", new object[1] { Math.Round(num45 + num47 + num49, 3) }, null, null, OptimisticSet: false, RValueBase: true);
					}
				}
				obj5 = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 24 }, null, null, array2 = new bool[2] { true, false });
				if (array2[0])
				{
					i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
				}
				if (Conversion.Val(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(obj5, null, "Value", new object[0], null, null, null))) == 0.0)
				{
					object[] obj55 = new object[2] { i, 43 };
					array = obj55;
					bool[] obj56 = new bool[2] { true, false };
					array2 = obj56;
					obj5 = NewLateBinding.LateGet(ws, null, "Cells", obj55, null, null, obj56);
					if (array2[0])
					{
						i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
					}
					if (Conversion.Val(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(obj5, null, "Value", new object[0], null, null, null))) == 0.0 && num4 != 671 && num4 != 672)
					{
						object[] obj57 = new object[2] { i, 7 };
						array = obj57;
						bool[] obj58 = new bool[2] { true, false };
						array2 = obj58;
						obj5 = NewLateBinding.LateGet(ws, null, "Cells", obj57, null, null, obj58);
						if (array2[0])
						{
							i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
						}
						double value = num3 * Conversion.Val(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(obj5, null, "Value", new object[0], null, null, null))) * 7850.0;
						obj5 = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 43 }, null, null, array2 = new bool[2] { true, false });
						if (array2[0])
						{
							i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
						}
						NewLateBinding.LateSetComplex(obj5, null, "Value", new object[1] { Math.Round(value, 3) }, null, null, OptimisticSet: false, RValueBase: true);
					}
				}
				if ((Operators.CompareString(left, "TUBE", TextCompare: false) == 0) | (Operators.CompareString(left, "USER DEFINED TUBE", TextCompare: false) == 0))
				{
					object[] obj59 = new object[2] { i, 11 };
					array = obj59;
					bool[] obj60 = new bool[2] { true, false };
					array2 = obj60;
					obj5 = NewLateBinding.LateGet(ws, null, "Cells", obj59, null, null, obj60);
					if (array2[0])
					{
						i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
					}
					double num50 = Conversion.Val(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(obj5, null, "Value", new object[0], null, null, null)));
					obj5 = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 12 }, null, null, array2 = new bool[2] { true, false });
					if (array2[0])
					{
						i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
					}
					double num51 = Conversion.Val(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(obj5, null, "Value", new object[0], null, null, null)));
					obj5 = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 13 }, null, null, array2 = new bool[2] { true, false });
					if (array2[0])
					{
						i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
					}
					double num52 = Conversion.Val(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(obj5, null, "Value", new object[0], null, null, null)));
					double num53 = (num51 + num52 - num50 * 2.0) * 2.0 * 1000.0;
					double num54 = num50 * 1000.0;
					obj5 = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 36 }, null, null, array2 = new bool[2] { true, false });
					if (array2[0])
					{
						i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
					}
					NewLateBinding.LateSetComplex(obj5, null, "Value", new object[1] { Math.Round(num53) }, null, null, OptimisticSet: false, RValueBase: true);
					obj5 = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 25 }, null, null, array2 = new bool[2] { true, false });
					if (array2[0])
					{
						i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
					}
					NewLateBinding.LateSetComplex(obj5, null, "Value", new object[1] { Math.Round(num54) }, null, null, OptimisticSet: false, RValueBase: true);
					double num55 = Math.Round(num3 * num53 * num54 * 7.85 / 1000.0, 3);
					obj5 = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 40 }, null, null, array2 = new bool[2] { true, false });
					if (array2[0])
					{
						i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
					}
					NewLateBinding.LateSetComplex(obj5, null, "Value", new object[1] { num55 }, null, null, OptimisticSet: false, RValueBase: true);
					obj5 = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 43 }, null, null, array2 = new bool[2] { true, false });
					if (array2[0])
					{
						i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
					}
					NewLateBinding.LateSetComplex(obj5, null, "Value", new object[1] { num55 }, null, null, OptimisticSet: false, RValueBase: true);
					obj5 = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 37 }, null, null, array2 = new bool[2] { true, false });
					if (array2[0])
					{
						i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
					}
					NewLateBinding.LateSetComplex(obj5, null, "Value", new object[1] { "TUBE " + Conversions.ToString(Math.Round(num51 * 1000.0)) + "x" + Conversions.ToString(Math.Round(num52 * 1000.0)) + "x" + Conversions.ToString(Math.Round(num50 * 1000.0, 1)) }, null, null, OptimisticSet: false, RValueBase: true);
				}
				if (!((Operators.CompareString(left, "PIPE", TextCompare: false) == 0) | (Operators.CompareString(left, "USER DEFINED PIPE", TextCompare: false) == 0)))
				{
					continue;
				}
				object[] obj61 = new object[2] { i, 11 };
				array = obj61;
				bool[] obj62 = new bool[2] { true, false };
				array2 = obj62;
				obj5 = NewLateBinding.LateGet(ws, null, "Cells", obj61, null, null, obj62);
				if (array2[0])
				{
					i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
				}
				double num56 = Conversion.Val(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(obj5, null, "Value", new object[0], null, null, null)));
				obj5 = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 12 }, null, null, array2 = new bool[2] { true, false });
				if (array2[0])
				{
					i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
				}
				double num57 = Conversion.Val(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(obj5, null, "Value", new object[0], null, null, null)));
				double num58 = Math.PI / 4.0 * (num56 * num56 - num57 * num57);
				double num59 = Math.Round(num3 * num58 * 7850.0, 3);
				obj5 = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 43 }, null, null, array2 = new bool[2] { true, false });
				if (array2[0])
				{
					i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
				}
				NewLateBinding.LateSetComplex(obj5, null, "Value", new object[1] { num59 }, null, null, OptimisticSet: false, RValueBase: true);
				if (num57 == 0.0)
				{
					object[] obj63 = new object[2] { i, 37 };
					array = obj63;
					bool[] obj64 = new bool[2] { true, false };
					array2 = obj64;
					obj5 = NewLateBinding.LateGet(ws, null, "Cells", obj63, null, null, obj64);
					if (array2[0])
					{
						i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
					}
					NewLateBinding.LateSetComplex(obj5, null, "Value", new object[1] { "ROD " + Conversions.ToString(Math.Round(num56 * 1000.0)) }, null, null, OptimisticSet: false, RValueBase: true);
				}
				else
				{
					object[] obj65 = new object[2] { i, 37 };
					array = obj65;
					bool[] obj66 = new bool[2] { true, false };
					array2 = obj66;
					obj5 = NewLateBinding.LateGet(ws, null, "Cells", obj65, null, null, obj66);
					if (array2[0])
					{
						i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
					}
					NewLateBinding.LateSetComplex(obj5, null, "Value", new object[1] { "PIPE " + Conversions.ToString(Math.Round(num56 * 1000.0)) + "*" + Conversions.ToString(Math.Round((num56 - num57) * 1000.0 / 2.0)) }, null, null, OptimisticSet: false, RValueBase: true);
				}
				obj5 = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 38 }, null, null, array2 = new bool[2] { true, false });
				if (array2[0])
				{
					i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
				}
				NewLateBinding.LateSetComplex(obj5, null, "Value", new object[1] { "" }, null, null, OptimisticSet: false, RValueBase: true);
				obj5 = NewLateBinding.LateGet(ws, null, "Cells", array = new object[2] { i, 39 }, null, null, array2 = new bool[2] { true, false });
				if (array2[0])
				{
					i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
				}
				NewLateBinding.LateSetComplex(obj5, null, "Value", new object[1] { "" }, null, null, OptimisticSet: false, RValueBase: true);
			}
			NewLateBinding.LateSetComplex(NewLateBinding.LateGet(NewLateBinding.LateGet(ws, null, "Range", new object[1] { "X7:AQ" + Conversions.ToString(num) }, null, null, null), null, "Borders", new object[0], null, null, null), null, "LineStyle", new object[1] { 1 }, null, null, OptimisticSet: false, RValueBase: true);
			NewLateBinding.LateCall(NewLateBinding.LateGet(NewLateBinding.LateGet(ws, null, "Range", new object[1] { "X:AQ" }, null, null, null), null, "Columns", new object[0], null, null, null), null, "AutoFit", new object[0], null, null, null, IgnoreReturn: true);
		}
	}

	private void UnprotectSheetAndWorkbook(object ws)
	{
		int try0001_dispatch = -1;
		int num2 = default(int);
		object objectValue = default(object);
		int num = default(int);
		int num3 = default(int);
		while (true)
		{
			try
			{
				/*Note: ILSpy has introduced the following switch to emulate a goto from catch-block to try-block*/;
				switch (try0001_dispatch)
				{
				default:
					num2 = 1;
					objectValue = RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(ws, null, "Parent", new object[0], null, null, null));
					goto IL_001f;
				case 181:
					{
						num = num2;
						switch ((num3 <= -2) ? 1 : num3)
						{
						case 1:
							break;
						default:
							goto end_IL_0001;
						}
						int num4 = num + 1;
						num = 0;
						switch (num4)
						{
						case 1:
							break;
						case 2:
							goto IL_001f;
						case 3:
							goto IL_0028;
						case 4:
							goto IL_0049;
						case 5:
							goto IL_0051;
						case 6:
							goto IL_005a;
						case 7:
							goto end_IL_0001_2;
						default:
							goto end_IL_0001;
						case 8:
							goto end_IL_0001_3;
						}
						goto default;
					}
					IL_0049:
					ProjectData.ClearProjectError();
					num3 = 0;
					goto IL_0051;
					IL_0051:
					ProjectData.ClearProjectError();
					num3 = -3;
					goto IL_005a;
					IL_0028:
					num2 = 3;
					NewLateBinding.LateCall(ws, null, "Unprotect", new object[1] { "2022" }, null, null, null, IgnoreReturn: true);
					goto IL_0049;
					IL_005a:
					num2 = 6;
					NewLateBinding.LateCall(objectValue, null, "Unprotect", new object[1] { "2022" }, null, null, null, IgnoreReturn: true);
					break;
					IL_001f:
					ProjectData.ClearProjectError();
					num3 = -2;
					goto IL_0028;
					end_IL_0001_2:
					break;
				}
				ProjectData.ClearProjectError();
				num3 = 0;
				break;
				end_IL_0001:;
			}
			catch (object obj) when (obj is Exception && num3 != 0 && num == 0)
			{
				ProjectData.SetProjectError((Exception)obj);
				try0001_dispatch = 181;
				continue;
			}
			throw ProjectData.CreateProjectError(-2146828237);
			continue;
			end_IL_0001_3:
			break;
		}
		if (num != 0)
		{
			ProjectData.ClearProjectError();
		}
	}

	private void ProtectSheetAndWorkbook(object ws)
	{
		int try0001_dispatch = -1;
		int num2 = default(int);
		object objectValue = default(object);
		int num = default(int);
		int num3 = default(int);
		while (true)
		{
			try
			{
				/*Note: ILSpy has introduced the following switch to emulate a goto from catch-block to try-block*/;
				switch (try0001_dispatch)
				{
				default:
					num2 = 1;
					objectValue = RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(ws, null, "Parent", new object[0], null, null, null));
					goto IL_001f;
				case 744:
					{
						num = num2;
						switch ((num3 <= -2) ? 1 : num3)
						{
						case 1:
							break;
						default:
							goto end_IL_0001;
						}
						int num4 = num + 1;
						num = 0;
						switch (num4)
						{
						case 1:
							break;
						case 2:
							goto IL_001f;
						case 3:
							goto IL_0028;
						case 4:
							goto IL_0049;
						case 5:
							goto IL_006a;
						case 6:
							goto IL_0072;
						case 7:
							goto IL_00a8;
						case 8:
							goto IL_00e6;
						case 9:
							goto IL_020f;
						case 10:
							goto IL_0267;
						case 11:
							goto end_IL_0001_2;
						default:
							goto end_IL_0001;
						case 12:
						case 13:
							goto end_IL_0001_3;
						}
						goto default;
					}
					IL_00e6:
					num2 = 8;
					NewLateBinding.LateCall(ws, null, "Protect", new object[14]
					{
						"2022", true, true, true, false, false, false, false, false, false,
						false, false, false, false
					}, new string[14]
					{
						"Password", "DrawingObjects", "Contents", "Scenarios", "AllowFormattingCells", "AllowFormattingColumns", "AllowFormattingRows", "AllowInsertingColumns", "AllowInsertingRows", "AllowDeletingColumns",
						"AllowDeletingRows", "AllowSorting", "AllowFiltering", "AllowUsingPivotTables"
					}, null, null, IgnoreReturn: true);
					goto IL_020f;
					IL_020f:
					num2 = 9;
					NewLateBinding.LateCall(objectValue, null, "Protect", new object[3] { "2022", true, false }, new string[3] { "Password", "Structure", "Windows" }, null, null, IgnoreReturn: true);
					goto IL_0267;
					IL_00a8:
					num2 = 7;
					NewLateBinding.LateSetComplex(NewLateBinding.LateGet(ws, null, "Range", new object[1] { "AS2:AS5,AR2,AW2" }, null, null, null), null, "Locked", new object[1] { false }, null, null, OptimisticSet: false, RValueBase: true);
					goto IL_00e6;
					IL_0267:
					num2 = 10;
					if (!Conversions.ToBoolean(Operators.NotObject(NewLateBinding.LateGet(objectValue, null, "ProtectStructure", new object[0], null, null, null))))
					{
						goto end_IL_0001_3;
					}
					break;
					IL_001f:
					ProjectData.ClearProjectError();
					num3 = -2;
					goto IL_0028;
					IL_0028:
					num2 = 3;
					NewLateBinding.LateCall(ws, null, "Unprotect", new object[1] { "2022" }, null, null, null, IgnoreReturn: true);
					goto IL_0049;
					IL_0049:
					num2 = 4;
					NewLateBinding.LateCall(objectValue, null, "Unprotect", new object[1] { "2022" }, null, null, null, IgnoreReturn: true);
					goto IL_006a;
					IL_006a:
					ProjectData.ClearProjectError();
					num3 = 0;
					goto IL_0072;
					IL_0072:
					num2 = 6;
					NewLateBinding.LateSetComplex(NewLateBinding.LateGet(ws, null, "Cells", new object[0], null, null, null), null, "Locked", new object[1] { true }, null, null, OptimisticSet: false, RValueBase: true);
					goto IL_00a8;
					end_IL_0001_2:
					break;
				}
				num2 = 11;
				throw new Exception("Workbook structure protection failed — cannot lock sheet add/delete.");
				end_IL_0001:;
			}
			catch (object obj) when (obj is Exception && num3 != 0 && num == 0)
			{
				ProjectData.SetProjectError((Exception)obj);
				try0001_dispatch = 744;
				continue;
			}
			throw ProjectData.CreateProjectError(-2146828237);
			continue;
			end_IL_0001_3:
			break;
		}
		if (num != 0)
		{
			ProjectData.ClearProjectError();
		}
	}
}
