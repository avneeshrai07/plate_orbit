using System;
using System.IO;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Threading;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace RaghavStaadExtractor;

[ComVisible(true)]
[Guid("A1234567-B89C-4DEF-0123-456789ABCDEF")]
public class SaveSheetsManager
{
	public void Save_Sheet1_BeamOutput()
	{
		Application application = null;
		Workbook workbook = null;
		Worksheet worksheet = null;
		Workbook workbook2 = null;
		Worksheet worksheet2 = null;
		int i = 1;
		try
		{
			application = GetExcelInstance();
			if (application == null)
			{
				return;
			}
			application.ScreenUpdating = false;
			application.DisplayAlerts = false;
			application.EnableEvents = false;
			workbook = application.ActiveWorkbook;
			if (workbook == null)
			{
				Interaction.MsgBox("No active workbook found.", MsgBoxStyle.Critical);
				return;
			}
			if (!(workbook.Sheets["Sheet1"] is Worksheet worksheet3))
			{
				Interaction.MsgBox("Sheet1 not found in the active workbook.", MsgBoxStyle.Critical);
				return;
			}
			SafeUnprotectSheet(worksheet3);
			string text = Path.Combine(workbook.Path, "PlateNovaBeamOutput");
			if (!Directory.Exists(text))
			{
				Directory.CreateDirectory(text);
			}
			string text2 = GetCellTextSafely(((_Worksheet)worksheet3).get_Range((object)"E3", RuntimeHelpers.GetObjectValue(Missing.Value)));
			if (string.IsNullOrWhiteSpace(text2))
			{
				text2 = "Unnamed";
			}
			string text3 = text2 + " BEAM OUTPUT.xlsx";
			string text4 = Path.Combine(text, text3);
			if (File.Exists(text4))
			{
				MsgBoxResult msgBoxResult = Interaction.MsgBox("File already exists. Overwrite?", MsgBoxStyle.YesNoCancel | MsgBoxStyle.Question);
				if (msgBoxResult == MsgBoxResult.Yes)
				{
					try
					{
						File.Delete(text4);
					}
					catch (Exception ex)
					{
						ProjectData.SetProjectError(ex);
						Exception ex2 = ex;
						Interaction.MsgBox("Error deleting existing file: " + ex2.Message, MsgBoxStyle.Critical);
						ProjectData.ClearProjectError();
						return;
					}
				}
				else
				{
					if (msgBoxResult != MsgBoxResult.No)
					{
						goto IL_03ee;
					}
					for (; File.Exists(Path.Combine(text, text2 + " BEAM OUTPUT (" + Conversions.ToString(i) + ").xlsx")); i = checked(i + 1))
					{
					}
					text3 = text2 + " BEAM OUTPUT (" + Conversions.ToString(i) + ").xlsx";
					text4 = Path.Combine(text, text3);
				}
			}
			try
			{
				worksheet3.Copy(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
				workbook2 = application.Workbooks[application.Workbooks.Count];
				worksheet2 = (Worksheet)workbook2.Sheets[1];
			}
			catch (Exception ex3)
			{
				ProjectData.SetProjectError(ex3);
				Exception ex4 = ex3;
				Interaction.MsgBox("Error copying sheet: " + ex4.Message, MsgBoxStyle.Critical);
				ProjectData.ClearProjectError();
				goto IL_03ee;
			}
			try
			{
				SafeUnprotectSheet(worksheet2);
				SafeDeleteRange(((_Worksheet)worksheet2).get_Range((object)"F:AJ", RuntimeHelpers.GetObjectValue(Missing.Value)));
				SafeDeleteRange(((_Worksheet)worksheet2).get_Range((object)"M:N", RuntimeHelpers.GetObjectValue(Missing.Value)));
				SafeDeleteShapes(worksheet2);
				string text5 = Interaction.InputBox("Enter Company Name:", "Company Name");
				if (!string.IsNullOrWhiteSpace(text5))
				{
					UpdateCompanyName(worksheet2, text5);
				}
			}
			catch (Exception ex5)
			{
				ProjectData.SetProjectError(ex5);
				Exception ex6 = ex5;
				Interaction.MsgBox("Error modifying copied sheet: " + ex6.Message, MsgBoxStyle.Critical);
				ProjectData.ClearProjectError();
			}
			try
			{
				workbook2.SaveAs(text4, XlFileFormat.xlOpenXMLWorkbook, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), XlSaveAsAccessMode.xlNoChange, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
				workbook2.Close(false, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
				workbook2 = null;
			}
			catch (Exception ex7)
			{
				ProjectData.SetProjectError(ex7);
				Exception ex8 = ex7;
				Interaction.MsgBox("Error saving the workbook: " + ex8.Message, MsgBoxStyle.Critical);
				ProjectData.ClearProjectError();
			}
			Interaction.MsgBox("Sheet1 saved successfully as '" + text3 + "' in folder:\r\n" + text, MsgBoxStyle.Information);
			goto IL_03ee;
			IL_03ee:
			if (worksheet2 != null)
			{
				Marshal.ReleaseComObject(worksheet2);
			}
			if (workbook2 != null)
			{
				workbook2.Close(false, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
				Marshal.ReleaseComObject(workbook2);
			}
			SafeProtectSheet(worksheet3);
			if (worksheet3 != null)
			{
				Marshal.ReleaseComObject(worksheet3);
			}
			if (workbook != null)
			{
				Marshal.ReleaseComObject(workbook);
			}
			if (application != null)
			{
				application.ScreenUpdating = true;
				application.DisplayAlerts = true;
				application.EnableEvents = true;
			}
		}
		catch (Exception ex9)
		{
			ProjectData.SetProjectError(ex9);
			Exception ex10 = ex9;
			Interaction.MsgBox("Unexpected error: " + ex10.Message, MsgBoxStyle.Critical);
			ProjectData.ClearProjectError();
		}
		finally
		{
			if (application != null)
			{
				Marshal.ReleaseComObject(application);
			}
		}
	}

	public void Save_Sheet2_PlateSummary()
	{
		Application application = null;
		Workbook workbook = null;
		Worksheet worksheet = null;
		Workbook workbook2 = null;
		Worksheet worksheet2 = null;
		try
		{
			application = GetRunningExcel();
			if (application == null)
			{
				Interaction.MsgBox("Could not connect to Excel application.", MsgBoxStyle.Critical);
				return;
			}
			Application application2 = application;
			application2.ScreenUpdating = false;
			application2.DisplayAlerts = false;
			application2.EnableEvents = false;
			application2.Calculation = XlCalculation.xlCalculationManual;
			application2.AskToUpdateLinks = false;
			application2 = null;
			workbook = application.ActiveWorkbook;
			if (workbook == null)
			{
				Interaction.MsgBox("No active workbook found.", MsgBoxStyle.Critical);
			}
			else
			{
				worksheet = workbook.Sheets["Sheet2"] as Worksheet;
				if (worksheet == null)
				{
					Interaction.MsgBox("'Sheet2' does not exist in the active workbook.", MsgBoxStyle.Critical);
				}
				else
				{
					string text = Path.Combine(workbook.Path, "PlateNovaPlateSummary");
					try
					{
						if (!Directory.Exists(text))
						{
							Directory.CreateDirectory(text);
						}
					}
					catch (Exception ex)
					{
						ProjectData.SetProjectError(ex);
						Exception ex2 = ex;
						Interaction.MsgBox("Error creating output directory: " + ex2.Message, MsgBoxStyle.Critical);
						ProjectData.ClearProjectError();
						goto IL_0511;
					}
					string text2 = "Unnamed";
					string text3 = "";
					try
					{
						if (((_Worksheet)worksheet).get_Range((object)"C3", RuntimeHelpers.GetObjectValue(Missing.Value)).get_Value(RuntimeHelpers.GetObjectValue(Missing.Value)) != null)
						{
							text2 = ((_Worksheet)worksheet).get_Range((object)"C3", RuntimeHelpers.GetObjectValue(Missing.Value)).get_Value(RuntimeHelpers.GetObjectValue(Missing.Value)).ToString()
								.Trim();
							if (string.IsNullOrWhiteSpace(text2))
							{
								text2 = "Unnamed";
							}
						}
						if (((_Worksheet)worksheet).get_Range((object)"C2", RuntimeHelpers.GetObjectValue(Missing.Value)).get_Value(RuntimeHelpers.GetObjectValue(Missing.Value)) != null)
						{
							text3 = ((_Worksheet)worksheet).get_Range((object)"C2", RuntimeHelpers.GetObjectValue(Missing.Value)).get_Value(RuntimeHelpers.GetObjectValue(Missing.Value)).ToString()
								.Trim();
						}
					}
					catch (Exception projectError)
					{
						ProjectData.SetProjectError(projectError);
						ProjectData.ClearProjectError();
					}
					string text4 = ((!string.IsNullOrWhiteSpace(text3)) ? (text2 + " " + text3 + " MATERIAL SUMMARY LIST.xlsx") : (text2 + " MATERIAL SUMMARY LIST.xlsx"));
					string text5 = Path.Combine(text, text4);
					int num = 1;
					if (File.Exists(text5))
					{
						MsgBoxResult msgBoxResult = Interaction.MsgBox("File '" + text4 + "' already exists. Overwrite?", MsgBoxStyle.YesNoCancel | MsgBoxStyle.Question);
						if (msgBoxResult == MsgBoxResult.Cancel)
						{
							goto IL_0511;
						}
						if (msgBoxResult != MsgBoxResult.Yes)
						{
							if (msgBoxResult == MsgBoxResult.No)
							{
								do
								{
									text4 = ((!string.IsNullOrWhiteSpace(text3)) ? (text2 + " " + text3 + " MATERIAL SUMMARY LIST (" + Conversions.ToString(num) + ").xlsx") : (text2 + " MATERIAL SUMMARY LIST (" + Conversions.ToString(num) + ").xlsx"));
									text5 = Path.Combine(text, text4);
									num = checked(num + 1);
								}
								while (File.Exists(text5));
							}
						}
						else
						{
							try
							{
								File.Delete(text5);
							}
							catch (Exception ex3)
							{
								ProjectData.SetProjectError(ex3);
								Exception ex4 = ex3;
								Interaction.MsgBox("Could not delete existing file: " + ex4.Message, MsgBoxStyle.Critical);
								ProjectData.ClearProjectError();
								goto IL_0511;
							}
						}
					}
					try
					{
						worksheet.Copy(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
						Thread.Sleep(500);
						workbook2 = application.Workbooks[application.Workbooks.Count];
						worksheet2 = (Worksheet)workbook2.Sheets[1];
						SafeRemoveShapes(worksheet2);
						int num2 = 0;
						int num3 = 3;
						bool flag = false;
						while (true)
						{
							if (num2 < num3 && !flag)
							{
								try
								{
									workbook2.SaveAs(text5, XlFileFormat.xlOpenXMLWorkbook, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), XlSaveAsAccessMode.xlNoChange, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
									flag = true;
								}
								catch (Exception ex5)
								{
									ProjectData.SetProjectError(ex5);
									Exception ex6 = ex5;
									num2 = checked(num2 + 1);
									if (num2 >= num3)
									{
										Interaction.MsgBox("Failed to save after " + Conversions.ToString(num3) + " attempts: " + ex6.Message, MsgBoxStyle.Critical);
										ProjectData.ClearProjectError();
										break;
									}
									Thread.Sleep(1000);
									ProjectData.ClearProjectError();
								}
								continue;
							}
							Interaction.MsgBox("Sheet2 saved successfully as '" + text4 + "' in:\r\n" + text, MsgBoxStyle.Information);
							break;
						}
					}
					catch (Exception ex7)
					{
						ProjectData.SetProjectError(ex7);
						Exception ex8 = ex7;
						Interaction.MsgBox("Error during save operation: " + ex8.Message, MsgBoxStyle.Critical);
						ProjectData.ClearProjectError();
					}
				}
			}
			goto IL_0511;
			IL_0511:
			if (workbook2 != null)
			{
				try
				{
					workbook2.Close(false, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
					Marshal.ReleaseComObject(worksheet2);
					Marshal.ReleaseComObject(workbook2);
				}
				catch (Exception projectError2)
				{
					ProjectData.SetProjectError(projectError2);
					ProjectData.ClearProjectError();
				}
			}
			if (worksheet != null)
			{
				Marshal.ReleaseComObject(worksheet);
			}
			if (workbook != null)
			{
				Marshal.ReleaseComObject(workbook);
			}
			if (application != null)
			{
				try
				{
					Application application3 = application;
					application3.ScreenUpdating = true;
					application3.DisplayAlerts = true;
					application3.EnableEvents = true;
					application3.Calculation = XlCalculation.xlCalculationAutomatic;
					application3 = null;
					Marshal.ReleaseComObject(application);
					return;
				}
				catch (Exception projectError3)
				{
					ProjectData.SetProjectError(projectError3);
					ProjectData.ClearProjectError();
					return;
				}
			}
		}
		catch (Exception ex9)
		{
			ProjectData.SetProjectError(ex9);
			Exception ex10 = ex9;
			Interaction.MsgBox("Unexpected error: " + ex10.Message, MsgBoxStyle.Critical);
			ProjectData.ClearProjectError();
		}
	}

	private Application GetRunningExcel()
	{
		Application result;
		try
		{
			result = (Application)Marshal.GetActiveObject("Excel.Application");
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			try
			{
				result = (Application)Activator.CreateInstance(Marshal.GetTypeFromCLSID(new Guid("00024500-0000-0000-C000-000000000046")));
				ProjectData.ClearProjectError();
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				Interaction.MsgBox("Could not start Excel: " + ex2.Message, MsgBoxStyle.Critical);
				result = null;
				ProjectData.ClearProjectError();
			}
		}
		return result;
	}

	private Application GetExcelInstance()
	{
		Application result;
		try
		{
			result = (Application)Marshal.GetActiveObject("Excel.Application");
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			Interaction.MsgBox("Excel is not running or could not be accessed: " + ex2.Message, MsgBoxStyle.Critical);
			result = null;
			ProjectData.ClearProjectError();
		}
		return result;
	}

	private void SafeUnprotectSheet(Worksheet ws)
	{
		if (ws == null)
		{
			return;
		}
		try
		{
			ws.Unprotect("2022");
			if (ws.Parent != null)
			{
				Workbook workbook = (Workbook)ws.Parent;
				workbook.Unprotect("2022");
				Marshal.ReleaseComObject(workbook);
			}
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
	}

	private void SafeProtectSheet(Worksheet ws)
	{
		if (ws == null)
		{
			return;
		}
		try
		{
			ws.Protect("2022", true, true, true, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
			if (ws.Parent != null)
			{
				Workbook workbook = (Workbook)ws.Parent;
				workbook.Protect("2022", true, false);
				Marshal.ReleaseComObject(workbook);
			}
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
	}

	private void SafeDeleteRange(Range rng)
	{
		if (rng == null)
		{
			return;
		}
		try
		{
			rng.Delete(RuntimeHelpers.GetObjectValue(Missing.Value));
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
		finally
		{
			if (rng != null)
			{
				Marshal.ReleaseComObject(rng);
			}
		}
	}

	private void SafeRemoveShapes(Worksheet ws)
	{
		if (ws == null)
		{
			return;
		}
		checked
		{
			try
			{
				int count = ws.Shapes.Count;
				for (int i = count; i >= 1; i += -1)
				{
					Shape shape = null;
					try
					{
						shape = ws.Shapes.Item(i);
						shape.Delete();
					}
					catch (Exception projectError)
					{
						ProjectData.SetProjectError(projectError);
						ProjectData.ClearProjectError();
					}
					finally
					{
						if (shape != null)
						{
							Marshal.ReleaseComObject(shape);
						}
					}
				}
			}
			catch (Exception projectError2)
			{
				ProjectData.SetProjectError(projectError2);
				ProjectData.ClearProjectError();
			}
			try
			{
				int num = Conversions.ToInteger(NewLateBinding.LateGet(ws.OLEObjects(RuntimeHelpers.GetObjectValue(Missing.Value)), null, "Count", new object[0], null, null, null));
				for (int j = num; j >= 1; j += -1)
				{
					OLEObject oLEObject = null;
					try
					{
						object instance = ws.OLEObjects(RuntimeHelpers.GetObjectValue(Missing.Value));
						object[] obj = new object[1] { j };
						object[] array = obj;
						bool[] obj2 = new bool[1] { true };
						bool[] array2 = obj2;
						object obj3 = NewLateBinding.LateGet(instance, null, "Item", obj, null, null, obj2);
						if (array2[0])
						{
							j = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
						}
						oLEObject = (OLEObject)obj3;
						oLEObject.Delete();
					}
					catch (Exception projectError3)
					{
						ProjectData.SetProjectError(projectError3);
						ProjectData.ClearProjectError();
					}
					finally
					{
						if (oLEObject != null)
						{
							Marshal.ReleaseComObject(oLEObject);
						}
					}
				}
			}
			catch (Exception projectError4)
			{
				ProjectData.SetProjectError(projectError4);
				ProjectData.ClearProjectError();
			}
		}
	}

	private void SafeDeleteShapes(Worksheet ws)
	{
		if (ws == null)
		{
			return;
		}
		try
		{
			foreach (Shape shape in ws.Shapes)
			{
				try
				{
					shape.Delete();
				}
				catch (Exception projectError)
				{
					ProjectData.SetProjectError(projectError);
					ProjectData.ClearProjectError();
				}
				finally
				{
					Marshal.ReleaseComObject(shape);
				}
			}
		}
		catch (Exception projectError2)
		{
			ProjectData.SetProjectError(projectError2);
			ProjectData.ClearProjectError();
		}
	}

	private string GetCellTextSafely(Range rng)
	{
		string result;
		if (rng == null)
		{
			result = string.Empty;
		}
		else
		{
			try
			{
				result = rng.Text.ToString().Trim();
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				result = string.Empty;
				ProjectData.ClearProjectError();
			}
			finally
			{
				if (rng != null)
				{
					Marshal.ReleaseComObject(rng);
				}
			}
		}
		return result;
	}

	private void UpdateCompanyName(Worksheet ws, string companyName)
	{
		if (ws == null)
		{
			return;
		}
		try
		{
			Range range = ((_Worksheet)ws).get_Range((object)"A1", RuntimeHelpers.GetObjectValue(Missing.Value));
			if (Conversions.ToBoolean(range.MergeCells))
			{
				Range mergeArea = range.MergeArea;
				string cell = mergeArea.get_Address(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
				mergeArea.UnMerge();
				((_Worksheet)ws).get_Range((object)cell, RuntimeHelpers.GetObjectValue(Missing.Value)).set_Value(RuntimeHelpers.GetObjectValue(Missing.Value), (object)companyName);
				((_Worksheet)ws).get_Range((object)cell, RuntimeHelpers.GetObjectValue(Missing.Value)).Merge(RuntimeHelpers.GetObjectValue(Missing.Value));
				Range range2 = ((_Worksheet)ws).get_Range((object)cell, RuntimeHelpers.GetObjectValue(Missing.Value));
				range2.HorizontalAlignment = XlHAlign.xlHAlignCenter;
				range2.VerticalAlignment = XlVAlign.xlVAlignCenter;
				range2.Font.Bold = true;
				range2.Font.Size = 14;
				range2.Locked = true;
				range2 = null;
				Marshal.ReleaseComObject(mergeArea);
			}
			else
			{
				range.set_Value(RuntimeHelpers.GetObjectValue(Missing.Value), (object)companyName);
				range.HorizontalAlignment = XlHAlign.xlHAlignCenter;
				range.VerticalAlignment = XlVAlign.xlVAlignCenter;
				range.Font.Bold = true;
				range.Font.Size = 14;
				range.Locked = true;
			}
			Marshal.ReleaseComObject(range);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}
}
