// ===== RaghavStaadExtractor.Unlock =====

// Decompiled member-by-member from RaghavStaadExtractor.dll



public Unlock()
{
}


using System;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

public void UnlockSheetRange()
{
	Application application = null;
	Workbook workbook = null;
	Worksheet worksheet = null;
	try
	{
		application = (Application)Marshal.GetActiveObject("Excel.Application");
		workbook = application.ActiveWorkbook;
		worksheet = (Worksheet)workbook.Sheets["Sheet1"];
		worksheet.Unprotect("2022");
		Range range = ((_Worksheet)worksheet).get_Range((object)"E2:AQ5", RuntimeHelpers.GetObjectValue(Missing.Value));
		range.Locked = false;
		worksheet.Protect("2022", true, true, true, true, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
	}
	catch (Exception ex)
	{
		ProjectData.SetProjectError(ex);
		Exception ex2 = ex;
		Interaction.MsgBox("Error unlocking specific range: " + ex2.Message);
		ProjectData.ClearProjectError();
	}
	finally
	{
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
			Marshal.ReleaseComObject(application);
		}
		GC.Collect();
		GC.WaitForPendingFinalizers();
	}
}


using System;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

public void ToggleUniqueFilter()
{
	Application application = null;
	Workbook workbook = null;
	Worksheet worksheet = null;
	try
	{
		application = (Application)Marshal.GetActiveObject("Excel.Application");
		workbook = application.ActiveWorkbook;
		worksheet = (Worksheet)workbook.Sheets["Sheet1"];
		try
		{
			worksheet.Unprotect("2022");
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
		long num = Conversions.ToLong(NewLateBinding.LateGet(NewLateBinding.LateGet(worksheet.Cells[worksheet.Rows.Count, "E"], null, "End", new object[1] { XlDirection.xlUp }, null, null, null), null, "Row", new object[0], null, null, null));
		bool flag = false;
		long num2 = num;
		for (long num3 = 9L; num3 <= num2; num3 = checked(num3 + 1))
		{
			if (NewLateBinding.LateGet(worksheet.Cells[num3, 5], null, "Value", new object[0], null, null, null) != null && NewLateBinding.LateGet(worksheet.Cells[num3, 5], null, "Value", new object[0], null, null, null).ToString().Contains("FILTERED DATA"))
			{
				flag = true;
				break;
			}
		}
		if (flag)
		{
			RemoveFilteredData(worksheet);
		}
		else
		{
			CreateFilteredData(worksheet);
		}
	}
	catch (Exception ex)
	{
		ProjectData.SetProjectError(ex);
		Exception ex2 = ex;
		Interaction.MsgBox("Error in ToggleUniqueFilter: " + ex2.Message);
		ProjectData.ClearProjectError();
	}
	finally
	{
		try
		{
			worksheet?.Protect("2022", RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), true, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
		}
		catch (Exception projectError2)
		{
			ProjectData.SetProjectError(projectError2);
			ProjectData.ClearProjectError();
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
			Marshal.ReleaseComObject(application);
		}
		GC.Collect();
		GC.WaitForPendingFinalizers();
	}
}


using System.Collections.Generic;
using System.Reflection;
using System.Runtime.CompilerServices;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

private void CreateFilteredData(Worksheet sheet)
{
	long num = Conversions.ToLong(NewLateBinding.LateGet(NewLateBinding.LateGet(sheet.Cells[sheet.Rows.Count, "E"], null, "End", new object[1] { XlDirection.xlUp }, null, null, null), null, "Row", new object[0], null, null, null));
	if (num < 9)
	{
		Interaction.MsgBox("No data found from row 9 onwards in column E");
		return;
	}
	long num2 = num;
	checked
	{
		for (long num3 = 9L; num3 <= num2; num3++)
		{
			NewLateBinding.LateSetComplex(sheet.Rows[num3, RuntimeHelpers.GetObjectValue(Missing.Value)], null, "Hidden", new object[1] { true }, null, null, OptimisticSet: false, RValueBase: true);
		}
		Dictionary<string, long> dictionary = new Dictionary<string, long>();
		long num4 = num;
		for (long num5 = 9L; num5 <= num4; num5++)
		{
			string text = Conversions.ToString(NewLateBinding.LateGet(sheet.Cells[num5, 5], null, "Value", new object[0], null, null, null));
			if (!string.IsNullOrEmpty(text) && text.StartsWith("R") && !dictionary.ContainsKey(text))
			{
				dictionary.Add(text, num5);
			}
		}
		List<KeyValuePair<string, long>> list = new List<KeyValuePair<string, long>>(dictionary);
		list.Sort((KeyValuePair<string, long> x, KeyValuePair<string, long> y) => CompareAlphanumeric(x.Key, y.Key));
		long num6 = num + 1;
		NewLateBinding.LateSetComplex(sheet.Cells[num6, 5], null, "Value", new object[1] { "UNIQUE PROPERTIES FILTERED DATA - DO NOT EDIT" }, null, null, OptimisticSet: false, RValueBase: true);
		((_Worksheet)sheet).get_Range(RuntimeHelpers.GetObjectValue(sheet.Cells[num6, 5]), RuntimeHelpers.GetObjectValue(sheet.Cells[num6, 43])).Interior.Color = Information.RGB(255, 255, 255);
		long num7 = num6 + 1;
		foreach (KeyValuePair<string, long> item in list)
		{
			long value = item.Value;
			Range range = ((_Worksheet)sheet).get_Range(RuntimeHelpers.GetObjectValue(sheet.Cells[value, 1]), RuntimeHelpers.GetObjectValue(sheet.Cells[value, 43]));
			Range destination = ((_Worksheet)sheet).get_Range(RuntimeHelpers.GetObjectValue(sheet.Cells[num7, 1]), RuntimeHelpers.GetObjectValue(sheet.Cells[num7, 43]));
			range.Copy(destination);
			num7++;
		}
		UnlockCopyColumns(sheet, num7 - 1);
		Interaction.MsgBox("Filter Applied - Unique rows displayed below original data");
	}
}


using System.Reflection;
using System.Runtime.CompilerServices;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

private void RemoveFilteredData(Worksheet sheet)
{
	long num = Conversions.ToLong(NewLateBinding.LateGet(NewLateBinding.LateGet(sheet.Cells[sheet.Rows.Count, "E"], null, "End", new object[1] { XlDirection.xlUp }, null, null, null), null, "Row", new object[0], null, null, null));
	long num2 = 0L;
	long num3 = num;
	for (long num4 = 9L; num4 <= num3; num4 = checked(num4 + 1))
	{
		if (NewLateBinding.LateGet(sheet.Cells[num4, 5], null, "Value", new object[0], null, null, null) != null && NewLateBinding.LateGet(sheet.Cells[num4, 5], null, "Value", new object[0], null, null, null).ToString().Contains("UNIQUE PROPERTIES"))
		{
			num2 = num4;
			break;
		}
	}
	if (num2 == 0)
	{
		sheet.Rows.Hidden = false;
		Interaction.MsgBox("Filter Removed");
		return;
	}
	((_Worksheet)sheet).get_Range(RuntimeHelpers.GetObjectValue(sheet.Rows[num2, RuntimeHelpers.GetObjectValue(Missing.Value)]), RuntimeHelpers.GetObjectValue(sheet.Rows[num, RuntimeHelpers.GetObjectValue(Missing.Value)])).Delete(RuntimeHelpers.GetObjectValue(Missing.Value));
	sheet.Rows.Hidden = false;
	LockCopyColumns(sheet);
	Interaction.MsgBox("Filter Removed - Filtered data deleted");
}


using System;
using System.IO;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

public void SaveAs_filter()
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
		string text = Path.Combine(workbook.Path, "PROPERTY");
		if (!Directory.Exists(text))
		{
			Directory.CreateDirectory(text);
		}
		string text2 = GetCellTextSafely(((_Worksheet)worksheet3).get_Range((object)"E3", RuntimeHelpers.GetObjectValue(Missing.Value)));
		if (string.IsNullOrWhiteSpace(text2))
		{
			text2 = "Unnamed";
		}
		string path = text2 + " PROPERTIES.xlsx";
		string text3 = Path.Combine(text, path);
		if (File.Exists(text3))
		{
			MsgBoxResult msgBoxResult = Interaction.MsgBox("File already exists. Overwrite?", MsgBoxStyle.YesNoCancel | MsgBoxStyle.Question);
			if (msgBoxResult == MsgBoxResult.Yes)
			{
				try
				{
					File.Delete(text3);
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
					goto IL_0358;
				}
				for (; File.Exists(Path.Combine(text, text2 + " PROPERTIES (" + Conversions.ToString(i) + ").xlsx")); i = checked(i + 1))
				{
				}
				path = text2 + " PROPERTIES (" + Conversions.ToString(i) + ").xlsx";
				text3 = Path.Combine(text, path);
			}
		}
		workbook2 = application.Workbooks.Add(RuntimeHelpers.GetObjectValue(Missing.Value));
		worksheet2 = (Worksheet)workbook2.Sheets[1];
		worksheet2.Name = "Properties";
		try
		{
			Range range = ((_Worksheet)worksheet2).get_Range((object)"B1:D1", RuntimeHelpers.GetObjectValue(Missing.Value));
			range.Merge(RuntimeHelpers.GetObjectValue(Missing.Value));
			range.HorizontalAlignment = XlHAlign.xlHAlignCenter;
			range.VerticalAlignment = XlVAlign.xlVAlignCenter;
		}
		catch (Exception ex3)
		{
			ProjectData.SetProjectError(ex3);
			Exception ex4 = ex3;
			ProjectData.ClearProjectError();
		}
		CopyVisibleColumnsData(worksheet3, worksheet2, application);
		try
		{
			workbook2.SaveAs(text3, XlFileFormat.xlOpenXMLWorkbook, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), XlSaveAsAccessMode.xlNoChange, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
			workbook2.Close(false, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
			workbook2 = null;
		}
		catch (Exception ex5)
		{
			ProjectData.SetProjectError(ex5);
			Exception ex6 = ex5;
			Interaction.MsgBox("Error saving the workbook: " + ex6.Message, MsgBoxStyle.Critical);
			ProjectData.ClearProjectError();
		}
		Interaction.MsgBox("Unique Properties SaveAs Successfully");
		goto IL_0358;
		IL_0358:
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
	catch (Exception ex7)
	{
		ProjectData.SetProjectError(ex7);
		Exception ex8 = ex7;
		Interaction.MsgBox("Unexpected error: " + ex8.Message, MsgBoxStyle.Critical);
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


#define DEBUG
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Reflection;
using System.Runtime.CompilerServices;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

private void CopyVisibleColumnsData(Worksheet sourceSheet, Worksheet targetSheet, Application excelApp)
{
	checked
	{
		try
		{
			long num = Conversions.ToLong(NewLateBinding.LateGet(NewLateBinding.LateGet(sourceSheet.Cells[sourceSheet.Rows.Count, "E"], null, "End", new object[1] { XlDirection.xlUp }, null, null, null), null, "Row", new object[0], null, null, null));
			if (num < 7)
			{
				return;
			}
			long num2 = 1L;
			int[] array = new int[4] { 5, 37, 38, 39 };
			string[] array2 = new string[4] { "A", "B", "C", "D" };
			Dictionary<string, bool> dictionary = new Dictionary<string, bool>();
			long num3 = num;
			for (long num4 = 7L; num4 <= num3; num4++)
			{
				if (!Conversions.ToBoolean(Operators.NotObject(NewLateBinding.LateGet(sourceSheet.Rows[num4, RuntimeHelpers.GetObjectValue(Missing.Value)], null, "Hidden", new object[0], null, null, null))))
				{
					continue;
				}
				int num5 = array.Length - 1;
				for (int i = 0; i <= num5; i++)
				{
					Range range = (Range)sourceSheet.Cells[num4, array[i]];
					Range range2 = (Range)targetSheet.Cells[num2, i + 1];
					if (Conversions.ToBoolean(!Information.IsNothing(RuntimeHelpers.GetObjectValue(range.get_Value(RuntimeHelpers.GetObjectValue(Missing.Value)))) || Conversions.ToBoolean(range.MergeCells)))
					{
						range2.set_Value(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(range.get_Value(RuntimeHelpers.GetObjectValue(Missing.Value))));
					}
					Range range3 = range;
					range2.Font.Name = RuntimeHelpers.GetObjectValue(range3.Font.Name);
					range2.Font.Size = RuntimeHelpers.GetObjectValue(range3.Font.Size);
					range2.Font.Bold = RuntimeHelpers.GetObjectValue(range3.Font.Bold);
					range2.Font.Italic = RuntimeHelpers.GetObjectValue(range3.Font.Italic);
					range2.Font.Color = RuntimeHelpers.GetObjectValue(range3.Font.Color);
					range2.Interior.Color = RuntimeHelpers.GetObjectValue(range3.Interior.Color);
					range2.HorizontalAlignment = RuntimeHelpers.GetObjectValue(range3.HorizontalAlignment);
					range2.VerticalAlignment = RuntimeHelpers.GetObjectValue(range3.VerticalAlignment);
					range2.WrapText = RuntimeHelpers.GetObjectValue(range3.WrapText);
					range2.NumberFormat = RuntimeHelpers.GetObjectValue(range3.NumberFormat);
					range2.Borders[XlBordersIndex.xlEdgeLeft].LineStyle = RuntimeHelpers.GetObjectValue(range3.Borders[XlBordersIndex.xlEdgeLeft].LineStyle);
					range2.Borders[XlBordersIndex.xlEdgeLeft].Weight = RuntimeHelpers.GetObjectValue(range3.Borders[XlBordersIndex.xlEdgeLeft].Weight);
					range2.Borders[XlBordersIndex.xlEdgeRight].LineStyle = RuntimeHelpers.GetObjectValue(range3.Borders[XlBordersIndex.xlEdgeRight].LineStyle);
					range2.Borders[XlBordersIndex.xlEdgeRight].Weight = RuntimeHelpers.GetObjectValue(range3.Borders[XlBordersIndex.xlEdgeRight].Weight);
					range2.Borders[XlBordersIndex.xlEdgeTop].LineStyle = RuntimeHelpers.GetObjectValue(range3.Borders[XlBordersIndex.xlEdgeTop].LineStyle);
					range2.Borders[XlBordersIndex.xlEdgeTop].Weight = RuntimeHelpers.GetObjectValue(range3.Borders[XlBordersIndex.xlEdgeTop].Weight);
					range2.Borders[XlBordersIndex.xlEdgeBottom].LineStyle = RuntimeHelpers.GetObjectValue(range3.Borders[XlBordersIndex.xlEdgeBottom].LineStyle);
					range2.Borders[XlBordersIndex.xlEdgeBottom].Weight = RuntimeHelpers.GetObjectValue(range3.Borders[XlBordersIndex.xlEdgeBottom].Weight);
					range3 = null;
					if (!Conversions.ToBoolean(range.MergeCells))
					{
						continue;
					}
					try
					{
						Range mergeArea = range.MergeArea;
						string key = $"{num4}_{array[i]}";
						if (dictionary.ContainsKey(key))
						{
							continue;
						}
						dictionary.Add(key, value: true);
						int num6 = 0;
						long val = mergeArea.Row;
						long num7 = mergeArea.Row + mergeArea.Rows.Count - 1;
						long num8 = Math.Max(val, 7L);
						long num9 = num7;
						for (long num10 = num8; num10 <= num9; num10++)
						{
							if (Conversions.ToBoolean(Operators.NotObject(NewLateBinding.LateGet(sourceSheet.Rows[num10, RuntimeHelpers.GetObjectValue(Missing.Value)], null, "Hidden", new object[0], null, null, null))))
							{
								num6++;
							}
						}
						if (num6 > 1)
						{
							Range range4 = ((_Worksheet)targetSheet).get_Range((object)range2, (object)range2.get_Offset((object)(num6 - 1), (object)0));
							if (Conversions.ToBoolean(Operators.NotObject(range4.MergeCells)))
							{
								range4.Merge(RuntimeHelpers.GetObjectValue(Missing.Value));
								range4.HorizontalAlignment = RuntimeHelpers.GetObjectValue(range.HorizontalAlignment);
								range4.VerticalAlignment = RuntimeHelpers.GetObjectValue(range.VerticalAlignment);
							}
						}
					}
					catch (Exception ex)
					{
						ProjectData.SetProjectError(ex);
						Exception ex2 = ex;
						Debug.WriteLine("Merge error: " + ex2.Message);
						ProjectData.ClearProjectError();
					}
				}
				try
				{
					NewLateBinding.LateSetComplex(targetSheet.Rows[num2, RuntimeHelpers.GetObjectValue(Missing.Value)], null, "RowHeight", new object[1] { NewLateBinding.LateGet(sourceSheet.Rows[num4, RuntimeHelpers.GetObjectValue(Missing.Value)], null, "RowHeight", new object[0], null, null, null) }, null, null, OptimisticSet: false, RValueBase: true);
				}
				catch (Exception projectError)
				{
					ProjectData.SetProjectError(projectError);
					ProjectData.ClearProjectError();
				}
				num2++;
			}
			try
			{
				NewLateBinding.LateSetComplex(targetSheet.Columns["A:A", RuntimeHelpers.GetObjectValue(Missing.Value)], null, "ColumnWidth", new object[1] { NewLateBinding.LateGet(sourceSheet.Columns["E:E", RuntimeHelpers.GetObjectValue(Missing.Value)], null, "ColumnWidth", new object[0], null, null, null) }, null, null, OptimisticSet: false, RValueBase: true);
				NewLateBinding.LateSetComplex(targetSheet.Columns["B:B", RuntimeHelpers.GetObjectValue(Missing.Value)], null, "ColumnWidth", new object[1] { NewLateBinding.LateGet(sourceSheet.Columns["AK:AK", RuntimeHelpers.GetObjectValue(Missing.Value)], null, "ColumnWidth", new object[0], null, null, null) }, null, null, OptimisticSet: false, RValueBase: true);
				NewLateBinding.LateSetComplex(targetSheet.Columns["C:C", RuntimeHelpers.GetObjectValue(Missing.Value)], null, "ColumnWidth", new object[1] { NewLateBinding.LateGet(sourceSheet.Columns["AL:AL", RuntimeHelpers.GetObjectValue(Missing.Value)], null, "ColumnWidth", new object[0], null, null, null) }, null, null, OptimisticSet: false, RValueBase: true);
				NewLateBinding.LateSetComplex(targetSheet.Columns["D:D", RuntimeHelpers.GetObjectValue(Missing.Value)], null, "ColumnWidth", new object[1] { NewLateBinding.LateGet(sourceSheet.Columns["AM:AM", RuntimeHelpers.GetObjectValue(Missing.Value)], null, "ColumnWidth", new object[0], null, null, null) }, null, null, OptimisticSet: false, RValueBase: true);
			}
			catch (Exception projectError2)
			{
				ProjectData.SetProjectError(projectError2);
				targetSheet.Columns.AutoFit();
				ProjectData.ClearProjectError();
			}
		}
		catch (Exception ex3)
		{
			ProjectData.SetProjectError(ex3);
			Exception ex4 = ex3;
			Interaction.MsgBox("Error copying visible columns data: " + ex4.Message, MsgBoxStyle.Critical);
			ProjectData.ClearProjectError();
		}
	}
}


using System;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

private Application GetExcelInstance()
{
	try
	{
		return (Application)Marshal.GetActiveObject("Excel.Application");
	}
	catch (Exception ex)
	{
		ProjectData.SetProjectError(ex);
		Exception ex2 = ex;
		Interaction.MsgBox("Excel is not running or could not be accessed: " + ex2.Message, MsgBoxStyle.Critical);
		Application result = null;
		ProjectData.ClearProjectError();
		return result;
	}
}


using System;
using System.Reflection;
using System.Runtime.CompilerServices;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

private void CopyFilteredColumnsData(Worksheet sourceSheet, Worksheet targetSheet, Application excelApp)
{
	checked
	{
		try
		{
			long num = Conversions.ToLong(NewLateBinding.LateGet(NewLateBinding.LateGet(sourceSheet.Cells[sourceSheet.Rows.Count, "E"], null, "End", new object[1] { XlDirection.xlUp }, null, null, null), null, "Row", new object[0], null, null, null));
			long num2 = 0L;
			long num3 = 0L;
			long num4 = num;
			for (long num5 = 9L; num5 <= num4; num5++)
			{
				if (NewLateBinding.LateGet(sourceSheet.Cells[num5, 5], null, "Value", new object[0], null, null, null) != null && NewLateBinding.LateGet(sourceSheet.Cells[num5, 5], null, "Value", new object[0], null, null, null).ToString().Contains("FILTERED DATA"))
				{
					num2 = num5 + 1;
					break;
				}
			}
			if (num2 == 0)
			{
				Interaction.MsgBox("No filtered data found", MsgBoxStyle.Exclamation);
				return;
			}
			long num6 = num2;
			long num7 = num;
			for (long num8 = num6; num8 <= num7; num8++)
			{
				if (NewLateBinding.LateGet(sourceSheet.Cells[num8, 5], null, "Value", new object[0], null, null, null) == null || string.IsNullOrEmpty(NewLateBinding.LateGet(sourceSheet.Cells[num8, 5], null, "Value", new object[0], null, null, null).ToString()))
				{
					num3 = num8 - 1;
					break;
				}
			}
			if (num3 == 0)
			{
				num3 = num;
			}
			long num9 = 1L;
			int[] array = new int[4] { 5, 37, 38, 39 };
			string[] array2 = new string[4] { "A", "B", "C", "D" };
			long num10 = num2;
			long num11 = num3;
			for (long num12 = num10; num12 <= num11; num12++)
			{
				int num13 = array.Length - 1;
				for (int i = 0; i <= num13; i++)
				{
					Range range = (Range)sourceSheet.Cells[num12, array[i]];
					Range range2 = (Range)targetSheet.Cells[num9, i + 1];
					if (!Information.IsNothing(RuntimeHelpers.GetObjectValue(range.get_Value(RuntimeHelpers.GetObjectValue(Missing.Value)))))
					{
						range2.set_Value(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(range.get_Value(RuntimeHelpers.GetObjectValue(Missing.Value))));
					}
					Range range3 = range;
					range2.Font.Name = RuntimeHelpers.GetObjectValue(range3.Font.Name);
					range2.Font.Size = RuntimeHelpers.GetObjectValue(range3.Font.Size);
					range2.Font.Bold = RuntimeHelpers.GetObjectValue(range3.Font.Bold);
					range2.Font.Italic = RuntimeHelpers.GetObjectValue(range3.Font.Italic);
					range2.Font.Color = RuntimeHelpers.GetObjectValue(range3.Font.Color);
					range2.Interior.Color = RuntimeHelpers.GetObjectValue(range3.Interior.Color);
					range2.HorizontalAlignment = RuntimeHelpers.GetObjectValue(range3.HorizontalAlignment);
					range2.VerticalAlignment = RuntimeHelpers.GetObjectValue(range3.VerticalAlignment);
					range2.WrapText = RuntimeHelpers.GetObjectValue(range3.WrapText);
					range2.NumberFormat = RuntimeHelpers.GetObjectValue(range3.NumberFormat);
					range2.Borders[XlBordersIndex.xlEdgeLeft].LineStyle = RuntimeHelpers.GetObjectValue(range3.Borders[XlBordersIndex.xlEdgeLeft].LineStyle);
					range2.Borders[XlBordersIndex.xlEdgeLeft].Weight = RuntimeHelpers.GetObjectValue(range3.Borders[XlBordersIndex.xlEdgeLeft].Weight);
					range2.Borders[XlBordersIndex.xlEdgeRight].LineStyle = RuntimeHelpers.GetObjectValue(range3.Borders[XlBordersIndex.xlEdgeRight].LineStyle);
					range2.Borders[XlBordersIndex.xlEdgeRight].Weight = RuntimeHelpers.GetObjectValue(range3.Borders[XlBordersIndex.xlEdgeRight].Weight);
					range2.Borders[XlBordersIndex.xlEdgeTop].LineStyle = RuntimeHelpers.GetObjectValue(range3.Borders[XlBordersIndex.xlEdgeTop].LineStyle);
					range2.Borders[XlBordersIndex.xlEdgeTop].Weight = RuntimeHelpers.GetObjectValue(range3.Borders[XlBordersIndex.xlEdgeTop].Weight);
					range2.Borders[XlBordersIndex.xlEdgeBottom].LineStyle = RuntimeHelpers.GetObjectValue(range3.Borders[XlBordersIndex.xlEdgeBottom].LineStyle);
					range2.Borders[XlBordersIndex.xlEdgeBottom].Weight = RuntimeHelpers.GetObjectValue(range3.Borders[XlBordersIndex.xlEdgeBottom].Weight);
					range3 = null;
				}
				try
				{
					NewLateBinding.LateSetComplex(targetSheet.Rows[num9, RuntimeHelpers.GetObjectValue(Missing.Value)], null, "RowHeight", new object[1] { NewLateBinding.LateGet(sourceSheet.Rows[num12, RuntimeHelpers.GetObjectValue(Missing.Value)], null, "RowHeight", new object[0], null, null, null) }, null, null, OptimisticSet: false, RValueBase: true);
				}
				catch (Exception projectError)
				{
					ProjectData.SetProjectError(projectError);
					ProjectData.ClearProjectError();
				}
				num9++;
			}
			try
			{
				NewLateBinding.LateSetComplex(targetSheet.Columns["A:A", RuntimeHelpers.GetObjectValue(Missing.Value)], null, "ColumnWidth", new object[1] { NewLateBinding.LateGet(sourceSheet.Columns["E:E", RuntimeHelpers.GetObjectValue(Missing.Value)], null, "ColumnWidth", new object[0], null, null, null) }, null, null, OptimisticSet: false, RValueBase: true);
				NewLateBinding.LateSetComplex(targetSheet.Columns["B:B", RuntimeHelpers.GetObjectValue(Missing.Value)], null, "ColumnWidth", new object[1] { NewLateBinding.LateGet(sourceSheet.Columns["AK:AK", RuntimeHelpers.GetObjectValue(Missing.Value)], null, "ColumnWidth", new object[0], null, null, null) }, null, null, OptimisticSet: false, RValueBase: true);
				NewLateBinding.LateSetComplex(targetSheet.Columns["C:C", RuntimeHelpers.GetObjectValue(Missing.Value)], null, "ColumnWidth", new object[1] { NewLateBinding.LateGet(sourceSheet.Columns["AL:AL", RuntimeHelpers.GetObjectValue(Missing.Value)], null, "ColumnWidth", new object[0], null, null, null) }, null, null, OptimisticSet: false, RValueBase: true);
				NewLateBinding.LateSetComplex(targetSheet.Columns["D:D", RuntimeHelpers.GetObjectValue(Missing.Value)], null, "ColumnWidth", new object[1] { NewLateBinding.LateGet(sourceSheet.Columns["AM:AM", RuntimeHelpers.GetObjectValue(Missing.Value)], null, "ColumnWidth", new object[0], null, null, null) }, null, null, OptimisticSet: false, RValueBase: true);
			}
			catch (Exception projectError2)
			{
				ProjectData.SetProjectError(projectError2);
				targetSheet.Columns.AutoFit();
				ProjectData.ClearProjectError();
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			Interaction.MsgBox("Error copying filtered columns data: " + ex2.Message, MsgBoxStyle.Critical);
			ProjectData.ClearProjectError();
		}
	}
}


private int CompareAlphanumeric(string x, string y)
{
	string letter = "";
	int number = 0;
	string letter2 = "";
	int number2 = 0;
	ExtractLetterNumber(x, ref letter, ref number);
	ExtractLetterNumber(y, ref letter2, ref number2);
	int num = string.Compare(letter, letter2);
	if (num != 0)
	{
		return num;
	}
	return number.CompareTo(number2);
}


private void ExtractLetterNumber(string input, ref string letter, ref int number)
{
	if (string.IsNullOrEmpty(input))
	{
		letter = "";
		number = 0;
		return;
	}
	int i;
	for (i = 0; i < input.Length && !char.IsDigit(input[i]); i = checked(i + 1))
	{
	}
	letter = ((i > 0) ? input.Substring(0, i) : "");
	if (i < input.Length)
	{
		string s = input.Substring(i);
		int.TryParse(s, out number);
	}
	else
	{
		number = 0;
	}
}


using System;
using System.Reflection;
using System.Runtime.CompilerServices;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

private void UnlockCopyColumns(Worksheet sheet, long lastFilteredRow)
{
	checked
	{
		try
		{
			long num = Conversions.ToLong(NewLateBinding.LateGet(NewLateBinding.LateGet(sheet.Cells[sheet.Rows.Count, "E"], null, "End", new object[1] { XlDirection.xlUp }, null, null, null), null, "Row", new object[0], null, null, null));
			long num2 = 0L;
			long num3 = num;
			for (long num4 = 9L; num4 <= num3; num4++)
			{
				if (NewLateBinding.LateGet(sheet.Cells[num4, 5], null, "Value", new object[0], null, null, null) != null && NewLateBinding.LateGet(sheet.Cells[num4, 5], null, "Value", new object[0], null, null, null).ToString().Contains("FILTERED DATA"))
				{
					num2 = num4 + 1;
					break;
				}
			}
			if (num2 != 0)
			{
				Range range = ((_Worksheet)sheet).get_Range((object)$"E{num2}:E{lastFilteredRow}", RuntimeHelpers.GetObjectValue(Missing.Value));
				range.Locked = false;
				Range range2 = ((_Worksheet)sheet).get_Range((object)$"AK{num2}:AK{lastFilteredRow}", RuntimeHelpers.GetObjectValue(Missing.Value));
				Range range3 = ((_Worksheet)sheet).get_Range((object)$"AL{num2}:AL{lastFilteredRow}", RuntimeHelpers.GetObjectValue(Missing.Value));
				Range range4 = ((_Worksheet)sheet).get_Range((object)$"AM{num2}:AM{lastFilteredRow}", RuntimeHelpers.GetObjectValue(Missing.Value));
				range2.Locked = false;
				range3.Locked = false;
				range4.Locked = false;
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			Interaction.MsgBox("Error unlocking copy columns: " + ex2.Message);
			ProjectData.ClearProjectError();
		}
	}
}


using System;
using System.Reflection;
using System.Runtime.CompilerServices;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

private void LockCopyColumns(Worksheet sheet)
{
	try
	{
		long num = Conversions.ToLong(NewLateBinding.LateGet(NewLateBinding.LateGet(sheet.Cells[sheet.Rows.Count, "E"], null, "End", new object[1] { XlDirection.xlUp }, null, null, null), null, "Row", new object[0], null, null, null));
		Range range = ((_Worksheet)sheet).get_Range((object)$"E9:E{num}", RuntimeHelpers.GetObjectValue(Missing.Value));
		range.Locked = true;
		Range range2 = ((_Worksheet)sheet).get_Range((object)$"AK9:AK{num}", RuntimeHelpers.GetObjectValue(Missing.Value));
		Range range3 = ((_Worksheet)sheet).get_Range((object)$"AL9:AL{num}", RuntimeHelpers.GetObjectValue(Missing.Value));
		Range range4 = ((_Worksheet)sheet).get_Range((object)$"AM9:AM{num}", RuntimeHelpers.GetObjectValue(Missing.Value));
		range2.Locked = true;
		range3.Locked = true;
		range4.Locked = true;
	}
	catch (Exception ex)
	{
		ProjectData.SetProjectError(ex);
		Exception ex2 = ex;
		Interaction.MsgBox("Error locking copy columns: " + ex2.Message);
		ProjectData.ClearProjectError();
	}
}


using System;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

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


using System;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

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


using System;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

private string GetCellTextSafely(Range rng)
{
	if (rng == null)
	{
		return string.Empty;
	}
	try
	{
		return rng.Text.ToString().Trim();
	}
	catch (Exception projectError)
	{
		ProjectData.SetProjectError(projectError);
		string empty = string.Empty;
		ProjectData.ClearProjectError();
		return empty;
	}
	finally
	{
		if (rng != null)
		{
			Marshal.ReleaseComObject(rng);
		}
	}
}


using System.Collections.Generic;
using System.Runtime.CompilerServices;

[SpecialName]
[CompilerGenerated]
private int _Lambda$__3-0(KeyValuePair<string, long> x, KeyValuePair<string, long> y)
{
	return CompareAlphanumeric(x.Key, y.Key);
}
