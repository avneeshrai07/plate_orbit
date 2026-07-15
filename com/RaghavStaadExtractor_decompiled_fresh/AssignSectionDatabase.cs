// ===== RaghavStaadExtractor.AssignSectionDatabase =====

// Decompiled member-by-member from RaghavStaadExtractor.dll



public AssignSectionDatabase()
{
}


using System;
using System.IO;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

public void CopySectionDatabaseToSheet3()
{
	Application application = null;
	Workbook workbook = null;
	Worksheet worksheet = null;
	Workbook workbook2 = null;
	Worksheet worksheet2 = null;
	try
	{
		application = (Application)Marshal.GetActiveObject("Excel.Application");
		application.ScreenUpdating = false;
		workbook2 = application.ActiveWorkbook;
		worksheet2 = (Worksheet)workbook2.Sheets[3];
		worksheet2.Cells.Clear();
		string path = workbook2.Path + "\\";
		string text = "RAGHAV DATABASE.xlsm";
		string text2 = Path.Combine(path, text);
		if (!File.Exists(text2))
		{
			Interaction.MsgBox("File '" + text + "' not found in the current folder.", MsgBoxStyle.Exclamation);
			return;
		}
		workbook = application.Workbooks.Open(text2, RuntimeHelpers.GetObjectValue(Missing.Value), true, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
		worksheet = (Worksheet)workbook.Sheets[1];
		long num = Conversions.ToLong(NewLateBinding.LateGet(NewLateBinding.LateGet(worksheet.Cells[worksheet.Rows.Count, 1], null, "End", new object[1] { XlDirection.xlUp }, null, null, null), null, "Row", new object[0], null, null, null));
		long num2 = Conversions.ToLong(NewLateBinding.LateGet(NewLateBinding.LateGet(worksheet.Cells[5, worksheet.Columns.Count], null, "End", new object[1] { XlDirection.xlToLeft }, null, null, null), null, "Column", new object[0], null, null, null));
		Range range = ((_Worksheet)worksheet).get_Range(RuntimeHelpers.GetObjectValue(worksheet.Cells[5, 1]), RuntimeHelpers.GetObjectValue(worksheet.Cells[num, num2]));
		Range range2 = ((_Worksheet)worksheet2).get_Range((object)"A1", RuntimeHelpers.GetObjectValue(Missing.Value)).get_Resize((object)range.Rows.Count, (object)range.Columns.Count);
		range2.set_Value(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(range.get_Value(RuntimeHelpers.GetObjectValue(Missing.Value))));
		range2.Font.Color = Information.RGB(0, 0, 139);
		workbook.Close(false, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
		Interaction.MsgBox("Staad Section Database, Perfectly Aligned!", MsgBoxStyle.Information);
	}
	catch (Exception ex)
	{
		ProjectData.SetProjectError(ex);
		Exception ex2 = ex;
		Interaction.MsgBox("Error: " + ex2.Message, MsgBoxStyle.Critical);
		ProjectData.ClearProjectError();
	}
	finally
	{
		if (workbook != null)
		{
			Marshal.ReleaseComObject(workbook);
		}
		if (worksheet != null)
		{
			Marshal.ReleaseComObject(worksheet);
		}
		if (worksheet2 != null)
		{
			Marshal.ReleaseComObject(worksheet2);
		}
		if (workbook2 != null)
		{
			Marshal.ReleaseComObject(workbook2);
		}
		if (application != null)
		{
			application.ScreenUpdating = true;
		}
		worksheet = null;
		workbook = null;
		worksheet2 = null;
		workbook2 = null;
		application = null;
		GC.Collect();
		GC.WaitForPendingFinalizers();
	}
}
