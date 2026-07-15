// ===== RaghavStaadExtractor.RaghavStaadExtractor.RaghavDatabase =====

// Decompiled member-by-member from RaghavStaadExtractor.dll



private const string pwd = "2022";


public RaghavDatabase()
{
}


using System;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

[ComVisible(true)]
public void DataRewiewUnlocked()
{
	Application application = null;
	try
	{
		application = (Application)Marshal.GetActiveObject("Excel.Application");
		Worksheet worksheet = (Worksheet)application.ActiveSheet;
		Workbook activeWorkbook = application.ActiveWorkbook;
		application.ScreenUpdating = false;
		application.EnableEvents = false;
		application.Calculation = XlCalculation.xlCalculationManual;
		activeWorkbook.Unprotect("2022");
		int count = activeWorkbook.Sheets.Count;
		for (int i = 2; i <= count; i = checked(i + 1))
		{
			Worksheet worksheet2 = (Worksheet)activeWorkbook.Sheets[i];
			worksheet2.Unprotect("2022");
			worksheet2.Visible = XlSheetVisibility.xlSheetVisible;
			worksheet2.Cells.Locked = true;
			worksheet2.Protect("2022", RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), false, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), false, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), false, false, false, RuntimeHelpers.GetObjectValue(Missing.Value));
		}
		activeWorkbook.Protect("2022", true, RuntimeHelpers.GetObjectValue(Missing.Value));
		worksheet.Activate();
		application.ScreenUpdating = true;
		Interaction.MsgBox("Step Inside: The Data Room is Now Open.", MsgBoxStyle.Information);
	}
	catch (Exception ex)
	{
		ProjectData.SetProjectError(ex);
		Exception ex2 = ex;
		Interaction.MsgBox("Error in DataRewiewUnlocked: " + ex2.Message, MsgBoxStyle.Critical);
		ProjectData.ClearProjectError();
	}
	finally
	{
		if (application != null)
		{
			application.ScreenUpdating = true;
			application.EnableEvents = true;
			application.Calculation = XlCalculation.xlCalculationAutomatic;
		}
	}
}


using System;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

[ComVisible(true)]
public void LetsGenrateUnitWt()
{
	Application application = null;
	checked
	{
		try
		{
			application = (Application)Marshal.GetActiveObject("Excel.Application");
			Workbook activeWorkbook = application.ActiveWorkbook;
			Worksheet worksheet = (Worksheet)activeWorkbook.Sheets["Sheet1"];
			long num = 6L;
			long num2 = 1L;
			application.ScreenUpdating = false;
			application.EnableEvents = false;
			application.Calculation = XlCalculation.xlCalculationManual;
			activeWorkbook.Unprotect("2022");
			int count = activeWorkbook.Sheets.Count;
			for (int i = 1; i <= count; i++)
			{
				Worksheet worksheet2 = (Worksheet)activeWorkbook.Sheets[i];
				worksheet2.Visible = XlSheetVisibility.xlSheetVisible;
				worksheet2.Unprotect("2022");
			}
			worksheet.Cells.Clear();
			Range range = ((_Worksheet)worksheet).get_Range((object)"A1:E1", RuntimeHelpers.GetObjectValue(Missing.Value));
			range.Merge(RuntimeHelpers.GetObjectValue(Missing.Value));
			range.set_Value(RuntimeHelpers.GetObjectValue(Missing.Value), (object)"RAGHAV DATABASE");
			range.HorizontalAlignment = XlHAlign.xlHAlignCenter;
			range.VerticalAlignment = XlVAlign.xlVAlignCenter;
			range.Font.Bold = true;
			range.Font.Size = 14;
			range.Interior.Color = Information.RGB(247, 199, 172);
			range = null;
			((_Worksheet)worksheet).get_Range((object)"A2:E3", RuntimeHelpers.GetObjectValue(Missing.Value)).Merge(RuntimeHelpers.GetObjectValue(Missing.Value));
			((_Worksheet)worksheet).get_Range((object)"A2:E3", RuntimeHelpers.GetObjectValue(Missing.Value)).set_Value(RuntimeHelpers.GetObjectValue(Missing.Value), (object)"");
			((_Worksheet)worksheet).get_Range((object)"A4:E4", RuntimeHelpers.GetObjectValue(Missing.Value)).Merge(RuntimeHelpers.GetObjectValue(Missing.Value));
			((_Worksheet)worksheet).get_Range((object)"A4:E4", RuntimeHelpers.GetObjectValue(Missing.Value)).Interior.Color = Information.RGB(247, 199, 172);
			Range range2 = ((_Worksheet)worksheet).get_Range((object)"A5:E5", RuntimeHelpers.GetObjectValue(Missing.Value));
			range2.set_Value(RuntimeHelpers.GetObjectValue(Missing.Value), (object)new string[5] { "S.No", "Name", "Staad Name", "AREA Ax(cm2)", "Kg/M" });
			range2.Font.Bold = true;
			range2.Font.Color = Information.RGB(0, 0, 0);
			range2.Interior.Color = Information.RGB(173, 216, 230);
			range2 = null;
			((_Worksheet)worksheet).get_Range((object)"A1:E4", RuntimeHelpers.GetObjectValue(Missing.Value)).Borders.LineStyle = XlLineStyle.xlContinuous;
			int count2 = activeWorkbook.Sheets.Count;
			for (int j = 2; j <= count2; j++)
			{
				Worksheet worksheet3 = (Worksheet)activeWorkbook.Sheets[j];
				long num3 = Conversions.ToLong(NewLateBinding.LateGet(NewLateBinding.LateGet(worksheet3.Cells[worksheet3.Rows.Count, "B"], null, "End", new object[1] { XlDirection.xlUp }, null, null, null), null, "Row", new object[0], null, null, null));
				if (num3 < 3)
				{
					continue;
				}
				Range range3 = ((_Worksheet)worksheet3).get_Range((object)("B3:D" + Conversions.ToString(num3)), RuntimeHelpers.GetObjectValue(Missing.Value));
				NewLateBinding.LateSetComplex(NewLateBinding.LateGet(worksheet.Cells[num, "B"], null, "Resize", new object[2]
				{
					range3.Rows.Count,
					3
				}, null, null, null), null, "Value", new object[1] { range3.get_Value(RuntimeHelpers.GetObjectValue(Missing.Value)) }, null, null, OptimisticSet: false, RValueBase: true);
				long num4 = num;
				long num5 = num + range3.Rows.Count - 1;
				for (long num6 = num4; num6 <= num5; num6++)
				{
					NewLateBinding.LateSetComplex(worksheet.Cells[num6, "A"], null, "Value", new object[1] { num2 }, null, null, OptimisticSet: false, RValueBase: true);
					if (Versioned.IsNumeric(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(worksheet.Cells[num6, "D"], null, "Value", new object[0], null, null, null))))
					{
						NewLateBinding.LateSetComplex(worksheet.Cells[num6, "E"], null, "Value", new object[1] { Math.Round(Conversions.ToDouble(NewLateBinding.LateGet(worksheet.Cells[num6, "D"], null, "Value", new object[0], null, null, null)) * 0.785, 3) }, null, null, OptimisticSet: false, RValueBase: true);
					}
					else
					{
						NewLateBinding.LateSetComplex(worksheet.Cells[num6, "E"], null, "Value", new object[1] { "N/A" }, null, null, OptimisticSet: false, RValueBase: true);
					}
					num2++;
				}
				num = Conversions.ToLong(Operators.AddObject(NewLateBinding.LateGet(NewLateBinding.LateGet(worksheet.Cells[worksheet.Rows.Count, "B"], null, "End", new object[1] { XlDirection.xlUp }, null, null, null), null, "Row", new object[0], null, null, null), 1));
			}
			long num7 = Conversions.ToLong(NewLateBinding.LateGet(NewLateBinding.LateGet(worksheet.Cells[worksheet.Rows.Count, "A"], null, "End", new object[1] { XlDirection.xlUp }, null, null, null), null, "Row", new object[0], null, null, null));
			((_Worksheet)worksheet).get_Range((object)("A6:E" + Conversions.ToString(num7)), RuntimeHelpers.GetObjectValue(Missing.Value)).Borders.LineStyle = XlLineStyle.xlContinuous;
			NewLateBinding.LateCall(worksheet.Columns["A:E", RuntimeHelpers.GetObjectValue(Missing.Value)], null, "AutoFit", new object[0], null, null, null, IgnoreReturn: true);
			int count3 = activeWorkbook.Sheets.Count;
			for (int k = 2; k <= count3; k++)
			{
				Worksheet worksheet4 = (Worksheet)activeWorkbook.Sheets[k];
				worksheet4.Visible = XlSheetVisibility.xlSheetVeryHidden;
				worksheet4.Protect("2022", RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), false, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), false, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), false, false, false, RuntimeHelpers.GetObjectValue(Missing.Value));
			}
			worksheet.Unprotect("2022");
			worksheet.Cells.Locked = false;
			activeWorkbook.Protect("2022", true, RuntimeHelpers.GetObjectValue(Missing.Value));
			application.ScreenUpdating = true;
			Interaction.MsgBox("All set! PlateNova can now use the Raghav Database.", MsgBoxStyle.Information);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			Interaction.MsgBox("Error in LetsGenrateUnitWt: " + ex2.Message, MsgBoxStyle.Critical);
			ProjectData.ClearProjectError();
		}
		finally
		{
			if (application != null)
			{
				application.ScreenUpdating = true;
				application.EnableEvents = true;
				application.Calculation = XlCalculation.xlCalculationAutomatic;
			}
		}
	}
}


using System;
using System.Collections;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

[ComVisible(true)]
public void UnprotectSheetsAllowAdd()
{
	Application application = null;
	try
	{
		application = (Application)Marshal.GetActiveObject("Excel.Application");
		Workbook activeWorkbook = application.ActiveWorkbook;
		Worksheet worksheet = (Worksheet)application.ActiveSheet;
		application.ScreenUpdating = false;
		application.EnableEvents = false;
		application.Calculation = XlCalculation.xlCalculationManual;
		activeWorkbook.Unprotect("2022");
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = activeWorkbook.Worksheets.GetEnumerator();
			while (enumerator.MoveNext())
			{
				Worksheet worksheet2 = (Worksheet)enumerator.Current;
				worksheet2.Visible = XlSheetVisibility.xlSheetVisible;
				worksheet2.Unprotect("2022");
				worksheet2.Protect("2022", RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), true, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), true, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), true, true, true, RuntimeHelpers.GetObjectValue(Missing.Value));
			}
		}
		finally
		{
			if (enumerator is IDisposable)
			{
				(enumerator as IDisposable).Dispose();
			}
		}
		worksheet.Activate();
		application.ScreenUpdating = true;
		Interaction.MsgBox("Blank pages await your brilliance — go ahead, add those sheets and let ideas flow!", MsgBoxStyle.Information);
	}
	catch (Exception ex)
	{
		ProjectData.SetProjectError(ex);
		Exception ex2 = ex;
		Interaction.MsgBox("Error in UnprotectSheetsAllowAdd: " + ex2.Message, MsgBoxStyle.Critical);
		ProjectData.ClearProjectError();
	}
	finally
	{
		if (application != null)
		{
			application.ScreenUpdating = true;
			application.EnableEvents = true;
			application.Calculation = XlCalculation.xlCalculationAutomatic;
		}
	}
}
