using System;
using System.Drawing;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace RaghavStaadExtractor;

public class ClearSheet
{
	private const string ProtectionPassword = "2022";

	public void ClearAndResetSheet1()
	{
		ClearSpecificSheet(1);
	}

	public void ClearAndResetSheet2()
	{
		ClearSpecificSheet(2);
	}

	public void ClearAllSheets()
	{
		Application application = (Application)Marshal.GetActiveObject("Excel.Application");
		Workbook activeWorkbook = application.ActiveWorkbook;
		if (Interaction.MsgBox("Do you want to proceed with clearing all sheets?", MsgBoxStyle.OkCancel | MsgBoxStyle.Question, "Confirm Action") != MsgBoxResult.Ok)
		{
			Interaction.MsgBox("Action canceled by the user.", MsgBoxStyle.Information, "Canceled");
			return;
		}
		int num = Math.Min(2, activeWorkbook.Sheets.Count);
		for (int i = 1; i <= num; i = checked(i + 1))
		{
			Worksheet worksheet = (Worksheet)activeWorkbook.Sheets[i];
			worksheet.Unprotect("2022");
			switch (i)
			{
			case 1:
				ClearSheet1(worksheet);
				break;
			case 2:
				ClearSheet2(worksheet);
				break;
			}
			worksheet.Protect("2022", RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
		}
	}

	private void ClearSpecificSheet(int sheetIndex)
	{
		Application application = (Application)Marshal.GetActiveObject("Excel.Application");
		Workbook activeWorkbook = application.ActiveWorkbook;
		if (sheetIndex < 1 || sheetIndex > activeWorkbook.Sheets.Count)
		{
			Interaction.MsgBox("Invalid sheet index.", MsgBoxStyle.Critical, "Error");
			return;
		}
		Worksheet worksheet = (Worksheet)activeWorkbook.Sheets[sheetIndex];
		worksheet.Unprotect("2022");
		string prompt = ((sheetIndex == 1) ? "Do you want to proceed with clearing Master Data List?" : "Do you want to proceed with clearing Summary List?");
		if (Interaction.MsgBox(prompt, MsgBoxStyle.OkCancel | MsgBoxStyle.Question, "Confirm Action") != MsgBoxResult.Ok)
		{
			Interaction.MsgBox("Action canceled by the user.", MsgBoxStyle.Information, "Canceled");
			return;
		}
		switch (sheetIndex)
		{
		case 1:
			ClearSheet1(worksheet);
			break;
		case 2:
			ClearSheet2(worksheet);
			break;
		}
		worksheet.Protect("2022", RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
	}

	private void ClearSheet1(Worksheet ws)
	{
		object objectValue = RuntimeHelpers.GetObjectValue(((_Worksheet)ws).get_Range((object)"AS2", RuntimeHelpers.GetObjectValue(Missing.Value)).get_Value(RuntimeHelpers.GetObjectValue(Missing.Value)));
		object objectValue2 = RuntimeHelpers.GetObjectValue(((_Worksheet)ws).get_Range((object)"AS3", RuntimeHelpers.GetObjectValue(Missing.Value)).get_Value(RuntimeHelpers.GetObjectValue(Missing.Value)));
		object objectValue3 = RuntimeHelpers.GetObjectValue(((_Worksheet)ws).get_Range((object)"AS4", RuntimeHelpers.GetObjectValue(Missing.Value)).get_Value(RuntimeHelpers.GetObjectValue(Missing.Value)));
		object objectValue4 = RuntimeHelpers.GetObjectValue(((_Worksheet)ws).get_Range((object)"AS5", RuntimeHelpers.GetObjectValue(Missing.Value)).get_Value(RuntimeHelpers.GetObjectValue(Missing.Value)));
		object objectValue5 = RuntimeHelpers.GetObjectValue(((_Worksheet)ws).get_Range((object)"AS6", RuntimeHelpers.GetObjectValue(Missing.Value)).get_Value(RuntimeHelpers.GetObjectValue(Missing.Value)));
		object objectValue6 = RuntimeHelpers.GetObjectValue(((_Worksheet)ws).get_Range((object)"AR7", RuntimeHelpers.GetObjectValue(Missing.Value)).get_Value(RuntimeHelpers.GetObjectValue(Missing.Value)));
		object objectValue7 = RuntimeHelpers.GetObjectValue(((_Worksheet)ws).get_Range((object)"AS7", RuntimeHelpers.GetObjectValue(Missing.Value)).get_Value(RuntimeHelpers.GetObjectValue(Missing.Value)));
		ws.Cells.Clear();
		ws.Cells.Interior.Color = Color.White.ToArgb();
		((_Worksheet)ws).get_Range((object)"AS2", RuntimeHelpers.GetObjectValue(Missing.Value)).set_Value(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(objectValue));
		((_Worksheet)ws).get_Range((object)"AS3", RuntimeHelpers.GetObjectValue(Missing.Value)).set_Value(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(objectValue2));
		((_Worksheet)ws).get_Range((object)"AS4", RuntimeHelpers.GetObjectValue(Missing.Value)).set_Value(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(objectValue3));
		((_Worksheet)ws).get_Range((object)"AS5", RuntimeHelpers.GetObjectValue(Missing.Value)).set_Value(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(objectValue4));
		((_Worksheet)ws).get_Range((object)"AS6", RuntimeHelpers.GetObjectValue(Missing.Value)).set_Value(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(objectValue5));
		((_Worksheet)ws).get_Range((object)"AR7", RuntimeHelpers.GetObjectValue(Missing.Value)).set_Value(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(objectValue6));
		((_Worksheet)ws).get_Range((object)"AS7", RuntimeHelpers.GetObjectValue(Missing.Value)).set_Value(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(objectValue7));
		Range range = ((_Worksheet)ws).get_Range((object)"C2", (object)"AM5");
		range.Merge(RuntimeHelpers.GetObjectValue(Missing.Value));
		range.set_Value(RuntimeHelpers.GetObjectValue(Missing.Value), (object)"A Vision Brought to Life by Raghav\r\nPRESS BUTTONS FOR [STAAD OPERATIONS]");
		range.HorizontalAlignment = XlHAlign.xlHAlignCenter;
		range.VerticalAlignment = XlVAlign.xlVAlignCenter;
		range.Font.Size = 14;
		range.Font.Bold = true;
		range.Font.Color = Information.RGB(0, 0, 0);
		range.Interior.Color = Information.RGB(0, 176, 240);
		range.Borders.LineStyle = XlLineStyle.xlContinuous;
		range.Borders.Weight = XlBorderWeight.xlThin;
		range = null;
		string[] array = new string[17]
		{
			"A1: Extract structural data directly from STAAD model.", "A2: Generate a comprehensive Material Summary List.", "SAVEAS: Save current output specifically for Beam details.", "MPS: Redirect to the MPS (Material Processing Sheet).", "CLEAR: Clears the active worksheet only.", "CLEAR ALL: Clears all worksheets in the workbook.", "DATABASE: Assign Raghav Database to Plate Nova (one-time setup only).", "DRAW: Activate Drawing Mode for custom visual outputs.", "RESET: Revert to default interface and settings.", "CREATED BY: Edit project or job-related metadata (button-enabled).",
			"FRONT: Generate front elevation view (ideal for frames and full 3D models).", "LEFT: Generate left side elevation (recommended for side views).", "RIGHT: Generate right side elevation (recommended for opposite side views).", "TOP: Generate top view (ideal for roof and anchor bolt plans).", "PLACEMENT: Enter U for upward text, D for downward, N for no text.", "TEXT HEIGHT: Specify desired text height.", "TEXT COLOR: Specify text color (0 = multi-color, 100=Enable group mode,101-Enble property+group mode)."
		};
		int num = 9;
		bool flag = true;
		string[] array2 = array;
		foreach (string text in array2)
		{
			Range range2 = (Range)ws.Cells[num, 1];
			range2.set_Value(RuntimeHelpers.GetObjectValue(Missing.Value), (object)text);
			if (flag)
			{
				range2.Font.Color = Information.RGB(0, 0, 0);
			}
			else
			{
				range2.Font.Color = Information.RGB(0, 0, 146);
			}
			flag = !flag;
			num = checked(num + 1);
		}
	}

	private void ClearSheet2(Worksheet ws)
	{
		ws.Cells.Clear();
		ws.Cells.Interior.Color = Information.RGB(255, 255, 255);
		Range range = ((_Worksheet)ws).get_Range((object)"A2:F4", RuntimeHelpers.GetObjectValue(Missing.Value));
		range.Merge(RuntimeHelpers.GetObjectValue(Missing.Value));
		range.set_Value(RuntimeHelpers.GetObjectValue(Missing.Value), (object)"A Vision Brought to Life by Raghav\r\nPRESS BUTTONS FOR [MATERIAL SUMMARY]");
		range.HorizontalAlignment = XlHAlign.xlHAlignCenter;
		range.VerticalAlignment = XlVAlign.xlVAlignCenter;
		range.Font.Size = 14;
		range.Font.Bold = true;
		range.Font.Color = Information.RGB(0, 0, 0);
		range.Interior.Color = Information.RGB(0, 176, 240);
		range.Borders.LineStyle = XlLineStyle.xlContinuous;
		range.Borders.Weight = XlBorderWeight.xlThin;
		range = null;
		string[] array = new string[17]
		{
			"A1: Extract structural data directly from STAAD model.", "A2: Generate a comprehensive Material Summary List.", "SAVEAS: Save current output specifically for Beam details.", "MPS: Redirect to the MPS (Material Processing Sheet).", "CLEAR: Clears the active worksheet only.", "CLEAR ALL: Clears all worksheets in the workbook.", "DATABASE: Assign Raghav Database to Plate Nova (one-time setup only).", "DRAW: Activate Drawing Mode for custom visual outputs.", "RESET: Revert to default interface and settings.", "CREATED BY: Edit project or job-related metadata (button-enabled).",
			"FRONT: Generate front elevation view (ideal for frames and full 3D models).", "LEFT: Generate left side elevation (recommended for side views).", "RIGHT: Generate right side elevation (recommended for opposite side views).", "TOP: Generate top view (ideal for roof and anchor bolt plans).", "PLACEMENT: Enter U for upward text, D for downward, N for no text.", "TEXT HEIGHT: Specify desired text height.", "TEXT COLOR: Specify text color (0 = multi-color, recommended)."
		};
		int num = 9;
		string[] array2 = array;
		foreach (string text in array2)
		{
			NewLateBinding.LateSetComplex(ws.Cells[num, "A"], null, "Value", new object[1] { text }, null, null, OptimisticSet: false, RValueBase: true);
			num = checked(num + 1);
		}
	}
}
