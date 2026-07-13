using System;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace RaghavStaadExtractor;

[ComVisible(true)]
[Guid("B1234567-89CD-4ABC-8123-DEF456789ABC")]
[ProgId("RaghavStaadExtractor.ColumnHider")]
public class ColumnHider
{
	private const string PASSWORD = "2022";

	private const string SHEET_NAME = "Sheet1";

	private void ApplyVisibility(Dictionary<string, bool> columnStates)
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
			foreach (KeyValuePair<string, bool> columnState in columnStates)
			{
				((_Worksheet)worksheet).get_Range((object)columnState.Key, RuntimeHelpers.GetObjectValue(Missing.Value)).EntireColumn.Hidden = columnState.Value;
			}
			((_Worksheet)worksheet).get_Range((object)"F:AJ", RuntimeHelpers.GetObjectValue(Missing.Value)).EntireColumn.Hidden = true;
			((_Worksheet)worksheet).get_Range((object)"AY:BC", RuntimeHelpers.GetObjectValue(Missing.Value)).EntireColumn.Hidden = true;
			worksheet.Protect("2022", true, true, true, true, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			Interaction.MsgBox("Error modifying column visibility: " + ex2.Message);
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

	public void ShowAAtoAQ_HideARtoAW()
	{
		Dictionary<string, bool> columnStates = new Dictionary<string, bool>
		{
			{ "A:AQ", false },
			{ "AR:AW", true }
		};
		ApplyVisibility(columnStates);
	}

	public void HideAAtoAQ_ShowARtoAW()
	{
		Dictionary<string, bool> columnStates = new Dictionary<string, bool>
		{
			{ "A:AQ", true },
			{ "AR:AW", false }
		};
		ApplyVisibility(columnStates);
	}

	public void NormalViewExceptFAJ()
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
			worksheet.Cells.EntireColumn.Hidden = false;
			((_Worksheet)worksheet).get_Range((object)"F:AJ", RuntimeHelpers.GetObjectValue(Missing.Value)).EntireColumn.Hidden = true;
			((_Worksheet)worksheet).get_Range((object)"AY:BC", RuntimeHelpers.GetObjectValue(Missing.Value)).EntireColumn.Hidden = true;
			worksheet.Protect("2022", true, true, true, true, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			Interaction.MsgBox("Error restoring normal view: " + ex2.Message);
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
}
