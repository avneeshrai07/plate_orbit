// ===== RaghavStaadExtractor.GridSystem =====

// Decompiled member-by-member from RaghavStaadExtractor.dll



public GridSystem()
{
}


using System;
using System.Windows.Forms;
using Microsoft.VisualBasic.CompilerServices;

public void ShowGridInputForm()
{
	try
	{
		string[] initialValues = ReadValuesFromExcel();
		using GridInputForm gridInputForm = new GridInputForm(initialValues);
		if (gridInputForm.ShowDialog() == DialogResult.OK)
		{
			WriteValuesToExcel(gridInputForm.Values);
		}
	}
	catch (Exception ex)
	{
		ProjectData.SetProjectError(ex);
		Exception ex2 = ex;
		MessageBox.Show("Error: " + ex2.Message);
		ProjectData.ClearProjectError();
	}
}


using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;

private string[] ReadValuesFromExcel()
{
	Application application = (Application)Marshal.GetActiveObject("Excel.Application");
	Workbook activeWorkbook = application.ActiveWorkbook;
	Worksheet worksheet = (Worksheet)activeWorkbook.Sheets["Sheet4"];
	return new string[10]
	{
		IfNull(RuntimeHelpers.GetObjectValue(((_Worksheet)worksheet).get_Range((object)"B2", RuntimeHelpers.GetObjectValue(Missing.Value)).get_Value(RuntimeHelpers.GetObjectValue(Missing.Value)))),
		IfNull(RuntimeHelpers.GetObjectValue(((_Worksheet)worksheet).get_Range((object)"B3", RuntimeHelpers.GetObjectValue(Missing.Value)).get_Value(RuntimeHelpers.GetObjectValue(Missing.Value)))),
		IfNull(RuntimeHelpers.GetObjectValue(((_Worksheet)worksheet).get_Range((object)"B4", RuntimeHelpers.GetObjectValue(Missing.Value)).get_Value(RuntimeHelpers.GetObjectValue(Missing.Value)))),
		IfNull(RuntimeHelpers.GetObjectValue(((_Worksheet)worksheet).get_Range((object)"C2", RuntimeHelpers.GetObjectValue(Missing.Value)).get_Value(RuntimeHelpers.GetObjectValue(Missing.Value)))),
		IfNull(RuntimeHelpers.GetObjectValue(((_Worksheet)worksheet).get_Range((object)"C3", RuntimeHelpers.GetObjectValue(Missing.Value)).get_Value(RuntimeHelpers.GetObjectValue(Missing.Value)))),
		IfNull(RuntimeHelpers.GetObjectValue(((_Worksheet)worksheet).get_Range((object)"C4", RuntimeHelpers.GetObjectValue(Missing.Value)).get_Value(RuntimeHelpers.GetObjectValue(Missing.Value)))),
		IfNull(RuntimeHelpers.GetObjectValue(((_Worksheet)worksheet).get_Range((object)"D2", RuntimeHelpers.GetObjectValue(Missing.Value)).get_Value(RuntimeHelpers.GetObjectValue(Missing.Value)))),
		IfNull(RuntimeHelpers.GetObjectValue(((_Worksheet)worksheet).get_Range((object)"D3", RuntimeHelpers.GetObjectValue(Missing.Value)).get_Value(RuntimeHelpers.GetObjectValue(Missing.Value)))),
		IfNull(RuntimeHelpers.GetObjectValue(((_Worksheet)worksheet).get_Range((object)"D4", RuntimeHelpers.GetObjectValue(Missing.Value)).get_Value(RuntimeHelpers.GetObjectValue(Missing.Value)))),
		IfNull(RuntimeHelpers.GetObjectValue(((_Worksheet)worksheet).get_Range((object)"E2", RuntimeHelpers.GetObjectValue(Missing.Value)).get_Value(RuntimeHelpers.GetObjectValue(Missing.Value))))
	};
}


private string IfNull(object val)
{
	return (val == null) ? "" : val.ToString();
}


using System;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

private void WriteValuesToExcel(string[] values)
{
	try
	{
		Microsoft.Office.Interop.Excel.Application application = (Microsoft.Office.Interop.Excel.Application)Marshal.GetActiveObject("Excel.Application");
		Workbook activeWorkbook = application.ActiveWorkbook;
		Worksheet worksheet = (Worksheet)activeWorkbook.Sheets["Sheet4"];
		((_Worksheet)worksheet).get_Range((object)"B2", RuntimeHelpers.GetObjectValue(Missing.Value)).set_Value(RuntimeHelpers.GetObjectValue(Missing.Value), (object)values[0]);
		((_Worksheet)worksheet).get_Range((object)"B3", RuntimeHelpers.GetObjectValue(Missing.Value)).set_Value(RuntimeHelpers.GetObjectValue(Missing.Value), (object)values[1]);
		((_Worksheet)worksheet).get_Range((object)"B4", RuntimeHelpers.GetObjectValue(Missing.Value)).set_Value(RuntimeHelpers.GetObjectValue(Missing.Value), (object)values[2]);
		((_Worksheet)worksheet).get_Range((object)"C2", RuntimeHelpers.GetObjectValue(Missing.Value)).set_Value(RuntimeHelpers.GetObjectValue(Missing.Value), (object)values[3]);
		((_Worksheet)worksheet).get_Range((object)"C3", RuntimeHelpers.GetObjectValue(Missing.Value)).set_Value(RuntimeHelpers.GetObjectValue(Missing.Value), (object)values[4]);
		((_Worksheet)worksheet).get_Range((object)"C4", RuntimeHelpers.GetObjectValue(Missing.Value)).set_Value(RuntimeHelpers.GetObjectValue(Missing.Value), (object)values[5]);
		((_Worksheet)worksheet).get_Range((object)"D2", RuntimeHelpers.GetObjectValue(Missing.Value)).set_Value(RuntimeHelpers.GetObjectValue(Missing.Value), (object)values[6]);
		((_Worksheet)worksheet).get_Range((object)"D3", RuntimeHelpers.GetObjectValue(Missing.Value)).set_Value(RuntimeHelpers.GetObjectValue(Missing.Value), (object)values[7]);
		((_Worksheet)worksheet).get_Range((object)"D4", RuntimeHelpers.GetObjectValue(Missing.Value)).set_Value(RuntimeHelpers.GetObjectValue(Missing.Value), (object)values[8]);
		((_Worksheet)worksheet).get_Range((object)"E2", RuntimeHelpers.GetObjectValue(Missing.Value)).set_Value(RuntimeHelpers.GetObjectValue(Missing.Value), (object)values[9]);
	}
	catch (Exception ex)
	{
		ProjectData.SetProjectError(ex);
		Exception ex2 = ex;
		MessageBox.Show("Error writing values to Excel: " + ex2.Message);
		ProjectData.ClearProjectError();
	}
}
