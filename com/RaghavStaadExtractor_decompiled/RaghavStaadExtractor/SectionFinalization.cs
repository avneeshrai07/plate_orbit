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
[Guid("1B3A4D71-5A2D-4B7E-9A2F-1EFA2E4C0A55")]
[ProgId("RaghavStaadExtractor.SectionFinalization")]
public class SectionFinalization
{
	[ComVisible(true)]
	public void FinalSummarySheet2()
	{
		Application application = null;
		Worksheet worksheet = null;
		Worksheet worksheet2 = null;
		Dictionary<string, double[]> dictionary = new Dictionary<string, double[]>();
		long num = 9L;
		long num2 = 9L;
		long num3 = 1L;
		checked
		{
			try
			{
				application = (Application)Marshal.GetActiveObject("Excel.Application");
				worksheet = (Worksheet)application.ActiveWorkbook.Sheets["Sheet2"];
				worksheet2 = (Worksheet)application.ActiveWorkbook.Sheets["Sheet1"];
				if (((_Worksheet)worksheet).get_Range((object)"A2", RuntimeHelpers.GetObjectValue(Missing.Value)).get_Value(RuntimeHelpers.GetObjectValue(Missing.Value)) != null && Operators.CompareString(Strings.Trim(Conversions.ToString(((_Worksheet)worksheet).get_Range((object)"A2", RuntimeHelpers.GetObjectValue(Missing.Value)).get_Value(RuntimeHelpers.GetObjectValue(Missing.Value)))), "", TextCompare: false) != 0)
				{
					Interaction.MsgBox("Data Already Summarized.\r\nTo Summarize Again, Press B1 in Sheet 2", MsgBoxStyle.Exclamation);
					return;
				}
				((_Worksheet)worksheet).get_Range((object)"A1:F7", RuntimeHelpers.GetObjectValue(Missing.Value)).ClearContents();
				((_Worksheet)worksheet).get_Range((object)"C2", RuntimeHelpers.GetObjectValue(Missing.Value)).set_Value(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(((_Worksheet)worksheet2).get_Range((object)"E2", RuntimeHelpers.GetObjectValue(Missing.Value)).get_Value(RuntimeHelpers.GetObjectValue(Missing.Value))));
				((_Worksheet)worksheet).get_Range((object)"C3", RuntimeHelpers.GetObjectValue(Missing.Value)).set_Value(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(((_Worksheet)worksheet2).get_Range((object)"E3", RuntimeHelpers.GetObjectValue(Missing.Value)).get_Value(RuntimeHelpers.GetObjectValue(Missing.Value))));
				((_Worksheet)worksheet).get_Range((object)"C4", RuntimeHelpers.GetObjectValue(Missing.Value)).set_Value(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(((_Worksheet)worksheet2).get_Range((object)"E4", RuntimeHelpers.GetObjectValue(Missing.Value)).get_Value(RuntimeHelpers.GetObjectValue(Missing.Value))));
				((_Worksheet)worksheet).get_Range((object)"C5", RuntimeHelpers.GetObjectValue(Missing.Value)).set_Value(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(((_Worksheet)worksheet2).get_Range((object)"E5", RuntimeHelpers.GetObjectValue(Missing.Value)).get_Value(RuntimeHelpers.GetObjectValue(Missing.Value))));
				((_Worksheet)worksheet).get_Range((object)"C6", RuntimeHelpers.GetObjectValue(Missing.Value)).set_Value(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(((_Worksheet)worksheet2).get_Range((object)"AO5", RuntimeHelpers.GetObjectValue(Missing.Value)).get_Value(RuntimeHelpers.GetObjectValue(Missing.Value))));
				((_Worksheet)worksheet).get_Range((object)"A8", RuntimeHelpers.GetObjectValue(Missing.Value)).set_Value(RuntimeHelpers.GetObjectValue(Missing.Value), (object)"SR.NO ");
				((_Worksheet)worksheet).get_Range((object)"A1", RuntimeHelpers.GetObjectValue(Missing.Value)).set_Value(RuntimeHelpers.GetObjectValue(Missing.Value), (object)"MATERIAL SUMMARY LIST");
				((_Worksheet)worksheet).get_Range((object)"A7", RuntimeHelpers.GetObjectValue(Missing.Value)).set_Value(RuntimeHelpers.GetObjectValue(Missing.Value), (object)"");
				((_Worksheet)worksheet).get_Range((object)"A2:B2", RuntimeHelpers.GetObjectValue(Missing.Value)).Merge(RuntimeHelpers.GetObjectValue(Missing.Value));
				((_Worksheet)worksheet).get_Range((object)"A3:B3", RuntimeHelpers.GetObjectValue(Missing.Value)).Merge(RuntimeHelpers.GetObjectValue(Missing.Value));
				((_Worksheet)worksheet).get_Range((object)"A4:B4", RuntimeHelpers.GetObjectValue(Missing.Value)).Merge(RuntimeHelpers.GetObjectValue(Missing.Value));
				((_Worksheet)worksheet).get_Range((object)"A5:B5", RuntimeHelpers.GetObjectValue(Missing.Value)).Merge(RuntimeHelpers.GetObjectValue(Missing.Value));
				((_Worksheet)worksheet).get_Range((object)"A6:B6", RuntimeHelpers.GetObjectValue(Missing.Value)).Merge(RuntimeHelpers.GetObjectValue(Missing.Value));
				((_Worksheet)worksheet).get_Range((object)"A2", RuntimeHelpers.GetObjectValue(Missing.Value)).set_Value(RuntimeHelpers.GetObjectValue(Missing.Value), (object)"JOB NAME:-");
				((_Worksheet)worksheet).get_Range((object)"A3", RuntimeHelpers.GetObjectValue(Missing.Value)).set_Value(RuntimeHelpers.GetObjectValue(Missing.Value), (object)"JOB NUMBER:-");
				((_Worksheet)worksheet).get_Range((object)"A4", RuntimeHelpers.GetObjectValue(Missing.Value)).set_Value(RuntimeHelpers.GetObjectValue(Missing.Value), (object)"ENGG. NAME:-");
				((_Worksheet)worksheet).get_Range((object)"A5", RuntimeHelpers.GetObjectValue(Missing.Value)).set_Value(RuntimeHelpers.GetObjectValue(Missing.Value), (object)"PART:-");
				((_Worksheet)worksheet).get_Range((object)"A6", RuntimeHelpers.GetObjectValue(Missing.Value)).set_Value(RuntimeHelpers.GetObjectValue(Missing.Value), (object)"REVISION:-");
				((_Worksheet)worksheet).get_Range((object)"D6", RuntimeHelpers.GetObjectValue(Missing.Value)).set_Value(RuntimeHelpers.GetObjectValue(Missing.Value), (object)"DATE:-");
				((_Worksheet)worksheet).get_Range((object)"E6", RuntimeHelpers.GetObjectValue(Missing.Value)).set_Value(RuntimeHelpers.GetObjectValue(Missing.Value), (object)Strings.Format(DateAndTime.Now.Date, "dd-MMM-yyyy"));
				((_Worksheet)worksheet).get_Range((object)"E4", RuntimeHelpers.GetObjectValue(Missing.Value)).set_Value(RuntimeHelpers.GetObjectValue(Missing.Value), (object)"TOTAL WT:-");
				((_Worksheet)worksheet).get_Range((object)"A1:F8", RuntimeHelpers.GetObjectValue(Missing.Value)).Borders.LineStyle = XlLineStyle.xlContinuous;
				((_Worksheet)worksheet).get_Range((object)"A1:F1,A7:F7", RuntimeHelpers.GetObjectValue(Missing.Value)).Merge(RuntimeHelpers.GetObjectValue(Missing.Value));
				((_Worksheet)worksheet).get_Range((object)"A1:E7", RuntimeHelpers.GetObjectValue(Missing.Value)).Font.Size = 11.5;
				((_Worksheet)worksheet).get_Range((object)"A1:F8", RuntimeHelpers.GetObjectValue(Missing.Value)).Font.Bold = true;
				((_Worksheet)worksheet).get_Range((object)"A1:F1,A7:F7", RuntimeHelpers.GetObjectValue(Missing.Value)).Interior.Color = Information.RGB(248, 203, 173);
				((_Worksheet)worksheet).get_Range((object)"A1,A7", RuntimeHelpers.GetObjectValue(Missing.Value)).HorizontalAlignment = XlHAlign.xlHAlignCenter;
				((_Worksheet)worksheet).get_Range((object)"C2:F2", RuntimeHelpers.GetObjectValue(Missing.Value)).Merge(RuntimeHelpers.GetObjectValue(Missing.Value));
				((_Worksheet)worksheet).get_Range((object)"C3:F3", RuntimeHelpers.GetObjectValue(Missing.Value)).Merge(RuntimeHelpers.GetObjectValue(Missing.Value));
				((_Worksheet)worksheet).get_Range((object)"C4:D4", RuntimeHelpers.GetObjectValue(Missing.Value)).Merge(RuntimeHelpers.GetObjectValue(Missing.Value));
				((_Worksheet)worksheet).get_Range((object)"C5:E5", RuntimeHelpers.GetObjectValue(Missing.Value)).Merge(RuntimeHelpers.GetObjectValue(Missing.Value));
				long num4 = Conversions.ToLong(NewLateBinding.LateGet(NewLateBinding.LateGet(worksheet.Cells[worksheet.Rows.Count, "B"], null, "End", new object[1] { XlDirection.xlUp }, null, null, null), null, "Row", new object[0], null, null, null));
				long num5 = num;
				long num6 = num4;
				for (long num7 = num5; num7 <= num6; num7++)
				{
					string text = Strings.Trim(Conversions.ToString(NewLateBinding.LateGet(worksheet.Cells[num7, "B"], null, "Value", new object[0], null, null, null)));
					if (Operators.CompareString(text, "", TextCompare: false) != 0)
					{
						double num8 = Conversion.Val(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(worksheet.Cells[num7, "C"], null, "Value", new object[0], null, null, null)));
						double num9 = Conversion.Val(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(worksheet.Cells[num7, "D"], null, "Value", new object[0], null, null, null)));
						double num10 = Conversion.Val(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(worksheet.Cells[num7, "E"], null, "Value", new object[0], null, null, null)));
						if (!dictionary.ContainsKey(text))
						{
							dictionary.Add(text, new double[3]);
						}
						double[] array = dictionary[text];
						array[0] += num8;
						array[1] += num9;
						array[2] += num10;
						dictionary[text] = array;
					}
				}
				((_Worksheet)worksheet).get_Range((object)("A" + Conversions.ToString(num2) + ":G" + Conversions.ToString(worksheet.Rows.Count)), RuntimeHelpers.GetObjectValue(Missing.Value)).ClearContents();
				long num11 = num2;
				foreach (string key in dictionary.Keys)
				{
					double[] array2 = dictionary[key];
					double num12 = array2[0] * array2[1] / 1000000.0;
					NewLateBinding.LateSetComplex(worksheet.Cells[num11, "A"], null, "Value", new object[1] { num3 }, null, null, OptimisticSet: false, RValueBase: true);
					NewLateBinding.LateSetComplex(worksheet.Cells[num11, "B"], null, "Value", new object[1] { key }, null, null, OptimisticSet: false, RValueBase: true);
					NewLateBinding.LateSetComplex(worksheet.Cells[num11, "E"], null, "Value", new object[1] { array2[2] }, null, null, OptimisticSet: false, RValueBase: true);
					if (num12 > 0.0)
					{
						NewLateBinding.LateSetComplex(worksheet.Cells[num11, "C"], null, "Value", new object[1] { "-" }, null, null, OptimisticSet: false, RValueBase: true);
						NewLateBinding.LateSetComplex(worksheet.Cells[num11, "D"], null, "Value", new object[1] { "-" }, null, null, OptimisticSet: false, RValueBase: true);
					}
					else
					{
						NewLateBinding.LateSetComplex(worksheet.Cells[num11, "C"], null, "Value", new object[1] { array2[0] }, null, null, OptimisticSet: false, RValueBase: true);
						NewLateBinding.LateSetComplex(worksheet.Cells[num11, "D"], null, "Value", new object[1] { "-" }, null, null, OptimisticSet: false, RValueBase: true);
					}
					NewLateBinding.LateSetComplex(worksheet.Cells[num11, "G"], null, "Value", new object[1] { Strings.Len(Conversions.ToString(NewLateBinding.LateGet(worksheet.Cells[num11, "B"], null, "Value", new object[0], null, null, null))) }, null, null, OptimisticSet: false, RValueBase: true);
					num11++;
					num3++;
				}
				long num13 = num11 - 1;
				worksheet.Sort.SortFields.Clear();
				worksheet.Sort.SortFields.Add(((_Worksheet)worksheet).get_Range((object)("G" + Conversions.ToString(num2) + ":G" + Conversions.ToString(num13)), RuntimeHelpers.GetObjectValue(Missing.Value)), XlSortOn.xlSortOnValues, XlSortOrder.xlAscending, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
				worksheet.Sort.SetRange(((_Worksheet)worksheet).get_Range((object)("A" + Conversions.ToString(num2) + ":G" + Conversions.ToString(num13)), RuntimeHelpers.GetObjectValue(Missing.Value)));
				worksheet.Sort.Header = XlYesNoGuess.xlNo;
				worksheet.Sort.Apply();
				long num14 = num2;
				long num15 = num13;
				for (long num16 = num14; num16 <= num15; num16++)
				{
					NewLateBinding.LateSetComplex(worksheet.Cells[num16, "A"], null, "Value", new object[1] { num16 - num2 + 1 }, null, null, OptimisticSet: false, RValueBase: true);
				}
				((_Worksheet)worksheet).get_Range((object)("G" + Conversions.ToString(num2) + ":G" + Conversions.ToString(num13)), RuntimeHelpers.GetObjectValue(Missing.Value)).ClearContents();
				if (num13 >= num2)
				{
					Range range = ((_Worksheet)worksheet).get_Range((object)("A" + Conversions.ToString(num2) + ":F" + Conversions.ToString(num13)), RuntimeHelpers.GetObjectValue(Missing.Value));
					range.Borders.LineStyle = XlLineStyle.xlContinuous;
					range.HorizontalAlignment = XlHAlign.xlHAlignLeft;
					range = null;
				}
				((_Worksheet)worksheet).get_Range((object)("F8:F" + Conversions.ToString(num13)), RuntimeHelpers.GetObjectValue(Missing.Value)).Merge(RuntimeHelpers.GetObjectValue(Missing.Value));
				Interaction.MsgBox("DATA SUMMERIZED.", MsgBoxStyle.Information);
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
				dictionary.Clear();
				dictionary = null;
				worksheet = null;
				worksheet2 = null;
				application = null;
			}
		}
	}
}
