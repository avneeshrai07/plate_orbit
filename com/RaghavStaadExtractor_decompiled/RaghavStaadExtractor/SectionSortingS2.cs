using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace RaghavStaadExtractor;

[ComVisible(true)]
[Guid("FC14406F-BD26-4667-B360-98D9A389B651")]
[ClassInterface(ClassInterfaceType.AutoDual)]
public class SectionSortingS2
{
	public SectionSortingS2()
	{
		EnsureLicense();
	}

	private void EnsureLicense()
	{
		ExecutionValidation executionValidation = new ExecutionValidation();
		if (!executionValidation.IsLicenceValid())
		{
			throw new UnauthorizedAccessException("License check failed.");
		}
	}

	public void SummerizedToSheet2()
	{
		checked
		{
			try
			{
				Application application = (Application)Marshal.GetActiveObject("Excel.Application");
				Workbook activeWorkbook = application.ActiveWorkbook;
				Worksheet worksheet = (Worksheet)activeWorkbook.Sheets["Sheet1"];
				Worksheet worksheet2 = (Worksheet)activeWorkbook.Sheets["Sheet2"];
				UnprotectSheetAndWorkbook(worksheet2);
				worksheet2.Cells.Clear();
				((_Worksheet)worksheet2).get_Range((object)"B8", RuntimeHelpers.GetObjectValue(Missing.Value)).set_Value(RuntimeHelpers.GetObjectValue(Missing.Value), (object)"THK./SECTION");
				((_Worksheet)worksheet2).get_Range((object)"C8", RuntimeHelpers.GetObjectValue(Missing.Value)).set_Value(RuntimeHelpers.GetObjectValue(Missing.Value), (object)"LENGTH (mm)");
				((_Worksheet)worksheet2).get_Range((object)"D8", RuntimeHelpers.GetObjectValue(Missing.Value)).set_Value(RuntimeHelpers.GetObjectValue(Missing.Value), (object)"WIDTH(mm)");
				((_Worksheet)worksheet2).get_Range((object)"E8", RuntimeHelpers.GetObjectValue(Missing.Value)).set_Value(RuntimeHelpers.GetObjectValue(Missing.Value), (object)"WEIGHT(Kg)");
				int num = Conversions.ToInteger(NewLateBinding.LateGet(NewLateBinding.LateGet(worksheet.Cells[worksheet.Rows.Count, "C"], null, "End", new object[1] { XlDirection.xlUp }, null, null, null), null, "Row", new object[0], null, null, null));
				object[,] array = (object[,])((_Worksheet)worksheet).get_Range((object)"A9", (object)$"AQ{num}").Value2;
				List<object[]> list = new List<object[]>();
				int length = array.GetLength(0);
				for (int i = 1; i <= length; i++)
				{
					double num2 = ValObj(RuntimeHelpers.GetObjectValue(array[i, 6]));
					double length2 = ValObj(RuntimeHelpers.GetObjectValue(array[i, 3])) * 1000.0;
					if (num2 == 697.0)
					{
						object objectValue = RuntimeHelpers.GetObjectValue(array[i, 37]);
						object objectValue2 = RuntimeHelpers.GetObjectValue(array[i, 38]);
						if (!Information.IsNothing(RuntimeHelpers.GetObjectValue(objectValue)) && Operators.CompareString(Convert.ToString(RuntimeHelpers.GetObjectValue(objectValue)).Trim(), "", TextCompare: false) != 0 && (Information.IsNothing(RuntimeHelpers.GetObjectValue(objectValue2)) || Operators.CompareString(Convert.ToString(RuntimeHelpers.GetObjectValue(objectValue2)).Trim(), "", TextCompare: false) == 0))
						{
							AddIfValid(list, RuntimeHelpers.GetObjectValue(array[i, 25]), length2, RuntimeHelpers.GetObjectValue(array[i, 36]), RuntimeHelpers.GetObjectValue(array[i, 40]));
							AddIfValid(list, RuntimeHelpers.GetObjectValue(array[i, 27]), length2, RuntimeHelpers.GetObjectValue(array[i, 26]), RuntimeHelpers.GetObjectValue(array[i, 41]));
							AddIfValid(list, RuntimeHelpers.GetObjectValue(array[i, 29]), length2, RuntimeHelpers.GetObjectValue(array[i, 28]), RuntimeHelpers.GetObjectValue(array[i, 42]));
						}
						else
						{
							AddIfValidFor697(list, RuntimeHelpers.GetObjectValue(array[i, 37]), length2, RuntimeHelpers.GetObjectValue(array[i, 40]));
							AddIfValidFor697(list, RuntimeHelpers.GetObjectValue(array[i, 38]), length2, RuntimeHelpers.GetObjectValue(array[i, 41]));
							AddIfValidFor697(list, RuntimeHelpers.GetObjectValue(array[i, 39]), length2, RuntimeHelpers.GetObjectValue(array[i, 42]));
						}
					}
					else
					{
						AddIfValid(list, RuntimeHelpers.GetObjectValue(array[i, 25]), length2, RuntimeHelpers.GetObjectValue(array[i, 36]), RuntimeHelpers.GetObjectValue(array[i, 40]));
						AddIfValid(list, RuntimeHelpers.GetObjectValue(array[i, 27]), length2, RuntimeHelpers.GetObjectValue(array[i, 26]), RuntimeHelpers.GetObjectValue(array[i, 41]));
						AddIfValid(list, RuntimeHelpers.GetObjectValue(array[i, 29]), length2, RuntimeHelpers.GetObjectValue(array[i, 28]), RuntimeHelpers.GetObjectValue(array[i, 42]));
					}
				}
				int length3 = array.GetLength(0);
				for (int j = 1; j <= length3; j++)
				{
					double num3 = ValObj(RuntimeHelpers.GetObjectValue(array[j, 6]));
					double num4 = ValObj(RuntimeHelpers.GetObjectValue(array[j, 3])) * 1000.0;
					bool flag = true;
					int num5 = 24;
					do
					{
						if (ValObj(RuntimeHelpers.GetObjectValue(array[j, num5])) != 0.0)
						{
							flag = false;
							break;
						}
						num5++;
					}
					while (num5 <= 36);
					if (!flag)
					{
						continue;
					}
					if (num3 == 697.0)
					{
						object objectValue3 = RuntimeHelpers.GetObjectValue(array[j, 37]);
						object objectValue4 = RuntimeHelpers.GetObjectValue(array[j, 38]);
						if (!Information.IsNothing(RuntimeHelpers.GetObjectValue(objectValue3)) && Operators.CompareString(Convert.ToString(RuntimeHelpers.GetObjectValue(objectValue3)).Trim(), "", TextCompare: false) != 0 && (Information.IsNothing(RuntimeHelpers.GetObjectValue(objectValue4)) || Operators.CompareString(Convert.ToString(RuntimeHelpers.GetObjectValue(objectValue4)).Trim(), "", TextCompare: false) == 0))
						{
							object objectValue5 = RuntimeHelpers.GetObjectValue(array[j, 37]);
							object objectValue6 = RuntimeHelpers.GetObjectValue(array[j, 43]);
							list.Add(new object[4] { objectValue5, num4, "", objectValue6 });
						}
						else if ((Information.IsNothing(RuntimeHelpers.GetObjectValue(array[j, 37])) || Operators.CompareString(Convert.ToString(RuntimeHelpers.GetObjectValue(array[j, 37])).Trim(), "", TextCompare: false) == 0) && (Information.IsNothing(RuntimeHelpers.GetObjectValue(array[j, 38])) || Operators.CompareString(Convert.ToString(RuntimeHelpers.GetObjectValue(array[j, 38])).Trim(), "", TextCompare: false) == 0) && (Information.IsNothing(RuntimeHelpers.GetObjectValue(array[j, 39])) || Operators.CompareString(Convert.ToString(RuntimeHelpers.GetObjectValue(array[j, 39])).Trim(), "", TextCompare: false) == 0))
						{
							object objectValue7 = RuntimeHelpers.GetObjectValue(array[j, 37]);
							object objectValue8 = RuntimeHelpers.GetObjectValue(array[j, 43]);
							string text = Convert.ToString(RuntimeHelpers.GetObjectValue(objectValue7)).Trim();
							if (text.StartsWith("2x", StringComparison.OrdinalIgnoreCase))
							{
								string sectionName = text.Substring(2);
								(string, string) tuple = ParseSectionName(sectionName);
								double num6 = ValObj(RuntimeHelpers.GetObjectValue(objectValue8)) / 2.0;
								list.Add(new object[4] { tuple.Item1, num4, tuple.Item2, num6 });
								list.Add(new object[4] { tuple.Item1, num4, tuple.Item2, num6 });
							}
							else
							{
								(string, string) tuple2 = ParseSectionName(Convert.ToString(RuntimeHelpers.GetObjectValue(objectValue7)));
								list.Add(new object[4] { tuple2.Item1, num4, tuple2.Item2, objectValue8 });
							}
						}
					}
					else
					{
						object objectValue9 = RuntimeHelpers.GetObjectValue(array[j, 37]);
						object objectValue10 = RuntimeHelpers.GetObjectValue(array[j, 43]);
						list.Add(new object[4] { objectValue9, num4, "", objectValue10 });
					}
				}
				int length4 = array.GetLength(0);
				for (int k = 1; k <= length4; k++)
				{
					double num7 = ValObj(RuntimeHelpers.GetObjectValue(array[k, 6]));
					if (num7 == 613.0 || num7 == 614.0 || num7 == 615.0)
					{
						object objectValue11 = RuntimeHelpers.GetObjectValue(array[k, 37]);
						double num8 = ValObj(RuntimeHelpers.GetObjectValue(array[k, 3])) * 1000.0;
						object objectValue12 = RuntimeHelpers.GetObjectValue(array[k, 40]);
						if (!Information.IsNothing(RuntimeHelpers.GetObjectValue(objectValue11)) && Operators.CompareString(Convert.ToString(RuntimeHelpers.GetObjectValue(objectValue11)).Trim(), "", TextCompare: false) != 0 && ValObj(RuntimeHelpers.GetObjectValue(objectValue12)) != 0.0)
						{
							list.Add(new object[4] { objectValue11, num8, "", objectValue12 });
						}
					}
				}
				List<object[]> list2 = list.OrderBy([SpecialName] (object[] r) => Convert.ToString(RuntimeHelpers.GetObjectValue(r[0]))).ToList();
				Range range = ((_Worksheet)worksheet2).get_Range((object)"B9", (object)$"E{8 + list2.Count}");
				object[,] array2 = new object[list2.Count - 1 + 1, 4];
				int num9 = list2.Count - 1;
				for (int num10 = 0; num10 <= num9; num10++)
				{
					int num11 = 0;
					do
					{
						array2[num10, num11] = RuntimeHelpers.GetObjectValue(list2[num10][num11]);
						num11++;
					}
					while (num11 <= 3);
				}
				range.Value2 = array2;
				((_Worksheet)worksheet2).get_Range((object)"A:E", RuntimeHelpers.GetObjectValue(Missing.Value)).EntireColumn.AutoFit();
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				ProjectData.ClearProjectError();
			}
		}
	}

	private void UnprotectSheetAndWorkbook(Worksheet ws)
	{
		Workbook workbook = (Workbook)ws.Parent;
		try
		{
			ws.Unprotect("2022");
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
		try
		{
			workbook.Unprotect("2022");
		}
		catch (Exception projectError2)
		{
			ProjectData.SetProjectError(projectError2);
			ProjectData.ClearProjectError();
		}
	}

	private void AddIfValidFor697(List<object[]> list, object thk, double length, object weight)
	{
		if (!Information.IsNothing(RuntimeHelpers.GetObjectValue(thk)) && Operators.CompareString(Convert.ToString(RuntimeHelpers.GetObjectValue(thk)).Trim(), "", TextCompare: false) != 0 && ValObj(RuntimeHelpers.GetObjectValue(weight)) != 0.0)
		{
			string text = Convert.ToString(RuntimeHelpers.GetObjectValue(thk)).Trim();
			double num = ValObj(RuntimeHelpers.GetObjectValue(weight));
			if (text.StartsWith("2x", StringComparison.OrdinalIgnoreCase))
			{
				string sectionName = text.Substring(2);
				(string, string) tuple = ParseSectionName(sectionName);
				double num2 = num / 2.0;
				list.Add(new object[4] { tuple.Item1, length, tuple.Item2, num2 });
				list.Add(new object[4] { tuple.Item1, length, tuple.Item2, num2 });
			}
			else
			{
				(string, string) tuple2 = ParseSectionName(text);
				list.Add(new object[4] { tuple2.Item1, length, tuple2.Item2, num });
			}
		}
	}

	private (string Thickness, string Width) ParseSectionName(string sectionName)
	{
		string text = Convert.ToString(sectionName).Trim();
		(string, string) result;
		if (text.Contains("x"))
		{
			string[] array = text.Split('x');
			if (array.Length == 2)
			{
				result = (array[1].Trim(), array[0].Trim());
				goto IL_0065;
			}
		}
		result = (text, "");
		goto IL_0065;
		IL_0065:
		return result;
	}

	private void AddIfValid(List<object[]> list, object thk, double length, object width, object weight)
	{
		if (ValObj(RuntimeHelpers.GetObjectValue(thk)) != 0.0 && ValObj(RuntimeHelpers.GetObjectValue(weight)) != 0.0)
		{
			list.Add(new object[4] { thk, length, width, weight });
		}
	}

	private double ValObj(object v)
	{
		if (double.TryParse(Convert.ToString(RuntimeHelpers.GetObjectValue(v)), out var result))
		{
			return result;
		}
		return 0.0;
	}
}
