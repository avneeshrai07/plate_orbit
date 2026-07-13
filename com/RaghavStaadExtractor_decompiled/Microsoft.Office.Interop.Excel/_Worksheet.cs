using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.Excel;

[ComImport]
[CompilerGenerated]
[Guid("000208D8-0000-0000-C000-000000000046")]
[TypeIdentifier]
public interface _Worksheet
{
	void _VtblGap1_2();

	[DispId(150)]
	object Parent
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(150)]
		[return: MarshalAs(UnmanagedType.IDispatch)]
		get;
	}

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[LCIDConversion(0)]
	[DispId(304)]
	void Activate();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(551)]
	[LCIDConversion(2)]
	void Copy([Optional][In][MarshalAs(UnmanagedType.Struct)] object Before, [Optional][In][MarshalAs(UnmanagedType.Struct)] object After);

	void _VtblGap2_6();

	[DispId(110)]
	string Name
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(110)]
		[return: MarshalAs(UnmanagedType.BStr)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(110)]
		[param: In]
		[param: MarshalAs(UnmanagedType.BStr)]
		set;
	}

	void _VtblGap3_18();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[LCIDConversion(1)]
	[DispId(285)]
	void Unprotect([Optional][In][MarshalAs(UnmanagedType.Struct)] object Password);

	[DispId(558)]
	XlSheetVisibility Visible
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[LCIDConversion(0)]
		[DispId(558)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(558)]
		[LCIDConversion(0)]
		[param: In]
		set;
	}

	[DispId(1377)]
	Shapes Shapes
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1377)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap4_10();

	[DispId(238)]
	Range Cells
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(238)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap5_5();

	[DispId(241)]
	Range Columns
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(241)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap6_26();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(799)]
	[LCIDConversion(1)]
	[return: MarshalAs(UnmanagedType.IDispatch)]
	object OLEObjects([Optional][In][MarshalAs(UnmanagedType.Struct)] object Index);

	void _VtblGap7_14();

	[DispId(197)]
	Range Range
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(197)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap8_1();

	[DispId(258)]
	Range Rows
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(258)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap9_39();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(2029)]
	void Protect([Optional][In][MarshalAs(UnmanagedType.Struct)] object Password, [Optional][In][MarshalAs(UnmanagedType.Struct)] object DrawingObjects, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Contents, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Scenarios, [Optional][In][MarshalAs(UnmanagedType.Struct)] object UserInterfaceOnly, [Optional][In][MarshalAs(UnmanagedType.Struct)] object AllowFormattingCells, [Optional][In][MarshalAs(UnmanagedType.Struct)] object AllowFormattingColumns, [Optional][In][MarshalAs(UnmanagedType.Struct)] object AllowFormattingRows, [Optional][In][MarshalAs(UnmanagedType.Struct)] object AllowInsertingColumns, [Optional][In][MarshalAs(UnmanagedType.Struct)] object AllowInsertingRows, [Optional][In][MarshalAs(UnmanagedType.Struct)] object AllowInsertingHyperlinks, [Optional][In][MarshalAs(UnmanagedType.Struct)] object AllowDeletingColumns, [Optional][In][MarshalAs(UnmanagedType.Struct)] object AllowDeletingRows, [Optional][In][MarshalAs(UnmanagedType.Struct)] object AllowSorting, [Optional][In][MarshalAs(UnmanagedType.Struct)] object AllowFiltering, [Optional][In][MarshalAs(UnmanagedType.Struct)] object AllowUsingPivotTables);

	void _VtblGap10_6();

	[DispId(880)]
	Sort Sort
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(880)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}
}
