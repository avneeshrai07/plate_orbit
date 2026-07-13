using System.Collections;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.Excel;

[ComImport]
[CompilerGenerated]
[Guid("000244AA-0000-0000-C000-000000000046")]
[InterfaceType(2)]
[TypeIdentifier]
public interface SortFields : IEnumerable
{
	void _VtblGap1_3();

	[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(181)]
	[return: MarshalAs(UnmanagedType.Interface)]
	SortField Add([In][MarshalAs(UnmanagedType.Interface)] Range Key, [Optional][In][MarshalAs(UnmanagedType.Struct)] object SortOn, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Order, [Optional][In][MarshalAs(UnmanagedType.Struct)] object CustomOrder, [Optional][In][MarshalAs(UnmanagedType.Struct)] object DataOption);

	void _VtblGap2_2();

	[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(111)]
	void Clear();

	[IndexerName("_Default")]
	[DispId(0)]
	SortField this[[In][MarshalAs(UnmanagedType.Struct)] object Index]
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(0)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}
}
