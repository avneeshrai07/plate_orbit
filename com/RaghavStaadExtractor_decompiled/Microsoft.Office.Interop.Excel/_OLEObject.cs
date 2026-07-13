using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.Excel;

[ComImport]
[CompilerGenerated]
[InterfaceType(2)]
[Guid("000208A2-0000-0000-C000-000000000046")]
[TypeIdentifier]
public interface _OLEObject
{
	void _VtblGap1_8();

	[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(-2147417995)]
	[return: MarshalAs(UnmanagedType.Struct)]
	object Delete();
}
