using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Core;

[ComImport]
[CompilerGenerated]
[Guid("000C0304-0000-0000-C000-000000000046")]
[TypeIdentifier]
public interface CommandBar : _IMsoOleAccDispObj
{
	void _VtblGap1_53();

	[DispId(1610874910)]
	bool Visible
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1610874910)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1610874910)]
		[param: In]
		set;
	}
}
