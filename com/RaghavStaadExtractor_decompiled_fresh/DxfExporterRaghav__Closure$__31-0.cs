// ===== RaghavStaadExtractor.DxfExporterRaghav._Closure$__31-0 =====

// Decompiled member-by-member from RaghavStaadExtractor.dll



using netDxf;

public Vector3 $VB$Local_localY;


using netDxf;

public Vector3 $VB$Local_localZ;


public _Closure$__31-0()
{
}


using System.Runtime.CompilerServices;
using netDxf;

[SpecialName]
internal Vector3 _Lambda$__0(double offsetY, double offsetZ, Vector3 basePoint, double webCenterY)
{
	return basePoint + $VB$Local_localY * (offsetY - webCenterY) + $VB$Local_localZ * offsetZ;
}
