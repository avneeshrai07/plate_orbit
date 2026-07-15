// ===== RaghavStaadExtractor.DxfExporterRaghav._Closure$__30-0 =====

// Decompiled member-by-member from RaghavStaadExtractor.dll



using netDxf;

public Vector3 $VB$Local_centerPoint;


using netDxf;

public Vector3 $VB$Local_localY;


using netDxf;

public Vector3 $VB$Local_localZ;


public _Closure$__30-0()
{
}


using System.Runtime.CompilerServices;
using netDxf;

[SpecialName]
internal Vector3 _Lambda$__0(double offsetY, double offsetZ)
{
	return $VB$Local_centerPoint + $VB$Local_localY * offsetY + $VB$Local_localZ * offsetZ;
}
