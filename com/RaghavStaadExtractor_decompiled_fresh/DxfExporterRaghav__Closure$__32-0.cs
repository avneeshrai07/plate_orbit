// ===== RaghavStaadExtractor.DxfExporterRaghav._Closure$__32-0 =====

// Decompiled member-by-member from RaghavStaadExtractor.dll



using netDxf;

public Vector3 $VB$Local_centerPoint;


using netDxf;

public Vector3 $VB$Local_localY;


public double $VB$Local_webCenterY;


using netDxf;

public Vector3 $VB$Local_localZ;


public _Closure$__32-0()
{
}


using System.Runtime.CompilerServices;
using netDxf;

[SpecialName]
internal Vector3 _Lambda$__0(double offsetY, double offsetZ)
{
	return $VB$Local_centerPoint + $VB$Local_localY * (offsetY - $VB$Local_webCenterY) + $VB$Local_localZ * offsetZ;
}
