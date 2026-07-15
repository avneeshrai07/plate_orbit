// ===== RaghavStaadExtractor.StaadDataExtractor._Closure$__ =====

// Decompiled member-by-member from RaghavStaadExtractor.dll



public static readonly _Closure$__ $I;


using System;

public static Func<double, double> $I27-0;


using System;

public static Func<double, bool> $I27-1;


using System;

public static Func<double, double> $I27-2;


using System;

public static Func<double, bool> $I27-3;


using System;

public static Func<double, bool> $I27-4;


using System;

public static Func<double, double> $I27-5;


using System;
using System.Collections.Generic;

public static Func<KeyValuePair<string, List<BeamData>>, bool> $I34-0;


using System;
using System.Collections.Generic;

public static Func<KeyValuePair<string, List<BeamData>>, string> $I34-1;


using System;
using System.Collections.Generic;

public static Func<KeyValuePair<string, List<BeamData>>, List<BeamData>> $I34-2;


using System;

public static Func<BeamData, double> $I36-0;


using System;

public static Func<BeamData, double> $I36-1;


using System;

public static Func<BeamData, double> $I36-2;


using System;

public static Func<BeamData, double> $I36-3;


using System;

public static Func<BeamData, double> $I36-4;


using System;

public static Func<BeamData, double> $I36-5;


using System;

public static Func<BeamData, double> $I36-6;


using System;

public static Func<BeamData, double> $I36-7;


using System;

public static Func<BeamData, double> $I36-8;


using System;

public static Func<BeamData, double> $I36-9;


using System;

public static Func<BeamData, double> $I36-10;


using System;

public static Func<BeamData, double> $I36-11;


using System;

public static Func<BeamData, double> $I36-12;


using System;

public static Func<BeamData, double> $I36-13;


using System;

public static Func<BeamData, double> $I36-14;


using System;

public static Func<BeamData, double> $I36-15;


using System;
using System.Collections.Generic;

public static Func<KeyValuePair<string, List<BeamData>>, string> $I39-0;


using System;

public static Func<BeamData, int> $I39-2;


public _Closure$__()
{
}


static _Closure$__()
{
	$I = new _Closure$__();
}


using System.Runtime.CompilerServices;

[SpecialName]
internal double _Lambda$__27-0(double x)
{
	return x;
}


using System.Runtime.CompilerServices;

[SpecialName]
internal bool _Lambda$__27-1(double c)
{
	return c < -0.005;
}


using System.Runtime.CompilerServices;

[SpecialName]
internal double _Lambda$__27-2(double c)
{
	return c;
}


using System;
using System.Runtime.CompilerServices;

[SpecialName]
internal bool _Lambda$__27-3(double c)
{
	return Math.Abs(c) < 0.005;
}


using System.Runtime.CompilerServices;

[SpecialName]
internal bool _Lambda$__27-4(double c)
{
	return c > 0.005;
}


using System.Runtime.CompilerServices;

[SpecialName]
internal double _Lambda$__27-5(double c)
{
	return c;
}


using System.Collections.Generic;
using System.Runtime.CompilerServices;

[SpecialName]
internal bool _Lambda$__34-0(KeyValuePair<string, List<BeamData>> g)
{
	return g.Value.Count >= 2;
}


using System.Collections.Generic;
using System.Runtime.CompilerServices;

[SpecialName]
internal string _Lambda$__34-1(KeyValuePair<string, List<BeamData>> k)
{
	return k.Key;
}


using System.Collections.Generic;
using System.Runtime.CompilerServices;

[SpecialName]
internal List<BeamData> _Lambda$__34-2(KeyValuePair<string, List<BeamData>> v)
{
	return v.Value;
}


using System;
using System.Runtime.CompilerServices;

[SpecialName]
internal double _Lambda$__36-0(BeamData b)
{
	return Math.Min(b.XA, b.XB);
}


using System;
using System.Runtime.CompilerServices;

[SpecialName]
internal double _Lambda$__36-1(BeamData b)
{
	return Math.Min(b.YA, b.YB);
}


using System;
using System.Runtime.CompilerServices;

[SpecialName]
internal double _Lambda$__36-2(BeamData b)
{
	return Math.Min(b.XA, b.XB);
}


using System;
using System.Runtime.CompilerServices;

[SpecialName]
internal double _Lambda$__36-3(BeamData b)
{
	return Math.Min(b.YA, b.YB);
}


using System;
using System.Runtime.CompilerServices;

[SpecialName]
internal double _Lambda$__36-4(BeamData b)
{
	return Math.Min(b.YA, b.YB);
}


using System;
using System.Runtime.CompilerServices;

[SpecialName]
internal double _Lambda$__36-5(BeamData b)
{
	return Math.Min(b.ZA, b.ZB);
}


using System;
using System.Runtime.CompilerServices;

[SpecialName]
internal double _Lambda$__36-6(BeamData b)
{
	return Math.Min(b.YA, b.YB);
}


using System;
using System.Runtime.CompilerServices;

[SpecialName]
internal double _Lambda$__36-7(BeamData b)
{
	return Math.Min(b.ZA, b.ZB);
}


using System;
using System.Runtime.CompilerServices;

[SpecialName]
internal double _Lambda$__36-8(BeamData b)
{
	return Math.Min(b.XA, b.XB);
}


using System;
using System.Runtime.CompilerServices;

[SpecialName]
internal double _Lambda$__36-9(BeamData b)
{
	return Math.Min(b.YA, b.YB);
}


using System;
using System.Runtime.CompilerServices;

[SpecialName]
internal double _Lambda$__36-10(BeamData b)
{
	return Math.Min(b.XA, b.XB);
}


using System;
using System.Runtime.CompilerServices;

[SpecialName]
internal double _Lambda$__36-11(BeamData b)
{
	return Math.Min(b.YA, b.YB);
}


using System;
using System.Runtime.CompilerServices;

[SpecialName]
internal double _Lambda$__36-12(BeamData b)
{
	return Math.Min(b.YA, b.YB);
}


using System;
using System.Runtime.CompilerServices;

[SpecialName]
internal double _Lambda$__36-13(BeamData b)
{
	return Math.Min(b.ZA, b.ZB);
}


using System;
using System.Runtime.CompilerServices;

[SpecialName]
internal double _Lambda$__36-14(BeamData b)
{
	return Math.Min(b.YA, b.YB);
}


using System;
using System.Runtime.CompilerServices;

[SpecialName]
internal double _Lambda$__36-15(BeamData b)
{
	return Math.Min(b.ZA, b.ZB);
}


using System.Collections.Generic;
using System.Runtime.CompilerServices;

[SpecialName]
internal string _Lambda$__39-0(KeyValuePair<string, List<BeamData>> g)
{
	return g.Key.Split('(')[0].Trim();
}


using System.Runtime.CompilerServices;

[SpecialName]
internal int _Lambda$__39-2(BeamData b)
{
	return b.BeamNo;
}
