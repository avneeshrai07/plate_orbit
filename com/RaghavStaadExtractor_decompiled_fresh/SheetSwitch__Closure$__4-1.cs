// ===== RaghavStaadExtractor.SheetSwitch._Closure$__4-1 =====

// Decompiled member-by-member from RaghavStaadExtractor.dll



using System.Windows.Forms;

public Timer $VB$Local_fadeTimer;


public _Closure$__4-0 $VB$NonLocal_$VB$Closure_2;


public _Closure$__4-1()
{
}


using System;
using System.Runtime.CompilerServices;

[SpecialName]
internal void _Lambda$__5(object s, EventArgs args)
{
	$VB$NonLocal_$VB$Closure_2.$VB$Local_form.Opacity += 0.05;
	if ($VB$NonLocal_$VB$Closure_2.$VB$Local_form.Opacity >= 1.0)
	{
		$VB$Local_fadeTimer.Stop();
		$VB$Local_fadeTimer.Dispose();
	}
}
