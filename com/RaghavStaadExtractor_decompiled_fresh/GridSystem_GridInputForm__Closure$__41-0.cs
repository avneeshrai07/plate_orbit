// ===== RaghavStaadExtractor.GridSystem.GridInputForm._Closure$__41-0 =====

// Decompiled member-by-member from RaghavStaadExtractor.dll



using System.Drawing;

public Color $VB$Local_originalColor;


using System.Windows.Forms;

public Timer $VB$Local_timer;


public _Closure$__41-1 $VB$NonLocal_$VB$Closure_2;


public _Closure$__41-0()
{
}


using System;
using System.Runtime.CompilerServices;

[SpecialName]
internal void _Lambda$__0(object s, EventArgs args)
{
	$VB$NonLocal_$VB$Closure_2.$VB$Me.txtB2.Parent.BackColor = $VB$Local_originalColor;
	$VB$Local_timer.Stop();
	$VB$Local_timer.Dispose();
	if ($VB$NonLocal_$VB$Closure_2.$VB$Local_previouslyFocusedControl != null)
	{
		$VB$NonLocal_$VB$Closure_2.$VB$Local_previouslyFocusedControl.Focus();
	}
}
