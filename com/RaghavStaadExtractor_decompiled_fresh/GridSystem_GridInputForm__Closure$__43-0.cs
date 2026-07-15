// ===== RaghavStaadExtractor.GridSystem.GridInputForm._Closure$__43-0 =====

// Decompiled member-by-member from RaghavStaadExtractor.dll



using System.Windows.Forms;

public Timer $VB$Local_timer;


using System.Windows.Forms;

public Control $VB$Local_previouslyFocusedControl;


public GridInputForm $VB$Me;


public _Closure$__43-0()
{
}


using System;
using System.Drawing;
using System.Runtime.CompilerServices;

[SpecialName]
internal void _Lambda$__0(object s, EventArgs args)
{
	$VB$Me.txtB2.Parent.BackColor = (string.IsNullOrWhiteSpace($VB$Me.txtB2.Text) ? NeutralLight : Color.FromArgb(212, 237, 218));
	$VB$Me.txtB4.Parent.BackColor = (string.IsNullOrWhiteSpace($VB$Me.txtB4.Text) ? NeutralLight : Color.FromArgb(212, 237, 218));
	$VB$Local_timer.Stop();
	$VB$Local_timer.Dispose();
	if ($VB$Local_previouslyFocusedControl != null)
	{
		$VB$Local_previouslyFocusedControl.Focus();
	}
}
