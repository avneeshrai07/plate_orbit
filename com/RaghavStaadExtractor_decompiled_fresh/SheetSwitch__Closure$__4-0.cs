// ===== RaghavStaadExtractor.SheetSwitch._Closure$__4-0 =====

// Decompiled member-by-member from RaghavStaadExtractor.dll



using System.Windows.Forms;

public Form $VB$Local_form;


using System.Windows.Forms;

public Button $VB$Local_btnOK;


using System.Windows.Forms;

public Button $VB$Local_btnCancel;


public string $VB$Local_passwordResult;


using System.Windows.Forms;

public TextBox $VB$Local_txtPassword;


public bool $VB$Local_isDragging;


using System.Drawing;

public Point $VB$Local_dragStartPoint;


public _Closure$__4-0()
{
}


using System.Drawing;
using System.Runtime.CompilerServices;
using System.Windows.Forms;

[SpecialName]
internal void _Lambda$__0(object sender, PaintEventArgs e)
{
	checked
	{
		using Pen pen = new Pen(Color.FromArgb(120, 120, 120), 2f);
		e.Graphics.DrawRectangle(pen, 0, 0, $VB$Local_form.Width - 1, $VB$Local_form.Height - 1);
	}
}


using System.Runtime.CompilerServices;
using System.Windows.Forms;

[SpecialName]
internal void _Lambda$__1(object sender, KeyPressEventArgs e)
{
	if (e.KeyChar == '\r')
	{
		e.Handled = true;
		$VB$Local_btnOK.PerformClick();
	}
	else if (e.KeyChar == '\u001b')
	{
		e.Handled = true;
		$VB$Local_btnCancel.PerformClick();
	}
}


using System;
using System.Runtime.CompilerServices;
using System.Windows.Forms;

[SpecialName]
internal void _Lambda$__2(object sender, EventArgs e)
{
	$VB$Local_passwordResult = $VB$Local_txtPassword.Text;
	$VB$Local_form.DialogResult = DialogResult.OK;
	$VB$Local_form.Close();
}


using System;
using System.Runtime.CompilerServices;
using System.Windows.Forms;

[SpecialName]
internal void _Lambda$__3(object sender, EventArgs e)
{
	$VB$Local_passwordResult = null;
	$VB$Local_form.DialogResult = DialogResult.Cancel;
	$VB$Local_form.Close();
}


using System;
using System.Runtime.CompilerServices;
using System.Windows.Forms;

[SpecialName]
internal void _Lambda$__4(object sender, EventArgs e)
{
	_Closure$__4-1 CS$<>8__locals0 = new _Closure$__4-1
	{
		$VB$NonLocal_$VB$Closure_2 = this
	};
	$VB$Local_txtPassword.Focus();
	$VB$Local_txtPassword.Select();
	$VB$Local_form.Opacity = 0.0;
	CS$<>8__locals0.$VB$Local_fadeTimer = new Timer
	{
		Interval = 10
	};
	CS$<>8__locals0.$VB$Local_fadeTimer.Tick += delegate
	{
		CS$<>8__locals0.$VB$NonLocal_$VB$Closure_2.$VB$Local_form.Opacity += 0.05;
		if (CS$<>8__locals0.$VB$NonLocal_$VB$Closure_2.$VB$Local_form.Opacity >= 1.0)
		{
			CS$<>8__locals0.$VB$Local_fadeTimer.Stop();
			CS$<>8__locals0.$VB$Local_fadeTimer.Dispose();
		}
	};
	CS$<>8__locals0.$VB$Local_fadeTimer.Start();
}


using System.Runtime.CompilerServices;
using System.Windows.Forms;

[SpecialName]
internal void _Lambda$__6(object sender, MouseEventArgs e)
{
	if (e.Button == MouseButtons.Left)
	{
		$VB$Local_isDragging = true;
		$VB$Local_dragStartPoint = e.Location;
	}
}


using System.Drawing;
using System.Runtime.CompilerServices;
using System.Windows.Forms;

[SpecialName]
internal void _Lambda$__7(object sender, MouseEventArgs e)
{
	checked
	{
		if ($VB$Local_isDragging)
		{
			Point location = $VB$Local_form.Location;
			location.X += e.X - $VB$Local_dragStartPoint.X;
			location.Y += e.Y - $VB$Local_dragStartPoint.Y;
			$VB$Local_form.Location = location;
		}
	}
}


using System.Runtime.CompilerServices;
using System.Windows.Forms;

[SpecialName]
internal void _Lambda$__8(object sender, MouseEventArgs e)
{
	$VB$Local_isDragging = false;
}
