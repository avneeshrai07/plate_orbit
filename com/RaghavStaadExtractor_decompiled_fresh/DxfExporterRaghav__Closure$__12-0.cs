// ===== RaghavStaadExtractor.DxfExporterRaghav._Closure$__12-0 =====

// Decompiled member-by-member from RaghavStaadExtractor.dll



using System.Windows.Forms;

public Form $VB$Local_inputForm;


using System.Windows.Forms;

public Panel $VB$Local_headerPanel;


using System.Windows.Forms;

public TextBox $VB$Local_txtInput;


using System.Windows.Forms;

public Button $VB$Local_btnOK;


using System.Windows.Forms;

public Button $VB$Local_btnCancel;


public _Closure$__12-0()
{
}


using System;
using System.Diagnostics;
using System.Runtime.CompilerServices;

[SpecialName]
[DebuggerHidden]
internal void _Lambda$__R1(object a0, EventArgs a1)
{
	_Lambda$__0();
}


using System.Drawing;
using System.Drawing.Drawing2D;
using System.Runtime.CompilerServices;

[SpecialName]
internal void _Lambda$__0()
{
	Rectangle rectangle = new Rectangle(0, 0, $VB$Local_inputForm.Width, $VB$Local_inputForm.Height);
	checked
	{
		using GraphicsPath graphicsPath = new GraphicsPath();
		int num = 12;
		graphicsPath.AddArc(rectangle.X, rectangle.Y, num, num, 180f, 90f);
		graphicsPath.AddArc(rectangle.X + rectangle.Width - num, rectangle.Y, num, num, 270f, 90f);
		graphicsPath.AddArc(rectangle.X + rectangle.Width - num, rectangle.Y + rectangle.Height - num, num, num, 0f, 90f);
		graphicsPath.AddArc(rectangle.X, rectangle.Y + rectangle.Height - num, num, num, 90f, 90f);
		graphicsPath.CloseFigure();
		$VB$Local_inputForm.Region = new Region(graphicsPath);
	}
}


using System.Drawing;
using System.Drawing.Drawing2D;
using System.Runtime.CompilerServices;
using System.Windows.Forms;

[SpecialName]
internal void _Lambda$__1(object sender, PaintEventArgs e)
{
	e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
	checked
	{
		Rectangle rectangle = new Rectangle(2, 2, $VB$Local_inputForm.Width - 2, $VB$Local_inputForm.Height - 2);
		using GraphicsPath graphicsPath = new GraphicsPath();
		int num = 12;
		graphicsPath.AddArc(rectangle.X, rectangle.Y, num, num, 180f, 90f);
		graphicsPath.AddArc(rectangle.X + rectangle.Width - num, rectangle.Y, num, num, 270f, 90f);
		graphicsPath.AddArc(rectangle.X + rectangle.Width - num, rectangle.Y + rectangle.Height - num, num, num, 0f, 90f);
		graphicsPath.AddArc(rectangle.X, rectangle.Y + rectangle.Height - num, num, num, 90f, 90f);
		graphicsPath.CloseFigure();
		using SolidBrush brush = new SolidBrush(Color.FromArgb(30, 0, 0, 0));
		e.Graphics.FillPath(brush, graphicsPath);
	}
}


using System.Drawing;
using System.Drawing.Drawing2D;
using System.Runtime.CompilerServices;
using System.Windows.Forms;

[SpecialName]
internal void _Lambda$__2(object sender, PaintEventArgs e)
{
	Rectangle rect = new Rectangle(0, 0, $VB$Local_headerPanel.Width, $VB$Local_headerPanel.Height);
	using LinearGradientBrush brush = new LinearGradientBrush(rect, Color.FromArgb(16, 137, 230), Color.FromArgb(0, 120, 215), LinearGradientMode.Vertical);
	e.Graphics.FillRectangle(brush, rect);
}


using System.Drawing;
using System.Drawing.Drawing2D;
using System.Runtime.CompilerServices;
using System.Windows.Forms;

[SpecialName]
internal void _Lambda$__3(object sender, PaintEventArgs e)
{
	checked
	{
		Rectangle rectangle = new Rectangle(0, 0, $VB$Local_txtInput.Width - 1, $VB$Local_txtInput.Height - 1);
		using Pen pen = new Pen(Color.FromArgb(180, 180, 180), 1f);
		e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
		using GraphicsPath graphicsPath = new GraphicsPath();
		int num = 4;
		graphicsPath.AddArc(rectangle.X, rectangle.Y, num, num, 180f, 90f);
		graphicsPath.AddArc(rectangle.X + rectangle.Width - num, rectangle.Y, num, num, 270f, 90f);
		graphicsPath.AddArc(rectangle.X + rectangle.Width - num, rectangle.Y + rectangle.Height - num, num, num, 0f, 90f);
		graphicsPath.AddArc(rectangle.X, rectangle.Y + rectangle.Height - num, num, num, 90f, 90f);
		graphicsPath.CloseFigure();
		e.Graphics.DrawPath(pen, graphicsPath);
	}
}


using System.Drawing;
using System.Drawing.Drawing2D;
using System.Runtime.CompilerServices;
using System.Windows.Forms;

[SpecialName]
internal void _Lambda$__4(object sender, PaintEventArgs e)
{
	checked
	{
		using GraphicsPath graphicsPath = new GraphicsPath();
		Rectangle rectangle = new Rectangle(0, 0, $VB$Local_btnOK.Width, $VB$Local_btnOK.Height);
		int num = 6;
		graphicsPath.AddArc(rectangle.X, rectangle.Y, num, num, 180f, 90f);
		graphicsPath.AddArc(rectangle.X + rectangle.Width - num, rectangle.Y, num, num, 270f, 90f);
		graphicsPath.AddArc(rectangle.X + rectangle.Width - num, rectangle.Y + rectangle.Height - num, num, num, 0f, 90f);
		graphicsPath.AddArc(rectangle.X, rectangle.Y + rectangle.Height - num, num, num, 90f, 90f);
		graphicsPath.CloseFigure();
		$VB$Local_btnOK.Region = new Region(graphicsPath);
	}
}


using System.Drawing;
using System.Drawing.Drawing2D;
using System.Runtime.CompilerServices;
using System.Windows.Forms;

[SpecialName]
internal void _Lambda$__5(object sender, PaintEventArgs e)
{
	checked
	{
		using GraphicsPath graphicsPath = new GraphicsPath();
		Rectangle rectangle = new Rectangle(0, 0, $VB$Local_btnCancel.Width, $VB$Local_btnCancel.Height);
		int num = 6;
		graphicsPath.AddArc(rectangle.X, rectangle.Y, num, num, 180f, 90f);
		graphicsPath.AddArc(rectangle.X + rectangle.Width - num, rectangle.Y, num, num, 270f, 90f);
		graphicsPath.AddArc(rectangle.X + rectangle.Width - num, rectangle.Y + rectangle.Height - num, num, num, 0f, 90f);
		graphicsPath.AddArc(rectangle.X, rectangle.Y + rectangle.Height - num, num, num, 90f, 90f);
		graphicsPath.CloseFigure();
		$VB$Local_btnCancel.Region = new Region(graphicsPath);
	}
}


using System.Runtime.CompilerServices;
using System.Windows.Forms;

[SpecialName]
internal void _Lambda$__6(object sender, KeyEventArgs e)
{
	if (e.KeyCode == Keys.Return)
	{
		$VB$Local_btnOK.PerformClick();
	}
	else if (e.KeyCode == Keys.Escape)
	{
		$VB$Local_btnCancel.PerformClick();
	}
}


using System;
using System.Diagnostics;
using System.Runtime.CompilerServices;

[SpecialName]
[DebuggerHidden]
internal void _Lambda$__R2(object a0, EventArgs a1)
{
	_Lambda$__7();
}


using System.Drawing;
using System.Runtime.CompilerServices;

[SpecialName]
internal void _Lambda$__7()
{
	$VB$Local_btnOK.BackColor = Color.FromArgb(16, 137, 230);
}


using System;
using System.Diagnostics;
using System.Runtime.CompilerServices;

[SpecialName]
[DebuggerHidden]
internal void _Lambda$__R3(object a0, EventArgs a1)
{
	_Lambda$__8();
}


using System.Drawing;
using System.Runtime.CompilerServices;

[SpecialName]
internal void _Lambda$__8()
{
	$VB$Local_btnOK.BackColor = Color.FromArgb(0, 120, 215);
}


using System;
using System.Diagnostics;
using System.Runtime.CompilerServices;

[SpecialName]
[DebuggerHidden]
internal void _Lambda$__R4(object a0, EventArgs a1)
{
	_Lambda$__9();
}


using System.Drawing;
using System.Runtime.CompilerServices;

[SpecialName]
internal void _Lambda$__9()
{
	$VB$Local_btnCancel.BackColor = Color.FromArgb(225, 225, 230);
}


using System;
using System.Diagnostics;
using System.Runtime.CompilerServices;

[SpecialName]
[DebuggerHidden]
internal void _Lambda$__R5(object a0, EventArgs a1)
{
	_Lambda$__10();
}


using System.Drawing;
using System.Runtime.CompilerServices;

[SpecialName]
internal void _Lambda$__10()
{
	$VB$Local_btnCancel.BackColor = Color.FromArgb(235, 235, 240);
}
