// ===== RaghavStaadExtractor.DxfExporterRaghav._Closure$__12-1 =====

// Decompiled member-by-member from RaghavStaadExtractor.dll



public bool $VB$Local_isDragging;


using System.Drawing;

public Point $VB$Local_lastCursor;


public _Closure$__12-0 $VB$NonLocal_$VB$Closure_2;


public _Closure$__12-1()
{
}


using System.Runtime.CompilerServices;
using System.Windows.Forms;

[SpecialName]
internal void _Lambda$__11(object sender, MouseEventArgs e)
{
	$VB$Local_isDragging = true;
	$VB$Local_lastCursor = Cursor.Position;
}


using System.Drawing;
using System.Runtime.CompilerServices;
using System.Windows.Forms;

[SpecialName]
internal void _Lambda$__12(object sender, MouseEventArgs e)
{
	if ($VB$Local_isDragging)
	{
		Point position = Cursor.Position;
		$VB$NonLocal_$VB$Closure_2.$VB$Local_inputForm.Location = checked(new Point($VB$NonLocal_$VB$Closure_2.$VB$Local_inputForm.Location.X + (position.X - $VB$Local_lastCursor.X), $VB$NonLocal_$VB$Closure_2.$VB$Local_inputForm.Location.Y + (position.Y - $VB$Local_lastCursor.Y)));
		$VB$Local_lastCursor = position;
	}
}


using System.Runtime.CompilerServices;
using System.Windows.Forms;

[SpecialName]
internal void _Lambda$__13(object sender, MouseEventArgs e)
{
	$VB$Local_isDragging = false;
}


using System.Runtime.CompilerServices;
using System.Windows.Forms;

[SpecialName]
internal void _Lambda$__14(object sender, MouseEventArgs e)
{
	$VB$Local_isDragging = true;
	$VB$Local_lastCursor = Cursor.Position;
}


using System.Drawing;
using System.Runtime.CompilerServices;
using System.Windows.Forms;

[SpecialName]
internal void _Lambda$__15(object sender, MouseEventArgs e)
{
	if ($VB$Local_isDragging)
	{
		Point position = Cursor.Position;
		$VB$NonLocal_$VB$Closure_2.$VB$Local_inputForm.Location = checked(new Point($VB$NonLocal_$VB$Closure_2.$VB$Local_inputForm.Location.X + (position.X - $VB$Local_lastCursor.X), $VB$NonLocal_$VB$Closure_2.$VB$Local_inputForm.Location.Y + (position.Y - $VB$Local_lastCursor.Y)));
		$VB$Local_lastCursor = position;
	}
}


using System.Runtime.CompilerServices;
using System.Windows.Forms;

[SpecialName]
internal void _Lambda$__16(object sender, MouseEventArgs e)
{
	$VB$Local_isDragging = false;
}
