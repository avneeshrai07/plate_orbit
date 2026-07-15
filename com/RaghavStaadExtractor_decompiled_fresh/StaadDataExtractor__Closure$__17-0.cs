// ===== RaghavStaadExtractor.StaadDataExtractor._Closure$__17-0 =====

// Decompiled member-by-member from RaghavStaadExtractor.dll



public string $VB$Local_logFilePath;


public _Closure$__17-0()
{
}


using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.CompilerServices;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

[SpecialName]
internal void _Lambda$__0(object sender, EventArgs e)
{
	try
	{
		Process.Start("explorer.exe", "/select,\"" + $VB$Local_logFilePath + "\"");
	}
	catch (Exception ex)
	{
		ProjectData.SetProjectError(ex);
		Exception ex2 = ex;
		try
		{
			Process.Start("explorer.exe", Path.GetDirectoryName($VB$Local_logFilePath));
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			Interaction.MsgBox("Could not open log file location.", MsgBoxStyle.Exclamation);
			ProjectData.ClearProjectError();
		}
		ProjectData.ClearProjectError();
	}
}
