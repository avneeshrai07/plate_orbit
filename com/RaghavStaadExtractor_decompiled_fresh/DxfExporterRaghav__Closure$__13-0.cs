// ===== RaghavStaadExtractor.DxfExporterRaghav._Closure$__13-0 =====

// Decompiled member-by-member from RaghavStaadExtractor.dll



public string $VB$Local_filePath;


public string $VB$Local_folderPath;


using System.Windows.Forms;

public Form $VB$Local_customDialog;


public _Closure$__13-0()
{
}


using System;
using System.Diagnostics;
using System.Runtime.CompilerServices;
using System.Windows.Forms;
using Microsoft.VisualBasic.CompilerServices;

[SpecialName]
internal void _Lambda$__0(object sender, EventArgs e)
{
	try
	{
		Process.Start("explorer.exe", $"/select,\"{$VB$Local_filePath}\"");
	}
	catch (Exception ex)
	{
		ProjectData.SetProjectError(ex);
		Exception ex2 = ex;
		try
		{
			Process.Start("explorer.exe", $VB$Local_folderPath);
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			MessageBox.Show("Could not open file location.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
			ProjectData.ClearProjectError();
		}
		ProjectData.ClearProjectError();
	}
	$VB$Local_customDialog.Close();
}


using System;
using System.Diagnostics;
using System.Runtime.CompilerServices;
using System.Windows.Forms;
using Microsoft.VisualBasic.CompilerServices;

[SpecialName]
internal void _Lambda$__1(object sender, EventArgs e)
{
	try
	{
		Process.Start($VB$Local_filePath);
	}
	catch (Exception ex)
	{
		ProjectData.SetProjectError(ex);
		Exception ex2 = ex;
		try
		{
			Process.Start(new ProcessStartInfo
			{
				FileName = $VB$Local_filePath,
				UseShellExecute = true,
				Verb = "open"
			});
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			MessageBox.Show("Could not open DXF file. Please ensure you have a CAD application installed.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
			ProjectData.ClearProjectError();
		}
		ProjectData.ClearProjectError();
	}
	$VB$Local_customDialog.Close();
}
