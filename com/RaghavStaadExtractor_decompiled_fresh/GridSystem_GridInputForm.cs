// ===== RaghavStaadExtractor.GridSystem.GridInputForm =====

// Decompiled member-by-member from RaghavStaadExtractor.dll



using System.Drawing;

private static readonly Color PrimaryDark;


using System.Drawing;

private static readonly Color PrimaryMedium;


using System.Drawing;

private static readonly Color AccentBlue;


using System.Drawing;

private static readonly Color AccentGreen;


using System.Drawing;

private static readonly Color AccentOrange;


using System.Drawing;

private static readonly Color NeutralLight;


using System.Drawing;

private static readonly Color NeutralMedium;


using System.Drawing;

private static readonly Color TextDark;


using System.Drawing;

private static readonly Color TextLight;


using System.Drawing;

private static readonly Color SuccessGreen;


using System.Drawing;

private static readonly Color ErrorRed;


using System.Drawing;

private static readonly Color FocusBlue;


using System.Drawing;

private static readonly Color FlipPurple;


using System.Windows.Forms;

private TextBox txtB2;


using System.Windows.Forms;

private TextBox txtB3;


using System.Windows.Forms;

private TextBox txtB4;


using System.Windows.Forms;

private TextBox txtC2;


using System.Windows.Forms;

private TextBox txtC3;


using System.Windows.Forms;

private TextBox txtC4;


using System.Windows.Forms;

private TextBox txtD2;


using System.Windows.Forms;

private TextBox txtD3;


using System.Windows.Forms;

private TextBox txtD4;


using System.Windows.Forms;

private TextBox txtE2;


using System.Windows.Forms;

private Button btnOk;


using System.Windows.Forms;

private Button btnCancel;


using System.Windows.Forms;

private Button btnFlipX;


using System.Windows.Forms;

private Button btnFlipZ;


using System.Windows.Forms;

private Button btnAuto;


private const int WM_NCLBUTTONDOWN = 161;


private const int HTCAPTION = 2;


public string[] Values => new string[10] { txtB2.Text, txtB3.Text, txtB4.Text, txtC2.Text, txtC3.Text, txtC4.Text, txtD2.Text, txtD3.Text, txtD4.Text, txtE2.Text };


using System.Drawing;

static GridInputForm()
{
	PrimaryDark = Color.FromArgb(44, 62, 80);
	PrimaryMedium = Color.FromArgb(52, 73, 94);
	AccentBlue = Color.FromArgb(41, 128, 185);
	AccentGreen = Color.FromArgb(39, 174, 96);
	AccentOrange = Color.FromArgb(230, 126, 34);
	NeutralLight = Color.FromArgb(236, 240, 241);
	NeutralMedium = Color.FromArgb(189, 195, 199);
	TextDark = Color.FromArgb(44, 62, 80);
	TextLight = Color.White;
	SuccessGreen = Color.FromArgb(46, 204, 113);
	ErrorRed = Color.FromArgb(231, 76, 60);
	FocusBlue = Color.FromArgb(52, 152, 219);
	FlipPurple = Color.FromArgb(155, 89, 182);
}


using System.Drawing;
using System.Windows.Forms;
using Microsoft.VisualBasic.CompilerServices;

public GridInputForm(string[] initialValues)
{
	_Closure$__31-0 arg = default(_Closure$__31-0);
	_Closure$__31-0 CS$<>8__locals0 = new _Closure$__31-0(arg);
	base..ctor();
	txtB2 = new TextBox();
	txtB3 = new TextBox();
	txtB4 = new TextBox();
	txtC2 = new TextBox();
	txtC3 = new TextBox();
	txtC4 = new TextBox();
	txtD2 = new TextBox();
	txtD3 = new TextBox();
	txtD4 = new TextBox();
	txtE2 = new TextBox();
	btnOk = new Button();
	btnCancel = new Button();
	btnFlipX = new Button();
	btnFlipZ = new Button();
	btnAuto = new Button();
	if (initialValues == null || initialValues.Length < 10)
	{
		initialValues = (string[])Utils.CopyArray(initialValues, new string[10]);
		int num = 0;
		do
		{
			if (initialValues[num] == null)
			{
				initialValues[num] = "";
			}
			num = checked(num + 1);
		}
		while (num <= 9);
	}
	base.FormBorderStyle = FormBorderStyle.None;
	base.StartPosition = FormStartPosition.Manual;
	base.Location = new Point(checked(Screen.PrimaryScreen.WorkingArea.Width - 760) / 2, checked(Screen.PrimaryScreen.WorkingArea.Height - 400) / 2);
	base.ClientSize = new Size(760, 400);
	BackColor = NeutralLight;
	Panel panel = new Panel
	{
		Height = 50,
		Dock = DockStyle.Top,
		BackColor = PrimaryDark
	};
	base.Controls.Add(panel);
	Label label = new Label
	{
		Text = "DEFINE GRIDS AS PER STAAD'S AXIS SYSTEM",
		Font = new Font("Segoe UI", 11f, FontStyle.Bold),
		ForeColor = TextLight,
		AutoSize = false,
		Dock = DockStyle.Fill,
		TextAlign = ContentAlignment.MiddleCenter
	};
	panel.Controls.Add(label);
	CS$<>8__locals0.$VB$Local_btnClose = new Button
	{
		Text = "✖",
		ForeColor = TextLight,
		Font = new Font("Segoe UI", 11f, FontStyle.Bold),
		FlatStyle = FlatStyle.Flat,
		Size = new Size(50, 50),
		Dock = DockStyle.Right,
		BackColor = PrimaryDark
	};
	CS$<>8__locals0.$VB$Local_btnClose.FlatAppearance.BorderSize = 0;
	CS$<>8__locals0.$VB$Local_btnClose.Click += delegate
	{
		_Lambda$__31-0();
	};
	CS$<>8__locals0.$VB$Local_btnClose.MouseEnter += delegate
	{
		CS$<>8__locals0._Lambda$__1();
	};
	CS$<>8__locals0.$VB$Local_btnClose.MouseLeave += delegate
	{
		CS$<>8__locals0._Lambda$__2();
	};
	panel.Controls.Add(CS$<>8__locals0.$VB$Local_btnClose);
	label.MouseDown += TitleBar_MouseDown;
	Panel panel2 = new Panel
	{
		Location = new Point(20, 70),
		Size = new Size(720, 260),
		BackColor = Color.White,
		BorderStyle = BorderStyle.None
	};
	base.Controls.Add(panel2);
	Panel panel3 = new Panel
	{
		Location = new Point(22, 72),
		Size = new Size(720, 260),
		BackColor = Color.FromArgb(50, 0, 0, 0)
	};
	base.Controls.Add(panel3);
	panel3.SendToBack();
	int num2 = 60;
	int num3 = 40;
	Panel panel4 = new Panel
	{
		BackColor = NeutralMedium,
		Width = 1,
		Height = panel2.Height,
		Location = new Point(200, 0)
	};
	Panel panel5 = new Panel
	{
		BackColor = NeutralMedium,
		Width = 1,
		Height = panel2.Height,
		Location = new Point(520, 0)
	};
	Panel panel6 = new Panel
	{
		BackColor = NeutralMedium,
		Height = 1,
		Width = panel2.Width,
		Location = new Point(0, 50)
	};
	panel2.Controls.AddRange(new Control[3] { panel4, panel5, panel6 });
	Panel panel7 = new Panel
	{
		BackColor = PrimaryMedium,
		Size = new Size(200, 50),
		Location = new Point(0, 0)
	};
	Panel panel8 = new Panel
	{
		BackColor = AccentBlue,
		Size = new Size(320, 50),
		Location = new Point(200, 0)
	};
	Panel panel9 = new Panel
	{
		BackColor = AccentOrange,
		Size = new Size(200, 50),
		Location = new Point(520, 0)
	};
	Label value = new Label
	{
		Text = "GRID NAMING",
		Dock = DockStyle.Fill,
		Font = new Font("Segoe UI", 10f, FontStyle.Bold),
		ForeColor = TextLight,
		TextAlign = ContentAlignment.MiddleCenter
	};
	Label value2 = new Label
	{
		Text = "GRID CO-ORDINATES",
		Dock = DockStyle.Fill,
		Font = new Font("Segoe UI", 10f, FontStyle.Bold),
		ForeColor = TextLight,
		TextAlign = ContentAlignment.MiddleCenter
	};
	Label value3 = new Label
	{
		Text = "LINE EXTENSIONS",
		Dock = DockStyle.Fill,
		Font = new Font("Segoe UI", 10f, FontStyle.Bold),
		ForeColor = TextLight,
		TextAlign = ContentAlignment.MiddleCenter
	};
	panel7.Controls.Add(value);
	panel8.Controls.Add(value2);
	panel9.Controls.Add(value3);
	panel2.Controls.AddRange(new Control[3] { panel7, panel8, panel9 });
	string[] array = new string[3] { "X:", "Y:", "Z:" };
	TextBox[] array2 = new TextBox[3] { txtB2, txtB3, txtB4 };
	int num4 = 0;
	checked
	{
		do
		{
			int num5 = num2 + num4 * num3;
			Label value4 = new Label
			{
				Text = array[num4],
				Location = new Point(15, num5 + 8),
				Size = new Size(25, 20),
				Font = new Font("Segoe UI", 9f, FontStyle.Bold),
				ForeColor = PrimaryMedium
			};
			panel2.Controls.Add(value4);
			Panel panel10 = new Panel
			{
				Location = new Point(40, num5),
				Size = new Size(144, 32),
				BackColor = NeutralLight
			};
			array2[num4].Location = new Point(2, 2);
			array2[num4].Size = new Size(140, 28);
			array2[num4].BorderStyle = BorderStyle.None;
			array2[num4].Font = new Font("Segoe UI", 9f);
			array2[num4].BackColor = Color.White;
			array2[num4].Text = initialValues[num4];
			panel10.Controls.Add(array2[num4]);
			panel2.Controls.Add(panel10);
			if (!string.IsNullOrWhiteSpace(array2[num4].Text))
			{
				panel10.BackColor = Color.FromArgb(230, 247, 255);
			}
			array2[num4].GotFocus += TextBox_GotFocus;
			array2[num4].LostFocus += TextBox_LostFocus;
			num4++;
		}
		while (num4 <= 2);
		int num6 = 350;
		int num7 = 20;
		btnFlipX.Text = "FLIP X";
		btnFlipX.Size = new Size(110, 38);
		btnFlipX.Location = new Point(num7, num6);
		btnFlipX.Font = new Font("Segoe UI", 9f, FontStyle.Bold);
		btnFlipX.BackColor = Color.FromArgb(22, 160, 133);
		btnFlipX.ForeColor = TextLight;
		btnFlipX.FlatStyle = FlatStyle.Flat;
		btnFlipX.FlatAppearance.BorderSize = 0;
		btnFlipX.Cursor = Cursors.Hand;
		btnFlipX.MouseEnter += delegate
		{
			_Lambda$__31-3();
		};
		btnFlipX.MouseLeave += delegate
		{
			_Lambda$__31-4();
		};
		btnFlipX.Click += FlipXValues;
		base.Controls.Add(btnFlipX);
		btnFlipZ.Text = "FLIP Z";
		btnFlipZ.Size = new Size(110, 38);
		btnFlipZ.Location = new Point(num7 + btnFlipX.Width + 10, num6);
		btnFlipZ.Font = new Font("Segoe UI", 9f, FontStyle.Bold);
		btnFlipZ.BackColor = Color.FromArgb(142, 68, 173);
		btnFlipZ.ForeColor = TextLight;
		btnFlipZ.FlatStyle = FlatStyle.Flat;
		btnFlipZ.FlatAppearance.BorderSize = 0;
		btnFlipZ.Cursor = Cursors.Hand;
		btnFlipZ.MouseEnter += delegate
		{
			_Lambda$__31-5();
		};
		btnFlipZ.MouseLeave += delegate
		{
			_Lambda$__31-6();
		};
		btnFlipZ.Click += FlipZValues;
		base.Controls.Add(btnFlipZ);
		btnAuto.Text = "AUTO";
		btnAuto.Size = new Size(110, 38);
		btnAuto.Location = new Point(base.ClientSize.Width - btnAuto.Width - 20, num6);
		btnAuto.Font = new Font("Segoe UI", 9f, FontStyle.Bold);
		btnAuto.BackColor = Color.FromArgb(52, 152, 219);
		btnAuto.ForeColor = TextLight;
		btnAuto.FlatStyle = FlatStyle.Flat;
		btnAuto.FlatAppearance.BorderSize = 0;
		btnAuto.Cursor = Cursors.Hand;
		btnAuto.MouseEnter += delegate
		{
			_Lambda$__31-7();
		};
		btnAuto.MouseLeave += delegate
		{
			_Lambda$__31-8();
		};
		btnAuto.Click += AutoOperation;
		base.Controls.Add(btnAuto);
		string[] array3 = new string[3] { "X-COORDS:", "Y-COORDS:", "Z-COORDS:" };
		TextBox[] array4 = new TextBox[3] { txtC2, txtC3, txtC4 };
		int num8 = 0;
		do
		{
			int num9 = num2 + num8 * num3;
			Label value5 = new Label
			{
				Text = array3[num8],
				Location = new Point(210, num9 + 8),
				Size = new Size(90, 20),
				Font = new Font("Segoe UI", 9f, FontStyle.Bold),
				ForeColor = AccentBlue
			};
			panel2.Controls.Add(value5);
			Panel panel11 = new Panel
			{
				Location = new Point(310, num9),
				Size = new Size(184, 32),
				BackColor = NeutralLight
			};
			array4[num8].Location = new Point(2, 2);
			array4[num8].Size = new Size(180, 28);
			array4[num8].BorderStyle = BorderStyle.None;
			array4[num8].Font = new Font("Segoe UI", 9f);
			array4[num8].BackColor = Color.White;
			array4[num8].Text = initialValues[num8 + 3];
			panel11.Controls.Add(array4[num8]);
			panel2.Controls.Add(panel11);
			if (!string.IsNullOrWhiteSpace(array4[num8].Text))
			{
				panel11.BackColor = Color.FromArgb(230, 247, 255);
			}
			array4[num8].GotFocus += TextBox_GotFocus;
			array4[num8].LostFocus += TextBox_LostFocus;
			num8++;
		}
		while (num8 <= 2);
		string[] array5 = new string[3] { "EXT-X:", "EXT-Y:", "EXT-Z:" };
		TextBox[] array6 = new TextBox[3] { txtD2, txtD3, txtD4 };
		int num10 = 0;
		do
		{
			int num11 = num2 + num10 * num3;
			Label value6 = new Label
			{
				Text = array5[num10],
				Location = new Point(540, num11 + 8),
				Size = new Size(70, 20),
				Font = new Font("Segoe UI", 9f, FontStyle.Bold),
				ForeColor = AccentOrange
			};
			panel2.Controls.Add(value6);
			Panel panel12 = new Panel
			{
				Location = new Point(620, num11),
				Size = new Size(104, 32),
				BackColor = NeutralLight
			};
			array6[num10].Location = new Point(2, 2);
			array6[num10].Size = new Size(100, 28);
			array6[num10].BorderStyle = BorderStyle.None;
			array6[num10].Font = new Font("Segoe UI", 9f);
			array6[num10].BackColor = Color.White;
			array6[num10].Text = initialValues[num10 + 6];
			panel12.Controls.Add(array6[num10]);
			panel2.Controls.Add(panel12);
			if (!string.IsNullOrWhiteSpace(array6[num10].Text))
			{
				panel12.BackColor = Color.FromArgb(230, 247, 255);
			}
			array6[num10].GotFocus += TextBox_GotFocus;
			array6[num10].LostFocus += TextBox_LostFocus;
			num10++;
		}
		while (num10 <= 2);
		int num12 = num2 + 3 * num3;
		Label value7 = new Label
		{
			Text = "HEIGHT:",
			Location = new Point(540, num12 + 8),
			Size = new Size(70, 20),
			Font = new Font("Segoe UI", 9f, FontStyle.Bold),
			ForeColor = AccentGreen
		};
		panel2.Controls.Add(value7);
		Panel panel13 = new Panel
		{
			Location = new Point(620, num12),
			Size = new Size(104, 32),
			BackColor = NeutralLight
		};
		txtE2.Location = new Point(2, 2);
		txtE2.Size = new Size(100, 28);
		txtE2.BorderStyle = BorderStyle.None;
		txtE2.Font = new Font("Segoe UI", 9f);
		txtE2.BackColor = Color.White;
		txtE2.Text = initialValues[9];
		panel13.Controls.Add(txtE2);
		panel2.Controls.Add(panel13);
		if (!string.IsNullOrWhiteSpace(txtE2.Text))
		{
			panel13.BackColor = Color.FromArgb(230, 247, 255);
		}
		txtE2.GotFocus += TextBox_GotFocus;
		txtE2.LostFocus += TextBox_LostFocus;
		btnOk.Text = "✓ SAVE";
		StyleButton(btnOk, SuccessGreen);
		btnOk.Location = new Point(280, 350);
		btnOk.Click += delegate
		{
			_Lambda$__31-9();
		};
		btnCancel.Text = "✗ CANCEL";
		StyleButton(btnCancel, ErrorRed);
		btnCancel.Location = new Point(400, 350);
		btnCancel.Click += delegate
		{
			_Lambda$__31-10();
		};
		base.Controls.AddRange(new Control[2] { btnOk, btnCancel });
		txtB2.KeyDown += TextBox_KeyDown;
		txtB3.KeyDown += TextBox_KeyDown;
		txtB4.KeyDown += TextBox_KeyDown;
		txtC2.KeyDown += TextBox_KeyDown;
		txtC3.KeyDown += TextBox_KeyDown;
		txtC4.KeyDown += TextBox_KeyDown;
		txtD2.KeyDown += TextBox_KeyDown;
		txtD3.KeyDown += TextBox_KeyDown;
		txtD4.KeyDown += TextBox_KeyDown;
		txtE2.KeyDown += TextBox_KeyDown;
	}
}


using System.Drawing;
using System.Windows.Forms;

private void StyleButton(Button btn, Color backColor)
{
	_Closure$__32-0 CS$<>8__locals0 = new _Closure$__32-0();
	CS$<>8__locals0.$VB$Local_btn = btn;
	CS$<>8__locals0.$VB$Local_backColor = backColor;
	CS$<>8__locals0.$VB$Local_btn.Font = new Font("Segoe UI", 9f, FontStyle.Bold);
	CS$<>8__locals0.$VB$Local_btn.BackColor = CS$<>8__locals0.$VB$Local_backColor;
	CS$<>8__locals0.$VB$Local_btn.ForeColor = TextLight;
	CS$<>8__locals0.$VB$Local_btn.FlatStyle = FlatStyle.Flat;
	CS$<>8__locals0.$VB$Local_btn.FlatAppearance.BorderSize = 0;
	CS$<>8__locals0.$VB$Local_btn.Size = new Size(110, 38);
	CS$<>8__locals0.$VB$Local_btn.Cursor = Cursors.Hand;
	CS$<>8__locals0.$VB$Local_btn.MouseEnter += delegate
	{
		CS$<>8__locals0._Lambda$__0();
	};
	CS$<>8__locals0.$VB$Local_btn.MouseLeave += delegate
	{
		CS$<>8__locals0._Lambda$__1();
	};
}


using System.Runtime.InteropServices;

[DllImport("user32.dll")]
private static extern bool ReleaseCapture();


using System;
using System.Runtime.InteropServices;

[DllImport("user32.dll")]
private static extern int SendMessage(IntPtr hWnd, int wMsg, int wParam, int lParam);


using System.Windows.Forms;

private void TitleBar_MouseDown(object sender, MouseEventArgs e)
{
	if (e.Button == MouseButtons.Left)
	{
		ReleaseCapture();
		SendMessage(base.Handle, 161, 2, 0);
	}
}


using System.Windows.Forms;

private void TextBox_KeyDown(object sender, KeyEventArgs e)
{
	if (e.Control && e.KeyCode == Keys.A)
	{
		((TextBox)sender).SelectAll();
		e.Handled = true;
		e.SuppressKeyPress = true;
	}
}


using System;
using System.Windows.Forms;

private void TextBox_GotFocus(object sender, EventArgs e)
{
	TextBox textBox = (TextBox)sender;
	Panel panel = (Panel)textBox.Parent;
	panel.BackColor = FocusBlue;
}


using System;
using System.Drawing;
using System.Windows.Forms;

private void TextBox_LostFocus(object sender, EventArgs e)
{
	TextBox textBox = (TextBox)sender;
	Panel panel = (Panel)textBox.Parent;
	if (string.IsNullOrWhiteSpace(textBox.Text))
	{
		panel.BackColor = NeutralLight;
	}
	else
	{
		panel.BackColor = Color.FromArgb(212, 237, 218);
	}
}


using System;
using System.Drawing;
using System.Windows.Forms;
using Microsoft.VisualBasic.CompilerServices;

private void FlipXValues(object sender, EventArgs e)
{
	try
	{
		Control control = null;
		if (base.ActiveControl != null)
		{
			control = base.ActiveControl;
		}
		if (string.IsNullOrWhiteSpace(txtB2.Text))
		{
			return;
		}
		txtB2.Text = FlipTextValues(txtB2.Text);
		Color backColor = txtB2.Parent.BackColor;
		txtB2.Parent.BackColor = Color.FromArgb(255, 230, 255);
		Timer timer = new Timer();
		timer.Interval = 200;
		timer.Tick += delegate
		{
			txtB2.Parent.BackColor = backColor;
			timer.Stop();
			timer.Dispose();
			if (control != null)
			{
				control.Focus();
			}
		};
		timer.Start();
	}
	catch (Exception ex)
	{
		ProjectData.SetProjectError(ex);
		Exception ex2 = ex;
		MessageBox.Show("Error flipping X values: " + ex2.Message);
		ProjectData.ClearProjectError();
	}
}


using System;
using System.Drawing;
using System.Windows.Forms;
using Microsoft.VisualBasic.CompilerServices;

private void FlipZValues(object sender, EventArgs e)
{
	try
	{
		Control control = null;
		if (base.ActiveControl != null)
		{
			control = base.ActiveControl;
		}
		if (string.IsNullOrWhiteSpace(txtB4.Text))
		{
			return;
		}
		txtB4.Text = FlipTextValues(txtB4.Text);
		Color backColor = txtB4.Parent.BackColor;
		txtB4.Parent.BackColor = Color.FromArgb(255, 230, 255);
		Timer timer = new Timer();
		timer.Interval = 200;
		timer.Tick += delegate
		{
			txtB4.Parent.BackColor = backColor;
			timer.Stop();
			timer.Dispose();
			if (control != null)
			{
				control.Focus();
			}
		};
		timer.Start();
	}
	catch (Exception ex)
	{
		ProjectData.SetProjectError(ex);
		Exception ex2 = ex;
		MessageBox.Show("Error flipping Z values: " + ex2.Message);
		ProjectData.ClearProjectError();
	}
}


using System;
using System.Drawing;
using System.Windows.Forms;
using Microsoft.VisualBasic.CompilerServices;

private void AutoOperation(object sender, EventArgs e)
{
	try
	{
		string text = "A B C D E F G J K L M N P Q R S T V W X Y Z AA AB AC AD AE AF AG AJ AK AL AM AN AP AQ";
		string text2 = "1 2 3 4 5 6 7 8 9 10 11 12 13 14 15 16 17 18 19 20 21 22 23 24 25 26 27 28 29 30 31 32 33 34 35 36 37 38 39 40 41 42 43 44 45 46 47 48 49 50 51 52 53 54 55 56 57 58 59 60 61 62 63 64 65 66 67 68 69 70 71 72 73 74 75 76 77 78 79 80";
		Control control = null;
		if (base.ActiveControl != null)
		{
			control = base.ActiveControl;
		}
		if (string.IsNullOrWhiteSpace(txtB2.Text) && string.IsNullOrWhiteSpace(txtB4.Text))
		{
			txtB2.Text = text;
			txtB4.Text = text2;
		}
		else
		{
			string text3 = txtB2.Text;
			txtB2.Text = txtB4.Text;
			txtB4.Text = text3;
		}
		Color backColor = txtB2.Parent.BackColor;
		Color backColor2 = txtB4.Parent.BackColor;
		txtB2.Parent.BackColor = Color.FromArgb(230, 255, 230);
		txtB4.Parent.BackColor = Color.FromArgb(230, 255, 230);
		Timer timer = new Timer();
		timer.Interval = 200;
		timer.Tick += delegate
		{
			txtB2.Parent.BackColor = (string.IsNullOrWhiteSpace(txtB2.Text) ? NeutralLight : Color.FromArgb(212, 237, 218));
			txtB4.Parent.BackColor = (string.IsNullOrWhiteSpace(txtB4.Text) ? NeutralLight : Color.FromArgb(212, 237, 218));
			timer.Stop();
			timer.Dispose();
			if (control != null)
			{
				control.Focus();
			}
		};
		timer.Start();
	}
	catch (Exception ex)
	{
		ProjectData.SetProjectError(ex);
		Exception ex2 = ex;
		MessageBox.Show("Error in AUTO operation: " + ex2.Message);
		ProjectData.ClearProjectError();
	}
}


using System;
using Microsoft.VisualBasic.CompilerServices;

private string FlipTextValues(string inputText)
{
	try
	{
		string[] array = inputText.Trim().Split(new char[1] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
		Array.Reverse(array);
		return string.Join(" ", array);
	}
	catch (Exception projectError)
	{
		ProjectData.SetProjectError(projectError);
		ProjectData.ClearProjectError();
		return inputText;
	}
}


using System;
using System.Diagnostics;
using System.Runtime.CompilerServices;

[SpecialName]
[CompilerGenerated]
[DebuggerHidden]
private void _Lambda$__R31-1(object a0, EventArgs a1)
{
	_Lambda$__31-0();
}


using System.Runtime.CompilerServices;
using System.Windows.Forms;

[SpecialName]
[CompilerGenerated]
private void _Lambda$__31-0()
{
	base.DialogResult = DialogResult.Cancel;
}


using System;
using System.Diagnostics;
using System.Runtime.CompilerServices;

[SpecialName]
[CompilerGenerated]
[DebuggerHidden]
private void _Lambda$__R31-4(object a0, EventArgs a1)
{
	_Lambda$__31-3();
}


using System;
using System.Drawing;
using System.Runtime.CompilerServices;

[SpecialName]
[CompilerGenerated]
private void _Lambda$__31-3()
{
	btnFlipX.BackColor = Color.FromArgb(Math.Max(0, 2), Math.Max(0, 140), Math.Max(0, 113));
}


using System;
using System.Diagnostics;
using System.Runtime.CompilerServices;

[SpecialName]
[CompilerGenerated]
[DebuggerHidden]
private void _Lambda$__R31-5(object a0, EventArgs a1)
{
	_Lambda$__31-4();
}


using System.Drawing;
using System.Runtime.CompilerServices;

[SpecialName]
[CompilerGenerated]
private void _Lambda$__31-4()
{
	btnFlipX.BackColor = Color.FromArgb(22, 160, 133);
}


using System;
using System.Diagnostics;
using System.Runtime.CompilerServices;

[SpecialName]
[CompilerGenerated]
[DebuggerHidden]
private void _Lambda$__R31-6(object a0, EventArgs a1)
{
	_Lambda$__31-5();
}


using System;
using System.Drawing;
using System.Runtime.CompilerServices;

[SpecialName]
[CompilerGenerated]
private void _Lambda$__31-5()
{
	btnFlipZ.BackColor = Color.FromArgb(Math.Max(0, 122), Math.Max(0, 48), Math.Max(0, 153));
}


using System;
using System.Diagnostics;
using System.Runtime.CompilerServices;

[SpecialName]
[CompilerGenerated]
[DebuggerHidden]
private void _Lambda$__R31-7(object a0, EventArgs a1)
{
	_Lambda$__31-6();
}


using System.Drawing;
using System.Runtime.CompilerServices;

[SpecialName]
[CompilerGenerated]
private void _Lambda$__31-6()
{
	btnFlipZ.BackColor = Color.FromArgb(142, 68, 173);
}


using System;
using System.Diagnostics;
using System.Runtime.CompilerServices;

[SpecialName]
[CompilerGenerated]
[DebuggerHidden]
private void _Lambda$__R31-8(object a0, EventArgs a1)
{
	_Lambda$__31-7();
}


using System;
using System.Drawing;
using System.Runtime.CompilerServices;

[SpecialName]
[CompilerGenerated]
private void _Lambda$__31-7()
{
	btnAuto.BackColor = Color.FromArgb(Math.Max(0, 32), Math.Max(0, 132), Math.Max(0, 199));
}


using System;
using System.Diagnostics;
using System.Runtime.CompilerServices;

[SpecialName]
[CompilerGenerated]
[DebuggerHidden]
private void _Lambda$__R31-9(object a0, EventArgs a1)
{
	_Lambda$__31-8();
}


using System.Drawing;
using System.Runtime.CompilerServices;

[SpecialName]
[CompilerGenerated]
private void _Lambda$__31-8()
{
	btnAuto.BackColor = Color.FromArgb(52, 152, 219);
}


using System;
using System.Diagnostics;
using System.Runtime.CompilerServices;

[SpecialName]
[CompilerGenerated]
[DebuggerHidden]
private void _Lambda$__R31-10(object a0, EventArgs a1)
{
	_Lambda$__31-9();
}


using System.Runtime.CompilerServices;
using System.Windows.Forms;

[SpecialName]
[CompilerGenerated]
private void _Lambda$__31-9()
{
	base.DialogResult = DialogResult.OK;
}


using System;
using System.Diagnostics;
using System.Runtime.CompilerServices;

[SpecialName]
[CompilerGenerated]
[DebuggerHidden]
private void _Lambda$__R31-11(object a0, EventArgs a1)
{
	_Lambda$__31-10();
}


using System.Runtime.CompilerServices;
using System.Windows.Forms;

[SpecialName]
[CompilerGenerated]
private void _Lambda$__31-10()
{
	base.DialogResult = DialogResult.Cancel;
}
