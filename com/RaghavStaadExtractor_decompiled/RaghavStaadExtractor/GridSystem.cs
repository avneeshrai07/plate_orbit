using System;
using System.Diagnostics;
using System.Drawing;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace RaghavStaadExtractor;

[ComVisible(true)]
[Guid("F1234567-89AB-4CDE-BCDE-ABCDEF012399")]
[ProgId("RaghavStaadExtractor.GridSystem")]
public class GridSystem
{
	private class GridInputForm : Form
	{
		[CompilerGenerated]
		internal sealed class _Closure_0024__31_002D0
		{
			public Button _0024VB_0024Local_btnClose;

			public _Closure_0024__31_002D0(_Closure_0024__31_002D0 arg0)
			{
				if (arg0 != null)
				{
					_0024VB_0024Local_btnClose = arg0._0024VB_0024Local_btnClose;
				}
			}

			[SpecialName]
			[DebuggerHidden]
			internal void _Lambda_0024__R2(object a0, EventArgs a1)
			{
				_Lambda_0024__1();
			}

			[SpecialName]
			internal void _Lambda_0024__1()
			{
				_0024VB_0024Local_btnClose.BackColor = ErrorRed;
			}

			[SpecialName]
			[DebuggerHidden]
			internal void _Lambda_0024__R3(object a0, EventArgs a1)
			{
				_Lambda_0024__2();
			}

			[SpecialName]
			internal void _Lambda_0024__2()
			{
				_0024VB_0024Local_btnClose.BackColor = PrimaryDark;
			}
		}

		[CompilerGenerated]
		internal sealed class _Closure_0024__32_002D0
		{
			public Button _0024VB_0024Local_btn;

			public Color _0024VB_0024Local_backColor;

			[SpecialName]
			[DebuggerHidden]
			internal void _Lambda_0024__R12(object a0, EventArgs a1)
			{
				_Lambda_0024__0();
			}

			[SpecialName]
			internal void _Lambda_0024__0()
			{
				_0024VB_0024Local_btn.BackColor = checked(Color.FromArgb(Math.Max(0, _0024VB_0024Local_backColor.R - 20), Math.Max(0, _0024VB_0024Local_backColor.G - 20), Math.Max(0, _0024VB_0024Local_backColor.B - 20)));
			}

			[SpecialName]
			[DebuggerHidden]
			internal void _Lambda_0024__R13(object a0, EventArgs a1)
			{
				_Lambda_0024__1();
			}

			[SpecialName]
			internal void _Lambda_0024__1()
			{
				_0024VB_0024Local_btn.BackColor = _0024VB_0024Local_backColor;
			}
		}

		private static readonly Color PrimaryDark = Color.FromArgb(44, 62, 80);

		private static readonly Color PrimaryMedium = Color.FromArgb(52, 73, 94);

		private static readonly Color AccentBlue = Color.FromArgb(41, 128, 185);

		private static readonly Color AccentGreen = Color.FromArgb(39, 174, 96);

		private static readonly Color AccentOrange = Color.FromArgb(230, 126, 34);

		private static readonly Color NeutralLight = Color.FromArgb(236, 240, 241);

		private static readonly Color NeutralMedium = Color.FromArgb(189, 195, 199);

		private static readonly Color TextDark = Color.FromArgb(44, 62, 80);

		private static readonly Color TextLight = Color.White;

		private static readonly Color SuccessGreen = Color.FromArgb(46, 204, 113);

		private static readonly Color ErrorRed = Color.FromArgb(231, 76, 60);

		private static readonly Color FocusBlue = Color.FromArgb(52, 152, 219);

		private static readonly Color FlipPurple = Color.FromArgb(155, 89, 182);

		private TextBox txtB2;

		private TextBox txtB3;

		private TextBox txtB4;

		private TextBox txtC2;

		private TextBox txtC3;

		private TextBox txtC4;

		private TextBox txtD2;

		private TextBox txtD3;

		private TextBox txtD4;

		private TextBox txtE2;

		private Button btnOk;

		private Button btnCancel;

		private Button btnFlipX;

		private Button btnFlipZ;

		private Button btnAuto;

		private const int WM_NCLBUTTONDOWN = 161;

		private const int HTCAPTION = 2;

		public string[] Values => new string[10] { txtB2.Text, txtB3.Text, txtB4.Text, txtC2.Text, txtC3.Text, txtC4.Text, txtD2.Text, txtD3.Text, txtD4.Text, txtE2.Text };

		public GridInputForm(string[] initialValues)
		{
			_Closure_0024__31_002D0 arg = default(_Closure_0024__31_002D0);
			_Closure_0024__31_002D0 CS_0024_003C_003E8__locals8 = new _Closure_0024__31_002D0(arg);
			base._002Ector();
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
				Font = new System.Drawing.Font("Segoe UI", 11f, FontStyle.Bold),
				ForeColor = TextLight,
				AutoSize = false,
				Dock = DockStyle.Fill,
				TextAlign = ContentAlignment.MiddleCenter
			};
			panel.Controls.Add(label);
			CS_0024_003C_003E8__locals8._0024VB_0024Local_btnClose = new Button
			{
				Text = "✖",
				ForeColor = TextLight,
				Font = new System.Drawing.Font("Segoe UI", 11f, FontStyle.Bold),
				FlatStyle = FlatStyle.Flat,
				Size = new Size(50, 50),
				Dock = DockStyle.Right,
				BackColor = PrimaryDark
			};
			CS_0024_003C_003E8__locals8._0024VB_0024Local_btnClose.FlatAppearance.BorderSize = 0;
			CS_0024_003C_003E8__locals8._0024VB_0024Local_btnClose.Click += [SpecialName] [DebuggerHidden] (object a0, EventArgs a1) =>
			{
				_Lambda_0024__31_002D0();
			};
			CS_0024_003C_003E8__locals8._0024VB_0024Local_btnClose.MouseEnter += [SpecialName] [DebuggerHidden] (object a0, EventArgs a1) =>
			{
				CS_0024_003C_003E8__locals8._Lambda_0024__1();
			};
			CS_0024_003C_003E8__locals8._0024VB_0024Local_btnClose.MouseLeave += [SpecialName] [DebuggerHidden] (object a0, EventArgs a1) =>
			{
				CS_0024_003C_003E8__locals8._Lambda_0024__2();
			};
			panel.Controls.Add(CS_0024_003C_003E8__locals8._0024VB_0024Local_btnClose);
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
				Font = new System.Drawing.Font("Segoe UI", 10f, FontStyle.Bold),
				ForeColor = TextLight,
				TextAlign = ContentAlignment.MiddleCenter
			};
			Label value2 = new Label
			{
				Text = "GRID CO-ORDINATES",
				Dock = DockStyle.Fill,
				Font = new System.Drawing.Font("Segoe UI", 10f, FontStyle.Bold),
				ForeColor = TextLight,
				TextAlign = ContentAlignment.MiddleCenter
			};
			Label value3 = new Label
			{
				Text = "LINE EXTENSIONS",
				Dock = DockStyle.Fill,
				Font = new System.Drawing.Font("Segoe UI", 10f, FontStyle.Bold),
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
						Font = new System.Drawing.Font("Segoe UI", 9f, FontStyle.Bold),
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
					array2[num4].Font = new System.Drawing.Font("Segoe UI", 9f);
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
				btnFlipX.Font = new System.Drawing.Font("Segoe UI", 9f, FontStyle.Bold);
				btnFlipX.BackColor = Color.FromArgb(22, 160, 133);
				btnFlipX.ForeColor = TextLight;
				btnFlipX.FlatStyle = FlatStyle.Flat;
				btnFlipX.FlatAppearance.BorderSize = 0;
				btnFlipX.Cursor = Cursors.Hand;
				btnFlipX.MouseEnter += [SpecialName] [DebuggerHidden] (object a0, EventArgs a1) =>
				{
					_Lambda_0024__31_002D3();
				};
				btnFlipX.MouseLeave += [SpecialName] [DebuggerHidden] (object a0, EventArgs a1) =>
				{
					_Lambda_0024__31_002D4();
				};
				btnFlipX.Click += FlipXValues;
				base.Controls.Add(btnFlipX);
				btnFlipZ.Text = "FLIP Z";
				btnFlipZ.Size = new Size(110, 38);
				btnFlipZ.Location = new Point(num7 + btnFlipX.Width + 10, num6);
				btnFlipZ.Font = new System.Drawing.Font("Segoe UI", 9f, FontStyle.Bold);
				btnFlipZ.BackColor = Color.FromArgb(142, 68, 173);
				btnFlipZ.ForeColor = TextLight;
				btnFlipZ.FlatStyle = FlatStyle.Flat;
				btnFlipZ.FlatAppearance.BorderSize = 0;
				btnFlipZ.Cursor = Cursors.Hand;
				btnFlipZ.MouseEnter += [SpecialName] [DebuggerHidden] (object a0, EventArgs a1) =>
				{
					_Lambda_0024__31_002D5();
				};
				btnFlipZ.MouseLeave += [SpecialName] [DebuggerHidden] (object a0, EventArgs a1) =>
				{
					_Lambda_0024__31_002D6();
				};
				btnFlipZ.Click += FlipZValues;
				base.Controls.Add(btnFlipZ);
				btnAuto.Text = "AUTO";
				btnAuto.Size = new Size(110, 38);
				btnAuto.Location = new Point(base.ClientSize.Width - btnAuto.Width - 20, num6);
				btnAuto.Font = new System.Drawing.Font("Segoe UI", 9f, FontStyle.Bold);
				btnAuto.BackColor = Color.FromArgb(52, 152, 219);
				btnAuto.ForeColor = TextLight;
				btnAuto.FlatStyle = FlatStyle.Flat;
				btnAuto.FlatAppearance.BorderSize = 0;
				btnAuto.Cursor = Cursors.Hand;
				btnAuto.MouseEnter += [SpecialName] [DebuggerHidden] (object a0, EventArgs a1) =>
				{
					_Lambda_0024__31_002D7();
				};
				btnAuto.MouseLeave += [SpecialName] [DebuggerHidden] (object a0, EventArgs a1) =>
				{
					_Lambda_0024__31_002D8();
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
						Font = new System.Drawing.Font("Segoe UI", 9f, FontStyle.Bold),
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
					array4[num8].Font = new System.Drawing.Font("Segoe UI", 9f);
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
						Font = new System.Drawing.Font("Segoe UI", 9f, FontStyle.Bold),
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
					array6[num10].Font = new System.Drawing.Font("Segoe UI", 9f);
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
					Font = new System.Drawing.Font("Segoe UI", 9f, FontStyle.Bold),
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
				txtE2.Font = new System.Drawing.Font("Segoe UI", 9f);
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
				btnOk.Click += [SpecialName] [DebuggerHidden] (object a0, EventArgs a1) =>
				{
					_Lambda_0024__31_002D9();
				};
				btnCancel.Text = "✗ CANCEL";
				StyleButton(btnCancel, ErrorRed);
				btnCancel.Location = new Point(400, 350);
				btnCancel.Click += [SpecialName] [DebuggerHidden] (object a0, EventArgs a1) =>
				{
					_Lambda_0024__31_002D10();
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

		private void StyleButton(Button btn, Color backColor)
		{
			_Closure_0024__32_002D0 CS_0024_003C_003E8__locals14 = new _Closure_0024__32_002D0();
			CS_0024_003C_003E8__locals14._0024VB_0024Local_btn = btn;
			CS_0024_003C_003E8__locals14._0024VB_0024Local_backColor = backColor;
			CS_0024_003C_003E8__locals14._0024VB_0024Local_btn.Font = new System.Drawing.Font("Segoe UI", 9f, FontStyle.Bold);
			CS_0024_003C_003E8__locals14._0024VB_0024Local_btn.BackColor = CS_0024_003C_003E8__locals14._0024VB_0024Local_backColor;
			CS_0024_003C_003E8__locals14._0024VB_0024Local_btn.ForeColor = TextLight;
			CS_0024_003C_003E8__locals14._0024VB_0024Local_btn.FlatStyle = FlatStyle.Flat;
			CS_0024_003C_003E8__locals14._0024VB_0024Local_btn.FlatAppearance.BorderSize = 0;
			CS_0024_003C_003E8__locals14._0024VB_0024Local_btn.Size = new Size(110, 38);
			CS_0024_003C_003E8__locals14._0024VB_0024Local_btn.Cursor = Cursors.Hand;
			CS_0024_003C_003E8__locals14._0024VB_0024Local_btn.MouseEnter += [SpecialName] [DebuggerHidden] (object a0, EventArgs a1) =>
			{
				CS_0024_003C_003E8__locals14._Lambda_0024__0();
			};
			CS_0024_003C_003E8__locals14._0024VB_0024Local_btn.MouseLeave += [SpecialName] [DebuggerHidden] (object a0, EventArgs a1) =>
			{
				CS_0024_003C_003E8__locals14._Lambda_0024__1();
			};
		}

		[DllImport("user32.dll")]
		private static extern bool ReleaseCapture();

		[DllImport("user32.dll")]
		private static extern int SendMessage(IntPtr hWnd, int wMsg, int wParam, int lParam);

		private void TitleBar_MouseDown(object sender, MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Left)
			{
				ReleaseCapture();
				SendMessage(base.Handle, 161, 2, 0);
			}
		}

		private void TextBox_KeyDown(object sender, KeyEventArgs e)
		{
			if (e.Control && e.KeyCode == Keys.A)
			{
				((TextBox)sender).SelectAll();
				e.Handled = true;
				e.SuppressKeyPress = true;
			}
		}

		private void TextBox_GotFocus(object sender, EventArgs e)
		{
			TextBox textBox = (TextBox)sender;
			Panel panel = (Panel)textBox.Parent;
			panel.BackColor = FocusBlue;
		}

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
				timer.Tick += [SpecialName] (object s, EventArgs args) =>
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
				timer.Tick += [SpecialName] (object s, EventArgs args) =>
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
				timer.Tick += [SpecialName] (object s, EventArgs args) =>
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

		private string FlipTextValues(string inputText)
		{
			string result;
			try
			{
				string[] array = inputText.Trim().Split(new char[1] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
				Array.Reverse(array);
				result = string.Join(" ", array);
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				result = inputText;
				ProjectData.ClearProjectError();
			}
			return result;
		}

		[SpecialName]
		[CompilerGenerated]
		private void _Lambda_0024__31_002D0()
		{
			base.DialogResult = DialogResult.Cancel;
		}

		[SpecialName]
		[CompilerGenerated]
		private void _Lambda_0024__31_002D3()
		{
			btnFlipX.BackColor = Color.FromArgb(Math.Max(0, 2), Math.Max(0, 140), Math.Max(0, 113));
		}

		[SpecialName]
		[CompilerGenerated]
		private void _Lambda_0024__31_002D4()
		{
			btnFlipX.BackColor = Color.FromArgb(22, 160, 133);
		}

		[SpecialName]
		[CompilerGenerated]
		private void _Lambda_0024__31_002D5()
		{
			btnFlipZ.BackColor = Color.FromArgb(Math.Max(0, 122), Math.Max(0, 48), Math.Max(0, 153));
		}

		[SpecialName]
		[CompilerGenerated]
		private void _Lambda_0024__31_002D6()
		{
			btnFlipZ.BackColor = Color.FromArgb(142, 68, 173);
		}

		[SpecialName]
		[CompilerGenerated]
		private void _Lambda_0024__31_002D7()
		{
			btnAuto.BackColor = Color.FromArgb(Math.Max(0, 32), Math.Max(0, 132), Math.Max(0, 199));
		}

		[SpecialName]
		[CompilerGenerated]
		private void _Lambda_0024__31_002D8()
		{
			btnAuto.BackColor = Color.FromArgb(52, 152, 219);
		}

		[SpecialName]
		[CompilerGenerated]
		private void _Lambda_0024__31_002D9()
		{
			base.DialogResult = DialogResult.OK;
		}

		[SpecialName]
		[CompilerGenerated]
		private void _Lambda_0024__31_002D10()
		{
			base.DialogResult = DialogResult.Cancel;
		}
	}

	public void ShowGridInputForm()
	{
		try
		{
			string[] initialValues = ReadValuesFromExcel();
			using GridInputForm gridInputForm = new GridInputForm(initialValues);
			if (gridInputForm.ShowDialog() == DialogResult.OK)
			{
				WriteValuesToExcel(gridInputForm.Values);
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			MessageBox.Show("Error: " + ex2.Message);
			ProjectData.ClearProjectError();
		}
	}

	private string[] ReadValuesFromExcel()
	{
		Microsoft.Office.Interop.Excel.Application application = (Microsoft.Office.Interop.Excel.Application)Marshal.GetActiveObject("Excel.Application");
		Workbook activeWorkbook = application.ActiveWorkbook;
		Worksheet worksheet = (Worksheet)activeWorkbook.Sheets["Sheet4"];
		return new string[10]
		{
			IfNull(RuntimeHelpers.GetObjectValue(((_Worksheet)worksheet).get_Range((object)"B2", RuntimeHelpers.GetObjectValue(Missing.Value)).get_Value(RuntimeHelpers.GetObjectValue(Missing.Value)))),
			IfNull(RuntimeHelpers.GetObjectValue(((_Worksheet)worksheet).get_Range((object)"B3", RuntimeHelpers.GetObjectValue(Missing.Value)).get_Value(RuntimeHelpers.GetObjectValue(Missing.Value)))),
			IfNull(RuntimeHelpers.GetObjectValue(((_Worksheet)worksheet).get_Range((object)"B4", RuntimeHelpers.GetObjectValue(Missing.Value)).get_Value(RuntimeHelpers.GetObjectValue(Missing.Value)))),
			IfNull(RuntimeHelpers.GetObjectValue(((_Worksheet)worksheet).get_Range((object)"C2", RuntimeHelpers.GetObjectValue(Missing.Value)).get_Value(RuntimeHelpers.GetObjectValue(Missing.Value)))),
			IfNull(RuntimeHelpers.GetObjectValue(((_Worksheet)worksheet).get_Range((object)"C3", RuntimeHelpers.GetObjectValue(Missing.Value)).get_Value(RuntimeHelpers.GetObjectValue(Missing.Value)))),
			IfNull(RuntimeHelpers.GetObjectValue(((_Worksheet)worksheet).get_Range((object)"C4", RuntimeHelpers.GetObjectValue(Missing.Value)).get_Value(RuntimeHelpers.GetObjectValue(Missing.Value)))),
			IfNull(RuntimeHelpers.GetObjectValue(((_Worksheet)worksheet).get_Range((object)"D2", RuntimeHelpers.GetObjectValue(Missing.Value)).get_Value(RuntimeHelpers.GetObjectValue(Missing.Value)))),
			IfNull(RuntimeHelpers.GetObjectValue(((_Worksheet)worksheet).get_Range((object)"D3", RuntimeHelpers.GetObjectValue(Missing.Value)).get_Value(RuntimeHelpers.GetObjectValue(Missing.Value)))),
			IfNull(RuntimeHelpers.GetObjectValue(((_Worksheet)worksheet).get_Range((object)"D4", RuntimeHelpers.GetObjectValue(Missing.Value)).get_Value(RuntimeHelpers.GetObjectValue(Missing.Value)))),
			IfNull(RuntimeHelpers.GetObjectValue(((_Worksheet)worksheet).get_Range((object)"E2", RuntimeHelpers.GetObjectValue(Missing.Value)).get_Value(RuntimeHelpers.GetObjectValue(Missing.Value))))
		};
	}

	private string IfNull(object val)
	{
		return (val == null) ? "" : val.ToString();
	}

	private void WriteValuesToExcel(string[] values)
	{
		try
		{
			Microsoft.Office.Interop.Excel.Application application = (Microsoft.Office.Interop.Excel.Application)Marshal.GetActiveObject("Excel.Application");
			Workbook activeWorkbook = application.ActiveWorkbook;
			Worksheet worksheet = (Worksheet)activeWorkbook.Sheets["Sheet4"];
			((_Worksheet)worksheet).get_Range((object)"B2", RuntimeHelpers.GetObjectValue(Missing.Value)).set_Value(RuntimeHelpers.GetObjectValue(Missing.Value), (object)values[0]);
			((_Worksheet)worksheet).get_Range((object)"B3", RuntimeHelpers.GetObjectValue(Missing.Value)).set_Value(RuntimeHelpers.GetObjectValue(Missing.Value), (object)values[1]);
			((_Worksheet)worksheet).get_Range((object)"B4", RuntimeHelpers.GetObjectValue(Missing.Value)).set_Value(RuntimeHelpers.GetObjectValue(Missing.Value), (object)values[2]);
			((_Worksheet)worksheet).get_Range((object)"C2", RuntimeHelpers.GetObjectValue(Missing.Value)).set_Value(RuntimeHelpers.GetObjectValue(Missing.Value), (object)values[3]);
			((_Worksheet)worksheet).get_Range((object)"C3", RuntimeHelpers.GetObjectValue(Missing.Value)).set_Value(RuntimeHelpers.GetObjectValue(Missing.Value), (object)values[4]);
			((_Worksheet)worksheet).get_Range((object)"C4", RuntimeHelpers.GetObjectValue(Missing.Value)).set_Value(RuntimeHelpers.GetObjectValue(Missing.Value), (object)values[5]);
			((_Worksheet)worksheet).get_Range((object)"D2", RuntimeHelpers.GetObjectValue(Missing.Value)).set_Value(RuntimeHelpers.GetObjectValue(Missing.Value), (object)values[6]);
			((_Worksheet)worksheet).get_Range((object)"D3", RuntimeHelpers.GetObjectValue(Missing.Value)).set_Value(RuntimeHelpers.GetObjectValue(Missing.Value), (object)values[7]);
			((_Worksheet)worksheet).get_Range((object)"D4", RuntimeHelpers.GetObjectValue(Missing.Value)).set_Value(RuntimeHelpers.GetObjectValue(Missing.Value), (object)values[8]);
			((_Worksheet)worksheet).get_Range((object)"E2", RuntimeHelpers.GetObjectValue(Missing.Value)).set_Value(RuntimeHelpers.GetObjectValue(Missing.Value), (object)values[9]);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			MessageBox.Show("Error writing values to Excel: " + ex2.Message);
			ProjectData.ClearProjectError();
		}
	}
}
