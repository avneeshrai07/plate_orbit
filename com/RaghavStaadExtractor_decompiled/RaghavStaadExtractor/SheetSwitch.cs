using System;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace RaghavStaadExtractor;

[ComVisible(true)]
[Guid("B1E75C9F-AB12-4E38-B5BB-ABCDEF123456")]
[ProgId("RaghavStaadExtractor.SheetSwitch")]
[ClassInterface(ClassInterfaceType.AutoDual)]
public class SheetSwitch
{
	public bool SwitchToSheet2()
	{
		bool result;
		try
		{
			Microsoft.Office.Interop.Excel.Application application = (Microsoft.Office.Interop.Excel.Application)Marshal.GetActiveObject("Excel.Application");
			Workbook activeWorkbook = application.ActiveWorkbook;
			Worksheet worksheet = (Worksheet)activeWorkbook.Sheets["Sheet2"];
			worksheet.Activate();
			result = true;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			Interaction.MsgBox("SwitchToSheet2 Error: " + ex2.Message);
			result = false;
			ProjectData.ClearProjectError();
		}
		return result;
	}

	public bool SwitchToSheet1()
	{
		bool result;
		try
		{
			Microsoft.Office.Interop.Excel.Application application = (Microsoft.Office.Interop.Excel.Application)Marshal.GetActiveObject("Excel.Application");
			Workbook activeWorkbook = application.ActiveWorkbook;
			Worksheet worksheet = (Worksheet)activeWorkbook.Sheets["Sheet1"];
			worksheet.Activate();
			result = true;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			Interaction.MsgBox("SwitchToSheet1 Error: " + ex2.Message);
			result = false;
			ProjectData.ClearProjectError();
		}
		return result;
	}

	public bool NormalizeView()
	{
		bool result;
		try
		{
			string text = PromptForPassword("Enter owner password to proceed to Normal View:", "Owner Access Required");
			if (string.IsNullOrWhiteSpace(text) || Operators.CompareString(text, "AVPR", TextCompare: false) != 0)
			{
				Interaction.MsgBox("Hold Up! Only the Boss Gets In Here.", MsgBoxStyle.Critical);
				result = false;
			}
			else
			{
				Microsoft.Office.Interop.Excel.Application application = (Microsoft.Office.Interop.Excel.Application)Marshal.GetActiveObject("Excel.Application");
				Worksheet worksheet = (Worksheet)application.ActiveSheet;
				try
				{
					worksheet.Unprotect("2022");
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					Interaction.MsgBox("Warning: Could not unprotected sheet. " + ex2.Message, MsgBoxStyle.Exclamation);
					ProjectData.ClearProjectError();
				}
				application.ActiveWindow.DisplayHeadings = true;
				application.ActiveWindow.DisplayWorkbookTabs = true;
				application.ActiveWindow.DisplayGridlines = true;
				application.ActiveWindow.DisplayHorizontalScrollBar = true;
				application.ActiveWindow.DisplayVerticalScrollBar = true;
				application.ExecuteExcel4Macro("SHOW.TOOLBAR(\"Ribbon\", True)");
				application.CommandBars["Ribbon"].Visible = true;
				application.DisplayFormulaBar = true;
				application.DisplayFullScreen = false;
				worksheet.Activate();
				application.ActiveWindow.ScrollRow = 1;
				application.ActiveWindow.ScrollColumn = 1;
				application.ActiveWindow.Zoom = 100;
				result = true;
			}
		}
		catch (Exception ex3)
		{
			ProjectData.SetProjectError(ex3);
			Exception ex4 = ex3;
			Interaction.MsgBox("NormalizeView Error: " + ex4.Message, MsgBoxStyle.Critical);
			result = false;
			ProjectData.ClearProjectError();
		}
		return result;
	}

	private string PromptForPassword(string prompt, string title)
	{
		string result = null;
		Form form = new Form();
		Form form2 = form;
		form2.Size = new Size(450, 280);
		form2.StartPosition = FormStartPosition.CenterScreen;
		form2.FormBorderStyle = FormBorderStyle.None;
		form2.BackColor = Color.FromArgb(200, 200, 200);
		form2.Font = new System.Drawing.Font("Segoe UI", 9f);
		form2.ShowInTaskbar = false;
		form2.TopMost = true;
		form2.Text = title;
		form2 = null;
		Panel panel = new Panel
		{
			Dock = DockStyle.Top,
			Height = 60,
			BackColor = Color.FromArgb(51, 122, 183)
		};
		Label label = new Label
		{
			Text = title,
			ForeColor = Color.White,
			Font = new System.Drawing.Font("Segoe UI", 14f, FontStyle.Bold),
			Location = new Point(60, 18),
			AutoSize = true,
			BackColor = Color.Transparent
		};
		PictureBox pictureBox = new PictureBox
		{
			Size = new Size(32, 32),
			Location = new Point(15, 14),
			BackColor = Color.Transparent,
			SizeMode = PictureBoxSizeMode.CenterImage
		};
		Bitmap image = new Bitmap(32, 32);
		using (Graphics graphics = Graphics.FromImage(image))
		{
			graphics.SmoothingMode = SmoothingMode.AntiAlias;
			GraphicsPath graphicsPath = new GraphicsPath();
			graphicsPath.AddLines(new Point[6]
			{
				new Point(16, 2),
				new Point(26, 8),
				new Point(26, 18),
				new Point(16, 30),
				new Point(6, 18),
				new Point(6, 8)
			});
			graphicsPath.CloseFigure();
			using (LinearGradientBrush brush = new LinearGradientBrush(new Rectangle(0, 0, 32, 32), Color.FromArgb(255, 215, 0), Color.FromArgb(255, 165, 0), LinearGradientMode.Vertical))
			{
				graphics.FillPath(brush, graphicsPath);
			}
			using (Pen pen = new Pen(Color.FromArgb(200, 140, 0), 2f))
			{
				graphics.DrawPath(pen, graphicsPath);
			}
			using Pen pen2 = new Pen(Color.White, 2f);
			graphics.DrawRectangle(pen2, 12, 16, 8, 6);
			graphics.DrawArc(pen2, 13, 12, 6, 6, 0, -180);
		}
		pictureBox.Image = image;
		Panel panel2 = new Panel
		{
			Dock = DockStyle.Fill,
			BackColor = Color.White,
			Padding = new Padding(30, 25, 30, 20)
		};
		Label label2 = new Label
		{
			Text = prompt,
			Font = new System.Drawing.Font("Segoe UI", 10f),
			ForeColor = Color.FromArgb(64, 64, 64),
			Location = new Point(30, 25),
			AutoSize = true,
			MaximumSize = new Size(360, 0)
		};
		TextBox textBox = new TextBox
		{
			UseSystemPasswordChar = true,
			Font = new System.Drawing.Font("Segoe UI", 11f),
			Size = new Size(360, 30),
			Location = new Point(1, 1),
			BorderStyle = BorderStyle.None,
			BackColor = Color.FromArgb(250, 250, 250)
		};
		Panel panel3 = new Panel
		{
			Size = new Size(362, 32),
			Location = new Point(29, 54),
			BackColor = Color.FromArgb(200, 200, 200)
		};
		panel3.Controls.Add(textBox);
		Panel panel4 = new Panel
		{
			Dock = DockStyle.Bottom,
			Height = 70,
			BackColor = Color.FromArgb(248, 248, 248)
		};
		Button button = new Button
		{
			Text = "OK",
			Size = new Size(100, 35),
			Location = new Point(240, 18),
			BackColor = Color.FromArgb(51, 122, 183),
			ForeColor = Color.White,
			FlatStyle = FlatStyle.Flat,
			Font = new System.Drawing.Font("Segoe UI", 9.75f, FontStyle.Bold),
			Cursor = Cursors.Hand
		};
		button.FlatAppearance.BorderSize = 0;
		button.FlatAppearance.MouseOverBackColor = Color.FromArgb(40, 96, 144);
		Button button2 = new Button
		{
			Text = "Cancel",
			Size = new Size(100, 35),
			Location = new Point(350, 18),
			BackColor = Color.FromArgb(108, 117, 125),
			ForeColor = Color.White,
			FlatStyle = FlatStyle.Flat,
			Font = new System.Drawing.Font("Segoe UI", 9.75f),
			Cursor = Cursors.Hand
		};
		button2.FlatAppearance.BorderSize = 0;
		button2.FlatAppearance.MouseOverBackColor = Color.FromArgb(90, 98, 104);
		panel.Controls.AddRange(new Control[2] { label, pictureBox });
		panel2.Controls.AddRange(new Control[2] { label2, panel3 });
		panel4.Controls.AddRange(new Control[2] { button, button2 });
		form.Controls.AddRange(new Control[3] { panel2, panel, panel4 });
		bool flag = false;
		checked
		{
			form.Paint += [SpecialName] (object sender, PaintEventArgs e) =>
			{
				using Pen pen3 = new Pen(Color.FromArgb(120, 120, 120), 2f);
				e.Graphics.DrawRectangle(pen3, 0, 0, form.Width - 1, form.Height - 1);
			};
			textBox.KeyPress += [SpecialName] (object sender, KeyPressEventArgs e) =>
			{
				if (e.KeyChar == '\r')
				{
					e.Handled = true;
					button.PerformClick();
				}
				else if (e.KeyChar == '\u001b')
				{
					e.Handled = true;
					button2.PerformClick();
				}
			};
			button.Click += [SpecialName] (object sender, EventArgs e) =>
			{
				result = textBox.Text;
				form.DialogResult = DialogResult.OK;
				form.Close();
			};
			button2.Click += [SpecialName] (object sender, EventArgs e) =>
			{
				result = null;
				form.DialogResult = DialogResult.Cancel;
				form.Close();
			};
			form.Load += [SpecialName] (object sender, EventArgs e) =>
			{
				textBox.Focus();
				textBox.Select();
				form.Opacity = 0.0;
				Timer timer = new Timer
				{
					Interval = 10
				};
				timer.Tick += [SpecialName] (object s, EventArgs args) =>
				{
					form.Opacity += 0.05;
					if (form.Opacity >= 1.0)
					{
						timer.Stop();
						timer.Dispose();
					}
				};
				timer.Start();
			};
			Point location2 = default(Point);
			MouseEventHandler value = [SpecialName] (object sender, MouseEventArgs e) =>
			{
				if (e.Button == MouseButtons.Left)
				{
					flag = true;
					location2 = e.Location;
				}
			};
			MouseEventHandler value2 = [SpecialName] (object sender, MouseEventArgs e) =>
			{
				if (flag)
				{
					Point location = form.Location;
					location.X += e.X - location2.X;
					location.Y += e.Y - location2.Y;
					form.Location = location;
				}
			};
			MouseEventHandler value3 = [SpecialName] (object sender, MouseEventArgs e) =>
			{
				flag = false;
			};
			panel.MouseDown += value;
			panel.MouseMove += value2;
			panel.MouseUp += value3;
			label.MouseDown += value;
			label.MouseMove += value2;
			label.MouseUp += value3;
			form.ShowDialog();
			form.Dispose();
			return result;
		}
	}

	public bool OpenMyLinkedInProfile()
	{
		bool result;
		try
		{
			string fileName = "https://www.linkedin.com/in/peeyushraghav15/";
			Process.Start(new ProcessStartInfo(fileName)
			{
				UseShellExecute = true
			});
			result = true;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			Interaction.MsgBox("OpenMyLinkedInProfile Error: " + ex2.Message, MsgBoxStyle.Critical);
			result = false;
			ProjectData.ClearProjectError();
		}
		return result;
	}
}
