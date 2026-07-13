Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Drawing
Imports System.Drawing.Drawing2D

<ComVisible(True)>
<Guid("B1E75C9F-AB12-4E38-B5BB-ABCDEF123456")>
<ProgId("RaghavStaadExtractor.SheetSwitch")>
<ClassInterface(ClassInterfaceType.AutoDual)>
Public Class SheetSwitch

    ' === Switch to Sheet2 ===
    Public Function SwitchToSheet2() As Boolean
        Try
            Dim excelApp As Excel.Application = CType(Marshal.GetActiveObject("Excel.Application"), Excel.Application)
            Dim wb As Excel.Workbook = excelApp.ActiveWorkbook
            Dim ws As Excel.Worksheet = CType(wb.Sheets("Sheet2"), Excel.Worksheet)
            ws.Activate()
            Return True
        Catch ex As Exception
            MsgBox("SwitchToSheet2 Error: " & ex.Message)
            Return False
        End Try
    End Function

    ' === Switch to Sheet1 ===
    Public Function SwitchToSheet1() As Boolean
        Try
            Dim excelApp As Excel.Application = CType(Marshal.GetActiveObject("Excel.Application"), Excel.Application)
            Dim wb As Excel.Workbook = excelApp.ActiveWorkbook
            Dim ws As Excel.Worksheet = CType(wb.Sheets("Sheet1"), Excel.Worksheet)
            ws.Activate()
            Return True
        Catch ex As Exception
            MsgBox("SwitchToSheet1 Error: " & ex.Message)
            Return False
        End Try
    End Function

    ' === Normalize Excel to "clean but functional" view ===
    ' === Normalize Excel to "clean but functional" view ===
    ' === Normalize Excel to "clean but functional" view ===
    Public Function NormalizeView() As Boolean
        Try
            ' 🔐 Ask for secure password using modern dialog
            Dim userPass As String = PromptForPassword("Enter owner password to proceed to Normal View:", "Owner Access Required")
            If String.IsNullOrWhiteSpace(userPass) OrElse userPass <> "AVPR" Then
                MsgBox("Hold Up! Only the Boss Gets In Here.", MsgBoxStyle.Critical)
                Return False
            End If

            Dim excelApp As Excel.Application = CType(Marshal.GetActiveObject("Excel.Application"), Excel.Application)
            Dim ws As Excel.Worksheet = CType(excelApp.ActiveSheet, Excel.Worksheet)

            ' 🔓 Unprotect sheet with password
            Try
                ws.Unprotect("2022")
            Catch ex As Exception
                MsgBox("Warning: Could not unprotected sheet. " & ex.Message, MsgBoxStyle.Exclamation)
            End Try

            ' ✅ Restore essential UI
            excelApp.ActiveWindow.DisplayHeadings = True             ' Show row numbers & column letters
            excelApp.ActiveWindow.DisplayWorkbookTabs = True         ' Show sheet tabs at bottom
            excelApp.ActiveWindow.DisplayGridlines = True            ' Show cell gridlines
            excelApp.ActiveWindow.DisplayHorizontalScrollBar = True  ' Show horizontal scrollbar
            excelApp.ActiveWindow.DisplayVerticalScrollBar = True    ' Show vertical scrollbar

            ' ✅ Show ribbon reliably
            excelApp.ExecuteExcel4Macro("SHOW.TOOLBAR(""Ribbon"", True)")
            excelApp.CommandBars("Ribbon").Visible = True

            ' ✅ Show formula bar & exit fullscreen
            excelApp.DisplayFormulaBar = True
            excelApp.DisplayFullScreen = False

            ' ✅ Reset position and zoom
            ws.Activate()
            excelApp.ActiveWindow.ScrollRow = 1
            excelApp.ActiveWindow.ScrollColumn = 1
            excelApp.ActiveWindow.Zoom = 100

            Return True
        Catch ex As Exception
            MsgBox("NormalizeView Error: " & ex.Message, MsgBoxStyle.Critical)
            Return False
        End Try
    End Function

    Private Function PromptForPassword(prompt As String, title As String) As String
        Dim passwordResult As String = Nothing

        ' Create the main form
        Dim form As New Form()

        ' Form properties
        With form
            .Size = New Drawing.Size(450, 280)
            .StartPosition = FormStartPosition.CenterScreen
            .FormBorderStyle = FormBorderStyle.None
            .BackColor = Drawing.Color.FromArgb(200, 200, 200)
            .Font = New Drawing.Font("Segoe UI", 9.0F)
            .ShowInTaskbar = False
            .TopMost = True
            .Text = title
        End With

        ' Header Panel
        Dim panelHeader As New Panel With {
        .Dock = DockStyle.Top,
        .Height = 60,
        .BackColor = Drawing.Color.FromArgb(51, 122, 183)
    }

        ' Title Label
        Dim lblTitle As New Label With {
        .Text = title,
        .ForeColor = Drawing.Color.White,
        .Font = New Drawing.Font("Segoe UI", 14.0F, Drawing.FontStyle.Bold),
        .Location = New Drawing.Point(60, 18),
        .AutoSize = True,
        .BackColor = Drawing.Color.Transparent
    }

        ' Security Icon
        Dim iconPictureBox As New PictureBox With {
        .Size = New Drawing.Size(32, 32),
        .Location = New Drawing.Point(15, 14),
        .BackColor = Drawing.Color.Transparent,
        .SizeMode = PictureBoxSizeMode.CenterImage
    }

        ' Create security icon
        Dim iconBmp As New Bitmap(32, 32)
        Using g As Graphics = Graphics.FromImage(iconBmp)
            g.SmoothingMode = Drawing2D.SmoothingMode.AntiAlias

            ' Draw shield background
            Dim shieldPath As New Drawing2D.GraphicsPath()
            shieldPath.AddLines({
            New Drawing.Point(16, 2),
            New Drawing.Point(26, 8),
            New Drawing.Point(26, 18),
            New Drawing.Point(16, 30),
            New Drawing.Point(6, 18),
            New Drawing.Point(6, 8)
        })
            shieldPath.CloseFigure()

            Using brush As New Drawing2D.LinearGradientBrush(New Drawing.Rectangle(0, 0, 32, 32),
                                                        Drawing.Color.FromArgb(255, 215, 0),
                                                        Drawing.Color.FromArgb(255, 165, 0),
                                                        Drawing2D.LinearGradientMode.Vertical)
                g.FillPath(brush, shieldPath)
            End Using

            ' Draw shield outline
            Using pen As New Drawing.Pen(Drawing.Color.FromArgb(200, 140, 0), 2)
                g.DrawPath(pen, shieldPath)
            End Using

            ' Draw lock symbol
            Using pen As New Drawing.Pen(Drawing.Color.White, 2)
                g.DrawRectangle(pen, 12, 16, 8, 6)
                g.DrawArc(pen, 13, 12, 6, 6, 0, -180)
            End Using
        End Using
        iconPictureBox.Image = iconBmp

        ' Main Panel
        Dim panelMain As New Panel With {
        .Dock = DockStyle.Fill,
        .BackColor = Drawing.Color.White,
        .Padding = New Padding(30, 25, 30, 20)
    }

        ' Prompt Label
        Dim lblPrompt As New Label With {
        .Text = prompt,
        .Font = New Drawing.Font("Segoe UI", 10.0F),
        .ForeColor = Drawing.Color.FromArgb(64, 64, 64),
        .Location = New Drawing.Point(30, 25),
        .AutoSize = True,
        .MaximumSize = New Drawing.Size(360, 0)
    }

        ' Password TextBox
        Dim txtPassword As New TextBox With {
        .UseSystemPasswordChar = True,
        .Font = New Drawing.Font("Segoe UI", 11.0F),
        .Size = New Drawing.Size(360, 30),
        .Location = New Drawing.Point(1, 1),
        .BorderStyle = BorderStyle.None,
        .BackColor = Drawing.Color.FromArgb(250, 250, 250)
    }

        ' Create custom border for textbox
        Dim txtPanel As New Panel With {
        .Size = New Drawing.Size(362, 32),
        .Location = New Drawing.Point(29, 54),
        .BackColor = Drawing.Color.FromArgb(200, 200, 200)
    }
        txtPanel.Controls.Add(txtPassword)

        ' Button Panel
        Dim panelButtons As New Panel With {
        .Dock = DockStyle.Bottom,
        .Height = 70,
        .BackColor = Drawing.Color.FromArgb(248, 248, 248)
    }

        ' OK Button
        Dim btnOK As New Button With {
        .Text = "OK",
        .Size = New Drawing.Size(100, 35),
        .Location = New Drawing.Point(240, 18),
        .BackColor = Drawing.Color.FromArgb(51, 122, 183),
        .ForeColor = Drawing.Color.White,
        .FlatStyle = FlatStyle.Flat,
        .Font = New Drawing.Font("Segoe UI", 9.75F, Drawing.FontStyle.Bold),
        .Cursor = Cursors.Hand
    }
        btnOK.FlatAppearance.BorderSize = 0
        btnOK.FlatAppearance.MouseOverBackColor = Drawing.Color.FromArgb(40, 96, 144)

        ' Cancel Button
        Dim btnCancel As New Button With {
        .Text = "Cancel",
        .Size = New Drawing.Size(100, 35),
        .Location = New Drawing.Point(350, 18),
        .BackColor = Drawing.Color.FromArgb(108, 117, 125),
        .ForeColor = Drawing.Color.White,
        .FlatStyle = FlatStyle.Flat,
        .Font = New Drawing.Font("Segoe UI", 9.75F),
        .Cursor = Cursors.Hand
    }
        btnCancel.FlatAppearance.BorderSize = 0
        btnCancel.FlatAppearance.MouseOverBackColor = Drawing.Color.FromArgb(90, 98, 104)

        ' Add controls to panels
        panelHeader.Controls.AddRange({lblTitle, iconPictureBox})
        panelMain.Controls.AddRange({lblPrompt, txtPanel})
        panelButtons.Controls.AddRange({btnOK, btnCancel})

        ' Add panels to form
        form.Controls.AddRange({panelMain, panelHeader, panelButtons})

        ' Variables for dragging functionality
        Dim isDragging As Boolean = False
        Dim dragStartPoint As Drawing.Point

        ' Event Handlers

        ' Form border painting
        AddHandler form.Paint, Sub(sender, e)
                                   Using pen As New Drawing.Pen(Drawing.Color.FromArgb(120, 120, 120), 2)
                                       e.Graphics.DrawRectangle(pen, 0, 0, form.Width - 1, form.Height - 1)
                                   End Using
                               End Sub

        ' Password textbox key handling
        AddHandler txtPassword.KeyPress, Sub(sender, e)
                                             If e.KeyChar = ChrW(Keys.Enter) Then
                                                 e.Handled = True
                                                 btnOK.PerformClick()
                                             ElseIf e.KeyChar = ChrW(Keys.Escape) Then
                                                 e.Handled = True
                                                 btnCancel.PerformClick()
                                             End If
                                         End Sub

        ' OK button click
        AddHandler btnOK.Click, Sub(sender, e)
                                    passwordResult = txtPassword.Text
                                    form.DialogResult = DialogResult.OK
                                    form.Close()
                                End Sub

        ' Cancel button click
        AddHandler btnCancel.Click, Sub(sender, e)
                                        passwordResult = Nothing
                                        form.DialogResult = DialogResult.Cancel
                                        form.Close()
                                    End Sub

        ' Form load - focus and animation
        AddHandler form.Load, Sub(sender, e)
                                  txtPassword.Focus()
                                  txtPassword.Select()

                                  ' Fade-in animation
                                  form.Opacity = 0
                                  Dim fadeTimer As New Timer With {.Interval = 10}
                                  AddHandler fadeTimer.Tick, Sub(s, args)
                                                                 form.Opacity += 0.05
                                                                 If form.Opacity >= 1.0 Then
                                                                     fadeTimer.Stop()
                                                                     fadeTimer.Dispose()
                                                                 End If
                                                             End Sub
                                  fadeTimer.Start()
                              End Sub

        ' Header dragging functionality
        Dim headerMouseDown As MouseEventHandler = Sub(sender, e)
                                                       If e.Button = MouseButtons.Left Then
                                                           isDragging = True
                                                           dragStartPoint = e.Location
                                                       End If
                                                   End Sub

        Dim headerMouseMove As MouseEventHandler = Sub(sender, e)
                                                       If isDragging Then
                                                           Dim p As Drawing.Point = form.Location
                                                           p.X += e.X - dragStartPoint.X
                                                           p.Y += e.Y - dragStartPoint.Y
                                                           form.Location = p
                                                       End If
                                                   End Sub

        Dim headerMouseUp As MouseEventHandler = Sub(sender, e)
                                                     isDragging = False
                                                 End Sub

        ' Add dragging events to header and title
        AddHandler panelHeader.MouseDown, headerMouseDown
        AddHandler panelHeader.MouseMove, headerMouseMove
        AddHandler panelHeader.MouseUp, headerMouseUp
        AddHandler lblTitle.MouseDown, headerMouseDown
        AddHandler lblTitle.MouseMove, headerMouseMove
        AddHandler lblTitle.MouseUp, headerMouseUp

        ' Show the dialog and return result
        form.ShowDialog()
        form.Dispose()

        Return passwordResult
    End Function

    ' === Open LinkedIn Profile in Default Browser ===
    Public Function OpenMyLinkedInProfile() As Boolean
        Try
            Dim linkedInURL As String = "https://www.linkedin.com/in/peeyushraghav15/"

            ' Use System.Diagnostics.Process to launch default browser
            System.Diagnostics.Process.Start(New ProcessStartInfo(linkedInURL) With {.UseShellExecute = True})

            Return True
        Catch ex As Exception
            MsgBox("OpenMyLinkedInProfile Error: " & ex.Message, MsgBoxStyle.Critical)
            Return False
        End Try
    End Function


End Class
