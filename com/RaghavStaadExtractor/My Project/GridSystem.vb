Imports System.Drawing
Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports Microsoft.Office.Interop

<ComVisible(True)>
<Guid("F1234567-89AB-4CDE-BCDE-ABCDEF012399")>
<ProgId("RaghavStaadExtractor.GridSystem")>
Public Class GridSystem

    Public Sub ShowGridInputForm()
        Try
            Dim currentValues As String() = ReadValuesFromExcel()
            Using frm As New GridInputForm(currentValues)
                If frm.ShowDialog() = DialogResult.OK Then
                    WriteValuesToExcel(frm.Values)
                End If
            End Using
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message)
        End Try
    End Sub

    ' ========================
    ' Read values from Excel
    ' ========================
    Private Function ReadValuesFromExcel() As String()
        Dim excelApp As Excel.Application = CType(Marshal.GetActiveObject("Excel.Application"), Excel.Application)
        Dim wb As Excel.Workbook = excelApp.ActiveWorkbook
        Dim ws As Excel.Worksheet = CType(wb.Sheets("Sheet4"), Excel.Worksheet)

        Return {
            IfNull(ws.Range("B2").Value),
            IfNull(ws.Range("B3").Value),
            IfNull(ws.Range("B4").Value),
            IfNull(ws.Range("C2").Value),
            IfNull(ws.Range("C3").Value),
            IfNull(ws.Range("C4").Value),
            IfNull(ws.Range("D2").Value),
            IfNull(ws.Range("D3").Value),
            IfNull(ws.Range("D4").Value),
            IfNull(ws.Range("E2").Value)   ' NEW ENTRY
        }
    End Function

    Private Function IfNull(val As Object) As String
        Return If(val Is Nothing, "", val.ToString())
    End Function

    ' ========================
    ' Write values to Excel
    ' ========================
    Private Sub WriteValuesToExcel(values As String())
        Try
            Dim excelApp As Excel.Application = CType(Marshal.GetActiveObject("Excel.Application"), Excel.Application)
            Dim wb As Excel.Workbook = excelApp.ActiveWorkbook
            Dim ws As Excel.Worksheet = CType(wb.Sheets("Sheet4"), Excel.Worksheet)

            ws.Range("B2").Value = values(0)
            ws.Range("B3").Value = values(1)
            ws.Range("B4").Value = values(2)
            ws.Range("C2").Value = values(3)
            ws.Range("C3").Value = values(4)
            ws.Range("C4").Value = values(5)
            ws.Range("D2").Value = values(6)
            ws.Range("D3").Value = values(7)
            ws.Range("D4").Value = values(8)
            ws.Range("E2").Value = values(9)   ' NEW ENTRY
        Catch ex As Exception
            MessageBox.Show("Error writing values to Excel: " & ex.Message)
        End Try
    End Sub

    Private Class GridInputForm
        Inherits Form

        ' === Professional Color Palette ===
        Private Shared ReadOnly PrimaryDark As Color = Color.FromArgb(44, 62, 80)        ' Dark Blue-Gray
        Private Shared ReadOnly PrimaryMedium As Color = Color.FromArgb(52, 73, 94)     ' Medium Blue-Gray
        Private Shared ReadOnly AccentBlue As Color = Color.FromArgb(41, 128, 185)      ' Professional Blue
        Private Shared ReadOnly AccentGreen As Color = Color.FromArgb(39, 174, 96)      ' Professional Green
        Private Shared ReadOnly AccentOrange As Color = Color.FromArgb(230, 126, 34)    ' Professional Orange
        Private Shared ReadOnly NeutralLight As Color = Color.FromArgb(236, 240, 241)   ' Light Gray
        Private Shared ReadOnly NeutralMedium As Color = Color.FromArgb(189, 195, 199)  ' Medium Gray
        Private Shared ReadOnly TextDark As Color = Color.FromArgb(44, 62, 80)          ' Dark Text
        Private Shared ReadOnly TextLight As Color = Color.White                        ' Light Text
        Private Shared ReadOnly SuccessGreen As Color = Color.FromArgb(46, 204, 113)    ' Success Green
        Private Shared ReadOnly ErrorRed As Color = Color.FromArgb(231, 76, 60)         ' Error Red
        Private Shared ReadOnly FocusBlue As Color = Color.FromArgb(52, 152, 219)       ' Focus Blue
        Private Shared ReadOnly FlipPurple As Color = Color.FromArgb(155, 89, 182)      ' Purple for FLIP buttons

        ' === Controls ===
        Private txtB2 As New TextBox()
        Private txtB3 As New TextBox()
        Private txtB4 As New TextBox()
        Private txtC2 As New TextBox()
        Private txtC3 As New TextBox()
        Private txtC4 As New TextBox()
        Private txtD2 As New TextBox()
        Private txtD3 As New TextBox()
        Private txtD4 As New TextBox()
        Private txtE2 As New TextBox()
        Private btnOk As New Button()
        Private btnCancel As New Button()
        Private btnFlipX As New Button()  ' NEW FLIP X BUTTON
        Private btnFlipZ As New Button()  ' NEW FLIP Z BUTTON
        Private btnAuto As New Button()   ' NEW AUTO BUTTON

        ' ===== Public Property =====
        Public ReadOnly Property Values As String()
            Get
                Return {
                txtB2.Text, txtB3.Text, txtB4.Text,
                txtC2.Text, txtC3.Text, txtC4.Text,
                txtD2.Text, txtD3.Text, txtD4.Text,
                txtE2.Text
            }
            End Get
        End Property

        ' ===== Constructor =====
        Public Sub New(initialValues As String())
            ' Validate and initialize values
            If initialValues Is Nothing OrElse initialValues.Length < 10 Then
                ReDim Preserve initialValues(9)
                For i = 0 To 9
                    If initialValues(i) Is Nothing Then initialValues(i) = ""
                Next
            End If

            ' Form styling
            Me.FormBorderStyle = FormBorderStyle.None
            Me.StartPosition = FormStartPosition.Manual
            Me.Location = New Drawing.Point((Screen.PrimaryScreen.WorkingArea.Width - 760) \ 2, (Screen.PrimaryScreen.WorkingArea.Height - 400) \ 2)
            Me.ClientSize = New Drawing.Size(760, 400)
            Me.BackColor = NeutralLight

            ' ==== Professional Title Bar ====
            Dim titleBar As New Panel() With {
                .Height = 50,
                .Dock = DockStyle.Top,
                .BackColor = PrimaryDark
            }
            Me.Controls.Add(titleBar)

            Dim lblTitle As New Label() With {
                .Text = "DEFINE GRIDS AS PER STAAD'S AXIS SYSTEM",
                .Font = New Drawing.Font("Segoe UI", 11, Drawing.FontStyle.Bold),
                .ForeColor = TextLight,
                .AutoSize = False,
                .Dock = DockStyle.Fill,
                .TextAlign = Drawing.ContentAlignment.MiddleCenter
            }
            titleBar.Controls.Add(lblTitle)

            ' Close button with hover effect
            Dim btnClose As New Button() With {
                .Text = "✖",
                .ForeColor = TextLight,
                .Font = New Drawing.Font("Segoe UI", 11, Drawing.FontStyle.Bold),
                .FlatStyle = FlatStyle.Flat,
                .Size = New Drawing.Size(50, 50),
                .Dock = DockStyle.Right,
                .BackColor = PrimaryDark
            }
            btnClose.FlatAppearance.BorderSize = 0
            AddHandler btnClose.Click, Sub() Me.DialogResult = DialogResult.Cancel
            AddHandler btnClose.MouseEnter, Sub() btnClose.BackColor = ErrorRed
            AddHandler btnClose.MouseLeave, Sub() btnClose.BackColor = PrimaryDark
            titleBar.Controls.Add(btnClose)

            ' Enable dragging from titleBar
            AddHandler lblTitle.MouseDown, AddressOf TitleBar_MouseDown

            ' === Main Input Panel ===
            Dim panelInputs As New Panel() With {
                .Location = New Drawing.Point(20, 70),
                .Size = New Drawing.Size(720, 260),
                .BackColor = Color.White,
                .BorderStyle = BorderStyle.None
            }
            Me.Controls.Add(panelInputs)

            ' Add subtle shadow effect
            Dim shadowPanel As New Panel() With {
                .Location = New Drawing.Point(22, 72),
                .Size = New Drawing.Size(720, 260),
                .BackColor = Color.FromArgb(50, 0, 0, 0)
            }
            Me.Controls.Add(shadowPanel)
            shadowPanel.SendToBack()

            Dim startY As Integer = 60
            Dim rowHeight As Integer = 40

            ' === Professional Separator Lines ===
            Dim verticalLine1 As New Panel() With {.BackColor = NeutralMedium, .Width = 1, .Height = panelInputs.Height, .Location = New Drawing.Point(200, 0)}
            Dim verticalLine2 As New Panel() With {.BackColor = NeutralMedium, .Width = 1, .Height = panelInputs.Height, .Location = New Drawing.Point(520, 0)}
            Dim topLine As New Panel() With {.BackColor = NeutralMedium, .Height = 1, .Width = panelInputs.Width, .Location = New Drawing.Point(0, 50)}
            panelInputs.Controls.AddRange({verticalLine1, verticalLine2, topLine})

            ' === Professional Header Sections ===
            Dim headerBox1 As New Panel() With {.BackColor = PrimaryMedium, .Size = New Drawing.Size(200, 50), .Location = New Drawing.Point(0, 0)}
            Dim headerBox2 As New Panel() With {.BackColor = AccentBlue, .Size = New Drawing.Size(320, 50), .Location = New Drawing.Point(200, 0)}
            Dim headerBox3 As New Panel() With {.BackColor = AccentOrange, .Size = New Drawing.Size(200, 50), .Location = New Drawing.Point(520, 0)}

            Dim lblB As New Label() With {
                .Text = "GRID NAMING",
                .Dock = DockStyle.Fill,
                .Font = New Drawing.Font("Segoe UI", 10, Drawing.FontStyle.Bold),
                .ForeColor = TextLight,
                .TextAlign = Drawing.ContentAlignment.MiddleCenter
            }
            Dim lblC As New Label() With {
                .Text = "GRID CO-ORDINATES",
                .Dock = DockStyle.Fill,
                .Font = New Drawing.Font("Segoe UI", 10, Drawing.FontStyle.Bold),
                .ForeColor = TextLight,
                .TextAlign = Drawing.ContentAlignment.MiddleCenter
            }
            Dim lblD As New Label() With {
                .Text = "LINE EXTENSIONS",
                .Dock = DockStyle.Fill,
                .Font = New Drawing.Font("Segoe UI", 10, Drawing.FontStyle.Bold),
                .ForeColor = TextLight,
                .TextAlign = Drawing.ContentAlignment.MiddleCenter
            }

            headerBox1.Controls.Add(lblB)
            headerBox2.Controls.Add(lblC)
            headerBox3.Controls.Add(lblD)

            panelInputs.Controls.AddRange({headerBox1, headerBox2, headerBox3})

            ' === B column (Grid Naming) ===
            Dim lblBcells = {"X:", "Y:", "Z:"}
            Dim txtBcells = {txtB2, txtB3, txtB4}
            For i = 0 To 2
                Dim yPos = startY + (i * rowHeight)
                Dim lbl As New Label() With {
                    .Text = lblBcells(i),
                    .Location = New Drawing.Point(15, yPos + 8),
                    .Size = New Drawing.Size(25, 20),
                    .Font = New Drawing.Font("Segoe UI", 9, Drawing.FontStyle.Bold),
                    .ForeColor = PrimaryMedium
                }
                panelInputs.Controls.Add(lbl)

                Dim borderPanel As New Panel() With {
                    .Location = New Drawing.Point(40, yPos),
                    .Size = New Drawing.Size(144, 32),
                    .BackColor = NeutralLight
                }
                txtBcells(i).Location = New Drawing.Point(2, 2)
                txtBcells(i).Size = New Drawing.Size(140, 28)
                txtBcells(i).BorderStyle = BorderStyle.None
                txtBcells(i).Font = New Drawing.Font("Segoe UI", 9)
                txtBcells(i).BackColor = Color.White
                txtBcells(i).Text = initialValues(i)
                borderPanel.Controls.Add(txtBcells(i))
                panelInputs.Controls.Add(borderPanel)

                If Not String.IsNullOrWhiteSpace(txtBcells(i).Text) Then borderPanel.BackColor = Color.FromArgb(230, 247, 255)

                AddHandler txtBcells(i).GotFocus, AddressOf TextBox_GotFocus
                AddHandler txtBcells(i).LostFocus, AddressOf TextBox_LostFocus
            Next

            ' === FLIP BUTTONS (Positioned 5mm left with distinct colors) ===
            Dim flipButtonsY As Integer = 350 ' Same Y position as Save/Cancel buttons
            Dim fiveMmLeft As Integer = 20 ' Approximately 5mm from left edge

            ' FLIP X Button - Using a professional teal/cyan color
            btnFlipX.Text = "FLIP X"
            btnFlipX.Size = New Drawing.Size(110, 38) ' Same size as Save button
            btnFlipX.Location = New Drawing.Point(fiveMmLeft, flipButtonsY) ' 5mm from left
            btnFlipX.Font = New Drawing.Font("Segoe UI", 9, Drawing.FontStyle.Bold)
            btnFlipX.BackColor = Color.FromArgb(22, 160, 133) ' Professional Teal
            btnFlipX.ForeColor = TextLight
            btnFlipX.FlatStyle = FlatStyle.Flat
            btnFlipX.FlatAppearance.BorderSize = 0
            btnFlipX.Cursor = Cursors.Hand
            AddHandler btnFlipX.MouseEnter, Sub() btnFlipX.BackColor = Color.FromArgb(Math.Max(0, 22 - 20), Math.Max(0, 160 - 20), Math.Max(0, 133 - 20))
            AddHandler btnFlipX.MouseLeave, Sub() btnFlipX.BackColor = Color.FromArgb(22, 160, 133)
            AddHandler btnFlipX.Click, AddressOf FlipXValues
            Me.Controls.Add(btnFlipX) ' Add to form instead of panel

            ' FLIP Z Button - Using a professional purple color
            btnFlipZ.Text = "FLIP Z"
            btnFlipZ.Size = New Drawing.Size(110, 38) ' Same size as Save button
            btnFlipZ.Location = New Drawing.Point(fiveMmLeft + btnFlipX.Width + 10, flipButtonsY) ' Next to FLIP X button with 10px spacing
            btnFlipZ.Font = New Drawing.Font("Segoe UI", 9, Drawing.FontStyle.Bold)
            btnFlipZ.BackColor = Color.FromArgb(142, 68, 173) ' Professional Purple
            btnFlipZ.ForeColor = TextLight
            btnFlipZ.FlatStyle = FlatStyle.Flat
            btnFlipZ.FlatAppearance.BorderSize = 0
            btnFlipZ.Cursor = Cursors.Hand
            AddHandler btnFlipZ.MouseEnter, Sub() btnFlipZ.BackColor = Color.FromArgb(Math.Max(0, 142 - 20), Math.Max(0, 68 - 20), Math.Max(0, 173 - 20))
            AddHandler btnFlipZ.MouseLeave, Sub() btnFlipZ.BackColor = Color.FromArgb(142, 68, 173)
            AddHandler btnFlipZ.Click, AddressOf FlipZValues
            Me.Controls.Add(btnFlipZ) ' Add to form instead of panel

            ' === AUTO BUTTON (Positioned on the right side) ===
            btnAuto.Text = "AUTO"
            btnAuto.Size = New Drawing.Size(110, 38) ' Same size as other buttons
            btnAuto.Location = New Drawing.Point(Me.ClientSize.Width - btnAuto.Width - 20, flipButtonsY) ' 20px from right edge
            btnAuto.Font = New Drawing.Font("Segoe UI", 9, Drawing.FontStyle.Bold)
            btnAuto.BackColor = Color.FromArgb(52, 152, 219) ' Professional Blue (matches focus color)
            btnAuto.ForeColor = TextLight
            btnAuto.FlatStyle = FlatStyle.Flat
            btnAuto.FlatAppearance.BorderSize = 0
            btnAuto.Cursor = Cursors.Hand
            AddHandler btnAuto.MouseEnter, Sub() btnAuto.BackColor = Color.FromArgb(Math.Max(0, 52 - 20), Math.Max(0, 152 - 20), Math.Max(0, 219 - 20))
            AddHandler btnAuto.MouseLeave, Sub() btnAuto.BackColor = Color.FromArgb(52, 152, 219)
            AddHandler btnAuto.Click, AddressOf AutoOperation
            Me.Controls.Add(btnAuto)

            ' === C column (Coordinates) ===
            Dim lblCcells = {"X-COORDS:", "Y-COORDS:", "Z-COORDS:"}
            Dim txtCcells = {txtC2, txtC3, txtC4}
            For i = 0 To 2
                Dim yPos = startY + (i * rowHeight)
                Dim lbl As New Label() With {
                    .Text = lblCcells(i),
                    .Location = New Drawing.Point(210, yPos + 8),
                    .Size = New Drawing.Size(90, 20),
                    .Font = New Drawing.Font("Segoe UI", 9, Drawing.FontStyle.Bold),
                    .ForeColor = AccentBlue
                }
                panelInputs.Controls.Add(lbl)

                Dim borderPanel As New Panel() With {
                    .Location = New Drawing.Point(310, yPos),
                    .Size = New Drawing.Size(184, 32),
                    .BackColor = NeutralLight
                }
                txtCcells(i).Location = New Drawing.Point(2, 2)
                txtCcells(i).Size = New Drawing.Size(180, 28)
                txtCcells(i).BorderStyle = BorderStyle.None
                txtCcells(i).Font = New Drawing.Font("Segoe UI", 9)
                txtCcells(i).BackColor = Color.White
                txtCcells(i).Text = initialValues(i + 3)
                borderPanel.Controls.Add(txtCcells(i))
                panelInputs.Controls.Add(borderPanel)

                If Not String.IsNullOrWhiteSpace(txtCcells(i).Text) Then borderPanel.BackColor = Color.FromArgb(230, 247, 255)

                AddHandler txtCcells(i).GotFocus, AddressOf TextBox_GotFocus
                AddHandler txtCcells(i).LostFocus, AddressOf TextBox_LostFocus
            Next

            ' === D column (Extensions) ===
            Dim lblDcells = {"EXT-X:", "EXT-Y:", "EXT-Z:"}
            Dim txtDcells = {txtD2, txtD3, txtD4}
            For i = 0 To 2
                Dim yPos = startY + (i * rowHeight)
                Dim lbl As New Label() With {
                    .Text = lblDcells(i),
                    .Location = New Drawing.Point(540, yPos + 8),
                    .Size = New Drawing.Size(70, 20),
                    .Font = New Drawing.Font("Segoe UI", 9, Drawing.FontStyle.Bold),
                    .ForeColor = AccentOrange
                }
                panelInputs.Controls.Add(lbl)

                Dim borderPanel As New Panel() With {
                    .Location = New Drawing.Point(620, yPos),
                    .Size = New Drawing.Size(104, 32),
                    .BackColor = NeutralLight
                }
                txtDcells(i).Location = New Drawing.Point(2, 2)
                txtDcells(i).Size = New Drawing.Size(100, 28)
                txtDcells(i).BorderStyle = BorderStyle.None
                txtDcells(i).Font = New Drawing.Font("Segoe UI", 9)
                txtDcells(i).BackColor = Color.White
                txtDcells(i).Text = initialValues(i + 6)
                borderPanel.Controls.Add(txtDcells(i))
                panelInputs.Controls.Add(borderPanel)

                If Not String.IsNullOrWhiteSpace(txtDcells(i).Text) Then borderPanel.BackColor = Color.FromArgb(230, 247, 255)

                AddHandler txtDcells(i).GotFocus, AddressOf TextBox_GotFocus
                AddHandler txtDcells(i).LostFocus, AddressOf TextBox_LostFocus
            Next

            ' === HEIGHT field (E2) ===
            Dim yPosHeight = startY + (3 * rowHeight)
            Dim lblHeight As New Label() With {
                .Text = "HEIGHT:",
                .Location = New Drawing.Point(540, yPosHeight + 8),
                .Size = New Drawing.Size(70, 20),
                .Font = New Drawing.Font("Segoe UI", 9, Drawing.FontStyle.Bold),
                .ForeColor = AccentGreen
            }
            panelInputs.Controls.Add(lblHeight)

            Dim borderPanelHeight As New Panel() With {
                .Location = New Drawing.Point(620, yPosHeight),
                .Size = New Drawing.Size(104, 32),
                .BackColor = NeutralLight
            }
            txtE2.Location = New Drawing.Point(2, 2)
            txtE2.Size = New Drawing.Size(100, 28)
            txtE2.BorderStyle = BorderStyle.None
            txtE2.Font = New Drawing.Font("Segoe UI", 9)
            txtE2.BackColor = Color.White
            txtE2.Text = initialValues(9)
            borderPanelHeight.Controls.Add(txtE2)
            panelInputs.Controls.Add(borderPanelHeight)

            If Not String.IsNullOrWhiteSpace(txtE2.Text) Then borderPanelHeight.BackColor = Color.FromArgb(230, 247, 255)

            AddHandler txtE2.GotFocus, AddressOf TextBox_GotFocus
            AddHandler txtE2.LostFocus, AddressOf TextBox_LostFocus

            ' ==== Professional Action Buttons ====
            btnOk.Text = "✓ SAVE"
            StyleButton(btnOk, SuccessGreen)
            btnOk.Location = New Drawing.Point(280, 350)
            AddHandler btnOk.Click, Sub() Me.DialogResult = DialogResult.OK

            btnCancel.Text = "✗ CANCEL"
            StyleButton(btnCancel, ErrorRed)
            btnCancel.Location = New Drawing.Point(400, 350)
            AddHandler btnCancel.Click, Sub() Me.DialogResult = DialogResult.Cancel

            Me.Controls.AddRange({btnOk, btnCancel})

            ' Enable Ctrl+A for all textboxes
            AddHandler txtB2.KeyDown, AddressOf TextBox_KeyDown
            AddHandler txtB3.KeyDown, AddressOf TextBox_KeyDown
            AddHandler txtB4.KeyDown, AddressOf TextBox_KeyDown
            AddHandler txtC2.KeyDown, AddressOf TextBox_KeyDown
            AddHandler txtC3.KeyDown, AddressOf TextBox_KeyDown
            AddHandler txtC4.KeyDown, AddressOf TextBox_KeyDown
            AddHandler txtD2.KeyDown, AddressOf TextBox_KeyDown
            AddHandler txtD3.KeyDown, AddressOf TextBox_KeyDown
            AddHandler txtD4.KeyDown, AddressOf TextBox_KeyDown
            AddHandler txtE2.KeyDown, AddressOf TextBox_KeyDown
        End Sub

        ' ==== Professional Button Styling ====
        Private Sub StyleButton(btn As Button, backColor As Drawing.Color)
            btn.Font = New Drawing.Font("Segoe UI", 9, Drawing.FontStyle.Bold)
            btn.BackColor = backColor
            btn.ForeColor = TextLight
            btn.FlatStyle = FlatStyle.Flat
            btn.FlatAppearance.BorderSize = 0
            btn.Size = New Drawing.Size(110, 38)
            btn.Cursor = Cursors.Hand

            ' Add hover effects
            AddHandler btn.MouseEnter, Sub() btn.BackColor = Color.FromArgb(Math.Max(0, backColor.R - 20), Math.Max(0, backColor.G - 20), Math.Max(0, backColor.B - 20))
            AddHandler btn.MouseLeave, Sub() btn.BackColor = backColor
        End Sub

        ' ==== Dragging Logic ====
        <Runtime.InteropServices.DllImport("user32.dll")>
        Private Shared Function ReleaseCapture() As Boolean
        End Function

        <Runtime.InteropServices.DllImport("user32.dll")>
        Private Shared Function SendMessage(hWnd As IntPtr, wMsg As Integer, wParam As Integer, lParam As Integer) As Integer
        End Function

        Private Const WM_NCLBUTTONDOWN As Integer = &HA1
        Private Const HTCAPTION As Integer = &H2

        Private Sub TitleBar_MouseDown(sender As Object, e As MouseEventArgs)
            If e.Button = MouseButtons.Left Then
                ReleaseCapture()
                SendMessage(Me.Handle, WM_NCLBUTTONDOWN, HTCAPTION, 0)
            End If
        End Sub

        ' ==== Ctrl+A Support ====
        Private Sub TextBox_KeyDown(sender As Object, e As KeyEventArgs)
            If e.Control AndAlso e.KeyCode = Keys.A Then
                DirectCast(sender, TextBox).SelectAll()
                e.Handled = True
                e.SuppressKeyPress = True
            End If
        End Sub

        ' ==== Professional Focus Handlers ====
        Private Sub TextBox_GotFocus(sender As Object, e As EventArgs)
            Dim tb As TextBox = DirectCast(sender, TextBox)
            Dim parent As Panel = DirectCast(tb.Parent, Panel)
            parent.BackColor = FocusBlue
        End Sub

        Private Sub TextBox_LostFocus(sender As Object, e As EventArgs)
            Dim tb As TextBox = DirectCast(sender, TextBox)
            Dim parent As Panel = DirectCast(tb.Parent, Panel)
            If String.IsNullOrWhiteSpace(tb.Text) Then
                parent.BackColor = NeutralLight
            Else
                parent.BackColor = Color.FromArgb(212, 237, 218)  ' Light success green
            End If
        End Sub

        ' ==== FLIP OPERATIONS ====
        Private Sub FlipXValues(sender As Object, e As EventArgs)
            Try
                ' Store which control had focus before the flip
                Dim previouslyFocusedControl As Control = Nothing
                If Me.ActiveControl IsNot Nothing Then
                    previouslyFocusedControl = Me.ActiveControl
                End If

                If Not String.IsNullOrWhiteSpace(txtB2.Text) Then
                    txtB2.Text = FlipTextValues(txtB2.Text)
                    ' Trigger visual feedback without changing focus
                    Dim originalColor = txtB2.Parent.BackColor
                    txtB2.Parent.BackColor = Color.FromArgb(255, 230, 255)  ' Light purple flash

                    ' Use a timer to restore the color without blocking the UI
                    Dim timer As New Timer()
                    timer.Interval = 200
                    AddHandler timer.Tick, Sub(s, args)
                                               txtB2.Parent.BackColor = originalColor
                                               timer.Stop()
                                               timer.Dispose()

                                               ' Restore focus to the previously focused control
                                               If previouslyFocusedControl IsNot Nothing Then
                                                   previouslyFocusedControl.Focus()
                                               End If
                                           End Sub
                    timer.Start()
                End If
            Catch ex As Exception
                MessageBox.Show("Error flipping X values: " & ex.Message)
            End Try
        End Sub

        Private Sub FlipZValues(sender As Object, e As EventArgs)
            Try
                ' Store which control had focus before the flip
                Dim previouslyFocusedControl As Control = Nothing
                If Me.ActiveControl IsNot Nothing Then
                    previouslyFocusedControl = Me.ActiveControl
                End If

                If Not String.IsNullOrWhiteSpace(txtB4.Text) Then
                    txtB4.Text = FlipTextValues(txtB4.Text)
                    ' Trigger visual feedback without changing focus
                    Dim originalColor = txtB4.Parent.BackColor
                    txtB4.Parent.BackColor = Color.FromArgb(255, 230, 255)  ' Light purple flash

                    ' Use a timer to restore the color without blocking the UI
                    Dim timer As New Timer()
                    timer.Interval = 200
                    AddHandler timer.Tick, Sub(s, args)
                                               txtB4.Parent.BackColor = originalColor
                                               timer.Stop()
                                               timer.Dispose()

                                               ' Restore focus to the previously focused control
                                               If previouslyFocusedControl IsNot Nothing Then
                                                   previouslyFocusedControl.Focus()
                                               End If
                                           End Sub
                    timer.Start()
                End If
            Catch ex As Exception
                MessageBox.Show("Error flipping Z values: " & ex.Message)
            End Try
        End Sub

        ' ==== AUTO OPERATION ====
        Private Sub AutoOperation(sender As Object, e As EventArgs)
            Try
                ' Define the default values
                Dim alphabetValues As String = "A B C D E F G J K L M N P Q R S T V W X Y Z AA AB AC AD AE AF AG AJ AK AL AM AN AP AQ"
                Dim numberValues As String = "1 2 3 4 5 6 7 8 9 10 11 12 13 14 15 16 17 18 19 20 21 22 23 24 25 26 27 28 29 30 31 32 33 34 35 36 37 38 39 40 41 42 43 44 45 46 47 48 49 50 51 52 53 54 55 56 57 58 59 60 61 62 63 64 65 66 67 68 69 70 71 72 73 74 75 76 77 78 79 80"

                ' Store which control had focus before the operation
                Dim previouslyFocusedControl As Control = Nothing
                If Me.ActiveControl IsNot Nothing Then
                    previouslyFocusedControl = Me.ActiveControl
                End If

                ' Check if both fields are empty - if so, fill with default values
                If String.IsNullOrWhiteSpace(txtB2.Text) AndAlso String.IsNullOrWhiteSpace(txtB4.Text) Then
                    txtB2.Text = alphabetValues
                    txtB4.Text = numberValues
                Else
                    ' Swap the values
                    Dim tempValue As String = txtB2.Text
                    txtB2.Text = txtB4.Text
                    txtB4.Text = tempValue
                End If

                ' Visual feedback for both textboxes
                Dim originalColorB2 = txtB2.Parent.BackColor
                Dim originalColorB4 = txtB4.Parent.BackColor
                txtB2.Parent.BackColor = Color.FromArgb(230, 255, 230)  ' Light green flash
                txtB4.Parent.BackColor = Color.FromArgb(230, 255, 230)  ' Light green flash

                ' Use a timer to restore the colors
                Dim timer As New Timer()
                timer.Interval = 200
                AddHandler timer.Tick, Sub(s, args)
                                           txtB2.Parent.BackColor = If(String.IsNullOrWhiteSpace(txtB2.Text), NeutralLight, Color.FromArgb(212, 237, 218))
                                           txtB4.Parent.BackColor = If(String.IsNullOrWhiteSpace(txtB4.Text), NeutralLight, Color.FromArgb(212, 237, 218))
                                           timer.Stop()
                                           timer.Dispose()

                                           ' Restore focus to the previously focused control
                                           If previouslyFocusedControl IsNot Nothing Then
                                               previouslyFocusedControl.Focus()
                                           End If
                                       End Sub
                timer.Start()

            Catch ex As Exception
                MessageBox.Show("Error in AUTO operation: " & ex.Message)
            End Try
        End Sub

        Private Function FlipTextValues(inputText As String) As String
            Try
                ' Split the text by spaces and reverse the order
                Dim values As String() = inputText.Trim().Split(New Char() {" "c}, StringSplitOptions.RemoveEmptyEntries)
                Array.Reverse(values)
                Return String.Join(" ", values)
            Catch
                ' If any error occurs, return original text
                Return inputText
            End Try
        End Function
    End Class

End Class