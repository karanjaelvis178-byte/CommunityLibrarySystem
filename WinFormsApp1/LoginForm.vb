Imports System.Drawing
Imports Microsoft.Data.SqlClient


Public Class LoginForm

    Private Sub LoginForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load 'event procedures that runs whrn the login form loads
        Try
            Db.Init() ' calls the unit method belonging to the db object
        Catch ex As Exception ' code to handle the erros
            MessageBox.Show("Could not connect to the database." & vbCrLf & vbCrLf & ex.Message & vbCrLf & vbCrLf & 'error handling block
                            "Make sure SQL Server LocalDB is installed.", "Database error",
                            MessageBoxButtons.OK, MessageBoxIcon.Error)
            Application.Exit()
            Return
        End Try

        Me.Text = "Library Management System" 'form title
        Me.Size = New Size(1100, 650) ' form size x.y
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.BackColor = Color.FromArgb(240, 243, 247)
        Me.AcceptButton = btnlogin
        BuildLoginUI()
    End Sub

    Private Sub CollectButtons(parent As Control, list As List(Of Button)) 'the procedure collects all button and put the in one list for when performing similar operation
        For Each c As Control In parent.Controls
            If TypeOf c Is Button Then list.Add(DirectCast(c, Button))
            CollectButtons(c, list)
        Next
    End Sub

    ' ---------------- login logic ----------------
    Private Sub btnlogin_Click(sender As Object, e As EventArgs) Handles btnlogin.Click
        If txtusername.Text.Trim() = "" OrElse txtpassword.Text = "" Then
            MessageBox.Show("Please enter username and password.")
            Return
        End If

        Try
            Dim u = AppData.FindUser(txtusername.Text)
            If u Is Nothing OrElse Not Db.VerifyPassword(txtpassword.Text, u.PasswordHash) Then
                MessageBox.Show("Invalid username or password.", "Login Failed",
                                MessageBoxButtons.OK, MessageBoxIcon.Error)
                txtpassword.Clear()
                txtpassword.Focus()
                Return
            End If

            If u.Status <> "Active" Then
                MessageBox.Show("This account has been deactivated.", "Login Failed",
                                MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            AppData.TouchLogin(u.Id)
            txtusername.Clear()
            txtpassword.Clear()

            Dim dash As New DashboardForm(u.Username, u.Role)
            dash.Show()
            Me.Hide()
        Catch ex As Exception
            MessageBox.Show("Database error: " & ex.Message)
        End Try
    End Sub

    Private Sub RegisterAdmin_Click(sender As Object, e As EventArgs)
        Dim v = FormDialog("Register Admin", {"Username", "Email", "Password", "Confirm Password"},
                           {False, False, True, True})
        If v Is Nothing Then Return

        Try
            If v(0) = "" OrElse v(2) = "" Then
                MessageBox.Show("Username and password are required.")
            ElseIf v(2).Length < 4 Then
                MessageBox.Show("Password must be at least 4 characters.")
            ElseIf v(2) <> v(3) Then
                MessageBox.Show("Passwords do not match.")
            ElseIf Not AppData.AddUser(v(0), v(1), v(2)) Then
                MessageBox.Show("That username already exists.")
            Else
                MessageBox.Show("Account created. You can now log in.")
                txtusername.Text = v(0)
                txtpassword.Focus()
            End If
        Catch ex As Exception
            MessageBox.Show("Database error: " & ex.Message)
        End Try
    End Sub

    Private Sub ForgotPassword_Click(sender As Object, e As EventArgs)
        Dim v = FormDialog("Reset Password", {"Username", "Email", "New Password", "Confirm Password"},
                           {False, False, True, True})
        If v Is Nothing Then Return

        Try
            Dim u = AppData.FindUser(v(0))
            If u Is Nothing OrElse Not String.Equals(u.Email, v(1), StringComparison.OrdinalIgnoreCase) Then
                MessageBox.Show("Username and email do not match any account.")
            ElseIf v(2).Length < 4 Then
                MessageBox.Show("Password must be at least 4 characters.")
            ElseIf v(2) <> v(3) Then
                MessageBox.Show("Passwords do not match.")
            Else
                AppData.SetPassword(u.Id, v(2))
                MessageBox.Show("Password reset. You can now log in.")
            End If
        Catch ex As Exception
            MessageBox.Show("Database error: " & ex.Message)
        End Try
    End Sub

    ' Reusable popup: one label + textbox per entry. Returns the values or Nothing.
    Private Function FormDialog(title As String, labels As String(), isPwd As Boolean()) As String()
        Dim n = labels.Length
        Dim f As New Form With {
            .Text = title, .ClientSize = New Size(380, 40 + n * 45 + 60),
            .StartPosition = FormStartPosition.CenterParent,
            .FormBorderStyle = FormBorderStyle.FixedDialog, .MaximizeBox = False, .MinimizeBox = False
        }
        Dim boxes(n - 1) As TextBox
        For i = 0 To n - 1
            f.Controls.Add(New Label With {.Text = labels(i), .Location = New Point(20, 28 + i * 45), .AutoSize = True})
            boxes(i) = New TextBox With {.Location = New Point(150, 25 + i * 45), .Width = 200,
                                         .UseSystemPasswordChar = isPwd(i)}
            f.Controls.Add(boxes(i))
        Next

        Dim y = 40 + n * 45
        Dim ok As New Button With {.Text = "Save", .Location = New Point(150, y), .Size = New Size(95, 32), .DialogResult = DialogResult.OK}
        Dim cancel As New Button With {.Text = "Cancel", .Location = New Point(255, y), .Size = New Size(95, 32), .DialogResult = DialogResult.Cancel}
        f.Controls.Add(ok)
        f.Controls.Add(cancel)
        f.AcceptButton = ok
        f.CancelButton = cancel

        If f.ShowDialog(Me) <> DialogResult.OK Then Return Nothing
        Dim result(n - 1) As String
        For i = 0 To n - 1
            result(i) = If(isPwd(i), boxes(i).Text, boxes(i).Text.Trim())
        Next
        Return result
    End Function

    ' ---------------- screen layout ----------------
    Private Sub BuildLoginUI()
        Dim oldButtons As New List(Of Button)
        CollectButtons(Me, oldButtons)
        HideOldLabels(Me)

        Dim tbl As New TableLayoutPanel With {.Dock = DockStyle.Fill, .ColumnCount = 2, .RowCount = 1,
                                              .BackColor = Me.BackColor}
        tbl.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 55))
        tbl.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 45))
        tbl.RowStyles.Add(New RowStyle(SizeType.Percent, 100))

        Dim leftPnl As New Panel With {.Dock = DockStyle.Fill}
        AddHandler leftPnl.Paint, AddressOf DrawLibraryScene
        AddHandler leftPnl.Resize, Sub() leftPnl.Invalidate()

        Dim rightPnl As New Panel With {.Dock = DockStyle.Fill}
        Dim box As New Panel With {.Size = New Size(470, 440), .BackColor = Me.BackColor}
        rightPnl.Controls.Add(box)
        AddHandler rightPnl.Resize, Sub()
                                        box.Location = New Point(Math.Max(0, (rightPnl.Width - box.Width) \ 2),
                                                                 Math.Max(0, (rightPnl.Height - box.Height) \ 2))
                                    End Sub

        ' icon, title, subtitle
        box.Controls.Add(New Label With {
            .Text = Char.ConvertFromUtf32(&H1F4D6), .Font = New Font("Segoe UI Emoji", 36),
            .ForeColor = Color.FromArgb(33, 150, 243), .AutoSize = False, .Size = New Size(470, 75),
            .Location = New Point(0, 0), .TextAlign = ContentAlignment.MiddleCenter
        })
        box.Controls.Add(MakeLabel("Library Management System", 0, 80, 20, Color.FromArgb(33, 150, 243), 470, True))
        box.Controls.Add(MakeLabel("Please login to continue", 0, 135, 12, Color.FromArgb(165, 42, 42), 470, True))

        ' login group
        Dim grp As New GroupBox With {.Text = "Login", .Location = New Point(0, 185), .Size = New Size(470, 230)}
        grp.Controls.Add(MakeLabel("Username:", 25, 55, 11, Color.FromArgb(165, 42, 42)))
        grp.Controls.Add(MakeLabel("Password:", 25, 105, 11, Color.FromArgb(165, 42, 42)))

        txtusername.Font = New Font("Segoe UI", 11)
        txtpassword.Font = New Font("Segoe UI", 11)
        txtusername.Location = New Point(150, 52)
        txtpassword.Location = New Point(150, 102)
        txtusername.Width = 295
        txtpassword.Width = 295
        txtpassword.UseSystemPasswordChar = True
        grp.Controls.Add(txtusername)
        grp.Controls.Add(txtpassword)

        For Each b In oldButtons
            Dim t = b.Text.ToLower()
            If b Is btnlogin Then
                StyleButton(b, "Login", 25, 165, 120)
                grp.Controls.Add(b)
            ElseIf t.Contains("register") Then
                StyleButton(b, "Register Admin", 155, 165, 140)
                AddHandler b.Click, AddressOf RegisterAdmin_Click
                grp.Controls.Add(b)
            ElseIf t.Contains("forgot") Then
                StyleButton(b, "Forgot Password?", 305, 165, 140)
                AddHandler b.Click, AddressOf ForgotPassword_Click
                grp.Controls.Add(b)
            End If
        Next
        box.Controls.Add(grp)

        tbl.Controls.Add(leftPnl, 0, 0)
        tbl.Controls.Add(rightPnl, 1, 0)
        Me.Controls.Add(tbl)
        tbl.BringToFront()
        txtusername.Focus()
    End Sub

    Private Sub HideOldLabels(parent As Control)
        For Each c As Control In parent.Controls
            If TypeOf c Is Label Then c.Visible = False
            HideOldLabels(c)
        Next
    End Sub

    Private Function MakeLabel(text As String, x As Integer, y As Integer, size As Single, color As Color,
                               Optional width As Integer = 0, Optional center As Boolean = False) As Label
        Dim l As New Label With {.Text = text, .Location = New Point(x, y),
                                 .Font = New Font("Segoe UI", size, FontStyle.Bold), .ForeColor = color}
        If width > 0 Then
            l.AutoSize = False
            l.Size = New Size(width, CInt(size * 2.4))
            If center Then l.TextAlign = ContentAlignment.MiddleCenter
        Else
            l.AutoSize = True
        End If
        Return l
    End Function

    Private Sub StyleButton(b As Button, text As String, x As Integer, y As Integer, w As Integer)
        b.Text = text
        b.Location = New Point(x, y)
        b.Size = New Size(w, 42)
        b.FlatStyle = FlatStyle.Flat
        b.FlatAppearance.BorderColor = Color.FromArgb(210, 215, 225)
        b.BackColor = Color.White
        b.Font = New Font("Segoe UI", 9, FontStyle.Bold)
        b.Cursor = Cursors.Hand
    End Sub

    ' ---------- drawn illustration (library desk with bookshelves) ----------
    Private Sub FillR(g As Graphics, c As Color, x As Single, y As Single, w As Single, h As Single)
        Using br As New SolidBrush(c)
            g.FillRectangle(br, x, y, w, h)
        End Using
    End Sub

    Private Sub FillE(g As Graphics, c As Color, x As Single, y As Single, w As Single, h As Single)
        Using br As New SolidBrush(c)
            g.FillEllipse(br, x, y, w, h)
        End Using
    End Sub

    Private Sub DrawShelf(g As Graphics, x As Single, y As Single, w As Single, h As Single, seed As Integer)
        Dim rd As New Random(seed)
        Dim pal = {Color.FromArgb(139, 35, 35), Color.FromArgb(200, 90, 40), Color.FromArgb(40, 70, 140),
                   Color.FromArgb(90, 50, 30), Color.FromArgb(220, 150, 60), Color.FromArgb(60, 120, 90)}
        FillR(g, Color.FromArgb(235, 242, 252), x, y, w, h)
        Dim rowH As Single = h / 4
        For r = 0 To 3
            Dim bx As Single = x + 8
            Dim baseY As Single = y + (r + 1) * rowH - 4
            Do While bx < x + w - 14
                Dim bw = rd.Next(8, 16)
                Dim bh As Single = rowH - 12 - rd.Next(0, 14)
                FillR(g, pal(rd.Next(pal.Length)), bx, baseY - bh, bw, bh)
                bx += bw + 1
            Loop
            FillR(g, Color.FromArgb(60, 90, 160), x, baseY, w, 4)
        Next
        Using p As New Pen(Color.FromArgb(60, 90, 160), 6)
            g.DrawRectangle(p, x, y, w, h)
        End Using
    End Sub

    Private Sub DrawLibraryScene(sender As Object, e As PaintEventArgs)
        Dim pnl = DirectCast(sender, Panel)
        Dim g = e.Graphics
        g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias

        Dim sc As Single = Math.Min(pnl.Width / 560.0F, pnl.Height / 380.0F) * 0.9F
        g.TranslateTransform((pnl.Width - 560 * sc) / 2, (pnl.Height - 380 * sc) / 2)
        g.ScaleTransform(sc, sc)

        FillE(g, Color.FromArgb(226, 236, 248), 0, 30, 560, 330)
        DrawShelf(g, 30, 70, 150, 230, 3)
        DrawShelf(g, 380, 70, 150, 230, 7)

        FillE(g, Color.FromArgb(90, 160, 70), 398, 22, 40, 40)
        FillE(g, Color.FromArgb(120, 190, 90), 420, 12, 30, 35)
        FillR(g, Color.FromArgb(230, 140, 40), 405, 58, 30, 12)

        FillE(g, Color.FromArgb(40, 30, 30), 252, 112, 56, 62)
        FillE(g, Color.FromArgb(245, 205, 175), 260, 124, 40, 44)
        FillR(g, Color.FromArgb(90, 150, 220), 252, 168, 56, 55)

        FillR(g, Color.FromArgb(160, 50, 40), 195, 218, 190, 82)
        FillR(g, Color.FromArgb(120, 35, 30), 185, 208, 210, 12)
        FillR(g, Color.FromArgb(230, 150, 50), 215, 238, 150, 34)
        FillR(g, Color.FromArgb(60, 90, 160), 330, 196, 40, 12)
        FillR(g, Color.FromArgb(200, 90, 40), 335, 186, 32, 10)
        Using f As New Font("Segoe UI", 13, FontStyle.Bold),
              br As New SolidBrush(Color.FromArgb(139, 35, 35)),
              sf As New StringFormat With {.Alignment = StringAlignment.Center, .LineAlignment = StringAlignment.Center}
            g.DrawString("LIBRARY", f, br, New RectangleF(215, 238, 150, 34), sf)
        End Using

        Using p As New Pen(Color.FromArgb(120, 130, 150), 3)
            g.DrawLine(p, 20, 300, 540, 300)
        End Using
    End Sub

End Class

