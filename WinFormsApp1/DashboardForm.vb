Imports System.Drawing
Imports Microsoft.Data.SqlClient

Public Class DashboardForm

    Private userName As String = "Admin"
    Private userRole As String = "Librarian"
    Private loggingOut As Boolean = False
    Private loading As Boolean = True
    Private dirty As Boolean = False
    Private WithEvents saveTimer As New System.Windows.Forms.Timer With {.Interval = 700}

    Private ReadOnly Blue As Color = Color.FromArgb(33, 150, 243)

    ' dashboard
    Private pnlDashboard As Panel
    Private lblClock As Label
    Private lblTotalBooks As Label
    Private lblTotalMembers As Label
    Private lblBooksIssued As Label
    Private lblOverdue As Label
    Private dgvRecent As DataGridView

    ' books
    Private pnlBooks As Panel
    Private dgvBooks As DataGridView
    Private txtSearch As TextBox

    ' members
    Private pnlMembers As Panel
    Private dgvMembers As DataGridView
    Private txtSearchMembers As TextBox

    ' issue / return
    Private pnlIssue As Panel
    Private cboBook As ComboBox
    Private cboMember As ComboBox
    Private numDays As NumericUpDown
    Private dgvReturn As DataGridView
    Private dgvIssued As DataGridView
    Private nextIssueId As Integer = 0

    ' librarians
    Private pnlLibrarians As Panel
    Private dgvLibrarians As DataGridView

    ' reports
    Private pnlReports As Panel
    Private repLabels As New Dictionary(Of String, Label)

    Public Sub New()
        InitializeComponent()
    End Sub

    Public Sub New(user As String, role As String)
        InitializeComponent()
        userName = user
        userRole = role
    End Sub

    Private Sub DashboardForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Text = "Library Management System"
        Me.Size = New Size(1100, 650)
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.BackColor = Color.White

        loading = True
        BuildUI()
        Try
            LoadBooks()
            LoadMembers()
            LoadIssues()
            LoadLibrarians()
        Catch ex As Exception
            MessageBox.Show("Could not load data: " & ex.Message)
        End Try
        HookSaveEvents()
        loading = False
        RefreshAll()
    End Sub

    ' =====================================================================
    '  SAVING / LOADING
    ' =====================================================================
    Private Sub HookSaveEvents()
        For Each g In {dgvBooks, dgvMembers, dgvIssued}
            AddHandler g.RowsAdded, Sub(s, e) MarkDirty()
            AddHandler g.RowsRemoved, Sub(s, e) MarkDirty()
            AddHandler g.CellValueChanged, Sub(s, e) MarkDirty()
        Next
    End Sub

    Private Sub MarkDirty()
        If loading Then Return
        dirty = True
        saveTimer.Stop()
        saveTimer.Start()
    End Sub

    Private Sub saveTimer_Tick(sender As Object, e As EventArgs) Handles saveTimer.Tick
        saveTimer.Stop()
        If dirty Then SaveAll()
    End Sub

    Private Function CellStr(r As DataGridViewRow, i As Integer) As String
        Dim v = r.Cells(i).Value
        Return If(v Is Nothing, "", v.ToString())
    End Function

    Private Function CellInt(r As DataGridViewRow, i As Integer) As Integer
        Dim n As Integer
        Integer.TryParse(CellStr(r, i), n)
        Return n
    End Function

    ' Writes Books, Members and Issues to the database in one transaction
    Private Sub SaveAll()
        Try
            Using c As SqlConnection = Db.Open()
                Using tx As SqlTransaction = c.BeginTransaction()
                    Db.RunTx(c, tx, "DELETE FROM Issues")
                    Db.RunTx(c, tx, "DELETE FROM Books")
                    Db.RunTx(c, tx, "DELETE FROM Members")

                    For Each r As DataGridViewRow In dgvBooks.Rows
                        Db.RunTx(c, tx,
                            "INSERT INTO Books (Title, Author, Category, ISBN, Total, Available, Publisher, PubYear, AddedDate) " &
                            "VALUES (@a, @b, @c, @d, @e, @f, @g, @h, @i)",
                            "@a", CellStr(r, 0), "@b", CellStr(r, 1), "@c", CellStr(r, 2),
                            "@d", CellStr(r, 3), "@e", CellInt(r, 4), "@f", CellInt(r, 5),
                            "@g", CellStr(r, 6), "@h", CellInt(r, 7), "@i", CellStr(r, 8))
                    Next

                    For Each r As DataGridViewRow In dgvMembers.Rows
                        Db.RunTx(c, tx,
                            "INSERT INTO Members (Id, Name, Email, Phone, Address, Joined, Status) " &
                            "VALUES (@a, @b, @c, @d, @e, @f, @g)",
                            "@a", CellInt(r, 0), "@b", CellStr(r, 1), "@c", CellStr(r, 2),
                            "@d", CellStr(r, 3), "@e", CellStr(r, 4), "@f", CellStr(r, 5),
                            "@g", CellStr(r, 6))
                    Next

                    For Each r As DataGridViewRow In dgvIssued.Rows
                        Db.RunTx(c, tx,
                            "INSERT INTO Issues (Id, BookTitle, MemberName, IssueDate, DueDate, Status, IssuedBy) " &
                            "VALUES (@a, @b, @c, @d, @e, @f, @g)",
                            "@a", CellInt(r, 0), "@b", CellStr(r, 1), "@c", CellStr(r, 2),
                            "@d", CellStr(r, 3), "@e", CellStr(r, 4), "@f", CellStr(r, 5),
                            "@g", CellStr(r, 6))
                    Next

                    tx.Commit()
                End Using
            End Using
            dirty = False
        Catch ex As Exception
            MessageBox.Show("Could not save data: " & ex.Message)
        End Try
    End Sub

    Private Sub FillGrid(g As DataGridView, sql As String)
        g.Rows.Clear()
        Using c As SqlConnection = Db.Open()
            Using cmd As New SqlCommand(sql, c)
                Using rd As SqlDataReader = cmd.ExecuteReader()
                    While rd.Read()
                        Dim vals(rd.FieldCount - 1) As Object
                        rd.GetValues(vals)
                        g.Rows.Add(vals)
                    End While
                End Using
            End Using
        End Using
    End Sub

    Private Sub LoadBooks()
        FillGrid(dgvBooks, "SELECT Title, Author, Category, ISBN, Total, Available, Publisher, PubYear, AddedDate FROM Books ORDER BY Id")
    End Sub

    Private Sub LoadMembers()
        FillGrid(dgvMembers, "SELECT Id, Name, Email, Phone, Address, Joined, Status FROM Members ORDER BY Id")
    End Sub

    Private Sub LoadIssues()
        FillGrid(dgvIssued, "SELECT Id, BookTitle, MemberName, IssueDate, DueDate, Status, IssuedBy FROM Issues ORDER BY Id")
        nextIssueId = 0
        For Each r As DataGridViewRow In dgvIssued.Rows
            nextIssueId = Math.Max(nextIssueId, CInt(r.Cells(0).Value))
        Next
    End Sub

    Private Sub LoadLibrarians()
        dgvLibrarians.Rows.Clear()
        For Each u In AppData.GetUsers()
            dgvLibrarians.Rows.Add(u.Id, u.Username, u.Email, u.Created, u.LastLogin, u.Status)
        Next
    End Sub

    ' =====================================================================
    '  MAIN LAYOUT
    ' =====================================================================
    Private Sub BuildUI()
        '---------- DASHBOARD CONTENT ----------
        Dim content As New Panel With {.Dock = DockStyle.Fill, .Padding = New Padding(20), .BackColor = Color.White}

        dgvRecent = New DataGridView With {
            .Dock = DockStyle.Fill, .ReadOnly = True, .AllowUserToAddRows = False,
            .RowHeadersVisible = False, .BackgroundColor = Color.White,
            .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
            .SelectionMode = DataGridViewSelectionMode.FullRowSelect
        }
        For Each col In {"Issue ID", "Book Title", "Member Name", "Issued By", "Issue Date", "Status"}
            dgvRecent.Columns.Add(col, col)
        Next
        content.Controls.Add(dgvRecent)

        content.Controls.Add(New Label With {
            .Text = "Recently Issued Books", .Dock = DockStyle.Top, .Height = 40,
            .Font = New Font("Segoe UI", 13, FontStyle.Bold), .TextAlign = ContentAlignment.BottomLeft
        })

        Dim cards As New TableLayoutPanel With {.Dock = DockStyle.Top, .Height = 150, .ColumnCount = 4, .RowCount = 1}
        For i = 1 To 4
            cards.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 25))
        Next
        lblTotalBooks = AddCard(cards, 0, "Total Books", Color.FromArgb(33, 150, 243))
        lblTotalMembers = AddCard(cards, 1, "Total Members", Color.FromArgb(76, 175, 80))
        lblBooksIssued = AddCard(cards, 2, "Books Issued", Color.FromArgb(255, 152, 0))
        lblOverdue = AddCard(cards, 3, "Overdue Books", Color.FromArgb(244, 67, 54))
        content.Controls.Add(cards)

        content.Controls.Add(New Label With {
            .Text = "Admin Dashboard", .Dock = DockStyle.Top, .Height = 50,
            .Font = New Font("Segoe UI", 18, FontStyle.Bold), .ForeColor = Blue,
            .TextAlign = ContentAlignment.MiddleCenter
        })

        Me.Controls.Add(content)
        pnlDashboard = content

        '---------- OTHER PAGES (added before sidebar/header so docking works) ----------
        BuildBooksPage()
        BuildMembersPage()
        BuildIssuePage()
        BuildLibrariansPage()
        BuildReportsPage()

        '---------- SIDEBAR ----------
        Dim sidebar As New Panel With {.Dock = DockStyle.Left, .Width = 200, .Padding = New Padding(10), .BackColor = Color.FromArgb(245, 247, 250)}
        Dim items = {"Dashboard", "Books", "Members", "Issue / Return", "Librarians", "Reports"}
        For i = items.Length - 1 To 0 Step -1
            Dim b As New Button With {
                .Text = items(i), .Dock = DockStyle.Top, .Height = 45,
                .FlatStyle = FlatStyle.Flat, .BackColor = Color.White,
                .Font = New Font("Segoe UI", 10, FontStyle.Bold),
                .TextAlign = ContentAlignment.MiddleLeft, .Padding = New Padding(10, 0, 0, 0),
                .Cursor = Cursors.Hand
            }
            b.FlatAppearance.BorderColor = Color.LightGray
            If items(i) = "Librarians" AndAlso userRole <> "Admin" Then b.Visible = False
            AddHandler b.Click, AddressOf MenuButton_Click
            sidebar.Controls.Add(b)
            sidebar.Controls.Add(New Panel With {.Dock = DockStyle.Top, .Height = 10, .BackColor = sidebar.BackColor})
        Next
        Me.Controls.Add(sidebar)

        '---------- HEADER ----------
        Dim header As New Panel With {.Dock = DockStyle.Top, .Height = 70, .BackColor = Color.White,
                                      .Padding = New Padding(0, 15, 15, 15)}

        Dim lblTitle As New Label With {
            .Text = "Library Management System", .AutoSize = True, .Location = New Point(20, 22),
            .Font = New Font("Segoe UI", 14, FontStyle.Bold), .ForeColor = Blue
        }

        Dim lblWelcome As New Label With {
    .Text = "Welcome, " & userName & " (" & userRole & ")",
    .AutoSize = True,
    .Location = New Point(340, 27),
    .Font = New Font("Segoe UI", 10)
}


        lblClock = New Label With {
            .AutoSize = False, .Width = 190, .Dock = DockStyle.Right,
            .TextAlign = ContentAlignment.MiddleCenter, .Font = New Font("Segoe UI", 9)
        }
        Dim btnLogout As New Button With {
            .Text = "Logout", .Width = 100, .Dock = DockStyle.Right,
            .FlatStyle = FlatStyle.Flat, .BackColor = Color.White, .Cursor = Cursors.Hand
        }
        AddHandler btnLogout.Click, AddressOf btnLogout_Click

        header.Controls.AddRange({lblTitle, lblWelcome, lblClock, btnLogout})
        Me.Controls.Add(header)

        '---------- CLOCK ----------
        Dim t As New System.Windows.Forms.Timer With {.Interval = 1000}
        AddHandler t.Tick, Sub() lblClock.Text = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
        t.Start()
        lblClock.Text = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
    End Sub

    Private Function AddCard(parent As TableLayoutPanel, col As Integer, title As String, numColor As Color) As Label
        Dim card As New Panel With {.Dock = DockStyle.Fill, .Margin = New Padding(10), .BorderStyle = BorderStyle.FixedSingle}
        Dim lblNum As New Label With {
            .Text = "0", .Dock = DockStyle.Fill, .TextAlign = ContentAlignment.MiddleCenter,
            .Font = New Font("Segoe UI", 24, FontStyle.Bold), .ForeColor = numColor
        }
        Dim lblName As New Label With {
            .Text = title, .Dock = DockStyle.Top, .Height = 50, .TextAlign = ContentAlignment.BottomCenter,
            .Font = New Font("Segoe UI", 11, FontStyle.Bold)
        }
        card.Controls.Add(lblNum)
        card.Controls.Add(lblName)
        parent.Controls.Add(card, col, 0)
        Return lblNum
    End Function
    Private Function IsAdmin() As Boolean
        If userRole = "Admin" Then Return True
        MessageBox.Show("Only an Admin can do this.", "Access denied",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning)
        Return False
    End Function
    Private Sub MenuButton_Click(sender As Object, e As EventArgs)
        Select Case DirectCast(sender, Button).Text
            Case "Dashboard" : ShowPage(pnlDashboard)
            Case "Books" : ShowPage(pnlBooks)
            Case "Members" : ShowPage(pnlMembers)
            Case "Issue / Return" : ShowPage(pnlIssue)
            Case "Librarians"
                LoadLibrarians()
                ShowPage(pnlLibrarians)
            Case "Reports" : ShowPage(pnlReports)
        End Select
    End Sub

    Private Sub ShowPage(p As Panel)
        RefreshAll()
        For Each x In {pnlDashboard, pnlBooks, pnlMembers, pnlIssue, pnlLibrarians, pnlReports}
            x.Visible = (x Is p)
        Next
    End Sub

    Private Sub btnLogout_Click(sender As Object, e As EventArgs)
        If MessageBox.Show("Log out?", "Logout", MessageBoxButtons.YesNo) = DialogResult.Yes Then
            loggingOut = True
            Dim lf = Application.OpenForms.OfType(Of LoginForm)().FirstOrDefault()
            If lf IsNot Nothing Then lf.Show()
            Me.Close()
        End If
    End Sub

    Private Sub DashboardForm_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        If dirty Then SaveAll()
    End Sub

    Private Sub DashboardForm_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed
        If Not loggingOut Then Application.Exit()
    End Sub

    ' =====================================================================
    '  SHARED HELPERS
    ' =====================================================================
    Private Function NewGrid(cols As String()) As DataGridView
        Dim g As New DataGridView With {
            .Dock = DockStyle.Fill, .ReadOnly = True, .AllowUserToAddRows = False,
            .RowHeadersVisible = False, .BackgroundColor = Color.White,
            .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
            .SelectionMode = DataGridViewSelectionMode.FullRowSelect, .MultiSelect = False
        }
        For Each c In cols
            g.Columns.Add(c, c)
        Next
        Return g
    End Function

    Private Sub FilterGrid(grid As DataGridView, text As String)
        Dim q = text.Trim().ToLower()
        grid.CurrentCell = Nothing
        For Each row As DataGridViewRow In grid.Rows
            Dim match = (q = "")
            For Each cell As DataGridViewCell In row.Cells
                If cell.Value IsNot Nothing AndAlso cell.Value.ToString().ToLower().Contains(q) Then
                    match = True
                    Exit For
                End If
            Next
            row.Visible = match
        Next
    End Sub

    Private Function PageTitle(text As String) As Label
        Return New Label With {
            .Text = text, .Dock = DockStyle.Top, .Height = 50,
            .Font = New Font("Segoe UI", 18, FontStyle.Bold), .ForeColor = Blue,
            .TextAlign = ContentAlignment.MiddleCenter
        }
    End Function

    ' =====================================================================
    '  BOOKS
    ' =====================================================================
    Private Sub BuildBooksPage()
        pnlBooks = New Panel With {.Dock = DockStyle.Fill, .Padding = New Padding(20),
                                   .BackColor = Color.White, .Visible = False}
        Me.Controls.Add(pnlBooks)

        dgvBooks = NewGrid({"Title", "Author", "Category", "ISBN", "Total", "Available",
                            "Publisher", "Publication Year", "Added Date"})
        pnlBooks.Controls.Add(dgvBooks)

        Dim bar As New Panel With {.Dock = DockStyle.Top, .Height = 50}
        Dim lblSearch As New Label With {.Text = "Search:", .AutoSize = True, .Location = New Point(0, 15)}
        txtSearch = New TextBox With {.Location = New Point(60, 12), .Width = 220}
        AddHandler txtSearch.TextChanged, Sub() FilterGrid(dgvBooks, txtSearch.Text)

        Dim buttons As New FlowLayoutPanel With {.Dock = DockStyle.Right, .AutoSize = True, .WrapContents = False}
        For Each t In {"Add Book", "Edit Book", "Delete Book", "Refresh"}

            Dim b As New Button With {.Text = t, .AutoSize = True, .Height = 30,
                              .FlatStyle = FlatStyle.Flat, .Cursor = Cursors.Hand}
            AddHandler b.Click, AddressOf BookButton_Click
            If t = "Delete Book" AndAlso userRole <> "Admin" Then b.Visible = False
            buttons.Controls.Add(b)
        Next
        bar.Controls.Add(buttons)
        bar.Controls.Add(lblSearch)
        bar.Controls.Add(txtSearch)
        pnlBooks.Controls.Add(bar)

        pnlBooks.Controls.Add(PageTitle("Books Management"))
    End Sub

    Private Sub BookButton_Click(sender As Object, e As EventArgs)
        Select Case DirectCast(sender, Button).Text
            Case "Add Book"
                Dim v = BookDialog()
                If v Is Nothing Then Return
                For Each r As DataGridViewRow In dgvBooks.Rows
                    If CStr(r.Cells(3).Value).ToLower() = v(3).ToLower() AndAlso v(3) <> "" Then
                        MessageBox.Show("A book with that ISBN already exists.")
                        Return
                    End If
                Next
                Dim total = CInt(v(5))
                dgvBooks.Rows.Add(v(0), v(1), v(2), v(3), total, total, v(4), CInt(v(6)),
                                  Date.Today.ToString("yyyy-MM-dd"))
                RefreshAll()

            Case "Edit Book"
                If dgvBooks.SelectedRows.Count = 0 Then
                    MessageBox.Show("Select a book first.")
                    Return
                End If
                Dim r = dgvBooks.SelectedRows(0)
                Dim oldTotal = CInt(r.Cells(4).Value)
                Dim oldAvail = CInt(r.Cells(5).Value)
                Dim v = BookDialog({CStr(r.Cells(0).Value), CStr(r.Cells(1).Value), CStr(r.Cells(2).Value),
                                    CStr(r.Cells(3).Value), CStr(r.Cells(6).Value),
                                    CStr(r.Cells(4).Value), CStr(r.Cells(7).Value)})
                If v Is Nothing Then Return
                Dim newTotal = CInt(v(5))
                r.Cells(0).Value = v(0)
                r.Cells(1).Value = v(1)
                r.Cells(2).Value = v(2)
                r.Cells(3).Value = v(3)
                r.Cells(6).Value = v(4)
                r.Cells(4).Value = newTotal
                r.Cells(7).Value = CInt(v(6))
                r.Cells(5).Value = Math.Min(newTotal, Math.Max(0, oldAvail + (newTotal - oldTotal)))
                RefreshAll()

            Case "Delete Book"
                If Not IsAdmin() Then Return
                If dgvBooks.SelectedRows.Count = 0 Then
                    MessageBox.Show("Select a book first.")
                ElseIf MessageBox.Show("Delete selected book?", "Confirm", MessageBoxButtons.YesNo) = DialogResult.Yes Then
                    dgvBooks.Rows.Remove(dgvBooks.SelectedRows(0))
                    RefreshAll()
                End If

            Case "Refresh"
                txtSearch.Clear()
                dgvBooks.ClearSelection()
        End Select
    End Sub

    ' Returns {title, author, category, isbn, publisher, total, year} or Nothing
    Private Function BookDialog(Optional v As String() = Nothing) As String()
        Dim f As New Form With {
            .Text = If(v Is Nothing, "Add Book", "Edit Book"),
            .Size = New Size(400, 470), .StartPosition = FormStartPosition.CenterParent,
            .FormBorderStyle = FormBorderStyle.FixedDialog, .MaximizeBox = False, .MinimizeBox = False
        }
        Dim names = {"Title", "Author", "Category", "ISBN", "Publisher"}
        Dim boxes(4) As TextBox
        For i = 0 To 4
            f.Controls.Add(New Label With {.Text = names(i), .Location = New Point(20, 25 + i * 45), .AutoSize = True})
            boxes(i) = New TextBox With {.Location = New Point(130, 22 + i * 45), .Width = 230}
            If v IsNot Nothing Then boxes(i).Text = v(i)
            f.Controls.Add(boxes(i))
        Next

        f.Controls.Add(New Label With {.Text = "Total Copies", .Location = New Point(20, 250), .AutoSize = True})
        Dim numTotal As New NumericUpDown With {.Location = New Point(130, 247), .Width = 100,
            .Minimum = 0, .Maximum = 1000, .Value = If(v Is Nothing, 1, CInt(v(5)))}
        f.Controls.Add(numTotal)

        f.Controls.Add(New Label With {.Text = "Pub. Year", .Location = New Point(20, 295), .AutoSize = True})
        Dim numYear As New NumericUpDown With {.Location = New Point(130, 292), .Width = 100,
            .Minimum = 1000, .Maximum = 2100, .Value = If(v Is Nothing, Date.Today.Year, CInt(v(6)))}
        f.Controls.Add(numYear)

        Dim ok As New Button With {.Text = "Save", .Location = New Point(130, 360), .Size = New Size(100, 32), .DialogResult = DialogResult.OK}
        Dim cancel As New Button With {.Text = "Cancel", .Location = New Point(260, 360), .Size = New Size(100, 32), .DialogResult = DialogResult.Cancel}
        f.Controls.Add(ok)
        f.Controls.Add(cancel)
        f.AcceptButton = ok
        f.CancelButton = cancel

        If f.ShowDialog(Me) = DialogResult.OK Then
            If boxes(0).Text.Trim() = "" Then
                MessageBox.Show("Title is required.")
                Return Nothing
            End If
            Return {boxes(0).Text.Trim(), boxes(1).Text.Trim(), boxes(2).Text.Trim(), boxes(3).Text.Trim(),
                    boxes(4).Text.Trim(), numTotal.Value.ToString(), numYear.Value.ToString()}
        End If
        Return Nothing
    End Function

    ' =====================================================================
    '  MEMBERS
    ' =====================================================================
    Private Sub BuildMembersPage()
        pnlMembers = New Panel With {.Dock = DockStyle.Fill, .Padding = New Padding(20),
                                     .BackColor = Color.White, .Visible = False}
        Me.Controls.Add(pnlMembers)

        dgvMembers = NewGrid({"ID", "Name", "Email", "Phone", "Address", "Joined Date", "Status"})
        pnlMembers.Controls.Add(dgvMembers)

        Dim bar As New Panel With {.Dock = DockStyle.Top, .Height = 50}
        Dim lblS As New Label With {.Text = "Search:", .AutoSize = True, .Location = New Point(0, 15)}
        txtSearchMembers = New TextBox With {.Location = New Point(60, 12), .Width = 220}
        AddHandler txtSearchMembers.TextChanged, Sub() FilterGrid(dgvMembers, txtSearchMembers.Text)

        Dim buttons As New FlowLayoutPanel With {.Dock = DockStyle.Right, .AutoSize = True, .WrapContents = False}
        For Each t In {"Add Member", "Edit Member", "Delete Member", "Refresh"}

            Dim b As New Button With {.Text = t, .AutoSize = True, .Height = 30,
                                      .FlatStyle = FlatStyle.Flat, .Cursor = Cursors.Hand}
            AddHandler b.Click, AddressOf MemberButton_Click
            If t = "Delete Member" AndAlso userRole <> "Admin" Then b.Visible = False

            buttons.Controls.Add(b)
        Next
        bar.Controls.Add(buttons)
        bar.Controls.Add(lblS)
        bar.Controls.Add(txtSearchMembers)
        pnlMembers.Controls.Add(bar)

        pnlMembers.Controls.Add(PageTitle("Members Management"))
    End Sub

    Private Sub MemberButton_Click(sender As Object, e As EventArgs)
        Select Case DirectCast(sender, Button).Text


            Case "Add Member"
                Dim v = MemberDialog()
                If v IsNot Nothing Then
                    Dim nextId = 1
                    For Each r As DataGridViewRow In dgvMembers.Rows
                        nextId = Math.Max(nextId, CInt(r.Cells(0).Value) + 1)
                    Next
                    dgvMembers.Rows.Add(nextId, v(0), v(1), v(2), v(3),
                                        DateTime.Today.ToString("yyyy-MM-dd"), v(4))
                    RefreshAll()
                End If

            Case "Edit Member"
                If dgvMembers.SelectedRows.Count = 0 Then
                    MessageBox.Show("Select a member first.")
                    Return
                End If
                Dim r = dgvMembers.SelectedRows(0)
                Dim v = MemberDialog({CStr(r.Cells(1).Value), CStr(r.Cells(2).Value),
                                      CStr(r.Cells(3).Value), CStr(r.Cells(4).Value),
                                      CStr(r.Cells(6).Value)})
                If v IsNot Nothing Then
                    r.Cells(1).Value = v(0)
                    r.Cells(2).Value = v(1)
                    r.Cells(3).Value = v(2)
                    r.Cells(4).Value = v(3)
                    r.Cells(6).Value = v(4)
                    RefreshAll()
                End If

            Case "Delete Member"
                If Not IsAdmin() Then Return
                If dgvMembers.SelectedRows.Count = 0 Then
                    MessageBox.Show("Select a member first.")
                ElseIf MessageBox.Show("Delete selected member?", "Confirm", MessageBoxButtons.YesNo) = DialogResult.Yes Then
                    dgvMembers.Rows.Remove(dgvMembers.SelectedRows(0))
                    RefreshAll()
                End If

            Case "Refresh"
                txtSearchMembers.Clear()
                dgvMembers.ClearSelection()
        End Select
    End Sub

    ' Popup form for adding/editing. Returns {name, email, phone, address, status} or Nothing
    Private Function MemberDialog(Optional vals As String() = Nothing) As String()
        Dim f As New Form With {
            .Text = If(vals Is Nothing, "Add Member", "Edit Member"),
            .Size = New Size(380, 380), .StartPosition = FormStartPosition.CenterParent,
            .FormBorderStyle = FormBorderStyle.FixedDialog, .MaximizeBox = False, .MinimizeBox = False
        }
        Dim names = {"Name", "Email", "Phone", "Address"}
        Dim boxes(3) As TextBox
        For i = 0 To 3
            f.Controls.Add(New Label With {.Text = names(i), .Location = New Point(20, 25 + i * 50), .AutoSize = True})
            boxes(i) = New TextBox With {.Location = New Point(100, 22 + i * 50), .Width = 230}
            If vals IsNot Nothing Then boxes(i).Text = vals(i)
            f.Controls.Add(boxes(i))
        Next

        f.Controls.Add(New Label With {.Text = "Status", .Location = New Point(20, 225), .AutoSize = True})
        Dim cbo As New ComboBox With {.Location = New Point(100, 222), .Width = 230,
                                      .DropDownStyle = ComboBoxStyle.DropDownList}
        cbo.Items.Add("Active")
        cbo.Items.Add("Inactive")
        cbo.SelectedIndex = If(vals IsNot Nothing AndAlso vals(4) = "Inactive", 1, 0)
        f.Controls.Add(cbo)

        Dim ok As New Button With {.Text = "Save", .Location = New Point(100, 280), .Size = New Size(100, 32),
                                   .DialogResult = DialogResult.OK}
        Dim cancel As New Button With {.Text = "Cancel", .Location = New Point(230, 280), .Size = New Size(100, 32),
                                       .DialogResult = DialogResult.Cancel}
        f.Controls.Add(ok)
        f.Controls.Add(cancel)
        f.AcceptButton = ok
        f.CancelButton = cancel

        If f.ShowDialog(Me) = DialogResult.OK Then
            If boxes(0).Text.Trim() = "" Then
                MessageBox.Show("Name is required.")
                Return Nothing
            End If
            Return {boxes(0).Text.Trim(), boxes(1).Text.Trim(), boxes(2).Text.Trim(),
                    boxes(3).Text.Trim(), cbo.SelectedItem.ToString()}
        End If
        Return Nothing
    End Function

    ' =====================================================================
    '  ISSUE / RETURN
    ' =====================================================================
    Private Sub BuildIssuePage()
        pnlIssue = New Panel With {.Dock = DockStyle.Fill, .Padding = New Padding(20),
                                   .BackColor = Color.White, .Visible = False}
        Me.Controls.Add(pnlIssue)

        Dim tabs As New TabControl With {.Dock = DockStyle.Fill}
        Dim tpIssue As New TabPage("Issue Book")
        Dim tpReturn As New TabPage("Return Book")
        Dim tpList As New TabPage("Currently Issued")
        tabs.TabPages.Add(tpIssue)
        tabs.TabPages.Add(tpReturn)
        tabs.TabPages.Add(tpList)
        pnlIssue.Controls.Add(tabs)

        pnlIssue.Controls.Add(PageTitle("Issue/Return Books"))

        ' ---- Issue tab ----
        Dim bold As New Font("Segoe UI", 10, FontStyle.Bold)
        tpIssue.Controls.Add(New Label With {.Text = "Select Book:", .Font = bold, .AutoSize = True, .Location = New Point(40, 40)})
        tpIssue.Controls.Add(New Label With {.Text = "Select Member:", .Font = bold, .AutoSize = True, .Location = New Point(40, 80)})
        tpIssue.Controls.Add(New Label With {.Text = "Due Days:", .Font = bold, .AutoSize = True, .Location = New Point(40, 120)})

        cboBook = New ComboBox With {.Location = New Point(180, 37), .Width = 300, .DropDownStyle = ComboBoxStyle.DropDownList}
        cboMember = New ComboBox With {.Location = New Point(180, 77), .Width = 300, .DropDownStyle = ComboBoxStyle.DropDownList}
        numDays = New NumericUpDown With {.Location = New Point(180, 117), .Width = 80, .Minimum = 1, .Maximum = 90, .Value = 14}

        Dim btnIssue As New Button With {.Text = "Issue Book", .Location = New Point(180, 170), .Size = New Size(130, 34), .FlatStyle = FlatStyle.Flat}
        Dim btnRefresh As New Button With {.Text = "Refresh Lists", .Location = New Point(180, 215), .Size = New Size(130, 34), .FlatStyle = FlatStyle.Flat}
        AddHandler btnIssue.Click, AddressOf IssueBook_Click
        AddHandler btnRefresh.Click, Sub() RefreshAll()

        tpIssue.Controls.Add(cboBook)
        tpIssue.Controls.Add(cboMember)
        tpIssue.Controls.Add(numDays)
        tpIssue.Controls.Add(btnIssue)
        tpIssue.Controls.Add(btnRefresh)

        ' ---- Return tab ----
        Dim cols = {"Issue ID", "Book Title", "Member Name", "Issue Date", "Due Date", "Status"}
        dgvReturn = NewGrid(cols)
        tpReturn.Controls.Add(dgvReturn)
        Dim bottom As New Panel With {.Dock = DockStyle.Bottom, .Height = 55}
        Dim btnReturn As New Button With {.Text = "Return Selected Book", .Location = New Point(10, 10),
                                          .Size = New Size(180, 34), .FlatStyle = FlatStyle.Flat}
        AddHandler btnReturn.Click, AddressOf ReturnBook_Click
        bottom.Controls.Add(btnReturn)
        tpReturn.Controls.Add(bottom)

        ' ---- Currently issued tab (7th column = who issued it, hidden) ----
        dgvIssued = NewGrid(cols)
        dgvIssued.Columns.Add("IssuedBy", "Issued By")
        dgvIssued.Columns("IssuedBy").Visible = False
        tpList.Controls.Add(dgvIssued)
    End Sub

    Private Function FindBook(title As String) As DataGridViewRow
        For Each r As DataGridViewRow In dgvBooks.Rows
            If CStr(r.Cells(0).Value) = title Then Return r
        Next
        Return Nothing
    End Function

    Private Sub IssueBook_Click(sender As Object, e As EventArgs)
        If cboBook.SelectedItem Is Nothing OrElse cboMember.SelectedItem Is Nothing Then
            MessageBox.Show("Select a book and a member.")
            Return
        End If

        Dim title = cboBook.SelectedItem.ToString()
        Dim bookRow = FindBook(title)
        If bookRow Is Nothing OrElse CInt(bookRow.Cells(5).Value) <= 0 Then
            MessageBox.Show("No copies of this book are available.")
            Return
        End If

        bookRow.Cells(5).Value = CInt(bookRow.Cells(5).Value) - 1
        nextIssueId += 1
        Dim issued = Date.Today
        dgvIssued.Rows.Add(nextIssueId, title, cboMember.SelectedItem.ToString(),
                           issued.ToString("yyyy-MM-dd"),
                           issued.AddDays(CInt(numDays.Value)).ToString("yyyy-MM-dd"), "Issued", userName)
        RefreshAll()
        MessageBox.Show("Book issued successfully.")
    End Sub

    Private Sub ReturnBook_Click(sender As Object, e As EventArgs)
        If dgvReturn.SelectedRows.Count = 0 Then
            MessageBox.Show("Select an issued book first.")
            Return
        End If

        Dim id = CInt(dgvReturn.SelectedRows(0).Cells(0).Value)
        For Each r As DataGridViewRow In dgvIssued.Rows
            If CInt(r.Cells(0).Value) = id Then
                r.Cells(5).Value = "Returned"
                Dim b = FindBook(CStr(r.Cells(1).Value))
                If b IsNot Nothing Then b.Cells(5).Value = CInt(b.Cells(5).Value) + 1
                Exit For
            End If
        Next
        RefreshAll()
        MessageBox.Show("Book returned.")
    End Sub

    ' Refreshes drop-downs, statuses, return list and dashboard numbers
    Private Sub RefreshAll()
        ' mark overdue
        For Each r As DataGridViewRow In dgvIssued.Rows
            Dim st = CStr(r.Cells(5).Value)
            If st <> "Returned" Then
                Dim ns = If(Date.Parse(CStr(r.Cells(4).Value)) < Date.Today, "Overdue", "Issued")
                If ns <> st Then r.Cells(5).Value = ns
            End If
        Next

        ' drop-downs
        cboBook.Items.Clear()
        For Each r As DataGridViewRow In dgvBooks.Rows
            If CInt(r.Cells(5).Value) > 0 Then cboBook.Items.Add(CStr(r.Cells(0).Value))
        Next
        cboMember.Items.Clear()
        For Each r As DataGridViewRow In dgvMembers.Rows
            If CStr(r.Cells(6).Value) = "Active" Then cboMember.Items.Add(CStr(r.Cells(1).Value))
        Next

        ' return list = everything not yet returned
        dgvReturn.Rows.Clear()
        For Each r As DataGridViewRow In dgvIssued.Rows
            If CStr(r.Cells(5).Value) <> "Returned" Then
                dgvReturn.Rows.Add(r.Cells(0).Value, r.Cells(1).Value, r.Cells(2).Value,
                                   r.Cells(3).Value, r.Cells(4).Value, r.Cells(5).Value)
            End If
        Next

        ' dashboard cards
        Dim issuedCount = 0
        Dim overdueCount = 0
        For Each r As DataGridViewRow In dgvIssued.Rows
            Select Case CStr(r.Cells(5).Value)
                Case "Issued" : issuedCount += 1
                Case "Overdue" : issuedCount += 1 : overdueCount += 1
            End Select
        Next
        lblTotalBooks.Text = dgvBooks.Rows.Count.ToString()
        lblTotalMembers.Text = dgvMembers.Rows.Count.ToString()
        lblBooksIssued.Text = issuedCount.ToString()
        lblOverdue.Text = overdueCount.ToString()

        ' recently issued table on the dashboard (latest 10)
        dgvRecent.Rows.Clear()
        For i = dgvIssued.Rows.Count - 1 To Math.Max(0, dgvIssued.Rows.Count - 10) Step -1
            Dim r = dgvIssued.Rows(i)
            dgvRecent.Rows.Add(r.Cells(0).Value, r.Cells(1).Value, r.Cells(2).Value,
                               r.Cells(6).Value, r.Cells(3).Value, r.Cells(5).Value)
        Next

        UpdateReportStats()
    End Sub

    ' =====================================================================
    '  LIBRARIANS (stored in the Users table, passwords hashed)
    ' =====================================================================
    Private Sub BuildLibrariansPage()
        pnlLibrarians = New Panel With {.Dock = DockStyle.Fill, .Padding = New Padding(20),
                                        .BackColor = Color.White, .Visible = False}
        Me.Controls.Add(pnlLibrarians)

        dgvLibrarians = NewGrid({"ID", "Username", "Email", "Created Date", "Last Login", "Status"})
        pnlLibrarians.Controls.Add(dgvLibrarians)

        Dim bar As New Panel With {.Dock = DockStyle.Top, .Height = 50}
        Dim buttons As New FlowLayoutPanel With {.Dock = DockStyle.Right, .AutoSize = True, .WrapContents = False}
        For Each t In {"Add Librarian", "Reset Password", "Deactivate", "Delete Librarian", "Refresh"}
            Dim b As New Button With {.Text = t, .AutoSize = True, .Height = 30,
                                      .FlatStyle = FlatStyle.Flat, .Cursor = Cursors.Hand}
            AddHandler b.Click, AddressOf LibrarianButton_Click
            buttons.Controls.Add(b)
        Next
        bar.Controls.Add(buttons)
        pnlLibrarians.Controls.Add(bar)

        pnlLibrarians.Controls.Add(PageTitle("Librarians Management"))
    End Sub

    Private Sub LibrarianButton_Click(sender As Object, e As EventArgs)
        Dim action = DirectCast(sender, Button).Text
        If DirectCast(sender, Button).Text <> "Refresh" AndAlso Not IsAdmin() Then Return

        Try
            Select Case action
                Case "Add Librarian"
                    Dim v = LibrarianDialog()
                    If v Is Nothing Then Return
                    If AppData.AddUser(v(0), v(1), v(2), "Librarian") Then
                        LoadLibrarians()
                    Else
                        MessageBox.Show("That username already exists.")
                    End If

                Case "Reset Password"
                    If dgvLibrarians.SelectedRows.Count = 0 Then
                        MessageBox.Show("Select a librarian first.")
                        Return
                    End If
                    Dim id = CInt(dgvLibrarians.SelectedRows(0).Cells(0).Value)
                    Dim pw = InputDialog("Reset Password", "New password:", True)
                    If pw IsNot Nothing Then
                        If pw.Length < 4 Then
                            MessageBox.Show("Password must be at least 4 characters.")
                        Else
                            AppData.SetPassword(id, pw)
                            MessageBox.Show("Password updated.")
                        End If
                    End If

                Case "Deactivate"
                    If dgvLibrarians.SelectedRows.Count = 0 Then
                        MessageBox.Show("Select a librarian first.")
                        Return
                    End If
                    Dim row = dgvLibrarians.SelectedRows(0)
                    If CStr(row.Cells(1).Value).ToLower() = userName.ToLower() Then
                        MessageBox.Show("You can't deactivate the account you are logged in with.")
                        Return
                    End If
                    Dim newStatus = If(CStr(row.Cells(5).Value) = "Active", "Inactive", "Active")
                    AppData.SetStatus(CInt(row.Cells(0).Value), newStatus)
                    LoadLibrarians()

                Case "Delete Librarian"
                    If dgvLibrarians.SelectedRows.Count = 0 Then
                        MessageBox.Show("Select a librarian first.")
                        Return
                    End If
                    Dim row = dgvLibrarians.SelectedRows(0)
                    Dim uname = CStr(row.Cells(1).Value)
                    If uname.ToLower() = userName.ToLower() Then
                        MessageBox.Show("You can't delete the account you are logged in with.")
                        Return
                    End If
                    If MessageBox.Show("Delete librarian '" & uname & "'? This cannot be undone.",
                                       "Confirm Delete", MessageBoxButtons.YesNo,
                                       MessageBoxIcon.Warning) = DialogResult.Yes Then
                        AppData.DeleteUser(CInt(row.Cells(0).Value))
                        LoadLibrarians()
                    End If

                Case "Refresh"
                    LoadLibrarians()
            End Select
        Catch ex As Exception
            MessageBox.Show("Database error: " & ex.Message)
        End Try
    End Sub

    ' Returns {username, email, password} or Nothing
    Private Function LibrarianDialog() As String()
        Dim f As New Form With {
            .Text = "Add Librarian", .Size = New Size(380, 300), .StartPosition = FormStartPosition.CenterParent,
            .FormBorderStyle = FormBorderStyle.FixedDialog, .MaximizeBox = False, .MinimizeBox = False
        }
        Dim names = {"Username", "Email", "Password"}
        Dim boxes(2) As TextBox
        For i = 0 To 2
            f.Controls.Add(New Label With {.Text = names(i), .Location = New Point(20, 25 + i * 50), .AutoSize = True})
            boxes(i) = New TextBox With {.Location = New Point(110, 22 + i * 50), .Width = 220}
            f.Controls.Add(boxes(i))
        Next
        boxes(2).UseSystemPasswordChar = True

        Dim ok As New Button With {.Text = "Save", .Location = New Point(110, 190), .Size = New Size(100, 32), .DialogResult = DialogResult.OK}
        Dim cancel As New Button With {.Text = "Cancel", .Location = New Point(230, 190), .Size = New Size(100, 32), .DialogResult = DialogResult.Cancel}
        f.Controls.Add(ok)
        f.Controls.Add(cancel)
        f.AcceptButton = ok
        f.CancelButton = cancel

        If f.ShowDialog(Me) = DialogResult.OK Then
            If boxes(0).Text.Trim() = "" OrElse boxes(2).Text = "" Then
                MessageBox.Show("Username and password are required.")
                Return Nothing
            End If
            If boxes(2).Text.Length < 4 Then
                MessageBox.Show("Password must be at least 4 characters.")
                Return Nothing
            End If
            Return {boxes(0).Text.Trim(), boxes(1).Text.Trim(), boxes(2).Text}
        End If
        Return Nothing
    End Function

    ' Small one-box popup. Returns the text or Nothing if cancelled/empty
    Private Function InputDialog(title As String, prompt As String, isPassword As Boolean) As String
        Dim f As New Form With {
            .Text = title, .Size = New Size(360, 190), .StartPosition = FormStartPosition.CenterParent,
            .FormBorderStyle = FormBorderStyle.FixedDialog, .MaximizeBox = False, .MinimizeBox = False
        }
        f.Controls.Add(New Label With {.Text = prompt, .Location = New Point(20, 20), .AutoSize = True})
        Dim tb As New TextBox With {.Location = New Point(20, 50), .Width = 300, .UseSystemPasswordChar = isPassword}
        f.Controls.Add(tb)
        Dim ok As New Button With {.Text = "OK", .Location = New Point(130, 95), .Size = New Size(90, 30), .DialogResult = DialogResult.OK}
        f.Controls.Add(ok)
        f.AcceptButton = ok

        If f.ShowDialog(Me) = DialogResult.OK AndAlso tb.Text <> "" Then Return tb.Text
        Return Nothing
    End Function

    ' =====================================================================
    '  REPORTS
    ' =====================================================================
    Private Sub BuildReportsPage()
        pnlReports = New Panel With {.Dock = DockStyle.Fill, .Padding = New Padding(20),
                                     .BackColor = Color.White, .Visible = False}
        Me.Controls.Add(pnlReports)

        Dim gbGen As New GroupBox With {.Text = "Generate Reports", .Dock = DockStyle.Top, .Height = 110}
        Dim flow As New FlowLayoutPanel With {.Dock = DockStyle.Fill, .Padding = New Padding(10, 15, 0, 0)}
        For Each t In {"All Books Report", "Members Report", "Issued Books Report", "Overdue Books Report"}
            Dim b As New Button With {.Text = t, .Size = New Size(180, 45), .FlatStyle = FlatStyle.Flat, .Cursor = Cursors.Hand}
            AddHandler b.Click, AddressOf ReportButton_Click
            flow.Controls.Add(b)
        Next
        gbGen.Controls.Add(flow)
        pnlReports.Controls.Add(gbGen)

        Dim gbStats As New GroupBox With {.Text = "System Reports", .Dock = DockStyle.Top, .Height = 230}
        Dim tbl As New TableLayoutPanel With {.Dock = DockStyle.Fill, .ColumnCount = 2, .RowCount = 1}
        tbl.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 50))
        tbl.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 50))

        Dim gbLib As New GroupBox With {.Text = "Library Statistics", .Dock = DockStyle.Fill, .Margin = New Padding(10)}
        Dim fl1 As New FlowLayoutPanel With {.Dock = DockStyle.Fill, .FlowDirection = FlowDirection.TopDown, .Padding = New Padding(15, 10, 0, 0)}
        For Each k In {"Total Books", "Total Members", "Currently Issued", "Overdue Books"}
            AddStat(fl1, k)
        Next
        gbLib.Controls.Add(fl1)

        Dim gbAct As New GroupBox With {.Text = "Recent Activity Summary", .Dock = DockStyle.Fill, .Margin = New Padding(10)}
        Dim fl2 As New FlowLayoutPanel With {.Dock = DockStyle.Fill, .FlowDirection = FlowDirection.TopDown, .Padding = New Padding(15, 10, 0, 0)}
        For Each k In {"Issues This Month", "Returns This Month", "New Members", "Net Activity"}
            AddStat(fl2, k)
        Next
        gbAct.Controls.Add(fl2)

        tbl.Controls.Add(gbLib, 0, 0)
        tbl.Controls.Add(gbAct, 1, 0)
        gbStats.Controls.Add(tbl)
        pnlReports.Controls.Add(gbStats)

        pnlReports.Controls.Add(PageTitle("Reports & Analytics"))
    End Sub

    Private Sub AddStat(parent As FlowLayoutPanel, key As String)
        Dim l As New Label With {.Text = key & ": 0", .AutoSize = True,
                                 .Font = New Font("Segoe UI", 12), .Margin = New Padding(0, 6, 0, 6)}
        repLabels(key) = l
        parent.Controls.Add(l)
    End Sub

    Private Sub SetStat(key As String, value As Integer)
        If repLabels.ContainsKey(key) Then repLabels(key).Text = key & ": " & value
    End Sub

    Private Sub UpdateReportStats()
        If repLabels.Count = 0 Then Return
        Dim monthStart As New Date(Date.Today.Year, Date.Today.Month, 1)
        Dim issuesM = 0, returnsM = 0, current = 0, overdue = 0, newMembers = 0

        For Each r As DataGridViewRow In dgvIssued.Rows
            Dim st = CStr(r.Cells(5).Value)
            Dim d As Date
            If Date.TryParse(CStr(r.Cells(3).Value), d) AndAlso d >= monthStart Then issuesM += 1
            If st = "Returned" Then
                returnsM += 1
            Else
                current += 1
                If st = "Overdue" Then overdue += 1
            End If
        Next

        For Each r As DataGridViewRow In dgvMembers.Rows
            Dim d As Date
            If Date.TryParse(CStr(r.Cells(5).Value), d) AndAlso d >= monthStart Then newMembers += 1
        Next

        SetStat("Total Books", dgvBooks.Rows.Count)
        SetStat("Total Members", dgvMembers.Rows.Count)
        SetStat("Currently Issued", current)
        SetStat("Overdue Books", overdue)
        SetStat("Issues This Month", issuesM)
        SetStat("Returns This Month", returnsM)
        SetStat("New Members", newMembers)
        SetStat("Net Activity", issuesM - returnsM)
    End Sub

    Private Sub ReportButton_Click(sender As Object, e As EventArgs)
        RefreshAll()
        Select Case DirectCast(sender, Button).Text
            Case "All Books Report"
                ShowReport("All Books Report", dgvBooks, Nothing)
            Case "Members Report"
                ShowReport("Members Report", dgvMembers, Nothing)
            Case "Issued Books Report"
                ShowReport("Issued Books Report", dgvIssued, Function(r) CStr(r.Cells(5).Value) <> "Returned")
            Case "Overdue Books Report"
                ShowReport("Overdue Books Report", dgvIssued, Function(r) CStr(r.Cells(5).Value) = "Overdue")
        End Select
    End Sub

    ' Opens a popup table copied from a source grid, with CSV export
    Private Sub ShowReport(title As String, src As DataGridView, filter As Func(Of DataGridViewRow, Boolean))
        Dim f As New Form With {.Text = title, .Size = New Size(900, 520), .StartPosition = FormStartPosition.CenterParent}
        Dim g As New DataGridView With {
            .Dock = DockStyle.Fill, .ReadOnly = True, .AllowUserToAddRows = False,
            .RowHeadersVisible = False, .BackgroundColor = Color.White,
            .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        }

        Dim visibleCols As New List(Of Integer)
        For i = 0 To src.Columns.Count - 1
            If src.Columns(i).Visible Then
                visibleCols.Add(i)
                g.Columns.Add(src.Columns(i).HeaderText, src.Columns(i).HeaderText)
            End If
        Next

        For Each r As DataGridViewRow In src.Rows
            If filter Is Nothing OrElse filter(r) Then
                Dim vals(visibleCols.Count - 1) As Object
                For i = 0 To visibleCols.Count - 1
                    vals(i) = r.Cells(visibleCols(i)).Value
                Next
                g.Rows.Add(vals)
            End If
        Next

        Dim bottom As New Panel With {.Dock = DockStyle.Bottom, .Height = 55}
        Dim lblCount As New Label With {.Text = g.Rows.Count & " record(s)", .AutoSize = True, .Location = New Point(15, 18)}
        Dim btnExport As New Button With {.Text = "Export to CSV", .Size = New Size(130, 34), .Anchor = AnchorStyles.Top Or AnchorStyles.Right,
                                          .Location = New Point(f.ClientSize.Width - 290, 10)}
        Dim btnClose As New Button With {.Text = "Close", .Size = New Size(110, 34), .Anchor = AnchorStyles.Top Or AnchorStyles.Right,
                                         .Location = New Point(f.ClientSize.Width - 145, 10)}
        AddHandler btnExport.Click, Sub() ExportGridToCsv(g, title)
        AddHandler btnClose.Click, Sub() f.Close()
        bottom.Controls.Add(lblCount)
        bottom.Controls.Add(btnExport)
        bottom.Controls.Add(btnClose)

        f.Controls.Add(g)
        f.Controls.Add(bottom)
        f.ShowDialog(Me)
    End Sub

    Private Sub ExportGridToCsv(g As DataGridView, title As String)
        Using sfd As New SaveFileDialog With {.Filter = "CSV file (*.csv)|*.csv", .FileName = title.Replace(" ", "_") & ".csv"}
            If sfd.ShowDialog() <> DialogResult.OK Then Return
            Dim sb As New System.Text.StringBuilder()

            Dim head As New List(Of String)
            For Each c As DataGridViewColumn In g.Columns
                head.Add(CsvQuote(c.HeaderText))
            Next
            sb.AppendLine(String.Join(",", head))

            For Each r As DataGridViewRow In g.Rows
                Dim line As New List(Of String)
                For Each c As DataGridViewCell In r.Cells
                    line.Add(CsvQuote(If(c.Value Is Nothing, "", c.Value.ToString())))
                Next
                sb.AppendLine(String.Join(",", line))
            Next

            System.IO.File.WriteAllText(sfd.FileName, sb.ToString(), System.Text.Encoding.UTF8)
            MessageBox.Show("Report saved.")
        End Using
    End Sub

    Private Function CsvQuote(s As String) As String
        Return """" & s.Replace("""", """""") & """"
    End Function

End Class
