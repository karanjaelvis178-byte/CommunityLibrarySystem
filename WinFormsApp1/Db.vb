Imports Microsoft.Data.SqlClient
Imports System.Security.Cryptography

Public Module Db

    Private Const ServerPart As String = "Server=(localdb)\MSSQLLocalDB;Integrated Security=True;Encrypt=False;TrustServerCertificate=True;"

    ' Opens a connection to the LibraryDB database
    Public Function Open() As SqlConnection
        Dim c As New SqlConnection(ServerPart & "Database=LibraryDB;")
        c.Open()
        Return c
    End Function

    ' Runs one INSERT/UPDATE/DELETE. Parameters are passed as name, value, name, value...
    Public Sub Run(sql As String, ParamArray args As Object())
        Using c As SqlConnection = Open()
            RunTx(c, Nothing, sql, args)
        End Using
    End Sub

    Public Sub RunTx(c As SqlConnection, tx As SqlTransaction, sql As String, ParamArray args As Object())
        Using cmd As New SqlCommand(sql, c, tx)
            For i As Integer = 0 To args.Length - 2 Step 2
                cmd.Parameters.AddWithValue(CStr(args(i)), args(i + 1))
            Next
            cmd.ExecuteNonQuery()
        End Using
    End Sub

    ' Creates the database and tables the first time the app runs
    Public Sub Init()
        Using c As New SqlConnection(ServerPart & "Database=master;")
            c.Open()
            RunTx(c, Nothing, "IF DB_ID('LibraryDB') IS NULL CREATE DATABASE LibraryDB")
        End Using

        Using c As SqlConnection = Open()
            RunTx(c, Nothing,
                "IF OBJECT_ID('dbo.Users','U') IS NULL CREATE TABLE dbo.Users (" &
                "Id INT IDENTITY(1,1) PRIMARY KEY, " &
                "Username NVARCHAR(100) NOT NULL UNIQUE, " &
                "Email NVARCHAR(200) NOT NULL DEFAULT '', " &
                "PasswordHash NVARCHAR(300) NOT NULL, " &
                "Created NVARCHAR(20) NOT NULL, " &
                "LastLogin NVARCHAR(30) NOT NULL DEFAULT '-', " &
                "Status NVARCHAR(20) NOT NULL DEFAULT 'Active')")

            ' Adds the Role column; existing accounts become Admin
            RunTx(c, Nothing,
                "IF COL_LENGTH('dbo.Users','Role') IS NULL " &
                "ALTER TABLE dbo.Users ADD Role NVARCHAR(20) NOT NULL DEFAULT 'Admin'")
            RunTx(c, Nothing,
                "IF OBJECT_ID('dbo.Books','U') IS NULL CREATE TABLE dbo.Books (" &
                "Id INT IDENTITY(1,1) PRIMARY KEY, " &
                "Title NVARCHAR(300) NOT NULL, " &
                "Author NVARCHAR(200) NOT NULL DEFAULT '', " &
                "Category NVARCHAR(100) NOT NULL DEFAULT '', " &
                "ISBN NVARCHAR(50) NOT NULL DEFAULT '', " &
                "Total INT NOT NULL DEFAULT 0, " &
                "Available INT NOT NULL DEFAULT 0, " &
                "Publisher NVARCHAR(200) NOT NULL DEFAULT '', " &
                "PubYear INT NOT NULL DEFAULT 0, " &
                "AddedDate NVARCHAR(20) NOT NULL DEFAULT '')")

            RunTx(c, Nothing,
                "IF OBJECT_ID('dbo.Members','U') IS NULL CREATE TABLE dbo.Members (" &
                "Id INT PRIMARY KEY, " &
                "Name NVARCHAR(200) NOT NULL, " &
                "Email NVARCHAR(200) NOT NULL DEFAULT '', " &
                "Phone NVARCHAR(50) NOT NULL DEFAULT '', " &
                "Address NVARCHAR(300) NOT NULL DEFAULT '', " &
                "Joined NVARCHAR(20) NOT NULL DEFAULT '', " &
                "Status NVARCHAR(20) NOT NULL DEFAULT 'Active')")

            RunTx(c, Nothing,
                "IF OBJECT_ID('dbo.Issues','U') IS NULL CREATE TABLE dbo.Issues (" &
                "Id INT PRIMARY KEY, " &
                "BookTitle NVARCHAR(300) NOT NULL, " &
                "MemberName NVARCHAR(200) NOT NULL, " &
                "IssueDate NVARCHAR(20) NOT NULL, " &
                "DueDate NVARCHAR(20) NOT NULL, " &
                "Status NVARCHAR(20) NOT NULL, " &
                "IssuedBy NVARCHAR(100) NOT NULL DEFAULT '')")

            ' First run: create the default admin account
            Dim count As Integer
            Using cmd As New SqlCommand("SELECT COUNT(*) FROM Users", c)
                count = CInt(cmd.ExecuteScalar())
            End Using
            If count = 0 Then
                RunTx(c, Nothing,
                      "INSERT INTO Users (Username, Email, PasswordHash, Created, Role) VALUES (@u, @e, @p, @d, 'Admin')",
                      "@u", "admin", "@e", "admin@library.com",
                      "@p", HashPassword("admin123"), "@d", Date.Today.ToString("yyyy-MM-dd"))
            End If
        End Using
    End Sub

    ' ---------- password hashing (passwords are never stored as plain text) ----------
    Public Function HashPassword(password As String) As String
        Dim salt As Byte() = RandomNumberGenerator.GetBytes(16)
        Dim hash As Byte() = Rfc2898DeriveBytes.Pbkdf2(password, salt, 100000, HashAlgorithmName.SHA256, 32)
        Return Convert.ToBase64String(salt) & ":" & Convert.ToBase64String(hash)
    End Function

    Public Function VerifyPassword(password As String, stored As String) As Boolean
        Dim parts = stored.Split(":"c)
        If parts.Length <> 2 Then Return False
        Dim salt As Byte() = Convert.FromBase64String(parts(0))
        Dim expected As Byte() = Convert.FromBase64String(parts(1))
        Dim actual As Byte() = Rfc2898DeriveBytes.Pbkdf2(password, salt, 100000, HashAlgorithmName.SHA256, 32)
        Return CryptographicOperations.FixedTimeEquals(expected, actual)
    End Function

End Module
