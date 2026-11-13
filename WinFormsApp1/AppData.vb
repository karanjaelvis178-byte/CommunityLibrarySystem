Imports Microsoft.Data.SqlClient

Public Class UserAccount
    Public Property Id As Integer
    Public Property Username As String = ""
    Public Property Email As String = ""
    Public Property PasswordHash As String = ""
    Public Property Created As String = ""
    Public Property LastLogin As String = "-"
    Public Property Status As String = "Active"
    Public Property Role As String = "Admin"
End Class

Public Module AppData

    Private Const UserColumns As String = "Id, Username, Email, PasswordHash, Created, LastLogin, Status, Role"

    Private Function ReadUser(r As SqlDataReader) As UserAccount
        Return New UserAccount With {
            .Id = r.GetInt32(0),
            .Username = r.GetString(1),
            .Email = r.GetString(2),
            .PasswordHash = r.GetString(3),
            .Created = r.GetString(4),
            .LastLogin = r.GetString(5),
            .Status = r.GetString(6),
            .Role = r.GetString(7)
        }
    End Function

    Public Function GetUsers() As List(Of UserAccount)
        Dim list As New List(Of UserAccount)
        Using c As SqlConnection = Db.Open()
            Using cmd As New SqlCommand("SELECT " & UserColumns & " FROM Users ORDER BY Id", c)
                Using r As SqlDataReader = cmd.ExecuteReader()
                    While r.Read()
                        list.Add(ReadUser(r))
                    End While
                End Using
            End Using
        End Using
        Return list
    End Function

    Public Function FindUser(username As String) As UserAccount
        Using c As SqlConnection = Db.Open()
            Using cmd As New SqlCommand("SELECT " & UserColumns & " FROM Users WHERE Username = @u", c)
                cmd.Parameters.AddWithValue("@u", username.Trim())
                Using r As SqlDataReader = cmd.ExecuteReader()
                    If r.Read() Then Return ReadUser(r)
                End Using
            End Using
        End Using
        Return Nothing
    End Function

    ' Returns False if the username already exists. Role is "Admin" or "Librarian".
    Public Function AddUser(username As String, email As String, password As String,
                            Optional role As String = "Admin") As Boolean
        Try
            Db.Run("INSERT INTO Users (Username, Email, PasswordHash, Created, Role) VALUES (@u, @e, @p, @d, @r)",
                   "@u", username.Trim(), "@e", email.Trim(),
                   "@p", Db.HashPassword(password), "@d", Date.Today.ToString("yyyy-MM-dd"),
                   "@r", role)
            Return True
        Catch ex As SqlException When ex.Number = 2627 OrElse ex.Number = 2601
            Return False
        End Try
    End Function

    Public Sub SetPassword(id As Integer, newPassword As String)
        Db.Run("UPDATE Users SET PasswordHash = @p WHERE Id = @i", "@p", Db.HashPassword(newPassword), "@i", id)
    End Sub

    Public Sub SetStatus(id As Integer, status As String)
        Db.Run("UPDATE Users SET Status = @s WHERE Id = @i", "@s", status, "@i", id)
    End Sub

    Public Sub DeleteUser(id As Integer)
        Db.Run("DELETE FROM Users WHERE Id = @i", "@i", id)
    End Sub

    Public Sub TouchLogin(id As Integer)
        Db.Run("UPDATE Users SET LastLogin = @t WHERE Id = @i",
               "@t", DateTime.Now.ToString("yyyy-MM-dd HH:mm"), "@i", id)
    End Sub

End Module