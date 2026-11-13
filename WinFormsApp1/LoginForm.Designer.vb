<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class LoginForm
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        lblusername = New Label()
        lblPassword = New Label()
        txtusername = New TextBox()
        txtpassword = New TextBox()
        btnregisterAdmin = New Button()
        btnlogin = New Button()
        BackgroundWorker1 = New ComponentModel.BackgroundWorker()
        btnforgotpassword = New Button()
        SuspendLayout()
        ' 
        ' lblusername
        ' 
        lblusername.AutoSize = True
        lblusername.Location = New Point(634, 142)
        lblusername.Name = "lblusername"
        lblusername.Size = New Size(95, 25)
        lblusername.TabIndex = 0
        lblusername.Text = "Username:"
        ' 
        ' lblPassword
        ' 
        lblPassword.AutoSize = True
        lblPassword.Location = New Point(638, 233)
        lblPassword.Name = "lblPassword"
        lblPassword.Size = New Size(91, 25)
        lblPassword.TabIndex = 1
        lblPassword.Text = "Password:"
        ' 
        ' txtusername
        ' 
        txtusername.Location = New Point(778, 142)
        txtusername.Name = "txtusername"
        txtusername.Size = New Size(240, 31)
        txtusername.TabIndex = 2
        ' 
        ' txtpassword
        ' 
        txtpassword.Location = New Point(778, 227)
        txtpassword.Name = "txtpassword"
        txtpassword.PasswordChar = "*"c
        txtpassword.Size = New Size(240, 31)
        txtpassword.TabIndex = 3
        ' 
        ' btnregisterAdmin
        ' 
        btnregisterAdmin.Location = New Point(655, 297)
        btnregisterAdmin.Name = "btnregisterAdmin"
        btnregisterAdmin.Size = New Size(153, 51)
        btnregisterAdmin.TabIndex = 4
        btnregisterAdmin.Text = "registerAdmin"
        btnregisterAdmin.UseVisualStyleBackColor = True
        ' 
        ' btnlogin
        ' 
        btnlogin.Location = New Point(492, 297)
        btnlogin.Name = "btnlogin"
        btnlogin.Size = New Size(112, 51)
        btnlogin.TabIndex = 5
        btnlogin.Text = "Login"
        btnlogin.UseVisualStyleBackColor = True
        ' 
        ' btnforgotpassword
        ' 
        btnforgotpassword.Location = New Point(869, 297)
        btnforgotpassword.Name = "btnforgotpassword"
        btnforgotpassword.Size = New Size(176, 51)
        btnforgotpassword.TabIndex = 6
        btnforgotpassword.Text = "Forgot Password"
        btnforgotpassword.UseVisualStyleBackColor = True
        ' 
        ' LoginForm
        ' 
        AutoScaleDimensions = New SizeF(10F, 25F)
        AutoScaleMode = AutoScaleMode.Font
        ClientSize = New Size(1202, 494)
        Controls.Add(btnforgotpassword)
        Controls.Add(btnlogin)
        Controls.Add(btnregisterAdmin)
        Controls.Add(txtpassword)
        Controls.Add(txtusername)
        Controls.Add(lblPassword)
        Controls.Add(lblusername)
        Name = "LoginForm"
        ResumeLayout(False)
        PerformLayout()
    End Sub

    Friend WithEvents lblusername As Label
    Friend WithEvents lblPassword As Label
    Friend WithEvents txtusername As TextBox
    Friend WithEvents txtpassword As TextBox
    Friend WithEvents btnregisterAdmin As Button
    Friend WithEvents btnlogin As Button
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents btnforgotpassword As Button

End Class
