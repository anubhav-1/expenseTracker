Imports System.Data.OleDb
Imports System.Windows.Forms

Public Class Login
    Inherits Form

    Private txtUsername As TextBox
    Private txtPassword As TextBox
    Private btnLogin As Button
    Private connectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\ExpenseTracker.accdb;"

    Public Sub New()
        Me.Text = "Login"
        Me.Size = New Size(400, 300)
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.BackColor = Color.FromArgb(34, 40, 49)
        Me.FormBorderStyle = FormBorderStyle.FixedSingle
        Me.MaximizeBox = False

        InitializeComponents()
    End Sub

    Private Sub InitializeComponents()
        Dim pnlLogin As New Panel()
        pnlLogin.Size = New Size(350, 200)
        pnlLogin.Location = New Point((Me.Width - pnlLogin.Width) \ 2, (Me.Height - pnlLogin.Height) \ 2)
        pnlLogin.BackColor = Color.FromArgb(45, 52, 64)
        pnlLogin.BorderStyle = BorderStyle.FixedSingle
        Me.Controls.Add(pnlLogin)

        ' Username controls
        Dim lblUsername As New Label()
        lblUsername.Text = "Username:"
        lblUsername.Location = New Point(20, 20)
        lblUsername.AutoSize = True
        lblUsername.ForeColor = Color.White
        lblUsername.Font = New Font("Segoe UI", 12, FontStyle.Bold)
        pnlLogin.Controls.Add(lblUsername)

        txtUsername = New TextBox()
        txtUsername.Location = New Point(120, 20)
        txtUsername.Size = New Size(200, 30)
        txtUsername.BackColor = Color.FromArgb(57, 62, 70)
        txtUsername.ForeColor = Color.White
        txtUsername.Font = New Font("Segoe UI", 12)
        txtUsername.BorderStyle = BorderStyle.FixedSingle
        pnlLogin.Controls.Add(txtUsername)

        ' Password controls
        Dim lblPassword As New Label()
        lblPassword.Text = "Password:"
        lblPassword.Location = New Point(20, 70)
        lblPassword.AutoSize = True
        lblPassword.ForeColor = Color.White
        lblPassword.Font = New Font("Segoe UI", 12, FontStyle.Bold)
        pnlLogin.Controls.Add(lblPassword)

        txtPassword = New TextBox()
        txtPassword.Location = New Point(120, 70)
        txtPassword.Size = New Size(200, 30)
        txtPassword.BackColor = Color.FromArgb(57, 62, 70)
        txtPassword.ForeColor = Color.White
        txtPassword.Font = New Font("Segoe UI", 12)
        txtPassword.BorderStyle = BorderStyle.FixedSingle
        txtPassword.PasswordChar = "*"c
        pnlLogin.Controls.Add(txtPassword)

        ' Login button
        btnLogin = New Button()
        btnLogin.Text = "Login"
        btnLogin.Location = New Point(120, 120)
        btnLogin.Size = New Size(100, 40)
        btnLogin.FlatStyle = FlatStyle.Flat
        btnLogin.FlatAppearance.BorderSize = 0
        btnLogin.BackColor = Color.FromArgb(0, 173, 181)
        btnLogin.ForeColor = Color.White
        btnLogin.Font = New Font("Segoe UI", 12, FontStyle.Bold)
        btnLogin.Cursor = Cursors.Hand
        AddHandler btnLogin.Click, AddressOf btnLogin_Click
        pnlLogin.Controls.Add(btnLogin)
    End Sub

    Private Sub btnLogin_Click(sender As Object, e As EventArgs)
        Try
            If ValidateUser(txtUsername.Text, txtPassword.Text) Then
                Dim dashboard As New Dashboard()
                dashboard.Show()
                Me.Hide()
            Else
                MessageBox.Show("Invalid username or password", "Login Failed", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        Catch ex As Exception
            MessageBox.Show($"Login error: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Function ValidateUser(username As String, password As String) As Boolean
        Using connection As New OleDbConnection(connectionString)
            connection.Open()
            Dim query = "SELECT COUNT(*) FROM Users WHERE Username = @Username AND Password = @Password"
            Using command As New OleDbCommand(query, connection)
                command.Parameters.AddWithValue("@Username", username)
                command.Parameters.AddWithValue("@Password", password)
                Return Convert.ToInt32(command.ExecuteScalar()) > 0
            End Using
        End Using
    End Function
End Class