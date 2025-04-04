Imports System.Data.OleDb
Imports System.Windows.Forms

Public Class Login
    Inherits Form

    Private txtUsername As TextBox
    Private txtPassword As TextBox
    Private btnLogin As Button
    Private btnRegister As Button  ' New register button
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
        AddHandler btnLogin.MouseEnter, AddressOf Button_MouseEnter
        AddHandler btnLogin.MouseLeave, AddressOf Button_MouseLeave
        pnlLogin.Controls.Add(btnLogin)

        ' Register button (new)
        btnRegister = New Button()
        btnRegister.Text = "Register"
        btnRegister.Location = New Point(230, 120)
        btnRegister.Size = New Size(100, 40)
        btnRegister.FlatStyle = FlatStyle.Flat
        btnRegister.FlatAppearance.BorderSize = 0
        btnRegister.BackColor = Color.FromArgb(76, 187, 23)  ' Green color for register
        btnRegister.ForeColor = Color.White
        btnRegister.Font = New Font("Segoe UI", 12, FontStyle.Bold)
        btnRegister.Cursor = Cursors.Hand
        AddHandler btnRegister.Click, AddressOf btnRegister_Click
        AddHandler btnRegister.MouseEnter, AddressOf RegisterButton_MouseEnter
        AddHandler btnRegister.MouseLeave, AddressOf RegisterButton_MouseLeave
        pnlLogin.Controls.Add(btnRegister)
    End Sub

    Private Sub btnLogin_Click(sender As Object, e As EventArgs)
        Try
            If ValidateUser(txtUsername.Text, txtPassword.Text) Then
                ' Pass reference to this form to Dashboard
                Dim dashboard As New Dashboard(Me)
                dashboard.Show()
                Me.Hide()
            Else
                MessageBox.Show("Invalid username or password", "Login Failed", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        Catch ex As Exception
            MessageBox.Show($"Login error: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' New Register button handler
    ' New Register button handler
    Private Sub btnRegister_Click(sender As Object, e As EventArgs)
        ' Check if fields are filled
        If String.IsNullOrWhiteSpace(txtUsername.Text) OrElse String.IsNullOrWhiteSpace(txtPassword.Text) Then
            MessageBox.Show("Please enter both username and password", "Registration Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        ' Check if username already exists
        If UsernameExists(txtUsername.Text) Then
            MessageBox.Show("Username already exists. Please choose a different username.", "Registration Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        ' Register the new user
        Try
            Using connection As New OleDbConnection(connectionString)
                connection.Open()
                ' Fixed syntax for MS Access - use square brackets for field names
                Dim query As String = "INSERT INTO Users ([Username], [Password]) VALUES (?, ?)"

                Using command As New OleDbCommand(query, connection)
                    ' Use Add method without parameter names for OleDb
                    command.Parameters.Add(New OleDbParameter("", txtUsername.Text))
                    command.Parameters.Add(New OleDbParameter("", txtPassword.Text))

                    Dim rowsAffected = command.ExecuteNonQuery()

                    If rowsAffected > 0 Then
                        MessageBox.Show("Registration successful! You can now log in.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        ' Clear fields
                        txtPassword.Clear()
                    Else
                        MessageBox.Show("Registration failed. Please try again.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End If
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show($"Registration error: {ex.Message}", "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' Check if username already exists
    Private Function UsernameExists(username As String) As Boolean
        Try
            Using connection As New OleDbConnection(connectionString)
                connection.Open()
                Dim query = "SELECT COUNT(*) FROM Users WHERE Username = ?"

                Using command As New OleDbCommand(query, connection)
                    command.Parameters.AddWithValue("@Username", username)
                    Dim count = Convert.ToInt32(command.ExecuteScalar())
                    Return count > 0
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show($"Database error checking username: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End Try
    End Function

    Private Function ValidateUser(username As String, password As String) As Boolean
        ' Check for empty inputs
        If String.IsNullOrWhiteSpace(username) OrElse String.IsNullOrWhiteSpace(password) Then
            Return False
        End If

        Try
            Using connection As New OleDbConnection(connectionString)
                connection.Open()
                Dim query = "SELECT COUNT(*) FROM Users WHERE Username = ? AND Password = ?"
                Using command As New OleDbCommand(query, connection)
                    command.Parameters.AddWithValue("@Username", username)
                    command.Parameters.AddWithValue("@Password", password)
                    Return Convert.ToInt32(command.ExecuteScalar()) > 0
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show($"Database error: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End Try
    End Function

    ' Button hover effects
    Private Sub Button_MouseEnter(sender As Object, e As EventArgs)
        Dim btn = CType(sender, Button)
        btn.BackColor = Color.FromArgb(0, 150, 160)
    End Sub

    Private Sub Button_MouseLeave(sender As Object, e As EventArgs)
        Dim btn = CType(sender, Button)
        btn.BackColor = Color.FromArgb(0, 173, 181)
    End Sub

    ' Register button hover effects
    Private Sub RegisterButton_MouseEnter(sender As Object, e As EventArgs)
        Dim btn = CType(sender, Button)
        btn.BackColor = Color.FromArgb(60, 160, 20)  ' Darker green on hover
    End Sub

    Private Sub RegisterButton_MouseLeave(sender As Object, e As EventArgs)
        Dim btn = CType(sender, Button)
        btn.BackColor = Color.FromArgb(76, 187, 23)  ' Original green
    End Sub
End Class