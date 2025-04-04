Imports System.Data.OleDb
Imports System.Windows.Forms
Imports System.Drawing

Partial Public Class Login
    Inherits Form

    ' Database connection
    Private connectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\ExpenseTracker.accdb;"

    ' Effects and Animation
    Private currentOpacity As Double = 0.0

    ' Form dragging support
    Private isMouseDown As Boolean = False
    Private mouseOffset As Point

    Public Sub New()
        ' This call is required by the designer
        InitializeComponent()

        ' Add additional initialization after the InitializeComponent() call
        Me.Opacity = 0  ' Start invisible for fade-in effect
        InitializeAnimation()
    End Sub

    ' Initialize fade-in animation
    Private Sub InitializeAnimation()
        tmrFade.Interval = 15
        tmrFade.Start()
    End Sub

    ' Fade-in timer event
    Private Sub tmrFade_Tick(sender As Object, e As EventArgs) Handles tmrFade.Tick
        currentOpacity += 0.05
        If currentOpacity >= 1 Then
            currentOpacity = 1
            tmrFade.Stop()
        End If
        Me.Opacity = currentOpacity
    End Sub

    Private Sub btnLogin_Click(sender As Object, e As EventArgs) Handles btnLogin.Click
        Try
            If ValidateUser(txtUsername.Text, txtPassword.Text) Then
                ' Create dashboard with reference to this form
                Dim dashboard As New Dashboard(Me)
                dashboard.Show()
                Me.Hide()
            Else
                ShowErrorMessage("Invalid username or password")
            End If
        Catch ex As Exception
            ShowErrorMessage($"Login error: {ex.Message}")
        End Try
    End Sub

    Private Sub btnRegister_Click(sender As Object, e As EventArgs) Handles btnRegister.Click
        ' Check if fields are filled
        If String.IsNullOrWhiteSpace(txtUsername.Text) OrElse String.IsNullOrWhiteSpace(txtPassword.Text) Then
            ShowErrorMessage("Please enter both username and password")
            Return
        End If

        ' Check if username already exists
        If UsernameExists(txtUsername.Text) Then
            ShowErrorMessage("Username already exists. Please choose a different username.")
            Return
        End If

        ' Register the new user
        Try
            Using connection As New OleDbConnection(connectionString)
                connection.Open()
                Dim query As String = "INSERT INTO Users ([Username], [Password]) VALUES (?, ?)"

                Using command As New OleDbCommand(query, connection)
                    command.Parameters.Add(New OleDbParameter("", txtUsername.Text))
                    command.Parameters.Add(New OleDbParameter("", txtPassword.Text))

                    Dim rowsAffected = command.ExecuteNonQuery()

                    If rowsAffected > 0 Then
                        ShowSuccessMessage("Registration successful! You can now log in.")
                        txtPassword.Clear()
                    Else
                        ShowErrorMessage("Registration failed. Please try again.")
                    End If
                End Using
            End Using
        Catch ex As Exception
            ShowErrorMessage($"Registration error: {ex.Message}")
        End Try
    End Sub

    ' Custom message box for errors
    Private Sub ShowErrorMessage(message As String)
        Dim frmError As New Form()
        frmError.Size = New Size(400, 200)
        frmError.StartPosition = FormStartPosition.CenterParent
        frmError.FormBorderStyle = FormBorderStyle.None
        frmError.BackColor = Color.FromArgb(45, 52, 64)

        Dim lblMessage As New Label()
        lblMessage.Text = message
        lblMessage.ForeColor = Color.White
        lblMessage.Font = New Font("Segoe UI", 11)
        lblMessage.TextAlign = ContentAlignment.MiddleCenter
        lblMessage.Dock = DockStyle.Fill
        frmError.Controls.Add(lblMessage)

        Dim btnOk As New Button()
        btnOk.Text = "OK"
        btnOk.Size = New Size(100, 40)
        btnOk.Location = New Point((frmError.Width - btnOk.Width) \ 2, frmError.Height - 60)
        btnOk.FlatStyle = FlatStyle.Flat
        btnOk.BackColor = Color.FromArgb(255, 77, 77)
        btnOk.ForeColor = Color.White
        btnOk.FlatAppearance.BorderSize = 0
        btnOk.Font = New Font("Segoe UI", 10, FontStyle.Bold)
        AddHandler btnOk.Click, Sub(s, e) frmError.Close()
        frmError.Controls.Add(btnOk)

        frmError.ShowDialog()
    End Sub

    ' Custom message box for success
    Private Sub ShowSuccessMessage(message As String)
        Dim frmSuccess As New Form()
        frmSuccess.Size = New Size(400, 200)
        frmSuccess.StartPosition = FormStartPosition.CenterParent
        frmSuccess.FormBorderStyle = FormBorderStyle.None
        frmSuccess.BackColor = Color.FromArgb(45, 52, 64)

        Dim lblMessage As New Label()
        lblMessage.Text = message
        lblMessage.ForeColor = Color.White
        lblMessage.Font = New Font("Segoe UI", 11)
        lblMessage.TextAlign = ContentAlignment.MiddleCenter
        lblMessage.Dock = DockStyle.Fill
        frmSuccess.Controls.Add(lblMessage)

        Dim btnOk As New Button()
        btnOk.Text = "OK"
        btnOk.Size = New Size(100, 40)
        btnOk.Location = New Point((frmSuccess.Width - btnOk.Width) \ 2, frmSuccess.Height - 60)
        btnOk.FlatStyle = FlatStyle.Flat
        btnOk.BackColor = Color.FromArgb(0, 173, 181)
        btnOk.ForeColor = Color.White
        btnOk.FlatAppearance.BorderSize = 0
        btnOk.Font = New Font("Segoe UI", 10, FontStyle.Bold)
        AddHandler btnOk.Click, Sub(s, e) frmSuccess.Close()
        frmSuccess.Controls.Add(btnOk)

        frmSuccess.ShowDialog()
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
            ShowErrorMessage($"Database error checking username: {ex.Message}")
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
            ShowErrorMessage($"Database error: {ex.Message}")
            Return False
        End Try
    End Function

    ' Event handlers for buttons
    Private Sub Button_MouseEnter(sender As Object, e As EventArgs) Handles btnLogin.MouseEnter
        Dim btn = CType(sender, Button)
        btn.BackColor = Color.FromArgb(0, 150, 160)
    End Sub

    Private Sub Button_MouseLeave(sender As Object, e As EventArgs) Handles btnLogin.MouseLeave
        Dim btn = CType(sender, Button)
        btn.BackColor = Color.FromArgb(0, 173, 181)
    End Sub

    ' Register button hover effects
    Private Sub RegisterButton_MouseEnter(sender As Object, e As EventArgs) Handles btnRegister.MouseEnter
        Dim btn = CType(sender, Button)
        btn.BackColor = Color.FromArgb(45, 52, 64)
        btn.ForeColor = Color.White
    End Sub

    Private Sub RegisterButton_MouseLeave(sender As Object, e As EventArgs) Handles btnRegister.MouseLeave
        Dim btn = CType(sender, Button)
        btn.BackColor = Color.Transparent
        btn.ForeColor = Color.FromArgb(0, 173, 181)
    End Sub

    ' Close button handlers
    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub btnClose_MouseEnter(sender As Object, e As EventArgs) Handles btnClose.MouseEnter
        btnClose.ForeColor = Color.FromArgb(0, 173, 181)
    End Sub

    Private Sub btnClose_MouseLeave(sender As Object, e As EventArgs) Handles btnClose.MouseLeave
        btnClose.ForeColor = Color.White
    End Sub

    ' Make form draggable
    Protected Overrides Sub OnMouseDown(e As MouseEventArgs)
        MyBase.OnMouseDown(e)
        If e.Button = MouseButtons.Left Then
            isMouseDown = True
            mouseOffset = New Point(-e.X, -e.Y)
        End If
    End Sub

    Protected Overrides Sub OnMouseMove(e As MouseEventArgs)
        MyBase.OnMouseMove(e)
        If isMouseDown Then
            Dim mousePos = Control.MousePosition
            mousePos.Offset(mouseOffset.X, mouseOffset.Y)
            Location = mousePos
        End If
    End Sub

    Protected Overrides Sub OnMouseUp(e As MouseEventArgs)
        MyBase.OnMouseUp(e)
        If e.Button = MouseButtons.Left Then
            isMouseDown = False
        End If
    End Sub
End Class