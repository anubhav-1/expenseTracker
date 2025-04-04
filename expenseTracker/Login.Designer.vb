Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.Windows.Forms
Imports System.Runtime.CompilerServices

' Module for extension methods
Module GraphicsExtensions
    ' Extension method for drawing rounded rectangles
    <Extension()>
    Public Sub DrawRoundedRectangle(ByVal graphics As Graphics, ByVal pen As Pen, ByVal bounds As Rectangle, ByVal radius As Integer)
        If radius = 0 Then
            graphics.DrawRectangle(pen, bounds)
            Return
        End If

        Dim diameter As Integer = radius * 2
        Dim size As Size = New Size(diameter, diameter)
        Dim arc As Rectangle = New Rectangle(bounds.Location, size)

        ' Top left arc
        graphics.DrawArc(pen, arc, 180, 90)

        ' Top right arc
        arc.X = bounds.Right - diameter
        graphics.DrawArc(pen, arc, 270, 90)

        ' Bottom right arc
        arc.Y = bounds.Bottom - diameter
        graphics.DrawArc(pen, arc, 0, 90)

        ' Bottom left arc
        arc.X = bounds.Left
        graphics.DrawArc(pen, arc, 90, 90)

        ' Draw lines connecting the arcs
        graphics.DrawLine(pen, bounds.Left + radius, bounds.Top, bounds.Right - radius, bounds.Top)
        graphics.DrawLine(pen, bounds.Right, bounds.Top + radius, bounds.Right, bounds.Bottom - radius)
        graphics.DrawLine(pen, bounds.Right - radius, bounds.Bottom, bounds.Left + radius, bounds.Bottom)
        graphics.DrawLine(pen, bounds.Left, bounds.Bottom - radius, bounds.Left, bounds.Top + radius)
    End Sub
End Module

Partial Public Class Login
    ' Form Shadow
    Private Const CS_DROPSHADOW As Integer = &H20000

    ' UI Controls
    Friend WithEvents txtUsername As TextBox
    Friend WithEvents txtPassword As TextBox
    Friend WithEvents btnLogin As Button
    Friend WithEvents btnRegister As Button
    Friend WithEvents btnClose As Button
    Friend WithEvents pnlLeft As Panel
    Friend WithEvents pnlRight As Panel
    Friend WithEvents lblTitle As Label
    Friend WithEvents lblSubtitle As Label
    Friend WithEvents pnlLogin As Panel
    Friend WithEvents picLogo As PictureBox
    Friend WithEvents tmrFade As Timer
    Private components As System.ComponentModel.IContainer

    ' Enable form shadow
    Protected Overrides ReadOnly Property CreateParams() As CreateParams
        Get
            Dim cp As CreateParams = MyBase.CreateParams
            cp.ClassStyle = cp.ClassStyle Or CS_DROPSHADOW
            Return cp
        End Get
    End Property

    'Required by the Windows Form Designer
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()

        ' Initialize the timer for fade-in effect
        Me.tmrFade = New Timer(Me.components)

        ' Set form properties
        Me.Text = "Expense Tracker - Login"
        Me.Size = New Size(900, 600)
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.BackColor = Color.FromArgb(34, 40, 49)
        Me.FormBorderStyle = FormBorderStyle.None
        Me.DoubleBuffered = True

        ' Left panel with logo and branding
        pnlLeft = New Panel()
        pnlLeft.Size = New Size(400, 600)
        pnlLeft.Location = New Point(0, 0)
        pnlLeft.BackColor = Color.FromArgb(0, 173, 181)  ' Teal accent color
        Me.Controls.Add(pnlLeft)

        ' Create circular logo panel
        picLogo = New PictureBox()
        picLogo.Size = New Size(120, 120)
        picLogo.Location = New Point((pnlLeft.Width - 120) \ 2, 100)
        picLogo.BackColor = Color.White
        AddHandler picLogo.Paint, AddressOf DrawLogo
        pnlLeft.Controls.Add(picLogo)

        ' App title
        lblTitle = New Label()
        lblTitle.Text = "EXPENSE TRACKER"
        lblTitle.Font = New Font("Segoe UI", 24, FontStyle.Bold)
        lblTitle.ForeColor = Color.White
        lblTitle.AutoSize = True
        lblTitle.Location = New Point((pnlLeft.Width - lblTitle.PreferredWidth) \ 2, 250)
        pnlLeft.Controls.Add(lblTitle)

        ' App subtitle
        lblSubtitle = New Label()
        lblSubtitle.Text = "Manage your finances with ease"
        lblSubtitle.Font = New Font("Segoe UI", 12)
        lblSubtitle.ForeColor = Color.White
        lblSubtitle.AutoSize = True
        lblSubtitle.Location = New Point((pnlLeft.Width - lblSubtitle.PreferredWidth) \ 2, 300)
        pnlLeft.Controls.Add(lblSubtitle)

        ' Add decorative elements
        AddDecorations(pnlLeft)

        ' Right panel for login form
        pnlRight = New Panel()
        pnlRight.Size = New Size(500, 600)
        pnlRight.Location = New Point(400, 0)
        pnlRight.BackColor = Color.FromArgb(34, 40, 49)  ' Dark background
        Me.Controls.Add(pnlRight)

        ' Login panel
        pnlLogin = New Panel()
        pnlLogin.Size = New Size(400, 450)  ' Increased height to accommodate button
        pnlLogin.Location = New Point(50, 75)  ' Moved up slightly
        pnlLogin.BackColor = Color.FromArgb(45, 52, 64)  ' Slightly lighter than background
        AddRoundedBorder(pnlLogin)
        pnlRight.Controls.Add(pnlLogin)

        ' Login label
        Dim lblLogin As New Label()
        lblLogin.Text = "Welcome Back"
        lblLogin.Font = New Font("Segoe UI", 20, FontStyle.Bold)
        lblLogin.ForeColor = Color.White
        lblLogin.AutoSize = True
        lblLogin.Location = New Point((pnlLogin.Width - lblLogin.PreferredWidth) \ 2, 30)
        pnlLogin.Controls.Add(lblLogin)

        ' Login subtitle
        Dim lblLoginSubtitle As New Label()
        lblLoginSubtitle.Text = "Sign in to continue"
        lblLoginSubtitle.Font = New Font("Segoe UI", 10)
        lblLoginSubtitle.ForeColor = Color.LightGray
        lblLoginSubtitle.AutoSize = True
        lblLoginSubtitle.Location = New Point((pnlLogin.Width - lblLoginSubtitle.PreferredWidth) \ 2, 70)
        pnlLogin.Controls.Add(lblLoginSubtitle)

        ' Username label
        Dim lblUsername As New Label()
        lblUsername.Text = "USERNAME"
        lblUsername.Font = New Font("Segoe UI", 9, FontStyle.Bold)
        lblUsername.ForeColor = Color.LightGray
        lblUsername.Location = New Point(50, 120)
        lblUsername.AutoSize = True
        pnlLogin.Controls.Add(lblUsername)

        ' Username textbox with icon
        txtUsername = New TextBox()
        txtUsername.Location = New Point(50, 145)
        txtUsername.Size = New Size(300, 40)
        txtUsername.BackColor = Color.FromArgb(57, 62, 70)
        txtUsername.ForeColor = Color.White
        txtUsername.Font = New Font("Segoe UI", 12)
        txtUsername.BorderStyle = BorderStyle.None
        pnlLogin.Controls.Add(txtUsername)

        ' Username underline
        Dim lineUsername As New Panel()
        lineUsername.Location = New Point(50, txtUsername.Bottom + 2)
        lineUsername.Size = New Size(300, 2)
        lineUsername.BackColor = Color.FromArgb(0, 173, 181)  ' Teal
        pnlLogin.Controls.Add(lineUsername)

        ' Password label
        Dim lblPassword As New Label()
        lblPassword.Text = "PASSWORD"
        lblPassword.Font = New Font("Segoe UI", 9, FontStyle.Bold)
        lblPassword.ForeColor = Color.LightGray
        lblPassword.Location = New Point(50, 190)
        lblPassword.AutoSize = True
        pnlLogin.Controls.Add(lblPassword)

        ' Password textbox
        txtPassword = New TextBox()
        txtPassword.Location = New Point(50, 215)
        txtPassword.Size = New Size(300, 40)
        txtPassword.BackColor = Color.FromArgb(57, 62, 70)
        txtPassword.ForeColor = Color.White
        txtPassword.Font = New Font("Segoe UI", 12)
        txtPassword.BorderStyle = BorderStyle.None
        txtPassword.PasswordChar = "●"c
        pnlLogin.Controls.Add(txtPassword)

        ' Password underline
        Dim linePassword As New Panel()
        linePassword.Location = New Point(50, txtPassword.Bottom + 2)
        linePassword.Size = New Size(300, 2)
        linePassword.BackColor = Color.FromArgb(0, 173, 181)  ' Teal
        pnlLogin.Controls.Add(linePassword)


        ' Login button
        btnLogin = New Button()
        btnLogin.Text = "LOGIN"
        btnLogin.Location = New Point(50, 290)
        btnLogin.Size = New Size(300, 45)
        btnLogin.FlatStyle = FlatStyle.Flat
        btnLogin.FlatAppearance.BorderSize = 0
        btnLogin.BackColor = Color.FromArgb(0, 173, 181)
        btnLogin.ForeColor = Color.White
        btnLogin.Font = New Font("Segoe UI", 12, FontStyle.Bold)
        btnLogin.Cursor = Cursors.Hand
        pnlLogin.Controls.Add(btnLogin)

        ' Or divider
        Dim lblOr As New Label()
        lblOr.Text = "OR"
        lblOr.Font = New Font("Segoe UI", 9)
        lblOr.ForeColor = Color.LightGray
        lblOr.AutoSize = True
        lblOr.Location = New Point((pnlLogin.Width - lblOr.PreferredWidth) \ 2, 355)
        pnlLogin.Controls.Add(lblOr)

        Dim lineLeft As New Panel()
        lineLeft.Location = New Point(60, 362)
        lineLeft.Size = New Size((pnlLogin.Width - lblOr.Width) \ 2 - 70, 1)
        lineLeft.BackColor = Color.DarkGray
        pnlLogin.Controls.Add(lineLeft)

        Dim lineRight As New Panel()
        lineRight.Location = New Point(lblOr.Right + 10, 362)
        lineRight.Size = New Size((pnlLogin.Width - lblOr.Width) \ 2 - 70, 1)
        lineRight.BackColor = Color.DarkGray
        pnlLogin.Controls.Add(lineRight)

        ' Register button
        btnRegister = New Button()
        btnRegister.Text = "CREATE NEW ACCOUNT"
        btnRegister.Location = New Point(50, 385)
        btnRegister.Size = New Size(300, 45)
        btnRegister.FlatStyle = FlatStyle.Flat
        btnRegister.FlatAppearance.BorderColor = Color.FromArgb(0, 173, 181)
        btnRegister.FlatAppearance.BorderSize = 1
        btnRegister.BackColor = Color.Transparent
        btnRegister.ForeColor = Color.FromArgb(0, 173, 181)
        btnRegister.Font = New Font("Segoe UI", 12, FontStyle.Bold)
        btnRegister.Cursor = Cursors.Hand
        pnlLogin.Controls.Add(btnRegister)

        ' Close button
        btnClose = New Button()
        btnClose.Text = "×"
        btnClose.Font = New Font("Arial", 14, FontStyle.Bold)
        btnClose.ForeColor = Color.White
        btnClose.FlatStyle = FlatStyle.Flat
        btnClose.FlatAppearance.BorderSize = 0
        btnClose.BackColor = Color.Transparent
        btnClose.Size = New Size(30, 30)
        btnClose.Location = New Point(Me.Width - 40, 10)
        btnClose.Cursor = Cursors.Hand
        Me.Controls.Add(btnClose)
    End Sub

    ' Draw logo in the circular PictureBox
    Private Sub DrawLogo(sender As Object, e As PaintEventArgs)
        Dim g As Graphics = e.Graphics
        g.SmoothingMode = SmoothingMode.AntiAlias

        ' Draw circular background
        Using brush As New SolidBrush(Color.White)
            g.FillEllipse(brush, 0, 0, picLogo.Width - 1, picLogo.Height - 1)
        End Using

        ' Draw dollar sign or custom icon
        Using font As New Font("Arial", 70, FontStyle.Bold)
            Using brush As New SolidBrush(Color.FromArgb(0, 173, 181))
                g.DrawString("$", font, brush, 35, 0)
            End Using
        End Using
    End Sub

    ' Add decorative elements
    Private Sub AddDecorations(panel As Panel)
        ' Add some decorative circles/shapes
        Dim random As New Random()
        For i As Integer = 0 To 10
            Dim size As Integer = random.Next(6, 20)
            Dim x As Integer = random.Next(30, panel.Width - 30)
            Dim y As Integer = random.Next(350, panel.Height - 30)

            Dim decoration As New Panel()
            decoration.Size = New Size(size, size)
            decoration.Location = New Point(x, y)
            decoration.BackColor = Color.FromArgb(255, 255, 255, 70) ' Semi-transparent white
            panel.Controls.Add(decoration)

            ' Make some of them circular
            If i Mod 2 = 0 Then
                AddHandler decoration.Paint, Sub(sender As Object, e As PaintEventArgs)
                                                 e.Graphics.SmoothingMode = SmoothingMode.AntiAlias
                                                 e.Graphics.FillEllipse(New SolidBrush(decoration.BackColor), 0, 0, decoration.Width - 1, decoration.Height - 1)
                                             End Sub
            End If
        Next
    End Sub

    ' Add rounded border to panel
    Private Sub AddRoundedBorder(panel As Panel)
        AddHandler panel.Paint, Sub(sender As Object, e As PaintEventArgs)
                                    Dim graphics As Graphics = e.Graphics
                                    graphics.SmoothingMode = SmoothingMode.AntiAlias

                                    ' Create a Rectangle for the panel border
                                    Dim roundRect As Rectangle = New Rectangle(0, 0, panel.Width - 1, panel.Height - 1)
                                    Dim radius As Integer = 15

                                    ' Draw shadow effect
                                    For i As Integer = 1 To 5
                                        Using shadowPen As New Pen(Color.FromArgb(10, 0, 0, 0), i)
                                            ' Fixed parameter passing
                                            graphics.DrawRoundedRectangle(shadowPen, New Rectangle(roundRect.X + i, roundRect.Y + i,
                                                                        roundRect.Width - i * 2, roundRect.Height - i * 2), radius)
                                        End Using
                                    Next

                                    ' Draw panel border - also fixed here
                                    Using pen As New Pen(Color.FromArgb(60, 0, 173, 181), 1)
                                        graphics.DrawRoundedRectangle(pen, roundRect, radius)
                                    End Using
                                End Sub
    End Sub

    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub
End Class