Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.Windows.Forms

Public Class ThemeHelper
    ' Color scheme constants 
    Public Shared ReadOnly PurpleColor As Color = Color.FromArgb(156, 39, 176)  ' Main purple
    Public Shared ReadOnly PurpleDarkColor As Color = Color.FromArgb(126, 29, 146)  ' Darker purple
    Public Shared ReadOnly PinkColor As Color = Color.FromArgb(233, 30, 99)  ' Pink
    Public Shared ReadOnly PinkDarkColor As Color = Color.FromArgb(203, 0, 69)  ' Darker pink
    Public Shared ReadOnly GreenColor As Color = Color.FromArgb(76, 187, 23)  ' Green for positive values
    Public Shared ReadOnly RedColor As Color = Color.FromArgb(255, 77, 77)  ' Red for negative/delete
    Public Shared ReadOnly BackgroundLightColor As Color = Color.FromArgb(240, 240, 245)  ' Light background
    Public Shared ReadOnly BackgroundDarkColor As Color = Color.FromArgb(45, 52, 64)  ' Dark background

    ' Common styling methods
    Public Shared Sub StyleDataGridView(dgv As DataGridView)
        dgv.BackgroundColor = Color.FromArgb(240, 240, 245) ' Light background
        dgv.ForeColor = Color.FromArgb(80, 80, 80)
        dgv.BorderStyle = BorderStyle.None
        dgv.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.None

        ' Headers
        dgv.ColumnHeadersDefaultCellStyle.BackColor = PurpleColor ' Purple
        dgv.ColumnHeadersDefaultCellStyle.ForeColor = Color.White
        dgv.ColumnHeadersDefaultCellStyle.Font = New Font("Century Gothic", 11, FontStyle.Bold)
        dgv.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        dgv.ColumnHeadersHeight = 40
        dgv.EnableHeadersVisualStyles = False

        ' Cells
        dgv.DefaultCellStyle.BackColor = Color.White
        dgv.DefaultCellStyle.ForeColor = Color.FromArgb(80, 80, 80)
        dgv.DefaultCellStyle.Font = New Font("Century Gothic", 10)
        dgv.DefaultCellStyle.SelectionBackColor = PinkColor ' Pink for selection
        dgv.DefaultCellStyle.SelectionForeColor = Color.White

        ' Remove row headers
        dgv.RowHeadersVisible = False
        dgv.RowTemplate.Height = 35
        dgv.RowTemplate.DefaultCellStyle.Padding = New Padding(5, 0, 0, 0)

        ' Grid color
        dgv.GridColor = Color.FromArgb(230, 230, 235) ' Light grid lines

        ' Other settings
        dgv.AllowUserToAddRows = False
        dgv.AllowUserToResizeRows = False
        dgv.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        dgv.MultiSelect = False
        dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill

        ' Apply double buffering to reduce flicker
        Try
            Dim dgvType As Type = dgv.GetType()
            Dim pi As Reflection.PropertyInfo = dgvType.GetProperty("DoubleBuffered", Reflection.BindingFlags.Instance Or Reflection.BindingFlags.NonPublic)
            pi.SetValue(dgv, True, Nothing)
        Catch ex As Exception
            Debug.WriteLine("Failed to apply double buffering: " & ex.Message)
        End Try

        ' Add alternating row styling
        dgv.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(245, 245, 250)
    End Sub

    Public Shared Sub ApplyRoundedCorners(control As Control, radius As Integer)
        Dim path As New GraphicsPath()
        path.AddArc(0, 0, radius, radius, 180, 90)
        path.AddArc(control.Width - radius, 0, radius, radius, 270, 90)
        path.AddArc(control.Width - radius, control.Height - radius, radius, radius, 0, 90)
        path.AddArc(0, control.Height - radius, radius, radius, 90, 90)
        control.Region = New Region(path)
    End Sub

    Public Shared Sub AddShadowEffect(control As Control)
        AddHandler control.Paint, AddressOf PaintShadow
    End Sub

    Private Shared Sub PaintShadow(sender As Object, e As PaintEventArgs)
        Dim control As Control = DirectCast(sender, Control)
        Dim g As Graphics = e.Graphics
        Dim rect As New Rectangle(0, 0, control.Width, control.Height)
        g.SmoothingMode = SmoothingMode.AntiAlias

        ' Draw shadow
        Using shadowBrush As New SolidBrush(Color.FromArgb(30, 0, 0, 0))
            g.FillRectangle(shadowBrush, New Rectangle(5, 5, rect.Width, rect.Height))
        End Using
    End Sub

    Public Shared Sub CreatePrettyPanel(panel As Panel, Optional radius As Integer = 15)
        panel.BackColor = Color.White
        panel.BorderStyle = BorderStyle.None

        ApplyRoundedCorners(panel, radius)
        AddShadowEffect(panel)
    End Sub
End Class