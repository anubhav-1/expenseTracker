Imports System.Windows.Forms
Imports System.Drawing

Partial Public Class ReportsAnalytics
    ' Control declarations
    Friend WithEvents pnlFilters As Panel
    Friend WithEvents lblReportType As Label
    Friend WithEvents lblTimeFrame As Label
    Friend WithEvents lblYear As Label
    Friend WithEvents lblMonth As Label
    Friend WithEvents cmbTimeFrame As ComboBox
    Friend WithEvents cmbYear As ComboBox
    Friend WithEvents cmbMonth As ComboBox
    Friend WithEvents btnGenerateReport As Button
    Friend WithEvents pnlChartArea As Panel
    Friend WithEvents pnlDataGrid As Panel
    Friend WithEvents dgvReportData As DataGridView
    Friend WithEvents pnlPieChart As Panel
    Friend WithEvents pnlBarChart As Panel

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    Private Sub InitializeComponent()
        ' Form setup
        Me.Text = "Income vs Expenses Report"
        Me.Size = New Size(1000, 800)
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.BackColor = Color.FromArgb(34, 40, 49)
        Me.FormBorderStyle = FormBorderStyle.None
        Me.Dock = DockStyle.Fill

        ' Top panel for filters
        pnlFilters = New Panel()
        pnlFilters.Dock = DockStyle.Top
        pnlFilters.Height = 100
        pnlFilters.BackColor = Color.FromArgb(45, 52, 64)
        pnlFilters.Padding = New Padding(10)
        Me.Controls.Add(pnlFilters)

        ' Report Type label - static text now
        lblReportType = New Label()
        lblReportType.Text = "Report Type: Income vs Expenses"
        lblReportType.ForeColor = Color.White
        lblReportType.Font = New Font("Segoe UI", 12, FontStyle.Bold)
        lblReportType.Location = New Point(20, 15)
        lblReportType.AutoSize = True
        pnlFilters.Controls.Add(lblReportType)

        ' Time Frame selector
        lblTimeFrame = New Label()
        lblTimeFrame.Text = "Time Frame:"
        lblTimeFrame.ForeColor = Color.White
        lblTimeFrame.Font = New Font("Segoe UI", 10, FontStyle.Bold)
        lblTimeFrame.Location = New Point(320, 15)
        lblTimeFrame.AutoSize = True
        pnlFilters.Controls.Add(lblTimeFrame)

        cmbTimeFrame = New ComboBox()
        cmbTimeFrame.Location = New Point(420, 12)
        cmbTimeFrame.Size = New Size(150, 28)
        cmbTimeFrame.BackColor = Color.FromArgb(57, 62, 70)
        cmbTimeFrame.ForeColor = Color.White
        cmbTimeFrame.DropDownStyle = ComboBoxStyle.DropDownList
        cmbTimeFrame.Font = New Font("Segoe UI", 10)
        cmbTimeFrame.Items.AddRange(New Object() {"Current Month", "Last 3 Months", "Last 6 Months", "This Year", "Custom"})
        cmbTimeFrame.SelectedIndex = 0
        pnlFilters.Controls.Add(cmbTimeFrame)

        ' Year selector (for custom time frame)
        lblYear = New Label()
        lblYear.Text = "Year:"
        lblYear.ForeColor = Color.White
        lblYear.Font = New Font("Segoe UI", 10, FontStyle.Bold)
        lblYear.Location = New Point(20, 55)
        lblYear.AutoSize = True
        lblYear.Visible = False
        pnlFilters.Controls.Add(lblYear)

        cmbYear = New ComboBox()
        cmbYear.Location = New Point(120, 52)
        cmbYear.Size = New Size(100, 28)
        cmbYear.BackColor = Color.FromArgb(57, 62, 70)
        cmbYear.ForeColor = Color.White
        cmbYear.DropDownStyle = ComboBoxStyle.DropDownList
        cmbYear.Font = New Font("Segoe UI", 10)
        cmbYear.Visible = False
        ' Add years
        Dim currentYear As Integer = DateTime.Now.Year
        For i As Integer = 0 To 5
            cmbYear.Items.Add(currentYear - i)
        Next
        cmbYear.SelectedIndex = 0
        pnlFilters.Controls.Add(cmbYear)

        ' Month selector (for custom time frame)
        lblMonth = New Label()
        lblMonth.Text = "Month:"
        lblMonth.ForeColor = Color.White
        lblMonth.Font = New Font("Segoe UI", 10, FontStyle.Bold)
        lblMonth.Location = New Point(240, 55)
        lblMonth.AutoSize = True
        lblMonth.Visible = False
        pnlFilters.Controls.Add(lblMonth)

        cmbMonth = New ComboBox()
        cmbMonth.Location = New Point(320, 52)
        cmbMonth.Size = New Size(120, 28)
        cmbMonth.BackColor = Color.FromArgb(57, 62, 70)
        cmbMonth.ForeColor = Color.White
        cmbMonth.DropDownStyle = ComboBoxStyle.DropDownList
        cmbMonth.Font = New Font("Segoe UI", 10)
        cmbMonth.Visible = False
        ' Add months
        cmbMonth.Items.AddRange(New Object() {"January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"})
        cmbMonth.SelectedIndex = DateTime.Now.Month - 1
        pnlFilters.Controls.Add(cmbMonth)

        ' Generate Report button
        btnGenerateReport = New Button()
        btnGenerateReport.Text = "Generate Report"
        btnGenerateReport.Location = New Point(600, 30)
        btnGenerateReport.Size = New Size(150, 40)
        btnGenerateReport.FlatStyle = FlatStyle.Flat
        btnGenerateReport.FlatAppearance.BorderSize = 0
        btnGenerateReport.BackColor = Color.FromArgb(0, 173, 181)
        btnGenerateReport.ForeColor = Color.White
        btnGenerateReport.Font = New Font("Segoe UI", 11, FontStyle.Bold)
        btnGenerateReport.Cursor = Cursors.Hand
        pnlFilters.Controls.Add(btnGenerateReport)

        ' Chart area
        pnlChartArea = New Panel()
        pnlChartArea.Dock = DockStyle.Top
        pnlChartArea.Height = 400
        pnlChartArea.BackColor = Color.FromArgb(57, 62, 70)
        Me.Controls.Add(pnlChartArea)

        ' Pie chart panel
        pnlPieChart = New Panel()
        pnlPieChart.Size = New Size(450, 380)
        pnlPieChart.Location = New Point(20, 10)
        pnlPieChart.BackColor = Color.FromArgb(45, 52, 64)
        pnlChartArea.Controls.Add(pnlPieChart)

        ' Bar chart panel
        pnlBarChart = New Panel()
        pnlBarChart.Size = New Size(450, 380)
        pnlBarChart.Location = New Point(490, 10)
        pnlBarChart.BackColor = Color.FromArgb(45, 52, 64)
        pnlChartArea.Controls.Add(pnlBarChart)

        ' Data grid area
        pnlDataGrid = New Panel()
        pnlDataGrid.Dock = DockStyle.Fill
        pnlDataGrid.BackColor = Color.FromArgb(57, 62, 70)
        pnlDataGrid.Padding = New Padding(20)
        Me.Controls.Add(pnlDataGrid)

        ' Initialize DataGridView
        dgvReportData = New DataGridView()
        dgvReportData.Dock = DockStyle.Fill
        dgvReportData.BackgroundColor = Color.FromArgb(45, 52, 64)
        dgvReportData.ForeColor = Color.White
        dgvReportData.BorderStyle = BorderStyle.None
        dgvReportData.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.None
        dgvReportData.ColumnHeadersDefaultCellStyle.BackColor = Color.Black
        dgvReportData.ColumnHeadersDefaultCellStyle.ForeColor = Color.White
        dgvReportData.ColumnHeadersDefaultCellStyle.Font = New Font("Segoe UI", 12, FontStyle.Bold)
        dgvReportData.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        dgvReportData.ColumnHeadersHeight = 40
        dgvReportData.EnableHeadersVisualStyles = False
        dgvReportData.DefaultCellStyle.BackColor = Color.FromArgb(57, 62, 70)
        dgvReportData.DefaultCellStyle.ForeColor = Color.White
        dgvReportData.DefaultCellStyle.Font = New Font("Segoe UI", 11)
        dgvReportData.DefaultCellStyle.SelectionBackColor = Color.FromArgb(0, 150, 160)
        dgvReportData.DefaultCellStyle.SelectionForeColor = Color.White
        dgvReportData.RowHeadersVisible = False
        dgvReportData.RowTemplate.Height = 35
        dgvReportData.RowTemplate.DefaultCellStyle.Padding = New Padding(5, 0, 0, 0)
        dgvReportData.GridColor = Color.FromArgb(34, 40, 49)
        dgvReportData.AllowUserToAddRows = False
        dgvReportData.AllowUserToResizeRows = False
        dgvReportData.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        dgvReportData.MultiSelect = False
        dgvReportData.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        dgvReportData.ReadOnly = True

        ' Apply double buffering to reduce flicker
        Try
            Dim dgvType As Type = dgvReportData.GetType()
            Dim pi As Reflection.PropertyInfo = dgvType.GetProperty("DoubleBuffered", Reflection.BindingFlags.Instance Or Reflection.BindingFlags.NonPublic)
            pi.SetValue(dgvReportData, True, Nothing)
        Catch ex As Exception
            Debug.WriteLine("Failed to apply double buffering: " & ex.Message)
        End Try

        pnlDataGrid.Controls.Add(dgvReportData)
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