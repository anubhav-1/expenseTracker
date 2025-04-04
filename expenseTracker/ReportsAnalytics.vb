Imports System.Windows.Forms
Imports System.Drawing
Imports System.Diagnostics

Public Class ReportsAnalytics
    Inherits Form

    ' Database connection
    Private connectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\ExpenseTracker.accdb;Persist Security Info=False;"

    ' Control declarations
    Private pnlFilters As Panel
    Private lblReportType As Label
    Private lblTimeFrame As Label
    Private lblYear As Label
    Private lblMonth As Label
    Private cmbTimeFrame As ComboBox
    Private cmbYear As ComboBox
    Private cmbMonth As ComboBox
    Private btnGenerateReport As Button
    Private pnlChartArea As Panel
    Private pnlDataGrid As Panel
    Private dgvReportData As DataGridView

    ' Chart canvases
    Private pnlPieChart As Panel
    Private pnlBarChart As Panel

    ' Report generator
    Private incomeExpenseReportGenerator As IncomeExpenseReportGenerator

    ' Track if a report has been generated
    Private reportGenerated As Boolean = False

    Public Sub New()
        ' Form setup
        Me.Text = "Income vs Expenses Report"
        Me.Size = New Size(1000, 800)
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.BackColor = Color.FromArgb(34, 40, 49)
        Me.FormBorderStyle = FormBorderStyle.None
        Me.Dock = DockStyle.Fill

        InitializeComponents()
        InitializeReportGenerator()
    End Sub

    Private Sub InitializeComponents()
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
        AddHandler cmbTimeFrame.SelectedIndexChanged, AddressOf OnTimeFrameChanged
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
        AddHandler btnGenerateReport.Click, AddressOf OnGenerateReport
        AddHandler btnGenerateReport.MouseEnter, AddressOf OnButtonMouseEnter
        AddHandler btnGenerateReport.MouseLeave, AddressOf OnButtonMouseLeave
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

    Private Sub InitializeReportGenerator()
        ' Initialize only the income expense report generator
        incomeExpenseReportGenerator = New IncomeExpenseReportGenerator(connectionString, dgvReportData, pnlPieChart, pnlBarChart)

        ' Add default welcome message to chart panels - this prevents automatic data loading
        AddHandler pnlPieChart.Paint, AddressOf DrawWelcomeMessage
        AddHandler pnlBarChart.Paint, AddressOf DrawWelcomeMessage
    End Sub

    ' Display welcome message instead of trying to load data automatically
    Private Sub DrawWelcomeMessage(sender As Object, e As PaintEventArgs)
        ' Only draw welcome message if no report has been generated yet
        If Not reportGenerated Then
            Using brush As New SolidBrush(Color.White)
                e.Graphics.DrawString("Select time frame, then click 'Generate Report'",
                                     New Font("Segoe UI", 11), brush, 20, 150)
            End Using
        End If
    End Sub

    ' Event handlers
    Private Sub OnTimeFrameChanged(sender As Object, e As EventArgs)
        ' Show/hide custom time frame controls based on selection
        Dim isCustom As Boolean = cmbTimeFrame.SelectedItem.ToString() = "Custom"

        lblYear.Visible = isCustom
        cmbYear.Visible = isCustom
        lblMonth.Visible = isCustom
        cmbMonth.Visible = isCustom
    End Sub

    Private Sub OnGenerateReport(sender As Object, e As EventArgs)
        ' Generate the report
        Try
            ' Get date range
            Dim dateRange As DateRange = GetDateRange()

            ' Debug info to see the exact dates being used
            Debug.WriteLine($"Generating income vs expenses report from {dateRange.StartDate} to {dateRange.EndDate}")
            Debug.WriteLine($"Formatted dates: {dateRange.GetFormattedStartDate()} to {dateRange.GetFormattedEndDate()}")

            ' Before generating a new report, remove the welcome message handlers
            If Not reportGenerated Then
                RemoveHandler pnlPieChart.Paint, AddressOf DrawWelcomeMessage
                RemoveHandler pnlBarChart.Paint, AddressOf DrawWelcomeMessage
                reportGenerated = True
            End If

            ' Generate income vs expenses report
            Debug.WriteLine("Calling incomeExpenseReportGenerator.GenerateReport...")
            incomeExpenseReportGenerator.GenerateReport(dateRange)
            Debug.WriteLine("Report generation complete")

        Catch ex As Exception
            MessageBox.Show("Error generating report: " & ex.Message, "Report Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Debug.WriteLine("Report error: " & ex.Message & vbCrLf & ex.StackTrace)
        End Try
    End Sub

    ' Helper method to get date range based on selected time frame
    Private Function GetDateRange() As DateRange
        Dim startDate As DateTime
        Dim endDate As DateTime = DateTime.Now
        Dim timeFrame As String = cmbTimeFrame.SelectedItem.ToString()

        Select Case timeFrame
            Case "Current Month"
                startDate = New DateTime(DateTime.Now.Year, DateTime.Now.Month, 1)
            Case "Last 3 Months"
                startDate = DateTime.Now.AddMonths(-3)
            Case "Last 6 Months"
                startDate = DateTime.Now.AddMonths(-6)
            Case "This Year"
                startDate = New DateTime(DateTime.Now.Year, 1, 1)
            Case "Custom"
                Dim selectedYear As Integer = Convert.ToInt32(cmbYear.SelectedItem)
                Dim selectedMonth As Integer = cmbMonth.SelectedIndex + 1
                startDate = New DateTime(selectedYear, selectedMonth, 1)

                ' For Income vs Expenses report, we want to show the whole period
                endDate = New DateTime(selectedYear, 12, 31)

            Case Else
                startDate = New DateTime(DateTime.Now.Year, DateTime.Now.Month, 1)
        End Select

        ' Log the date range for debugging
        Debug.WriteLine($"Date range: {startDate.ToString("MM/dd/yyyy")} to {endDate.ToString("MM/dd/yyyy")}")

        Return New DateRange(startDate, endDate)
    End Function

    ' Button hover effects
    Private Sub OnButtonMouseEnter(sender As Object, e As EventArgs)
        Dim btn As Button = CType(sender, Button)
        btn.BackColor = Color.FromArgb(0, 150, 160)
    End Sub

    Private Sub OnButtonMouseLeave(sender As Object, e As EventArgs)
        Dim btn As Button = CType(sender, Button)
        btn.BackColor = Color.FromArgb(0, 173, 181)
    End Sub
End Class