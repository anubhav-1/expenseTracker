Imports System.Windows.Forms
Imports System.Drawing
Imports System.Diagnostics

Partial Public Class ReportsAnalytics
    Inherits Form

    ' Database connection
    Private connectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\ExpenseTracker.accdb;Persist Security Info=False;"

    ' Report generator
    Private incomeExpenseReportGenerator As IncomeExpenseReportGenerator

    ' Track if a report has been generated
    Private reportGenerated As Boolean = False

    Public Sub New()
        ' This call is required by the designer
        InitializeComponent()

        ' Additional initialization
        InitializeReportGenerator()
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
    Private Sub OnTimeFrameChanged(sender As Object, e As EventArgs) Handles cmbTimeFrame.SelectedIndexChanged
        ' Show/hide custom time frame controls based on selection
        If cmbTimeFrame IsNot Nothing AndAlso lblYear IsNot Nothing AndAlso
           cmbYear IsNot Nothing AndAlso lblMonth IsNot Nothing AndAlso cmbMonth IsNot Nothing Then

            Dim isCustom As Boolean = cmbTimeFrame.SelectedItem.ToString() = "Custom"

            lblYear.Visible = isCustom
            cmbYear.Visible = isCustom
            lblMonth.Visible = isCustom
            cmbMonth.Visible = isCustom
        End If
    End Sub

    Private Sub OnGenerateReport(sender As Object, e As EventArgs) Handles btnGenerateReport.Click
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
    Private Sub OnButtonMouseEnter(sender As Object, e As EventArgs) Handles btnGenerateReport.MouseEnter
        Dim btn As Button = CType(sender, Button)
        btn.BackColor = Color.FromArgb(0, 150, 160)
    End Sub

    Private Sub OnButtonMouseLeave(sender As Object, e As EventArgs) Handles btnGenerateReport.MouseLeave
        Dim btn As Button = CType(sender, Button)
        btn.BackColor = Color.FromArgb(0, 173, 181)
    End Sub
End Class