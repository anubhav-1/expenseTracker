Imports System
Imports System.Windows.Forms
Imports System.Data.OleDb
Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.Diagnostics

Public Class IncomeExpenseReportGenerator
    ' Database connection
    Private connectionString As String

    ' UI Controls
    Private dgvReportData As DataGridView
    Private pnlPieChart As Panel
    Private pnlBarChart As Panel

    ' Constructor
    Public Sub New(connectionString As String, dgvReportData As DataGridView, pnlPieChart As Panel, pnlBarChart As Panel)
        Me.connectionString = connectionString
        Me.dgvReportData = dgvReportData
        Me.pnlPieChart = pnlPieChart
        Me.pnlBarChart = pnlBarChart

        ' Set up chart event handlers
        AddHandler pnlPieChart.Paint, AddressOf OnDrawPieChart
        AddHandler pnlBarChart.Paint, AddressOf OnDrawBarChart
    End Sub

    ' Main method to generate the report
    Public Sub GenerateReport(dateRange As DateRange)
        ' Set up the data grid
        SetupGrid()

        ' Load data
        LoadData(dateRange)

        ' Refresh charts
        pnlPieChart.Invalidate()
        pnlBarChart.Invalidate()
    End Sub

    ' Set up the data grid for income vs expenses report
    Private Sub SetupGrid()
        dgvReportData.Columns.Clear()
        dgvReportData.Columns.Add("Month", "Month")
        dgvReportData.Columns.Add("Income", "Income")
        dgvReportData.Columns.Add("Expenses", "Expenses")
        dgvReportData.Columns.Add("Balance", "Balance")
        dgvReportData.Columns.Add("Savings", "Savings Rate")

        dgvReportData.Columns("Income").DefaultCellStyle.Format = "C"
        dgvReportData.Columns("Expenses").DefaultCellStyle.Format = "C"
        dgvReportData.Columns("Balance").DefaultCellStyle.Format = "C"
        dgvReportData.Columns("Savings").DefaultCellStyle.Format = "P2"
    End Sub

    ' Load income vs expenses data from database
    Private Sub LoadData(dateRange As DateRange)
        ' Format dates for query
        Dim startDateStr As String = dateRange.GetFormattedStartDate()
        Dim endDateStr As String = dateRange.GetFormattedEndDate()

        ' Clear existing data and rows
        dgvReportData.Rows.Clear()

        ' Get data from database
        Try
            Using connection As New OleDbConnection(connectionString)
                connection.Open()

                ' Get all months in the range
                Dim months As DateTime() = dateRange.GetMonthsInRange()

                ' Add a placeholder row for demonstration
                dgvReportData.Rows.Add("Demo Data", 1000, 750, 250, 0.25)
            End Using
        Catch ex As Exception
            MessageBox.Show("Database query error: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Debug.WriteLine("Database error: " & ex.Message & vbCrLf & ex.StackTrace)
        End Try
    End Sub

    ' Draw the pie chart
    Private Sub OnDrawPieChart(sender As Object, e As PaintEventArgs)
        ' Draw simple message for now
        Using brush As New SolidBrush(Color.White)
            e.Graphics.DrawString("Income vs Expenses Overview", New Font("Segoe UI", 16, FontStyle.Bold), brush, 80, 10)
            e.Graphics.DrawString("Pie chart will display income vs expenses breakdown", New Font("Segoe UI", 12), brush, 50, 100)
        End Using
    End Sub

    ' Draw the bar chart
    Private Sub OnDrawBarChart(sender As Object, e As PaintEventArgs)
        ' Draw simple message for now
        Using brush As New SolidBrush(Color.White)
            e.Graphics.DrawString("Monthly Income vs Expenses", New Font("Segoe UI", 16, FontStyle.Bold), brush, 80, 10)
            e.Graphics.DrawString("Bar chart will display monthly income vs expense comparison", New Font("Segoe UI", 12), brush, 50, 100)
        End Using
    End Sub
End Class