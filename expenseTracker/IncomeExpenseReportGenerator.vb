Imports System
Imports System.Windows.Forms
Imports System.Data.OleDb
Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.Diagnostics

Partial Public Class IncomeExpenseReportGenerator
    ' Database connection
    Private connectionString As String

    ' UI Controls
    Private WithEvents dgvReportData As DataGridView
    Private pnlPieChart As Panel
    Private pnlBarChart As Panel

    ' Flag to track if data has been loaded
    Private dataLoaded As Boolean = False

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

        ' Mark that data has been loaded
        dataLoaded = True

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

        ' Debug info
        Debug.WriteLine($"Loading income/expense data from {startDateStr} to {endDateStr}")

        ' Clear existing data and rows
        dgvReportData.Rows.Clear()

        ' Dictionary to store income and expenses by month
        Dim monthlyIncome As New Dictionary(Of String, Decimal)()
        Dim monthlyExpenses As New Dictionary(Of String, Decimal)()

        ' Get data from database
        Try
            Using connection As New OleDbConnection(connectionString)
                connection.Open()

                ' Debug log connection
                Debug.WriteLine("Database connection opened successfully")

                ' Query for income by month
                ' Use Format function to extract year and month
                Dim incomeQuery As String = "SELECT Format([Timestamp],'yyyy-mm') AS YearMonth, Sum([Amount]) AS TotalIncome " &
                                          "FROM Salary " &
                                          "WHERE [Timestamp] BETWEEN #" & startDateStr & "# AND #" & endDateStr & "# " &
                                          "GROUP BY Format([Timestamp],'yyyy-mm')"

                ' Debug log query
                Debug.WriteLine("Income Query: " & incomeQuery)

                Using command As New OleDbCommand(incomeQuery, connection)
                    Using reader As OleDbDataReader = command.ExecuteReader()
                        Debug.WriteLine("Income query executed, reading results...")

                        While reader.Read()
                            Dim yearMonth As String = reader("YearMonth").ToString()
                            Dim amount As Decimal = 0

                            If Not IsDBNull(reader("TotalIncome")) Then
                                amount = Convert.ToDecimal(reader("TotalIncome"))
                            End If

                            Debug.WriteLine($"Read income data: {yearMonth} = {amount}")
                            monthlyIncome(yearMonth) = amount
                        End While
                    End Using
                End Using

                ' Query for expenses by month
                ' Use Format function to extract year and month
                Dim expensesQuery As String = "SELECT Format([Timestamp],'yyyy-mm') AS YearMonth, Sum([Amount]) AS TotalExpenses " &
                                            "FROM Expenses " &
                                            "WHERE [Timestamp] BETWEEN #" & startDateStr & "# AND #" & endDateStr & "# " &
                                            "GROUP BY Format([Timestamp],'yyyy-mm')"

                ' Debug log query
                Debug.WriteLine("Expenses Query: " & expensesQuery)

                Using command As New OleDbCommand(expensesQuery, connection)
                    Using reader As OleDbDataReader = command.ExecuteReader()
                        Debug.WriteLine("Expenses query executed, reading results...")

                        While reader.Read()
                            Dim yearMonth As String = reader("YearMonth").ToString()
                            Dim amount As Decimal = 0

                            If Not IsDBNull(reader("TotalExpenses")) Then
                                amount = Convert.ToDecimal(reader("TotalExpenses"))
                            End If

                            Debug.WriteLine($"Read expense data: {yearMonth} = {amount}")
                            monthlyExpenses(yearMonth) = amount
                        End While
                    End Using
                End Using

                ' Get all months in the range
                Dim allMonths As New List(Of String)()

                ' Add all months from income data
                For Each month As String In monthlyIncome.Keys
                    If Not allMonths.Contains(month) Then
                        allMonths.Add(month)
                    End If
                Next

                ' Add all months from expense data
                For Each month As String In monthlyExpenses.Keys
                    If Not allMonths.Contains(month) Then
                        allMonths.Add(month)
                    End If
                Next

                ' If no data found, add a message to the debug log
                If allMonths.Count = 0 Then
                    Debug.WriteLine("No data found for the selected date range")
                    MessageBox.Show("No data found for the selected date range.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Return
                End If

                ' Sort months chronologically
                allMonths.Sort()

                ' Add data to grid
                For Each yearMonth As String In allMonths
                    ' Parse year and month from yearMonth string (format: yyyy-mm)
                    Dim parts As String() = yearMonth.Split("-"c)
                    If parts.Length >= 2 Then
                        Dim year As Integer = Convert.ToInt32(parts(0))
                        Dim month As Integer = Convert.ToInt32(parts(1))

                        ' Format month name
                        Dim monthDate As New DateTime(year, month, 1)
                        Dim monthName As String = monthDate.ToString("MMM yyyy")

                        ' Get income and expenses for this month
                        Dim income As Decimal = 0
                        If monthlyIncome.ContainsKey(yearMonth) Then
                            income = monthlyIncome(yearMonth)
                        End If

                        Dim expenses As Decimal = 0
                        If monthlyExpenses.ContainsKey(yearMonth) Then
                            expenses = monthlyExpenses(yearMonth)
                        End If

                        ' Calculate balance and savings rate
                        Dim balance As Decimal = income - expenses
                        Dim savingsRate As Double = 0
                        If income > 0 Then
                            savingsRate = Convert.ToDouble(balance / income)
                        End If

                        ' Add to grid
                        dgvReportData.Rows.Add(monthName, income, expenses, balance, savingsRate)
                    End If
                Next
            End Using
        Catch ex As Exception
            MessageBox.Show("Database query error: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Debug.WriteLine("Database error: " & ex.Message & vbCrLf & ex.StackTrace)
        End Try
    End Sub
End Class