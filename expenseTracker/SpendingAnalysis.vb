Imports System.Windows.Forms
Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.Diagnostics
Imports System.Collections.Generic
Imports System.Data.OleDb

Partial Public Class SpendingAnalysis
    Inherits Form

    ' Database connection
    Private connectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\ExpenseTracker.accdb;Persist Security Info=False;"

    ' Data containers
    Private categoryData As New Dictionary(Of String, Decimal)
    Private previousPeriodData As New Dictionary(Of String, Decimal)
    Private monthlyTotalData As New Dictionary(Of String, Decimal)
    Private monthlyTrends As New List(Of KeyValuePair(Of String, Decimal))
    Private anomalies As New List(Of SpendingAnomaly)

    ' Track if data has been loaded
    Private dataLoaded As Boolean = False

    Public Sub New()
        ' Initialize form
        InitializeComponent()

        ' Don't automatically run analysis - just show welcome message
        ShowWelcomeMessage()
    End Sub

    Private Sub ShowWelcomeMessage()
        ' Display welcome messages on each panel
        pnlTopCategories.Invalidate()
        pnlSavingOpportunities.Invalidate()
        pnlSpendingTrends.Invalidate()
        pnlBudgetRecommendations.Invalidate()
    End Sub

    Private Sub OnTimeFrameChanged(sender As Object, e As EventArgs) Handles cmbTimeFrame.SelectedIndexChanged
        ' Only proceed if controls are initialized
        If cmbTimeFrame Is Nothing OrElse lblYear Is Nothing OrElse
           cmbYear Is Nothing OrElse lblMonth Is Nothing OrElse cmbMonth Is Nothing Then
            Return
        End If

        ' Show/hide custom time frame controls based on selection
        Dim isCustom As Boolean = cmbTimeFrame.SelectedItem.ToString() = "Custom"

        lblYear.Visible = isCustom
        cmbYear.Visible = isCustom
        lblMonth.Visible = isCustom
        cmbMonth.Visible = isCustom
    End Sub

    Private Sub OnAnalyzeClick(sender As Object, e As EventArgs) Handles btnAnalyze.Click
        ' User clicked Analyze button
        Cursor = Cursors.WaitCursor
        Try
            Debug.WriteLine("------ Beginning Spending Analysis ------")
            PerformAnalysis(GetDateRange())
            dataLoaded = True
        Catch ex As Exception
            MessageBox.Show("Error analyzing spending: " & ex.Message, "Analysis Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Debug.WriteLine("Analysis error: " & ex.Message & vbCrLf & ex.StackTrace)
        Finally
            Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub PerformAnalysis(dateRange As DateRange)
        Try
            ' Clear previous data
            categoryData.Clear()
            previousPeriodData.Clear()
            monthlyTotalData.Clear()
            monthlyTrends.Clear()
            anomalies.Clear()

            Debug.WriteLine($"Date Range: {dateRange.StartDate.ToString("MM/dd/yyyy")} to {dateRange.EndDate.ToString("MM/dd/yyyy")}")

            ' Load spending data for the selected period
            LoadCategoryData(dateRange)

            ' Load data from previous period for comparison
            Dim previousRange As DateRange = GetPreviousPeriod(dateRange)
            Debug.WriteLine($"Previous Period: {previousRange.StartDate.ToString("MM/dd/yyyy")} to {previousRange.EndDate.ToString("MM/dd/yyyy")}")
            LoadPreviousPeriodData(previousRange)

            ' Load monthly trends for the past 6 months regardless of selection
            LoadMonthlyTrends()

            ' Find spending anomalies
            DetectSpendingAnomalies()

            ' Generate budget recommendations
            ' (Done dynamically in paint handler)

            ' Refresh all panels
            pnlTopCategories.Invalidate()
            pnlSavingOpportunities.Invalidate()
            pnlSpendingTrends.Invalidate()
            pnlBudgetRecommendations.Invalidate()

        Catch ex As Exception
            MessageBox.Show("Error performing analysis: " & ex.Message, "Analysis Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Debug.WriteLine("Analysis error: " & ex.Message & vbCrLf & ex.StackTrace)
        End Try
    End Sub

    Private Sub LoadCategoryData(dateRange As DateRange)
        ' Format dates for query - ensuring proper format for Access with FORWARD SLASHES
        Dim startDateStr As String = dateRange.StartDate.ToString("MM/dd/yyyy")
        Dim endDateStr As String = dateRange.EndDate.ToString("MM/dd/yyyy")

        Debug.WriteLine($"Loading category data from {startDateStr} to {endDateStr}")

        Try
            Using connection As New OleDbConnection(connectionString)
                connection.Open()
                Debug.WriteLine("Database connection opened successfully")

                ' Get total expenses first - using parameterized query
                Dim totalExpenses As Decimal = 0
                Dim totalQuery As String = "SELECT SUM([Amount]) FROM Expenses WHERE [Timestamp] BETWEEN @StartDate AND @EndDate"

                Debug.WriteLine("Total query: " + totalQuery)

                Using command As New OleDbCommand(totalQuery, connection)
                    command.Parameters.AddWithValue("@StartDate", startDateStr)
                    command.Parameters.AddWithValue("@EndDate", endDateStr)

                    Dim result = command.ExecuteScalar()
                    If result IsNot Nothing AndAlso Not IsDBNull(result) Then
                        totalExpenses = Convert.ToDecimal(result)
                    End If
                End Using

                Debug.WriteLine($"Total expenses: {totalExpenses}")

                ' If no expenses found, display message and return
                If totalExpenses = 0 Then
                    MessageBox.Show("No expenses found for the selected date range.", "No Data", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Return
                End If

                ' Query to get expenses by category - using parameterized query
                Dim query As String = "SELECT [Category], SUM([Amount]) AS TotalAmount " &
                                     "FROM Expenses " &
                                     "WHERE [Timestamp] BETWEEN @StartDate AND @EndDate " &
                                     "GROUP BY [Category] " &
                                     "ORDER BY SUM([Amount]) DESC"

                Debug.WriteLine("Category query: " + query)

                Using command As New OleDbCommand(query, connection)
                    command.Parameters.AddWithValue("@StartDate", startDateStr)
                    command.Parameters.AddWithValue("@EndDate", endDateStr)

                    Using reader As OleDbDataReader = command.ExecuteReader()
                        While reader.Read()
                            Dim category As String = If(IsDBNull(reader("Category")), "Uncategorized", reader("Category").ToString())
                            Dim amount As Decimal = Convert.ToDecimal(reader("TotalAmount"))
                            Dim percentage As Double = If(totalExpenses = 0, 0, Convert.ToDouble(amount / totalExpenses))

                            Debug.WriteLine($"Category: {category}, Amount: {amount}, Percentage: {percentage}")

                            ' Add to dictionary for charts
                            categoryData(category) = amount
                        End While
                    End Using
                End Using

                Debug.WriteLine($"Loaded {categoryData.Count} categories with data")
            End Using
        Catch ex As Exception
            MessageBox.Show("Database query error: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Debug.WriteLine("Database error: " & ex.Message & vbCrLf & ex.StackTrace)
        End Try
    End Sub

    Private Sub LoadPreviousPeriodData(dateRange As DateRange)
        ' Format dates for query - ensuring proper format for Access with FORWARD SLASHES
        Dim startDateStr As String = dateRange.StartDate.ToString("MM/dd/yyyy")
        Dim endDateStr As String = dateRange.EndDate.ToString("MM/dd/yyyy")

        Debug.WriteLine($"Loading previous period data from {startDateStr} to {endDateStr}")

        Try
            Using connection As New OleDbConnection(connectionString)
                connection.Open()

                ' Query to get expenses by category for previous period
                Dim query As String = "SELECT [Category], SUM([Amount]) AS TotalAmount " &
                                     "FROM Expenses " &
                                     "WHERE [Timestamp] BETWEEN @StartDate AND @EndDate " &
                                     "GROUP BY [Category] " &
                                     "ORDER BY TotalAmount DESC"

                Debug.WriteLine("Previous period query: " + query)

                Using command As New OleDbCommand(query, connection)
                    command.Parameters.AddWithValue("@StartDate", startDateStr)
                    command.Parameters.AddWithValue("@EndDate", endDateStr)

                    Using reader As OleDbDataReader = command.ExecuteReader()
                        While reader.Read()
                            Dim category As String = If(IsDBNull(reader("Category")), "Uncategorized", reader("Category").ToString())
                            Dim amount As Decimal = Convert.ToDecimal(reader("TotalAmount"))

                            Debug.WriteLine($"Read previous period data: {category} = {amount}")

                            ' Add to dictionary for comparison
                            previousPeriodData(category) = amount
                        End While
                    End Using
                End Using

                Debug.WriteLine($"Loaded {previousPeriodData.Count} categories from previous period")
            End Using
        Catch ex As Exception
            Debug.WriteLine("Previous period data error: " & ex.Message & vbCrLf & ex.StackTrace)
        End Try
    End Sub

    Private Sub LoadMonthlyTrends()
        ' Get the past 6 months for trend analysis
        Dim endDate As DateTime = DateTime.Now
        Dim startDate As DateTime = endDate.AddMonths(-6)

        ' Format dates for query - ensuring proper format for Access with FORWARD SLASHES
        Dim startDateStr As String = startDate.ToString("MM/dd/yyyy")
        Dim endDateStr As String = endDate.ToString("MM/dd/yyyy")

        Debug.WriteLine($"Loading monthly trends from {startDateStr} to {endDateStr}")

        Try
            Using connection As New OleDbConnection(connectionString)
                connection.Open()

                ' Query to get monthly totals with proper date formatting
                Dim query As String = "SELECT Format([Timestamp],'yyyy-mm') AS YearMonth, SUM([Amount]) AS TotalAmount " &
                                     "FROM Expenses " &
                                     "WHERE [Timestamp] BETWEEN @StartDate AND @EndDate " &
                                     "GROUP BY Format([Timestamp],'yyyy-mm') " &
                                     "ORDER BY YearMonth"

                Debug.WriteLine("Monthly trends query: " + query)

                Using command As New OleDbCommand(query, connection)
                    command.Parameters.AddWithValue("@StartDate", startDateStr)
                    command.Parameters.AddWithValue("@EndDate", endDateStr)

                    Using reader As OleDbDataReader = command.ExecuteReader()
                        While reader.Read()
                            Dim yearMonth As String = reader("YearMonth").ToString()
                            Dim amount As Decimal = Convert.ToDecimal(reader("TotalAmount"))

                            Debug.WriteLine($"Read monthly trend: {yearMonth} = {amount}")

                            ' Parse year and month for better display
                            Dim parts As String() = yearMonth.Split("-"c)
                            If parts.Length >= 2 Then
                                Dim year As Integer = Convert.ToInt32(parts(0))
                                Dim month As Integer = Convert.ToInt32(parts(1))

                                ' Format month name
                                Dim monthDate As New DateTime(year, month, 1)
                                Dim monthName As String = monthDate.ToString("MMM yyyy")

                                ' Add to dictionary for trend analysis
                                monthlyTotalData(monthName) = amount
                                monthlyTrends.Add(New KeyValuePair(Of String, Decimal)(monthName, amount))
                            End If
                        End While
                    End Using
                End Using

                Debug.WriteLine($"Loaded {monthlyTrends.Count} months of trend data")
            End Using
        Catch ex As Exception
            Debug.WriteLine("Monthly trends error: " & ex.Message & vbCrLf & ex.StackTrace)
        End Try
    End Sub

    Private Sub DetectSpendingAnomalies()
        Debug.WriteLine("Detecting spending anomalies...")

        ' Compare current spending with previous period to detect anomalies
        For Each category In categoryData.Keys
            If previousPeriodData.ContainsKey(category) Then
                Dim currentAmount As Decimal = categoryData(category)
                Dim previousAmount As Decimal = previousPeriodData(category)

                ' Skip categories with very small amounts
                If previousAmount < 100 And currentAmount < 100 Then
                    Debug.WriteLine($"Skipping small category: {category} (current: {currentAmount}, previous: {previousAmount})")
                    Continue For
                End If

                ' Calculate percentage change
                Dim percentChange As Double = If(previousAmount = 0, 100, (Convert.ToDouble(currentAmount - previousAmount) / Convert.ToDouble(previousAmount)) * 100)

                Debug.WriteLine($"Category {category}: Previous: {previousAmount}, Current: {currentAmount}, Change: {percentChange:0.0}%")

                ' Flag significant increases (30% or more)
                If percentChange >= 30 Then
                    Debug.WriteLine($"Found anomaly: {category} increased by {percentChange:0.0}%")

                    Dim anomaly As New SpendingAnomaly() With {
                        .Category = category,
                        .CurrentAmount = currentAmount,
                        .PreviousAmount = previousAmount,
                        .PercentChange = percentChange
                    }
                    anomalies.Add(anomaly)
                End If
            End If
        Next

        ' Sort anomalies by percent change (descending)
        anomalies.Sort(Function(a, b) b.PercentChange.CompareTo(a.PercentChange))

        Debug.WriteLine($"Found {anomalies.Count} spending anomalies")
    End Sub

    Private Function GetPreviousPeriod(currentRange As DateRange) As DateRange
        ' Calculate a previous period of the same length
        Dim duration As TimeSpan = currentRange.EndDate - currentRange.StartDate
        Dim previousEnd As DateTime = currentRange.StartDate.AddDays(-1)
        Dim previousStart As DateTime = previousEnd.AddDays(-duration.Days)

        Return New DateRange(previousStart, previousEnd)
    End Function

    Private Function GetDateRange() As DateRange
        Debug.WriteLine("Getting date range for analysis")

        ' Check if controls are initialized
        If cmbTimeFrame Is Nothing Then
            ' Default to current month if controls aren't initialized
            Return New DateRange(New DateTime(DateTime.Now.Year, DateTime.Now.Month, 1), DateTime.Now)
        End If

        Dim startDate As DateTime
        Dim endDate As DateTime = DateTime.Now
        Dim timeFrame As String = cmbTimeFrame.SelectedItem.ToString()

        Debug.WriteLine($"Selected time frame: {timeFrame}")

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
                If cmbYear IsNot Nothing AndAlso cmbMonth IsNot Nothing AndAlso
                   cmbYear.SelectedItem IsNot Nothing AndAlso cmbMonth.SelectedIndex >= 0 Then

                    Dim selectedYear As Integer = Convert.ToInt32(cmbYear.SelectedItem)
                    Dim selectedMonth As Integer = cmbMonth.SelectedIndex + 1
                    startDate = New DateTime(selectedYear, selectedMonth, 1)
                    endDate = startDate.AddMonths(1).AddDays(-1)
                Else
                    MessageBox.Show("Please select both year and month for custom range", "Selection Required", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Return New DateRange(DateTime.Now.AddMonths(-1), DateTime.Now)
                End If
            Case Else
                startDate = New DateTime(DateTime.Now.Year, DateTime.Now.Month, 1)
        End Select

        Debug.WriteLine($"Date range: {startDate.ToString("MM/dd/yyyy")} to {endDate.ToString("MM/dd/yyyy")}")
        Return New DateRange(startDate, endDate)
    End Function

    ' Helper method to get colors for category chart
    Private Function GetCategoryColor(index As Integer) As Color
        Dim colors As Color() = {
            Color.FromArgb(0, 173, 181),    ' Teal
            Color.FromArgb(255, 77, 77),    ' Red
            Color.FromArgb(76, 187, 23),    ' Green
            Color.FromArgb(255, 190, 11),   ' Yellow
            Color.FromArgb(153, 102, 255),  ' Purple
            Color.FromArgb(58, 134, 255),   ' Blue
            Color.FromArgb(255, 128, 0),    ' Orange
            Color.FromArgb(240, 98, 146)    ' Pink
        }

        Return colors(index Mod colors.Length)
    End Function

    ' Button hover effects
    Private Sub OnButtonMouseEnter(sender As Object, e As EventArgs) Handles btnAnalyze.MouseEnter
        btnAnalyze.BackColor = Color.FromArgb(0, 150, 160)
    End Sub

    Private Sub OnButtonMouseLeave(sender As Object, e As EventArgs) Handles btnAnalyze.MouseLeave
        btnAnalyze.BackColor = Color.FromArgb(0, 173, 181)
    End Sub

    ' Nested class to represent spending anomalies
    Public Class SpendingAnomaly
        Public Property Category As String
        Public Property CurrentAmount As Decimal
        Public Property PreviousAmount As Decimal
        Public Property PercentChange As Double
    End Class
End Class