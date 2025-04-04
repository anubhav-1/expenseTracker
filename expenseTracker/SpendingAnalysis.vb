Imports System.Windows.Forms
Imports System.Drawing
Imports System.Data.OleDb
Imports System.Drawing.Drawing2D
Imports System.Diagnostics
Imports System.Collections.Generic

Public Class SpendingAnalysis
    Inherits Form

    ' Database connection
    Private connectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\ExpenseTracker.accdb;Persist Security Info=False;"

    ' Control declarations
    Private pnlFilters As Panel
    Private lblTimeFrame As Label
    Private cmbTimeFrame As ComboBox
    Private lblYear As Label
    Private lblMonth As Label
    Private cmbYear As ComboBox
    Private cmbMonth As ComboBox
    Private btnAnalyze As Button

    ' Analysis panels
    Private pnlTopCategories As Panel
    Private pnlSavingOpportunities As Panel
    Private pnlSpendingTrends As Panel
    Private pnlBudgetRecommendations As Panel

    ' Data containers
    Private categoryData As New Dictionary(Of String, Decimal)
    Private previousPeriodData As New Dictionary(Of String, Decimal)
    Private monthlyTotalData As New Dictionary(Of String, Decimal)
    Private monthlyTrends As New List(Of KeyValuePair(Of String, Decimal))
    Private anomalies As New List(Of SpendingAnomaly)

    ' Track if data has been loaded
    Private dataLoaded As Boolean = False

    Public Sub New()
        ' Form setup
        Me.Text = "Spending Analysis"
        Me.Size = New Size(1000, 800)
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.BackColor = Color.FromArgb(34, 40, 49)
        Me.FormBorderStyle = FormBorderStyle.None
        Me.Dock = DockStyle.Fill

        InitializeComponents()
    End Sub

    Private Sub InitializeComponents()
        ' Top panel for filters
        pnlFilters = New Panel()
        pnlFilters.Dock = DockStyle.Top
        pnlFilters.Height = 80
        pnlFilters.BackColor = Color.FromArgb(45, 52, 64)
        pnlFilters.Padding = New Padding(10)
        Me.Controls.Add(pnlFilters)

        ' Time Frame selector
        lblTimeFrame = New Label()
        lblTimeFrame.Text = "Analysis Period:"
        lblTimeFrame.ForeColor = Color.White
        lblTimeFrame.Font = New Font("Segoe UI", 10, FontStyle.Bold)
        lblTimeFrame.Location = New Point(20, 15)
        lblTimeFrame.AutoSize = True
        pnlFilters.Controls.Add(lblTimeFrame)

        cmbTimeFrame = New ComboBox()
        cmbTimeFrame.Location = New Point(150, 12)
        cmbTimeFrame.Size = New Size(180, 28)
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
        lblYear.Location = New Point(350, 15)
        lblYear.AutoSize = True
        lblYear.Visible = False
        pnlFilters.Controls.Add(lblYear)

        cmbYear = New ComboBox()
        cmbYear.Location = New Point(400, 12)
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
        lblMonth.Location = New Point(520, 15)
        lblMonth.AutoSize = True
        lblMonth.Visible = False
        pnlFilters.Controls.Add(lblMonth)

        cmbMonth = New ComboBox()
        cmbMonth.Location = New Point(580, 12)
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

        ' Analyze button
        btnAnalyze = New Button()
        btnAnalyze.Text = "Analyze Spending"
        btnAnalyze.Location = New Point(750, 12)
        btnAnalyze.Size = New Size(150, 40)
        btnAnalyze.FlatStyle = FlatStyle.Flat
        btnAnalyze.FlatAppearance.BorderSize = 0
        btnAnalyze.BackColor = Color.FromArgb(0, 173, 181)
        btnAnalyze.ForeColor = Color.White
        btnAnalyze.Font = New Font("Segoe UI", 11, FontStyle.Bold)
        btnAnalyze.Cursor = Cursors.Hand
        AddHandler btnAnalyze.Click, AddressOf OnAnalyzeClick
        AddHandler btnAnalyze.MouseEnter, AddressOf OnButtonMouseEnter
        AddHandler btnAnalyze.MouseLeave, AddressOf OnButtonMouseLeave
        pnlFilters.Controls.Add(btnAnalyze)

        ' Main content area setup
        CreateAnalysisPanels()

        ' Don't automatically run analysis - just show welcome message
        ShowWelcomeMessage()
    End Sub

    Private Sub CreateAnalysisPanels()
        ' Top Categories Panel (Top Left)
        pnlTopCategories = New Panel()
        pnlTopCategories.Location = New Point(20, 100)
        pnlTopCategories.Size = New Size(465, 330)
        pnlTopCategories.BackColor = Color.FromArgb(57, 62, 70)
        AddHandler pnlTopCategories.Paint, AddressOf OnPaintTopCategories
        Me.Controls.Add(pnlTopCategories)

        ' Saving Opportunities Panel (Top Right)
        pnlSavingOpportunities = New Panel()
        pnlSavingOpportunities.Location = New Point(505, 100)
        pnlSavingOpportunities.Size = New Size(465, 330)
        pnlSavingOpportunities.BackColor = Color.FromArgb(57, 62, 70)
        AddHandler pnlSavingOpportunities.Paint, AddressOf OnPaintSavingOpportunities
        Me.Controls.Add(pnlSavingOpportunities)

        ' Spending Trends Panel (Bottom Left)
        pnlSpendingTrends = New Panel()
        pnlSpendingTrends.Location = New Point(20, 450)
        pnlSpendingTrends.Size = New Size(465, 330)
        pnlSpendingTrends.BackColor = Color.FromArgb(57, 62, 70)
        AddHandler pnlSpendingTrends.Paint, AddressOf OnPaintSpendingTrends
        Me.Controls.Add(pnlSpendingTrends)

        ' Budget Recommendations Panel (Bottom Right)
        pnlBudgetRecommendations = New Panel()
        pnlBudgetRecommendations.Location = New Point(505, 450)
        pnlBudgetRecommendations.Size = New Size(465, 330)
        pnlBudgetRecommendations.BackColor = Color.FromArgb(57, 62, 70)
        AddHandler pnlBudgetRecommendations.Paint, AddressOf OnPaintBudgetRecommendations
        Me.Controls.Add(pnlBudgetRecommendations)
    End Sub

    Private Sub ShowWelcomeMessage()
        ' Display welcome messages on each panel
        pnlTopCategories.Invalidate()
        pnlSavingOpportunities.Invalidate()
        pnlSpendingTrends.Invalidate()
        pnlBudgetRecommendations.Invalidate()
    End Sub

    Private Sub OnTimeFrameChanged(sender As Object, e As EventArgs)
        ' Show/hide custom time frame controls based on selection
        Dim isCustom As Boolean = cmbTimeFrame.SelectedItem.ToString() = "Custom"

        lblYear.Visible = isCustom
        cmbYear.Visible = isCustom
        lblMonth.Visible = isCustom
        cmbMonth.Visible = isCustom
    End Sub

    Private Sub OnAnalyzeClick(sender As Object, e As EventArgs)
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
                If cmbYear.SelectedItem Is Nothing OrElse cmbMonth.SelectedItem Is Nothing Then
                    MessageBox.Show("Please select both year and month for custom range", "Selection Required", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Return New DateRange(DateTime.Now.AddMonths(-1), DateTime.Now)
                End If

                Dim selectedYear As Integer = Convert.ToInt32(cmbYear.SelectedItem)
                Dim selectedMonth As Integer = cmbMonth.SelectedIndex + 1
                startDate = New DateTime(selectedYear, selectedMonth, 1)
                endDate = startDate.AddMonths(1).AddDays(-1)
            Case Else
                startDate = New DateTime(DateTime.Now.Year, DateTime.Now.Month, 1)
        End Select

        Debug.WriteLine($"Date range: {startDate.ToString("MM/dd/yyyy")} to {endDate.ToString("MM/dd/yyyy")}")
        Return New DateRange(startDate, endDate)
    End Function

    ' Painting methods for the analysis panels
    Private Sub OnPaintTopCategories(sender As Object, e As PaintEventArgs)
        e.Graphics.SmoothingMode = SmoothingMode.AntiAlias

        If Not dataLoaded Then
            ' Show welcome message
            Using brush As New SolidBrush(Color.White)
                e.Graphics.DrawString("Top Spending Categories", New Font("Segoe UI", 14, FontStyle.Bold), brush, 20, 20)
                e.Graphics.DrawString("Click 'Analyze Spending' to see your top expense categories", New Font("Segoe UI", 10), brush, 20, 50)
            End Using
            Return
        End If

        ' Draw panel title
        Using brush As New SolidBrush(Color.White)
            e.Graphics.DrawString("Top Spending Categories", New Font("Segoe UI", 14, FontStyle.Bold), brush, 20, 20)
        End Using

        ' Check if we have data
        If categoryData.Count = 0 Then
            Using brush As New SolidBrush(Color.LightGray)
                e.Graphics.DrawString("No expense data found for the selected time period.", New Font("Segoe UI", 10), brush, 20, 60)
            End Using
            Return
        End If

        ' Get top 5 categories by amount
        Dim topCategories = categoryData.OrderByDescending(Function(kvp) kvp.Value).Take(5).ToList()

        ' Calculate total for percentage
        Dim total As Decimal = categoryData.Values.Sum()

        ' Display categories with amount and percentage
        Dim y As Integer = 60
        Dim index As Integer = 1
        Dim barWidth As Integer = 300
        Dim barHeight As Integer = 30
        Dim barSpacing As Integer = 40

        For Each category In topCategories
            Dim percentage As Double = Convert.ToDouble(category.Value / total * 100)
            Dim amount As Decimal = category.Value

            ' Draw category details
            Using brush As New SolidBrush(Color.White)
                ' Category name and amount
                e.Graphics.DrawString($"{index}. {category.Key}", New Font("Segoe UI", 10, FontStyle.Bold), brush, 20, y)
                e.Graphics.DrawString($"{amount:C}", New Font("Segoe UI", 10), brush, 350, y)

                ' Percentage
                e.Graphics.DrawString($"{percentage:0.0}%", New Font("Segoe UI", 9), brush, 20, y + 25)
            End Using

            ' Draw bar
            Dim barLength As Integer = CInt((percentage / 100) * barWidth)
            Using brush As New SolidBrush(GetCategoryColor(index - 1))
                e.Graphics.FillRectangle(brush, 110, y + 25, barLength, 10)
            End Using

            y += barSpacing
            index += 1
        Next

        ' Show total
        y += 10
        Using brush As New SolidBrush(Color.White)
            e.Graphics.DrawString("Total Spending:", New Font("Segoe UI", 10, FontStyle.Bold), brush, 20, y)
            e.Graphics.DrawString($"{total:C}", New Font("Segoe UI", 10, FontStyle.Bold), brush, 350, y)
        End Using
    End Sub

    Private Sub OnPaintSavingOpportunities(sender As Object, e As PaintEventArgs)
        e.Graphics.SmoothingMode = SmoothingMode.AntiAlias

        If Not dataLoaded Then
            ' Show welcome message
            Using brush As New SolidBrush(Color.White)
                e.Graphics.DrawString("Saving Opportunities", New Font("Segoe UI", 14, FontStyle.Bold), brush, 20, 20)
                e.Graphics.DrawString("Click 'Analyze Spending' to see potential savings", New Font("Segoe UI", 10), brush, 20, 50)
            End Using
            Return
        End If

        ' Draw panel title
        Using brush As New SolidBrush(Color.White)
            e.Graphics.DrawString("Spending Anomalies", New Font("Segoe UI", 14, FontStyle.Bold), brush, 20, 20)
        End Using

        ' Check if we have anomalies
        If anomalies.Count = 0 Then
            Using brush As New SolidBrush(Color.LightGray)
                e.Graphics.DrawString("No significant spending increases detected.", New Font("Segoe UI", 10), brush, 20, 60)
                e.Graphics.DrawString("Good job maintaining consistent spending!", New Font("Segoe UI", 10), brush, 20, 85)
            End Using
            Return
        End If

        ' Show subtitle with explanation
        Using brush As New SolidBrush(Color.LightGray)
            e.Graphics.DrawString("Categories with 30%+ spending increase compared to previous period:", New Font("Segoe UI", 9), brush, 20, 50)
        End Using

        ' Display anomalies
        Dim y As Integer = 80
        Dim count As Integer = 0

        For Each anomaly In anomalies
            If count >= 5 Then Exit For ' Limit to top 5 anomalies

            ' Draw anomaly details
            Using headerBrush As New SolidBrush(Color.White)
                ' Category name
                e.Graphics.DrawString(anomaly.Category, New Font("Segoe UI", 11, FontStyle.Bold), headerBrush, 20, y)
            End Using

            ' Calculate display values
            Dim increase As Decimal = anomaly.CurrentAmount - anomaly.PreviousAmount
            Dim arrowX As Integer = 280
            Dim arrowLength As Integer = Math.Min(CInt(anomaly.PercentChange), 100)

            ' Draw values
            Using valueBrush As New SolidBrush(Color.White)
                Dim valueFont As New Font("Segoe UI", 9)
                e.Graphics.DrawString($"Previous: {anomaly.PreviousAmount:C}", valueFont, valueBrush, 30, y + 25)
                e.Graphics.DrawString($"Current: {anomaly.CurrentAmount:C}", valueFont, valueBrush, 30, y + 45)
                e.Graphics.DrawString($"Increase: {increase:C} ({anomaly.PercentChange:0.0}%)", valueFont, valueBrush, 30, y + 65)
            End Using

            ' Draw arrow showing increase
            Using redPen As New Pen(Color.FromArgb(255, 77, 77), 2)
                ' Draw line
                e.Graphics.DrawLine(redPen, arrowX, y + 35, arrowX + arrowLength, y + 35)
                ' Draw arrowhead
                e.Graphics.DrawLine(redPen, arrowX + arrowLength - 10, y + 30, arrowX + arrowLength, y + 35)
                e.Graphics.DrawLine(redPen, arrowX + arrowLength - 10, y + 40, arrowX + arrowLength, y + 35)
            End Using

            y += 90
            count += 1
        Next

        ' Add recommendation if we have anomalies
        If anomalies.Count > 0 Then
            Using brush As New SolidBrush(Color.FromArgb(0, 173, 181))
                e.Graphics.DrawString("💡 Tip: Review these categories for potential savings opportunities.", New Font("Segoe UI", 10, FontStyle.Bold), brush, 20, y + 10)
            End Using
        End If
    End Sub

    Private Sub OnPaintSpendingTrends(sender As Object, e As PaintEventArgs)
        e.Graphics.SmoothingMode = SmoothingMode.AntiAlias

        If Not dataLoaded Then
            ' Show welcome message
            Using brush As New SolidBrush(Color.White)
                e.Graphics.DrawString("Spending Trends", New Font("Segoe UI", 14, FontStyle.Bold), brush, 20, 20)
                e.Graphics.DrawString("Click 'Analyze Spending' to see your spending trends", New Font("Segoe UI", 10), brush, 20, 50)
            End Using
            Return
        End If

        ' Draw panel title
        Using brush As New SolidBrush(Color.White)
            e.Graphics.DrawString("Monthly Spending Trends", New Font("Segoe UI", 14, FontStyle.Bold), brush, 20, 20)
        End Using

        ' Check if we have data
        If monthlyTrends.Count = 0 Then
            Using brush As New SolidBrush(Color.LightGray)
                e.Graphics.DrawString("No trend data available for the past 6 months.", New Font("Segoe UI", 10), brush, 20, 60)
            End Using
            Return
        End If

        ' Sort months chronologically
        monthlyTrends.Sort(Function(a, b) DateTime.Parse("01 " & a.Key).CompareTo(DateTime.Parse("01 " & b.Key)))

        ' Define chart area
        Dim chartRect As New Rectangle(30, 60, 400, 200)

        ' Draw axes
        Using axisPen As New Pen(Color.Gray, 1)
            ' X-axis (horizontal line)
            e.Graphics.DrawLine(axisPen, chartRect.Left, chartRect.Bottom, chartRect.Right, chartRect.Bottom)

            ' Y-axis (vertical line)
            e.Graphics.DrawLine(axisPen, chartRect.Left, chartRect.Top, chartRect.Left, chartRect.Bottom)
        End Using

        ' Find maximum value for scaling
        Dim maxValue As Decimal = If(monthlyTrends.Count > 0, monthlyTrends.Max(Function(t) t.Value), 0)
        If maxValue = 0 Then maxValue = 1 ' Avoid division by zero

        ' Round up max value for cleaner axis
        maxValue = Math.Ceiling(maxValue / 500) * 500

        ' Draw value markers on Y-axis
        Using grayBrush As New SolidBrush(Color.Gray)
            Dim valueFont As New Font("Segoe UI", 8)
            For i As Integer = 0 To 4
                Dim yValue As Decimal = maxValue * i / 4
                Dim y As Integer = chartRect.Bottom - (i * chartRect.Height / 4)
                e.Graphics.DrawString(yValue.ToString("C0"), valueFont, grayBrush, chartRect.Left - 50, y - 6)

                ' Draw horizontal grid line
                Using gridPen As New Pen(Color.FromArgb(60, 70, 80), 1)
                    gridPen.DashStyle = DashStyle.Dot
                    e.Graphics.DrawLine(gridPen, chartRect.Left, y, chartRect.Right, y)
                End Using
            Next
        End Using

        ' Draw the line chart
        If monthlyTrends.Count > 1 Then
            ' Calculate points for the line
            Dim points As New List(Of PointF)
            Dim barWidth As Integer = chartRect.Width / (monthlyTrends.Count + 1)
            Dim x As Single = chartRect.Left + barWidth / 2

            For Each item In monthlyTrends
                ' Calculate Y position (inverted, since 0 is at top in GDI+)
                Dim y As Single = chartRect.Bottom - (CSng(item.Value / maxValue) * chartRect.Height)
                points.Add(New PointF(x, y))
                x += barWidth
            Next

            ' Draw trend line
            Using trendPen As New Pen(Color.FromArgb(0, 173, 181), 3)
                If points.Count >= 2 Then
                    e.Graphics.DrawLines(trendPen, points.ToArray())
                End If
            End Using

            ' Draw data points and labels
            x = chartRect.Left + barWidth / 2
            Dim index As Integer = 0
            Using pointBrush As New SolidBrush(Color.FromArgb(0, 173, 181))
                Using whiteBrush As New SolidBrush(Color.White)
                    Using grayBrush As New SolidBrush(Color.LightGray)
                        For Each item In monthlyTrends
                            ' Get point position
                            Dim point As PointF = points(index)

                            ' Draw point circle
                            e.Graphics.FillEllipse(pointBrush, point.X - 5, point.Y - 5, 10, 10)
                            e.Graphics.DrawEllipse(New Pen(Color.White), point.X - 5, point.Y - 5, 10, 10)

                            ' Draw month label on X-axis
                            Dim labelFont As New Font("Segoe UI", 8)
                            Dim shortMonth As String = item.Key.Split(" "c)(0) ' Get just "Jan" from "Jan 2023"

                            ' Draw rotated month label
                            e.Graphics.TranslateTransform(point.X, chartRect.Bottom + 5)
                            e.Graphics.RotateTransform(45)
                            e.Graphics.DrawString(shortMonth, labelFont, grayBrush, 0, 0)
                            e.Graphics.ResetTransform()

                            ' Draw amount above point if there's enough space
                            If index Mod 2 = 0 Then ' Draw every other label to avoid crowding
                                Dim amountStr As String = item.Value.ToString("C0")
                                Dim textSize As SizeF = e.Graphics.MeasureString(amountStr, labelFont)
                                e.Graphics.DrawString(amountStr, labelFont, whiteBrush, point.X - textSize.Width / 2, point.Y - 25)
                            End If

                            x += barWidth
                            index += 1
                        Next
                    End Using
                End Using
            End Using

            ' Draw trend analysis
            If monthlyTrends.Count >= 2 Then
                Dim firstAmount As Decimal = monthlyTrends.First().Value
                Dim lastAmount As Decimal = monthlyTrends.Last().Value
                Dim percentChange As Double = If(firstAmount = 0, 0, Convert.ToDouble((lastAmount - firstAmount) / firstAmount) * 100)

                Dim trendText As String
                Dim trendColor As Color

                If percentChange > 15 Then
                    trendText = $"⚠️ Spending increased by {percentChange:0.0}% over this period"
                    trendColor = Color.FromArgb(255, 77, 77) ' Red
                ElseIf percentChange < -15 Then
                    trendText = $"👍 Spending decreased by {Math.Abs(percentChange):0.0}% over this period"
                    trendColor = Color.FromArgb(76, 187, 23) ' Green
                Else
                    trendText = $"✓ Spending remained relatively stable ({percentChange:0.0}% change)"
                    trendColor = Color.FromArgb(0, 173, 181) ' Teal
                End If

                Using brush As New SolidBrush(trendColor)
                    e.Graphics.DrawString(trendText, New Font("Segoe UI", 10, FontStyle.Bold), brush, 30, chartRect.Bottom + 70)
                End Using
            End If
        End If
    End Sub

    Private Sub OnPaintBudgetRecommendations(sender As Object, e As PaintEventArgs)
        e.Graphics.SmoothingMode = SmoothingMode.AntiAlias

        If Not dataLoaded Then
            ' Show welcome message
            Using brush As New SolidBrush(Color.White)
                e.Graphics.DrawString("Budget Recommendations", New Font("Segoe UI", 14, FontStyle.Bold), brush, 20, 20)
                e.Graphics.DrawString("Click 'Analyze Spending' to see budget recommendations", New Font("Segoe UI", 10), brush, 20, 50)
            End Using
            Return
        End If

        ' Draw panel title
        Using brush As New SolidBrush(Color.White)
            e.Graphics.DrawString("Budget Recommendations", New Font("Segoe UI", 14, FontStyle.Bold), brush, 20, 20)
        End Using

        ' Get total spending for the current period
        Dim totalSpending As Decimal = 0
        For Each amount In categoryData.Values
            totalSpending += amount
        Next

        ' Generate budget recommendations
        Dim y As Integer = 60
        Dim recommendations As New List(Of String)

        ' Check for no data
        If categoryData.Count = 0 Then
            Using brush As New SolidBrush(Color.LightGray)
                e.Graphics.DrawString("Not enough data to generate recommendations.", New Font("Segoe UI", 10), brush, 20, y)
            End Using
            Return
        End If

        ' Check total spending trend
        If monthlyTrends.Count >= 2 Then
            Dim firstAmount As Decimal = monthlyTrends.First().Value
            Dim lastAmount As Decimal = monthlyTrends.Last().Value
            Dim percentChange As Double = If(firstAmount = 0, 0, Convert.ToDouble((lastAmount - firstAmount) / firstAmount) * 100)

            If percentChange > 20 Then
                recommendations.Add("Your overall spending has increased by " & percentChange.ToString("0.0") & "% recently. Consider setting a monthly budget limit.")
            End If
        End If

        ' Check for anomalies
        If anomalies.Count > 0 Then
            Dim topAnomaly As SpendingAnomaly = anomalies(0)
            recommendations.Add("Consider checking your " & topAnomaly.Category & " spending which increased by " & topAnomaly.PercentChange.ToString("0.0") & "%.")
        End If

        ' Check for top categories
        Dim topCategory As String = ""
        Dim topAmount As Decimal = 0

        For Each category In categoryData
            If category.Value > topAmount Then
                topAmount = category.Value
                topCategory = category.Key
            End If
        Next

        If topAmount > 0 Then
            Dim percentage As Double = Convert.ToDouble(topAmount / totalSpending) * 100
            If percentage > 40 Then
                recommendations.Add("Your " & topCategory & " spending accounts for " & percentage.ToString("0.0") & "% of your total. Consider diversifying your budget.")
            End If
        End If

        ' If no specific recommendations, provide general advice
        If recommendations.Count = 0 Then
            recommendations.Add("Your spending patterns look good! Continue maintaining your current budget approach.")
            recommendations.Add("Consider setting aside " & Math.Round(totalSpending * 0.1, 2).ToString("C0") & " (10% of your expenses) as emergency savings.")
        End If

        ' Draw recommendations with icons
        Dim icons As String() = {"💡", "✅", "⚠️", "💰", "📊"}
        Dim iconBrush As New SolidBrush(Color.White)
        Dim textBrush As New SolidBrush(Color.FromArgb(0, 173, 181))
        Dim titleFont As New Font("Segoe UI", 12, FontStyle.Bold)
        Dim bodyFont As New Font("Segoe UI", 10)

        ' Draw subtitle
        Using grayBrush As New SolidBrush(Color.LightGray)
            e.Graphics.DrawString("Based on your spending patterns, we recommend:", bodyFont, grayBrush, 20, y)
        End Using

        y += 30

        ' Draw each recommendation
        For i As Integer = 0 To Math.Min(recommendations.Count - 1, 4)
            ' Draw icon
            Dim icon As String = icons(i Mod icons.Length)
            e.Graphics.DrawString(icon, New Font("Segoe UI", 14), iconBrush, 20, y)

            ' Draw recommendation text (with word wrap)
            Dim text As String = recommendations(i)
            Dim rect As New RectangleF(50, y, 390, 60)
            e.Graphics.DrawString(text, bodyFont, iconBrush, rect)

            y += 60
        Next

        ' Draw a box with suggested monthly budget
        Dim suggestedBudget As Decimal = totalSpending * 0.9 ' Suggest 10% less than current spending

        Using boxBrush As New SolidBrush(Color.FromArgb(45, 52, 64))
            Dim boxRect As New Rectangle(50, y, 350, 80)
            e.Graphics.FillRectangle(boxBrush, boxRect)
            e.Graphics.DrawRectangle(New Pen(Color.FromArgb(0, 173, 181), 2), boxRect)

            ' Draw budget suggestion
            Using brush As New SolidBrush(Color.White)
                e.Graphics.DrawString("Suggested Monthly Budget:", titleFont, brush, 70, y + 15)
                e.Graphics.DrawString(suggestedBudget.ToString("C0"), New Font("Segoe UI", 16, FontStyle.Bold), brush, 70, y + 40)
            End Using
        End Using
    End Sub

    ' Button event handlers
    Private Sub OnButtonMouseEnter(sender As Object, e As EventArgs)
        Dim btn As Button = CType(sender, Button)
        btn.BackColor = Color.FromArgb(0, 150, 160)
    End Sub

    Private Sub OnButtonMouseLeave(sender As Object, e As EventArgs)
        Dim btn As Button = CType(sender, Button)
        btn.BackColor = Color.FromArgb(0, 173, 181)
    End Sub

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

    ' Nested class to represent spending anomalies
    Public Class SpendingAnomaly
        Public Property Category As String
        Public Property CurrentAmount As Decimal
        Public Property PreviousAmount As Decimal
        Public Property PercentChange As Double
    End Class
End Class