Imports System.Windows.Forms
Imports System.Data.OleDb
Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.Diagnostics

Public Class MonthlyReportGenerator
    ' Database connection
    Private connectionString As String

    ' UI Controls
    Private dgvReportData As DataGridView
    Private pnlPieChart As Panel
    Private pnlBarChart As Panel

    ' Chart data
    Private monthlyData As New Dictionary(Of String, Decimal)

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
        LoadMonthlyData(dateRange)

        ' Refresh charts
        pnlPieChart.Invalidate()
        pnlBarChart.Invalidate()
    End Sub

    ' Set up the data grid for monthly report
    Private Sub SetupGrid()
        dgvReportData.Columns.Clear()
        dgvReportData.Columns.Add("Month", "Month")
        dgvReportData.Columns.Add("TotalExpenses", "Total Expenses")
        dgvReportData.Columns.Add("AvgPerDay", "Avg. Per Day")

        dgvReportData.Columns("TotalExpenses").DefaultCellStyle.Format = "C"
        dgvReportData.Columns("AvgPerDay").DefaultCellStyle.Format = "C"
    End Sub

    ' Load monthly data from database
    Private Sub LoadMonthlyData(dateRange As DateRange)
        ' Format dates for query
        Dim startDateStr As String = dateRange.GetFormattedStartDate()
        Dim endDateStr As String = dateRange.GetFormattedEndDate()

        ' Clear existing data
        monthlyData.Clear()
        dgvReportData.Rows.Clear()

        ' Get data from database
        Try
            Using connection As New OleDbConnection(connectionString)
                connection.Open()

                ' Query to get expenses by month - using Format to get month from date
                Dim query As String = "SELECT Format(Timestamp,'yyyy-mm') AS YearMonth, SUM(Amount) AS TotalAmount " &
                                     "FROM Expenses " &
                                     "WHERE Timestamp BETWEEN #" & startDateStr & "# AND #" & endDateStr & "# " &
                                     "GROUP BY Format(Timestamp,'yyyy-mm') " &
                                     "ORDER BY YearMonth"

                Using command As New OleDbCommand(query, connection)
                    Using reader As OleDbDataReader = command.ExecuteReader()
                        While reader.Read()
                            Dim yearMonth As String = reader("YearMonth").ToString()
                            Dim amount As Decimal = Convert.ToDecimal(reader("TotalAmount"))

                            ' Parse year and month
                            Dim parts As String() = yearMonth.Split("-"c)
                            If parts.Length >= 2 Then
                                Dim year As Integer = Convert.ToInt32(parts(0))
                                Dim month As Integer = Convert.ToInt32(parts(1))

                                ' Format month name
                                Dim monthDate As New DateTime(year, month, 1)
                                Dim monthName As String = monthDate.ToString("MMM yyyy")

                                ' Store for chart
                                monthlyData.Add(monthName, amount)

                                ' Calculate average per day
                                Dim daysInMonth As Integer = DateTime.DaysInMonth(year, month)
                                Dim avgPerDay As Decimal = amount / daysInMonth

                                ' Add to grid
                                dgvReportData.Rows.Add(monthName, amount, avgPerDay)
                            End If
                        End While
                    End Using
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show("Database query error: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Debug.WriteLine("Database error: " & ex.Message & vbCrLf & ex.StackTrace)
        End Try
    End Sub

    ' Draw the pie chart (placeholder - not really useful for monthly data)
    Private Sub OnDrawPieChart(sender As Object, e As PaintEventArgs)
        ' Draw instruction message since pie chart isn't ideal for monthly data
        Using brush As New SolidBrush(Color.White)
            e.Graphics.DrawString("Monthly Expense Breakdown", New Font("Segoe UI", 16, FontStyle.Bold), brush, 50, 50)
            e.Graphics.DrawString("A pie chart is not the ideal visualization for monthly trends.", New Font("Segoe UI", 12), brush, 50, 100)
            e.Graphics.DrawString("Please refer to the bar chart for a better visualization of monthly data.", New Font("Segoe UI", 12), brush, 50, 130)
        End Using
    End Sub

    ' Draw the bar chart (monthly trend)
    Private Sub OnDrawBarChart(sender As Object, e As PaintEventArgs)
        ' Check if we have data
        If monthlyData Is Nothing OrElse monthlyData.Count = 0 Then
            ' Draw message when no data available
            Using brush As New SolidBrush(Color.White)
                e.Graphics.DrawString("No data available for the selected time period", New Font("Segoe UI", 12), brush, 50, 150)
            End Using
            Return
        End If

        ' Initialize graphics
        e.Graphics.SmoothingMode = SmoothingMode.AntiAlias

        ' Define chart area
        Dim chartRect As New Rectangle(50, 50, 390, 280)

        ' Draw title
        Using brush As New SolidBrush(Color.White)
            e.Graphics.DrawString("Monthly Expense Trend", New Font("Segoe UI", 16, FontStyle.Bold), brush, 120, 10)
        End Using

        ' Find the maximum value for scaling
        Dim maxValue As Decimal = If(monthlyData.Values.Count > 0, monthlyData.Values.Max(), 0)
        If maxValue = 0 Then
            Using brush As New SolidBrush(Color.White)
                e.Graphics.DrawString("No expenses in the selected time period", New Font("Segoe UI", 12), brush, 50, 150)
            End Using
            Return
        End If

        ' Round up maxValue to nearest 1000 or 100 for nicer scale
        If maxValue > 10000 Then
            maxValue = Math.Ceiling(maxValue / 1000) * 1000
        ElseIf maxValue > 1000 Then
            maxValue = Math.Ceiling(maxValue / 500) * 500
        Else
            maxValue = Math.Ceiling(maxValue / 100) * 100
        End If

        ' Draw axes
        Using axisPen As New Pen(Color.White, 2)
            ' X-axis
            e.Graphics.DrawLine(axisPen, chartRect.Left, chartRect.Bottom, chartRect.Right, chartRect.Bottom)

            ' Y-axis
            e.Graphics.DrawLine(axisPen, chartRect.Left, chartRect.Top, chartRect.Left, chartRect.Bottom)

            ' Y-axis labels
            Using brush As New SolidBrush(Color.White)
                Dim font As New Font("Segoe UI", 8)

                ' Draw value labels on Y-axis
                For i As Integer = 0 To 4
                    Dim y As Integer = chartRect.Bottom - (i * chartRect.Height / 4)
                    Dim value As Decimal = maxValue * i / 4
                    e.Graphics.DrawLine(axisPen, chartRect.Left - 5, y, chartRect.Left, y)
                    e.Graphics.DrawString(value.ToString("C0"), font, brush, chartRect.Left - 45, y - 7)
                Next
            End Using
        End Using

        ' Draw bars
        Dim barWidth As Integer = 40
        Dim spacing As Integer = 60
        Dim x As Integer = chartRect.Left + 30
        Dim barColor As Color = Color.FromArgb(0, 173, 181)  ' Teal

        ' Order by date (months are already in order they were added from query)
        For Each kvp As KeyValuePair(Of String, Decimal) In monthlyData
            Dim monthName As String = kvp.Key
            Dim amount As Decimal = kvp.Value

            ' Skip if zero
            If amount = 0 Then Continue For

            ' Scale value to chart height
            Dim barHeight As Integer = CInt((amount / maxValue) * chartRect.Height)

            ' Draw bar
            Using brush As New SolidBrush(barColor)
                Dim barRect As New Rectangle(x, chartRect.Bottom - barHeight, barWidth, barHeight)
                e.Graphics.FillRectangle(brush, barRect)
                e.Graphics.DrawRectangle(Pens.White, barRect)
            End Using

            ' Draw label
            Using brush As New SolidBrush(Color.White)
                ' Rotate text for x-axis labels
                e.Graphics.TranslateTransform(x + barWidth / 2, chartRect.Bottom + 5)
                e.Graphics.RotateTransform(45)
                e.Graphics.DrawString(monthName, New Font("Segoe UI", 8), brush, 0, 0)
                e.Graphics.ResetTransform()
            End Using

            ' Draw amount on top of bar
            Using brush As New SolidBrush(Color.White)
                Dim amountText As String = amount.ToString("C0")
                Dim textSize As SizeF = e.Graphics.MeasureString(amountText, New Font("Segoe UI", 8))
                Dim textX As Single = x + (barWidth - textSize.Width) / 2
                Dim textY As Single = chartRect.Bottom - barHeight - textSize.Height - 5

                ' Only draw if there's enough space
                If barHeight > textSize.Height + 10 Then
                    e.Graphics.DrawString(amountText, New Font("Segoe UI", 8), brush, textX, textY)
                End If
            End Using

            x += spacing
        Next

        ' Draw trend line if we have more than one month
        If monthlyData.Count > 1 Then
            DrawTrendLine(e, chartRect, maxValue)
        End If
    End Sub

    ' Draw a trend line connecting the tops of the bars
    Private Sub DrawTrendLine(e As PaintEventArgs, chartRect As Rectangle, maxValue As Decimal)
        Dim trendPen As New Pen(Color.FromArgb(255, 255, 255), 2)
        trendPen.DashStyle = DashStyle.Dot

        Dim barWidth As Integer = 40
        Dim spacing As Integer = 60
        Dim x As Integer = chartRect.Left + 30 + barWidth / 2  ' Center of first bar

        Dim points As New List(Of Point)

        ' Collect points for each bar top
        For Each kvp As KeyValuePair(Of String, Decimal) In monthlyData
            Dim amount As Decimal = kvp.Value

            ' Skip if zero (but still advance x position)
            If amount = 0 Then
                x += spacing
                Continue For
            End If

            ' Scale value to chart height
            Dim barHeight As Integer = CInt((amount / maxValue) * chartRect.Height)

            ' Add point at top center of bar
            points.Add(New Point(x, chartRect.Bottom - barHeight))

            x += spacing
        Next

        ' Draw line connecting the points if we have at least 2
        If points.Count >= 2 Then
            e.Graphics.DrawLines(trendPen, points.ToArray())

            ' Draw small circles at each point
            Dim pointBrush As New SolidBrush(Color.White)
            For Each point As Point In points
                e.Graphics.FillEllipse(pointBrush, point.X - 3, point.Y - 3, 6, 6)
            Next
        End If
    End Sub
End Class