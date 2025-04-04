Imports System
Imports System.Windows.Forms
Imports System.Data.OleDb
Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.Diagnostics

Public Class CategoryReportGenerator
    ' Database connection
    Private connectionString As String

    ' UI Controls
    Private dgvReportData As DataGridView
    Private pnlPieChart As Panel
    Private pnlBarChart As Panel

    ' Chart data
    Private categoryData As New Dictionary(Of String, Decimal)

    ' Flag to track if data has been loaded
    Private dataLoaded As Boolean = False

    ' Constructor
    Public Sub New(connectionString As String, dgvReportData As DataGridView, pnlPieChart As Panel, pnlBarChart As Panel)
        Me.connectionString = connectionString
        Me.dgvReportData = dgvReportData
        Me.pnlPieChart = pnlPieChart
        Me.pnlBarChart = pnlBarChart

        ' Set up chart event handlers - but don't load data automatically
        AddHandler pnlPieChart.Paint, AddressOf OnDrawPieChart
        AddHandler pnlBarChart.Paint, AddressOf OnDrawBarChart
    End Sub

    ' Main method to generate the report
    Public Sub GenerateReport(dateRange As DateRange)
        ' Set up the data grid
        SetupGrid()

        ' Load data
        LoadCategoryData(dateRange)

        ' Mark that data has been loaded
        dataLoaded = True

        ' Refresh charts
        pnlPieChart.Invalidate()
        pnlBarChart.Invalidate()
    End Sub

    ' Set up the data grid for category report
    Private Sub SetupGrid()
        dgvReportData.Columns.Clear()
        dgvReportData.Columns.Add("Category", "Category")
        dgvReportData.Columns.Add("Amount", "Amount")
        dgvReportData.Columns.Add("Percentage", "Percentage")

        dgvReportData.Columns("Amount").DefaultCellStyle.Format = "C"
        dgvReportData.Columns("Percentage").DefaultCellStyle.Format = "P2"
    End Sub

    ' Load category data from database
    Private Sub LoadCategoryData(dateRange As DateRange)
        ' Format dates for query using consistent MM/dd/yyyy format for Access
        Dim startDateStr As String = dateRange.GetFormattedStartDate()
        Dim endDateStr As String = dateRange.GetFormattedEndDate()

        ' Debug info
        Debug.WriteLine($"Loading expense categories from {startDateStr} to {endDateStr}")

        ' Clear existing data
        categoryData.Clear()
        dgvReportData.Rows.Clear()

        ' Get data from database
        Try
            Using connection As New OleDbConnection(connectionString)
                connection.Open()

                ' Get total expenses first
                Dim totalExpenses As Decimal = 0
                Dim totalQuery As String = "SELECT SUM([Amount]) FROM Expenses WHERE [Timestamp] BETWEEN #" & startDateStr & "# AND #" & endDateStr & "#"

                Debug.WriteLine("Total query: " + totalQuery)

                Using command As New OleDbCommand(totalQuery, connection)
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

                ' Query to get expenses by category
                Dim query As String = "SELECT [Category], SUM([Amount]) AS TotalAmount " &
                                     "FROM Expenses " &
                                     "WHERE [Timestamp] BETWEEN #" & startDateStr & "# AND #" & endDateStr & "# " &
                                     "GROUP BY [Category] " &
                                     "ORDER BY TotalAmount DESC"

                Debug.WriteLine("Category query: " + query)

                ' Execute the query to get expenses by category
                Using command As New OleDbCommand(query, connection)
                    Using reader As OleDbDataReader = command.ExecuteReader()
                        While reader.Read()
                            Dim category As String = If(IsDBNull(reader("Category")), "Uncategorized", reader("Category").ToString())
                            Dim amount As Decimal = Convert.ToDecimal(reader("TotalAmount"))
                            Dim percentage As Double = If(totalExpenses = 0, 0, Convert.ToDouble(amount / totalExpenses))

                            Debug.WriteLine($"Category: {category}, Amount: {amount}, Percentage: {percentage}")

                            ' Add to dictionary for charts
                            categoryData(category) = amount

                            ' Add to grid
                            dgvReportData.Rows.Add(category, amount, percentage)
                        End While
                    End Using
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show("Database query error: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Debug.WriteLine("Database error: " & ex.Message & vbCrLf & ex.StackTrace)
        End Try
    End Sub

    ' Draw the pie chart
    Private Sub OnDrawPieChart(sender As Object, e As PaintEventArgs)
        ' Check if data has been loaded - if not, don't try to draw anything
        If Not dataLoaded Then
            Return
        End If

        ' Check if we have data
        If categoryData Is Nothing OrElse categoryData.Count = 0 Then
            ' Draw message when no data available
            Using brush As New SolidBrush(Color.White)
                e.Graphics.DrawString("No data available for the selected time period", New Font("Segoe UI", 12), brush, 50, 150)
            End Using
            Return
        End If

        ' Initialize graphics and brushes
        e.Graphics.SmoothingMode = SmoothingMode.AntiAlias

        ' Define chart area
        Dim chartRect As New Rectangle(50, 50, 300, 300)

        ' Calculate total for percentages
        Dim total As Decimal = categoryData.Values.Sum()
        If total = 0 Then
            Using brush As New SolidBrush(Color.White)
                e.Graphics.DrawString("No expenses in the selected time period", New Font("Segoe UI", 12), brush, 50, 150)
            End Using
            Return
        End If

        ' Colors for pie slices
        Dim colors As Color() = {
            Color.FromArgb(0, 173, 181),   ' Teal
            Color.FromArgb(255, 77, 77),   ' Red
            Color.FromArgb(76, 187, 23),   ' Green
            Color.FromArgb(255, 190, 11),  ' Yellow
            Color.FromArgb(153, 102, 255), ' Purple
            Color.FromArgb(58, 134, 255),  ' Blue
            Color.FromArgb(255, 128, 0),   ' Orange
            Color.FromArgb(0, 210, 180),   ' Turquoise
            Color.FromArgb(240, 98, 146),  ' Pink
            Color.FromArgb(124, 179, 66)   ' Lime
        }

        ' Draw title
        Using brush As New SolidBrush(Color.White)
            e.Graphics.DrawString("Expense Distribution by Category", New Font("Segoe UI", 16, FontStyle.Bold), brush, 50, 10)
        End Using

        ' Draw legend first
        Dim legendY As Integer = 50
        Dim colorIndex As Integer = 0

        For Each kvp As KeyValuePair(Of String, Decimal) In categoryData
            Dim category As String = kvp.Key
            Dim amount As Decimal = kvp.Value

            ' Skip if zero
            If amount = 0 Then Continue For

            Dim percentage As Single = CSng(amount / total * 100)

            ' Draw legend item
            Using brush As New SolidBrush(colors(colorIndex Mod colors.Length))
                ' Color box
                e.Graphics.FillRectangle(brush, New Rectangle(370, legendY, 15, 15))

                ' Category name and percentage
                Using textBrush As New SolidBrush(Color.White)
                    Dim legendText As String = $"{category}: {amount.ToString("C")}"
                    If percentage > 0 Then
                        legendText &= $" ({percentage:0.0}%)"
                    End If
                    e.Graphics.DrawString(legendText, New Font("Segoe UI", 9), textBrush, 390, legendY)
                End Using
            End Using

            legendY += 25
            colorIndex += 1

            ' Limit legend items to prevent overflow
            If colorIndex > 10 Then Exit For
        Next

        ' Now draw the pie slices
        colorIndex = 0
        Dim startAngle As Single = 0

        For Each kvp As KeyValuePair(Of String, Decimal) In categoryData
            Dim amount As Decimal = kvp.Value

            ' Skip if zero
            If amount = 0 Then Continue For

            Dim percentage As Single = CSng(amount / total * 100)
            Dim sweepAngle As Single = CSng(360 * percentage / 100)

            Using brush As New SolidBrush(colors(colorIndex Mod colors.Length))
                e.Graphics.FillPie(brush, chartRect, startAngle, sweepAngle)
                e.Graphics.DrawPie(Pens.White, chartRect, startAngle, sweepAngle)
            End Using

            startAngle += sweepAngle
            colorIndex += 1
        Next
    End Sub

    ' Draw the bar chart (category comparison)
    Private Sub OnDrawBarChart(sender As Object, e As PaintEventArgs)
        ' Check if data has been loaded - if not, don't try to draw anything
        If Not dataLoaded Then
            Return
        End If

        ' Check if we have data
        If categoryData Is Nothing OrElse categoryData.Count = 0 Then
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
            e.Graphics.DrawString("Expense Amounts by Category", New Font("Segoe UI", 16, FontStyle.Bold), brush, 80, 10)
        End Using

        ' Find the maximum value for scaling
        Dim maxValue As Decimal = If(categoryData.Values.Count > 0, categoryData.Values.Max(), 0)
        If maxValue = 0 Then
            Using brush As New SolidBrush(Color.White)
                e.Graphics.DrawString("No expenses in the selected time period", New Font("Segoe UI", 12), brush, 50, 150)
            End Using
            Return
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
        Dim barWidth As Integer = 30
        Dim x As Integer = chartRect.Left + 30
        Dim spacing As Integer = 60
        Dim barColor As Color = Color.FromArgb(0, 173, 181)  ' Teal

        ' Get top 6 categories to avoid cluttering
        Dim topCategories = categoryData.OrderByDescending(Function(kvp) kvp.Value).Take(6)

        ' Draw bars for each category
        For Each kvp As KeyValuePair(Of String, Decimal) In topCategories
            Dim category As String = kvp.Key
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
                e.Graphics.DrawString(category, New Font("Segoe UI", 8), brush, 0, 0)
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

        ' Draw note about top categories
        If categoryData.Count > 6 Then
            Using brush As New SolidBrush(Color.LightGray)
                e.Graphics.DrawString("* Showing top 6 categories by amount", New Font("Segoe UI", 8), brush, chartRect.Left, chartRect.Bottom + 50)
            End Using
        End If
    End Sub
End Class