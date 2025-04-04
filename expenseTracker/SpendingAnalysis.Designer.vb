Imports System.Windows.Forms
Imports System.Drawing
Imports System.Drawing.Drawing2D

Partial Public Class SpendingAnalysis
    ' Control declarations
    Friend WithEvents pnlFilters As Panel
    Friend WithEvents lblTimeFrame As Label
    Friend WithEvents lblYear As Label
    Friend WithEvents lblMonth As Label
    Friend WithEvents cmbTimeFrame As ComboBox
    Friend WithEvents cmbYear As ComboBox
    Friend WithEvents cmbMonth As ComboBox
    Friend WithEvents btnAnalyze As Button

    ' Analysis panels
    Friend WithEvents pnlTopCategories As Panel
    Friend WithEvents pnlSavingOpportunities As Panel
    Friend WithEvents pnlSpendingTrends As Panel
    Friend WithEvents pnlBudgetRecommendations As Panel

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    Private Sub InitializeComponent()
        ' Form setup
        Me.Text = "Spending Analysis"
        Me.Size = New Size(1000, 800)
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.BackColor = Color.FromArgb(34, 40, 49)
        Me.FormBorderStyle = FormBorderStyle.None
        Me.Dock = DockStyle.Fill

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
        pnlFilters.Controls.Add(btnAnalyze)

        ' Create analysis panels
        CreateAnalysisPanels()
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