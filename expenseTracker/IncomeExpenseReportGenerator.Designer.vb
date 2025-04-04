Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.Windows.Forms

Partial Public Class IncomeExpenseReportGenerator
    ' Draw the pie chart - with legend at the bottom
    Private Sub OnDrawPieChart(sender As Object, e As PaintEventArgs)
        ' Check if data has been loaded
        If Not dataLoaded Then
            ' Draw message when no data available
            Using brush As New SolidBrush(Color.White)
                e.Graphics.DrawString("No data available for the selected time period", New Font("Segoe UI", 12), brush, 50, 150)
            End Using
            Return
        End If

        ' Initialize graphics
        e.Graphics.SmoothingMode = SmoothingMode.AntiAlias

        ' Draw title
        Using brush As New SolidBrush(Color.White)
            e.Graphics.DrawString("Income vs Expenses Distribution", New Font("Segoe UI", 16, FontStyle.Bold), brush, 50, 20)
        End Using

        ' Count total Income and Expenses from the grid
        Dim totalIncome As Decimal = 0
        Dim totalExpenses As Decimal = 0

        For Each row As DataGridViewRow In dgvReportData.Rows
            If Not row.IsNewRow Then
                totalIncome += Convert.ToDecimal(row.Cells("Income").Value)
                totalExpenses += Convert.ToDecimal(row.Cells("Expenses").Value)
            End If
        Next

        ' If no data, show message
        If totalIncome = 0 And totalExpenses = 0 Then
            Using brush As New SolidBrush(Color.White)
                e.Graphics.DrawString("No income or expense data for the selected time period", New Font("Segoe UI", 12), brush, 50, 150)
            End Using
            Return
        End If

        ' Define chart area - make it larger
        Dim chartRect As New Rectangle(75, 60, 300, 250)

        ' Draw pie for Income vs Expenses
        Dim total As Decimal = totalIncome + totalExpenses
        Dim incomeAngle As Single = CSng(360 * totalIncome / total)
        Dim expensesAngle As Single = CSng(360 * totalExpenses / total)

        ' Draw Income slice
        Using incomeBrush As New SolidBrush(Color.FromArgb(0, 173, 181))
            e.Graphics.FillPie(incomeBrush, chartRect, 0, incomeAngle)
        End Using

        ' Draw Expenses slice
        Using expensesBrush As New SolidBrush(Color.FromArgb(255, 77, 77))
            e.Graphics.FillPie(expensesBrush, chartRect, incomeAngle, expensesAngle)
        End Using

        ' Draw legend at the bottom
        Dim legendY As Integer = chartRect.Bottom + 20

        ' Income legend
        Using brush As New SolidBrush(Color.FromArgb(0, 173, 181))
            e.Graphics.FillRectangle(brush, New Rectangle(100, legendY, 15, 15))
        End Using

        Using brush As New SolidBrush(Color.White)
            e.Graphics.DrawString($"Income: {totalIncome:C} ({totalIncome / total:P1})",
                             New Font("Segoe UI", 10), brush, 120, legendY)
        End Using

        ' Expenses legend
        Using brush As New SolidBrush(Color.FromArgb(255, 77, 77))
            e.Graphics.FillRectangle(brush, New Rectangle(100, legendY + 25, 15, 15))
        End Using

        Using brush As New SolidBrush(Color.White)
            e.Graphics.DrawString($"Expenses: {totalExpenses:C} ({totalExpenses / total:P1})",
                             New Font("Segoe UI", 10), brush, 120, legendY + 25)
        End Using

        ' Net income/loss
        Dim netAmount As Decimal = totalIncome - totalExpenses
        Dim netText As String = If(netAmount >= 0, "Net Savings:", "Net Loss:")
        Dim netColor As Color = If(netAmount >= 0, Color.FromArgb(76, 187, 23), Color.Red)

        Using brush As New SolidBrush(netColor)
            e.Graphics.DrawString($"{netText} {Math.Abs(netAmount):C}", New Font("Segoe UI", 12, FontStyle.Bold),
                             brush, 100, legendY + 55)
        End Using
    End Sub

    ' Draw the bar chart - with legend at the bottom
    Private Sub OnDrawBarChart(sender As Object, e As PaintEventArgs)
        ' Check if data has been loaded
        If Not dataLoaded Then
            ' Draw message when no data available
            Using brush As New SolidBrush(Color.White)
                e.Graphics.DrawString("No data available for the selected time period", New Font("Segoe UI", 12), brush, 50, 150)
            End Using
            Return
        End If

        ' Initialize graphics
        e.Graphics.SmoothingMode = SmoothingMode.AntiAlias

        ' Draw title
        Using brush As New SolidBrush(Color.White)
            e.Graphics.DrawString("Monthly Income vs Expenses", New Font("Segoe UI", 16, FontStyle.Bold), brush, 100, 20)
        End Using

        ' Check if we have data
        If dgvReportData.Rows.Count = 0 Then
            Using brush As New SolidBrush(Color.White)
                e.Graphics.DrawString("No data for the selected time period", New Font("Segoe UI", 12), brush, 50, 150)
            End Using
            Return
        End If

        ' Define chart area - make it larger
        Dim chartRect As New Rectangle(50, 60, 380, 250)

        ' Find the maximum value for scaling
        Dim maxValue As Decimal = 0
        For Each row As DataGridViewRow In dgvReportData.Rows
            If Not row.IsNewRow Then
                Dim income As Decimal = Convert.ToDecimal(row.Cells("Income").Value)
                Dim expenses As Decimal = Convert.ToDecimal(row.Cells("Expenses").Value)
                maxValue = Math.Max(maxValue, Math.Max(income, expenses))
            End If
        Next

        ' Round up maxValue for nicer scale
        If maxValue > 10000 Then
            maxValue = Math.Ceiling(maxValue / 1000) * 1000
        ElseIf maxValue > 1000 Then
            maxValue = Math.Ceiling(maxValue / 500) * 500
        Else
            maxValue = Math.Ceiling(maxValue / 100) * 100
        End If

        ' If no data, exit
        If maxValue = 0 Then
            Using brush As New SolidBrush(Color.White)
                e.Graphics.DrawString("No income or expense data for the selected time period", New Font("Segoe UI", 12), brush, 50, 150)
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
            Using brush As New SolidBrush(Color.LightGray)
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

        ' Draw grouped bars
        Dim barWidth As Integer = 30
        Dim groupWidth As Integer = 80
        Dim x As Integer = chartRect.Left + 20

        ' Income and Expense colors
        Dim incomeColor As Color = Color.FromArgb(0, 173, 181)  ' Teal
        Dim expenseColor As Color = Color.FromArgb(255, 77, 77) ' Red

        ' Draw up to 5 months to avoid cluttering
        Dim rowsToShow As Integer = Math.Min(dgvReportData.Rows.Count, 5)
        For i As Integer = 0 To rowsToShow - 1
            Dim row As DataGridViewRow = dgvReportData.Rows(i)
            If row.IsNewRow Then Continue For

            Dim monthName As String = row.Cells("Month").Value.ToString()
            Dim income As Decimal = Convert.ToDecimal(row.Cells("Income").Value)
            Dim expenses As Decimal = Convert.ToDecimal(row.Cells("Expenses").Value)

            ' Calculate bar heights (scaled)
            Dim incomeHeight As Integer = CInt((income / maxValue) * chartRect.Height)
            Dim expensesHeight As Integer = CInt((expenses / maxValue) * chartRect.Height)

            ' Draw income bar
            Using brush As New SolidBrush(incomeColor)
                Dim incomeRect As New Rectangle(x, chartRect.Bottom - incomeHeight, barWidth, incomeHeight)
                e.Graphics.FillRectangle(brush, incomeRect)
                e.Graphics.DrawRectangle(Pens.White, incomeRect)
            End Using

            ' Draw expenses bar
            Using brush As New SolidBrush(expenseColor)
                Dim expensesRect As New Rectangle(x + barWidth + 5, chartRect.Bottom - expensesHeight, barWidth, expensesHeight)
                e.Graphics.FillRectangle(brush, expensesRect)
                e.Graphics.DrawRectangle(Pens.White, expensesRect)
            End Using

            ' Draw month label
            Using brush As New SolidBrush(Color.White)
                ' Rotate text for x-axis labels
                e.Graphics.TranslateTransform(x + barWidth, chartRect.Bottom + 5)
                e.Graphics.RotateTransform(45)
                e.Graphics.DrawString(monthName, New Font("Segoe UI", 8), brush, 0, 0)
                e.Graphics.ResetTransform()
            End Using

            x += groupWidth
        Next

        ' Draw legend at the bottom
        Dim legendY As Integer = chartRect.Bottom + 40

        ' Income legend
        Using brush As New SolidBrush(incomeColor)
            e.Graphics.FillRectangle(brush, New Rectangle(140, legendY, 15, 15))
        End Using

        Using brush As New SolidBrush(Color.White)
            e.Graphics.DrawString("Income", New Font("Segoe UI", 10), brush, 160, legendY)
        End Using

        ' Expenses legend
        Using brush As New SolidBrush(expenseColor)
            e.Graphics.FillRectangle(brush, New Rectangle(240, legendY, 15, 15))
        End Using

        Using brush As New SolidBrush(Color.White)
            e.Graphics.DrawString("Expenses", New Font("Segoe UI", 10), brush, 260, legendY)
        End Using
    End Sub
End Class