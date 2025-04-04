Imports System.Windows.Forms
Imports System.Data.OleDb
Imports System.Diagnostics
Imports System.Globalization
Imports System.Text

Partial Public Class Dashboard
    Inherits Form

    ' Reference to Login form
    Private loginForm As Login

    ' Database connection
    Private connectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\ExpenseTracker.accdb;Persist Security Info=False;"

    Public Sub New()
        ' This call is required by the designer
        InitializeComponent()

        ' Add custom initialization after InitializeComponent
        ShowDashboard()
    End Sub

    Public Sub New(loginForm As Login)
        Me.New()
        Me.loginForm = loginForm
    End Sub

    Private Sub BtnDashboard_Click(sender As Object, e As EventArgs) Handles BtnDashboard.Click
        ShowDashboard()
    End Sub

    Private Sub BtnSalary_Click(sender As Object, e As EventArgs) Handles BtnSalary.Click
        ShowSalary()
    End Sub

    ' Reports button handler
    Private Sub BtnReports_Click(sender As Object, e As EventArgs) Handles BtnReports.Click
        ShowReports()
    End Sub

    ' Analysis button handler
    Private Sub BtnAnalysis_Click(sender As Object, e As EventArgs) Handles BtnAnalysis.Click
        ShowAnalysis()
    End Sub

    ' Logout button handler
    Private Sub BtnLogout_Click(sender As Object, e As EventArgs) Handles BtnLogout.Click
        Dim result As DialogResult = MessageBox.Show("Are you sure you want to log out?", "Confirm Logout",
                                                  MessageBoxButtons.YesNo, MessageBoxIcon.Question)

        If result = DialogResult.Yes Then
            ' Show login form
            If Me.loginForm IsNot Nothing Then
                Me.loginForm.Show()
            Else
                ' Create new login form if reference is missing
                Dim newLogin As New Login()
                newLogin.Show()
            End If

            ' Close dashboard
            Me.Close()
        End If
    End Sub

    Private Sub ShowDashboard()
        pnlMain.Controls.Clear()

        ' Content panel
        contentPanel = New Panel()
        contentPanel.Dock = DockStyle.Fill
        contentPanel.BackColor = Color.FromArgb(57, 62, 70)
        pnlMain.Controls.Add(contentPanel)

        ' Create the financial summary with placeholder values
        CreateFinancialSummary()

        ' Force immediate calculation and update of values
        Try
            ' Calculate values
            Dim totalSalary As Decimal = CalculateTotalSalary()
            Dim totalExpenses As Decimal = CalculateTotalExpenses()
            Dim remaining As Decimal = totalSalary - totalExpenses

            ' Directly update the labels with the calculated values
            If salaryValueLabel IsNot Nothing Then
                salaryValueLabel.Text = totalSalary.ToString("C")
            End If

            If expensesValueLabel IsNot Nothing Then
                expensesValueLabel.Text = totalExpenses.ToString("C")
            End If

            If remainingValueLabel IsNot Nothing Then
                remainingValueLabel.Text = remaining.ToString("C")
                remainingValueLabel.ForeColor = If(remaining >= 0, Color.FromArgb(76, 187, 23), Color.Red)
            End If

        Catch ex As Exception
            MessageBox.Show("Error calculating financial summary: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        ' Expense form
        Dim expenseFormPanel As New Panel()
        expenseFormPanel.Dock = DockStyle.None
        expenseFormPanel.Location = New Point(0, 110)
        expenseFormPanel.Size = New Size(contentPanel.Width, contentPanel.Height - 110)
        contentPanel.Controls.Add(expenseFormPanel)

        Dim expenseForm As New ExpenseEntryForm()
        expenseForm.TopLevel = False
        expenseForm.FormBorderStyle = FormBorderStyle.None
        expenseForm.Dock = DockStyle.Fill
        AddHandler expenseForm.ExpenseAdded, AddressOf OnExpenseChanged
        AddHandler expenseForm.ExpenseDeleted, AddressOf OnExpenseChanged
        expenseFormPanel.Controls.Add(expenseForm)
        expenseForm.Show()
    End Sub

    Private Sub CreateFinancialSummary()
        summaryPanel = New Panel()
        summaryPanel.Dock = DockStyle.None
        summaryPanel.Location = New Point(0, 0)
        summaryPanel.Size = New Size(contentPanel.Width, 100)
        summaryPanel.BackColor = Color.FromArgb(45, 52, 64)
        contentPanel.Controls.Add(summaryPanel)

        ' Create the boxes with placeholder values initially
        salaryBox = CreateSummaryBox("Current Month Salary", "₹0.00", 20, Color.FromArgb(0, 173, 181))
        expensesBox = CreateSummaryBox("Total Expenses", "₹0.00", 280, Color.FromArgb(255, 77, 77))
        remainingBox = CreateSummaryBox("Remaining Budget", "₹0.00", 540, Color.FromArgb(76, 187, 23))
    End Sub

    Private Function CreateSummaryBox(title As String, initialValue As String, x As Integer, color As Color) As Panel
        Dim box As New Panel()
        box.Location = New Point(x, 20)
        box.Size = New Size(220, 60)
        box.BackColor = Color.FromArgb(34, 40, 49)

        Dim border As New Panel()
        border.Location = New Point(0, 0)
        border.Size = New Size(5, 60)
        border.BackColor = color
        box.Controls.Add(border)

        Dim lblTitle As New Label()
        lblTitle.Text = title
        lblTitle.Location = New Point(15, 10)
        lblTitle.ForeColor = Color.White
        lblTitle.Font = New Font("Segoe UI", 9, FontStyle.Bold)
        box.Controls.Add(lblTitle)

        ' Create value label with a name that can be referenced later
        Dim lblValue As New Label()
        lblValue.Text = initialValue
        lblValue.Location = New Point(15, 30)
        lblValue.Size = New Size(190, 25) ' Make sure the label is big enough
        lblValue.ForeColor = color
        lblValue.BackColor = Color.FromArgb(34, 40, 49) ' Ensure consistent background
        lblValue.Font = New Font("Segoe UI", 14, FontStyle.Bold)
        lblValue.Name = "ValueLabel"  ' Add a name so we can find it later
        box.Controls.Add(lblValue)

        ' Store a reference to the appropriate value label based on the box type
        If title = "Current Month Salary" Then
            salaryValueLabel = lblValue
        ElseIf title = "Total Expenses" Then
            expensesValueLabel = lblValue
        ElseIf title = "Remaining Budget" Then
            remainingValueLabel = lblValue
        End If

        summaryPanel.Controls.Add(box)
        Return box
    End Function

    Private Sub OnExpenseChanged(sender As Object, e As EventArgs)
        UpdateFinancialSummary()
    End Sub

    Private Sub UpdateFinancialSummary()
        ' Initialize with default empty values
        Dim totalSalary As Decimal = 0
        Dim totalExpenses As Decimal = 0
        Dim remaining As Decimal = 0

        ' Try to calculate actual values
        Try
            totalSalary = CalculateTotalSalary()
            totalExpenses = CalculateTotalExpenses()
            remaining = totalSalary - totalExpenses
        Catch ex As Exception
            Debug.WriteLine("Summary update error: " & ex.Message)
            ' Keep 0 values if calculation fails
        End Try

        ' Always show some value, even if it's zero - don't just show the currency symbol
        Dim salaryString As String = totalSalary.ToString("C")
        Dim expensesString As String = totalExpenses.ToString("C")
        Dim remainingString As String = remaining.ToString("C")

        ' Set the text directly on the stored label references with debugging
        Debug.WriteLine("Updating labels with: Salary=" & salaryString & ", Expenses=" & expensesString & ", Remaining=" & remainingString)

        If salaryValueLabel IsNot Nothing Then
            salaryValueLabel.Text = salaryString
            salaryValueLabel.ForeColor = Color.FromArgb(0, 173, 181)
            Debug.WriteLine("Updated salary label: " & salaryValueLabel.Text)
        Else
            Debug.WriteLine("Salary label is null!")
        End If

        If expensesValueLabel IsNot Nothing Then
            expensesValueLabel.Text = expensesString
            expensesValueLabel.ForeColor = Color.FromArgb(255, 77, 77)
            Debug.WriteLine("Updated expenses label: " & expensesValueLabel.Text)
        Else
            Debug.WriteLine("Expenses label is null!")
        End If

        If remainingValueLabel IsNot Nothing Then
            remainingValueLabel.Text = remainingString
            remainingValueLabel.ForeColor = If(remaining >= 0, Color.FromArgb(76, 187, 23), Color.Red)
            Debug.WriteLine("Updated remaining label: " & remainingValueLabel.Text)
        Else
            Debug.WriteLine("Remaining label is null!")
        End If
    End Sub

    Private Function CalculateTotalSalary() As Decimal
        Try
            Using connection As New OleDbConnection(connectionString)
                connection.Open()
                Dim query = "SELECT SUM(Amount) FROM Salary WHERE MONTH(Timestamp) = MONTH(DATE()) AND YEAR(Timestamp) = YEAR(DATE())"
                Using command As New OleDbCommand(query, connection)
                    Dim result = command.ExecuteScalar()
                    If result IsNot Nothing AndAlso Not IsDBNull(result) Then
                        Return Convert.ToDecimal(result)
                    End If
                End Using
            End Using
        Catch ex As Exception
            Debug.WriteLine("Salary calculation error: " & ex.Message)
            Throw ' Re-throw to handle in calling method
        End Try
        Return 0
    End Function

    Private Function CalculateTotalExpenses() As Decimal
        Try
            Using connection As New OleDbConnection(connectionString)
                connection.Open()
                Dim query = "SELECT SUM(Amount) FROM Expenses WHERE MONTH(Timestamp) = MONTH(DATE()) AND YEAR(Timestamp) = YEAR(DATE())"
                Using command As New OleDbCommand(query, connection)
                    Dim result = command.ExecuteScalar()
                    If result IsNot Nothing AndAlso Not IsDBNull(result) Then
                        Return Convert.ToDecimal(result)
                    End If
                End Using
            End Using
        Catch ex As Exception
            Debug.WriteLine("Expenses calculation error: " & ex.Message)
            Throw ' Re-throw to handle in calling method
        End Try
        Return 0
    End Function

    Private Sub ShowSalary()
        pnlMain.Controls.Clear()

        Dim salaryPanel As New Panel()
        salaryPanel.Dock = DockStyle.Fill
        salaryPanel.BackColor = Color.FromArgb(57, 62, 70)
        pnlMain.Controls.Add(salaryPanel)

        ' Update financial summary to ensure it's in sync
        UpdateFinancialSummary()

        ' Salary controls
        Dim lblAmount As New Label()
        lblAmount.Text = "Salary Amount:"
        lblAmount.Location = New Point(20, 20)
        lblAmount.ForeColor = Color.White
        lblAmount.Font = New Font("Segoe UI", 12, FontStyle.Bold)
        salaryPanel.Controls.Add(lblAmount)

        Dim txtAmount As New TextBox()
        txtAmount.Name = "txtSalaryAmount"
        txtAmount.Location = New Point(150, 20)
        txtAmount.Size = New Size(200, 30)
        txtAmount.BackColor = Color.FromArgb(45, 52, 64)
        txtAmount.ForeColor = Color.White
        txtAmount.Font = New Font("Segoe UI", 12)
        txtAmount.BorderStyle = BorderStyle.FixedSingle
        salaryPanel.Controls.Add(txtAmount)

        ' Add button
        Dim btnAdd As New Button()
        btnAdd.Text = "Add Salary"
        btnAdd.Location = New Point(150, 70)
        btnAdd.Size = New Size(120, 40)
        btnAdd.FlatStyle = FlatStyle.Flat
        btnAdd.FlatAppearance.BorderSize = 0
        btnAdd.BackColor = Color.FromArgb(0, 173, 181)
        btnAdd.ForeColor = Color.White
        btnAdd.Font = New Font("Segoe UI", 12, FontStyle.Bold)
        btnAdd.Cursor = Cursors.Hand
        AddHandler btnAdd.Click, AddressOf BtnAddSalary_Click
        AddHandler btnAdd.MouseEnter, AddressOf Button_MouseEnter
        AddHandler btnAdd.MouseLeave, AddressOf Button_MouseLeave
        salaryPanel.Controls.Add(btnAdd)

        ' DataGridView with dark theme
        Dim dgv As New DataGridView()
        dgv.Name = "dgvSalary"
        dgv.Location = New Point(20, 130)
        dgv.Size = New Size(900, 400)
        dgv.BackgroundColor = Color.FromArgb(45, 52, 64)
        dgv.ForeColor = Color.White
        dgv.BorderStyle = BorderStyle.None
        dgv.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.None
        dgv.ColumnHeadersDefaultCellStyle.BackColor = Color.Black
        dgv.ColumnHeadersDefaultCellStyle.ForeColor = Color.White
        dgv.ColumnHeadersDefaultCellStyle.Font = New Font("Segoe UI", 12, FontStyle.Bold)
        dgv.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        dgv.ColumnHeadersHeight = 40
        dgv.EnableHeadersVisualStyles = False
        dgv.DefaultCellStyle.BackColor = Color.FromArgb(57, 62, 70)
        dgv.DefaultCellStyle.ForeColor = Color.White
        dgv.DefaultCellStyle.Font = New Font("Segoe UI", 11)
        dgv.DefaultCellStyle.SelectionBackColor = Color.FromArgb(0, 150, 160)
        dgv.DefaultCellStyle.SelectionForeColor = Color.White
        dgv.RowHeadersVisible = False
        dgv.RowTemplate.Height = 35
        dgv.RowTemplate.DefaultCellStyle.Padding = New Padding(5, 0, 0, 0)
        dgv.GridColor = Color.FromArgb(34, 40, 49)
        dgv.AllowUserToAddRows = False
        dgv.AllowUserToResizeRows = False
        dgv.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        dgv.MultiSelect = False
        dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill

        ' Add columns
        Dim idCol = New DataGridViewTextBoxColumn()
        idCol.Name = "ID"
        idCol.HeaderText = "ID"
        idCol.Visible = False
        dgv.Columns.Add(idCol)

        Dim amountCol = New DataGridViewTextBoxColumn()
        amountCol.Name = "Amount"
        amountCol.HeaderText = "Amount"
        amountCol.DefaultCellStyle.Format = "C"
        amountCol.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        dgv.Columns.Add(amountCol)

        Dim dateCol = New DataGridViewTextBoxColumn()
        dateCol.Name = "Timestamp"
        dateCol.HeaderText = "Date"
        dateCol.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        dgv.Columns.Add(dateCol)

        salaryPanel.Controls.Add(dgv)

        ' Delete button
        Dim btnDelete As New Button()
        btnDelete.Text = "Delete Selected"
        btnDelete.Location = New Point(300, 70)
        btnDelete.Size = New Size(150, 40)
        btnDelete.FlatStyle = FlatStyle.Flat
        btnDelete.FlatAppearance.BorderSize = 0
        btnDelete.BackColor = Color.FromArgb(255, 77, 77)
        btnDelete.ForeColor = Color.White
        btnDelete.Font = New Font("Segoe UI", 12, FontStyle.Bold)
        btnDelete.Cursor = Cursors.Hand
        AddHandler btnDelete.Click, AddressOf BtnDeleteSalary_Click
        AddHandler btnDelete.MouseEnter, AddressOf DeleteButton_MouseEnter
        AddHandler btnDelete.MouseLeave, AddressOf DeleteButton_MouseLeave
        salaryPanel.Controls.Add(btnDelete)

        Dim dgvType As Type = dgv.GetType()
        Dim pi As Reflection.PropertyInfo = dgvType.GetProperty("DoubleBuffered", Reflection.BindingFlags.Instance Or Reflection.BindingFlags.NonPublic)
        pi.SetValue(dgv, True, Nothing)

        LoadSalaryHistory(dgv)

        ' Add handler to refresh data when the form becomes visible again
        AddHandler Me.VisibleChanged, Sub(s, e)
                                          If Me.Visible Then
                                              LoadSalaryHistory(dgv)
                                              UpdateFinancialSummary()
                                          End If
                                      End Sub
    End Sub

    Private Sub BtnAddSalary_Click(sender As Object, e As EventArgs)
        ' Find the parent panel and controls
        Dim salaryPanel = CType(sender.Parent, Panel)
        Dim txtAmount = CType(salaryPanel.Controls("txtSalaryAmount"), TextBox)
        Dim dgv = CType(salaryPanel.Controls("dgvSalary"), DataGridView)

        ' Apply double-buffering to DataGridView if not already done
        Try
            Dim dgvType As Type = dgv.GetType()
            Dim pi As Reflection.PropertyInfo = dgvType.GetProperty("DoubleBuffered", Reflection.BindingFlags.Instance Or Reflection.BindingFlags.NonPublic)
            pi.SetValue(dgv, True, Nothing)
        Catch ex As Exception
            Debug.WriteLine("Failed to apply double-buffering: " & ex.Message)
        End Try

        ' Validate input
        Dim input = txtAmount.Text.Trim()
        If String.IsNullOrEmpty(input) Then
            MessageBox.Show("Please enter a salary amount", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        Dim amount As Decimal
        If Not Decimal.TryParse(input, NumberStyles.Number, CultureInfo.InvariantCulture, amount) Then
            MessageBox.Show("Please enter a valid amount (e.g. 2500 or 3000.50)", "Invalid Input", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        If amount <= 0 Then
            MessageBox.Show("Amount must be greater than zero", "Invalid Amount", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        ' Database operation
        Try
            Using connection As New OleDbConnection(connectionString)
                connection.Open()
                Dim query = "INSERT INTO Salary (Amount, [Timestamp]) VALUES (?, ?)"

                Using command As New OleDbCommand(query, connection)
                    command.Parameters.Add("Amount", OleDbType.Decimal).Value = amount
                    command.Parameters.Add("Timestamp", OleDbType.Date).Value = DateTime.Now

                    Dim rows = command.ExecuteNonQuery()
                    If rows > 0 Then
                        ' Clear the input
                        txtAmount.Clear()

                        ' Use our dedicated refresh method
                        RefreshDataGridView(dgv)

                        ' Update financial summary
                        UpdateFinancialSummary()

                        ' Show success message
                        MessageBox.Show("Salary added successfully", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)

                        ' One more refresh after message box closes
                        RefreshDataGridView(dgv)
                    End If
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show($"Failed to add salary: {ex.Message}", "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub BtnDeleteSalary_Click(sender As Object, e As EventArgs)
        Dim salaryPanel = CType(sender.Parent, Panel)
        Dim dgv = CType(salaryPanel.Controls("dgvSalary"), DataGridView)

        If dgv.SelectedRows.Count = 0 Then
            MessageBox.Show("Please select a record to delete", "Selection Required", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If

        Dim id = Convert.ToInt32(dgv.SelectedRows(0).Cells("ID").Value)
        Dim confirm = MessageBox.Show("Delete this salary record?", "Confirm", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

        If confirm = DialogResult.Yes Then
            Try
                Using connection As New OleDbConnection(connectionString)
                    connection.Open()
                    Dim query = "DELETE FROM Salary WHERE ID = ?"

                    Using command As New OleDbCommand(query, connection)
                        command.Parameters.Add("ID", OleDbType.Integer).Value = id
                        Dim rows = command.ExecuteNonQuery()

                        If rows > 0 Then
                            ' Use our dedicated refresh method
                            RefreshDataGridView(dgv)

                            ' Update financial summary
                            UpdateFinancialSummary()

                            ' Show success message
                            MessageBox.Show("Record deleted", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)

                            ' One more refresh after message box closes
                            RefreshDataGridView(dgv)
                        End If
                    End Using
                End Using
            Catch ex As Exception
                MessageBox.Show($"Delete failed: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End If
    End Sub

    Private Sub LoadSalaryHistory(dgv As DataGridView)
        ' Make sure the DataGridView exists and is accessible
        If dgv Is Nothing OrElse dgv.IsDisposed Then
            Debug.WriteLine("DataGridView is null or disposed!")
            Return
        End If

        ' Temporarily suspend layout to prevent flickering
        dgv.SuspendLayout()

        ' Clear existing rows
        dgv.Rows.Clear()

        ' Reset the selection
        dgv.ClearSelection()

        Try
            Using connection As New OleDbConnection(connectionString)
                connection.Open()
                Dim query = "SELECT ID, Amount, Timestamp FROM Salary ORDER BY Timestamp DESC"

                Using command As New OleDbCommand(query, connection)
                    Using reader = command.ExecuteReader()
                        ' Flag to check if we have any data
                        Dim hasData As Boolean = False

                        While reader.Read()
                            hasData = True

                            Dim id = reader("ID").ToString()
                            Dim amount = Convert.ToDecimal(reader("Amount"))
                            Dim dateStr = Convert.ToDateTime(reader("Timestamp")).ToString("dd-MM-yyyy")

                            ' Add row with proper formatting
                            Dim row As String() = {id, amount.ToString("C"), dateStr}
                            dgv.Rows.Add(row)
                        End While

                        ' If no data found, log information
                        If Not hasData Then
                            Debug.WriteLine("No salary records found in database")
                        End If
                    End Using
                End Using
            End Using
        Catch ex As Exception
            Debug.WriteLine($"Error loading salary history: {ex.Message}")
            MessageBox.Show($"Error loading history: {ex.Message}", "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            ' Resume layout
            dgv.ResumeLayout()

            ' Force multiple levels of refresh
            dgv.Refresh()
            dgv.Update()

            ' If the grid is visible and has a parent, refresh the parent too
            If dgv.Visible AndAlso dgv.Parent IsNot Nothing Then
                dgv.Parent.Refresh()
                dgv.Parent.Update()
            End If

            ' Force UI thread to process painting events
            Application.DoEvents()
        End Try
    End Sub

    ' Method to show Reports & Analytics
    Private Sub ShowReports()
        pnlMain.Controls.Clear()

        Dim reportsPanel As New Panel()
        reportsPanel.Dock = DockStyle.Fill
        reportsPanel.BackColor = Color.FromArgb(57, 62, 70)
        pnlMain.Controls.Add(reportsPanel)

        ' Create and add the ReportsAnalytics form
        Dim reportsAnalytics As New ReportsAnalytics()
        reportsAnalytics.TopLevel = False
        reportsAnalytics.FormBorderStyle = FormBorderStyle.None
        reportsAnalytics.Dock = DockStyle.Fill
        reportsPanel.Controls.Add(reportsAnalytics)
        reportsAnalytics.Show()

        ' Update financial summary
        UpdateFinancialSummary()
    End Sub

    ' Method to show Spending Analysis tab
    Private Sub ShowAnalysis()
        pnlMain.Controls.Clear()

        Dim analysisPanel As New Panel()
        analysisPanel.Dock = DockStyle.Fill
        analysisPanel.BackColor = Color.FromArgb(57, 62, 70)
        pnlMain.Controls.Add(analysisPanel)

        ' Create and add the SpendingAnalysis form
        Dim spendingAnalysis As New SpendingAnalysis()
        spendingAnalysis.TopLevel = False
        spendingAnalysis.FormBorderStyle = FormBorderStyle.None
        spendingAnalysis.Dock = DockStyle.Fill
        analysisPanel.Controls.Add(spendingAnalysis)
        spendingAnalysis.Show()

        ' Update financial summary
        UpdateFinancialSummary()
    End Sub

    Private Sub TestDatabaseConnection()
        Try
            Using connection As New OleDbConnection(connectionString)
                connection.Open()
                MessageBox.Show("Database connection successful", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End Using
        Catch ex As Exception
            MessageBox.Show($"Connection failed: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' Complete solution for direct DataGridView refresh
    Public Sub RefreshDataGridView(dgv As DataGridView)
        If dgv Is Nothing Then Return

        ' Force data reload
        LoadSalaryHistory(dgv)

        ' Multiple UI refresh approaches
        dgv.Invalidate()
        dgv.Refresh()
        dgv.Update()

        ' Force message processing
        Application.DoEvents()
    End Sub

    ' Button hover effects
    Private Sub Button_MouseEnter(sender As Object, e As EventArgs)
        Dim btn = CType(sender, Button)
        If btn.BackColor <> Color.FromArgb(255, 77, 77) Then
            btn.BackColor = Color.FromArgb(0, 150, 160)
        End If
    End Sub

    Private Sub Button_MouseLeave(sender As Object, e As EventArgs)
        Dim btn = CType(sender, Button)
        If btn.BackColor <> Color.FromArgb(255, 77, 77) Then
            btn.BackColor = Color.FromArgb(0, 173, 181)
        End If
    End Sub

    Private Sub DeleteButton_MouseEnter(sender As Object, e As EventArgs)
        Dim btn = CType(sender, Button)
        btn.BackColor = Color.FromArgb(220, 50, 50)
    End Sub

    Private Sub DeleteButton_MouseLeave(sender As Object, e As EventArgs)
        Dim btn = CType(sender, Button)
        btn.BackColor = Color.FromArgb(255, 77, 77)
    End Sub

    ' Logout button hover effects
    Private Sub LogoutButton_MouseEnter(sender As Object, e As EventArgs) Handles BtnLogout.MouseEnter
        BtnLogout.BackColor = Color.FromArgb(220, 50, 50)
    End Sub

    Private Sub LogoutButton_MouseLeave(sender As Object, e As EventArgs) Handles BtnLogout.MouseLeave
        BtnLogout.BackColor = Color.FromArgb(255, 77, 77)
    End Sub
End Class