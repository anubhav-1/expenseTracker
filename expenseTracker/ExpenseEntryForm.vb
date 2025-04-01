Imports System.Windows.Forms
Imports System.Data.OleDb
Imports System.Linq
Imports System.Text

Public Class ExpenseEntryForm
    Inherits Form

    ' Control declarations
    Private WithEvents txtName As TextBox
    Private WithEvents txtDescription As TextBox
    Private WithEvents txtAmount As TextBox
    Private WithEvents cmbCategory As ComboBox
    Private WithEvents btnAddExpense As Button
    Private WithEvents btnDeleteExpense As Button
    Private WithEvents dgvExpenses As DataGridView

    ' Database connection
    Private connectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\ExpenseTracker.accdb;Persist Security Info=False;"

    ' Define events for dashboard communication
    Public Event ExpenseAdded(sender As Object, e As EventArgs)
    Public Event ExpenseDeleted(sender As Object, e As EventArgs)

    Public Sub New()
        ' Form setup
        Me.Text = "Expense Entry"
        Me.Size = New Size(1000, 600)
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.BackColor = Color.FromArgb(34, 40, 49)

        ' Initialize all controls
        InitializeControls()
        SetupDataGridView()
        LoadHardcodedCategories()
        LoadExpenses()
    End Sub

    Private Sub InitializeControls()
        ' Name input
        Dim lblName As New Label()
        lblName.Text = "Expense Name:"
        lblName.Location = New Point(20, 20)
        lblName.AutoSize = True
        lblName.ForeColor = Color.White
        lblName.Font = New Font("Segoe UI", 12, FontStyle.Bold)
        Me.Controls.Add(lblName)

        txtName = New TextBox()
        txtName.Location = New Point(150, 20)
        txtName.Size = New Size(200, 30)
        txtName.BackColor = Color.FromArgb(57, 62, 70)
        txtName.ForeColor = Color.White
        txtName.Font = New Font("Segoe UI", 12)
        txtName.BorderStyle = BorderStyle.FixedSingle
        Me.Controls.Add(txtName)

        ' Description input
        Dim lblDescription As New Label()
        lblDescription.Text = "Description:"
        lblDescription.Location = New Point(20, 70)
        lblDescription.AutoSize = True
        lblDescription.ForeColor = Color.White
        lblDescription.Font = New Font("Segoe UI", 12, FontStyle.Bold)
        Me.Controls.Add(lblDescription)

        txtDescription = New TextBox()
        txtDescription.Location = New Point(150, 70)
        txtDescription.Size = New Size(300, 30)
        txtDescription.BackColor = Color.FromArgb(57, 62, 70)
        txtDescription.ForeColor = Color.White
        txtDescription.Font = New Font("Segoe UI", 12)
        txtDescription.BorderStyle = BorderStyle.FixedSingle
        Me.Controls.Add(txtDescription)

        ' Amount input
        Dim lblAmount As New Label()
        lblAmount.Text = "Amount:"
        lblAmount.Location = New Point(20, 120)
        lblAmount.AutoSize = True
        lblAmount.ForeColor = Color.White
        lblAmount.Font = New Font("Segoe UI", 12, FontStyle.Bold)
        Me.Controls.Add(lblAmount)

        txtAmount = New TextBox()
        txtAmount.Location = New Point(150, 120)
        txtAmount.Size = New Size(150, 30)
        txtAmount.BackColor = Color.FromArgb(57, 62, 70)
        txtAmount.ForeColor = Color.White
        txtAmount.Font = New Font("Segoe UI", 12)
        txtAmount.BorderStyle = BorderStyle.FixedSingle
        Me.Controls.Add(txtAmount)

        ' Category input
        Dim lblCategory As New Label()
        lblCategory.Text = "Category:"
        lblCategory.Location = New Point(20, 170)
        lblCategory.AutoSize = True
        lblCategory.ForeColor = Color.White
        lblCategory.Font = New Font("Segoe UI", 12, FontStyle.Bold)
        Me.Controls.Add(lblCategory)

        cmbCategory = New ComboBox()
        cmbCategory.Location = New Point(150, 170)
        cmbCategory.Size = New Size(200, 30)
        cmbCategory.BackColor = Color.FromArgb(57, 62, 70)
        cmbCategory.ForeColor = Color.White
        cmbCategory.Font = New Font("Segoe UI", 12)
        cmbCategory.DropDownStyle = ComboBoxStyle.DropDownList
        Me.Controls.Add(cmbCategory)

        ' Add expense button
        btnAddExpense = New Button()
        btnAddExpense.Text = "Add Expense"
        btnAddExpense.Location = New Point(150, 220)
        btnAddExpense.Size = New Size(150, 40)
        btnAddExpense.FlatStyle = FlatStyle.Flat
        btnAddExpense.FlatAppearance.BorderSize = 0
        btnAddExpense.BackColor = Color.FromArgb(0, 173, 181)
        btnAddExpense.ForeColor = Color.White
        btnAddExpense.Font = New Font("Segoe UI", 12, FontStyle.Bold)
        btnAddExpense.Cursor = Cursors.Hand
        AddHandler btnAddExpense.Click, AddressOf btnAddExpense_Click
        AddHandler btnAddExpense.MouseEnter, AddressOf Button_MouseEnter
        AddHandler btnAddExpense.MouseLeave, AddressOf Button_MouseLeave
        Me.Controls.Add(btnAddExpense)

        ' Delete expense button
        btnDeleteExpense = New Button()
        btnDeleteExpense.Text = "Delete Selected"
        btnDeleteExpense.Location = New Point(320, 220)
        btnDeleteExpense.Size = New Size(150, 40)
        btnDeleteExpense.FlatStyle = FlatStyle.Flat
        btnDeleteExpense.FlatAppearance.BorderSize = 0
        btnDeleteExpense.BackColor = Color.FromArgb(255, 77, 77) ' Red color for delete button
        btnDeleteExpense.ForeColor = Color.White
        btnDeleteExpense.Font = New Font("Segoe UI", 12, FontStyle.Bold)
        btnDeleteExpense.Cursor = Cursors.Hand
        AddHandler btnDeleteExpense.Click, AddressOf btnDeleteExpense_Click
        AddHandler btnDeleteExpense.MouseEnter, AddressOf DeleteButton_MouseEnter
        AddHandler btnDeleteExpense.MouseLeave, AddressOf DeleteButton_MouseLeave
        Me.Controls.Add(btnDeleteExpense)
    End Sub

    Private Sub SetupDataGridView()
        ' Initialize DataGridView
        dgvExpenses = New DataGridView()
        dgvExpenses.Location = New Point(20, 280)
        dgvExpenses.Size = New Size(900, 250)
        dgvExpenses.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        dgvExpenses.BackgroundColor = Color.FromArgb(45, 52, 64)
        dgvExpenses.ForeColor = Color.White
        dgvExpenses.Font = New Font("Segoe UI", 10)

        ' Set black headers with white text
        dgvExpenses.ColumnHeadersDefaultCellStyle.BackColor = Color.Black
        dgvExpenses.ColumnHeadersDefaultCellStyle.ForeColor = Color.White
        dgvExpenses.ColumnHeadersDefaultCellStyle.Font = New Font("Segoe UI", 12, FontStyle.Bold)

        dgvExpenses.DefaultCellStyle.BackColor = Color.FromArgb(45, 52, 64)
        dgvExpenses.DefaultCellStyle.ForeColor = Color.White
        dgvExpenses.DefaultCellStyle.SelectionBackColor = Color.FromArgb(0, 150, 160)
        dgvExpenses.DefaultCellStyle.SelectionForeColor = Color.White
        dgvExpenses.EnableHeadersVisualStyles = False
        dgvExpenses.RowHeadersVisible = False
        dgvExpenses.SelectionMode = DataGridViewSelectionMode.FullRowSelect ' Allow selecting entire rows
        dgvExpenses.MultiSelect = False ' Allow only one row to be selected at a time
        dgvExpenses.Columns.Add("ID", "ID")
        dgvExpenses.Columns.Add("Name", "Name")
        dgvExpenses.Columns.Add("Description", "Description")
        dgvExpenses.Columns.Add("Amount", "Amount")
        dgvExpenses.Columns.Add("Category", "Category")
        dgvExpenses.Columns.Add("Timestamp", "Date")
        dgvExpenses.Columns("ID").Visible = False

        ' Enable double buffering to reduce flicker
        Dim dgvType As Type = dgvExpenses.GetType()
        Dim pi As Reflection.PropertyInfo = dgvType.GetProperty("DoubleBuffered", Reflection.BindingFlags.Instance Or Reflection.BindingFlags.NonPublic)
        pi.SetValue(dgvExpenses, True, Nothing)

        Me.Controls.Add(dgvExpenses)
    End Sub

    Private Sub LoadHardcodedCategories()
        ' Clear existing items
        cmbCategory.Items.Clear()

        ' Add hardcoded categories as simple strings
        cmbCategory.Items.Add("Food")
        cmbCategory.Items.Add("Travel")
        cmbCategory.Items.Add("EMI")
        cmbCategory.Items.Add("Essentials")
        cmbCategory.Items.Add("Others")

        ' Select the first item
        cmbCategory.SelectedIndex = 0
    End Sub

    Private Sub btnAddExpense_Click(sender As Object, e As EventArgs)
        ' Validate input
        If String.IsNullOrWhiteSpace(txtName.Text) Then
            MessageBox.Show("Please enter an expense name.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        Dim amount As Double
        If Not Double.TryParse(txtAmount.Text, amount) Then
            MessageBox.Show("Please enter a valid amount.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        If cmbCategory.SelectedItem Is Nothing Then
            MessageBox.Show("Please select a category.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        ' Get the selected category as string
        Dim selectedCategory As String = cmbCategory.SelectedItem.ToString()

        ' Add expense to the database
        Try
            Using connection As New OleDbConnection(connectionString)
                connection.Open()

                ' Create the SQL statement with parameters
                Dim sql As String = "INSERT INTO Expenses ([Name], [Description], [Amount], [Category], [Timestamp]) VALUES (?, ?, ?, ?, ?)"

                Using cmd As New OleDbCommand(sql, connection)
                    ' Add parameters
                    cmd.Parameters.Add(New OleDbParameter("Name", OleDbType.VarChar)).Value = txtName.Text
                    cmd.Parameters.Add(New OleDbParameter("Description", OleDbType.VarChar)).Value = txtDescription.Text
                    cmd.Parameters.Add(New OleDbParameter("Amount", OleDbType.Double)).Value = amount
                    cmd.Parameters.Add(New OleDbParameter("Category", OleDbType.VarChar)).Value = selectedCategory
                    cmd.Parameters.Add(New OleDbParameter("Timestamp", OleDbType.Date)).Value = DateTime.Now

                    ' Execute the command
                    cmd.ExecuteNonQuery()
                End Using
            End Using

            ' Clear input fields
            txtName.Clear()
            txtDescription.Clear()
            txtAmount.Clear()

            ' Reload expenses
            LoadExpenses()

            ' Raise the ExpenseAdded event to notify the Dashboard
            RaiseEvent ExpenseAdded(Me, New EventArgs())

            MessageBox.Show("Expense added successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            MessageBox.Show("Error adding expense: " & ex.Message, "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnDeleteExpense_Click(sender As Object, e As EventArgs)
        ' Check if a row is selected
        If dgvExpenses.SelectedRows.Count = 0 Then
            MessageBox.Show("Please select an expense to delete.", "Selection Required", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If

        ' Get the ID of the selected expense
        Dim selectedRow As DataGridViewRow = dgvExpenses.SelectedRows(0)
        Dim expenseId As String = selectedRow.Cells("ID").Value.ToString()

        ' Confirm deletion
        Dim result As DialogResult = MessageBox.Show("Are you sure you want to delete this expense?", "Confirm Deletion",
                                                   MessageBoxButtons.YesNo, MessageBoxIcon.Question)

        If result = DialogResult.Yes Then
            ' Delete the expense from the database
            Try
                Using connection As New OleDbConnection(connectionString)
                    connection.Open()

                    ' Create the DELETE SQL statement
                    Dim sql As String = "DELETE FROM Expenses WHERE ID = ?"

                    Using cmd As New OleDbCommand(sql, connection)
                        ' Add parameter
                        cmd.Parameters.Add(New OleDbParameter("ID", OleDbType.Integer)).Value = Convert.ToInt32(expenseId)

                        ' Execute the command
                        Dim rowsAffected As Integer = cmd.ExecuteNonQuery()

                        If rowsAffected > 0 Then
                            MessageBox.Show("Expense deleted successfully.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)

                            ' Reload expenses
                            LoadExpenses()

                            ' Raise the ExpenseDeleted event to notify the Dashboard
                            RaiseEvent ExpenseDeleted(Me, New EventArgs())
                        Else
                            MessageBox.Show("No expense was deleted. The record may have been removed already.", "Warning",
                                           MessageBoxButtons.OK, MessageBoxIcon.Warning)
                        End If
                    End Using
                End Using
            Catch ex As Exception
                MessageBox.Show("Error deleting expense: " & ex.Message, "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End If
    End Sub

    Private Sub LoadExpenses()
        ' Temporarily suspend layout to prevent flickering
        dgvExpenses.SuspendLayout()

        ' Clear existing rows
        dgvExpenses.Rows.Clear()

        ' Load expenses from the database
        Try
            Using connection As New OleDbConnection(connectionString)
                connection.Open()
                ' Using simple query now that Timestamp is a proper Date/Time field
                Dim query As String = "SELECT [ID], [Name], [Description], [Amount], [Category], [Timestamp] FROM [Expenses] ORDER BY [Timestamp] DESC"

                Using command As New OleDbCommand(query, connection)
                    Using reader As OleDbDataReader = command.ExecuteReader()
                        While reader.Read()
                            ' Safe handling of potential NULL values
                            Dim id As String = If(IsDBNull(reader("ID")), "", reader("ID").ToString())
                            Dim name As String = If(IsDBNull(reader("Name")), "", reader("Name").ToString())
                            Dim description As String = If(IsDBNull(reader("Description")), "", reader("Description").ToString())

                            ' Safe handling of Amount field
                            Dim amountStr As String = ""
                            If Not IsDBNull(reader("Amount")) Then
                                Dim amountValue As Double = Convert.ToDouble(reader("Amount"))
                                amountStr = amountValue.ToString("C")
                            End If

                            Dim category As String = If(IsDBNull(reader("Category")), "", reader("Category").ToString())

                            ' Safe handling of Timestamp field
                            Dim timestampStr As String = ""
                            If Not IsDBNull(reader("Timestamp")) Then
                                Dim timestampValue As DateTime = Convert.ToDateTime(reader("Timestamp"))
                                timestampStr = timestampValue.ToString("dd-MM-yyyy")
                            End If

                            Dim row As String() = {
                                id,
                                name,
                                description,
                                amountStr,
                                category,
                                timestampStr
                            }
                            dgvExpenses.Rows.Add(row)
                        End While
                    End Using
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show("Error loading expenses: " & ex.Message, "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            ' Resume layout and refresh
            dgvExpenses.ResumeLayout()
            dgvExpenses.Refresh()
        End Try
    End Sub

    ' Button hover effects for regular buttons
    Private Sub Button_MouseEnter(sender As Object, e As EventArgs)
        Dim btn As Button = CType(sender, Button)
        btn.BackColor = Color.FromArgb(0, 150, 160)
    End Sub

    Private Sub Button_MouseLeave(sender As Object, e As EventArgs)
        Dim btn As Button = CType(sender, Button)
        btn.BackColor = Color.FromArgb(0, 173, 181)
    End Sub

    ' Button hover effects for delete button
    Private Sub DeleteButton_MouseEnter(sender As Object, e As EventArgs)
        Dim btn As Button = CType(sender, Button)
        btn.BackColor = Color.FromArgb(220, 50, 50) ' Darker red on hover
    End Sub

    Private Sub DeleteButton_MouseLeave(sender As Object, e As EventArgs)
        Dim btn As Button = CType(sender, Button)
        btn.BackColor = Color.FromArgb(255, 77, 77) ' Original red
    End Sub
End Class