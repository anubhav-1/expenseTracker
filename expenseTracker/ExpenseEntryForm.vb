Imports System.Windows.Forms
Imports System.Data.OleDb
Imports System.Linq
Imports System.Text

Partial Public Class ExpenseEntryForm
    Inherits Form

    ' Database connection
    Private connectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\ExpenseTracker.accdb;Persist Security Info=False;"

    ' Define events for dashboard communication
    Public Event ExpenseAdded(sender As Object, e As EventArgs)
    Public Event ExpenseDeleted(sender As Object, e As EventArgs)

    Public Sub New()
        ' This call is required by the designer
        InitializeComponent()

        ' Add initialization after the InitializeComponent() call
        LoadHardcodedCategories()
        LoadExpenses()
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

    Private Sub btnAddExpense_Click(sender As Object, e As EventArgs) Handles btnAddExpense.Click
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

    Private Sub btnDeleteExpense_Click(sender As Object, e As EventArgs) Handles btnDeleteExpense.Click
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
    Private Sub Button_MouseEnter(sender As Object, e As EventArgs) Handles btnAddExpense.MouseEnter
        Dim btn As Button = CType(sender, Button)
        btn.BackColor = Color.FromArgb(0, 150, 160)
    End Sub

    Private Sub Button_MouseLeave(sender As Object, e As EventArgs) Handles btnAddExpense.MouseLeave
        Dim btn As Button = CType(sender, Button)
        btn.BackColor = Color.FromArgb(0, 173, 181)
    End Sub

    ' Button hover effects for delete button
    Private Sub DeleteButton_MouseEnter(sender As Object, e As EventArgs) Handles btnDeleteExpense.MouseEnter
        Dim btn As Button = CType(sender, Button)
        btn.BackColor = Color.FromArgb(220, 50, 50) ' Darker red on hover
    End Sub

    Private Sub DeleteButton_MouseLeave(sender As Object, e As EventArgs) Handles btnDeleteExpense.MouseLeave
        Dim btn As Button = CType(sender, Button)
        btn.BackColor = Color.FromArgb(255, 77, 77) ' Original red
    End Sub
End Class