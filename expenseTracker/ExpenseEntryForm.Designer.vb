Imports System.Windows.Forms
Imports System.Drawing

Partial Public Class ExpenseEntryForm
    ' Control declarations
    Friend WithEvents txtName As TextBox
    Friend WithEvents txtDescription As TextBox
    Friend WithEvents txtAmount As TextBox
    Friend WithEvents cmbCategory As ComboBox
    Friend WithEvents btnAddExpense As Button
    Friend WithEvents btnDeleteExpense As Button
    Friend WithEvents dgvExpenses As DataGridView

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    Private Sub InitializeComponent()
        ' Form setup
        Me.Text = "Expense Entry"
        Me.Size = New Size(1000, 600)
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.BackColor = Color.FromArgb(34, 40, 49)

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
        Me.Controls.Add(btnDeleteExpense)

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