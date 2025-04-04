Imports System.Windows.Forms
Imports System.Drawing

Partial Public Class Dashboard
    ' Declare controls
    Friend WithEvents pnlMenu As Panel
    Friend WithEvents pnlMain As Panel
    Friend WithEvents BtnDashboard As Button
    Friend WithEvents BtnSalary As Button
    Friend WithEvents BtnTestDB As Button
    Friend WithEvents BtnReports As Button
    Friend WithEvents BtnAnalysis As Button
    Friend WithEvents BtnLogout As Button

    ' Financial summary controls
    Friend WithEvents contentPanel As Panel
    Friend WithEvents summaryPanel As Panel
    Friend WithEvents salaryBox As Panel
    Friend WithEvents salaryValueLabel As Label
    Friend WithEvents expensesBox As Panel
    Friend WithEvents expensesValueLabel As Label
    Friend WithEvents remainingBox As Panel
    Friend WithEvents remainingValueLabel As Label

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    Private Sub InitializeComponent()
        ' Form setup
        Me.Text = "Expense Tracker"
        Me.Size = New Size(1200, 800)
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.BackColor = Color.FromArgb(34, 40, 49)

        ' Menu panel
        pnlMenu = New Panel()
        pnlMenu.BackColor = Color.FromArgb(45, 52, 64)
        pnlMenu.Dock = DockStyle.Left
        pnlMenu.Width = 200
        Me.Controls.Add(pnlMenu)

        ' Dashboard button
        BtnDashboard = CreateMenuButton("Dashboard", 20)
        pnlMenu.Controls.Add(BtnDashboard)

        ' Salary button
        BtnSalary = CreateMenuButton("Salary", 80)
        pnlMenu.Controls.Add(BtnSalary)

        ' Reports button
        BtnReports = CreateMenuButton("Reports", 140)
        pnlMenu.Controls.Add(BtnReports)

        ' Analysis button
        BtnAnalysis = CreateMenuButton("Analysis", 200)
        pnlMenu.Controls.Add(BtnAnalysis)

        ' Logout button
        BtnLogout = CreateMenuButton("Logout", 260)
        BtnLogout.BackColor = Color.FromArgb(255, 77, 77)  ' Red color for logout
        pnlMenu.Controls.Add(BtnLogout)

        ' Main content panel
        pnlMain = New Panel()
        pnlMain.BackColor = Color.FromArgb(57, 62, 70)
        pnlMain.Location = New Point(pnlMenu.Width, 0)
        pnlMain.Size = New Size(Me.Width - pnlMenu.Width, Me.Height)
        Me.Controls.Add(pnlMain)
    End Sub

    Private Function CreateMenuButton(text As String, top As Integer) As Button
        Dim btn As New Button()
        btn.Text = text
        btn.Location = New Point(10, top)
        btn.Size = New Size(180, 40)
        btn.FlatStyle = FlatStyle.Flat
        btn.FlatAppearance.BorderSize = 0
        btn.BackColor = Color.FromArgb(0, 173, 181)
        btn.ForeColor = Color.White
        btn.Font = New Font("Segoe UI", 12, FontStyle.Bold)
        btn.Cursor = Cursors.Hand

        Return btn
    End Function

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