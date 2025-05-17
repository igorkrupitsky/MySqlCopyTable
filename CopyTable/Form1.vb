Imports System.Data.OleDb
Imports System.Data.Odbc

Public Class Form1
	Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

	Public Sub New()
		MyBase.New()

		'This call is required by the Windows Form Designer.
		InitializeComponent()

		'Add any initialization after the InitializeComponent() call

	End Sub

	'Form overrides dispose to clean up the component list.
	Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
		If disposing Then
			If Not (components Is Nothing) Then
				components.Dispose()
			End If
		End If
		MyBase.Dispose(disposing)
	End Sub

	'Required by the Windows Form Designer
	Private components As System.ComponentModel.IContainer

	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.  
	'Do not modify it using the code editor.
	Friend WithEvents btnExport As System.Windows.Forms.Button
	Friend WithEvents btnConnect As System.Windows.Forms.Button
	Friend WithEvents txtConnect As System.Windows.Forms.TextBox
	Friend WithEvents btnCancel As System.Windows.Forms.Button
	Friend WithEvents ProgressBar1 As System.Windows.Forms.ProgressBar
	Friend WithEvents Label1 As System.Windows.Forms.Label
	Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
	Friend WithEvents txtTest As System.Windows.Forms.Button
	Friend WithEvents txtPassword As System.Windows.Forms.TextBox
	Friend WithEvents txtUser As System.Windows.Forms.TextBox
	Friend WithEvents txtDatabase As System.Windows.Forms.TextBox
	Friend WithEvents txtServer As System.Windows.Forms.TextBox
	Friend WithEvents Label5 As System.Windows.Forms.Label
	Friend WithEvents Label4 As System.Windows.Forms.Label
	Friend WithEvents Label3 As System.Windows.Forms.Label
	Friend WithEvents Label2 As System.Windows.Forms.Label
	Friend WithEvents cboDriver As System.Windows.Forms.ComboBox
	Friend WithEvents txtLog As System.Windows.Forms.TextBox
	Friend WithEvents lbCount As System.Windows.Forms.Label
	Friend WithEvents txtPort As TextBox
	Friend WithEvents Label7 As Label
	Friend WithEvents btnSearch As Button
	Friend WithEvents txtSearch As TextBox
	Friend WithEvents dgTables As DataGridView
	Friend WithEvents ProgressBar2 As ProgressBar
	Friend WithEvents btnStop As LinkLabel
	Friend WithEvents chkDropTable As CheckBox
	Friend WithEvents chkDeleteData As CheckBox
	Friend WithEvents chkCreateTable As CheckBox
	Friend WithEvents txtSearchMax As TextBox
	Friend WithEvents Label8 As Label
	Friend WithEvents Label9 As Label
	Friend WithEvents chkChangedRecords As CheckBox
	Friend WithEvents chkNotCopied As CheckBox
	Friend WithEvents Label6 As System.Windows.Forms.Label
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Me.btnExport = New System.Windows.Forms.Button()
		Me.txtConnect = New System.Windows.Forms.TextBox()
		Me.btnConnect = New System.Windows.Forms.Button()
		Me.btnCancel = New System.Windows.Forms.Button()
		Me.ProgressBar1 = New System.Windows.Forms.ProgressBar()
		Me.Label1 = New System.Windows.Forms.Label()
		Me.GroupBox1 = New System.Windows.Forms.GroupBox()
		Me.txtPort = New System.Windows.Forms.TextBox()
		Me.Label7 = New System.Windows.Forms.Label()
		Me.cboDriver = New System.Windows.Forms.ComboBox()
		Me.Label6 = New System.Windows.Forms.Label()
		Me.txtTest = New System.Windows.Forms.Button()
		Me.txtPassword = New System.Windows.Forms.TextBox()
		Me.txtUser = New System.Windows.Forms.TextBox()
		Me.txtDatabase = New System.Windows.Forms.TextBox()
		Me.txtServer = New System.Windows.Forms.TextBox()
		Me.Label5 = New System.Windows.Forms.Label()
		Me.Label4 = New System.Windows.Forms.Label()
		Me.Label3 = New System.Windows.Forms.Label()
		Me.Label2 = New System.Windows.Forms.Label()
		Me.txtLog = New System.Windows.Forms.TextBox()
		Me.lbCount = New System.Windows.Forms.Label()
		Me.btnSearch = New System.Windows.Forms.Button()
		Me.txtSearch = New System.Windows.Forms.TextBox()
		Me.dgTables = New System.Windows.Forms.DataGridView()
		Me.ProgressBar2 = New System.Windows.Forms.ProgressBar()
		Me.btnStop = New System.Windows.Forms.LinkLabel()
		Me.chkDropTable = New System.Windows.Forms.CheckBox()
		Me.chkDeleteData = New System.Windows.Forms.CheckBox()
		Me.chkCreateTable = New System.Windows.Forms.CheckBox()
		Me.txtSearchMax = New System.Windows.Forms.TextBox()
		Me.Label8 = New System.Windows.Forms.Label()
		Me.Label9 = New System.Windows.Forms.Label()
		Me.chkChangedRecords = New System.Windows.Forms.CheckBox()
		Me.chkNotCopied = New System.Windows.Forms.CheckBox()
		Me.GroupBox1.SuspendLayout()
		CType(Me.dgTables, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.SuspendLayout()
		'
		'btnExport
		'
		Me.btnExport.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
		Me.btnExport.Location = New System.Drawing.Point(805, 603)
		Me.btnExport.Name = "btnExport"
		Me.btnExport.Size = New System.Drawing.Size(160, 34)
		Me.btnExport.TabIndex = 0
		Me.btnExport.Text = "Copy tables"
		'
		'txtConnect
		'
		Me.txtConnect.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
			Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
		Me.txtConnect.BackColor = System.Drawing.SystemColors.HighlightText
		Me.txtConnect.Location = New System.Drawing.Point(94, 12)
		Me.txtConnect.Multiline = True
		Me.txtConnect.Name = "txtConnect"
		Me.txtConnect.Size = New System.Drawing.Size(967, 35)
		Me.txtConnect.TabIndex = 1
		'
		'btnConnect
		'
		Me.btnConnect.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
		Me.btnConnect.Location = New System.Drawing.Point(1071, 13)
		Me.btnConnect.Name = "btnConnect"
		Me.btnConnect.Size = New System.Drawing.Size(69, 34)
		Me.btnConnect.TabIndex = 3
		Me.btnConnect.Text = "..."
		'
		'btnCancel
		'
		Me.btnCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
		Me.btnCancel.Location = New System.Drawing.Point(978, 602)
		Me.btnCancel.Name = "btnCancel"
		Me.btnCancel.Size = New System.Drawing.Size(167, 35)
		Me.btnCancel.TabIndex = 6
		Me.btnCancel.Text = "Cancel"
		'
		'ProgressBar1
		'
		Me.ProgressBar1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
			Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
		Me.ProgressBar1.Location = New System.Drawing.Point(322, 924)
		Me.ProgressBar1.Name = "ProgressBar1"
		Me.ProgressBar1.Size = New System.Drawing.Size(739, 12)
		Me.ProgressBar1.TabIndex = 7
		'
		'Label1
		'
		Me.Label1.AutoSize = True
		Me.Label1.Location = New System.Drawing.Point(24, 20)
		Me.Label1.Name = "Label1"
		Me.Label1.Size = New System.Drawing.Size(54, 20)
		Me.Label1.TabIndex = 8
		Me.Label1.Text = "To DB"
		'
		'GroupBox1
		'
		Me.GroupBox1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
			Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
		Me.GroupBox1.Controls.Add(Me.txtPort)
		Me.GroupBox1.Controls.Add(Me.Label7)
		Me.GroupBox1.Controls.Add(Me.cboDriver)
		Me.GroupBox1.Controls.Add(Me.Label6)
		Me.GroupBox1.Controls.Add(Me.txtTest)
		Me.GroupBox1.Controls.Add(Me.txtPassword)
		Me.GroupBox1.Controls.Add(Me.txtUser)
		Me.GroupBox1.Controls.Add(Me.txtDatabase)
		Me.GroupBox1.Controls.Add(Me.txtServer)
		Me.GroupBox1.Controls.Add(Me.Label5)
		Me.GroupBox1.Controls.Add(Me.Label4)
		Me.GroupBox1.Controls.Add(Me.Label3)
		Me.GroupBox1.Controls.Add(Me.Label2)
		Me.GroupBox1.Location = New System.Drawing.Point(13, 53)
		Me.GroupBox1.Name = "GroupBox1"
		Me.GroupBox1.Size = New System.Drawing.Size(1129, 303)
		Me.GroupBox1.TabIndex = 9
		Me.GroupBox1.TabStop = False
		Me.GroupBox1.Text = "From MySQL DB"
		'
		'txtPort
		'
		Me.txtPort.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
			Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
		Me.txtPort.Location = New System.Drawing.Point(102, 149)
		Me.txtPort.Name = "txtPort"
		Me.txtPort.Size = New System.Drawing.Size(1014, 26)
		Me.txtPort.TabIndex = 12
		Me.txtPort.Text = "3306"
		'
		'Label7
		'
		Me.Label7.AutoSize = True
		Me.Label7.Location = New System.Drawing.Point(11, 153)
		Me.Label7.Name = "Label7"
		Me.Label7.Size = New System.Drawing.Size(38, 20)
		Me.Label7.TabIndex = 11
		Me.Label7.Text = "Port"
		'
		'cboDriver
		'
		Me.cboDriver.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
			Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
		Me.cboDriver.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
		Me.cboDriver.FormattingEnabled = True
		Me.cboDriver.Location = New System.Drawing.Point(102, 25)
		Me.cboDriver.Name = "cboDriver"
		Me.cboDriver.Size = New System.Drawing.Size(1014, 28)
		Me.cboDriver.TabIndex = 10
		'
		'Label6
		'
		Me.Label6.AutoSize = True
		Me.Label6.Location = New System.Drawing.Point(11, 29)
		Me.Label6.Name = "Label6"
		Me.Label6.Size = New System.Drawing.Size(50, 20)
		Me.Label6.TabIndex = 9
		Me.Label6.Text = "Driver"
		'
		'txtTest
		'
		Me.txtTest.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
		Me.txtTest.Location = New System.Drawing.Point(996, 263)
		Me.txtTest.Name = "txtTest"
		Me.txtTest.Size = New System.Drawing.Size(120, 34)
		Me.txtTest.TabIndex = 8
		Me.txtTest.Text = "Connect"
		Me.txtTest.UseVisualStyleBackColor = True
		'
		'txtPassword
		'
		Me.txtPassword.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
			Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
		Me.txtPassword.Location = New System.Drawing.Point(102, 231)
		Me.txtPassword.Name = "txtPassword"
		Me.txtPassword.Size = New System.Drawing.Size(1014, 26)
		Me.txtPassword.TabIndex = 7
		'
		'txtUser
		'
		Me.txtUser.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
			Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
		Me.txtUser.Location = New System.Drawing.Point(102, 192)
		Me.txtUser.Name = "txtUser"
		Me.txtUser.Size = New System.Drawing.Size(1014, 26)
		Me.txtUser.TabIndex = 6
		'
		'txtDatabase
		'
		Me.txtDatabase.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
			Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
		Me.txtDatabase.Location = New System.Drawing.Point(102, 110)
		Me.txtDatabase.Name = "txtDatabase"
		Me.txtDatabase.Size = New System.Drawing.Size(1014, 26)
		Me.txtDatabase.TabIndex = 5
		'
		'txtServer
		'
		Me.txtServer.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
			Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
		Me.txtServer.Location = New System.Drawing.Point(102, 64)
		Me.txtServer.Name = "txtServer"
		Me.txtServer.Size = New System.Drawing.Size(1014, 26)
		Me.txtServer.TabIndex = 4
		'
		'Label5
		'
		Me.Label5.AutoSize = True
		Me.Label5.Location = New System.Drawing.Point(11, 235)
		Me.Label5.Name = "Label5"
		Me.Label5.Size = New System.Drawing.Size(78, 20)
		Me.Label5.TabIndex = 3
		Me.Label5.Text = "Password"
		'
		'Label4
		'
		Me.Label4.AutoSize = True
		Me.Label4.Location = New System.Drawing.Point(11, 196)
		Me.Label4.Name = "Label4"
		Me.Label4.Size = New System.Drawing.Size(43, 20)
		Me.Label4.TabIndex = 2
		Me.Label4.Text = "User"
		'
		'Label3
		'
		Me.Label3.AutoSize = True
		Me.Label3.Location = New System.Drawing.Point(11, 114)
		Me.Label3.Name = "Label3"
		Me.Label3.Size = New System.Drawing.Size(79, 20)
		Me.Label3.TabIndex = 1
		Me.Label3.Text = "Database"
		'
		'Label2
		'
		Me.Label2.AutoSize = True
		Me.Label2.Location = New System.Drawing.Point(11, 69)
		Me.Label2.Name = "Label2"
		Me.Label2.Size = New System.Drawing.Size(55, 20)
		Me.Label2.TabIndex = 0
		Me.Label2.Text = "Server"
		'
		'txtLog
		'
		Me.txtLog.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
			Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
		Me.txtLog.Location = New System.Drawing.Point(7, 684)
		Me.txtLog.Multiline = True
		Me.txtLog.Name = "txtLog"
		Me.txtLog.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
		Me.txtLog.Size = New System.Drawing.Size(1145, 234)
		Me.txtLog.TabIndex = 12
		'
		'lbCount
		'
		Me.lbCount.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
		Me.lbCount.AutoSize = True
		Me.lbCount.Location = New System.Drawing.Point(1115, 920)
		Me.lbCount.Name = "lbCount"
		Me.lbCount.Size = New System.Drawing.Size(45, 20)
		Me.lbCount.TabIndex = 14
		Me.lbCount.Text = "0000"
		'
		'btnSearch
		'
		Me.btnSearch.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
		Me.btnSearch.Location = New System.Drawing.Point(978, 645)
		Me.btnSearch.Name = "btnSearch"
		Me.btnSearch.Size = New System.Drawing.Size(167, 34)
		Me.btnSearch.TabIndex = 15
		Me.btnSearch.Text = "Search"
		'
		'txtSearch
		'
		Me.txtSearch.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
		Me.txtSearch.Location = New System.Drawing.Point(642, 652)
		Me.txtSearch.Name = "txtSearch"
		Me.txtSearch.Size = New System.Drawing.Size(157, 26)
		Me.txtSearch.TabIndex = 13
		'
		'dgTables
		'
		Me.dgTables.AllowUserToAddRows = False
		Me.dgTables.AllowUserToDeleteRows = False
		Me.dgTables.AllowUserToOrderColumns = True
		Me.dgTables.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
			Or System.Windows.Forms.AnchorStyles.Left) _
			Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
		Me.dgTables.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
		Me.dgTables.Location = New System.Drawing.Point(12, 356)
		Me.dgTables.Name = "dgTables"
		Me.dgTables.RowHeadersWidth = 62
		Me.dgTables.Size = New System.Drawing.Size(1128, 236)
		Me.dgTables.TabIndex = 21
		'
		'ProgressBar2
		'
		Me.ProgressBar2.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
		Me.ProgressBar2.Location = New System.Drawing.Point(7, 924)
		Me.ProgressBar2.Name = "ProgressBar2"
		Me.ProgressBar2.Size = New System.Drawing.Size(309, 16)
		Me.ProgressBar2.TabIndex = 39
		'
		'btnStop
		'
		Me.btnStop.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
		Me.btnStop.AutoSize = True
		Me.btnStop.Location = New System.Drawing.Point(1067, 918)
		Me.btnStop.Name = "btnStop"
		Me.btnStop.Size = New System.Drawing.Size(43, 20)
		Me.btnStop.TabIndex = 42
		Me.btnStop.TabStop = True
		Me.btnStop.Text = "Stop"
		Me.btnStop.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me.btnStop.Visible = False
		'
		'chkDropTable
		'
		Me.chkDropTable.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
		Me.chkDropTable.AutoSize = True
		Me.chkDropTable.Location = New System.Drawing.Point(13, 645)
		Me.chkDropTable.Name = "chkDropTable"
		Me.chkDropTable.Size = New System.Drawing.Size(165, 24)
		Me.chkDropTable.TabIndex = 46
		Me.chkDropTable.Text = "Drop table if exists"
		Me.chkDropTable.UseVisualStyleBackColor = True
		'
		'chkDeleteData
		'
		Me.chkDeleteData.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
		Me.chkDeleteData.AutoSize = True
		Me.chkDeleteData.Checked = True
		Me.chkDeleteData.CheckState = System.Windows.Forms.CheckState.Checked
		Me.chkDeleteData.Location = New System.Drawing.Point(187, 603)
		Me.chkDeleteData.Name = "chkDeleteData"
		Me.chkDeleteData.Size = New System.Drawing.Size(175, 24)
		Me.chkDeleteData.TabIndex = 45
		Me.chkDeleteData.Text = "Delete before insert"
		Me.chkDeleteData.UseVisualStyleBackColor = True
		'
		'chkCreateTable
		'
		Me.chkCreateTable.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
		Me.chkCreateTable.AutoSize = True
		Me.chkCreateTable.Checked = True
		Me.chkCreateTable.CheckState = System.Windows.Forms.CheckState.Checked
		Me.chkCreateTable.Location = New System.Drawing.Point(13, 604)
		Me.chkCreateTable.Name = "chkCreateTable"
		Me.chkCreateTable.Size = New System.Drawing.Size(168, 24)
		Me.chkCreateTable.TabIndex = 44
		Me.chkCreateTable.Text = "Create target table"
		Me.chkCreateTable.UseVisualStyleBackColor = True
		'
		'txtSearchMax
		'
		Me.txtSearchMax.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
		Me.txtSearchMax.Location = New System.Drawing.Point(543, 652)
		Me.txtSearchMax.Name = "txtSearchMax"
		Me.txtSearchMax.Size = New System.Drawing.Size(93, 26)
		Me.txtSearchMax.TabIndex = 47
		'
		'Label8
		'
		Me.Label8.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
		Me.Label8.AutoSize = True
		Me.Label8.Location = New System.Drawing.Point(638, 629)
		Me.Label8.Name = "Label8"
		Me.Label8.Size = New System.Drawing.Size(94, 20)
		Me.Label8.TabIndex = 13
		Me.Label8.Text = "Search Text"
		'
		'Label9
		'
		Me.Label9.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
		Me.Label9.AutoSize = True
		Me.Label9.Location = New System.Drawing.Point(539, 630)
		Me.Label9.Name = "Label9"
		Me.Label9.Size = New System.Drawing.Size(82, 20)
		Me.Label9.TabIndex = 48
		Me.Label9.Text = "Max Rows"
		'
		'chkChangedRecords
		'
		Me.chkChangedRecords.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
		Me.chkChangedRecords.AutoSize = True
		Me.chkChangedRecords.Checked = True
		Me.chkChangedRecords.CheckState = System.Windows.Forms.CheckState.Checked
		Me.chkChangedRecords.Location = New System.Drawing.Point(808, 651)
		Me.chkChangedRecords.Name = "chkChangedRecords"
		Me.chkChangedRecords.Size = New System.Drawing.Size(164, 24)
		Me.chkChangedRecords.TabIndex = 49
		Me.chkChangedRecords.Text = "Changed Records"
		Me.chkChangedRecords.UseVisualStyleBackColor = True
		'
		'chkNotCopied
		'
		Me.chkNotCopied.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
		Me.chkNotCopied.AutoSize = True
		Me.chkNotCopied.Checked = True
		Me.chkNotCopied.CheckState = System.Windows.Forms.CheckState.Checked
		Me.chkNotCopied.Location = New System.Drawing.Point(663, 602)
		Me.chkNotCopied.Name = "chkNotCopied"
		Me.chkNotCopied.Size = New System.Drawing.Size(136, 24)
		Me.chkNotCopied.TabIndex = 50
		Me.chkNotCopied.Text = "Not yet copied"
		Me.chkNotCopied.UseVisualStyleBackColor = True
		'
		'Form1
		'
		Me.AutoScaleBaseSize = New System.Drawing.Size(8, 19)
		Me.ClientSize = New System.Drawing.Size(1171, 937)
		Me.Controls.Add(Me.chkNotCopied)
		Me.Controls.Add(Me.chkChangedRecords)
		Me.Controls.Add(Me.Label9)
		Me.Controls.Add(Me.Label8)
		Me.Controls.Add(Me.txtSearchMax)
		Me.Controls.Add(Me.chkDropTable)
		Me.Controls.Add(Me.chkDeleteData)
		Me.Controls.Add(Me.chkCreateTable)
		Me.Controls.Add(Me.btnStop)
		Me.Controls.Add(Me.ProgressBar2)
		Me.Controls.Add(Me.dgTables)
		Me.Controls.Add(Me.txtSearch)
		Me.Controls.Add(Me.btnSearch)
		Me.Controls.Add(Me.lbCount)
		Me.Controls.Add(Me.txtLog)
		Me.Controls.Add(Me.GroupBox1)
		Me.Controls.Add(Me.Label1)
		Me.Controls.Add(Me.ProgressBar1)
		Me.Controls.Add(Me.btnCancel)
		Me.Controls.Add(Me.btnConnect)
		Me.Controls.Add(Me.txtConnect)
		Me.Controls.Add(Me.btnExport)
		Me.Name = "Form1"
		Me.Text = "Copy Table Data"
		Me.GroupBox1.ResumeLayout(False)
		Me.GroupBox1.PerformLayout()
		CType(Me.dgTables, System.ComponentModel.ISupportInitialize).EndInit()
		Me.ResumeLayout(False)
		Me.PerformLayout()

	End Sub

#End Region

	Dim oAppSetting As New AppSetting()
	Dim cnOdbc As OdbcConnection = Nothing
	Dim cn As OleDbConnection = Nothing
	Dim bStop As Boolean = False
	Dim sw As IO.StreamWriter

	Private Sub OpenConnections()

		If cnOdbc Is Nothing Then
			cnOdbc = New OdbcConnection
			cnOdbc.ConnectionString = GetMySqlConnectionString()
			cnOdbc.Open()
		End If

		If cn Is Nothing Then
			cn = New OleDbConnection
			cn.ConnectionString = txtConnect.Text
			cn.Open()
		End If

		If cn.State <> ConnectionState.Open Then
			cn.Open()
		End If

		If cnOdbc.State <> ConnectionState.Open Then
			cnOdbc.Open()
		End If

	End Sub

	Private Sub frmExport_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
		oAppSetting.LoadData()

		Dim mysqlDrivers As List(Of String) = GetMySqlOdbcDrivers()
		For Each driver As String In mysqlDrivers
			cboDriver.Items.Add(driver)
		Next

		Dim sDriver As String = oAppSetting.GetValue("Driver")
		If sDriver = "" Then
			cboDriver.SelectedIndex = 0
		Else
			cboDriver.SelectedItem = sDriver
		End If

		txtServer.Text = oAppSetting.GetValue("Server")
		txtDatabase.Text = oAppSetting.GetValue("Database")
		txtPort.Text = oAppSetting.GetValue("Port")
		txtUser.Text = oAppSetting.GetValue("User")
		txtPassword.Text = oAppSetting.GetValue("Password")
		txtConnect.Text = oAppSetting.GetValue("ConnectionString")

	End Sub

	Function GetMySqlOdbcDrivers() As List(Of String)
		Dim drivers As New List(Of String)
		Dim driverKeyPaths As String() = {
			"SOFTWARE\ODBC\ODBCINST.INI\ODBC Drivers",  ' Path for 64-bit applications on 64-bit OS
			"SOFTWARE\Wow6432Node\ODBC\ODBCINST.INI\ODBC Drivers"  ' Path for 32-bit applications on 64-bit OS
		}

		For Each keyPath As String In driverKeyPaths
			Using key As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(keyPath)
				If key IsNot Nothing Then
					For Each driverName As String In key.GetValueNames()
						If driverName.Contains("MySQL") Then
							drivers.Add(driverName)
						End If
					Next
				End If
			End Using
		Next

		Return drivers
	End Function

	Private Sub txtTest_Click(sender As System.Object, e As System.EventArgs) Handles txtTest.Click

		oAppSetting.SetValue("Driver", cboDriver.Text)
		oAppSetting.SetValue("Server", txtServer.Text)
		oAppSetting.SetValue("Database", txtDatabase.Text)
		oAppSetting.SetValue("Port", txtPort.Text)
		oAppSetting.SetValue("User", txtUser.Text)
		oAppSetting.SetValue("Password", txtPassword.Text)
		oAppSetting.SaveData()

		Try
			SetupGrid()
		Catch ex As Exception
			MsgBox(ex.Message)
			Exit Sub
		End Try

	End Sub

	Sub SetupGrid()

		OpenConnections()

		Dim sSqlSrc As String = "SELECT table_name, table_rows FROM information_schema.tables WHERE table_type = 'BASE TABLE' and table_rows > 0"
		Dim dtSrc As DataTable = GetTable(sSqlSrc, cnOdbc)

		Dim sSqlDest As String = "SELECT t.name AS table_name, SUM(p.rows) AS row_count
								FROM sys.tables t
								INNER JOIN sys.indexes i ON t.object_id = i.object_id
								INNER JOIN sys.partitions p ON i.object_id = p.object_id AND i.index_id = p.index_id
								WHERE t.type = 'U' AND i.type <= 1  
								GROUP BY t.name
								ORDER BY t.name"
		Dim dtDest As DataTable = GetTable(sSqlDest, cn)

		Dim oTable As New Data.DataTable
		oTable.Columns.Add(New Data.DataColumn("Checked", System.Type.GetType("System.Boolean"))) '<--
		oTable.Columns.Add(New Data.DataColumn("Name"))
		oTable.Columns.Add(New Data.DataColumn("RowCount", System.Type.GetType("System.Int64")))
		oTable.Columns.Add(New Data.DataColumn("DestRowCount", System.Type.GetType("System.Int64"))) '<--

		For iRow As Integer = 0 To dtSrc.Rows.Count - 1
			Dim oDataRow As DataRow = oTable.NewRow()
			Dim sTable As String = dtSrc.Rows(iRow)("table_name").ToString() & ""
			oDataRow("Name") = sTable
			oDataRow("RowCount") = dtSrc.Rows(iRow)("table_rows")

			Dim oDestRows As DataRow() = dtDest.Select("table_name = '" & PadQuotes(sTable) & "'")
			If oDestRows.Length > 0 Then
				oDataRow("DestRowCount") = oDestRows(0)("row_count")
			End If

			oTable.Rows.Add(oDataRow)
		Next

		SetupGrid2(oTable)
	End Sub

	Public Function GetTable(ByVal sSql As String, ByRef cn As OdbcConnection) As System.Data.DataTable
		Dim ad As New OdbcDataAdapter(sSql, cn)
		Dim ds As DataSet = New DataSet

		Try
			ad.Fill(ds)
		Catch ex As Exception
			Throw New Exception(ex.Message & ", SQL: " & sSql)
		Finally
		End Try

		Return ds.Tables(0)
	End Function

	Public Function GetTable(ByVal sSql As String, ByRef cn As OleDbConnection) As System.Data.DataTable
		Dim ad As New OleDbDataAdapter(sSql, cn)
		Dim ds As DataSet = New DataSet

		Try
			ad.Fill(ds)
		Catch ex As Exception
			Throw New Exception(ex.Message & ", SQL: " & sSql)
		Finally
		End Try

		Return ds.Tables(0)
	End Function

	Sub SetupGrid2(oTable As DataTable)
		dgTables.Rows.Clear()

		dgTables.DataSource = oTable
		dgTables.Update()

		Dim oCol As DataGridViewCheckBoxColumn = DirectCast(dgTables.Columns("Checked"), DataGridViewCheckBoxColumn)
		oCol.TrueValue = True
		oCol.SortMode = DataGridViewColumnSortMode.Automatic
		oCol.Width = 35
		oCol.HeaderText = ""

		'dgTables.Columns("DestRowCount").Visible = True
		UpdateDataColumn("RowCount", "#,#", "Src Row Count")
		UpdateDataColumn("DestRowCount", "#,#", "Dest Row Count")
		SetupBackground()
	End Sub

	Private Sub UpdateDataColumn(sColName As String, sFormat As String, sHeaderText As String)
		Dim oCol As DataGridViewColumn = dgTables.Columns(sColName)
		If sFormat <> "" Then oCol.DefaultCellStyle.Format = sFormat
		If sHeaderText <> "" Then oCol.HeaderText = sHeaderText
	End Sub

	Private Sub SetupBackground()
		For iRow = 0 To dgTables.RowCount - 1
			Dim sSrcCount As String = dgTables.Rows(iRow).Cells("RowCount").Value.ToString()
			Dim sDstCount As String = dgTables.Rows(iRow).Cells("DestRowCount").Value.ToString()
			If sSrcCount <> "" AndAlso sDstCount <> "" Then
				If CInt(sSrcCount) = CInt(sDstCount) Then
					dgTables.Rows(iRow).Cells("DestRowCount").Style.BackColor = Color.LightBlue
				Else
					dgTables.Rows(iRow).Cells("DestRowCount").Style.BackColor = Color.LightPink
				End If
			Else
				dgTables.Rows(iRow).Cells("DestRowCount").Style.BackColor = Color.White
			End If
		Next
	End Sub

	Function GetMySqlConnectionString() As String
		Return "Driver={" & cboDriver.Text & "};Option=3" &
		   ";Server=" & txtServer.Text &
		   ";Port=" & txtPort.Text &
		   ";Database=" & txtDatabase.Text &
		   ";User=" & txtUser.Text &
		   ";Password=" & txtPassword.Text & ";"
	End Function

	Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
		Me.Close()
	End Sub

	Private Sub btnConnect_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnConnect.Click
		Dim sConnectionString As String = "Provider=MSOLEDBSQL.1" 'MSOLEDBSQL.1 - Driver, 'SQLOLEDB.1 - Provider
		sConnectionString = EditConnectionString(sConnectionString)
		If sConnectionString = "" Then
			Exit Sub
		End If

		txtConnect.Text = sConnectionString
		oAppSetting.SetValue("ConnectionString", sConnectionString)
		oAppSetting.SaveData()
	End Sub

	Protected Function EditConnectionString(ByVal sConnectionString As String) As String
		Try
			Dim oDataLinks As Object = CreateObject("DataLinks")
			Dim cn As Object = CreateObject("ADODB.Connection")

			cn.ConnectionString = sConnectionString
			oDataLinks.hWnd = Me.Handle

			If Not oDataLinks.PromptEdit(cn) Then
				'User pressed cancel button
				Return ""
			End If

			cn.Open()

			Return cn.ConnectionString

		Catch ex As Exception
			MsgBox(ex.Message)
			Return ""
		End Try
	End Function

	Function GetSelectedTables() As List(Of String)
		Dim oRet As New List(Of String)

		For Each oRow As DataGridViewRow In dgTables.Rows
			Dim oCheckbox As DataGridViewCheckBoxCell = DirectCast(oRow.Cells.Item(0), DataGridViewCheckBoxCell)

			If oCheckbox.Value.ToString = oCheckbox.TrueValue.ToString() Then
				Dim sName As String = oRow.Cells(1).Value.ToString()
				oRet.Add(sName)
			End If
		Next

		Return oRet
	End Function

	Private Sub btnExport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExport.Click

		If cnOdbc Is Nothing Or cn Is Nothing Then
			MsgBox("Please connect to the source database.")
			Exit Sub
		End If

		Dim oTables As List(Of String) = GetSelectedTables()
		If oTables.Count = 0 Then
			MsgBox("Please select tables to copy.")
			Exit Sub
		End If

		txtLog.Clear()
		ProgressBar2.Maximum = oTables.Count
		bStop = False
		Dim i As Integer = 0

		For Each oRow As DataGridViewRow In dgTables.Rows
			Dim oCheckbox As DataGridViewCheckBoxCell = DirectCast(oRow.Cells.Item(0), DataGridViewCheckBoxCell)
			If oCheckbox.Value.ToString = oCheckbox.TrueValue.ToString() Then

				i += 1
				ProgressBar2.Value = i
				ProgressBar2.Refresh()
				Windows.Forms.Application.DoEvents()

				OpenConnections()

				Dim sTable As String = oRow.Cells("Name").Value.ToString()
				Dim iSrcCount As Integer = CInt(oRow.Cells("RowCount").Value)
				Dim sDestRecCount As String = oRow.Cells("DestRowCount").Value.ToString()
				ExportTable(sTable, iSrcCount, sDestRecCount)

				If bStop Then
					Exit For
				End If
			End If
		Next

		Try
			cnOdbc.Close()
			cn.Close()
		Catch ex As Exception
			'Ignore
		End Try

		ProgressBar1.Value = 0
		ProgressBar2.Value = 0

		MsgBox("Done")

	End Sub


	Private Sub ExportTable(ByVal sTableName As String, iSrcCount As Integer, sDestRecCount As String)

		If chkCreateTable.Checked Then

			If chkDropTable.Checked AndAlso sDestRecCount <> "" Then
				Log("Drop table: " & sTableName)

				Dim sSqlDrop As String = "DROP TABLE " & PadQuotes(sTableName)

				Try
					ExecuteNonQuery(sSqlDrop, cn)
					sDestRecCount = ""
				Catch ex As Exception
					Log("Could not drop table: " & sTableName & ", " & ex.Message & vbTab)
				End Try
			End If

			If sDestRecCount = "" Then
				Log("Create table: " & sTableName)

				Dim sCreateTableSql As String = GetCreateTableSql(sTableName, cnOdbc)

				Try
					ExecuteNonQuery(sCreateTableSql, cn)
					sDestRecCount = "0"
				Catch ex As Exception
					LogErrorToFile(sTableName, ex.Message, sCreateTableSql, True)
					Log(ex.Message & vbTab & "SQL: " & sCreateTableSql)
					Exit Sub
				End Try
			End If

		End If

		If chkDeleteData.Checked AndAlso sDestRecCount <> "" Then
			Log("Deleteting data from table: " & sTableName & ", Rows: " & sDestRecCount)

			Dim sSql1 As String = "DELETE FROM " & PadColumnName(sTableName)
			Try
				ExecuteNonQuery(sSql1, cn)
			Catch ex As Exception
				Log(ex.Message & vbTab & "SQL: " & sSql1)
			End Try
		End If

		If sDestRecCount = "" Then
			Log("Destination table does not exist: " & sTableName)
			Exit Sub
		End If

		If iSrcCount = 0 Then
			'Nothing to copy - Exit
			Exit Sub
		End If

		'Copy Data
		ProgressBar1.Maximum = iSrcCount
		lbCount.Visible = True

		Log("Copying " & iSrcCount & " rows from table: " & sTableName)

		Dim cmd As OdbcCommand = New OdbcCommand("SELECT * FROM `" & sTableName & "`", cnOdbc)
		Dim dr As OdbcDataReader = cmd.ExecuteReader()
		Dim oSchemaRows As Data.DataRowCollection = dr.GetSchemaTable.Rows

		Dim sRow As String
		Dim i As Integer
		Dim iRow As Integer = 0

		'Get Header
		Dim sHeader As String = ""
		For i = 0 To oSchemaRows.Count - 1
			Dim sColumn As String = oSchemaRows(i)("ColumnName")
			If i <> 0 Then
				sHeader += ", "
			End If
			sHeader += PadColumnName(sColumn)
		Next

		While dr.Read()
			sRow = ""
			For i = 0 To oSchemaRows.Count - 1
				If sRow <> "" Then
					sRow += ", "
				End If

				sRow += GetValueString(dr.GetValue(i))
			Next

			Dim sSql1 As String = "INSERT INTO " & PadColumnName(sTableName) & " (" & sHeader & ") VALUES (" & sRow & ")"

			Try
				ExecuteNonQuery(sSql1, cn)
			Catch ex As Exception
				LogErrorToFile(sTableName, ex.Message, sSql1, False)
				Log("Error inserting data into table: " & sTableName & ", Error:" & ex.Message)
			End Try

			iRow += 1

			If iRow <= iSrcCount Then
				ProgressBar1.Value = iRow
			End If

			lbCount.Text = iRow.ToString()
			lbCount.Refresh()
		End While
		dr.Close()

		ProgressBar1.Value = 0
		lbCount.Visible = False
		lbCount.Text = ""

		Log("Finished processing " & sTableName)

		If Not sw Is Nothing Then
			sw.Close()
			sw = Nothing
		End If

	End Sub

	Sub LogErrorToFile(ByVal sTableName As String, sError As String, sSql As String, bCloseFile As Boolean)
		If sw Is Nothing Then
			Dim sAssPath As String = System.Reflection.Assembly.GetExecutingAssembly().Location
			Dim sFolderPath As String = System.IO.Path.GetDirectoryName(sAssPath)
			Dim sFilePath As String = IO.Path.Combine(sFolderPath, PadFileName(sTableName) & ".sql")

			If IO.File.Exists(sFilePath) Then
				IO.File.Delete(sFilePath)
			End If

			sw = New IO.StreamWriter(sFilePath, False)
		End If

		sw.WriteLine("--" & sError)
		sw.WriteLine(sSql)

		If bCloseFile Then
			sw.Close()
			sw = Nothing
		End If

	End Sub

	Public Function PadFileName(ByVal s As String) As String
		s = Replace(s, "<", "")
		s = Replace(s, ">", "")
		s = Replace(s, ":", "-")
		s = Replace(s, """", "")
		s = Replace(s, "/", "")
		s = Replace(s, "\", "")
		s = Replace(s, "?", "")
		s = Replace(s, "'", "")
		s = Replace(s, ChrW(65533), "")
		's = Replace(s, " ", "_")
		Return Replace(s, "*", "")
	End Function

	Private Sub ExecuteNonQuery(sSql As String, cn As OleDbConnection)
		Dim cm As New OleDbCommand(sSql, cn)
		cm.ExecuteNonQuery()
	End Sub

	Private Sub Log(s As String)
		txtLog.AppendText(s & vbCrLf)
		txtLog.SelectionStart = txtLog.Text.Length
		txtLog.ScrollToCaret()
		txtLog.Refresh()
	End Sub

	Private Function GetValueString(ByVal obj As Object) As String
		If (IsDBNull(obj)) Then Return "NULL"

		Select Case obj.GetType.FullName

			Case "System.Boolean"
				If (obj = True) Then
					Return "1"
				Else
					Return "0"
				End If

			Case "System.String"
				Dim str As String = obj
				Return "'" + str.Replace("'", "''") + "'"

			Case "System.DateTime"
				Try
					Return "'" + ValidateSqlServerDateTime(CDate(obj)).ToString() + "'"
				Catch ex As Exception
					Return "'" + obj.ToString() + "'"
				End Try

			Case "System.Drawing.Image"
				Return "NULL"

			Case "System.Drawing.Bitmap"
				Return "NULL"

			Case "System.Byte[]"
				Return "0x" + GetHexString(obj)

			Case Else
				Return obj.ToString()

		End Select
	End Function
	Function ValidateSqlServerDateTime(inputDate As DateTime) As DateTime
		Dim minDate As New DateTime(1753, 1, 1)
		Dim maxDate As New DateTime(9999, 12, 31)

		If inputDate < minDate Then
			Return minDate
		ElseIf inputDate > maxDate Then
			Return maxDate
		Else
			Return inputDate
		End If
	End Function

	Private Function GetHexString(ByRef bytes() As Byte) As String
		Dim sb As New System.Text.StringBuilder
		Dim b As Byte
		Dim i As Integer = 0

		For Each b In bytes
			i += 1
			sb.Append(b.ToString("X2"))
			If i > 10 Then
				Return sb.ToString()
			End If
		Next

		Return sb.ToString()
	End Function

	Private Function GetCreateTableSql(ByVal sTableName As String, ByRef cn As OdbcConnection) As String
		Dim sb As New System.Text.StringBuilder()

		sb.Append("CREATE TABLE " & PadColumnName(sTableName) & " (" & vbCrLf)

		Dim sSql As String = "select * from INFORMATION_SCHEMA.COLUMNS where TABLE_NAME = '" & PadQuotes(sTableName) & "'"

		Dim cmd As New OdbcCommand(sSql, cn)
		Dim dr As OdbcDataReader = cmd.ExecuteReader()

		Dim i As Integer = 0
		While dr.Read
			Dim sColumn As String = dr.GetValue(dr.GetOrdinal("COLUMN_NAME")).ToString()
			Dim sDataType As String = dr.GetValue(dr.GetOrdinal("DATA_TYPE")).ToString()
			Dim bAllowDBNull As Boolean = dr.GetString(dr.GetOrdinal("IS_NULLABLE")) = "YES"
			Dim sColumnSize As String = dr.GetValue(dr.GetOrdinal("CHARACTER_MAXIMUM_LENGTH")).ToString()

			'MySql to SQL Server data type converter
			Select Case LCase(sDataType)
				Case "longtext" : sDataType = "text"
				Case "ntext" : sDataType = "text"
				Case "double" : sDataType = "float"
			End Select

			If sDataType = "decimal" OrElse sDataType = "numeric" Then
				Dim sPrecision As String = dr.GetValue(dr.GetOrdinal("NUMERIC_PRECISION")).ToString() & ""
				Dim sScale As String = dr.GetValue(dr.GetOrdinal("NUMERIC_SCALE")).ToString() & ""
				sDataType += "(" & sPrecision & ", " & sScale & ")"
			End If

			If i > 0 Then
				sb.Append(",")
				sb.Append(vbCrLf)
			End If

			sb.Append(PadColumnName(sColumn))
			sb.Append(" " & sDataType)

			If (LCase(sDataType) = "varchar" Or LCase(sDataType) = "nvarchar") And sColumnSize <> "" Then
				sb.Append("(" & sColumnSize & ")")
			End If

			If bAllowDBNull Then
				sb.Append(" NULL")
			Else
				sb.Append(" NOT NULL")
			End If

			i += 1
		End While

		sb.Append(")")

		dr.Close()

		If i = 0 Then
			Return ""
		Else
			Return sb.ToString()
		End If

	End Function

	Public Function PadQuotes(ByVal s As String) As String
		If s = "" Then
			Return ""
		End If
		Return (s & "").Replace("'", "''")
	End Function

	Public Function PadColumnName(ByVal sTable As String) As String
		Return "[" & sTable & "]"
	End Function

	Private Sub btnSearch_Click(sender As Object, e As EventArgs) Handles btnSearch.Click

		Dim sMax As String = txtSearchMax.Text

		For iRow = 0 To dgTables.RowCount - 1
			Dim oRow As DataGridViewRow = dgTables.Rows(iRow)
			Dim sName As String = oRow.Cells("Name").Value.ToString()
			Dim sSrcCount As String = oRow.Cells("RowCount").Value.ToString() & ""
			Dim sDestRecCount As String = oRow.Cells("DestRowCount").Value.ToString() & ""
			Dim bFound As Boolean = True

			If chkNotCopied.Checked Then
				bFound = sDestRecCount = ""
			ElseIf chkChangedRecords.Checked Then
				bFound = sSrcCount <> sDestRecCount
			End If

			If bFound AndAlso sMax <> "" Then
				If CInt(sSrcCount) > CInt(sMax) Then
					bFound = False
				End If
			End If

			If (txtSearch.Text = "" OrElse sName.IndexOf(txtSearch.Text) <> -1) AndAlso bFound Then
				dgTables.Rows(iRow).Cells("Checked").Value = True
			Else
				dgTables.Rows(iRow).Cells("Checked").Value = False
			End If
		Next
	End Sub

	Private Sub btnStop_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles btnStop.LinkClicked
		bStop = True
	End Sub

End Class
