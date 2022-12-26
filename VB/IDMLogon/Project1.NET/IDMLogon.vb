Option Strict Off
Option Explicit On
Friend Class frmLogon
	Inherits System.Windows.Forms.Form
#Region "Windows Form Designer generated code "
	Public Sub New()
		MyBase.New()
		If m_vb6FormDefInstance Is Nothing Then
			If m_InitializingDefInstance Then
				m_vb6FormDefInstance = Me
			Else
				Try 
					'For the start-up form, the first instance created is the default instance.
					If System.Reflection.Assembly.GetExecutingAssembly.EntryPoint.DeclaringType Is Me.GetType Then
						m_vb6FormDefInstance = Me
					End If
				Catch
				End Try
			End If
		End If
		'This call is required by the Windows Form Designer.
		InitializeComponent()
	End Sub
	'Form overrides dispose to clean up the component list.
	Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
		If Disposing Then
			Static fTerminateCalled As Boolean
			If Not fTerminateCalled Then
				Form_Terminate_renamed()
				fTerminateCalled = True
			End If
			If Not components Is Nothing Then
				components.Dispose()
			End If
		End If
		MyBase.Dispose(Disposing)
	End Sub
	'Required by the Windows Form Designer
	Private components As System.ComponentModel.IContainer
	Public ToolTip1 As System.Windows.Forms.ToolTip
	Public WithEvents cmbLibraries As System.Windows.Forms.ComboBox
	Public WithEvents btnUpdatePwd As System.Windows.Forms.Button
	Public WithEvents txtNewPwd2 As System.Windows.Forms.TextBox
	Public WithEvents txtNewPwd As System.Windows.Forms.TextBox
	Public WithEvents txtPassword As System.Windows.Forms.TextBox
	Public WithEvents txtUsername As System.Windows.Forms.TextBox
	Public WithEvents btnShowChange As System.Windows.Forms.Button
	Public WithEvents btnLogon As System.Windows.Forms.Button
	Public WithEvents Label4 As System.Windows.Forms.Label
	Public WithEvents Label1 As System.Windows.Forms.Label
	Public WithEvents Label3 As System.Windows.Forms.Label
	Public WithEvents Label2 As System.Windows.Forms.Label
	Public WithEvents lblCatalog As System.Windows.Forms.Label
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmLogon))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.ToolTip1.Active = True
		Me.cmbLibraries = New System.Windows.Forms.ComboBox
		Me.btnUpdatePwd = New System.Windows.Forms.Button
		Me.txtNewPwd2 = New System.Windows.Forms.TextBox
		Me.txtNewPwd = New System.Windows.Forms.TextBox
		Me.txtPassword = New System.Windows.Forms.TextBox
		Me.txtUsername = New System.Windows.Forms.TextBox
		Me.btnShowChange = New System.Windows.Forms.Button
		Me.btnLogon = New System.Windows.Forms.Button
		Me.Label4 = New System.Windows.Forms.Label
		Me.Label1 = New System.Windows.Forms.Label
		Me.Label3 = New System.Windows.Forms.Label
		Me.Label2 = New System.Windows.Forms.Label
		Me.lblCatalog = New System.Windows.Forms.Label
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
		Me.Text = "IDM Logon"
		Me.ClientSize = New System.Drawing.Size(389, 176)
		Me.Location = New System.Drawing.Point(3, 22)
		Me.Icon = CType(resources.GetObject("frmLogon.Icon"), System.Drawing.Icon)
		Me.MaximizeBox = False
		Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
		Me.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
		Me.BackColor = System.Drawing.SystemColors.Control
		Me.ControlBox = True
		Me.Enabled = True
		Me.KeyPreview = False
		Me.MinimizeBox = True
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.ShowInTaskbar = True
		Me.HelpButton = False
		Me.WindowState = System.Windows.Forms.FormWindowState.Normal
		Me.Name = "frmLogon"
		Me.cmbLibraries.Size = New System.Drawing.Size(153, 21)
		Me.cmbLibraries.Location = New System.Drawing.Point(120, 16)
		Me.cmbLibraries.TabIndex = 12
		Me.cmbLibraries.Text = "Combo1"
		Me.cmbLibraries.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmbLibraries.BackColor = System.Drawing.SystemColors.Window
		Me.cmbLibraries.CausesValidation = True
		Me.cmbLibraries.Enabled = True
		Me.cmbLibraries.ForeColor = System.Drawing.SystemColors.WindowText
		Me.cmbLibraries.IntegralHeight = True
		Me.cmbLibraries.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmbLibraries.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmbLibraries.Sorted = False
		Me.cmbLibraries.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown
		Me.cmbLibraries.TabStop = True
		Me.cmbLibraries.Visible = True
		Me.cmbLibraries.Name = "cmbLibraries"
		Me.btnUpdatePwd.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.btnUpdatePwd.Text = "&Update"
		Me.btnUpdatePwd.Size = New System.Drawing.Size(89, 25)
		Me.btnUpdatePwd.Location = New System.Drawing.Point(288, 136)
		Me.btnUpdatePwd.TabIndex = 11
		Me.btnUpdatePwd.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.btnUpdatePwd.BackColor = System.Drawing.SystemColors.Control
		Me.btnUpdatePwd.CausesValidation = True
		Me.btnUpdatePwd.Enabled = True
		Me.btnUpdatePwd.ForeColor = System.Drawing.SystemColors.ControlText
		Me.btnUpdatePwd.Cursor = System.Windows.Forms.Cursors.Default
		Me.btnUpdatePwd.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.btnUpdatePwd.TabStop = True
		Me.btnUpdatePwd.Name = "btnUpdatePwd"
		Me.txtNewPwd2.AutoSize = False
		Me.txtNewPwd2.Size = New System.Drawing.Size(153, 19)
		Me.txtNewPwd2.IMEMode = System.Windows.Forms.ImeMode.Disable
		Me.txtNewPwd2.Location = New System.Drawing.Point(120, 144)
		Me.txtNewPwd2.PasswordChar = ChrW(42)
		Me.txtNewPwd2.TabIndex = 9
		Me.txtNewPwd2.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtNewPwd2.AcceptsReturn = True
		Me.txtNewPwd2.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtNewPwd2.BackColor = System.Drawing.SystemColors.Window
		Me.txtNewPwd2.CausesValidation = True
		Me.txtNewPwd2.Enabled = True
		Me.txtNewPwd2.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtNewPwd2.HideSelection = True
		Me.txtNewPwd2.ReadOnly = False
		Me.txtNewPwd2.Maxlength = 0
		Me.txtNewPwd2.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtNewPwd2.MultiLine = False
		Me.txtNewPwd2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtNewPwd2.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtNewPwd2.TabStop = True
		Me.txtNewPwd2.Visible = True
		Me.txtNewPwd2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtNewPwd2.Name = "txtNewPwd2"
		Me.txtNewPwd.AutoSize = False
		Me.txtNewPwd.Size = New System.Drawing.Size(153, 19)
		Me.txtNewPwd.IMEMode = System.Windows.Forms.ImeMode.Disable
		Me.txtNewPwd.Location = New System.Drawing.Point(120, 112)
		Me.txtNewPwd.PasswordChar = ChrW(42)
		Me.txtNewPwd.TabIndex = 7
		Me.txtNewPwd.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtNewPwd.AcceptsReturn = True
		Me.txtNewPwd.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtNewPwd.BackColor = System.Drawing.SystemColors.Window
		Me.txtNewPwd.CausesValidation = True
		Me.txtNewPwd.Enabled = True
		Me.txtNewPwd.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtNewPwd.HideSelection = True
		Me.txtNewPwd.ReadOnly = False
		Me.txtNewPwd.Maxlength = 0
		Me.txtNewPwd.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtNewPwd.MultiLine = False
		Me.txtNewPwd.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtNewPwd.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtNewPwd.TabStop = True
		Me.txtNewPwd.Visible = True
		Me.txtNewPwd.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtNewPwd.Name = "txtNewPwd"
		Me.txtPassword.AutoSize = False
		Me.txtPassword.Size = New System.Drawing.Size(153, 19)
		Me.txtPassword.IMEMode = System.Windows.Forms.ImeMode.Disable
		Me.txtPassword.Location = New System.Drawing.Point(120, 80)
		Me.txtPassword.PasswordChar = ChrW(42)
		Me.txtPassword.TabIndex = 1
		Me.txtPassword.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtPassword.AcceptsReturn = True
		Me.txtPassword.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtPassword.BackColor = System.Drawing.SystemColors.Window
		Me.txtPassword.CausesValidation = True
		Me.txtPassword.Enabled = True
		Me.txtPassword.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtPassword.HideSelection = True
		Me.txtPassword.ReadOnly = False
		Me.txtPassword.Maxlength = 0
		Me.txtPassword.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtPassword.MultiLine = False
		Me.txtPassword.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtPassword.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtPassword.TabStop = True
		Me.txtPassword.Visible = True
		Me.txtPassword.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtPassword.Name = "txtPassword"
		Me.txtUsername.AutoSize = False
		Me.txtUsername.Size = New System.Drawing.Size(153, 19)
		Me.txtUsername.IMEMode = System.Windows.Forms.ImeMode.Disable
		Me.txtUsername.Location = New System.Drawing.Point(120, 48)
		Me.txtUsername.TabIndex = 0
		Me.txtUsername.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtUsername.AcceptsReturn = True
		Me.txtUsername.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtUsername.BackColor = System.Drawing.SystemColors.Window
		Me.txtUsername.CausesValidation = True
		Me.txtUsername.Enabled = True
		Me.txtUsername.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtUsername.HideSelection = True
		Me.txtUsername.ReadOnly = False
		Me.txtUsername.Maxlength = 0
		Me.txtUsername.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtUsername.MultiLine = False
		Me.txtUsername.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtUsername.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtUsername.TabStop = True
		Me.txtUsername.Visible = True
		Me.txtUsername.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtUsername.Name = "txtUsername"
		Me.btnShowChange.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.btnShowChange.Text = "&Change Pwd >>"
		Me.btnShowChange.Enabled = False
		Me.btnShowChange.Size = New System.Drawing.Size(89, 25)
		Me.btnShowChange.Location = New System.Drawing.Point(288, 48)
		Me.btnShowChange.TabIndex = 3
		Me.btnShowChange.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.btnShowChange.BackColor = System.Drawing.SystemColors.Control
		Me.btnShowChange.CausesValidation = True
		Me.btnShowChange.ForeColor = System.Drawing.SystemColors.ControlText
		Me.btnShowChange.Cursor = System.Windows.Forms.Cursors.Default
		Me.btnShowChange.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.btnShowChange.TabStop = True
		Me.btnShowChange.Name = "btnShowChange"
		Me.btnLogon.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.btnLogon.Text = "&Logon"
		Me.AcceptButton = Me.btnLogon
		Me.btnLogon.Size = New System.Drawing.Size(89, 25)
		Me.btnLogon.Location = New System.Drawing.Point(288, 16)
		Me.btnLogon.TabIndex = 2
		Me.btnLogon.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.btnLogon.BackColor = System.Drawing.SystemColors.Control
		Me.btnLogon.CausesValidation = True
		Me.btnLogon.Enabled = True
		Me.btnLogon.ForeColor = System.Drawing.SystemColors.ControlText
		Me.btnLogon.Cursor = System.Windows.Forms.Cursors.Default
		Me.btnLogon.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.btnLogon.TabStop = True
		Me.btnLogon.Name = "btnLogon"
		Me.Label4.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me.Label4.Text = "Confirm:"
		Me.Label4.Size = New System.Drawing.Size(65, 17)
		Me.Label4.Location = New System.Drawing.Point(48, 144)
		Me.Label4.TabIndex = 10
		Me.Label4.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label4.BackColor = System.Drawing.SystemColors.Control
		Me.Label4.Enabled = True
		Me.Label4.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label4.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label4.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label4.UseMnemonic = True
		Me.Label4.Visible = True
		Me.Label4.AutoSize = False
		Me.Label4.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label4.Name = "Label4"
		Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me.Label1.Text = "New Password:"
		Me.Label1.Size = New System.Drawing.Size(97, 17)
		Me.Label1.Location = New System.Drawing.Point(16, 112)
		Me.Label1.TabIndex = 8
		Me.Label1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label1.BackColor = System.Drawing.SystemColors.Control
		Me.Label1.Enabled = True
		Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label1.UseMnemonic = True
		Me.Label1.Visible = True
		Me.Label1.AutoSize = False
		Me.Label1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label1.Name = "Label1"
		Me.Label3.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me.Label3.Text = "Password:"
		Me.Label3.Size = New System.Drawing.Size(57, 17)
		Me.Label3.Location = New System.Drawing.Point(56, 80)
		Me.Label3.TabIndex = 6
		Me.Label3.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label3.BackColor = System.Drawing.SystemColors.Control
		Me.Label3.Enabled = True
		Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label3.UseMnemonic = True
		Me.Label3.Visible = True
		Me.Label3.AutoSize = False
		Me.Label3.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label3.Name = "Label3"
		Me.Label2.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me.Label2.Text = "Username:"
		Me.Label2.Size = New System.Drawing.Size(57, 17)
		Me.Label2.Location = New System.Drawing.Point(56, 48)
		Me.Label2.TabIndex = 5
		Me.Label2.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label2.BackColor = System.Drawing.SystemColors.Control
		Me.Label2.Enabled = True
		Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label2.UseMnemonic = True
		Me.Label2.Visible = True
		Me.Label2.AutoSize = False
		Me.Label2.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label2.Name = "Label2"
		Me.lblCatalog.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me.lblCatalog.Text = "Library:"
		Me.lblCatalog.Size = New System.Drawing.Size(57, 17)
		Me.lblCatalog.Location = New System.Drawing.Point(56, 16)
		Me.lblCatalog.TabIndex = 4
		Me.lblCatalog.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblCatalog.BackColor = System.Drawing.SystemColors.Control
		Me.lblCatalog.Enabled = True
		Me.lblCatalog.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblCatalog.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblCatalog.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblCatalog.UseMnemonic = True
		Me.lblCatalog.Visible = True
		Me.lblCatalog.AutoSize = False
		Me.lblCatalog.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblCatalog.Name = "lblCatalog"
		Me.Controls.Add(cmbLibraries)
		Me.Controls.Add(btnUpdatePwd)
		Me.Controls.Add(txtNewPwd2)
		Me.Controls.Add(txtNewPwd)
		Me.Controls.Add(txtPassword)
		Me.Controls.Add(txtUsername)
		Me.Controls.Add(btnShowChange)
		Me.Controls.Add(btnLogon)
		Me.Controls.Add(Label4)
		Me.Controls.Add(Label1)
		Me.Controls.Add(Label3)
		Me.Controls.Add(Label2)
		Me.Controls.Add(lblCatalog)
	End Sub
#End Region 
#Region "Upgrade Support "
	Private Shared m_vb6FormDefInstance As frmLogon
	Private Shared m_InitializingDefInstance As Boolean
	Public Shared Property DefInstance() As frmLogon
		Get
			If m_vb6FormDefInstance Is Nothing OrElse m_vb6FormDefInstance.IsDisposed Then
				m_InitializingDefInstance = True
				m_vb6FormDefInstance = New frmLogon()
				m_InitializingDefInstance = False
			End If
			DefInstance = m_vb6FormDefInstance
		End Get
		Set
			m_vb6FormDefInstance = Value
		End Set
	End Property
#End Region 
	' Sample logon program Copyright(C) 1997 FileNet Corporation
	'
	'Declare a user-defined variable to pass to the Shell_NotifyIcon
	'function.
	
	Private Structure NOTIFYICONDATA
		Dim cbSize As Integer
		Dim hWnd As Integer
		Dim uId As Integer
		Dim uFlags As Integer
		Dim uCallBackMessage As Integer
		Dim hIcon As Integer
		<VBFixedString(64),System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValTStr,SizeConst:=64)> Public szTip As String
	End Structure
	
	'Declare the constants for the API function. These constants can be
	'found in the header file Shellapi.h.
	
	'The following constants are the messages sent to the
	'Shell_NotifyIcon function to add, modify, or delete an icon from the
	'taskbar status area.
	Private Const NIM_ADD As Short = &H0s
	Private Const NIM_MODIFY As Short = &H1s
	Private Const NIM_DELETE As Short = &H2s
	
	'The following constant is the message sent when a mouse event occurs
	'within the rectangular boundaries of the icon in the taskbar status
	'area.
	Private Const WM_MOUSEMOVE As Short = &H200s
	
	'The following constants are the flags that indicate the valid
	'members of the NOTIFYICONDATA data type.
	Private Const NIF_MESSAGE As Short = &H1s
	Private Const NIF_ICON As Short = &H2s
	Private Const NIF_TIP As Short = &H4s
	
	'The following constants are used to determine the mouse input on the
	'the icon in the taskbar status area.
	
	'Left-click constants.
	Private Const WM_LBUTTONDBLCLK As Short = &H203s 'Double-click
	Private Const WM_LBUTTONDOWN As Short = &H201s 'Button down
	Private Const WM_LBUTTONUP As Short = &H202s 'Button up
	
	'Right-click constants.
	Private Const WM_RBUTTONDBLCLK As Short = &H206s 'Double-click
	Private Const WM_RBUTTONDOWN As Short = &H204s 'Button down
	Private Const WM_RBUTTONUP As Short = &H205s 'Button up
	
	'Declare the API function call.
	'UPGRADE_WARNING: Structure NOTIFYICONDATA may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1050"'
	Private Declare Function Shell_NotifyIcon Lib "shell32"  Alias "Shell_NotifyIconA"(ByVal dwMessage As Integer, ByRef pnid As NOTIFYICONDATA) As Boolean
	
	Dim nid As NOTIFYICONDATA
	
	' Declarations for hiding from the task manager
	Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Integer, ByVal nCmdShow As Integer) As Integer
	Private Declare Function GetWindow Lib "user32" (ByVal hWnd As Integer, ByVal wCmd As Integer) As Integer
	Const SW_HIDE As Short = 0
	Const SW_RESTORE As Short = 9
	Const GW_OWNER As Short = 4
	Dim OwnerhWnd As Integer
	
	Const ORIG_HEIGHT As Short = 2055
	' IDM declarations
	Dim oLibraries As IDMObjects.ObjectSet
	Dim oLib As IDMObjects.Library
	Dim bLoggedOn As Boolean
	Public oErrManager As IDMError.ErrorManager
	
	Public Sub ShowError()
		Dim oErrCollect As IDMError.Errors
		Dim oError As IDMError.Error
		Dim iCnt As Short
		oErrCollect = oErrManager.Errors
		If oErrCollect.Count > 1 Then
			iCnt = 1
			For	Each oError In oErrCollect
				MsgBox("Error " & iCnt & ": " & oError.Description & " : " & Hex(oError.Number))
				iCnt = iCnt + 1
			Next oError
		Else
			If oErrCollect.Count = 1 Then
				oErrManager.ShowErrorDialog()
			Else
				If Err.Number <> 0 Then
					MsgBox(Err.Description & " : " & Err.Number)
				End If
			End If
		End If
	End Sub
	
	
	' Logs on and hides if successful
	Private Function idmLogon() As Boolean
		
		On Error GoTo errHandler
		
		bLoggedOn = False
		idmLogon = False
		
		Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
		
		'Attempt to logon
		Dim oUser As IDMObjects.User
		If Not oLib.GetState(IDMObjects.idmLibraryState.idmLibraryLoggedOn) Then
			bLoggedOn = oLib.Logon(txtUsername, txtPassword, "", IDMObjects.idmLibraryLogon.idmLogonOptNoUI)
		Else
			oUser = oLib.ActiveUser
			If Not (oUser.Name = txtUsername.Text) Then
				'Can only be logged on as one user on the desktop
				'Do the logoff from the attach done by oLib.GetState
				oLib.Logoff()
				'Make the logon call to generat appropriate error message
				bLoggedOn = oLib.Logon(txtUsername, txtPassword, "", IDMObjects.idmLibraryLogon.idmLogonOptNoUI)
			Else
				bLoggedOn = True
			End If
		End If
		
		Me.Cursor = System.Windows.Forms.Cursors.Default
		
		idmLogon = bLoggedOn
		
		Exit Function
		
errHandler: 
		
		ShowError()
		Me.Cursor = System.Windows.Forms.Cursors.Default
		
	End Function
	
	Private Function idmLogoff() As Boolean
		
		oLib.Logoff()
		
		bLoggedOn = False
		
		idmLogoff = True
		
	End Function
	
	' Based on current state, either log on or log off
	Private Sub btnLogon_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnLogon.Click
		
		If (Not bLoggedOn) And cmbLibraries.Text <> "" Then
			
			If (idmLogon()) Then
				
				btnLogon.Text = "&Logoff"
				Me.Text = "IDM Logoff"
				
				btnShowChange.Enabled = True
				txtUsername.Enabled = False
				txtPassword.Text = ""
				txtNewPwd.Text = ""
				txtNewPwd2.Text = ""
				txtPassword.Enabled = False
				cmbLibraries.Enabled = False
				' Flesh out the tooltip text on the tray icon
				nid.szTip = oLib.Label & " logoff" & vbNullChar
				' Get ourselves put on the tray
				Shell_NotifyIcon(NIM_ADD, nid)
				' Discard oLibraries to save some memory...
				'UPGRADE_NOTE: Object oLibraries may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1029"'
				oLibraries = Nothing
				hideLogon()
				
			End If
			
		ElseIf (idmLogoff()) Then 
			
			Me.Height = VB6.TwipsToPixelsY(ORIG_HEIGHT)
			btnLogon.Text = "&Logon"
			Me.Text = "IDM Logon"
			btnShowChange.Enabled = False
			txtUsername.Enabled = True
			txtPassword.Text = ""
			txtNewPwd.Text = ""
			txtNewPwd2.Text = ""
			txtPassword.Enabled = True
			cmbLibraries.Enabled = True
			' Repopulate the cmbLibraries in case they changed
			Call PopulateLibraries((Me.cmbLibraries))
			' Get off the tray
			Shell_NotifyIcon(NIM_DELETE, nid)
			
		End If
		
	End Sub
	
	Private Sub btnShowChange_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnShowChange.Click
		
		Me.Height = VB6.TwipsToPixelsY(3015)
		
		txtPassword.Enabled = True
		btnShowChange.Enabled = False
		txtPassword.Focus()
		
	End Sub
	' Logic to change password
	Private Sub btnUpdatePwd_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnUpdatePwd.Click
		
		If (txtNewPwd.Text <> txtNewPwd2.Text) Then
			MsgBox("Your new password and confirmation do not match.  Try again.")
			Exit Sub
		End If
		
		Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
		
		On Error GoTo errHandler
		
		oLib.ChangePassword(txtPassword.Text, txtNewPwd.Text, IDMObjects.idmPasswordOptions.idmPasswordNoUI)
		
		MsgBox("Your password has been changed.")
		
		Me.Height = VB6.TwipsToPixelsY(ORIG_HEIGHT)
		txtPassword.Text = ""
		txtNewPwd.Text = ""
		txtNewPwd2.Text = ""
		btnShowChange.Enabled = True
		txtPassword.Enabled = False
		
		Me.Cursor = System.Windows.Forms.Cursors.Default
		
		Exit Sub
		
errHandler: 
		
		MsgBox("Failed to change password.  Make sure you have entered your old password as well as a new one.")
		
		Me.Cursor = System.Windows.Forms.Cursors.Default
		
	End Sub
	' Set the global oLib to point to user's library choice
	'UPGRADE_WARNING: Event cmbLibraries.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2075"'
	Private Sub cmbLibraries_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmbLibraries.SelectedIndexChanged
		oLib = oLibraries(cmbLibraries.SelectedIndex + 1)
	End Sub
	
	Private Sub PopulateLibraries(ByRef cmbLib As System.Windows.Forms.ComboBox)
		Dim nbHood As New IDMObjects.Neighborhood
		Dim oLib As IDMObjects.Library
		' Get global oLibraries so cmbLibraries_Click can use it
		oLibraries = nbHood.Libraries
		cmbLib.Items.Clear()
		For	Each oLib In oLibraries
			cmbLib.Items.Add(oLib.Label)
		Next oLib
		cmbLib.SelectedIndex = 0
	End Sub
	Private Sub frmLogon_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		
		oErrManager = CreateObject("IDMError.ErrorManager")
		
		On Error GoTo errHandler
		
		bLoggedOn = False
		
		' Populate the combo box with available libraries
		Call PopulateLibraries((Me.cmbLibraries))
		
		' Add to the system tray
		'Set the individual values of the NOTIFYICONDATA data type.
		nid.cbSize = Len(nid)
		nid.hWnd = Me.Handle.ToInt32
		nid.uId = VariantType.Null
		nid.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
		nid.uCallBackMessage = WM_MOUSEMOVE
		nid.hIcon = CInt(CObj(Me.Icon))
		
		'Set the height to not show password stuff
		Me.Height = VB6.TwipsToPixelsY(ORIG_HEIGHT)
		
		Exit Sub
		
errHandler: 
		
		ShowError()
		' Put error text where library would normally go
		lblCatalog.Text = "Error:"
		' txtCatalog = Err.Description
		btnLogon.Enabled = False
		
	End Sub
	
	'UPGRADE_WARNING: Event frmLogon.Resize may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2075"'
	Private Sub frmLogon_Resize(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Resize
		
		'Make sure we hide when the user minimizes window
		If (Me.WindowState = System.Windows.Forms.FormWindowState.Minimized) Then
			
			hideLogon()
			
		End If
		
	End Sub
	
	'UPGRADE_NOTE: Form_Terminate was upgraded to Form_Terminate_Renamed. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1061"'
	'UPGRADE_WARNING: frmLogon event Form.Terminate has a new behavior. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2065"'
	Private Sub Form_Terminate_Renamed()
		On Error GoTo errHandler
		' If we're logged on, better log off...
		If bLoggedOn Then
			oLib.Logoff()
		End If
		
		Exit Sub
		
errHandler: 
		
		ShowError()
		
	End Sub
	
	Private Sub frmLogon_MouseMove(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles MyBase.MouseMove
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim X As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim Y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		
		' This is not the real mousemove message, but is the message from
		' the Shell_NotifyIcon function.  See MS KB article Q162613 for more info.
		Dim msg As Integer
		msg = X / VB6.TwipsPerPixelX
		Select Case msg
			Case WM_LBUTTONDOWN, WM_LBUTTONUP, WM_LBUTTONDBLCLK, WM_RBUTTONDOWN, WM_RBUTTONUP, WM_RBUTTONDBLCLK
				' Make the app visible again
				ShowWindow(OwnerhWnd, SW_RESTORE)
				Me.Visible = True
				Me.WindowState = System.Windows.Forms.FormWindowState.Normal
		End Select
	End Sub
	
	Private Sub hideLogon()
		
		Dim ret As Integer
		
		' Grab the background or owner window:
		OwnerhWnd = GetWindow(Me.Handle.ToInt32, GW_OWNER)
		
		'Minimize window if not already
		If (Me.WindowState <> System.Windows.Forms.FormWindowState.Minimized) Then Me.WindowState = System.Windows.Forms.FormWindowState.Minimized
		
		' Hide from task list:
		ret = ShowWindow(OwnerhWnd, SW_HIDE)
		
		' Make sure the form is invisible:
		Me.Visible = False
		
	End Sub
End Class