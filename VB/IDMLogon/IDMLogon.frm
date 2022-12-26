VERSION 5.00
Begin VB.Form frmLogon 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "IDM Logon"
   ClientHeight    =   2640
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5835
   Icon            =   "IDMLogon.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2640
   ScaleWidth      =   5835
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbLibraries 
      Height          =   315
      Left            =   1800
      TabIndex        =   12
      Text            =   "Combo1"
      Top             =   240
      Width           =   2295
   End
   Begin VB.CommandButton btnUpdatePwd 
      Caption         =   "&Update"
      Height          =   375
      Left            =   4320
      TabIndex        =   11
      Top             =   2040
      Width           =   1335
   End
   Begin VB.TextBox txtNewPwd2 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1800
      PasswordChar    =   "*"
      TabIndex        =   9
      Top             =   2160
      Width           =   2295
   End
   Begin VB.TextBox txtNewPwd 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1800
      PasswordChar    =   "*"
      TabIndex        =   7
      Top             =   1680
      Width           =   2295
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1800
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1200
      Width           =   2295
   End
   Begin VB.TextBox txtUsername 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1800
      TabIndex        =   0
      Top             =   720
      Width           =   2295
   End
   Begin VB.CommandButton btnShowChange 
      Caption         =   "&Change Pwd >>"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4320
      TabIndex        =   3
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton btnLogon 
      Caption         =   "&Logon"
      Default         =   -1  'True
      Height          =   375
      Left            =   4320
      TabIndex        =   2
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Confirm:"
      Height          =   255
      Left            =   720
      TabIndex        =   10
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "New Password:"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Password:"
      Height          =   255
      Left            =   840
      TabIndex        =   6
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Username:"
      Height          =   255
      Left            =   840
      TabIndex        =   5
      Top             =   720
      Width           =   855
   End
   Begin VB.Label lblCatalog 
      Alignment       =   1  'Right Justify
      Caption         =   "Library:"
      Height          =   255
      Left            =   840
      TabIndex        =   4
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "frmLogon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Sample logon program Copyright(C) 1997 FileNet Corporation
'
'Declare a user-defined variable to pass to the Shell_NotifyIcon
'function.
Option Explicit

Private Type NOTIFYICONDATA
   cbSize As Long
   hWnd As Long
   uId As Long
   uFlags As Long
   uCallBackMessage As Long
   hIcon As Long
   szTip As String * 64
End Type

'Declare the constants for the API function. These constants can be
'found in the header file Shellapi.h.

'The following constants are the messages sent to the
'Shell_NotifyIcon function to add, modify, or delete an icon from the
'taskbar status area.
Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2

'The following constant is the message sent when a mouse event occurs
'within the rectangular boundaries of the icon in the taskbar status
'area.
Private Const WM_MOUSEMOVE = &H200

'The following constants are the flags that indicate the valid
'members of the NOTIFYICONDATA data type.
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4

'The following constants are used to determine the mouse input on the
'the icon in the taskbar status area.

'Left-click constants.
Private Const WM_LBUTTONDBLCLK = &H203   'Double-click
Private Const WM_LBUTTONDOWN = &H201     'Button down
Private Const WM_LBUTTONUP = &H202       'Button up

'Right-click constants.
Private Const WM_RBUTTONDBLCLK = &H206   'Double-click
Private Const WM_RBUTTONDOWN = &H204     'Button down
Private Const WM_RBUTTONUP = &H205       'Button up

'Declare the API function call.
Private Declare Function Shell_NotifyIcon Lib "shell32" _
 Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
 
Dim nid As NOTIFYICONDATA

' Declarations for hiding from the task manager
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Const SW_HIDE = 0
Const SW_RESTORE = 9
Const GW_OWNER = 4
Dim OwnerhWnd As Long

Const ORIG_HEIGHT = 2055
' IDM declarations
Dim oLibraries As ObjectSet
Dim oLib As IDMObjects.Library
Dim bLoggedOn As Boolean
Public oErrManager As idmError.ErrorManager

Public Sub ShowError()
Dim oErrCollect As idmError.Errors
Dim oError As idmError.Error
Dim iCnt As Integer
Set oErrCollect = oErrManager.Errors
If oErrCollect.Count > 1 Then
    iCnt = 1
    For Each oError In oErrCollect
        MsgBox "Error " & iCnt & ": " & oError.Description & " : " & Hex(oError.Number)
        iCnt = iCnt + 1
    Next
Else
    If oErrCollect.Count = 1 Then
        oErrManager.ShowErrorDialog
    Else
        If Err.Number <> 0 Then
            MsgBox Err.Description & " : " & Err.Number
        End If
    End If
End If
End Sub

        
' Logs on and hides if successful
Private Function idmLogon() As Boolean
    
    On Error GoTo errHandler
    
    bLoggedOn = False
    idmLogon = False
    
    Me.MousePointer = vbHourglass
        
    'Attempt to logon
    If Not oLib.GetState(idmLibraryLoggedOn) Then
        bLoggedOn = oLib.Logon(txtUsername, txtPassword, "", idmLogonOptNoUI)
    Else
        Dim oUser As IDMObjects.User
        Set oUser = oLib.ActiveUser
        If Not (oUser.Name = txtUsername) Then
            'Can only be logged on as one user on the desktop
            'Do the logoff from the attach done by oLib.GetState
            oLib.Logoff
            'Make the logon call to generat appropriate error message
            bLoggedOn = oLib.Logon(txtUsername, txtPassword, "", idmLogonOptNoUI)
        Else
            bLoggedOn = True
        End If
    End If
    
    Me.MousePointer = vbDefault
    
    idmLogon = bLoggedOn
    
    Exit Function
    
errHandler:

    ShowError
    Me.MousePointer = vbDefault
    
End Function

Private Function idmLogoff() As Boolean

    oLib.Logoff
    
    bLoggedOn = False
    
    idmLogoff = True

End Function

' Based on current state, either log on or log off
Private Sub btnLogon_Click()

    If (Not bLoggedOn) And cmbLibraries <> "" Then
    
        If (idmLogon()) Then
            
            btnLogon.Caption = "&Logoff"
            Me.Caption = "IDM Logoff"
            
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
            Shell_NotifyIcon NIM_ADD, nid
            ' Discard oLibraries to save some memory...
            Set oLibraries = Nothing
            hideLogon
            
        End If
    
    ElseIf (idmLogoff()) Then
        
        Me.Height = ORIG_HEIGHT
        btnLogon.Caption = "&Logon"
        Me.Caption = "IDM Logon"
        btnShowChange.Enabled = False
        txtUsername.Enabled = True
        txtPassword.Text = ""
        txtNewPwd.Text = ""
        txtNewPwd2.Text = ""
        txtPassword.Enabled = True
        cmbLibraries.Enabled = True
        ' Repopulate the cmbLibraries in case they changed
        Call PopulateLibraries(Me.cmbLibraries)
        ' Get off the tray
        Shell_NotifyIcon NIM_DELETE, nid
             
    End If
        
End Sub

Private Sub btnShowChange_Click()

    Me.Height = 3015
    
    txtPassword.Enabled = True
    btnShowChange.Enabled = False
    txtPassword.SetFocus
    
End Sub
' Logic to change password
Private Sub btnUpdatePwd_Click()

    If (txtNewPwd.Text <> txtNewPwd2.Text) Then
        MsgBox ("Your new password and confirmation do not match.  Try again.")
        Exit Sub
    End If
    
    Me.MousePointer = vbHourglass
    
    On Error GoTo errHandler
    
    oLib.ChangePassword txtPassword.Text, txtNewPwd.Text, idmPasswordNoUI
    
    MsgBox ("Your password has been changed.")
    
    Me.Height = ORIG_HEIGHT
    txtPassword.Text = ""
    txtNewPwd.Text = ""
    txtNewPwd2.Text = ""
    btnShowChange.Enabled = True
    txtPassword.Enabled = False
    
    Me.MousePointer = vbDefault
    
    Exit Sub
    
errHandler:

    MsgBox "Failed to change password.  Make sure you have entered your old password as well as a new one."
    
    Me.MousePointer = vbDefault
    
End Sub
' Set the global oLib to point to user's library choice
Private Sub cmbLibraries_Click()
Set oLib = oLibraries(cmbLibraries.ListIndex + 1)
End Sub

Private Sub PopulateLibraries(cmbLib As ComboBox)
Dim nbHood As New IDMObjects.Neighborhood
Dim oLib As IDMObjects.Library
' Get global oLibraries so cmbLibraries_Click can use it
Set oLibraries = nbHood.Libraries
cmbLib.Clear
For Each oLib In oLibraries
    cmbLib.AddItem oLib.Label
Next
cmbLib.ListIndex = 0
End Sub
Private Sub Form_Load()

Set oErrManager = CreateObject("IDMError.ErrorManager")

On Error GoTo errHandler

    bLoggedOn = False

    ' Populate the combo box with available libraries
    Call PopulateLibraries(Me.cmbLibraries)
    
    ' Add to the system tray
    'Set the individual values of the NOTIFYICONDATA data type.
    nid.cbSize = Len(nid)
    nid.hWnd = Me.hWnd
    nid.uId = vbNull
    nid.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    nid.uCallBackMessage = WM_MOUSEMOVE
    nid.hIcon = Me.Icon
    
    'Set the height to not show password stuff
    Me.Height = ORIG_HEIGHT
    
    Exit Sub
    
errHandler:

    ShowError
    ' Put error text where library would normally go
    lblCatalog.Caption = "Error:"
    ' txtCatalog = Err.Description
    btnLogon.Enabled = False
    
End Sub

Private Sub Form_Resize()

    'Make sure we hide when the user minimizes window
    If (Me.WindowState = vbMinimized) Then
        
        hideLogon
        
    End If
    
End Sub

Private Sub Form_Terminate()
On Error GoTo errHandler
    ' If we're logged on, better log off...
    If bLoggedOn Then
        oLib.Logoff
    End If
    
    Exit Sub
    
errHandler:
    
    ShowError

End Sub

Private Sub Form_MouseMove _
            (Button As Integer, _
            Shift As Integer, _
            X As Single, _
            Y As Single)
                 
    ' This is not the real mousemove message, but is the message from
    ' the Shell_NotifyIcon function.  See MS KB article Q162613 for more info.
    Dim msg As Long
    msg = X / Screen.TwipsPerPixelX
    Select Case msg
        Case WM_LBUTTONDOWN, WM_LBUTTONUP, WM_LBUTTONDBLCLK, WM_RBUTTONDOWN, WM_RBUTTONUP, WM_RBUTTONDBLCLK
            ' Make the app visible again
            ShowWindow OwnerhWnd, SW_RESTORE
            Me.Visible = True
            Me.WindowState = vbNormal
    End Select
End Sub

Private Sub hideLogon()

    Dim ret As Long
    
    ' Grab the background or owner window:
    OwnerhWnd = GetWindow(Me.hWnd, GW_OWNER)
    
    'Minimize window if not already
    If (Me.WindowState <> vbMinimized) Then Me.WindowState = vbMinimized
    
    ' Hide from task list:
    ret = ShowWindow(OwnerhWnd, SW_HIDE)
    
    ' Make sure the form is invisible:
    Me.Visible = False

End Sub

