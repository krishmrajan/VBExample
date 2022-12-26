VERSION 5.00
Begin VB.Form frmLogon 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Logon"
   ClientHeight    =   2895
   ClientLeft      =   2340
   ClientTop       =   1605
   ClientWidth     =   4530
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   4530
   StartUpPosition =   2  'CenterScreen
   Tag             =   "00900"
   Begin VB.TextBox txtGroup 
      BackColor       =   &H00C0C0C0&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1440
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1200
      Width           =   2880
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2400
      TabIndex        =   5
      Tag             =   "00701"
      Top             =   2280
      Width           =   1200
   End
   Begin VB.ComboBox cboLibraries 
      Height          =   288
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1680
      Width           =   2880
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1440
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   720
      Width           =   2880
   End
   Begin VB.TextBox txtUserName 
      Height          =   285
      Left            =   1440
      TabIndex        =   0
      Text            =   "SysAdmin"
      Top             =   210
      Width           =   2880
   End
   Begin VB.CommandButton cmdLogon 
      Caption         =   "&Logon"
      Default         =   -1  'True
      Height          =   375
      Left            =   840
      TabIndex        =   4
      Tag             =   "00700"
      Top             =   2280
      Width           =   1200
   End
   Begin VB.Label lblGroup 
      Caption         =   "&Group:"
      Height          =   252
      Left            =   240
      TabIndex        =   9
      Tag             =   "00202"
      Top             =   1212
      Width           =   1092
   End
   Begin VB.Label lblLibrary 
      Caption         =   "Li&brary:"
      Height          =   252
      Left            =   240
      TabIndex        =   8
      Tag             =   "00203"
      Top             =   1740
      Width           =   972
   End
   Begin VB.Label lblPassword 
      Caption         =   "&Password:"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Tag             =   "00201"
      Top             =   735
      Width           =   1095
   End
   Begin VB.Label lblUserName 
      Caption         =   "&User Name:"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Tag             =   "00200"
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "frmLogon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpbuffer As String, nSize As Long) As Long

Private Sub FillLogonIDs()
'This subroutine fills the logon user IDs and passwords for
'the logon form.  This is only for ease of demonstration.
'A normal user should never have SysAdmin/admin capabilities.

Dim loLibrary As IDMObjects.Library
Dim lsBuffer As String
Dim llSize As Long

    For Each loLibrary In goNeighborhood.Libraries
        If loLibrary.Label = cboLibraries.Text Then
            Exit For  'we have found the library we want to logon to
        End If
    Next
    If loLibrary.SystemType = idmSysTypeIS Then
        'NOTE:  This SysAdmin Logon option is used to assist on FileNET demo systems.
        '       A user program should not have the option of using this setting.
        If gbSettingUseSysAdminLogon = False Then
            'Get Windows user name and pre-fill UserName text box.
            lsBuffer = Space$(255)
            llSize = Len(lsBuffer)
            Call GetUserName(lsBuffer, llSize)
            If llSize > 0 Then
                txtUserName.Text = Left$(lsBuffer, llSize)
            Else
                txtUserName.Text = vbNullString
            End If
        Else
            txtUserName.Text = GS_SYSADMIN_LOGON_NAME
            txtPassword.Text = GS_SYSADMIN_PASSWORD
        End If
        txtGroup.Enabled = False
        txtGroup.BackColor = &HC0C0C0  'Gray
    ElseIf loLibrary.SystemType = idmSysTypeDS Then
        If gbSettingUseSysAdminLogon = False Then
            'Get Windows user name and pre-fill UserName text box.
            lsBuffer = Space$(255)
            llSize = Len(lsBuffer)
            Call GetUserName(lsBuffer, llSize)
            If llSize > 0 Then
                txtUserName.Text = Left$(lsBuffer, llSize)
            Else
                txtUserName.Text = vbNullString
            End If
        Else
            txtUserName.Text = GS_ADMIN_LOGON_NAME
            txtPassword.Text = GS_ADMIN_PASSWORD
            txtGroup.Text = GS_ADMIN_GROUP
        End If
        txtGroup.Enabled = True
        txtGroup.BackColor = RGB(255, 255, 255) 'White
    End If


End Sub

Private Sub cboLibraries_Click()
    
    FillLogonIDs
    
End Sub

Private Sub cmdCancel_Click()
    
    'The user may have logged on through other means, so
    ' let the app proceed...
    Unload Me
        
End Sub

Private Sub cmdLogon_Click()

Dim lsSelectedLibraryLabel As String
Dim liUserResponse As Integer

    gbSuccess = False
    
    'Get the label of the library the user selected.
    lsSelectedLibraryLabel = cboLibraries.Text
            
    Do
        MouseWait
        'Call routine to logon to FileNET library selected
        gbSuccess = LogonToLibrary(lsSelectedLibraryLabel, txtUserName.Text, txtPassword.Text, txtGroup.Text)
        MouseNormal
        If gbSuccess = False Then
            liUserResponse = MsgBox(LoadResString(GI_ERR_LOGON_UNSUCCESSFUL), vbExclamation + vbYesNo, LoadResString(GI_ERR_ERROR))
            Select Case liUserResponse
                Case vbNo
                    'User may have other libraries open, so go ahead with app...
                    AppTerminate False
                Case vbYes
                    'Break out of loop to show form again
                    Exit Do
            End Select
        Else   'Logon was successful
            Unload Me
        End If
    Loop Until gbSuccess = True
        
End Sub

Private Sub Form_Load()

Dim loLibrary As IDMObjects.Library
Dim iInx As Integer
Dim iSelector As Integer

On Error GoTo ErrorHandler
    
    'Load resource strings for form elements to create
    'language-specific UI.  Edit the .RC file to change languages.
    LoadResStrings Me
    
    'Add list of all libraries found in the global FileNET
    'Neighborhood to the Combo Box
    'NOTE:  Label property is the short name (i.e. domain on IDMIS)
    '       Name property is the full 3-part NCH name on IDMIS
    iInx = 0
    For Each loLibrary In goNeighborhood.Libraries
        cboLibraries.AddItem loLibrary.Label
        If loLibrary.Name = goNeighborhood.DefaultLibrary.Name Then
            iSelector = iInx
        End If
        iInx = iInx + 1
    Next
    
    'Pre-select the Default Library for the user in the Combo Box
    'NOTE:  Make sure the property 'Label' or 'Name' that you used
    '       above is also used here or an error will occur.
    cboLibraries.ListIndex = iSelector
    FillLogonIDs

Exit Sub

ErrorHandler:

    'display error message
    DisplayErrorMessage ("Form1.Form_Load")

    'cleanup error codes
    CleanupErrorCodes

    Resume Next

End Sub

