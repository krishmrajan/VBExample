VERSION 5.00
Begin VB.Form frmCommandIDMDS 
   Caption         =   "Create Command"
   ClientHeight    =   5376
   ClientLeft      =   132
   ClientTop       =   588
   ClientWidth     =   6360
   LinkTopic       =   "Form1"
   ScaleHeight     =   5376
   ScaleWidth      =   6360
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnSimpleSQL 
      Caption         =   "Simple"
      Height          =   375
      Left            =   360
      TabIndex        =   16
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton btnBrowse 
      Caption         =   "Select folder"
      Height          =   375
      Left            =   360
      TabIndex        =   15
      Top             =   1920
      Width           =   1095
   End
   Begin VB.TextBox txtFolderName 
      Enabled         =   0   'False
      Height          =   315
      Left            =   1560
      TabIndex        =   13
      Top             =   2400
      Width           =   3135
   End
   Begin VB.CommandButton btnSortingSQL 
      Caption         =   "Sorting"
      Height          =   375
      Left            =   1560
      TabIndex        =   12
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CheckBox chkObjSet 
      Caption         =   "&Show Results in IDM ListView"
      Height          =   255
      Left            =   2280
      TabIndex        =   11
      Top             =   3960
      Width           =   2412
   End
   Begin VB.ComboBox lstAccessLevel 
      Height          =   315
      ItemData        =   "CreateCommandMezz.frx":0000
      Left            =   2280
      List            =   "CreateCommandMezz.frx":0013
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   3360
      Width           =   1935
   End
   Begin VB.OptionButton chkADAll 
      Caption         =   "All Groups"
      Height          =   255
      Left            =   480
      TabIndex        =   8
      Top             =   3960
      Width           =   1455
   End
   Begin VB.OptionButton chkADActive 
      Caption         =   "Active Group"
      Height          =   255
      Left            =   480
      TabIndex        =   7
      Top             =   3600
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Access Domain"
      Height          =   975
      Left            =   360
      TabIndex        =   6
      Top             =   3360
      Width           =   1695
   End
   Begin VB.CheckBox chkSecure 
      Caption         =   "Secured Search"
      Height          =   252
      Left            =   360
      TabIndex        =   5
      Top             =   3000
      Value           =   1  'Checked
      Width           =   1695
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton cmdExecute 
      Caption         =   "&Execute"
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton cmdClrFolder 
      Caption         =   "Clear folder"
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   1920
      Width           =   1095
   End
   Begin VB.TextBox txtCommand 
      Height          =   975
      Left            =   360
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   360
      Width           =   5655
   End
   Begin VB.Label lblFolderName 
      Caption         =   "Folder Name:"
      Height          =   255
      Left            =   360
      TabIndex        =   14
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Access Level"
      Height          =   255
      Left            =   2280
      TabIndex        =   10
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Please enter command text or click a button to get sample command text:"
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   120
      Width           =   5655
   End
   Begin VB.Menu mnuProperties 
      Caption         =   "ADO &Properties"
   End
End
Attribute VB_Name = "frmCommandIDMDS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public cmd As New ADODB.Command
Dim bOnceText
'
' You can change the following IDMDS query strings
' to setup the simple queries as you wish
'
Const strSimpleSQL = "select idmId from FnDocument where idmId > 003000000"
Const strSortingSQL = "select idmId, idmName from FnDocument where idmId > 003000000 order by idmName"
' Browse the folder list
Private Sub btnBrowse_Click()
    frmFolderList.Show vbModal, Me
    If frmFolderList.strFolderName <> "" Then
        txtFolderName.Text = frmFolderList.strFolderName
    End If
End Sub
' Fill the SQL text area with our sample
Private Sub btnSimpleSQL_Click()
    txtCommand.Text = strSimpleSQL
End Sub

Private Sub btnSortingSQL_Click()
    txtCommand.Text = strSortingSQL
End Sub

Private Sub cmdClose_Click()
    Set cmd = Nothing
    Unload Me
End Sub

Private Sub cmdClrFolder_Click()
txtFolderName = ""
End Sub

' Setup the query properties, then call the right
' rowset function
Private Sub cmdExecute_Click()
On Error GoTo Handle
    cmd.CommandText = txtCommand.Text
    If txtFolderName.Text <> "" Then
        'If chkSubFolder.Value = 1 Then
         '   cmd.Properties("SearchSubfolder") = True
        'Else
         '   cmd.Properties("SearchSubfolder") = False
        'End If
        cmd.Properties("SearchFolderName") = frmFolderList.strFolderId
    Else
        cmd.Properties("SearchFolderName") = ""
    End If
    cmd.Properties("Access Level") = lstAccessLevel.ListIndex
    cmd.Properties("Secure Search") = chkSecure.Value
    If chkADActive.Value Then
        cmd.Properties("Access Domain") = 0
    Else
        cmd.Properties("Access Domain") = 1
    End If
    
    If chkObjSet.Value <> 1 Then
        Screen.MousePointer = vbHourglass
        cmd.Properties("SupportsObjSet") = False
        Set frmRowset.rs = cmd.Execute(, , adCmdText)
        Screen.MousePointer = vbDefault
        frmRowset.Show vbModal, Me
    Else
        Screen.MousePointer = vbHourglass
        cmd.Properties("SupportsObjSet") = True
        Set frmRowsetListView.rs = cmd.Execute(, , adCmdText)
        Screen.MousePointer = vbDefault
        frmRowsetListView.Show vbModal, Me
    End If
    Exit Sub
Handle:
    ShowError
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Load()
    Dim iADValue As Integer
    Dim iAccessLevel As Integer
    Dim bSecure As Boolean
    
    Set cmd.ActiveConnection = dbGlobals.ds
    cmdClrFolder.Enabled = False
    cmdExecute.Enabled = False
    cmdClose.Enabled = True
    bOnceText = True
    
    'Get command property values to init the controls
    If cmd.Properties("Secure Search") Then
        chkSecure = 1
    Else
        chkSecure = 0
    End If
    iADValue = cmd.Properties("Access Domain")
    If iADValue = 0 Then
        chkADActive = True
        chkADAll = False
    Else
        chkADActive = False
        chkADAll = True
    End If
    iAccessLevel = cmd.Properties("Access Level")
    lstAccessLevel.ListIndex = iAccessLevel
End Sub

Private Sub mnuProperties_Click()
    Set frmConnectionProperties.obj = cmd
    frmConnectionProperties.Show vbModal, Me
End Sub

Private Sub txtCommand_Change()
    If bOnceText Then
        cmdClrFolder.Enabled = True
        cmdExecute.Enabled = True
        bOnceText = False
    End If
End Sub

