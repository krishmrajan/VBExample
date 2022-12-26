VERSION 5.00
Begin VB.Form frmCommandIDMIS 
   Caption         =   "Create Command"
   ClientHeight    =   4644
   ClientLeft      =   132
   ClientTop       =   588
   ClientWidth     =   6252
   LinkTopic       =   "Form1"
   ScaleHeight     =   4644
   ScaleWidth      =   6252
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClrFolder 
      Caption         =   "Clear folder"
      Height          =   375
      Left            =   1560
      TabIndex        =   12
      Top             =   2040
      Width           =   1092
   End
   Begin VB.CommandButton btnBrowse 
      Caption         =   "Select folder"
      Height          =   375
      Left            =   360
      TabIndex        =   10
      Top             =   2040
      Width           =   1092
   End
   Begin VB.TextBox txtFolderName 
      Enabled         =   0   'False
      Height          =   405
      Left            =   1560
      TabIndex        =   9
      Top             =   2520
      Width           =   2655
   End
   Begin VB.CheckBox chkSubFolder 
      Caption         =   "Search Subfolder"
      Height          =   252
      Left            =   360
      TabIndex        =   8
      Top             =   3120
      Width           =   1692
   End
   Begin VB.CommandButton btnKeySQL 
      Caption         =   "Key Condition"
      Height          =   375
      Left            =   2760
      TabIndex        =   7
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton btnSortingSQL 
      Caption         =   "Sorting"
      Height          =   375
      Left            =   1560
      TabIndex        =   6
      Top             =   1440
      Width           =   1092
   End
   Begin VB.CommandButton btnSimpleSQL 
      Caption         =   "Simple"
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Top             =   1440
      Width           =   1092
   End
   Begin VB.CheckBox chkObjSet 
      Caption         =   "&Show Results in IDM ListView"
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   3480
      Width           =   2532
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton cmdExecute 
      Caption         =   "&Execute"
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   4080
      Width           =   1215
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
      TabIndex        =   11
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Please enter command text or click a button to get sample command text:"
      Height          =   252
      Left            =   360
      TabIndex        =   3
      Top             =   120
      Width           =   5532
   End
   Begin VB.Menu mnuProperties 
      Caption         =   "ADO &Properties"
   End
End
Attribute VB_Name = "frmCommandIDMIS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public cmd As New ADODB.Command
Dim bOnceText
'
' You can change the following SQL statements to
' make the simple button do what you want...
'
Const strSimpleSQL = "select F_DOCNUMBER, F_ENTRYDATE from FnDocument where F_DOCNUMBER > 100000"
Const strSortingSQL = "select F_DOCNUMBER, F_ENTRYDATE from FnDocument where F_DOCNUMBER > 100000 order by F_ENTRYDATE"
Const strKeySQL = "select F_DOCNUMBER, F_ENTRYDATE from FnDocument key condition F_DOCNUMBER > 100000"
' Bring up the folder browsing dialog
Private Sub btnBrowse_Click()
    frmFolderList.Show vbModal, Me
    If frmFolderList.strFolderName <> "" Then
        txtFolderName.Text = frmFolderList.strFolderName
    End If
End Sub

Private Sub cmdClose_Click()
    Set cmd = Nothing
    Unload Me
End Sub

Private Sub cmdClrFolder_Click()
txtFolderName = ""
End Sub

' Execute button; build up the command object, then call
' the appropriate rowset form
Private Sub cmdExecute_Click()
On Error GoTo Handle
    cmd.CommandText = txtCommand.Text
    cmd.Properties("SearchFolderName") = txtFolderName.Text
    If txtFolderName.Text <> "" Then
        If chkSubFolder.Value = 1 Then
            cmd.Properties("SearchSubfolder") = True
        Else
            cmd.Properties("SearchSubfolder") = False
        End If
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
    cmdExecute.Enabled = False
    cmdClose.Enabled = True
    bOnceText = True
    
End Sub

Private Sub mnuProperties_Click()
    Set frmConnectionProperties.obj = cmd
    frmConnectionProperties.Show vbModal, Me
End Sub

Private Sub btnSimpleSQL_Click()
    txtCommand.Text = strSimpleSQL
End Sub

Private Sub btnSortingSQL_Click()
    txtCommand.Text = strSortingSQL
End Sub

Private Sub btnKeySQL_Click()
    txtCommand.Text = strKeySQL
End Sub


Private Sub txtCommand_Change()
    If bOnceText Then
        cmdClrFolder.Enabled = True
        cmdExecute.Enabled = True
        bOnceText = False
    End If
End Sub

