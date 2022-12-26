VERSION 5.00
Begin VB.Form WSRename 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Rename Workspace"
   ClientHeight    =   3465
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7530
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   7530
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Height          =   495
      Left            =   2160
      TabIndex        =   1
      Top             =   2880
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   4080
      TabIndex        =   2
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Frame frDestQInfo 
      Caption         =   "New Queue Information"
      Height          =   1215
      Left            =   0
      TabIndex        =   8
      Top             =   1440
      Width           =   7455
      Begin VB.TextBox txtDestServiceDisp 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1680
         TabIndex        =   9
         Top             =   360
         Width           =   5655
      End
      Begin VB.TextBox txtDestWS 
         Height          =   285
         Left            =   1680
         TabIndex        =   0
         Top             =   720
         Width           =   5655
      End
      Begin VB.Label lblDestWS 
         Caption         =   "New Workspace:"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label lblDestService 
         Caption         =   "New Service:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Frame frSourceQInfo 
      Caption         =   "Current Queue Information"
      Height          =   1095
      Left            =   0
      TabIndex        =   3
      Top             =   120
      Width           =   7455
      Begin VB.Label lblSourceService 
         Caption         =   "Current Service:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label lblSourceWS 
         Caption         =   "Current Workspace:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label lblSourceServiceDisp 
         Height          =   285
         Left            =   1680
         TabIndex        =   5
         Top             =   360
         Width           =   5655
      End
      Begin VB.Label lblSourceWSDisp 
         Height          =   285
         Left            =   1680
         TabIndex        =   4
         Top             =   720
         Width           =   5655
      End
   End
End
Attribute VB_Name = "WSRename"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    
    lblSourceServiceDisp = "Default Service" 'oSourceQueue.ServiceName
    lblSourceWSDisp = QMaint.oWorkspace
    txtDestServiceDisp = "Default Service" ' oSourceQueue.ServiceName
    cmdOk.Enabled = (txtDestWS <> "")

End Sub

Private Sub Form_Activate()
    WSRename.txtDestWS.SetFocus
End Sub

Private Sub cmdOK_Click()
    Dim listCount As Integer
    Dim inx As Integer
    Dim txtOldWSName As String
    
    On Error GoTo ErrorHandler

    Screen.MousePointer = vbHourglass
    
    ' Save original workspace name.
    txtOldWSName = QMaint.oWorkspace.Name
    
    ' Change the workspace name and save it.
    QMaint.oWorkspace = txtDestWS.Text
    QMaint.oWorkspace.Save
       
    ' Remove old workspace from the combo box on qmaint form.
    QMaint.cmbWorkspace.RemoveItem (QMaint.cmbWorkspace.ListIndex)

    ' Add workspace to combo box on qmaint form.
    QMaint.cmbWorkspace.AddItem (txtDestWS.Text)
    
    ' Set the renamed workspace to be the currently displayed workspace.
    ' Note: The following implicitly calls cmbWorkspace.click()
    QMaint.cmbWorkspace.ListIndex = QMaint.cmbWorkspace.NewIndex
    
    Screen.MousePointer = vbDefault
    Unload Me
    
    Exit Sub
    
ErrorHandler:
    QMaint.oWorkspace = txtOldWSName        ' Restore original Workspace name.
    oErrorLog.logFNError errWarning, "Error renaming Workspace.", "Error renaming Workspace."
    Screen.MousePointer = vbDefault

End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub txtDestWS_Change()
    cmdOk.Enabled = (txtDestWS.Text <> "")
End Sub

Private Sub txtDestWS_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then   ' If the character entered is a space, ignore it.
        KeyAscii = 0
    End If
End Sub
