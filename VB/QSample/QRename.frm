VERSION 5.00
Begin VB.Form QRename 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Rename Queue"
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7515
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   7515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frSourceQInfo 
      Caption         =   "Current Queue Information"
      Height          =   1455
      Left            =   0
      TabIndex        =   9
      Top             =   120
      Width           =   7455
      Begin VB.Label lblSourceQDisp 
         Height          =   285
         Left            =   1680
         TabIndex        =   15
         Top             =   1080
         Width           =   5655
      End
      Begin VB.Label lblSourceWSDisp 
         Height          =   285
         Left            =   1680
         TabIndex        =   14
         Top             =   720
         Width           =   5655
      End
      Begin VB.Label lblSourceServiceDisp 
         Height          =   285
         Left            =   1680
         TabIndex        =   13
         Top             =   360
         Width           =   5655
      End
      Begin VB.Label lblSourceQ 
         Caption         =   "Current Queue:"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label lblSourceWS 
         Caption         =   "Current Workspace:"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label lblSourceService 
         Caption         =   "Current Service:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame frDestQInfo 
      Caption         =   "New Queue Information"
      Height          =   1455
      Left            =   0
      TabIndex        =   3
      Top             =   1680
      Width           =   7455
      Begin VB.TextBox txtDestWS 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1680
         TabIndex        =   5
         Top             =   720
         Width           =   5655
      End
      Begin VB.TextBox txtDestQ 
         Height          =   285
         Left            =   1680
         TabIndex        =   0
         Top             =   1080
         Width           =   5655
      End
      Begin VB.TextBox txtDestServiceDisp 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1680
         TabIndex        =   4
         Top             =   360
         Width           =   5655
      End
      Begin VB.Label lblDestService 
         Caption         =   "New Service:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label lblDestWS 
         Caption         =   "New Workspace:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label lblDestQ 
         Caption         =   "New Queue:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1080
         Width           =   1815
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   4080
      TabIndex        =   2
      Top             =   3240
      Width           =   1335
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Height          =   495
      Left            =   2160
      TabIndex        =   1
      Top             =   3240
      Width           =   1335
   End
End
Attribute VB_Name = "QRename"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private oSourceQueue As IDMObjects.queue                        ' The source queue
Private oDestQueue As IDMObjects.queue                          ' The dest queue
Private oDestWorkspace As IDMObjects.QueueWorkspace             ' The dest queue

Option Explicit

Private Sub Form_Load()
    
    Set oSourceQueue = QMaint.oQueue
    lblSourceServiceDisp = oSourceQueue.ServiceName
    lblSourceWSDisp = oSourceQueue.Workspace
    lblSourceQDisp = oSourceQueue.Name
    txtDestServiceDisp = oSourceQueue.ServiceName
    txtDestWS = oSourceQueue.Workspace
    cmdOK.Enabled = (txtDestWS <> "") And (txtDestQ <> "")

End Sub

Private Sub Form_Activate()
    QRename.txtDestQ.SetFocus
End Sub
Private Sub cmdOK_Click()
    Dim listCount As Integer
    Dim inx As Integer
    
    On Error GoTo ErrorHandler

    Screen.MousePointer = vbHourglass
    oSourceQueue.Name = txtDestQ.Text
    oSourceQueue.Save
    
    ' Remove old queue from the combo box on qmaint form.
    QMaint.cmbQueue.RemoveItem (QMaint.cmbQueue.ListIndex)

    ' Add Queue to combo box on qmaint form.
    QMaint.cmbQueue.AddItem (txtDestQ.Text)
    
    ' Set the renamed queue to be the currently displayed queue.
    ' Note: The following implicitly calls cmbQueue.click()
    QMaint.cmbQueue.ListIndex = QMaint.cmbQueue.NewIndex
    Set oSourceQueue = Nothing

    Screen.MousePointer = vbDefault
    Unload Me
    
    Exit Sub
    
ErrorHandler:
    oErrorLog.logFNError errWarning, "Error renaming queue.", "Error renaming queue."
    Screen.MousePointer = vbDefault
    Set oSourceQueue = Nothing
    QMaint.oQueue.Name = lblSourceQDisp     ' Revert to original name.

End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub txtDestQ_Change()
    cmdOK.Enabled = (txtDestWS <> "") And (txtDestQ <> "")
End Sub

Private Sub txtDestQ_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then   ' If the character entered is a space, ignore it.
        KeyAscii = 0
    End If
End Sub
