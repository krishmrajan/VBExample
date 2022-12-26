VERSION 5.00
Begin VB.Form DelConf 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Confirm Deletion"
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   360
      Picture         =   "DelConf.frx":0000
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   8
      Top             =   120
      Width           =   540
   End
   Begin VB.CheckBox cbFull 
      Alignment       =   1  'Right Justify
      Caption         =   "Even if full"
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Top             =   1920
      Width           =   1695
   End
   Begin VB.CommandButton btnOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   1018
      TabIndex        =   1
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CommandButton btnCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2567
      TabIndex        =   0
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label txtQueue 
      Height          =   255
      Left            =   1800
      TabIndex        =   7
      Top             =   1440
      Width           =   2535
   End
   Begin VB.Label txtWS 
      Height          =   255
      Left            =   1800
      TabIndex        =   6
      Top             =   960
      Width           =   2535
   End
   Begin VB.Label lblQueue 
      Caption         =   "Queue:"
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label lblWS 
      Caption         =   "Workspace:"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label lblStatic 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Are you sure you wish to delete?"
      Height          =   195
      Left            =   1320
      TabIndex        =   2
      Top             =   240
      Width           =   2325
   End
End
Attribute VB_Name = "DelConf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnCancel_Click()
    QMaint.MainStatusBar.SimpleText = "Deletion cancelled."
    QMaint.MainStatusBar.Refresh
    Unload Me
End Sub

Private Sub btnOK_Click()
    Dim listCount As Integer
    Dim inx As Integer
    
    On Error GoTo ErrorHandler
    Screen.MousePointer = vbHourglass
    QMaint.MainStatusBar.SimpleText = "Deleting queue..."
    QMaint.MainStatusBar.Refresh
    Sleep (5000)
    Set QMaint.oQueueQuerySpec = Nothing
    QMaint.oQueue.Delete (cbFull.Value = 1)
    listCount = QMaint.cmbQueue.listCount
    For inx = 0 To listCount - 1
        If QMaint.cmbQueue.List(inx) = QMaint.oQueue.Name Then
            QMaint.cmbQueue.RemoveItem (inx)
            Exit For
        End If
    Next inx
    QMaint.MainStatusBar.SimpleText = "Queue deleted."
    QMaint.QueueInfoFrame = "Queue Information"
    QMaint.ViewButton.Enabled = False
    QMaint.EditButton.Enabled = False
    QMaint.DeleteButton.Enabled = False
    QMaint.BusyButton.Enabled = False
    QMaint.UnbusyButton.Enabled = False
    QMaint.QueryButton.Enabled = False
    QMaint.InsertButton.Enabled = False
    QMaint.mnuQueueModify.Enabled = False
    QMaint.mnuQueueQuery.Enabled = False
    QMaint.mnuQueueCopy.Enabled = False
    QMaint.mnuQueueRename.Enabled = False
    QMaint.mnuQueueDelete.Enabled = False
    QMaint.mnuQueuePrint.Enabled = False
    Set QMaint.oQueue = Nothing
    QMaint.MainStatusBar.Refresh
    Screen.MousePointer = vbDefault
    QMaint.grdQueueData.Clear
    QMaint.grdQueueData.Rows = 0
    Unload Me
    
    Exit Sub
ErrorHandler:
    oErrorLog.logFNError errWarning, "Error deleting queue.", "Error deleting queue."
    QMaint.MainStatusBar.SimpleText = "Deleting queue failed..."
    QMaint.MainStatusBar.Refresh
    Call QMaint.cmbQueue_Click   ' Restore original settings...
    Screen.MousePointer = vbDefault
    Unload Me
End Sub

Private Sub Form_Load()
    txtWS = QMaint.oWorkspace.Name
    txtQueue = QMaint.oQueue.Name
End Sub
