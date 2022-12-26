VERSION 5.00
Begin VB.Form QCopy 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Copy Queue"
   ClientHeight    =   5430
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   7950
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5430
   ScaleWidth      =   7950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2040
      TabIndex        =   17
      Top             =   7320
      Width           =   5895
   End
   Begin VB.Frame frDestQInfo 
      Caption         =   "Destination Queue Information"
      Height          =   1455
      Left            =   120
      TabIndex        =   13
      Top             =   1680
      Width           =   7695
      Begin VB.TextBox txtDestServiceDisp 
         Height          =   285
         Left            =   1920
         TabIndex        =   20
         Top             =   360
         Width           =   5655
      End
      Begin VB.TextBox txtDestQ 
         Height          =   285
         Left            =   1920
         TabIndex        =   19
         Top             =   1080
         Width           =   5655
      End
      Begin VB.TextBox txtDestWS 
         Height          =   285
         Left            =   1920
         TabIndex        =   18
         Top             =   720
         Width           =   5655
      End
      Begin VB.Label lblDestQ 
         Caption         =   "Destination Queue:"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label lblDestWS 
         Caption         =   "Destination Workspace:"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label lblDestService 
         Caption         =   "Destination Service:"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Frame frSourceQInfo 
      Caption         =   "Source Queue Information"
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7695
      Begin VB.Label lblSourceService 
         Caption         =   "Source Service:"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label lblSourceWS 
         Caption         =   "Source Workspace:"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label lblSourceQ 
         Caption         =   "Source Queue:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label lblSourceServiceDisp 
         Height          =   285
         Left            =   1920
         TabIndex        =   9
         Top             =   360
         Width           =   5655
      End
      Begin VB.Label lblSourceWSDisp 
         Height          =   285
         Left            =   1920
         TabIndex        =   8
         Top             =   720
         Width           =   5655
      End
      Begin VB.Label lblSourceQDisp 
         Height          =   285
         Left            =   1920
         TabIndex        =   7
         Top             =   1080
         Width           =   5655
      End
   End
   Begin VB.Frame frCopyOptions 
      Caption         =   "Copy Options"
      Height          =   1335
      Left            =   2040
      TabIndex        =   6
      Top             =   3240
      Width           =   3855
      Begin VB.OptionButton optCopyQDataOnly 
         Caption         =   "Copy Queue Data Only"
         Height          =   255
         Left            =   360
         TabIndex        =   3
         Top             =   960
         Width           =   3255
      End
      Begin VB.OptionButton optCopyQDefAndData 
         Caption         =   "Copy Queue Definition and Data"
         Height          =   255
         Left            =   360
         TabIndex        =   2
         Top             =   600
         Width           =   3255
      End
      Begin VB.OptionButton optCopyQDefOnly 
         Caption         =   "Copy Queue Definition Only"
         Height          =   255
         Left            =   360
         TabIndex        =   1
         Top             =   240
         Value           =   -1  'True
         Width           =   3255
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   4320
      TabIndex        =   5
      Top             =   4800
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   495
      Left            =   2280
      TabIndex        =   4
      Top             =   4800
      Width           =   1335
   End
End
Attribute VB_Name = "QCopy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private oSourceQueue As IDMObjects.queue                        ' The source queue
Private oDestQueue As IDMObjects.queue                          ' The dest queue
Private oDestWorkspace As IDMObjects.QueueWorkspace             ' The dest queue
Private oBrowseSet As IDMObjects.QueueBrowseSet                 ' The browse set
Private oQueueQuerySpec As IDMObjects.QueueQuerySpecification   ' The queue query spec object
Private oSourceQueueEntry As IDMObjects.QueueEntry              ' The source QueueEntry to edit
Private oDestProperty As IDMObjects.Property                    ' The dest property

Option Explicit

Private Sub Form_Activate()
    QCopy.txtDestQ.SetFocus

End Sub

Private Sub Form_Load()
    
    Set oSourceQueue = QMaint.oQueue
    lblSourceServiceDisp = oSourceQueue.ServiceName
    lblSourceWSDisp = oSourceQueue.Workspace
    lblSourceQDisp = oSourceQueue.Name
    txtDestServiceDisp = oSourceQueue.ServiceName
    txtDestWS = oSourceQueue.Workspace
    cmdOK.Enabled = (txtDestWS <> "") And (txtDestQ <> "")

End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    
    Dim lMaxRet As Long                     ' The maximum number of queue entries to return
    Dim lCount As Long                      ' Approximate count of queue entries

    On Error GoTo ErrorHandler
    
    Screen.MousePointer = vbHourglass

    Set oDestWorkspace = QMaint.oLibrary.GetObject(idmObjTypeQueueWorkspace, txtDestWS.Text)
        
    If optCopyQDefOnly Then
        Set oDestQueue = oDestWorkspace.CreateQueue(txtDestQ)
        oDestQueue.ServiceName = txtDestServiceDisp               ' Set the queue service name
        If gCopyQDefinition(oSourceQueue, oDestQueue) Then
            'If new queue is in current workspace, add it to our list box
            If (lblSourceWSDisp = txtDestWS) Then
                QMaint.cmbQueue.AddItem (oDestQueue.Name)
            End If
        Else
            GoTo ErrorHandler
        End If
        
    ElseIf optCopyQDefAndData Then
        lMaxRet = 500
        Set oDestQueue = oDestWorkspace.CreateQueue(txtDestQ)       ' Create new queue
        oDestQueue.ServiceName = txtDestServiceDisp                 ' Set the queue service name
        
        Set oQueueQuerySpec = oSourceQueue.CreateQuerySpecification()
        Call gBuildBrowseSet(oBrowseSet, oSourceQueue, oQueueQuerySpec, lMaxRet, lCount)
        
        If gCopyQDefinition(oSourceQueue, oDestQueue) Then
            If gCopyQContents(oBrowseSet, oSourceQueue, oDestQueue, lCount) Then
                ' If new queue is in current workspace, add it to our list box
                If (lblSourceWSDisp = txtDestWS) Then
                    QMaint.cmbQueue.AddItem (oDestQueue.Name)
                End If
            Else
                GoTo ErrorHandler
            End If
        Else
            GoTo ErrorHandler
        End If
            
    ElseIf optCopyQDataOnly Then
        lMaxRet = 500
        Set oQueueQuerySpec = oSourceQueue.CreateQuerySpecification()
        Call gBuildBrowseSet(oBrowseSet, oSourceQueue, oQueueQuerySpec, lMaxRet, lCount)
        If lCount > 0 Then
            Set oDestQueue = oDestWorkspace.GetQueue(txtDestQ)
            oDestQueue.ServiceName = txtDestServiceDisp               ' Set the queue service name
            Call gCopyQContents(oBrowseSet, oSourceQueue, oDestQueue, lCount)
        Else
            MsgBox "No data to copy", vbInformation, AppName
        End If
    End If
    
    Set oDestQueue = Nothing
    Set oSourceQueue = Nothing
    Set oQueueQuerySpec = Nothing
    Set oDestWorkspace = Nothing
    
    QMaint.MainStatusBar.SimpleText = "Queue Copied..."
    QMaint.MainStatusBar.Refresh
    Screen.MousePointer = vbDefault
    Unload Me
    
    Exit Sub
    
ErrorHandler:
    Set oDestQueue = Nothing
    Set oSourceQueue = Nothing
    Set oQueueQuerySpec = Nothing
    Set oDestWorkspace = Nothing
    
    oErrorLog.logFNError errWarning, "Error copying queue.", "Error copying queue."
    Screen.MousePointer = vbDefault
        
End Sub
Private Sub txtDestWS_Change()

    txtDestWS.Text = Trim(txtDestWS.Text)
    cmdOK.Enabled = (txtDestWS <> "") And (txtDestQ <> "")
    
End Sub

Private Sub txtDestQ_Change()

    txtDestQ.Text = Trim(txtDestQ.Text)
    cmdOK.Enabled = (txtDestWS <> "") And (txtDestQ <> "")
    
End Sub

