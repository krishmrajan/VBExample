VERSION 5.00
Begin VB.Form WSCopy 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Copy Workspace"
   ClientHeight    =   4545
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   8160
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   8160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frCopyOptions 
      Caption         =   "Copy Options"
      Height          =   1335
      Left            =   1800
      TabIndex        =   8
      Top             =   2280
      Width           =   4215
      Begin VB.OptionButton optCopyWSDefOnly 
         Caption         =   "Copy Workspace Definition Only"
         Height          =   255
         Left            =   360
         TabIndex        =   11
         Top             =   240
         Value           =   -1  'True
         Width           =   3615
      End
      Begin VB.OptionButton optCopyWSAndQDef 
         Caption         =   "Copy Workspace and Queue Definitions Only"
         Height          =   255
         Left            =   360
         TabIndex        =   10
         Top             =   600
         Width           =   3615
      End
      Begin VB.OptionButton optCopyWSAndQData 
         Caption         =   "Copy Workspace and Queue Data"
         Height          =   255
         Left            =   360
         TabIndex        =   9
         Top             =   960
         Width           =   3615
      End
   End
   Begin VB.Frame frSourceWSInfo 
      Caption         =   "Source Workspace Information"
      Height          =   855
      Left            =   240
      TabIndex        =   5
      Top             =   120
      Width           =   7695
      Begin VB.Label lblSourceWSDisp 
         Height          =   285
         Left            =   1920
         TabIndex        =   7
         Top             =   360
         Width           =   5535
      End
      Begin VB.Label lblSourceWS 
         Caption         =   "Source Workspace:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame frDestWSInfo 
      Caption         =   "Destination Workspace Information"
      Height          =   975
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   7695
      Begin VB.TextBox txtDestWS 
         Height          =   285
         Left            =   1920
         TabIndex        =   3
         Top             =   360
         Width           =   5655
      End
      Begin VB.Label lblDestWS 
         Caption         =   "Destination Workspace:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   4080
      TabIndex        =   1
      Top             =   3840
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   495
      Left            =   2400
      TabIndex        =   0
      Top             =   3840
      Width           =   1335
   End
End
Attribute VB_Name = "WSCopy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private oSourceWorkspace As IDMObjects.QueueWorkspace           ' queue workspace we are working on
Private oDestWorkspace As IDMObjects.QueueWorkspace             ' queue workspace we are working on
Private oQueueQuerySpec As IDMObjects.QueueQuerySpecification   ' The queue query spec object
Private oBrowseSet As IDMObjects.QueueBrowseSet                 ' The browse set
Private oQueueList As IDMObjects.ObjectSet
Private oSourceQueue As IDMObjects.queue
Private oDestQueue As IDMObjects.queue

Option Explicit

Private Sub Form_Load()
    
    Set oSourceWorkspace = QMaint.oLibrary.GetObject(idmObjTypeQueueWorkspace, _
                                                         QMaint.cmbWorkspace.Text)
    lblSourceWSDisp = oSourceWorkspace.Name
  
    cmdOK.Enabled = (txtDestWS <> "")
        
End Sub

Private Sub Form_Activate()
    WSCopy.txtDestWS.SetFocus
End Sub

Private Sub cmdOK_Click()
    
    On Error GoTo ErrorHandler
    
    Dim lMaxRet As Long
    Dim lCount As Long
    
    If optCopyWSDefOnly Then
        ' Copy the workspace definition only to a new workspace
        Screen.MousePointer = vbHourglass
        QMaint.MainStatusBar.SimpleText = "Copying Workspace Definition: " & txtDestWS & "..."
        QMaint.MainStatusBar.Refresh
   
        Set oSourceQueue = Nothing
        Set oDestQueue = Nothing

        ' Create a new workspace
        Set oDestWorkspace = QMaint.oLibrary.CreateObject(idmObjTypeQueueWorkspace, "")
        oDestWorkspace.Name = txtDestWS
        oDestWorkspace.Description = oSourceWorkspace.Description
        oDestWorkspace.Permissions.Item(1).GranteeName = oSourceWorkspace.Permissions.Item(1).GranteeName
        oDestWorkspace.Permissions.Item(2).GranteeName = oSourceWorkspace.Permissions.Item(2).GranteeName
        oDestWorkspace.Permissions.Item(3).GranteeName = oSourceWorkspace.Permissions.Item(3).GranteeName
        oDestWorkspace.SaveNew
    
    ElseIf optCopyWSAndQDef Then
        ' Copy the workspace and queue definition only to a new workspace
        Screen.MousePointer = vbHourglass
        QMaint.MainStatusBar.SimpleText = "Copying Workspace Definition: " & txtDestWS & "..."
        QMaint.MainStatusBar.Refresh
   
        Set oSourceQueue = Nothing
        Set oDestQueue = Nothing

        ' Create a new workspace
        Set oDestWorkspace = QMaint.oLibrary.CreateObject(idmObjTypeQueueWorkspace, "")
        oDestWorkspace.Name = txtDestWS
        oDestWorkspace.Description = oSourceWorkspace.Description
        oDestWorkspace.Permissions.Item(1).GranteeName = oSourceWorkspace.Permissions.Item(1).GranteeName
        oDestWorkspace.Permissions.Item(2).GranteeName = oSourceWorkspace.Permissions.Item(2).GranteeName
        oDestWorkspace.Permissions.Item(3).GranteeName = oSourceWorkspace.Permissions.Item(3).GranteeName
        oDestWorkspace.SaveNew

        Set oQueueList = oSourceWorkspace.FilterQueues("")
    
        For Each oSourceQueue In oQueueList
            QMaint.MainStatusBar.SimpleText = "Copying Queue Definition: " & oSourceQueue.Name & " ..."
            QMaint.MainStatusBar.Refresh
            Set oDestQueue = oDestWorkspace.CreateQueue(oSourceQueue.Name)
            oDestQueue.ServiceName = oSourceQueue.ServiceName               ' Set the queue service name
            Call gCopyQDefinition(oSourceQueue, oDestQueue)
            Set oSourceQueue = Nothing
            Set oDestQueue = Nothing
        Next
    
    ElseIf optCopyWSAndQData Then
        ' Copy the workspace and queue definition and data to a new workspace
        Screen.MousePointer = vbHourglass
        QMaint.MainStatusBar.SimpleText = "Copying Workspace Definition: " & txtDestWS & "..."
        QMaint.MainStatusBar.Refresh
   
        Set oSourceQueue = Nothing
        Set oDestQueue = Nothing

        ' Create a new workspace
        Set oDestWorkspace = QMaint.oLibrary.CreateObject(idmObjTypeQueueWorkspace, "")
        oDestWorkspace.Name = txtDestWS
        oDestWorkspace.Description = oSourceWorkspace.Description
        oDestWorkspace.Permissions.Item(1).GranteeName = oSourceWorkspace.Permissions.Item(1).GranteeName
        oDestWorkspace.Permissions.Item(2).GranteeName = oSourceWorkspace.Permissions.Item(2).GranteeName
        oDestWorkspace.Permissions.Item(3).GranteeName = oSourceWorkspace.Permissions.Item(3).GranteeName
        oDestWorkspace.SaveNew

        Set oQueueList = oSourceWorkspace.FilterQueues("")
    
        For Each oSourceQueue In oQueueList
            QMaint.MainStatusBar.SimpleText = "Copying Queue Definition: " & oSourceQueue.Name & " ..."
            QMaint.MainStatusBar.Refresh
            
            lMaxRet = 500
            Set oDestQueue = oDestWorkspace.CreateQueue(oSourceQueue.Name)
            oDestQueue.ServiceName = oSourceQueue.ServiceName               ' Set the queue service name
            
            Call gCopyQDefinition(oSourceQueue, oDestQueue)
            
            Set oQueueQuerySpec = oSourceQueue.CreateQuerySpecification()
            Call gBuildBrowseSet(oBrowseSet, oSourceQueue, oQueueQuerySpec, lMaxRet, lCount)
            Call gCopyQContents(oBrowseSet, oSourceQueue, oDestQueue, lCount)
            
            Set oSourceQueue = Nothing
            Set oDestQueue = Nothing
            Set oBrowseSet = Nothing
            Set oDestQueue = Nothing
        Next
    
    End If
    
    ' Clean up
    QMaint.cmbWorkspace.AddItem (oDestWorkspace.Name)
    QMaint.MainStatusBar.SimpleText = "Workspace Copied..."
    QMaint.MainStatusBar.Refresh
    Screen.MousePointer = vbDefault
    Set oDestWorkspace = Nothing
    Set oQueueList = Nothing
    Set oSourceQueue = Nothing
    Set oDestQueue = Nothing

    Unload Me

    Exit Sub
    
ErrorHandler:
    oErrorLog.logFNError errWarning, "Error copying workspace.", "Error copying workspace."
    Screen.MousePointer = vbDefault

End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub txtDestWS_Change()
    
    txtDestWS.Text = Trim(txtDestWS.Text)
    cmdOK.Enabled = (txtDestWS <> "")

End Sub
