VERSION 5.00
Begin VB.Form QCreate 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Create Queue"
   ClientHeight    =   7800
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   10950
   LinkTopic       =   "QCreate"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7800
   ScaleWidth      =   10950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.VScrollBar sbProperties 
      Height          =   1455
      Left            =   10560
      TabIndex        =   42
      Top             =   5640
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Frame frProperties 
      BorderStyle     =   0  'None
      Height          =   1695
      Left            =   240
      TabIndex        =   20
      Top             =   5280
      Width           =   10095
      Begin VB.Frame frProperty 
         Height          =   735
         Index           =   0
         Left            =   0
         TabIndex        =   21
         Top             =   240
         Width           =   10095
         Begin VB.TextBox txtPropScale 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   0
            Left            =   9480
            MaxLength       =   2
            TabIndex        =   31
            Text            =   "4"
            Top             =   240
            Width           =   375
         End
         Begin VB.TextBox txtPropPrecision 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   0
            Left            =   8640
            MaxLength       =   2
            TabIndex        =   30
            Text            =   "12"
            Top             =   240
            Width           =   495
         End
         Begin VB.OptionButton optPropRendezvous 
            Height          =   255
            Index           =   0
            Left            =   6720
            TabIndex        =   29
            Top             =   255
            Width           =   255
         End
         Begin VB.CheckBox cbPropDisplay 
            Height          =   255
            Index           =   0
            Left            =   7800
            TabIndex        =   28
            Top             =   255
            Value           =   1  'Checked
            Width           =   255
         End
         Begin VB.CheckBox cbPropRequired 
            Height          =   255
            Index           =   0
            Left            =   5880
            TabIndex        =   27
            Top             =   255
            Width           =   255
         End
         Begin VB.CheckBox cbPropUnique 
            Height          =   255
            Index           =   0
            Left            =   5040
            TabIndex        =   26
            Top             =   255
            Width           =   255
         End
         Begin VB.CheckBox cbPropKey 
            Height          =   255
            Index           =   0
            Left            =   4440
            TabIndex        =   25
            Top             =   255
            Width           =   255
         End
         Begin VB.TextBox txtPropSize 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   0
            Left            =   3600
            TabIndex        =   24
            Text            =   "20"
            Top             =   240
            Width           =   615
         End
         Begin VB.ComboBox cmbPropType 
            Height          =   315
            Index           =   0
            ItemData        =   "QCreate.frx":0000
            Left            =   2040
            List            =   "QCreate.frx":0024
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   23
            Top             =   225
            Width           =   1335
         End
         Begin VB.Label lblPropName 
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   22
            Top             =   255
            Width           =   1695
         End
      End
      Begin VB.Label lblScale 
         AutoSize        =   -1  'True
         Caption         =   "Scale"
         Height          =   195
         Left            =   9480
         TabIndex        =   41
         Top             =   0
         Width           =   405
      End
      Begin VB.Label lblPrecision 
         AutoSize        =   -1  'True
         Caption         =   "Precision"
         Height          =   195
         Left            =   8520
         TabIndex        =   40
         Top             =   0
         Width           =   645
      End
      Begin VB.Label lblDisplay 
         AutoSize        =   -1  'True
         Caption         =   "Display"
         Height          =   195
         Left            =   7680
         TabIndex        =   39
         Top             =   0
         Width           =   510
      End
      Begin VB.Label lblRendezvous 
         AutoSize        =   -1  'True
         Caption         =   "Rendezvous"
         Height          =   195
         Left            =   6480
         TabIndex        =   38
         Top             =   0
         Width           =   900
      End
      Begin VB.Label lblRequired 
         AutoSize        =   -1  'True
         Caption         =   "Required"
         Height          =   195
         Left            =   5640
         TabIndex        =   37
         Top             =   0
         Width           =   645
      End
      Begin VB.Label lblUnique 
         AutoSize        =   -1  'True
         Caption         =   "Unique"
         Height          =   195
         Left            =   4920
         TabIndex        =   36
         Top             =   0
         Width           =   510
      End
      Begin VB.Label lblKey 
         AutoSize        =   -1  'True
         Caption         =   "Key"
         Height          =   195
         Left            =   4440
         TabIndex        =   35
         Top             =   0
         Width           =   270
      End
      Begin VB.Label lblSize 
         AutoSize        =   -1  'True
         Caption         =   "Size"
         Height          =   195
         Left            =   3720
         TabIndex        =   34
         Top             =   0
         Width           =   300
      End
      Begin VB.Label lblType 
         AutoSize        =   -1  'True
         Caption         =   "Data Type"
         Height          =   195
         Left            =   2280
         TabIndex        =   33
         Top             =   0
         Width           =   750
      End
      Begin VB.Label lblPropNameName 
         AutoSize        =   -1  'True
         Caption         =   "Name"
         Height          =   195
         Left            =   840
         TabIndex        =   32
         Top             =   0
         Width           =   420
      End
   End
   Begin VB.CommandButton btnCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   6120
      TabIndex        =   19
      Top             =   7200
      Width           =   1815
   End
   Begin VB.CommandButton btnSave 
      BackColor       =   &H00FF8080&
      Caption         =   "&Save"
      Default         =   -1  'True
      Height          =   375
      Left            =   2040
      MaskColor       =   &H0000C000&
      TabIndex        =   18
      Top             =   7200
      Width           =   1815
   End
   Begin VB.Frame frDefPermissions 
      Caption         =   "Definition Permissions"
      Height          =   1935
      Left            =   240
      TabIndex        =   4
      Top             =   1200
      Width           =   10095
      Begin VB.TextBox txtDefRead 
         Height          =   285
         Left            =   960
         MaxLength       =   82
         TabIndex        =   6
         Top             =   360
         Width           =   8895
      End
      Begin VB.TextBox txtDefWrite 
         Height          =   285
         Left            =   960
         MaxLength       =   82
         TabIndex        =   8
         Top             =   840
         Width           =   8895
      End
      Begin VB.TextBox txtDefAX 
         Height          =   285
         Left            =   960
         MaxLength       =   82
         TabIndex        =   10
         Top             =   1320
         Width           =   8895
      End
      Begin VB.Label lblDefRead 
         Caption         =   "Read"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblDefWrite 
         Caption         =   "Write"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   840
         Width           =   735
      End
      Begin VB.Label lblDefAX 
         Caption         =   "AX"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   1320
         Width           =   735
      End
   End
   Begin VB.Frame frContPermissions 
      Caption         =   "Content Permissions"
      Height          =   1935
      Left            =   240
      TabIndex        =   11
      Top             =   3240
      Width           =   10095
      Begin VB.TextBox txtContRead 
         Height          =   285
         Left            =   960
         MaxLength       =   82
         TabIndex        =   13
         Top             =   360
         Width           =   8895
      End
      Begin VB.TextBox txtContWrite 
         Height          =   285
         Left            =   960
         MaxLength       =   82
         TabIndex        =   15
         Top             =   840
         Width           =   8895
      End
      Begin VB.TextBox txtContAX 
         Height          =   285
         Left            =   960
         MaxLength       =   82
         TabIndex        =   17
         Top             =   1320
         Width           =   8895
      End
      Begin VB.Label lblContRead 
         Caption         =   "Read"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblContWrite 
         Caption         =   "Write"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   840
         Width           =   735
      End
      Begin VB.Label lblContAX 
         Caption         =   "AX"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   1320
         Width           =   735
      End
   End
   Begin VB.TextBox txtDescription 
      Height          =   495
      Left            =   1320
      MaxLength       =   800
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   840
      Width           =   9015
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   1320
      MaxLength       =   14
      TabIndex        =   1
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label lblDescription 
      Caption         =   "Description:"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   975
   End
   Begin VB.Label lblName 
      Caption         =   "Name:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   495
   End
   Begin VB.Menu mnuAddProp 
      Caption         =   "&Add Property Description..."
   End
   Begin VB.Menu mnuRemoveProp 
      Caption         =   "&Remove Property Description..."
   End
   Begin VB.Menu mnuRemRendez 
      Caption         =   "Remove Rendez&vous"
   End
   Begin VB.Menu mnuServices 
      Caption         =   "&Services"
   End
End
Attribute VB_Name = "QCreate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public isNew As Boolean                ' If is a new queue
Private RecreateQueue As Boolean       ' Flags need to delete and recreate
                                       ' a queue for deleting properties.
Private oldQueueName As String
Private theQueue As IDMObjects.queue   ' The queue that we are referencing
Private numPropDescs As Integer        ' The number of property descriptions
Private maxFieldNum As Integer         ' The maximum field number assigned
Private Const visibleProps = 2         ' The number of visible property descriptions on form

Private Sub btnCancel_Click()
    If isNew Then
        QMaint.MainStatusBar.SimpleText = "Queue creation cancelled."
    Else
        QMaint.MainStatusBar.SimpleText = "Queue modification cancelled."
    End If
    QMaint.MainStatusBar.Refresh
    
    Unload Me
End Sub

Private Sub btnSave_Click()
    Dim result As Boolean
    Dim inx As Integer
    Dim propDesc As IDMObjects.PropertyDescription
    Dim tempQueue As IDMObjects.queue     ' Store an origonal copy of the queue.
            
On Error GoTo ErrorHandler

    
    ' If any security fields changed for an existing queue, we need to recreate the queue.
    If Not (isNew) And _
      (txtDefRead <> theQueue.DefinitionPermissions(idmISAccessRead).GranteeName _
      Or txtDefWrite <> theQueue.DefinitionPermissions(idmISAccessWrite).GranteeName _
      Or txtDefAX <> theQueue.DefinitionPermissions(idmISAccessAX).GranteeName _
      Or txtContRead <> theQueue.ContentPermissions(idmISAccessRead).GranteeName _
      Or txtContWrite <> theQueue.ContentPermissions(idmISAccessWrite).GranteeName _
      Or txtContAX <> theQueue.ContentPermissions(idmISAccessAX).GranteeName) Then
           RecreateQueue = True
    End If
    
    ' Always confirm the deletion of data
    If RecreateQueue Then
        If MsgBox("This action will result in the deletion of all queue entries.  Do you wish to continue?" _
                  , vbYesNo + vbExclamation + vbDefaultButton2, AppName) = vbNo Then
            Exit Sub
        End If
    End If

    Screen.MousePointer = vbHourglass
    QMaint.MainStatusBar.SimpleText = "Saving Queue..."
    QMaint.MainStatusBar.Refresh
    
    ' Load Queue definations from the form
    oldQueueName = theQueue.Name
    theQueue.Name = txtName
    theQueue.Description = txtDescription
    theQueue.DefinitionPermissions(idmISAccessRead).GranteeName = txtDefRead
    theQueue.DefinitionPermissions(idmISAccessWrite).GranteeName = txtDefWrite
    theQueue.DefinitionPermissions(idmISAccessAX).GranteeName = txtDefAX
    theQueue.ContentPermissions(idmISAccessRead).GranteeName = txtContRead
    theQueue.ContentPermissions(idmISAccessWrite).GranteeName = txtContWrite
    theQueue.ContentPermissions(idmISAccessAX).GranteeName = txtContAX
    
    inx = 0
    For Each propDesc In theQueue.PropertyDescriptions
        If propDesc.GetState(idmPropCustom) Then
            propDesc.GetExtendedProperty("F_QUEUETYPEID") = cmbPropType(inx).ItemData(cmbPropType(inx).ListIndex)
            propDesc.GetExtendedProperty("F_FIELDSIZE") = CInt(txtPropSize(inx))
            propDesc.GetExtendedProperty("F_ISKEY") = fromCheckbox(cbPropKey(inx))
            propDesc.GetExtendedProperty("F_ISUNIQUE") = fromCheckbox(cbPropUnique(inx))
            propDesc.GetExtendedProperty("F_ISREQUIRED") = fromCheckbox(cbPropRequired(inx))
            propDesc.GetExtendedProperty("F_ISRENDEZVOUS") = optPropRendezvous(inx)
            propDesc.GetExtendedProperty("F_SHOULDDISPLAY") = fromCheckbox(cbPropDisplay(inx))
            propDesc.GetExtendedProperty("F_PRECISION") = CInt(txtPropPrecision(inx))
            propDesc.GetExtendedProperty("F_SCALE") = CInt(txtPropScale(inx))
            inx = inx + 1
        End If
    Next
    
    If isNew Then
        theQueue.SaveNew
        QMaint.cmbQueue.AddItem (txtName)
    Else
        If RecreateQueue Then
            Set tempQueue = QMaint.oWorkspace.CreateQueue("temp" & oldQueueName)
            
            ' Copy the Queue Defination and save it as a new temp queue
            result = gCopyQDefinition(theQueue, tempQueue)

            ' Release all links to the queue object
            Set QMaint.oQueue = Nothing
            Set QMaint.oQueueQuerySpec = Nothing
            Set QMaint.cQueueEntries = Nothing
            Set theQueue = Nothing
            Set tempQueue = Nothing
            Set propDesc = Nothing
            
            ' Clear the grid
            QMaint.grdQueueData.Clear
            QMaint.grdQueueData.Rows = 0
            
            ' Get original version of the queue
            Set theQueue = QMaint.oWorkspace.GetQueue(oldQueueName)
            
            DoEvents
            ' Wait for the server...
            Sleep (2000)
            ' Delete the original queue
            theQueue.Delete True
            
            ' Let go of the ref to the deleted queue.
            Set theQueue = Nothing
            
            ' Get the temp queue and change the name back to the original name.
            Set theQueue = QMaint.oWorkspace.GetQueue("temp" & oldQueueName)
            theQueue.Name = oldQueueName
            DoEvents
            theQueue.Save
                
        Else
            ' Release links to the queue object in qmaint.frm.
            Set QMaint.oQueue = Nothing
            Set QMaint.oQueueQuerySpec = Nothing
            DoEvents
            ' Save the new queue
            theQueue.Save
        End If
    
        ' The name of the queue has changed
        If txtName <> oldQueueName Then
            ' Handle renames of the queue.  Update combo box on QMaint form.
            Dim lstCount As Integer
            Dim listInx As Integer
            
            ' If existing queue, remove old name from combo box
            If Not isNew Then
                lstCount = QMaint.cmbQueue.listCount
                For listInx = 0 To lstCount - 1
                    If QMaint.cmbQueue.List(listInx) = oldQueueName Then
                        QMaint.cmbQueue.RemoveItem (listInx)
                        Exit For
                    End If
                Next listInx
            End If
            
            ' Add new name to combo box
            QMaint.cmbQueue.AddItem (txtName)
            QMaint.cmbQueue.Text = txtName
        End If
    End If
    
    QMaint.MainStatusBar.SimpleText = "Queue saved."
    QMaint.MainStatusBar.Refresh
    Screen.MousePointer = vbDefault
    Unload Me
    RecreateQueue = False
    
    ' Update the grid now...
    Call QMaint.cmbQueue_Click
            
    ' Release all local links to the queue object
    Set theQueue = Nothing
    Set propDesc = Nothing
    Set tempQueue = Nothing
    
    Exit Sub
ErrorHandler:
    oErrorLog.logFNError errWarning, "Error saving queue.", "Error saving queue."
    Call QMaint.cmbQueue_Click   ' Restore original settings...
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub Form_Load()
    Dim propDesc As IDMObjects.PropertyDescription
    
    If isNew Then
        Me.Caption = "Create Queue"
        Set theQueue = QMaint.oWorkspace.CreateQueue("Queue1")
    Else
        Me.Caption = "Modify Queue"
        Set theQueue = QMaint.oQueue
    End If
    txtName = theQueue.Name
    txtDescription = theQueue.Description
    txtDefRead = theQueue.DefinitionPermissions(idmISAccessRead).GranteeName
    txtDefWrite = theQueue.DefinitionPermissions(idmISAccessWrite).GranteeName
    txtDefAX = theQueue.DefinitionPermissions(idmISAccessAX).GranteeName
    txtContRead = theQueue.ContentPermissions(idmISAccessRead).GranteeName
    txtContWrite = theQueue.ContentPermissions(idmISAccessWrite).GranteeName
    txtContAX = theQueue.ContentPermissions(idmISAccessAX).GranteeName
    numPropDescs = 0
    For Each propDesc In theQueue.PropertyDescriptions
        If propDesc.GetState(idmPropCustom) Then
            If numPropDescs > frProperty.UBound Then
                Call loadControls(numPropDescs)
            End If
            numPropDescs = numPropDescs + 1
            Call displayPropDesc(propDesc, numPropDescs - 1)
        End If
    Next
    maxFieldNum = numPropDescs
    If numPropDescs = 0 Then
        Call setVisible(0, False)
        btnSave.Enabled = False
    Else
        btnSave.Enabled = (txtName <> "")
    End If
    RecreateQueue = False
    
    Set propDesc = Nothing
    
End Sub

Private Sub mnuAddProp_Click()
    Dim nameStr As String
    Dim propDesc As IDMObjects.PropertyDescription
    
    On Error GoTo ErrorHandler
    maxFieldNum = maxFieldNum + 1
    nameStr = InputBox("Property Name", "Add Property Description", "Field" & CStr(maxFieldNum))
    If nameStr = "" Then
        Exit Sub
    End If
    Set propDesc = theQueue.AddPropertyDescription(nameStr)
    If numPropDescs > frProperty.UBound Then
        Call loadControls(numPropDescs)
    End If
    If numPropDescs >= visibleProps Then
        sbProperties.Max = numPropDescs - visibleProps + 1
        sbProperties.Value = sbProperties.Max
        sbProperties.Visible = True
    End If
    numPropDescs = numPropDescs + 1
    Call displayPropDesc(propDesc, numPropDescs - 1)
    Set propDesc = Nothing
    btnSave.Enabled = (txtName <> "")
    Exit Sub
ErrorHandler:
    oErrorLog.logFNError errWarning, "Error adding property description.", "Error adding property description."
End Sub

Private Sub displayPropDesc(propDesc As IDMObjects.PropertyDescription, inx As Integer)
    Dim listInx As Integer
    Dim pos As Integer
    
    lblPropName(inx) = propDesc.Name
    For listInx = 0 To cmbPropType(inx).listCount - 1
        If cmbPropType(inx).ItemData(listInx) = propDesc.GetExtendedProperty("F_QUEUETYPEID") Then
            cmbPropType(inx).ListIndex = listInx
            Exit For
        End If
    Next listInx
    txtPropSize(inx) = propDesc.Size
    cbPropKey(inx) = toCheckbox(propDesc.GetState(idmPropKey))
    cbPropUnique(inx) = toCheckbox(propDesc.GetState(idmPropUnique))
    cbPropRequired(inx) = toCheckbox(propDesc.GetState(idmPropRequired))
    optPropRendezvous(inx) = propDesc.GetState(idmPropRendezvous)
    cbPropDisplay(inx) = toCheckbox(propDesc.GetState(idmPropShouldDisplay))
    txtPropPrecision(inx) = propDesc.GetExtendedProperty("F_PRECISION")
    txtPropScale(inx) = propDesc.GetExtendedProperty("F_SCALE")
    
    pos = sbProperties.Value
    If inx >= pos And inx < pos + visibleProps Then
        frProperty(inx).Top = frProperty(0).Top + _
                              frProperty(0).Height * (inx - pos)
        Call setVisible(inx, True)
    Else
        Call setVisible(inx, False)
    End If
End Sub

Private Function toCheckbox(val As Boolean) As Integer
    If val Then
        toCheckbox = 1
    Else
        toCheckbox = 0
    End If
End Function

Private Function fromCheckbox(val As Integer) As Boolean
    If val = 1 Then
        fromCheckbox = True
    Else
        fromCheckbox = False
    End If
End Function

Private Sub mnuRemoveProp_Click()
    Dim nameStr As String
    
    On Error GoTo ErrorHandler
    nameStr = InputBox("Property Name", "Remove Property Description", "")
    If nameStr = "" Then
        Exit Sub
    End If
    theQueue.RemovePropertyDescription (nameStr)
    Call removePropDesc(nameStr)
    If numPropDescs <= visibleProps Then
        sbProperties.Visible = False
        sbProperties.Value = 0
    Else
        If sbProperties.Value >= sbProperties.Max Then
            sbProperties.Value = sbProperties.Max - 1
        End If
        sbProperties.Max = sbProperties.Max - 1
    End If
    If numPropDescs = 0 Then
        btnSave.Enabled = False
    End If
    RecreateQueue = True
    Exit Sub
ErrorHandler:
    oErrorLog.logFNError errWarning, "Error removing property description.", "Error removing property description."
End Sub

Private Sub removePropDesc(nameStr As String)
    Dim inx As Integer
    Dim inx2 As Integer
    
    For inx = 0 To numPropDescs - 1
        If lblPropName(inx) = nameStr Then
            Exit For
        End If
    Next inx
    For inx2 = inx + 1 To numPropDescs - 1
        lblPropName(inx2 - 1) = lblPropName(inx2)
        cmbPropType(inx2 - 1) = cmbPropType(inx2)
        txtPropSize(inx2 - 1) = txtPropSize(inx2)
        cbPropKey(inx2 - 1) = cbPropKey(inx2)
        cbPropUnique(inx2 - 1) = cbPropUnique(inx2)
        cbPropRequired(inx2 - 1) = cbPropRequired(inx2)
        optPropRendezvous(inx2 - 1) = optPropRendezvous(inx2)
        cbPropDisplay(inx2 - 1) = cbPropDisplay(inx2)
        txtPropPrecision(inx2 - 1) = txtPropPrecision(inx2)
        txtPropScale(inx2 - 1) = txtPropScale(inx2)
    Next inx2
    Call setVisible(numPropDescs - 1, False)
    numPropDescs = numPropDescs - 1
End Sub

Private Sub setVisible(inx As Integer, visibility As Boolean)
    frProperty(inx).Visible = visibility
    lblPropName(inx).Visible = visibility
    cmbPropType(inx).Visible = visibility
    txtPropSize(inx).Visible = visibility
    cbPropKey(inx).Visible = visibility
    cbPropUnique(inx).Visible = visibility
    cbPropRequired(inx).Visible = visibility
    optPropRendezvous(inx).Visible = visibility
    cbPropDisplay(inx).Visible = visibility
    txtPropPrecision(inx).Visible = visibility
    txtPropScale(inx).Visible = visibility
End Sub

Private Sub loadControls(inx As Integer)
    Dim inx2 As Integer
    Dim SaveoptPropRendezvous As Boolean
    
    Load frProperty(inx)
    Load lblPropName(inx)
    Load cmbPropType(inx)
    For inx2 = 0 To cmbPropType(0).listCount - 1
        cmbPropType(inx).List(inx2) = cmbPropType(0).List(inx2)
        cmbPropType(inx).ItemData(inx2) = cmbPropType(0).ItemData(inx2)
    Next inx2
    Load txtPropSize(inx)
    Load cbPropKey(inx)
    Load cbPropUnique(inx)
    Load cbPropRequired(inx)
    SaveoptPropRendezvous = optPropRendezvous(0)
    optPropRendezvous(0) = False       'Must be done cause if first control is true, the new control
    Load optPropRendezvous(inx)        'will be set to True and change the first control to false.
    optPropRendezvous(0) = SaveoptPropRendezvous
    Load cbPropDisplay(inx)
    Load txtPropPrecision(inx)
    Load txtPropScale(inx)
    Set lblPropName(inx).Container = frProperty(inx)
    Set cmbPropType(inx).Container = frProperty(inx)
    Set txtPropSize(inx).Container = frProperty(inx)
    Set cbPropKey(inx).Container = frProperty(inx)
    Set cbPropUnique(inx).Container = frProperty(inx)
    Set cbPropRequired(inx).Container = frProperty(inx)
    Set optPropRendezvous(inx).Container = frProperty(inx)
    Set cbPropDisplay(inx).Container = frProperty(inx)
    Set txtPropPrecision(inx).Container = frProperty(inx)
    Set txtPropScale(inx).Container = frProperty(inx)
    If inx >= visibleProps Then
        sbProperties.Max = inx - visibleProps + 1
        sbProperties.Visible = True
    End If
End Sub

Private Sub mnuRemRendez_Click()
    Dim inx As Integer
    
    For inx = 0 To numPropDescs - 1
        optPropRendezvous(inx).Value = False
    Next inx
End Sub

Private Sub mnuServices_Click()
    Dim ServiceStr As String
    
    On Error GoTo ErrorHandler
    ServiceStr = InputBox("Change default Queue Service", AppName, "")
    If ServiceStr = "" Then
        Exit Sub
    End If
    theQueue.ServiceName = ServiceStr
    
    Exit Sub
ErrorHandler:
    oErrorLog.logFNError errWarning, "Error Setting Service Name.", "Error Setting Service Name."

End Sub

Private Sub optPropRendezvous_Click(Index As Integer)
    Dim inx As Integer
    
    For inx = 0 To numPropDescs - 1
        If Index <> inx Then
            optPropRendezvous(inx).Value = False
        End If
    Next inx
End Sub

Private Sub sbProperties_Change()
    Dim inx As Integer
    Dim pos As Integer
    
    pos = sbProperties.Value
    For inx = 0 To numPropDescs - 1
        If inx >= pos And inx < pos + visibleProps Then
            frProperty(inx).Top = frProperty(0).Top + _
                                  frProperty(0).Height * (inx - pos)
            Call setVisible(inx, True)
        Else
            Call setVisible(inx, False)
        End If
    Next inx
End Sub

Private Sub txtName_Change()
    txtName.Text = Trim(txtName.Text)
    btnSave.Enabled = (txtName.Text <> "" And numPropDescs > 0)
End Sub
