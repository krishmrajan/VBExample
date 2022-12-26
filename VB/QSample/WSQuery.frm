VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form WSQuery 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Workspace Query"
   ClientHeight    =   6585
   ClientLeft      =   3375
   ClientTop       =   825
   ClientWidth     =   11715
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6585
   ScaleWidth      =   11715
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComCtl2.UpDown udMaxPrio 
      Height          =   255
      Left            =   4680
      TabIndex        =   48
      Top             =   1095
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   450
      _Version        =   393216
      BuddyControl    =   "meMaxPrio"
      BuddyDispid     =   196642
      OrigLeft        =   4680
      OrigTop         =   1095
      OrigRight       =   4920
      OrigBottom      =   1350
      SyncBuddy       =   -1  'True
      BuddyProperty   =   22
      Enabled         =   -1  'True
   End
   Begin VB.Frame frSearchFields 
      Caption         =   "Search Fields"
      Height          =   5415
      Left            =   5715
      TabIndex        =   33
      Top             =   270
      Width           =   5940
      Begin VB.Frame frSearchField 
         Caption         =   "Search Field #3"
         Height          =   1575
         Index           =   2
         Left            =   135
         TabIndex        =   43
         Top             =   3735
         Width           =   5655
         Begin VB.TextBox txtFieldName 
            Height          =   315
            Index           =   2
            Left            =   135
            TabIndex        =   21
            Top             =   495
            Width           =   3015
         End
         Begin VB.TextBox txtFieldValue 
            Height          =   315
            Index           =   2
            Left            =   135
            TabIndex        =   23
            Top             =   1170
            Width           =   5415
         End
         Begin VB.ComboBox cmbPropType 
            Height          =   315
            Index           =   2
            ItemData        =   "WSQuery.frx":0000
            Left            =   3375
            List            =   "WSQuery.frx":0002
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   22
            Top             =   495
            Width           =   2175
         End
         Begin VB.Label lbFieldValue 
            Caption         =   "Field Value:"
            Height          =   195
            Index           =   2
            Left            =   135
            TabIndex        =   46
            Top             =   945
            Width           =   1725
         End
         Begin VB.Label lbFieldName 
            Caption         =   "Field Name:"
            Height          =   195
            Index           =   2
            Left            =   135
            TabIndex        =   45
            Top             =   270
            Width           =   1725
         End
         Begin VB.Label lbFieldType 
            Caption         =   "Field Type:"
            Height          =   195
            Index           =   2
            Left            =   3330
            TabIndex        =   44
            Top             =   270
            Width           =   1725
         End
      End
      Begin VB.Frame frSearchField 
         Caption         =   "Search Field #2"
         Height          =   1575
         Index           =   1
         Left            =   135
         TabIndex        =   39
         Top             =   2025
         Width           =   5655
         Begin VB.TextBox txtFieldName 
            Height          =   315
            Index           =   1
            Left            =   135
            TabIndex        =   18
            Top             =   495
            Width           =   3015
         End
         Begin VB.TextBox txtFieldValue 
            Height          =   315
            Index           =   1
            Left            =   135
            TabIndex        =   20
            Top             =   1170
            Width           =   5415
         End
         Begin VB.ComboBox cmbPropType 
            Height          =   315
            Index           =   1
            ItemData        =   "WSQuery.frx":0004
            Left            =   3375
            List            =   "WSQuery.frx":0006
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   495
            Width           =   2175
         End
         Begin VB.Label lbFieldValue 
            Caption         =   "Field Value:"
            Height          =   195
            Index           =   1
            Left            =   135
            TabIndex        =   42
            Top             =   945
            Width           =   1725
         End
         Begin VB.Label lbFieldName 
            Caption         =   "Field Name:"
            Height          =   195
            Index           =   1
            Left            =   135
            TabIndex        =   41
            Top             =   270
            Width           =   1725
         End
         Begin VB.Label lbFieldType 
            Caption         =   "Field Type:"
            Height          =   195
            Index           =   1
            Left            =   3330
            TabIndex        =   40
            Top             =   270
            Width           =   1725
         End
      End
      Begin VB.Frame frSearchField 
         Caption         =   "Search Field #1"
         Height          =   1575
         Index           =   0
         Left            =   135
         TabIndex        =   34
         Top             =   315
         Width           =   5655
         Begin VB.ComboBox cmbPropType 
            Height          =   315
            Index           =   0
            ItemData        =   "WSQuery.frx":0008
            Left            =   3375
            List            =   "WSQuery.frx":0029
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   495
            Width           =   2175
         End
         Begin VB.TextBox txtFieldValue 
            Height          =   315
            Index           =   0
            Left            =   135
            TabIndex        =   17
            Top             =   1170
            Width           =   5415
         End
         Begin VB.TextBox txtFieldName 
            Height          =   315
            Index           =   0
            Left            =   135
            TabIndex        =   15
            Top             =   495
            Width           =   3015
         End
         Begin VB.Label lbFieldType 
            Caption         =   "Field Type:"
            Height          =   195
            Index           =   0
            Left            =   3330
            TabIndex        =   38
            Top             =   270
            Width           =   1725
         End
         Begin VB.Label lbFieldName 
            Caption         =   "Field Name:"
            Height          =   195
            Index           =   0
            Left            =   135
            TabIndex        =   37
            Top             =   270
            Width           =   1725
         End
         Begin VB.Label lbFieldValue 
            Caption         =   "Field Value:"
            Height          =   195
            Index           =   0
            Left            =   135
            TabIndex        =   36
            Top             =   945
            Width           =   1725
         End
      End
   End
   Begin VB.Frame frBusy 
      Caption         =   "Busy"
      Height          =   615
      Left            =   360
      TabIndex        =   35
      Top             =   255
      Width           =   5055
      Begin VB.OptionButton optNotBusy 
         Alignment       =   1  'Right Justify
         Caption         =   "Not"
         Height          =   255
         Left            =   3840
         TabIndex        =   3
         Top             =   240
         Width           =   615
      End
      Begin VB.OptionButton optOnlyIfBusy 
         Alignment       =   1  'Right Justify
         Caption         =   "Only If"
         Height          =   255
         Left            =   1800
         TabIndex        =   2
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton optEvenIfBusy 
         Alignment       =   1  'Right Justify
         Caption         =   "Even If"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   855
      End
   End
   Begin MSMask.MaskEdBox meMaxPrio 
      Height          =   255
      Left            =   4440
      TabIndex        =   4
      Top             =   1095
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      _Version        =   393216
      MaxLength       =   1
      Mask            =   "#"
      PromptChar      =   "_"
   End
   Begin VB.CheckBox cbSortDesc 
      Alignment       =   1  'Right Justify
      Caption         =   "Sort Descending"
      Height          =   330
      Left            =   360
      TabIndex        =   14
      Top             =   5295
      Width           =   1650
   End
   Begin VB.TextBox txtSort 
      Height          =   285
      Left            =   1830
      MaxLength       =   18
      TabIndex        =   13
      Top             =   4815
      Width           =   3585
   End
   Begin VB.TextBox txtUser 
      Height          =   285
      Left            =   1830
      MaxLength       =   82
      TabIndex        =   6
      Top             =   2055
      Width           =   3585
   End
   Begin VB.Frame frIncomplete 
      Caption         =   "Incomplete"
      Height          =   615
      Left            =   360
      TabIndex        =   31
      Top             =   3975
      Width           =   5055
      Begin VB.OptionButton optNotInc 
         Alignment       =   1  'Right Justify
         Caption         =   "Not"
         Height          =   255
         Left            =   3840
         TabIndex        =   12
         Top             =   240
         Width           =   615
      End
      Begin VB.OptionButton optOnlyIfInc 
         Alignment       =   1  'Right Justify
         Caption         =   "Only If"
         Height          =   255
         Left            =   1800
         TabIndex        =   11
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton optEvenIfInc 
         Alignment       =   1  'Right Justify
         Caption         =   "Even If"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.TextBox txtGroup 
      Height          =   285
      Left            =   1830
      MaxLength       =   82
      TabIndex        =   7
      Top             =   2535
      Width           =   3585
   End
   Begin VB.CheckBox cbDelayed 
      Alignment       =   1  'Right Justify
      Caption         =   "Even If Delayed"
      Height          =   330
      Left            =   360
      TabIndex        =   8
      Top             =   3015
      Width           =   1650
   End
   Begin VB.TextBox txtDeadline 
      Height          =   285
      Left            =   1830
      TabIndex        =   9
      Top             =   3495
      Width           =   1575
   End
   Begin VB.CheckBox cbCheckUser 
      Alignment       =   1  'Right Justify
      Caption         =   "Check User"
      Height          =   330
      Left            =   360
      TabIndex        =   5
      Top             =   1575
      Width           =   1650
   End
   Begin VB.CommandButton ClearButton 
      Caption         =   "Clear"
      Height          =   375
      Left            =   8010
      TabIndex        =   26
      Top             =   6015
      Width           =   1455
   End
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5130
      TabIndex        =   25
      Top             =   6015
      Width           =   1455
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2250
      TabIndex        =   24
      Top             =   6000
      Width           =   1455
   End
   Begin MSMask.MaskEdBox meMinPrio 
      Height          =   255
      Left            =   1800
      TabIndex        =   49
      Top             =   1095
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      _Version        =   393216
      MaxLength       =   1
      Mask            =   "#"
      PromptChar      =   "_"
   End
   Begin MSComCtl2.UpDown udMinPrio 
      Height          =   255
      Left            =   2040
      TabIndex        =   47
      Top             =   1095
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   450
      _Version        =   393216
      BuddyControl    =   "meMinPrio"
      BuddyDispid     =   196643
      OrigLeft        =   2040
      OrigTop         =   1080
      OrigRight       =   2280
      OrigBottom      =   1335
      SyncBuddy       =   -1  'True
      BuddyProperty   =   22
      Enabled         =   -1  'True
   End
   Begin VB.Label lbMinPrio 
      Caption         =   "Min. Priority"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   1095
      Width           =   975
   End
   Begin VB.Label lbSort 
      Caption         =   "Sort Field"
      Height          =   255
      Left            =   360
      TabIndex        =   32
      Top             =   4815
      Width           =   975
   End
   Begin VB.Label lbUser 
      Caption         =   "User Name"
      Height          =   255
      Left            =   360
      TabIndex        =   28
      Top             =   2055
      Width           =   1095
   End
   Begin VB.Label lbMaxPrio 
      Caption         =   "Max. Priority"
      Height          =   255
      Left            =   3240
      TabIndex        =   27
      Top             =   1095
      Width           =   975
   End
   Begin VB.Label lbGroup 
      Caption         =   "Group Name"
      Height          =   255
      Left            =   360
      TabIndex        =   29
      Top             =   2535
      Width           =   1095
   End
   Begin VB.Label lbDeadline 
      Alignment       =   2  'Center
      Caption         =   "Deadline"
      Height          =   255
      Left            =   360
      TabIndex        =   30
      Top             =   3495
      Width           =   735
   End
End
Attribute VB_Name = "WSQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public oWorkspace As IDMObjects.QueueWorkspace  ' queue workspace we are working on
Public oQueueList As IDMObjects.ObjectSet
Public oQueue As IDMObjects.queue               ' queue we are working on
Public oQueueQuerySpec As IDMObjects.QueueQuerySpecification   ' the queue query spec object
Public cProperties As Collection                    ' the collection of properties for this QueueEntry
Public iTotalFields As Integer                      ' the total number of fields displayed on the form
Private lCount As Long                              ' Entries returned from browse set
Private Const iMaxFields = 3                        ' Max fields to search
Private Const iField1 = 0
Private Const iField2 = 1
Private Const iField3 = 2
Private Const lMaxRet = 500                         ' Max entries for building browse set
Private Const visibleProps = 20                     ' the number of visible properties on form
Private Const iSpacing = 120                        ' the vertical spacing between the datafields

' Cancel out of edit
Private Sub CancelButton_Click()
    QMaint.MainStatusBar.SimpleText = "Query Workspace cancelled at user's request."
    
    QMaint.MainStatusBar.Refresh
    Unload Me
End Sub


Private Sub Form_Load()

    Dim oProperties As IDMObjects.Properties        ' a collection of properties for a QueueEntry
    Dim oProperty As IDMObjects.Property            ' a property of a QueueEntry
    Dim iCounter As Integer                         ' an integer counter
    Dim iHeight As Integer                          ' the height of the standard DataField field
    Dim sQueueName As String                        ' the name of the queue being edited
    Dim sWSName As String                           ' the name of the workspace being edited
    Dim iVisibleFields As Integer                   ' the number of visible fields
    Dim iField As Integer
    
    Call setDefaultProps
    
End Sub
' Complete the setup of the queue query specification.
Private Function SetupQueueQuerySpec(QueueQuerySpec As IDMObjects.QueueQuerySpecification, _
                                oProperties As IDMObjects.Properties, _
                                ByVal iTotalFields As Integer) As Boolean
Dim oProperty As IDMObjects.Property
Dim iCounter As Integer
Dim iFoundFields As Integer

iFoundFields = 0
Call getSystemProps(QueueQuerySpec)

' For each Property, find its associated DataLabel/DataField pair
' cast the data to the correct type and write it to the value field
' Note that we do need to compair the TypeID of the property to what the
' type the user selected.  This prevents a field with the same name, but
' different type from being picked.
For Each oProperty In oProperties
    For iCounter = 0 To iTotalFields - 1
        ' Verify the name of the fields are the same.
        If txtFieldName(iCounter).Text = oProperty.Name Then
            Select Case oProperty.PropertyDescription.TypeID
                Case idmTypeBoolean
                    oProperty.Value = CBool(txtFieldValue(iCounter).Text)
                    iFoundFields = iFoundFields + 1
                Case idmTypeByte
                    oProperty.Value = CByte(txtFieldValue(iCounter).Text)
                    iFoundFields = iFoundFields + 1
                Case idmTypeCurrency
                    oProperty.Value = CCur(txtFieldValue(iCounter).Text)
                    iFoundFields = iFoundFields + 1
                Case idmTypeDate
                    If cmbPropType(iCounter).ItemData(cmbPropType(iCounter).ListIndex) = idmQueueTypeDate Then
                        oProperty.Value = CDate(txtFieldValue(iCounter).Text)
                        iFoundFields = iFoundFields + 1
                    End If
                Case idmTypeObject
                    If cmbPropType(iCounter).ItemData(cmbPropType(iCounter).ListIndex) = idmQueueTypeDecimal Or _
                       cmbPropType(iCounter).ItemData(cmbPropType(iCounter).ListIndex) = idmQueueTypeNumber Then
                        If oProperty.PropertyDescription.GetState(idmPropSupportsFPNumber) Then
                            ' Produce FnFPNumber from string, so no loss of precision
                            ' happens in a conversion to and from a double
                            If IsEmpty(oProperty.Value) Or IsNull(oProperty.Value) Then
                                Dim tempFP As New IDMObjects.FnFPNumber
                                tempFP.ValueAsString = txtFieldValue(iCounter).Text
                                oProperty.Value = tempFP
                            Else
                                oProperty.FnFPNumber.ValueAsString = txtFieldValue(iCounter).Text
                            End If
                        Else
                            oProperty.Value = txtFieldValue(iCounter).Text
                        End If
                        iFoundFields = iFoundFields + 1
                    End If
                Case idmTypeLong
                    If cmbPropType(iCounter).ItemData(cmbPropType(iCounter).ListIndex) = idmQueueTypeFolder Then
                        oProperty.Value = CLng(txtFieldValue(iCounter).Text)
                        iFoundFields = iFoundFields + 1
                    End If
                Case idmTypeUnsignedLong
                    If cmbPropType(iCounter).ItemData(cmbPropType(iCounter).ListIndex) = idmQueueTypeNumber Then
                        oProperty.Value = CDbl(txtFieldValue(iCounter).Text)
                        iFoundFields = iFoundFields + 1
                    End If
                Case idmTypeShort
                    If cmbPropType(iCounter).ItemData(cmbPropType(iCounter).ListIndex) = idmQueueTypeSelection Then
                        oProperty.Value = CInt(txtFieldValue(iCounter).Text)
                        iFoundFields = iFoundFields + 1
                    End If
                Case idmTypeUnsignedShort
                    If cmbPropType(iCounter).ItemData(cmbPropType(iCounter).ListIndex) = idmQueueTypeInteger Then
                        oProperty.Value = CInt(txtFieldValue(iCounter).Text)
                        iFoundFields = iFoundFields + 1
                    End If
                Case idmTypeString
                    If cmbPropType(iCounter).ItemData(cmbPropType(iCounter).ListIndex) = idmQueueTypeString Then
                        oProperty.Value = txtFieldValue(iCounter).Text
                        iFoundFields = iFoundFields + 1
                    End If
                Case Else
                    oProperty.Value = txtFieldValue(iCounter).Text
                    iFoundFields = iFoundFields + 1
            End Select
        End If
    Next
Next

    ' Make sure we found all the fields we are looking for in the queue.
    If iTotalFields = iFoundFields Then
        SetupQueueQuerySpec = True
    Else
        SetupQueueQuerySpec = False
    End If
    
End Function

Private Sub OKButton_Click()
    Dim oProperties As IDMObjects.Properties                ' a collection of properties for a QueueEntry
    Dim iIndex As Integer
    On Error GoTo Errors
    
    If Not (OKButton_Validate()) Then
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    QMaint.MainStatusBar.SimpleText = "Querying workspace " & QMaint.cmbWorkspace.Text
    QMaint.MainStatusBar.Refresh

    ' Clear QMaint cmbQueues combo box and associated arrays
    For iIndex = QMaint.cmbQueues.listCount - 1 To 0 Step -1
        QMaint.cmbQueues.RemoveItem (iIndex)
        Set gWSQQuerySpec(iIndex) = Nothing
        Set gQueue(iIndex) = Nothing
    Next iIndex
    
    ' Loop through the queues in the workspace
    Set oWorkspace = QMaint.oLibrary.GetObject(idmObjTypeQueueWorkspace, QMaint.cmbWorkspace.Text)
    Set oQueueList = QMaint.oWorkspace.FilterQueues("")

    iIndex = 0
    For Each oQueue In oQueueList
    
        QMaint.MainStatusBar.SimpleText = "Querying workspace: " & QMaint.cmbWorkspace.Text & _
                                          "  /  Queue: " & oQueue.Name
        QMaint.MainStatusBar.Refresh
        
        ' Setup up the global workspace query specification
        Set gWSQQuerySpec(iIndex) = oQueue.CreateQuerySpecification()
        
        ' Get the properties from the QueueQuerySpecification
        Set oProperties = gWSQQuerySpec(iIndex).Filters
        
        ' Clear current values
        gWSQQuerySpec(iIndex).Clear

        ' Get total number of fields to search
        For iTotalFields = 0 To iMaxFields - 1
            If txtFieldName(iTotalFields).Text = "" Then
                Exit For
            End If
        Next
        
        ' Finish the query specification
        If SetupQueueQuerySpec(gWSQQuerySpec(iIndex), oProperties, iTotalFields) Then
            ' Make sure we have at least one entry
            If gWSQQuerySpec(iIndex).Count > 0 Then
            
                ' Save the queue
                Set gQueue(iIndex) = oQueue
        
                ' Make the combo box visible
                If QMaint.cmbQueues.Enabled = False Then
                    QMaint.cmbQueues.Enabled = True
                End If
            
                ' Add Queue to combo box on qmaint form.
                QMaint.cmbQueues.AddItem (oQueue.Name)
            
                ' Add the index to the itemdata of the combo box
                QMaint.cmbQueues.ItemData(QMaint.cmbQueues.NewIndex) = iIndex
            
                ' Incriment index
                iIndex = iIndex + 1
                
            End If
        End If
    Next
    
    QMaint.MainStatusBar.SimpleText = "Workspace Query Completed..."
    QMaint.MainStatusBar.Refresh

    Unload Me
    Set oQueueList = Nothing
    Screen.MousePointer = vbDefault
    
    Exit Sub
    
Errors:
    MsgBox "Setting the query values failed: " & Err.Description, vbExclamation, AppName
    Screen.MousePointer = vbDefault
End Sub

Private Sub ClearButton_Click()
    
    Call setDefaultProps
    
End Sub

Private Sub setDefaultProps()
    Dim iField As Integer
    Dim inx As Integer
    
    optEvenIfBusy.Value = True
    optOnlyIfBusy.Value = False
    optNotBusy.Value = False
    udMinPrio.Value = 0
    udMaxPrio.Value = 9
    cbCheckUser.Value = 0
    txtUser.Text = ""
    txtGroup.Text = ""
    cbDelayed.Value = 1
    txtDeadline.Text = ""
    optEvenIfInc.Value = True
    optOnlyIfInc.Value = False
    optNotInc.Value = False
    txtSort.Text = ""
    cbSortDesc.Value = 0
        
    ' Clear all values of the search fields
    For iField = 0 To iMaxFields - 1
        txtFieldName(iField) = ""
        cmbPropType(iField).ListIndex = -1
        txtFieldValue(iField) = ""
        ' Set the other two combo box values based on the first
        For inx = 0 To cmbPropType(0).listCount - 1
            cmbPropType(iField).List(inx) = cmbPropType(0).List(inx)
            cmbPropType(iField).ItemData(inx) = cmbPropType(0).ItemData(inx)
        Next inx
    Next iField
    
    ' Disable search field #2 and #3
    Call EnableField(iField2, False)
    Call EnableField(iField3, False)
    
End Sub
Private Sub EnableField(iField As Integer, bEnable As Boolean)
        
        frSearchField(iField).Enabled = bEnable
        lbFieldName(iField).Enabled = bEnable
        txtFieldName(iField).Enabled = bEnable
        lbFieldType(iField).Enabled = bEnable
        cmbPropType(iField).Enabled = bEnable
        lbFieldValue(iField).Enabled = bEnable
        txtFieldValue(iField).Enabled = bEnable

End Sub

Private Sub getSystemProps(QueueQuerySpec As IDMObjects.QueueQuerySpecification)
    If optEvenIfBusy.Value Then
        QueueQuerySpec.Status = idmBusyOK
    ElseIf optOnlyIfBusy.Value Then
        QueueQuerySpec.Status = idmBusyOnly
    Else
        QueueQuerySpec.Status = idmBusyNotOK
    End If
    QueueQuerySpec.MinPriority = meMinPrio.Text
    QueueQuerySpec.MaxPriority = meMaxPrio.Text
    QueueQuerySpec.CheckUser = (cbCheckUser.Value = 1)
    QueueQuerySpec.UserName = txtUser.Text
    QueueQuerySpec.GroupName = txtGroup.Text
    QueueQuerySpec.EvenIfDelayed = (cbDelayed.Value = 1)
    If txtDeadline.Text = "" Then
        QueueQuerySpec.Deadline = CDate(idmQueueNoTimeOut)
    Else
        QueueQuerySpec.Deadline = CDate(txtDeadline.Text)
    End If
    If optEvenIfInc.Value Then
        QueueQuerySpec.Incomplete = idmIncompleteOK
    ElseIf optOnlyIfInc.Value Then
        QueueQuerySpec.Incomplete = idmIncompleteOnly
    Else
        QueueQuerySpec.Incomplete = idmIncompleteNotOK
    End If
    QueueQuerySpec.SortField = txtSort.Text
    QueueQuerySpec.SortDescending = (cbSortDesc.Value = 1)
    
End Sub

Private Sub txtFieldName_Validate(Index As Integer, Cancel As Boolean)

    If Index = iField1 Then
        If Not (txtFieldName(Index).Text = "") And _
           Not (cmbPropType(Index).ListIndex = -1) Then
           Call EnableField(iField2, True)
        End If
    ElseIf Index = iField2 Then
        If Not (txtFieldName(Index).Text = "") And _
           Not (cmbPropType(Index).ListIndex = -1) Then
           Call EnableField(iField3, True)
        End If
    End If
    
End Sub

Private Sub cmbPropType_Validate(Index As Integer, Cancel As Boolean)
    
    If Index = iField1 Then
        If Not (txtFieldName(Index).Text = "") And _
           Not (cmbPropType(Index).ListIndex = -1) Then
           Call EnableField(iField2, True)
        End If
    ElseIf Index = iField2 Then
        If Not (txtFieldName(Index).Text = "") And _
           Not (cmbPropType(Index).ListIndex = -1) Then
           Call EnableField(iField3, True)
        End If
    End If

End Sub
Private Function OKButton_Validate() As Boolean
    Dim iFieldsEnabled As Integer
    OKButton_Validate = True

    ' Sanity check for the Search Fields
    If (frSearchField(iField3).Enabled) And _
       ((txtFieldName(iField3).Text <> "") Or _
        (cmbPropType(iField3).ListIndex <> -1) Or _
        (txtFieldValue(iField3).Text <> "")) Then
        
        If Not (PreviousFieldsValid(iField3)) Then
            OKButton_Validate = False
        End If
        
    ElseIf (frSearchField(iField2).Enabled) And _
           ((txtFieldName(iField2).Text <> "") Or _
            (cmbPropType(iField2).ListIndex <> -1) Or _
            (txtFieldValue(iField2).Text <> "")) Then

        If Not (PreviousFieldsValid(iField2)) Then
            OKButton_Validate = False
        End If
        
    ElseIf (frSearchField(iField1).Enabled) And _
           ((txtFieldName(iField1).Text <> "") Or _
            (cmbPropType(iField1).ListIndex <> -1) Or _
            (txtFieldValue(iField1).Text <> "")) Then

            If Not (PreviousFieldsValid(iField1)) Then
            OKButton_Validate = False
        End If
        
    End If
    
End Function

Private Function PreviousFieldsValid(iField As Integer) As Boolean
    Dim Index As Integer
    PreviousFieldsValid = True
    
    ' Verify all fields are valid
    For Index = iField To 0 Step -1
        If txtFieldName(Index).Text = "" Then
            MsgBox "Field Name in Search Field #" & Index + 1 & " is required", vbExclamation, AppName
            PreviousFieldsValid = False
            Exit Function
        End If
        
        If cmbPropType(Index).ListIndex = -1 Then
            MsgBox "Field Type in Search Field #" & Index + 1 & " is required", vbExclamation, AppName
            PreviousFieldsValid = False
            Exit Function
        End If
           
    Next

End Function
