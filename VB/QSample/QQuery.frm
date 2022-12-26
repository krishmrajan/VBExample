VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form QQuery 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Queue Query"
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
      TabIndex        =   30
      Top             =   1095
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   450
      _Version        =   393216
      BuddyControl    =   "meMaxPrio"
      BuddyDispid     =   196638
      OrigLeft        =   4680
      OrigTop         =   1095
      OrigRight       =   4920
      OrigBottom      =   1350
      SyncBuddy       =   -1  'True
      BuddyProperty   =   22
      Enabled         =   -1  'True
   End
   Begin MSComCtl2.UpDown udMinPrio 
      Height          =   255
      Left            =   2040
      TabIndex        =   29
      Top             =   1095
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   450
      _Version        =   393216
      BuddyControl    =   "meMinPrio"
      BuddyDispid     =   196639
      OrigLeft        =   2040
      OrigTop         =   1095
      OrigRight       =   2280
      OrigBottom      =   1350
      SyncBuddy       =   -1  'True
      BuddyProperty   =   22
      Enabled         =   -1  'True
   End
   Begin VB.Frame frBusy 
      Caption         =   "Busy"
      Height          =   615
      Left            =   360
      TabIndex        =   28
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
      TabIndex        =   5
      Top             =   1095
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      _Version        =   393216
      MaxLength       =   1
      Mask            =   "#"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox meMinPrio 
      Height          =   255
      Left            =   1800
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
      TabIndex        =   15
      Top             =   5295
      Width           =   1650
   End
   Begin VB.TextBox txtSort 
      Height          =   285
      Left            =   2595
      MaxLength       =   18
      TabIndex        =   14
      Top             =   4815
      Width           =   2820
   End
   Begin VB.TextBox txtUser 
      Height          =   285
      Left            =   1830
      MaxLength       =   82
      TabIndex        =   7
      Top             =   2055
      Width           =   3585
   End
   Begin VB.Frame frIncomplete 
      Caption         =   "Incomplete"
      Height          =   615
      Left            =   360
      TabIndex        =   20
      Top             =   3975
      Width           =   5055
      Begin VB.OptionButton optNotInc 
         Alignment       =   1  'Right Justify
         Caption         =   "Not"
         Height          =   255
         Left            =   3840
         TabIndex        =   13
         Top             =   240
         Width           =   615
      End
      Begin VB.OptionButton optOnlyIfInc 
         Alignment       =   1  'Right Justify
         Caption         =   "Only If"
         Height          =   255
         Left            =   1800
         TabIndex        =   12
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton optEvenIfInc 
         Alignment       =   1  'Right Justify
         Caption         =   "Even If"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.TextBox txtGroup 
      Height          =   285
      Left            =   1830
      MaxLength       =   82
      TabIndex        =   8
      Top             =   2535
      Width           =   3585
   End
   Begin VB.CheckBox cbDelayed 
      Alignment       =   1  'Right Justify
      Caption         =   "Even If Delayed"
      Height          =   330
      Left            =   360
      TabIndex        =   9
      Top             =   3015
      Width           =   1650
   End
   Begin VB.TextBox txtDeadline 
      Height          =   285
      Left            =   1830
      TabIndex        =   10
      Top             =   3495
      Width           =   1575
   End
   Begin VB.CheckBox cbCheckUser 
      Alignment       =   1  'Right Justify
      Caption         =   "Check User"
      Height          =   330
      Left            =   360
      TabIndex        =   6
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
   Begin VB.TextBox DataField 
      Height          =   285
      Index           =   0
      Left            =   7680
      TabIndex        =   22
      Top             =   390
      Width           =   3015
   End
   Begin VB.VScrollBar sbProperties 
      Height          =   3930
      Left            =   11040
      TabIndex        =   23
      Top             =   240
      Visible         =   0   'False
      Width           =   255
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
      TabIndex        =   21
      Top             =   4815
      Width           =   975
   End
   Begin VB.Label lbUser 
      Caption         =   "User Name"
      Height          =   255
      Left            =   360
      TabIndex        =   17
      Top             =   2055
      Width           =   1095
   End
   Begin VB.Label lbMaxPrio 
      Caption         =   "Max. Priority"
      Height          =   255
      Left            =   3240
      TabIndex        =   16
      Top             =   1095
      Width           =   975
   End
   Begin VB.Label lbGroup 
      Caption         =   "Group Name"
      Height          =   255
      Left            =   360
      TabIndex        =   18
      Top             =   2535
      Width           =   1095
   End
   Begin VB.Label lbDeadline 
      Alignment       =   2  'Center
      Caption         =   "Deadline"
      Height          =   255
      Left            =   360
      TabIndex        =   19
      Top             =   3495
      Width           =   735
   End
   Begin VB.Label DataLabel 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   5850
      TabIndex        =   27
      Top             =   390
      Width           =   1455
   End
End
Attribute VB_Name = "QQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public iTotalFields As Integer                      ' the total number of fields displayed on the form
Private Const visibleProps = 20                     ' the number of visible properties on form
Private Const iSpacing = 120                        ' the vertical spacing between the datafields

' Cancel out of edit
Private Sub CancelButton_Click()
    QMaint.MainStatusBar.SimpleText = "Query cancelled at user's request."
    
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
    
    Call setSystemProps
    
    ' get the properties from the QueueQuerySpecification
    Set oProperties = QMaint.oQueueQuerySpec.Filters
    
    iCounter = 0
    iHeight = DataField(0).Height
    
    ' for each property (Name/Value), add a field and label to the form.  Fill the field and label
    ' DataLabel and DataField are declared as controls arrays on the form
    For Each oProperty In oProperties
        If iCounter >= 1 Then
            Load DataLabel(iCounter)
            Load DataField(iCounter)
            If iCounter >= visibleProps Then
                sbProperties.Max = iCounter - visibleProps + 1
                sbProperties.Visible = True
            End If
        End If
        DataLabel(iCounter).Caption = oProperty.Name & ":"
            
        If IsEmpty(oProperty.Value) Or IsNull(oProperty.Value) Or _
           Not oProperty.PropertyDescription.GetState(idmPropSupportsFPNumber) Then
                DataField(iCounter).Text = CStr(IIf(IsNull(oProperty.Value), "", oProperty.Value))
           Else
                ' Get string from FnFPNumber, so no loss of precision happens
                ' in a conversion to and from a double.
                DataField(iCounter).Text = oProperty.FnFPNumber.ValueAsString
        End If
        
        If iCounter < visibleProps Then
            DataLabel(iCounter).Top = DataLabel(0).Top + (iHeight + iSpacing) * iCounter
            DataField(iCounter).Top = DataField(0).Top + (iHeight + iSpacing) * iCounter
            Call setVisible(iCounter, True)
        Else
            Call setVisible(iCounter, False)
        End If
        DataLabel(iCounter).TabIndex = DataLabel(0).TabIndex + 2 * iCounter
        DataField(iCounter).TabIndex = DataLabel(0).TabIndex + 2 * iCounter + 1
        iCounter = iCounter + 1
    Next
    
    iTotalFields = iCounter
    If iTotalFields <= visibleProps Then
        iVisibleFields = iCounter
    Else
        iVisibleFields = visibleProps
    End If
    
    ' move the OKbutton, cancelbutton, and clearbutton below the last field and resize the form to fit everything
    ' adjust scroll bar
    OKButton.TabIndex = DataLabel(0).TabIndex + 2 * iCounter
    CancelButton.TabIndex = DataLabel(0).TabIndex + 2 * iCounter + 1
    ClearButton.TabIndex = DataLabel(0).TabIndex + 2 * iCounter + 2
    sbProperties.TabIndex = DataLabel(0).TabIndex + 2 * iCounter + 3
    OKButton.Top = DataField(0).Top + (iHeight + iSpacing) * iVisibleFields + iSpacing
    If OKButton.Top < cbSortDesc.Top + cbSortDesc.Height + iSpacing Then
        OKButton.Top = cbSortDesc.Top + cbSortDesc.Height + iSpacing
    End If
    CancelButton.Top = OKButton.Top
    ClearButton.Top = OKButton.Top
    sbProperties.Height = visibleProps * (iHeight + iSpacing) - iSpacing
    Me.Height = OKButton.Top + OKButton.Height + (5 * iSpacing)
    
End Sub
' Update the properties in the queue query
Private Sub UpdateQueueQuery(oProperties As IDMObjects.Properties, _
    ByVal iTotalFields As Integer)
Dim oProperty As IDMObjects.Property
Dim iCounter As Integer

Call getSystemProps

' for each Property, find its associated DataLabel/DataField pair
' cast the data to the correct type and write it to the value field
For Each oProperty In oProperties
    For iCounter = 0 To iTotalFields - 1
        If DataLabel(iCounter).Caption = oProperty.Name & ":" Then
            If DataField(iCounter).Text <> "" Then
                Select Case oProperty.PropertyDescription.TypeID
                Case idmTypeBoolean
                    oProperty.Value = CBool(DataField(iCounter).Text)
                Case idmTypeByte
                    oProperty.Value = CByte(DataField(iCounter).Text)
                Case idmTypeCurrency
                    oProperty.Value = CCur(DataField(iCounter).Text)
                Case idmTypeDate
                    oProperty.Value = CDate(DataField(iCounter).Text)
                Case idmTypeObject
                    If oProperty.PropertyDescription.GetState(idmPropSupportsFPNumber) Then
                        ' Produce FnFPNumber from string, so no loss of precision
                        ' happens in a conversion to and from a double
                        If IsEmpty(oProperty.Value) Or IsNull(oProperty.Value) Then
                            Dim tempFP As New IDMObjects.FnFPNumber
                            tempFP.ValueAsString = DataField(iCounter).Text
                            oProperty.Value = tempFP
                        Else
                            oProperty.FnFPNumber.ValueAsString = DataField(iCounter).Text
                        End If
                    Else
                        oProperty.Value = DataField(iCounter).Text
                    End If
                Case idmTypeLong
                    oProperty.Value = CLng(DataField(iCounter).Text)
                Case idmTypeUnsignedLong
                    oProperty.Value = CDbl(DataField(iCounter).Text)
                Case idmTypeShort
                    oProperty.Value = CInt(DataField(iCounter).Text)
                Case idmTypeUnsignedShort
                    oProperty.Value = CInt(DataField(iCounter).Text)
                Case idmTypeString
                    oProperty.Value = DataField(iCounter).Text
                Case Else
                    oProperty.Value = DataField(iCounter).Text
                End Select
            End If
        End If
    Next
Next
End Sub

Private Sub OKButton_Click()
    Dim oProperties As IDMObjects.Properties                ' a collection of properties for a QueueEntry
    
    On Error GoTo Errors
    
    Set oProperties = QMaint.oQueueQuerySpec.Filters
    QMaint.oQueueQuerySpec.Clear
      
    ' Get the fields
    Call UpdateQueueQuery(oProperties, iTotalFields)
    Call QMaint.RefreshGrid
    Unload Me
    
    Exit Sub
    
Errors:
    MsgBox "Setting the query values failed: " & Err.Description, vbExclamation, AppName
End Sub

Private Sub ClearButton_Click()
    Dim iCounter As Integer
    
    optEvenIfBusy.Value = True
    udMinPrio.Value = 0
    udMaxPrio.Value = 9
    cbCheckUser.Value = 0
    txtUser.Text = ""
    txtGroup.Text = ""
    cbDelayed.Value = 1
    txtDeadline.Text = ""
    optEvenIfInc.Value = True
    txtSort.Text = ""
    cbSortDesc.Value = 0
    For iCounter = 0 To iTotalFields - 1
        DataField(iCounter) = ""
    Next
    
End Sub

Private Sub setVisible(inx As Integer, visibility As Boolean)
    DataLabel(inx).Visible = visibility
    DataField(inx).Visible = visibility
End Sub

Private Sub sbProperties_Change()
    Dim inx As Integer
    Dim pos As Integer
    Dim iHeight As Integer                          ' the height of the standard DataField field
    
    iHeight = DataField(0).Height
    pos = sbProperties.Value
    For inx = 0 To iTotalFields - 1
        If inx >= pos And inx < pos + visibleProps Then
            DataLabel(inx).Top = DataLabel(0).Top + _
                                 (iHeight + iSpacing) * (inx - pos)
            DataField(inx).Top = DataField(0).Top + _
                                 (iHeight + iSpacing) * (inx - pos)
            Call setVisible(inx, True)
        Else
            Call setVisible(inx, False)
        End If
    Next inx
End Sub


Private Sub setSystemProps()
    If QMaint.oQueueQuerySpec.Status = idmBusyOK Then
        optEvenIfBusy.Value = True
    ElseIf QMaint.oQueueQuerySpec.Status = idmBusyOnly Then
        optOnlyIfBusy.Value = True
    Else
        optNotBusy.Value = True
    End If
    udMinPrio.Value = QMaint.oQueueQuerySpec.MinPriority
    udMaxPrio.Value = QMaint.oQueueQuerySpec.MaxPriority
    cbCheckUser.Value = IIf(QMaint.oQueueQuerySpec.CheckUser, 1, 0)
    txtUser.Text = QMaint.oQueueQuerySpec.UserName
    txtGroup.Text = QMaint.oQueueQuerySpec.GroupName
    cbDelayed.Value = IIf(QMaint.oQueueQuerySpec.EvenIfDelayed, 1, 0)
    txtDeadline.Text = IIf(QMaint.oQueueQuerySpec.Deadline = CDate(idmQueueNoTimeOut), "", CStr(QMaint.oQueueQuerySpec.Deadline))
    If QMaint.oQueueQuerySpec.Incomplete = idmIncompleteOK Then
        optEvenIfInc.Value = True
    ElseIf QMaint.oQueueQuerySpec.Incomplete = idmIncompleteOnly Then
        optOnlyIfInc.Value = True
    Else
        optNotInc.Value = True
    End If
    txtSort.Text = QMaint.oQueueQuerySpec.SortField
    cbSortDesc.Value = IIf(QMaint.oQueueQuerySpec.SortDescending, 1, 0)
End Sub

Private Sub getSystemProps()
    If optEvenIfBusy.Value Then
        QMaint.oQueueQuerySpec.Status = idmBusyOK
    ElseIf optOnlyIfBusy.Value Then
        QMaint.oQueueQuerySpec.Status = idmBusyOnly
    Else
        QMaint.oQueueQuerySpec.Status = idmBusyNotOK
    End If
    QMaint.oQueueQuerySpec.MinPriority = meMinPrio.Text
    QMaint.oQueueQuerySpec.MaxPriority = meMaxPrio.Text
    QMaint.oQueueQuerySpec.CheckUser = (cbCheckUser.Value = 1)
    QMaint.oQueueQuerySpec.UserName = txtUser.Text
    QMaint.oQueueQuerySpec.GroupName = txtGroup.Text
    QMaint.oQueueQuerySpec.EvenIfDelayed = (cbDelayed.Value = 1)
    If txtDeadline.Text = "" Then
        QMaint.oQueueQuerySpec.Deadline = CDate(idmQueueNoTimeOut)
    Else
        QMaint.oQueueQuerySpec.Deadline = CDate(txtDeadline.Text)
    End If
    If optEvenIfInc.Value Then
        QMaint.oQueueQuerySpec.Incomplete = idmIncompleteOK
    ElseIf optOnlyIfInc.Value Then
        QMaint.oQueueQuerySpec.Incomplete = idmIncompleteOnly
    Else
        QMaint.oQueueQuerySpec.Incomplete = idmIncompleteNotOK
    End If
    QMaint.oQueueQuerySpec.SortField = txtSort.Text
    QMaint.oQueueQuerySpec.SortDescending = (cbSortDesc.Value = 1)
End Sub
