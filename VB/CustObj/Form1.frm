VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "comctl32.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6675
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   5670
   LinkTopic       =   "Form1"
   ScaleHeight     =   6675
   ScaleWidth      =   5670
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   495
      Left            =   3240
      TabIndex        =   42
      Top             =   6120
      Width           =   1335
   End
   Begin VB.TextBox txtLibName 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1440
      TabIndex        =   41
      Top             =   120
      Width           =   2655
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   495
      Left            =   4440
      TabIndex        =   39
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   495
      Left            =   1080
      TabIndex        =   38
      Top             =   6120
      Width           =   1335
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4215
      Left            =   240
      TabIndex        =   7
      Top             =   1800
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   7435
      _Version        =   327681
      Style           =   1
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "General"
      TabPicture(0)   =   "Form1.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label4"
      Tab(0).Control(1)=   "Label5"
      Tab(0).Control(2)=   "Label6"
      Tab(0).Control(3)=   "Label7"
      Tab(0).Control(4)=   "Label8"
      Tab(0).Control(5)=   "Label9"
      Tab(0).Control(6)=   "Label10"
      Tab(0).Control(7)=   "Label11"
      Tab(0).Control(8)=   "Label12"
      Tab(0).Control(9)=   "lblClassDesc"
      Tab(0).Control(10)=   "lblID"
      Tab(0).Control(11)=   "lblObjectType"
      Tab(0).Control(12)=   "lblSystemType"
      Tab(0).Control(13)=   "lblCanDelete"
      Tab(0).Control(14)=   "lblCanModify"
      Tab(0).Control(15)=   "lblModified"
      Tab(0).Control(16)=   "lblReserved"
      Tab(0).Control(17)=   "lblName"
      Tab(0).ControlCount=   18
      TabCaption(1)   =   "Permissions"
      TabPicture(1)   =   "Form1.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lvPermissions"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Properties"
      TabPicture(2)   =   "Form1.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Frame1"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Frame2"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Frame3"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).ControlCount=   3
      Begin VB.Frame Frame3 
         Caption         =   "Property:"
         Height          =   1095
         Left            =   360
         TabIndex        =   33
         Top             =   2760
         Width           =   4455
         Begin VB.TextBox txtPropValue 
            Enabled         =   0   'False
            Height          =   285
            Left            =   840
            TabIndex        =   37
            Top             =   600
            Width           =   2535
         End
         Begin VB.Label lblPropName 
            AutoSize        =   -1  'True
            Caption         =   "lblPropName"
            Height          =   195
            Left            =   840
            TabIndex        =   36
            Top             =   360
            Visible         =   0   'False
            Width           =   900
         End
         Begin VB.Label Label16 
            Caption         =   "Value:"
            Height          =   255
            Left            =   120
            TabIndex        =   35
            Top             =   600
            Width           =   855
         End
         Begin VB.Label Label15 
            Caption         =   "Name:"
            Height          =   255
            Left            =   120
            TabIndex        =   34
            Top             =   360
            Width           =   735
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Extended Properties:"
         Height          =   855
         Left            =   360
         TabIndex        =   29
         Top             =   1800
         Width           =   4455
         Begin VB.CommandButton cmdGetExtProp 
            Caption         =   "Get"
            Height          =   495
            Left            =   3120
            TabIndex        =   32
            Top             =   240
            Width           =   1095
         End
         Begin VB.TextBox txtExtPropID 
            Height          =   285
            Left            =   840
            TabIndex        =   30
            Top             =   360
            Width           =   2175
         End
         Begin VB.Label Label13 
            Caption         =   "Name:"
            Height          =   255
            Left            =   120
            TabIndex        =   31
            Top             =   360
            Width           =   495
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Properties:"
         Height          =   1095
         Left            =   360
         TabIndex        =   27
         Top             =   600
         Width           =   4455
         Begin VB.ListBox lstProperties 
            Height          =   645
            Left            =   240
            TabIndex        =   28
            Top             =   240
            Width           =   3855
         End
      End
      Begin ComctlLib.ListView lvPermissions 
         Height          =   1455
         Left            =   -74640
         TabIndex        =   26
         Top             =   1080
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   2566
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   327682
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         Caption         =   "lblName"
         Height          =   195
         Left            =   -73320
         TabIndex        =   25
         Top             =   1080
         Visible         =   0   'False
         Width           =   570
      End
      Begin VB.Label lblReserved 
         AutoSize        =   -1  'True
         Caption         =   "lblReserved"
         Height          =   195
         Left            =   -73320
         TabIndex        =   24
         Top             =   2520
         Visible         =   0   'False
         Width           =   840
      End
      Begin VB.Label lblModified 
         AutoSize        =   -1  'True
         Caption         =   "lblModified"
         Height          =   195
         Left            =   -73320
         TabIndex        =   23
         Top             =   2280
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.Label lblCanModify 
         AutoSize        =   -1  'True
         Caption         =   "lblCanModify"
         Height          =   195
         Left            =   -73320
         TabIndex        =   22
         Top             =   2040
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.Label lblCanDelete 
         AutoSize        =   -1  'True
         Caption         =   "lblCanDelete"
         Height          =   195
         Left            =   -73320
         TabIndex        =   21
         Top             =   1800
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.Label lblSystemType 
         AutoSize        =   -1  'True
         Caption         =   "lblSystemType"
         Height          =   195
         Left            =   -73320
         TabIndex        =   20
         Top             =   1560
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.Label lblObjectType 
         AutoSize        =   -1  'True
         Caption         =   "lblObjectType"
         Height          =   195
         Left            =   -73320
         TabIndex        =   19
         Top             =   1320
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label lblID 
         AutoSize        =   -1  'True
         Caption         =   "lblID"
         Height          =   195
         Left            =   -73320
         TabIndex        =   18
         Top             =   840
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.Label lblClassDesc 
         AutoSize        =   -1  'True
         Caption         =   "lblClassDesc"
         Height          =   195
         Left            =   -73320
         TabIndex        =   17
         Top             =   600
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.Label Label12 
         Caption         =   "Reserved:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   16
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label Label11 
         Caption         =   "Modified:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   15
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Label Label10 
         Caption         =   "Can Modify:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   14
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label9 
         Caption         =   "Can Delete:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   13
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label8 
         Caption         =   "System Type:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   12
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "Object Type:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   11
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "Name:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   10
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "ID:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   9
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Class Description:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   8
         Top             =   600
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdGetObj 
      Caption         =   "Get"
      Height          =   495
      Left            =   4440
      TabIndex        =   6
      Top             =   480
      Width           =   1095
   End
   Begin VB.TextBox txtKey 
      Height          =   285
      Left            =   1440
      TabIndex        =   5
      Text            =   "SDMROOT"
      Top             =   1320
      Width           =   2655
   End
   Begin VB.TextBox txtObjectType 
      Height          =   285
      Left            =   1440
      TabIndex        =   4
      Text            =   "MTISDM"
      Top             =   960
      Width           =   2655
   End
   Begin VB.TextBox txtISVID 
      Height          =   285
      Left            =   1440
      TabIndex        =   3
      Text            =   "SAROSOPG"
      Top             =   600
      Width           =   2655
   End
   Begin VB.Label Label14 
      Caption         =   "Library:"
      Height          =   255
      Left            =   240
      TabIndex        =   40
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Key:"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Object Type:"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Isv Id:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   855
   End
   Begin VB.Menu mnuLogin 
      Caption         =   "Login"
   End
   Begin VB.Menu mnuDelete 
      Caption         =   "Delete"
   End
   Begin VB.Menu mnuRefresh 
      Caption         =   "Refresh"
      Begin VB.Menu mnuRefreshObject 
         Caption         =   "Refresh Object"
         Begin VB.Menu mnuRefreshObjectProperties 
            Caption         =   "Properties"
         End
      End
      Begin VB.Menu mnuRefreshTabs 
         Caption         =   "Refresh Tabs"
         Begin VB.Menu mnuRefreshTabsCurrent 
            Caption         =   "Refresh Current"
         End
         Begin VB.Menu mnuRefreshTabsAll 
            Caption         =   "Refresh All"
         End
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oCust As IDMObjects.CustomObject


Private Sub cmdAdd_Click()
    On Error GoTo ErrorHandler
    Dim oProp As IDMObjects.Property
    Set oCust = oLib.CreateObject(idmObjTypeCustomObject, txtISVID.Text)
    'set id, objecttype, key
    Set oProp = oCust.Properties("idmCustObjType")  'object type
    oProp.Value = txtObjectType.Text
    Set oProp = oCust.Properties("idmCustObjKey")   'key
    oProp.Value = txtKey.Text
    oCust.Save 'save
    RefreshTabs
    Exit Sub

ErrorHandler:
    Select Case gErrorHandler
    Case vbAbort
        Exit Sub
    Case vbRetry
        Resume
    Case vbIgnore
        Resume Next
    End Select
    
End Sub

Private Sub cmdGetExtProp_Click()
    On Error GoTo ErrorHandler
    Dim oProp As IDMObjects.Property
    Dim oTmp As Object
    If txtExtPropID <> "" Then
        Set oProp = oCust.GetExtendedProperty(txtExtPropID.Text)
        lblPropName.Caption = oProp.Name
    
        txtPropValue.Text = oProp.Value
        lblPropName.Visible = True
        txtPropValue.Visible = True
    Else
        MsgBox ("Please enter a customer property name")
    End If
    Exit Sub
ErrorHandler:
    Select Case gErrorHandler
    Case vbAbort
        Exit Sub
    Case vbRetry
        Resume
    Case vbIgnore
        Resume Next
    End Select
    
End Sub

Private Sub cmdGetObj_Click()
    On Error GoTo ErrorHandler
    Set oCust = oLib.GetObject(idmObjTypeCustomObject, txtISVID.Text & "^" & txtObjectType.Text & "^" & txtKey.Text)
    RefreshTabs
    Exit Sub
ErrorHandler:
    Select Case gErrorHandler
    Case vbAbort
        Exit Sub
    Case vbRetry
        Resume
    Case vbIgnore
        Resume Next
    End Select
    
End Sub

Private Sub cmdSave_Click()
    On Error GoTo ErrorHandler
    oCust.Save
    MsgBox ("Custom object saved...")
    Exit Sub
ErrorHandler:
    Select Case gErrorHandler
    Case vbAbort
        Exit Sub
    Case vbRetry
        Resume
    Case vbIgnore
        Resume Next
    End Select
    
End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
    'permissions list headers
    
    lvPermissions.ColumnHeaders.Add 1, , "Name"
    lvPermissions.ColumnHeaders.Add 2, , "Type"
    lvPermissions.ColumnHeaders.Add 3, , "Access"
    EnableAll False
    Call mnuLogin_Click
        
End Sub



Private Sub lstProperties_Click()
    On Error GoTo ErrorHandler
    Dim oProp As IDMObjects.Property
    Set oProp = oCust.Properties(lstProperties.ListIndex + 1)
    lblPropName.Caption = oProp.Name
    txtPropValue.Text = oProp.Value
    lblPropName.Visible = True
    txtPropValue.Visible = True
    
    Exit Sub
ErrorHandler:
    Select Case gErrorHandler
    Case vbAbort
        Exit Sub
    Case vbRetry
        Resume
    Case vbIgnore
        Resume Next
    End Select
    
    
End Sub

Private Sub mnuDelete_Click()
    On Error GoTo ErrorHandler
    If MsgBox("Delete Custom Object?", vbYesNo, "Delete") = vbYes Then
        oCust.Delete
    End If
    Exit Sub
ErrorHandler:
    Select Case gErrorHandler
    Case vbAbort
        Exit Sub
    Case vbRetry
        Resume
    Case vbIgnore
        Resume Next
    End Select
    
End Sub

Private Sub mnuLogin_Click()
   On Error GoTo ErrorHandler
    Set oLib = Nothing
    'frmLogon.Show vbModal
    Load frmLogon
    If frmLogon.CmbLibraries.ListCount = 0 Then
        Unload frmLogon
        Exit Sub
    End If
    frmLogon.Show vbModal
    If oLib.GetState(idmLibraryLoggedOn) Then
        EnableAll True
    txtLibName = oLib.Label
    End If
    Exit Sub
ErrorHandler:
    Select Case gErrorHandler
    Case vbAbort
        Exit Sub
    Case vbRetry
        Resume
    Case vbIgnore
        Resume Next
    End Select
    
End Sub

Private Sub LoadTabGeneral()
    On Error GoTo ErrorHandler
    lblClassDesc.Caption = oCust.ClassDescription.Name
    lblID.Caption = oCust.ID
    lblName.Caption = oCust.Name
    lblObjectType.Caption = oCust.ObjectType
    lblSystemType.Caption = oCust.SystemType
    lblCanDelete.Caption = oCust.GetState(idmCustomObjectCanDelete)
    lblCanModify.Caption = oCust.GetState(idmCustomObjectCanModify)
    lblModified.Caption = oCust.GetState(idmCustomObjectModified)
    lblReserved.Caption = oCust.GetState(idmCustomObjectReserved)
    
    lblClassDesc.Visible = True
    lblID.Visible = True
    lblName.Visible = True
    lblObjectType.Visible = True
    lblSystemType.Visible = True
    lblCanDelete.Visible = True
    lblCanModify.Visible = True
    lblModified.Visible = True
    lblReserved.Visible = True
    Exit Sub
ErrorHandler:
    Select Case gErrorHandler
    Case vbAbort
        Exit Sub
    Case vbRetry
        Resume
    Case vbIgnore
        Resume Next
    End Select
    
End Sub

Private Sub LoadTabPermissions()
    On Error GoTo ErrorHandler
    ' get permissions:
    Dim oPerms As IDMObjects.Permissions
    Dim lvItem As ListItem
    Dim ii As Integer
    lvPermissions.ListItems.Clear
    Set oPerms = oCust.Permissions
    For ii = 1 To oPerms.Count
        Set lvItem = lvPermissions.ListItems.Add(, , oPerms.Item(ii).GranteeName)
        Select Case oPerms.Item(ii).GranteeType
        Case idmObjTypeUser
            lvItem.SubItems(1) = "User"
        Case idmObjTypeGroup
            lvItem.SubItems(1) = "Group"
        End Select
        Select Case oPerms.Item(ii).Access
        Case idmDSAccessAdmin
            lvItem.SubItems(2) = "Admin"
        Case idmDSAccessAuthor
            lvItem.SubItems(2) = "Author"
        Case idmDSAccessViewer
            lvItem.SubItems(2) = "Viewer"
        Case idmDSAccessOwner
            lvItem.SubItems(2) = "Owner"
        Case idmDSAccessNone
            lvItem.SubItems(2) = "None"
        End Select
    Next ii
    Exit Sub
ErrorHandler:
    Select Case gErrorHandler
    Case vbAbort
        Exit Sub
    Case vbRetry
        Resume
    Case vbIgnore
        Resume Next
    End Select
    
End Sub
Private Sub LoadTabProperties()
    On Error GoTo ErrorHandler
    Dim oProp As IDMObjects.Property
    lstProperties.Clear
    For Each oProp In oCust.Properties
        lstProperties.AddItem oProp.Name
    Next oProp
    Exit Sub
ErrorHandler:
    Select Case gErrorHandler
    Case vbAbort
        Exit Sub
    Case vbRetry
        Resume
    Case vbIgnore
        Resume Next
    End Select
    
End Sub

Private Sub RefreshTabs(Optional iTab As Variant)
    If IsMissing(iTab) Then
        iTab = -1
    End If
    Select Case iTab
    Case 0  'general
        LoadTabGeneral
    Case 1  'permissions
        LoadTabPermissions
    Case 2  'props
        LoadTabProperties
    Case Else
        LoadTabGeneral
        LoadTabPermissions
        LoadTabProperties
    End Select
End Sub

Private Sub mnuRefreshObjectProperties_Click()
    On Error GoTo ErrorHandler
    oCust.Refresh idmCustomObjectRefreshProperties
    Exit Sub
ErrorHandler:
    Select Case gErrorHandler
    Case vbAbort
        Exit Sub
    Case vbRetry
        Resume
    Case vbIgnore
        Resume Next
    End Select

End Sub



Private Sub mnuRefreshTabsAll_Click()
    RefreshTabs
End Sub

Private Sub mnuRefreshTabsCurrent_Click()
    RefreshTabs SSTab1.Tab
End Sub

Private Sub EnableAll(bEnabled As Boolean)
    Dim ctl As Control
    For Each ctl In Me.Controls
        If ctl.Name <> "mnuLogin" And _
           ctl.Name <> "txtPropValue" And _
           ctl.Name <> "txtLibName" Then
            ctl.Enabled = bEnabled
        End If
    Next ctl
End Sub

Private Function gErrorHandler() As VbMsgBoxResult
    Dim sMessage As String
    If Err.Number = 0 Then
        sMessage = "Unknown Error"
    Else
        sMessage = "Error " & Err.Number & ": " & Err.Description
    End If
    gErrorHandler = MsgBox(sMessage, vbAbortRetryIgnore, "OOPS")
End Function

