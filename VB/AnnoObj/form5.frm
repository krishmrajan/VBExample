VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form Form5 
   Caption         =   "Modify Security"
   ClientHeight    =   5040
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   6156
   LinkTopic       =   "Form5"
   ScaleHeight     =   5040
   ScaleWidth      =   6156
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton Option1 
      Caption         =   "Group"
      Height          =   375
      Index           =   1
      Left            =   3000
      TabIndex        =   9
      Top             =   2640
      Width           =   2295
   End
   Begin VB.CommandButton Remove 
      Caption         =   "Remove"
      Height          =   375
      Left            =   2760
      TabIndex        =   8
      Top             =   4200
      Width           =   1095
   End
   Begin VB.CommandButton Add 
      Caption         =   "Add"
      Height          =   375
      Left            =   1200
      TabIndex        =   7
      Top             =   4200
      Width           =   1215
   End
   Begin VB.ComboBox AccessLevelCmb 
      Height          =   315
      Left            =   3360
      TabIndex        =   3
      Top             =   3600
      Width           =   2655
   End
   Begin VB.ComboBox NameCmb 
      Height          =   315
      Left            =   240
      TabIndex        =   2
      Top             =   3600
      Width           =   2775
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Users"
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   2640
      Value           =   -1  'True
      Width           =   2295
   End
   Begin ComctlLib.ListView SecurityList 
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   5895
      _ExtentX        =   10393
      _ExtentY        =   3831
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label Label3 
      Caption         =   "Annotation Security:"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label Label2 
      Caption         =   "Access Level:"
      Height          =   255
      Left            =   3360
      TabIndex        =   5
      Top             =   3360
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Name:"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   3360
      Width           =   1335
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Add_Click()
   Dim securityItem As ListItem
   If NameCmb.Text = "" Then Exit Sub
   Set securityItem = SecurityList.ListItems.Add(, , NameCmb.Text)
   accesslevel = AccessLevelCmb.Text
   If accesslevel = "" Then
      accesslevel = "None"
   End If
   securityItem.SubItems(1) = accesslevel
   Dim granteeObj As Object
   If Option1(0).Value Then  ' user
      Set granteeObj = Form1.objLibrary.GetObject(idmObjTypeUser, NameCmb.Text)
   Else
      Set granteeObj = Form1.objLibrary.GetObject(idmObjTypeGroup, NameCmb.Text)
   End If
   'AccessLevelCmb.List
   Form1.objAnno.Permissions.Add granteeObj, AccessLevelCmb.ItemData(AccessLevelCmb.ListIndex)
   Form1.objDocument.Save
End Sub

Private Sub Form_Activate()
    SecurityList.ColumnHeaders.Clear
    SecurityList.ListItems.Clear
    SecurityList.ColumnHeaders.Add , , "Name", SecurityList.Width / 2
    SecurityList.ColumnHeaders.Add , , "AccessLevel", SecurityList.Width / 2
    Dim securityItem As ListItem
    For i = 1 To Form1.objAnno.Permissions.Count
      accessName = Form1.objAnno.Permissions(i).GranteeName
      accesslevel = Form1.objAnno.Permissions(i).Access
      Select Case CInt(accesslevel)
      Case idmDSAccessNone
        accessLevelStr = "None"
      Case idmDSAccessViewer
        accessLevelStr = "Viewer"
      Case idmDSAccessAuthor
        accessLevelStr = "Author"
      Case idmDSAccessOwner
        accessLevelStr = "Owner"
      Case idmDSAccessAdmin
        accessLevelStr = "Admin"
      End Select
      Set securityItem = SecurityList.ListItems.Add(, , accessName)
      securityItem.SubItems(1) = accessLevelStr
    Next i
   
    AccessLevelCmb.Clear
    AccessLevelCmb.AddItem "None"
    AccessLevelCmb.AddItem "Viewer"
    AccessLevelCmb.AddItem "Author"
    AccessLevelCmb.AddItem "Owner"
    AccessLevelCmb.AddItem "Admin"
    AccessLevelCmb.ItemData(0) = idmDSAccessNone
    AccessLevelCmb.ItemData(1) = idmDSAccessViewer
    AccessLevelCmb.ItemData(2) = idmDSAccessAuthor
    AccessLevelCmb.ItemData(3) = idmDSAccessOwner
    AccessLevelCmb.ItemData(4) = idmDSAccessAdmin
    AccessLevelCmb.ListIndex = 1
End Sub

Private Sub NameCmb_DropDown()
    NameCmb.Clear
    Dim objUser As IDMObjects.User
    Dim objGroup As IDMObjects.Group
    If Option1(0).Value = True Then
    For i = 1 To Form1.objUsers.Count
        Set objUser = Form1.objUsers(i)
        NameCmb.AddItem objUser.Name
    Next
    Else
    For i = 1 To Form1.objGroups.Count
        Set objGroup = Form1.objGroups(i)
        NameCmb.AddItem objGroup.Name
    Next
    End If
End Sub

Private Sub Remove_Click()
    Dim selectedItem As ListItem
    Set selectedItem = SecurityList.selectedItem
    If Not selectedItem Is Nothing Then
        Form1.objAnno.Permissions.Remove selectedItem.Index
        SecurityList.ListItems.Remove selectedItem.Index
        Form1.objDocument.Save
    End If
End Sub
