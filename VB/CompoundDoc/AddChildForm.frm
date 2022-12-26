VERSION 5.00
Object = "{2A601BA0-B880-11CF-8185-444553540000}#2.0#0"; "FnList.ocx"
Object = "{54E6FE61-B93A-11CF-8185-444553540000}#2.0#0"; "FnTree.ocx"
Begin VB.Form AddChildForm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add Child documents"
   ClientHeight    =   2250
   ClientLeft      =   825
   ClientTop       =   2760
   ClientWidth     =   6780
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2250
   ScaleWidth      =   6780
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save Link"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5280
      TabIndex        =   11
      Top             =   720
      Width           =   1455
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "&Remove Child"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5280
      TabIndex        =   10
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancelAdd 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5280
      TabIndex        =   8
      Top             =   1680
      Width           =   1455
   End
   Begin IDMListView.IDMListView ChildrenListView 
      Height          =   1695
      Left            =   120
      TabIndex        =   6
      Top             =   360
      Width           =   4935
      _Version        =   131072
      _ExtentX        =   8705
      _ExtentY        =   2990
      _StockProps     =   239
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   1
      _ColumnHeaders  =   "AddChildForm.frx":0000
   End
   Begin VB.TextBox PathField 
      Height          =   285
      Left            =   1680
      TabIndex        =   5
      Top             =   4200
      Width           =   495
   End
   Begin VB.CheckBox chkAbsolutePath 
      Caption         =   "Absolute Path"
      Height          =   255
      Left            =   720
      TabIndex        =   4
      Top             =   3480
      Width           =   1575
   End
   Begin VB.CommandButton AddChildren 
      Caption         =   "&Add Child"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   5280
      TabIndex        =   3
      Top             =   240
      Width           =   1455
   End
   Begin VB.CommandButton CancelAdd 
      Caption         =   "Done"
      Height          =   495
      Index           =   0
      Left            =   8520
      TabIndex        =   2
      Top             =   9240
      Width           =   1455
   End
   Begin IDMListView.IDMListView AddChildListView 
      Height          =   975
      Left            =   3960
      TabIndex        =   1
      Top             =   3360
      Width           =   855
      _Version        =   131072
      _ExtentX        =   1508
      _ExtentY        =   1720
      _StockProps     =   239
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   1
      _ColumnHeaders  =   "AddChildForm.frx":0018
   End
   Begin IDMTreeView.IDMTreeView AddChildTreeView 
      Height          =   855
      Left            =   2520
      TabIndex        =   0
      Top             =   3360
      Width           =   735
      _Version        =   131072
      _ExtentX        =   1296
      _ExtentY        =   1508
      _StockProps     =   228
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   1
   End
   Begin VB.Frame Frame1 
      Caption         =   "Children Documents"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   0
      TabIndex        =   7
      Top             =   120
      Width           =   5175
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   4200
      Width           =   1335
   End
End
Attribute VB_Name = "AddChildForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oAddChildNeighborhood As New IDMObjects.Neighborhood
Dim oAddChildSelFolder As IDMObjects.Folder
Dim oParentDoc As IDMObjects.Document
Dim oChildDoc As IDMObjects.Document
Dim oCmnDlg As IDMObjects.CommonDialogs
Public Sub SpecifyParent(oDocument As IDMObjects.Document)
    Set oParentDoc = oDocument
    LoadChildrenListView
End Sub
Private Sub LoadChildrenListView()
    Dim oLinks As IDMObjects.Links
    Dim oLink As IDMObjects.Link
  
    On Error GoTo ErrHandler
    
    ChildrenListView.ClearItems
    
    Set oLinks = oParentDoc.Compound.Children
    For Each oLink In oLinks
        Dim oChild As IDMObjects.Document
        Set oChild = oLink.Child
        ChildrenListView.AddItem oChild, 0
    Next oLink
    
    Exit Sub

ErrHandler:
   If Err.Number <> 1 Then
      
   Else
      MsgBox Err.Description, vbCritical, "Load Child Document"
   End If
End Sub
Private Sub AddChildren_Click(Index As Integer)
    Dim oCommDialog As IDMObjects.CommonDialogs
    'Dim oChildDoc As IDMObjects.Document
    Dim oLink As IDMObjects.Link
    Dim oLinks As IDMObjects.Links
    Dim oLinkProp As IDMObjects.Document
    Dim oVer As IDMObjects.Version
    Dim oObject As Object
    
    On Error GoTo ErrHandler
    
    'use fileNET CommonDialog to get a child document object
    Set oCommDialog = CreateObject("IDMObjects.CommonDialogs")
    Call oCommDialog.SelectDocument(oChildDoc, idmOperationOpen)
    
    MainForm.RefreshListView
    'set link option
    frmLinkOption.Show vbModal, Me
    

    Set oLinks = oParentDoc.Compound.Children
    
    If oLinks.GetState(idmLinksEditable) = True Then
        Set oLink = New IDMObjects.Link
        oLink.Parent = oDocument
        oLink.Child = oChildDoc
        oLink.ClassID = idmStaticLink
        'Set oLinkProp = oLink.Properties("idmLinkChild")
   
         'set link option and path
         Select Case sLinkOption
                Case "Weak"
                     oLink.Properties("idmLinkStrength") = 1
                Case "Strong"
                     oLink.Properties("idmLinkStrength") = 2
         End Select
         Select Case sPath
                Case "Relative"
                     oLink.Properties("idmLinkUsesRelativePath") = 1
                     oLink.Properties("idmLinkRelativePath") = Trim(frmLinkOption.txtPath.Text)
                Case "Absolute"
                    oLink.Properties("idmLinkUsesRelativePath") = 2
                    oLink.Properties("idmLinkAbsolutePath") = Trim(frmLinkOption.txtPath.Text)
         End Select
        
         'add to the collection
         'Set oLinks = oParentDoc.Compound.Children
         oLinks.Add oLink
    Else
          MsgBox "The link collection is not editable, a new  link is not added.", vbInformation, "Add New Links"
    End If
    'show the children
    Call LoadChildrenListView
   
   Exit Sub
ErrHandler:
   MsgBox Err.Description, vbCritical, "Add Child Document"
End Sub
Private Sub CancelAdd_Click(Index As Integer)
    Me.Hide
End Sub

Private Sub ChildrenListView_ItemSelectChange(ByVal Item As Object, ByVal Key As Long, ByVal Selected As Boolean, ByVal ObjType As IDMListView.idmObjectType)
   If ObjType = idmObjTypeDocument Then
        Set oChildDoc = Item
   Else
        Set oChildDoc = Nothing
   End If
End Sub

Private Sub cmdCancelAdd_Click()
   Unload Me
End Sub
Private Sub cmdRemove_Click()
    Dim oLink As IDMObjects.Link
    Dim oLinks As IDMObjects.Links
    Dim iIndex As Integer
    
    On Error GoTo ErrHandler
    
    'Set oLink = New idmObjects.Link
    'oLink.Parent = oDocument
    
    If oDocument.Compound.Children.Count > 0 Then
        Set oLinks = oParentDoc.Compound.Children
    Else
        MsgBox "Can'remove a parent link", vbInformation, "Remove Link"
        Exit Sub
    End If
    If oLinks.GetState(idmLinksEditable) = True Then
        iIndex = oLinks.Find(oChildDoc)
        If iIndex > 0 Then
           oLinks.Remove iIndex
        End If
    Else
        MsgBox "The link collection is not editable, the link is not removed.", vbInformation, "Remove Link"
    End If
    Call LoadChildrenListView
    Exit Sub
ErrHandler:
   MsgBox Err.Description, vbCritical, "Remove Link"
End Sub

Private Sub cmdSave_Click()
    'save parent document and link collection
    oDocument.Save
End Sub

Private Sub Form_Load()
    Call CenterForm(Me)
    AddChildTreeView.AddRootItem oAddChildNeighborhood, True
    PathField.Text = "C:\Temp"
End Sub
Private Sub AddChildTreeView_ItemSelectChange(ByVal Item As Object, ByVal ObjType As IDMTreeView.idmObjectType)
    AddChildListView.ClearItems
    If ObjType = idmObjTypeFolder Then
        Set oAddChildSelFolder = Item
        AddChildListView.AddItems oAddChildSelFolder.SubFolders, -1
        AddChildListView.AddItems oAddChildSelFolder.GetContents(idmFolderContentDocument), -1
        AddChildListView.AddItems oAddChildSelFolder.GetContents(idmFolderContentStoredSearch), -1
    Else
        Set oAddChildSelFolder = Nothing
    End If
End Sub
Private Sub AddChildListView_DblClick()
    If AddChildListView.SelectedItem.ObjectType = idmObjTypeFolder Then
        Dim oFolder As IDMObjects.Folder
        Set oFolder = AddChildListView.SelectedItem
        AddChildTreeView.SelectChildItem oFolder.Name
    End If
End Sub


