VERSION 5.00
Object = "{2A601BA0-B880-11CF-8185-444553540000}#1.0#0"; "FNLIST.OCX"
Object = "{54E6FE61-B93A-11CF-8185-444553540000}#1.0#0"; "FNTREE.OCX"
Object = "{A9983B40-CE52-11CF-AE75-00A0248802BA}#1.0#0"; "FNVIEWER.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Begin VB.Form BenefitForm 
   Caption         =   "Benefits"
   ClientHeight    =   8595
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   10725
   LinkTopic       =   "Form1"
   Picture         =   "benefitsform.frx":0000
   ScaleHeight     =   8595
   ScaleWidth      =   10725
   StartUpPosition =   3  'Windows Default
   Begin IDMViewerCtrl.IDMViewerCtrl ViewerCtrl1 
      Height          =   6135
      Left            =   4320
      TabIndex        =   4
      Top             =   120
      Width           =   6255
      _Version        =   65536
      _ExtentX        =   11033
      _ExtentY        =   10821
      _StockProps     =   161
      SystemType      =   -4476
   End
   Begin IDMListView.IDMListView ListView1 
      Height          =   2895
      Left            =   120
      TabIndex        =   3
      Top             =   3360
      Width           =   4095
      _Version        =   65536
      _ExtentX        =   7223
      _ExtentY        =   5106
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
      _ColumnHeaders  =   "benefitsform.frx":9C844
   End
   Begin IDMTreeView.IDMTreeView TreeView1 
      Height          =   2895
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4095
      _Version        =   65536
      _ExtentX        =   7223
      _ExtentY        =   5106
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
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6120
      Top             =   8280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327681
   End
   Begin VB.CommandButton Checkin 
      Caption         =   "Save"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6600
      Width           =   1815
   End
   Begin VB.CommandButton Checkout 
      Caption         =   "Update"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6600
      Width           =   1815
   End
End
Attribute VB_Name = "BenefitForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim oDoc As IDMObjects.Document
Dim oFolder As IDMObjects.Folder
Dim Location As String
Private Sub RefreshListView(poFolder As IDMObjects.Folder)
    Dim oContents As IDMObjects.ObjectSet
    ListView1.ClearItems
    Set oContents = poFolder.GetContents(idmFolderContentDocument)
    If Not oContents Is Nothing And oContents.Count > 0 Then
       ListView1.AddItems oContents, -1
    End If
    Set oContents = poFolder.GetContents(idmFolderContentStoredSearch)
    If Not oContents Is Nothing And oContents.Count > 0 Then
        ListView1.AddItems oContents, -1
    End If

End Sub

Private Sub Checkin_Click()
    oDoc.Version.Checkin
    Checkout.Enabled = True
    Checkin.Enabled = False
    DocID = oDoc.ID
    Call RefreshListView(oFolder)
    Set oDoc = goDSLib.GetObject(idmObjTypeDocument, DocID)
    Me.ViewerCtrl1.Document = oDoc
End Sub

Private Sub Checkout_Click()
'    CommonDialog1.DialogTitle = "Select Location to Checkout To"
'    CommonDialog1.ShowOpen
    Dim checkedout As Boolean
    checkedout = False
    If Not oDoc.GetState(idmDocCheckedout) Then
        oDoc.Version.Checkout Location
        checkedout = True
    End If
    oDoc.Launch idmDocLaunchNativeApplication
    Checkin.Enabled = True
    Checkout.Enabled = True
    
    If checkedout Then
        ' do a bunch of refresh junk to work around bug in Document
        DocID = oDoc.ID
        oFolder.Refresh idmFolderRefreshContents
        Call RefreshListView(oFolder)
        Set oDoc = goDSLib.GetObject(idmObjTypeDocument, DocID)
    End If
End Sub

Private Sub Form_Load()
    Dim oFolder As IDMObjects.Folder
    Dim sFolderId As String
    sFolderId = gfSettings.txtBenFolder
    Set oFolder = goDSLib.GetObject(idmObjTypeFolder, sFolderId)
    TreeView1.ClearItems
    TreeView1.AddRootItem oFolder, True
End Sub

Private Sub ListView1_DblClick()
    ViewerCtrl1.Document = oDoc
End Sub

Private Sub ListView1_ItemSelectChange(ByVal Item As Object, ByVal Key As Long, ByVal Selected As Boolean, ByVal ObjType As IDMListView.idmObjectType)
    If Not Item Is Nothing Then
        Set oDoc = Item
        Checkout.Enabled = True
        Checkin.Enabled = oDoc.GetState(idmDocCheckedout)
    Else
        Checkout.Enabled = False
        Checkin.Enabled = False
    End If
End Sub

Private Sub TreeView1_ItemSelectChange(ByVal Item As Object, ByVal ObjType As IDMTreeView.idmObjectType)
    Set oFolder = Item
    Call RefreshListView(oFolder)
End Sub
