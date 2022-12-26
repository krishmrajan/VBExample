VERSION 5.00
Object = "{2A601BA0-B880-11CF-8185-444553540000}#3.0#0"; "fnlist.ocx"
Object = "{A9983B40-CE52-11CF-AE75-00A0248802BA}#3.0#0"; "fnviewer.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form ResumeForm 
   Caption         =   "Resumes"
   ClientHeight    =   7755
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   9540
   LinkTopic       =   "Form2"
   Picture         =   "BrowseNewResumeForm.frx":0000
   ScaleHeight     =   7755
   ScaleWidth      =   9540
   StartUpPosition =   3  'Windows Default
   Begin IDMViewerCtrl.IDMViewerCtrl ViewerCtrl1 
      Height          =   5415
      Left            =   120
      TabIndex        =   8
      Top             =   3000
      Width           =   6855
      _Version        =   196608
      _ExtentX        =   12091
      _ExtentY        =   9551
      _StockProps     =   161
      Appearance      =   1
      SystemType      =   -16522
   End
   Begin IDMListView.IDMListView ListView1 
      Height          =   1815
      Left            =   120
      TabIndex        =   7
      Top             =   960
      Width           =   6855
      _Version        =   196608
      _ExtentX        =   12091
      _ExtentY        =   3201
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
      _ColumnHeaders  =   "BrowseNewResumeForm.frx":9C844
   End
   Begin VB.CommandButton RequestInterview 
      Caption         =   "Request Interview..."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7200
      Picture         =   "BrowseNewResumeForm.frx":9C85C
      TabIndex        =   6
      Top             =   6600
      Width           =   1935
   End
   Begin VB.CommandButton AddNote 
      Caption         =   "Add Note..."
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7200
      Picture         =   "BrowseNewResumeForm.frx":9F3B0
      TabIndex        =   5
      Top             =   5475
      Width           =   1935
   End
   Begin VB.CommandButton NewResumes 
      Caption         =   "Find New Resumes"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   2775
   End
   Begin VB.CommandButton FindResume 
      Caption         =   "Find Resume..."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3300
      TabIndex        =   3
      Top             =   120
      Width           =   2775
   End
   Begin VB.CommandButton Highlight 
      Caption         =   "Highlight"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7200
      TabIndex        =   2
      Top             =   4365
      Width           =   1935
   End
   Begin VB.CommandButton AddNewResume 
      Caption         =   "Add New Resume..."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6480
      TabIndex        =   1
      Top             =   120
      Width           =   2895
   End
   Begin VB.CommandButton SaveNew 
      Caption         =   "Save New..."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7200
      TabIndex        =   0
      Top             =   3360
      Visible         =   0   'False
      Width           =   1935
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7200
      Top             =   9240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "ResumeForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim oFolder As IDMObjects.Folder
Dim oDocument As IDMObjects.Document

Private Sub FindResume_Click()
    frmFind.Show 1, ResumeForm
End Sub

Private Sub AddNote_Click()
    ViewerCtrl1.LeftButtonAction = idmActionStickyNote
End Sub

Private Sub SetLVHeaders(cHeadings As Collection, cPropNames As Collection)
Dim asClasses(1) As String
Dim oPropDesc As IDMObjects.PropertyDescription
asClasses(0) = gfSettings.txtResDocClass
Set goPropDescs = goISLib.FilterPropertyDescriptions(idmObjTypeDocument, _
    asClasses)
For Each oPropDesc In goPropDescs
    If Left(oPropDesc.Name, 2) <> "F_" Then
        Call ListView1.AddColumnHeader(goISLib, oPropDesc)
        cHeadings.Add (oPropDesc.Label)
        cPropNames.Add (oPropDesc.Name)
    End If
Next
Call ListView1.SwitchColumnHeaders(goISLib)
End Sub

Private Sub Form_Load()
Set gcHeadings = Nothing
Set gcPropNames = Nothing
Set gcHeadings = New Collection
Set gcPropNames = New Collection
Call SetLVHeaders(gcHeadings, gcPropNames)
End Sub


Private Sub NewResumes_Click()
    Dim sFolderId As String
    sFolderId = gfSettings.txtResFolder
    Set oFolder = goISLib.GetObject(idmObjTypeFolder, sFolderId)
    ListView1.ClearItems
    ListView1.AddItems oFolder.GetContents(idmFolderContentDocument), -1
    ListView1.View = idmViewReport
    ViewerCtrl1.Clear
End Sub

Private Sub Highlight_Click()
    ViewerCtrl1.LeftButtonAction = idmActionHighlight
End Sub

Private Sub SaveNew_Click()
    Dim MyFile As String
    Dim sClass As String
    MyFile = ViewerCtrl1.DocumentFilename
    Set oDocument = Nothing
    sClass = gfSettings.txtResDocClass
    Set oDocument = goISLib.CreateObject(idmObjTypeDocument, sClass)
    On Error GoTo ErrorHandler
    oDocument.SaveNew MyFile, idmDocSaveNewConfirmationUI
Exit Sub

ErrorHandler:
    MsgBox Err.Description
End Sub


Private Sub AddNewResume_Click()
    CommonDialog1.DialogTitle = "Select Resume to Add"
    CommonDialog1.ShowOpen
    ViewerCtrl1.DocumentFilename = CommonDialog1.FileName
    AddNote.Enabled = False
    Highlight.Enabled = False
    RequestInterview.Enabled = False
    NewResumes.Enabled = True
    FindResume.Enabled = True
    SaveNew.Visible = True
End Sub

Private Sub ListView1_DblClick()
    ViewerCtrl1.Document = oDocument
    If ViewerCtrl1.DocumentCategory = idmNativeImage Then
        AddNote.Enabled = True
        Highlight.Enabled = True
    Else
        AddNote.Enabled = False
        Highlight.Enabled = False
    End If
    RequestInterview.Enabled = True
End Sub

Private Sub ListView1_ItemSelectChange(ByVal Item As Object, ByVal Key As Long, ByVal Selected As Boolean, ByVal ObjType As IDMListView.idmObjectType)
    Set oDocument = Item
    AddNote.Enabled = True
    Highlight.Enabled = True
    NewResumes.Enabled = True
    FindResume.Enabled = True
    SaveNew.Visible = False
End Sub

Private Sub RequestInterview_Click()
    On Error GoTo ErrorHandler
    oDocument.Route
Exit Sub
    
ErrorHandler:
    MsgBox Err.Description
End Sub
