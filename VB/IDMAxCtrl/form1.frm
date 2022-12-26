VERSION 5.00
Object = "{2A601BA0-B880-11CF-8185-444553540000}#3.0#0"; "FnList.ocx"
Object = "{A9983B40-CE52-11CF-AE75-00A0248802BA}#3.0#0"; "fnviewer.ocx"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "IDM Demo"
   ClientHeight    =   9240
   ClientLeft      =   7470
   ClientTop       =   3720
   ClientWidth     =   7410
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9240
   ScaleWidth      =   7410
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox ShowAnnotations 
      Caption         =   "Show Annotations"
      Enabled         =   0   'False
      Height          =   252
      Left            =   480
      TabIndex        =   15
      Top             =   6000
      Width           =   1695
   End
   Begin VB.Frame FrameAnno 
      Height          =   2175
      Left            =   120
      TabIndex        =   9
      Top             =   6000
      Width           =   7095
      Begin VB.CommandButton Approve 
         Caption         =   "Approve"
         Enabled         =   0   'False
         Height          =   252
         Left            =   5880
         TabIndex        =   14
         Top             =   480
         Width           =   1092
      End
      Begin VB.CommandButton Reject 
         Caption         =   "Reject"
         Enabled         =   0   'False
         Height          =   252
         Left            =   5880
         TabIndex        =   13
         Top             =   840
         Width           =   1092
      End
      Begin VB.CommandButton AddNote 
         Caption         =   "Add Note"
         Enabled         =   0   'False
         Height          =   252
         Left            =   5880
         TabIndex        =   12
         Top             =   1200
         Width           =   1092
      End
      Begin VB.CommandButton Highlight 
         Caption         =   "Highlight"
         Enabled         =   0   'False
         Height          =   252
         Left            =   5880
         TabIndex        =   11
         Top             =   1560
         Width           =   1092
      End
      Begin IDMListView.IDMListView IDMListView2 
         Height          =   1575
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   5655
         _Version        =   196608
         _ExtentX        =   9975
         _ExtentY        =   2778
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
         Enabled         =   0   'False
         Appearance      =   1
         _ColumnHeaders  =   "Form1.frx":0000
      End
   End
   Begin VB.CommandButton CloseBtn 
      Caption         =   "Close"
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   8520
      Width           =   1095
   End
   Begin IDMListView.IDMListView IDMListView1 
      Height          =   1335
      Left            =   120
      TabIndex        =   7
      Top             =   1080
      Width           =   7095
      _Version        =   196608
      _ExtentX        =   12515
      _ExtentY        =   2355
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
      ShowAnnotations =   -1  'True
      _ColumnHeaders  =   "Form1.frx":0018
   End
   Begin IDMViewerCtrl.IDMViewerCtrl IDMViewerCtrl1 
      Height          =   3255
      Left            =   120
      TabIndex        =   6
      Top             =   2520
      Width           =   7095
      _Version        =   196608
      _ExtentX        =   12515
      _ExtentY        =   5741
      _StockProps     =   161
      Appearance      =   1
      SystemType      =   14318
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Find Now"
      Enabled         =   0   'False
      Height          =   372
      Left            =   5280
      TabIndex        =   4
      Top             =   240
      Width           =   1092
   End
   Begin VB.ComboBox cmbDocClasses 
      Height          =   288
      Left            =   2400
      TabIndex        =   3
      Text            =   "Combo2"
      Top             =   600
      Width           =   1932
   End
   Begin VB.ComboBox cmbLibraries 
      Height          =   288
      Left            =   2400
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   120
      Width           =   1932
   End
   Begin VB.Label DocumentID 
      Height          =   255
      Left            =   1800
      TabIndex        =   5
      Top             =   8640
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Select Document Type:"
      Height          =   252
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   2172
   End
   Begin VB.Label Label1 
      Caption         =   "Select Library:"
      Height          =   252
      Left            =   120
      TabIndex        =   0
      Top             =   144
      Width           =   2172
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim oNeighborhood As New IDMObjects.Neighborhood
Dim oLibraries As IDMObjects.ObjectSet
Dim oDocument As IDMObjects.Document
Dim oLibrary As IDMObjects.Library
Dim clsQuery As New clsSimpleQuery

' Handle the logon, catch any errors here
Private Sub MyLogon(oLibrary As IDMObjects.Library)
    On Error GoTo Errorhandler  ' Enable error-handling routine.
    If Not (oLibrary.GetState(idmLibraryLoggedOn)) Then
        oLibrary.Logon "", "", "", idmLogonOptWithUI
    End If
    Exit Sub        ' Exit to avoid handler.
    
Errorhandler:
    MsgBox Err.Description & Err.Number
End Sub

Private Sub CloseBtn_Click()
    Unload Me
End Sub

' When user picks a library, get logged on and populate
' the document class combo box
Private Sub cmbLibraries_Click()
    Dim oLib As IDMObjects.Library
    cmbDocClasses.Clear
    
    For Each oLib In oLibraries
        If oLib.Label = cmbLibraries.Text Then
            Exit For
        End If
    Next
    ' make sure we are logged on to the selected Library
    MyLogon oLib
    
    'make sure the user didn't cancel out the logon
    If Not (oLib.GetState(idmLibraryLoggedOn)) Then
        Exit Sub
    End If
    'enable find button
    Command1.Enabled = True
    'clear the list view items
    IDMListView1.ClearItems
    'clear viewer
    IDMViewerCtrl1.Clear
    'clear annotation list
    IDMListView2.ClearItems
    'blank out the document id
    DocumentID.Caption = ""
    
    Set oLibrary = oLib  ' global for everyone to flex
    
    ' Fill the combo box with document class description names
    Dim oClass As IDMObjects.ClassDescription
    Dim oDocClasses As IDMObjects.ObjectSet
    Set oDocClasses = oLibrary.FilterClassDescriptions(idmObjTypeDocument)
    For Each oClass In oDocClasses
        cmbDocClasses.AddItem oClass.Name
    Next
End Sub

' Fire off a query to populate our IDMListView
Private Sub Command1_Click()
    Dim cHeadings As New Collection
    Dim sWhere As String
    
    ' create the where clause
    Select Case oLibrary.SystemType
    Case idmSysTypeDS
        sWhere = "idmDocProtected = 'Yes'" ' exclude external docs
        If cmbDocClasses <> "" Then
            sWhere = sWhere & " AND idmDocType = '" & cmbDocClasses & "'"
        End If
    Case idmSysTypeIS
        'Since we're looking at annotations, just grab images
        sWhere = "F_DOCTYPE = 'IMAGE'"
        If cmbDocClasses <> "" Then
            sWhere = sWhere & " AND F_DOCCLASSNAME = '" & cmbDocClasses & "'"
        End If
    End Select
    
    Call clsQuery.BindToLib(oLibrary, cHeadings)
    MousePointer = vbHourglass
    Call clsQuery.ExecQuery(IDMListView1, sWhere, "", 20)
    MousePointer = vbArrow
End Sub

' On start-up, populate the combo box with candidate libraries
Private Sub Form_Load()
    Set oLibraries = oNeighborhood.Libraries
    Dim oLib As IDMObjects.Library
    cmbLibraries.Clear
    cmbDocClasses.Clear
    ' Form1.Height = 2832
    ' For now, restrict this to libs which support annotations
    For Each oLib In oLibraries
        'If oLib.SystemType = idmSysTypeIS Then
         '   If oLib.Supports(idmSupportsAnnotations) Then
                cmbLibraries.AddItem oLib.Label
         '   End If
        'End If
    Next
    
    If cmbLibraries.ListCount <= 0 Then
        MsgBox "There is no IS library available. This demo works only for IS library supporting annotations."
        Unload Me
    End If
End Sub
' Log off from libraries on termination
Private Sub Form_Unload(Cancel As Integer)
    Dim oLib As IDMObjects.Library
    For Each oLib In oLibraries
        If oLib.GetState(idmLibraryLoggedOn) Then
            oLib.Logoff
        End If
    Next
End Sub

' Double click to view the document
Private Sub IDMListView1_DblClick()
    If Not IDMListView1.SelectedItem Is Nothing Then
        Set oDocument = IDMListView1.SelectedItem
        IDMListView2.ClearItems
        IDMViewerCtrl1.Document = oDocument
        
        'supports annotations
        If IDMViewerCtrl1.IsOperationSupported(idmOpAnnotations) Then
            ShowAnnotations.Enabled = True
        Else
            ShowAnnotations.Enabled = False
        End If
        
        ShowAnnotations_Click
        ' Form1.Height = 8760
        DocumentID = oDocument.Name
    End If
End Sub
' Double click on annotation => advance Viewer to correct page
Private Sub IDMListView2_DblClick()
    If Not IDMListView2.SelectedItem Is Nothing Then
        Dim oAnno As IDMObjects.Annotation
        Set oAnno = IDMListView2.SelectedItem
        IDMViewerCtrl1.PageNumber = oAnno.Properties("F_PAGENUMBER").Value
    End If
End Sub
' Populate the annotations list for selected document
Private Sub ShowAnnotations_Click()
    If (ShowAnnotations.Value = vbChecked) And IDMViewerCtrl1.IsOperationSupported(idmOpAnnotations) Then
        IDMViewerCtrl1.ShowAnnotations = True
        If oDocument.GetState(idmDocAnnotated) Then
            Dim oAnnos As IDMObjects.ObjectSet
            Set oAnnos = oDocument.Annotations
            If oAnnos.Count <> 0 Then
                IDMListView2.AddItems oAnnos, -1
            End If
        End If
        
        Approve.Enabled = True
        Reject.Enabled = True
        AddNote.Enabled = True
        Highlight.Enabled = True
        IDMListView2.Enabled = True
    Else
        IDMViewerCtrl1.ShowAnnotations = False
        IDMListView2.ClearItems
        
        Approve.Enabled = False
        Reject.Enabled = False
        AddNote.Enabled = False
        Highlight.Enabled = False
        IDMListView2.Enabled = False
    End If
End Sub
' Subroutines for handling annotation creation...
Private Sub AddNote_Click()
    Dim oAnno As IDMObjects.Annotation
    pg = IDMViewerCtrl1.PageNumber
    Set oAnno = oDocument.CreateAnnotation(pg, "Text")
    oAnno.Properties("F_LEFT").Value = 0.5
    oAnno.Properties("F_TOP").Value = 0.5
    oAnno.Properties("F_HEIGHT").Value = 0.5
    oAnno.Properties("F_WIDTH").Value = 1.5
    oAnno.Properties("F_FORECOLOR").Value = vbBlue
    oAnno.Properties("F_BACKCOLOR").Value = vbWhite
    oAnno.ShowPropertiesDialog
    oDocument.Save
    IDMListView2.AddItem oAnno, -1
End Sub

Private Sub MyCreateStamp(AnnoText As String)
    Dim oAnno As IDMObjects.Annotation
    pg = IDMViewerCtrl1.PageNumber
    Set oAnno = oDocument.CreateAnnotation(pg, "Stamp")
    oAnno.Properties("F_TEXT").Value = AnnoText
    oAnno.Properties("F_HASBORDER").Value = True
    oDocument.Save
    IDMListView2.AddItem oAnno, -1
End Sub
Private Sub Highlight_Click()
    IDMViewerCtrl1.LeftButtonAction = idmActionHighlight
    IDMViewerCtrl1.Document.Save
End Sub

Private Sub Reject_Click()
    MyCreateStamp "Reject"
End Sub
Private Sub Approve_Click()
    MyCreateStamp "Approve"
End Sub

Private Sub IDMViewerCtrl1_RequestDocumentClose(ByVal Document As Object, Cancel As Boolean)
    Document.Save
End Sub


