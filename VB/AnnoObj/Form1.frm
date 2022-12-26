VERSION 5.00
Object = "{2A601BA0-B880-11CF-8185-444553540000}#2.0#0"; "fnlist.ocx"
Begin VB.Form Form1 
   Caption         =   "Annotation Sample Application"
   ClientHeight    =   5544
   ClientLeft      =   1500
   ClientTop       =   1536
   ClientWidth     =   7908
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5544
   ScaleWidth      =   7908
   Begin IDMListView.IDMListView AnnoList 
      Height          =   4335
      Left            =   240
      TabIndex        =   12
      Top             =   1080
      Width           =   5655
      _Version        =   131072
      _ExtentX        =   9975
      _ExtentY        =   7646
      _StockProps     =   239
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   1
      _ColumnHeaders  =   "Form1.frx":0000
   End
   Begin VB.CommandButton ShowPage 
      Caption         =   "Show Property Page"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6000
      TabIndex        =   11
      Top             =   3840
      Width           =   1695
   End
   Begin VB.CommandButton ShowProperties 
      Caption         =   "Show Properties"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6000
      TabIndex        =   10
      Top             =   3360
      Width           =   1695
   End
   Begin VB.CommandButton ModifySecurity 
      Caption         =   "Modify Security"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6000
      TabIndex        =   9
      Top             =   2880
      Width           =   1695
   End
   Begin VB.CommandButton SaveAnnotations 
      Caption         =   "Save Annotations"
      Enabled         =   0   'False
      Height          =   370
      Left            =   6000
      TabIndex        =   7
      Top             =   1320
      Width           =   1695
   End
   Begin VB.CommandButton Exit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   6240
      TabIndex        =   6
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton PrevPage 
      Caption         =   "<< Prev Page"
      Enabled         =   0   'False
      Height          =   255
      Left            =   3480
      TabIndex        =   5
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton NextPage 
      Caption         =   "Next Page >>"
      Enabled         =   0   'False
      Height          =   255
      Left            =   4680
      TabIndex        =   4
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton SelectDocument 
      Caption         =   "Select Document..."
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton CreateAnnotation 
      Caption         =   "Create Annotation"
      Enabled         =   0   'False
      Height          =   370
      Left            =   6000
      TabIndex        =   1
      Top             =   840
      Width           =   1695
   End
   Begin VB.CommandButton DeleteAnnotation 
      Caption         =   "Delete "
      Enabled         =   0   'False
      Height          =   370
      Left            =   6000
      TabIndex        =   0
      Top             =   2400
      Width           =   1695
   End
   Begin VB.Label LibraryName 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   8
      Top             =   120
      Width           =   3375
   End
   Begin VB.Label DocumentID 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   840
      Width           =   3255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public objLibrary As idmobjects.Library
Public objCommonDialogs As New idmobjects.CommonDialogs
Public objDocument As idmobjects.Document
Public objAnno As idmobjects.Annotation
Public objUsers As idmobjects.ObjectSet
Public objGroups As idmobjects.ObjectSet
Public CurrentPage As Integer

Private Sub Exit_Click()
    Dim cmd As VbMsgBoxResult
    cmd = vbNo
    If Not objDocument Is Nothing Then
        If objDocument.GetState(idmDocAnnosModified) Then
            cmd = MsgBox("Save Modified Annotations?", vbYesNoCancel, Form1.Caption)
            If cmd = vbYes Then
                objDocument.Save
            End If
        End If
    End If
    If cmd <> vbCancel Then
        End
    End If
End Sub

Private Sub DeleteAnnotation_Click()
On Error GoTo errorhandler
    If Not AnnoList.SelectedItem Is Nothing Then
        Dim oAnno As idmobjects.Annotation
        Set oAnno = AnnoList.SelectedItem
        oAnno.Delete
        ShowAnnos (CurrentPage)
    End If
    Exit Sub
errorhandler:
    ShowError
End Sub


Private Sub Form_Unload(Cancel As Integer)
Set oErrManager = Nothing
Set objCommonDialogs = Nothing
End Sub

Private Sub ModifySecurity_Click()
    If AnnoList.SelectedItem Is Nothing Then
        MsgBox "Select an annotation"
    Else
        ' Hold onto the list of users and groups to increase
        ' performance on subsequent requests
        Form1.MousePointer = vbHourglass
        If Form1.objUsers Is Nothing Then
            Set Form1.objUsers = Form1.objLibrary.Users
        End If
        If Form1.objGroups Is Nothing Then
            Set Form1.objGroups = Form1.objLibrary.Groups
        End If
        Form1.MousePointer = vbDefault
        
        Set objAnno = AnnoList.SelectedItem
        If Form1.objLibrary.SystemType = idmSysTypeIS Then
            Form3.Show 1, Form1
        Else
            Form5.Show 1, Form1
        End If
    End If
End Sub

Private Sub SaveAnnotations_Click()
    objDocument.Save
End Sub

Private Sub SelectDocument_Click()
    Dim objDoc As idmobjects.Document
    Dim Operation As idmobjects.idmOperation
    If objLibrary Is Nothing Then
        Dim oHood As New idmobjects.Neighborhood
        Set objLibrary = oHood.DefaultLibrary
    End If
    objCommonDialogs.LookIn = objLibrary
    
    objCommonDialogs.Options = idmSelectHideDrives + _
        idmSelectHideOpenAs + idmSelectHideAdvanced
    objCommonDialogs.SelectDocument objDoc, Operation
    If Operation = idmOperationOpen Then
        Set objDocument = objDoc
        Set objLibrary = objDocument.Library
        
        ' Check to see if the selected Library supports creating annotations
        If objLibrary.Supports(idmSupportsAnnotations) Then
            LibraryName = "Library: " & objLibrary.Label
            CurrentPage = 1
            ShowAnnos (1)
            CreateAnnotation.Enabled = True
            DeleteAnnotation.Enabled = True
            SaveAnnotations.Enabled = True
            ModifySecurity.Enabled = True
            ShowProperties.Enabled = True
            ShowPage.Enabled = True
        Else
            MsgBox objLibrary.Label & " Library does not support annotations."
            Set objDocument = Nothing
            Set objLibrary = Nothing
            CreateAnnotation.Enabled = False
            DeleteAnnotation.Enabled = False
            SaveAnnotations.Enabled = False
            ModifySecurity.Enabled = False
            ShowProperties.Enabled = False
            ShowPage.Enabled = False
        End If
        Set objAnno = Nothing
        Set objUsers = Nothing
        Set objGroups = Nothing
    End If
End Sub

Private Sub CreateAnnotation_Click()
    If objDocument Is Nothing Then
        MsgBox "Please select a document"
    End If
    Form2.Show 1, Form1
End Sub

Public Sub ShowAnnos(ByVal PageNum As Integer)
 On Error GoTo errorhandler
 
    AnnoList.ClearItems
    DocumentID = "Document ID: " & objDocument.ID & " Page: " & PageNum
    If objDocument.GetState(idmDocAnnotated) Then
        Set objAnnos = objDocument.GetPageAnnotations(PageNum)
        AnnoList.AddItems objAnnos, -1
        PrevPage.Enabled = False
    End If
    If PageNum = objDocument.PageCount Then
        NextPage.Enabled = False
    Else
        NextPage.Enabled = True
    End If
    If PageNum = 1 Then
        PrevPage.Enabled = False
    Else
        PrevPage.Enabled = True
    End If
    CurrentPage = PageNum
    Exit Sub
errorhandler:
    ShowError
End Sub

Private Sub NextPage_Click()
    ShowAnnos (CurrentPage + 1)
End Sub

Private Sub PrevPage_Click()
    ShowAnnos (CurrentPage - 1)
End Sub

Private Sub ShowPage_Click()
    If AnnoList.SelectedItem Is Nothing Then
        MsgBox "Select an annotation"
    Else
        Set objAnno = AnnoList.SelectedItem
        objAnno.ShowPropertiesDialog
    End If
End Sub

Private Sub ShowProperties_Click()
    If AnnoList.SelectedItem Is Nothing Then
        MsgBox "Select an annotation"
    Else
        Set objAnno = AnnoList.SelectedItem
        Form4.Show 1, Form1
    End If
End Sub
