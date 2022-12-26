VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Begin VB.Form AddDocumentForm 
   Caption         =   "Add New Document"
   ClientHeight    =   4425
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6495
   LinkTopic       =   "Form2"
   ScaleHeight     =   4425
   ScaleWidth      =   6495
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   2160
      TabIndex        =   6
      Text            =   "Combo1"
      Top             =   360
      Width           =   2895
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   360
      Top             =   3720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327681
   End
   Begin VB.CommandButton Next 
      Caption         =   "Next >>"
      Height          =   375
      Left            =   3600
      TabIndex        =   5
      Top             =   3840
      Width           =   1335
   End
   Begin VB.CommandButton Cancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5040
      TabIndex        =   4
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton RemoveFile 
      Caption         =   "Remove File"
      Height          =   495
      Left            =   5160
      TabIndex        =   3
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton AddFile 
      Caption         =   "Add File..."
      Height          =   495
      Left            =   5160
      TabIndex        =   2
      Top             =   1680
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   2205
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   4935
   End
   Begin VB.Label Label1 
      Caption         =   "Document Class:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   360
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "Select Files to Save:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   1935
   End
End
Attribute VB_Name = "AddDocumentForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim MyFiles() As Variant

Private Sub AddFile_Click()
    CommonDialog1.ShowOpen
    List1.AddItem CommonDialog1.filename
End Sub

Private Sub Cancel_Click()
    AddDocumentForm.Hide
End Sub

Private Sub Form_Activate()
    ' Add the Document Classes defined for this Library to
    ' the drop down combo box
    Dim oClasses As IDMObjects.ObjectSet
    Dim oClass As IDMObjects.ClassDescription
    Set oClasses = MainForm.oCurrentLibrary.FilterClassDescriptions(idmObjTypeDocument)
    For Each oClass In oClasses
        Combo1.AddItem oClass.Name
    Next
    Combo1.ListIndex = 0
End Sub

Private Sub Next_Click()
    ' Add the list of files to an Array
    ReDim MyFiles(1 To List1.ListCount) As Variant
    ii = 1
    While ii <= List1.ListCount
    MyFiles(ii) = List1.List(ii - 1)
        ii = ii + 1
    Wend
            
    ' Create a new document object
    Dim oDoc As IDMObjects.Document
    ClassName = Combo1
    Set oDoc = MainForm.oCurrentLibrary.CreateObject(idmObjTypeDocument, ClassName)
    
    On Error GoTo ErrorHandler
    
    ' Display a message box to let the user look at property details
    ' before adding with the wizard
    If MsgBox("Would you like to display detailed property information before adding?", _
               vbYesNo, "Adding Document") = vbYes Then
        PropertiesForm.ShowPropertiesFor oDoc, AddDocumentForm, "Next", "New Document"
    End If
    
    ' Display the add wizard
    oDoc.SaveNew MyFiles, idmDocSaveNewKeep + idmDocSaveNewWithUIWizard
    
    ' Check to see if a folder is selected in ListView or TreeView
    If Not MainForm.oFolder Is Nothing Then
        MainForm.oFolder.File oDoc
    End If
    Exit Sub
ErrorHandler:
   MainForm.ShowError
End Sub

Private Sub RemoveFile_Click()
    If List1.ListIndex <> -1 Then
        List1.RemoveItem List1.ListIndex
    End If
End Sub
