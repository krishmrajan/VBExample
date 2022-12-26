VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Create Annotation"
   ClientHeight    =   2850
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4710
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   2850
   ScaleWidth      =   4710
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox AnnoClasses 
      Height          =   2205
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   3015
   End
   Begin VB.CommandButton Cancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3600
      TabIndex        =   2
      Top             =   840
      Width           =   975
   End
   Begin VB.CommandButton OK 
      Caption         =   "OK"
      Height          =   375
      Left            =   3600
      TabIndex        =   1
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Select Annotation Type:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cancel_Click()
    Form2.Hide
End Sub

Private Sub Form_Load()
    ' Add the Annotation class names to the AnnoClasses listbox
    Dim objClasses As idmobjects.ObjectSet
    Set objClasses = Form1.objLibrary.FilterClassDescriptions(idmObjTypeAnnotation)
    For Each objAnnoClass In objClasses
        AnnoClasses.AddItem objAnnoClass.Label
    Next
End Sub

Private Sub OK_Click()
    Dim objAnno As idmobjects.Annotation
    On Error GoTo ErrorHandler
    Dim Class As Variant
    Class = AnnoClasses
    Set objAnno = Form1.objDocument.CreateAnnotation(Form1.CurrentPage, Class)
    If objAnno.ShowPropertiesDialog = idmDialogExitCancel Then
        objAnno.Delete
    Else
        Form1.ShowAnnos (Form1.CurrentPage)
    End If
    Form2.Hide
ErrorHandler:
    ShowError
End Sub


