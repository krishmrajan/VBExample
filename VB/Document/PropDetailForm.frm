VERSION 5.00
Begin VB.Form PropDetailForm 
   Caption         =   "Property Details"
   ClientHeight    =   6315
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7785
   LinkTopic       =   "Form2"
   ScaleHeight     =   6315
   ScaleWidth      =   7785
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton BtnChoices 
      Caption         =   "Choices"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3360
      TabIndex        =   25
      Top             =   5880
      Width           =   1455
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   3120
      Locked          =   -1  'True
      TabIndex        =   22
      Text            =   "Combo1"
      Top             =   5280
      Width           =   3495
   End
   Begin VB.CommandButton Done 
      Caption         =   "Done"
      Height          =   375
      Left            =   5400
      TabIndex        =   20
      Top             =   5880
      Width           =   1215
   End
   Begin VB.TextBox Text10 
      Height          =   375
      Left            =   3120
      Locked          =   -1  'True
      TabIndex        =   19
      Text            =   "Text10"
      Top             =   4800
      Width           =   3495
   End
   Begin VB.TextBox Text9 
      Height          =   375
      Left            =   3120
      Locked          =   -1  'True
      TabIndex        =   17
      Text            =   "Text9"
      Top             =   4320
      Width           =   3495
   End
   Begin VB.TextBox Text8 
      Height          =   375
      Left            =   3120
      Locked          =   -1  'True
      TabIndex        =   16
      Text            =   "Text8"
      Top             =   3840
      Width           =   3495
   End
   Begin VB.TextBox Text7 
      Height          =   375
      Left            =   3120
      Locked          =   -1  'True
      TabIndex        =   15
      Text            =   "Text7"
      Top             =   3360
      Width           =   3495
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   3120
      Locked          =   -1  'True
      TabIndex        =   14
      Text            =   "Text6"
      Top             =   2880
      Width           =   3495
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   3120
      Locked          =   -1  'True
      TabIndex        =   13
      Text            =   "Text5"
      Top             =   2400
      Width           =   3495
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   3120
      Locked          =   -1  'True
      TabIndex        =   12
      Text            =   "Text4"
      Top             =   1920
      Width           =   3495
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   3120
      Locked          =   -1  'True
      TabIndex        =   11
      Text            =   "Text3"
      Top             =   1440
      Width           =   3495
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   3120
      Locked          =   -1  'True
      TabIndex        =   10
      Text            =   "Text2"
      Top             =   960
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   3120
      Locked          =   -1  'True
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   480
      Width           =   3495
   End
   Begin VB.Label Label13 
      Caption         =   "Label13"
      Height          =   255
      Left            =   1920
      TabIndex        =   24
      Top             =   120
      Width           =   4695
   End
   Begin VB.Label Label12 
      Caption         =   "Property Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label11 
      Caption         =   "Label11"
      Height          =   375
      Left            =   240
      TabIndex        =   21
      Top             =   5280
      Width           =   2775
   End
   Begin VB.Label Label10 
      Caption         =   "Label10"
      Height          =   255
      Left            =   240
      TabIndex        =   18
      Top             =   4920
      Width           =   2775
   End
   Begin VB.Label Label9 
      Caption         =   "Label9"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   4380
      Width           =   2655
   End
   Begin VB.Label Label8 
      Caption         =   "Label8"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   3900
      Width           =   2655
   End
   Begin VB.Label Label7 
      Caption         =   "Label7"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   3420
      Width           =   2655
   End
   Begin VB.Label Label6 
      Caption         =   "Label6"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   2940
      Width           =   2655
   End
   Begin VB.Label Label5 
      Caption         =   "Label5"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   2460
      Width           =   2655
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   1980
      Width           =   2655
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1500
      Width           =   2655
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   1020
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   540
      Width           =   2655
   End
End
Attribute VB_Name = "PropDetailForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub BtnChoices_Click()
Dim oPropDesc As IDMObjects.PropertyDescription
Dim oChoices As IDMObjects.Choices
Dim oChoice As IDMObjects.Choice
Dim fChoices As New frmChoices
Set oPropDesc = MainForm.oCurrentLibrary.GetObject(idmObjTypePropDesc, PropertiesForm.PropName, idmObjTypeDocument)
Set oChoices = oPropDesc.Choices
Call oChoices.Paging.NextPage(0, idmForward)
For Each oChoice In oChoices
    fChoices.List1.AddItem (oChoice.Value)
Next
fChoices.Show vbModal
End Sub

Private Sub Done_Click()
    PropDetailForm.Hide
End Sub

Private Sub Picture1_Click()

End Sub

