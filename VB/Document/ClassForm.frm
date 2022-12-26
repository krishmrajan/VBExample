VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "comctl32.ocx"
Begin VB.Form ClassForm 
   Caption         =   "Class Description  - properties"
   ClientHeight    =   5475
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6885
   LinkTopic       =   "Form1"
   ScaleHeight     =   5475
   ScaleWidth      =   6885
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Done 
      Caption         =   "Done"
      Height          =   375
      Left            =   5520
      TabIndex        =   1
      Top             =   5040
      Width           =   1095
   End
   Begin ComctlLib.ListView MsList1 
      Height          =   4215
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   7435
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327680
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      MouseIcon       =   "ClassForm.frx":0000
      NumItems        =   0
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   255
      Left            =   2280
      TabIndex        =   3
      Top             =   300
      Width           =   3135
   End
   Begin VB.Label Label1 
      Caption         =   "Class Name:"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   300
      Width           =   1815
   End
End
Attribute VB_Name = "ClassForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Done_Click()
    ClassForm.Hide
End Sub

Private Sub Form_Load()
    MsList1.ColumnHeaders.Clear
    MsList1.ColumnHeaders.Add , , "Property Name"
    MsList1.ColumnHeaders.Add , , "Data Type"
    MsList1.View = lvwReport
End Sub

