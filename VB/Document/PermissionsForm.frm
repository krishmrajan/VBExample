VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "comctl32.ocx"
Begin VB.Form PermissionsForm 
   Caption         =   "Permissions"
   ClientHeight    =   4050
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7440
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   4050
   ScaleWidth      =   7440
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Done 
      Caption         =   "Done"
      Height          =   375
      Left            =   6000
      TabIndex        =   2
      Top             =   3480
      Width           =   1215
   End
   Begin ComctlLib.ListView MsList2 
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   4683
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327680
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   7095
   End
End
Attribute VB_Name = "PermissionsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Done_Click()
    PermissionsForm.Hide
End Sub

Private Sub Form_Load()
    MsList2.ColumnHeaders.Clear
    MsList2.ColumnHeaders.Add , , "Access Rights"
    MsList2.ColumnHeaders.Add , , "Access Type"
    MsList2.ColumnHeaders.Add , , "Grantee Type"
    MsList2.View = lvwReport
End Sub
