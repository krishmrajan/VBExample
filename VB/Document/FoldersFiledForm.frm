VERSION 5.00
Begin VB.Form FoldersFiledForm 
   Caption         =   "Folders filed in"
   ClientHeight    =   4320
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6630
   LinkTopic       =   "Form1"
   ScaleHeight     =   4320
   ScaleWidth      =   6630
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   2985
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   6375
   End
   Begin VB.CommandButton Done 
      Caption         =   "Done"
      Height          =   375
      Left            =   5160
      TabIndex        =   1
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "FoldersFiledForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Done_Click()
    FoldersFiledForm.Hide
End Sub

