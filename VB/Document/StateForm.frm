VERSION 5.00
Begin VB.Form StateForm 
   Caption         =   "State"
   ClientHeight    =   5475
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6555
   LinkTopic       =   "Form2"
   ScaleHeight     =   5475
   ScaleWidth      =   6555
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Done 
      Caption         =   "Done"
      Height          =   375
      Left            =   5160
      TabIndex        =   20
      Top             =   5040
      Width           =   975
   End
   Begin VB.TextBox Text10 
      Height          =   285
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   19
      Text            =   "Text10"
      Top             =   4440
      Width           =   3015
   End
   Begin VB.TextBox Text9 
      Height          =   285
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   18
      Text            =   "Text9"
      Top             =   3960
      Width           =   3015
   End
   Begin VB.TextBox Text8 
      Height          =   285
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   17
      Text            =   "Text8"
      Top             =   3480
      Width           =   3015
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   16
      Text            =   "Text7"
      Top             =   3000
      Width           =   3015
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   15
      Text            =   "Text6"
      Top             =   2520
      Width           =   3015
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   14
      Text            =   "Text5"
      Top             =   2040
      Width           =   3015
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   13
      Text            =   "Text4"
      Top             =   1560
      Width           =   3015
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   12
      Text            =   "Text3"
      Top             =   1080
      Width           =   3015
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   11
      Text            =   "Text2"
      Top             =   600
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label Label10 
      Caption         =   "Label10"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   4560
      Width           =   2655
   End
   Begin VB.Label Label9 
      Caption         =   "Label9"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   4063
      Width           =   2655
   End
   Begin VB.Label Label8 
      Caption         =   "Label8"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   3572
      Width           =   2655
   End
   Begin VB.Label Label7 
      Caption         =   "Label7"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   3081
      Width           =   2655
   End
   Begin VB.Label Label6 
      Caption         =   "Label6"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   2590
      Width           =   2655
   End
   Begin VB.Label Label5 
      Caption         =   "Label5"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   2099
      Width           =   2655
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1608
      Width           =   2655
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   1117
      Width           =   2655
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   626
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   135
      Width           =   2655
   End
End
Attribute VB_Name = "StateForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Done_Click()
    StateForm.Hide
End Sub

