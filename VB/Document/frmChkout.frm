VERSION 5.00
Begin VB.Form frmChkout 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6885
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   6885
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnOk 
      Caption         =   "Ok"
      Height          =   495
      Left            =   1800
      TabIndex        =   6
      Top             =   2280
      Width           =   855
   End
   Begin VB.TextBox txtFilepart 
      Height          =   375
      Left            =   1800
      TabIndex        =   5
      Top             =   1560
      Width           =   2415
   End
   Begin VB.TextBox txtDirpart 
      Height          =   375
      Left            =   1800
      TabIndex        =   4
      Top             =   1020
      Width           =   2415
   End
   Begin VB.TextBox txtFullpath 
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Top             =   480
      Width           =   4575
   End
   Begin VB.Label Label3 
      Caption         =   "Filename part:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Directory part:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Full path param:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   1335
   End
End
Attribute VB_Name = "frmChkout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnOk_Click()
Me.Hide
End Sub
