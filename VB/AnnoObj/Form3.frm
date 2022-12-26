VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Modify Security"
   ClientHeight    =   4020
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6285
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   ScaleHeight     =   4020
   ScaleWidth      =   6285
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Cancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5040
      TabIndex        =   11
      Top             =   840
      Width           =   975
   End
   Begin VB.CommandButton OK 
      Caption         =   "OK"
      Height          =   375
      Left            =   5040
      TabIndex        =   10
      Top             =   360
      Width           =   975
   End
   Begin VB.CommandButton SetExecute 
      Caption         =   "Set Append/Execute"
      Height          =   375
      Left            =   4080
      TabIndex        =   9
      Top             =   3240
      Width           =   2055
   End
   Begin VB.CommandButton SetWrite 
      Caption         =   "Set Write"
      Height          =   375
      Left            =   4080
      TabIndex        =   8
      Top             =   2760
      Width           =   2055
   End
   Begin VB.CommandButton SetRead 
      Caption         =   "Set Read"
      Height          =   375
      Left            =   4080
      TabIndex        =   7
      Top             =   2280
      Width           =   2055
   End
   Begin VB.ListBox List1 
      Height          =   1425
      Left            =   240
      TabIndex        =   6
      Top             =   2280
      Width           =   3735
   End
   Begin VB.TextBox ExecuteSec 
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Text            =   "Text3"
      Top             =   1440
      Width           =   2535
   End
   Begin VB.TextBox WriteSec 
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   960
      Width           =   2535
   End
   Begin VB.TextBox ReadSec 
      Height          =   375
      Left            =   1800
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   480
      Width           =   2535
   End
   Begin VB.Label Label3 
      Caption         =   "Append/Execute"
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Write"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Read"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   480
      Width           =   1335
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cancel_Click()
    Form3.Hide
End Sub


Private Sub Form_Activate()
    ReadSec = Form1.objAnno.Permissions(1).GranteeName
    WriteSec = Form1.objAnno.Permissions(2).GranteeName
    ExecuteSec = Form1.objAnno.Permissions(3).GranteeName
    

    
    ' Show groups in the list box
    Dim objGroup As idmobjects.Group
    For Each objGroup In Form1.objGroups
        List1.AddItem objGroup.Name
    Next
    
    'Show users in the list box
    Dim objUser As idmobjects.User
    For Each objUser In Form1.objUsers
        List1.AddItem objUser.Name
    Next
End Sub

Private Sub OK_Click()
    Form1.objAnno.Permissions(1).GranteeName = ReadSec
    Form1.objAnno.Permissions(2).GranteeName = WriteSec
    Form1.objAnno.Permissions(3).GranteeName = ExecuteSec
    Form3.Hide
End Sub

Private Sub SetExecute_Click()
    ExecuteSec = List1
End Sub

Private Sub SetRead_Click()
    ReadSec = List1
End Sub

Private Sub SetWrite_Click()
    WriteSec = List1
End Sub


