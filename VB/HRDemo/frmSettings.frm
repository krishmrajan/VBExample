VERSION 5.00
Begin VB.Form frmSettings 
   Caption         =   "Form1"
   ClientHeight    =   6372
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   6876
   LinkTopic       =   "Form1"
   ScaleHeight     =   6372
   ScaleWidth      =   6876
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   615
      Left            =   5280
      TabIndex        =   13
      Top             =   5640
      Width           =   1455
   End
   Begin VB.CommandButton btnOk 
      Caption         =   "Ok"
      Height          =   615
      Left            =   3600
      TabIndex        =   12
      Top             =   5640
      Width           =   1455
   End
   Begin VB.TextBox txtResDocClass 
      Height          =   375
      Left            =   2880
      TabIndex        =   5
      Top             =   2400
      Width           =   2655
   End
   Begin VB.TextBox txtIMSLibName 
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      Top             =   240
      Width           =   2655
   End
   Begin VB.TextBox txtMZLibName 
      Height          =   375
      Left            =   2880
      TabIndex        =   6
      Top             =   3240
      Width           =   2655
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear settings"
      Height          =   615
      Left            =   1848
      TabIndex        =   11
      Top             =   5640
      Width           =   1455
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save settings"
      Height          =   615
      Left            =   128
      TabIndex        =   10
      Top             =   5640
      Width           =   1455
   End
   Begin VB.TextBox txtBenFolder 
      Height          =   375
      Left            =   2880
      TabIndex        =   9
      Top             =   4920
      Width           =   2655
   End
   Begin VB.TextBox txtMZPassword 
      Height          =   375
      Left            =   2880
      TabIndex        =   8
      Top             =   4320
      Width           =   2655
   End
   Begin VB.TextBox txtMZUser 
      Height          =   375
      Left            =   2880
      TabIndex        =   7
      Top             =   3765
      Width           =   2655
   End
   Begin VB.TextBox txtResFolder 
      Height          =   375
      Left            =   2880
      TabIndex        =   4
      Top             =   1875
      Width           =   2655
   End
   Begin VB.TextBox txtIMSPassword 
      Height          =   375
      Left            =   2880
      TabIndex        =   3
      Top             =   1290
      Width           =   2655
   End
   Begin VB.TextBox txtIMSUser 
      Height          =   375
      Left            =   2880
      TabIndex        =   2
      Top             =   720
      Width           =   2655
   End
   Begin VB.Label Label9 
      Caption         =   "Resume DocClass:"
      Height          =   375
      Left            =   720
      TabIndex        =   21
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Label Label8 
      Caption         =   "IDMIS Library"
      Height          =   375
      Left            =   720
      TabIndex        =   20
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label Label7 
      Caption         =   "IDMDS Library:"
      Height          =   375
      Left            =   720
      TabIndex        =   19
      Top             =   3360
      Width           =   1695
   End
   Begin VB.Label Label6 
      Caption         =   "Benefits Folder:"
      Height          =   375
      Left            =   720
      TabIndex        =   18
      Top             =   5040
      Width           =   1335
   End
   Begin VB.Label Label5 
      Caption         =   "IDMDS Password:"
      Height          =   375
      Left            =   720
      TabIndex        =   17
      Top             =   4470
      Width           =   1695
   End
   Begin VB.Label Label4 
      Caption         =   "IDMDS Usercode:"
      Height          =   375
      Left            =   720
      TabIndex        =   16
      Top             =   3885
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "Resume Folder:"
      Height          =   375
      Left            =   720
      TabIndex        =   15
      Top             =   1995
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "IDMIS Password:"
      Height          =   375
      Left            =   720
      TabIndex        =   14
      Top             =   1410
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "IDMIS Usercode:"
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   840
      Width           =   1335
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnCancel_Click()
Me.Hide
End Sub

Private Sub btnOk_Click()
Me.Hide
End Sub

Private Sub cmdClear_Click()
Dim oCtrl As Control
Call goPersist.DeleteSettings(gsAppName, gsSectionName)
Call ClearEntries(Me)
End Sub

Public Sub cmdRefresh_Click()
Call ClearEntries(Me)
Call goPersist.GetSettings(gsAppName, gsSectionName, Me)
End Sub

Private Sub cmdSave_Click()
Call goPersist.SaveSettings(gsAppName, gsSectionName, Me)
End Sub
Private Sub ClearEntries(fFrm As Form)
Dim oCtrl As Control
For Each oCtrl In Me.Controls
    If TypeOf oCtrl Is TextBox Then
        oCtrl.Text = ""
    End If
Next
End Sub
