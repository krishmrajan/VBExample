VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.0#0"; "COMCT232.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Begin VB.Form FormProg 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Performing Document Committal..."
   ClientHeight    =   4290
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7620
   LinkTopic       =   "Form1"
   ScaleHeight     =   4290
   ScaleWidth      =   7620
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   495
      Left            =   1800
      TabIndex        =   5
      Top             =   360
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   873
      _Version        =   327682
      Appearance      =   1
   End
   Begin ComCtl2.Animation Animation1 
      Height          =   615
      Left            =   2280
      TabIndex        =   3
      Top             =   1320
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   1085
      _Version        =   327681
      FullWidth       =   137
      FullHeight      =   41
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   2040
      TabIndex        =   1
      Top             =   2280
      Width           =   4815
   End
   Begin VB.Label txtMax 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   6240
      TabIndex        =   2
      Top             =   960
      Width           =   495
   End
   Begin VB.Label txtMin 
      BackColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   1920
      TabIndex        =   4
      Top             =   960
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Uploading:"
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   2400
      Width           =   1095
   End
End
Attribute VB_Name = "FormProg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub


