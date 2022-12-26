VERSION 5.00
Begin VB.Form AboutBox 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About QSample"
   ClientHeight    =   1665
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6270
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1665
   ScaleWidth      =   6270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   120
      Picture         =   "AboutBox.frx":0000
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   7
      Top             =   240
      Width           =   540
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   5040
      TabIndex        =   0
      Top             =   240
      Width           =   975
   End
   Begin VB.Label lblCopyright 
      Caption         =   "Copyright (c) FileNet Corp. 2003.  All rights reserved."
      Height          =   255
      Left            =   840
      TabIndex        =   6
      Top             =   1080
      Width           =   3855
   End
   Begin VB.Label lblBuildVersion 
      Caption         =   "???"
      Height          =   255
      Left            =   1320
      TabIndex        =   5
      Top             =   840
      Width           =   4215
   End
   Begin VB.Label lblBuild 
      Caption         =   "Build:"
      Height          =   255
      Left            =   840
      TabIndex        =   4
      Top             =   840
      Width           =   495
   End
   Begin VB.Label lblApplicationVersion 
      Caption         =   "???"
      Height          =   255
      Left            =   2400
      TabIndex        =   3
      Top             =   600
      Width           =   855
   End
   Begin VB.Label lblApplication 
      Caption         =   "QSample Application"
      Height          =   255
      Left            =   840
      TabIndex        =   2
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label lblCompany 
      Caption         =   "FileNet Corporation"
      Height          =   255
      Left            =   840
      TabIndex        =   1
      Top             =   360
      Width           =   2295
   End
End
Attribute VB_Name = "AboutBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Load()
    Dim sSourceFile As String
    
    sSourceFile = GetIDMInstallPath & "\FnLocDB.exe"
    GetBuildInfo (sSourceFile)
    lblApplicationVersion = gApplicationVersion
    lblBuildVersion = gFileVersion & " (" & gBuildVersion & ")"

End Sub
Private Sub cmdOK_Click()
    Unload Me
End Sub

