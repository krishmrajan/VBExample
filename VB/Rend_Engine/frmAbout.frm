VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About HTML Rendition Engine"
   ClientHeight    =   3390
   ClientLeft      =   5550
   ClientTop       =   3735
   ClientWidth     =   5835
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   5835
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2250
      TabIndex        =   0
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "tools used to create and debug procedures."
      Height          =   195
      Index           =   5
      Left            =   840
      TabIndex        =   8
      Top             =   1920
      Width           =   3090
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "familiar with the programming language being demonstrated and the"
      Height          =   195
      Index           =   4
      Left            =   840
      TabIndex        =   7
      Top             =   1680
      Width           =   4755
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "fitness for a particular purpose. This sample assumes that you are"
      Height          =   195
      Index           =   3
      Left            =   840
      TabIndex        =   6
      Top             =   1440
      Width           =   4590
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "not limited to, the implied warranties of merchantability and/or"
      Height          =   195
      Index           =   2
      Left            =   840
      TabIndex        =   5
      Top             =   1200
      Width           =   4290
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "only, without warranty either expressed or implied,  including, but"
      Height          =   195
      Index           =   1
      Left            =   840
      TabIndex        =   4
      Top             =   960
      Width           =   4515
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Disclaimer: FileNET provides programming examples for illustration"
      Height          =   195
      Index           =   0
      Left            =   840
      TabIndex        =   3
      Top             =   720
      Width           =   4635
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "IDM Custom Rendition Engine Sample"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   840
      TabIndex        =   2
      Top             =   120
      Width           =   4770
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Copyright © 1999, FileNET Corporation"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   840
      TabIndex        =   1
      Top             =   2400
      Width           =   3285
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmAbout.frx":014A
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' This program is an example which uses Publishing foundation objects
'----------------------------------------------------------------------
'   Disclaimer: FileNET provides programming examples for illustration
'   only, without warranty either expressed or implied,  including, but
'   not limited to, the implied warranties of merchantability and/or
'   fitness for a particular purpose. This sample assumes that you are
'   familiar with the programming language being demonstrated and the
'   tools used to create and debug procedures.
'----------------------------------------------------------------------

' Revision:   1.1
' Date:       November 19, 1999 12:35:54
' Author:     Vladimir Fridman
' Workfile:   frmAbout.frm

'----------------------------------------------------------------------
' All Rights Reserved.  Copyright (c) 1999 FileNet Corp.
'----------------------------------------------------------------------

Option Explicit

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    CenterForm Me
End Sub
