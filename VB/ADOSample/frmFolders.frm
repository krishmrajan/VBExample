VERSION 5.00
Begin VB.Form frmFolders 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select a Folder"
   ClientHeight    =   6516
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   4860
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6516
   ScaleWidth      =   4860
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClear 
      Caption         =   "&No Selection"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   5760
      Width           =   1215
   End
   Begin VB.PictureBox TreeView1 
      Height          =   5295
      Left            =   240
      ScaleHeight     =   5244
      ScaleWidth      =   4164
      TabIndex        =   1
      Top             =   240
      Width           =   4215
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "&Select Current"
      Height          =   375
      Left            =   3120
      TabIndex        =   0
      Top             =   5760
      Width           =   1335
   End
End
Attribute VB_Name = "frmFolders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public objCatalog As New IDMObjects.Library
Public strFolderID As String

Private Sub cmdClear_Click()
    strFolderID = ""
    Hide
End Sub

Private Sub cmdSelect_Click()
    strFolderID = "878090212"
    Hide
End Sub

