VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmStyleTemplate 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "StyleTemplate Properties"
   ClientHeight    =   5055
   ClientLeft      =   2505
   ClientTop       =   3540
   ClientWidth     =   5685
   Icon            =   "frmStyleTemplate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   5685
   ShowInTaskbar   =   0   'False
   Begin MSComDlg.CommonDialog cdDialog 
      Left            =   120
      Top             =   4560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Background color"
   End
   Begin VB.Frame fraFrame 
      Caption         =   "Rendering Options"
      Height          =   975
      Index           =   1
      Left            =   120
      TabIndex        =   18
      Top             =   3480
      Width           =   5415
      Begin VB.PictureBox picColor 
         Height          =   375
         Left            =   1920
         ScaleHeight     =   315
         ScaleWidth      =   3315
         TabIndex        =   20
         Top             =   360
         Width           =   3375
      End
      Begin VB.CommandButton cmdSelectColor 
         Caption         =   "Background color..."
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame fraFrame 
      Caption         =   "Style Template Properties"
      Height          =   3255
      Index           =   0
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   5415
      Begin VB.TextBox txtID 
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   240
         Width           =   3375
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   1920
         TabIndex        =   0
         Text            =   "HTML Template"
         Top             =   600
         Width           =   3375
      End
      Begin VB.TextBox txtInputMimeType 
         Height          =   285
         Left            =   1920
         TabIndex        =   1
         Text            =   "*/*"
         Top             =   960
         Width           =   3375
      End
      Begin VB.TextBox txtOutputMimeType 
         Height          =   285
         Left            =   1920
         TabIndex        =   2
         Text            =   "text/html"
         Top             =   1320
         Width           =   3375
      End
      Begin VB.TextBox txtRanking 
         Height          =   285
         Left            =   1920
         TabIndex        =   4
         Top             =   2040
         Width           =   3375
      End
      Begin VB.TextBox txtComment 
         Height          =   285
         Left            =   1920
         TabIndex        =   3
         Text            =   "Comment"
         Top             =   1680
         Width           =   3375
      End
      Begin VB.TextBox txtOutputName 
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   2760
         Width           =   3375
      End
      Begin VB.TextBox txtTemplateDefinition 
         Height          =   285
         Left            =   1920
         TabIndex        =   5
         Top             =   2400
         Width           =   3375
      End
      Begin VB.Label lblID 
         Alignment       =   1  'Right Justify
         Caption         =   "ID"
         Height          =   195
         Left            =   1590
         TabIndex        =   17
         Top             =   285
         Width           =   165
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   1260
         TabIndex        =   16
         Top             =   645
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Input Mime Type"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   5
         Left            =   330
         TabIndex        =   15
         Top             =   1005
         Width           =   1425
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Output Mime Type"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   6
         Left            =   195
         TabIndex        =   14
         Top             =   1365
         Width           =   1560
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Output Name"
         Height          =   195
         Index           =   7
         Left            =   810
         TabIndex        =   13
         Top             =   2805
         Width           =   945
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Ranking"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   8
         Left            =   1035
         TabIndex        =   12
         Top             =   2085
         Width           =   720
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Comment"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   9
         Left            =   975
         TabIndex        =   11
         Top             =   1725
         Width           =   780
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Template Definition"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   11
         Left            =   90
         TabIndex        =   10
         Top             =   2445
         Width           =   1665
      End
   End
   Begin VB.CommandButton cmdSaveMethod 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3000
      TabIndex        =   8
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4320
      TabIndex        =   7
      Top             =   4560
      Width           =   1215
   End
End
Attribute VB_Name = "frmStyleTemplate"
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
' Workfile:   frmStyleTemplate.frm

'----------------------------------------------------------------------
' All Rights Reserved.  Copyright (c) 1999 FileNet Corp.
'----------------------------------------------------------------------

Option Explicit
Public oStyleTemplate As StyleTemplate
Public bAddingMode As Boolean
Public oRenditionEngine As RenditionEngine


Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdSaveMethod_Click()
    On Error GoTo ErrorHandler
    
    If bAddingMode Then
        Set oStyleTemplate = oRenditionEngine.CreateStyleTemplate(txtID.Text)
    End If
    
    With oStyleTemplate
        .Comment = txtComment.Text
        .InputMimeType = txtInputMimeType.Text
        .OutputMimeType = txtOutputMimeType.Text
        .Name = txtName.Text
        .Ranking = txtRanking.Text
        .TemplateDefinition = txtTemplateDefinition.Text
        .Save
    End With
    
    Unload Me
Exit Sub
ErrorHandler:
    ShowError
End Sub


Private Sub cmdSelectColor_Click()
    
    cdDialog.ShowColor
    
    txtTemplateDefinition.Text = cdDialog.Color
    
    picColor.BackColor = cdDialog.Color
    
End Sub

Private Sub Form_Load()
    On Error GoTo ErrorHandler
    CenterForm Me
    
    If bAddingMode Then
        Caption = "Add Style Template"
        lblID.Font.Bold = True
        txtID.Locked = False
        txtID.BackColor = vbWhite
        txtID.Text = TEMPLATE_ID + CStr(Int((Rnd() * 1000000)))
    Else
        With oStyleTemplate
            
            txtID = .ID
            txtName = .Name
            txtInputMimeType = .InputMimeType
            txtOutputMimeType = .OutputMimeType
            txtOutputName = .OutputName
            txtRanking = .Ranking
            txtComment = .Comment
            txtTemplateDefinition = .TemplateDefinition
            
            ShowBackGroundColor
            
            
        End With
    End If
    
    
Exit Sub
ErrorHandler:
    ShowError
End Sub

Sub ShowBackGroundColor()
    If IsNumeric(txtTemplateDefinition.Text) Then
        picColor.BackColor = txtTemplateDefinition.Text
        cdDialog.Color = txtTemplateDefinition.Text
    End If
    
End Sub

