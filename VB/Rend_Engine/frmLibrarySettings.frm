VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComctl.ocx"
Begin VB.Form frmLibrarySettings 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Library Configuration"
   ClientHeight    =   4935
   ClientLeft      =   4410
   ClientTop       =   3540
   ClientWidth     =   6975
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   6975
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraStatistics 
      Caption         =   "Rendition Engine Statistics"
      Height          =   4215
      Left            =   3840
      TabIndex        =   16
      Top             =   120
      Width           =   3015
      Begin VB.TextBox txtVersionsAdded 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   1920
         TabIndex        =   25
         Text            =   "?"
         Top             =   2040
         Width           =   855
      End
      Begin VB.TextBox txtDocsAdded 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   1920
         TabIndex        =   24
         Text            =   "?"
         Top             =   1680
         Width           =   855
      End
      Begin VB.TextBox txtPreventedCollsions 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   1920
         TabIndex        =   22
         Text            =   "?"
         Top             =   1320
         Width           =   855
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "Refresh"
         Height          =   375
         Left            =   1560
         TabIndex        =   11
         Top             =   3720
         Width           =   1335
      End
      Begin VB.TextBox txtRequestsInQueue 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   1920
         TabIndex        =   10
         Text            =   "?"
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox txtDocsPublished 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   9
         Text            =   "?"
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox txtQueries 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   8
         Text            =   "?"
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lblLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Versions added"
         Height          =   195
         Index           =   8
         Left            =   675
         TabIndex        =   27
         Top             =   2085
         Width           =   1095
      End
      Begin VB.Label lblLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Documents added"
         Height          =   195
         Index           =   7
         Left            =   465
         TabIndex        =   26
         Top             =   1725
         Width           =   1305
      End
      Begin VB.Label lblLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Prevented collisions"
         Height          =   195
         Index           =   6
         Left            =   360
         TabIndex        =   23
         Top             =   1365
         Width           =   1410
      End
      Begin VB.Label lblLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Requests in queue"
         Height          =   195
         Index           =   5
         Left            =   435
         TabIndex        =   19
         Top             =   1005
         Width           =   1335
      End
      Begin VB.Label lblLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Documents published"
         Height          =   195
         Index           =   4
         Left            =   240
         TabIndex        =   18
         Top             =   645
         Width           =   1530
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Number of queries"
         Height          =   195
         Left            =   480
         TabIndex        =   17
         Top             =   285
         Width           =   1290
      End
   End
   Begin VB.Frame fraFrame 
      Caption         =   "Library Settings"
      Height          =   2295
      Index           =   1
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   3615
      Begin VB.TextBox txtRequestsPerCheck 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3000
         TabIndex        =   20
         Text            =   "1"
         Top             =   1800
         Width           =   495
      End
      Begin VB.TextBox txtPassword 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1080
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   1200
         Width           =   2415
      End
      Begin VB.TextBox txtLoginId 
         Height          =   285
         Left            =   1080
         TabIndex        =   1
         Text            =   "Admin"
         Top             =   840
         Width           =   2415
      End
      Begin MSComctlLib.ImageCombo cboLibs 
         Height          =   330
         Left            =   1080
         TabIndex        =   0
         Top             =   360
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   582
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Locked          =   -1  'True
      End
      Begin VB.Label lblLabel 
         Caption         =   "Jobs to process per one query"
         Height          =   255
         Index           =   0
         Left            =   600
         TabIndex        =   21
         Top             =   1800
         Width           =   2295
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Password"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   15
         Top             =   1245
         Width           =   690
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Login ID"
         Height          =   195
         Index           =   2
         Left            =   330
         TabIndex        =   14
         Top             =   885
         Width           =   600
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Library"
         Height          =   195
         Index           =   3
         Left            =   465
         TabIndex        =   13
         Top             =   435
         Width           =   465
      End
   End
   Begin VB.Frame fraFrame 
      Caption         =   "Compound Documents Settings"
      Height          =   1815
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   2520
      Width           =   3615
      Begin VB.CheckBox chkFileWithParent 
         Caption         =   "File children in the same folder with parents"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   3375
      End
      Begin VB.CheckBox chkVersionChildren 
         Caption         =   "Version children.  Uncheck if you want to add new set of children for each version of the parent"
         Height          =   615
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   3255
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1440
      TabIndex        =   7
      Top             =   4440
      Width           =   1215
   End
End
Attribute VB_Name = "frmLibrarySettings"
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
' Workfile:   frmLibrarySettings.frm

'----------------------------------------------------------------------
' All Rights Reserved.  Copyright (c) 1999 FileNet Corp.
'----------------------------------------------------------------------

Option Explicit
Public oServicedLibrary As ServicedLib
Public bAddLibraryMode As Boolean
Public bPressedOK As Boolean


Private Sub cmdCancel_Click()
    bPressedOK = False
    Hide
End Sub

Private Sub cmdOK_Click()
    
    
    If Not cboLibs.SelectedItem Is Nothing Then
        oServicedLibrary.sLibraryName = cboLibs.SelectedItem.Text
    End If
    
    
    oServicedLibrary.sLoginID = txtLoginId.Text
    oServicedLibrary.sPassword = txtPassword.Text
    
    oServicedLibrary.bFileWithParent = CBool(chkFileWithParent.Value)
    oServicedLibrary.bVersionChildren = CBool(chkVersionChildren.Value)
    oServicedLibrary.nRequestsEachTime = CInt(txtRequestsPerCheck.Text)
    
   
    bPressedOK = True
    Hide
End Sub

Private Sub cmdRefresh_Click()
'load statistics information

    txtQueries.Text = oServicedLibrary.nQueries
    txtDocsPublished.Text = oServicedLibrary.nDocsPublishedOK
    txtPreventedCollsions.Text = oServicedLibrary.nPreventedCollisions
    txtRequestsInQueue.Text = oServicedLibrary.RequestsInQueue
    txtDocsAdded.Text = oServicedLibrary.nDocumentsAdded
    txtVersionsAdded.Text = oServicedLibrary.nVersionsAdded
End Sub

Private Sub Form_Load()
    
    Set cboLibs.ImageList = frmMain.imgImages
    
    If bAddLibraryMode Then
        Caption = "Add Library"
        cboLibs.Visible = True
        LoadAvailableLibraries
        txtRequestsPerCheck.Text = "1"
        fraStatistics.Visible = False
        Width = 3945
    Else
        
        cboLibs.ComboItems.Add , , oServicedLibrary.sLibraryName, "library"
        cboLibs.Enabled = False
        txtLoginId.Text = oServicedLibrary.sLoginID
        txtPassword.Text = oServicedLibrary.sPassword
        chkFileWithParent.Value = Abs(oServicedLibrary.bFileWithParent)
        chkVersionChildren.Value = Abs(oServicedLibrary.bVersionChildren)
        txtRequestsPerCheck.Text = oServicedLibrary.nRequestsEachTime
        
        'statistics
        cmdRefresh_Click
        
    End If
    
    If cboLibs.ComboItems.Count <> 0 Then
        Set cboLibs.SelectedItem = cboLibs.ComboItems(1)
    End If

    CenterForm Me
    
End Sub


Sub LoadAvailableLibraries()
    Dim oTempLib As Library
    
    cboLibs.ComboItems.Clear
    For Each oTempLib In oHood.Libraries
        If Not IsLibraryServiced(oTempLib) Then
            cboLibs.ComboItems.Add , , oTempLib.Name, "library"
        End If
    Next
End Sub

Function IsLibraryServiced(oLib As Library) As Boolean
'checks if the library oLib is already Serviced

    Dim oServLib As ServicedLib
    IsLibraryServiced = False
    For Each oServLib In oServicedLibraries
        If oLib.Name = oServLib.sLibraryName Then
            IsLibraryServiced = True
            Exit For
        End If
    Next
End Function

