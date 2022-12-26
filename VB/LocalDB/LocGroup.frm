VERSION 5.00
Begin VB.Form frmLocGroup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Local Group"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5175
   Icon            =   "LocGroup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   5175
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   375
      Left            =   2640
      TabIndex        =   5
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton cmdExploreFolder 
      Caption         =   "Explore Folder"
      Height          =   375
      Left            =   3480
      TabIndex        =   9
      Top             =   2520
      Width           =   1575
   End
   Begin VB.CommandButton cmdOpenFolder 
      Caption         =   "Open Folder"
      Height          =   375
      Left            =   1800
      TabIndex        =   8
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Frame fraFrame 
      Caption         =   "Local Group Properties"
      Height          =   1455
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   4935
      Begin VB.TextBox txtUserString 
         Height          =   285
         Left            =   960
         TabIndex        =   1
         Top             =   720
         Width           =   3855
      End
      Begin VB.TextBox txtPath 
         Height          =   285
         Left            =   960
         TabIndex        =   0
         Top             =   360
         Width           =   3855
      End
      Begin VB.CheckBox chkIsKeptFolder 
         Caption         =   "Is Kept Folder"
         Height          =   195
         Left            =   960
         TabIndex        =   2
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "UserString"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   765
         Width           =   735
      End
      Begin VB.Label Path_label 
         AutoSize        =   -1  'True
         Caption         =   "Path"
         Height          =   195
         Left            =   525
         TabIndex        =   12
         Top             =   405
         Width           =   330
      End
   End
   Begin VB.CommandButton cmdLocalFolderProps 
      Caption         =   "Folder Properties..."
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   2520
      Width           =   1575
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      Height          =   375
      Left            =   3840
      TabIndex        =   6
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1440
      TabIndex        =   4
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "This group contains X Local Groups and X Local Records"
      Height          =   195
      Left            =   120
      TabIndex        =   10
      Top             =   1680
      Width           =   4080
   End
End
Attribute VB_Name = "frmLocGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' This program is an example which uses the Local DB foundation objects

'----------------------------------------------------------------------
'   Disclaimer: FileNET provides programming examples for illustration
'   only, without warranty either expressed or implied,  including, but
'   not limited to, the implied warranties of merchantability and/or
'   fitness for a particular purpose. This sample assumes that you are
'   familiar with the programming language being demonstrated and the
'   tools used to create and debug procedures.
'----------------------------------------------------------------------

'----------------------------------------------------------------------
' $Revision:   2.0  $
' $Date:   14 November 1999 14:43:12  $
' $Author:   Vladimir Fridman $
' $Workfile:   LocGroup.frm  $
'----------------------------------------------------------------------
' All Rights Reserved.  Copyright (c) 1988,1999 FileNet Corp.
'----------------------------------------------------------------------

Option Explicit
Public oLocalGroup As LocalGroup

Private Sub cmdApply_Click()
'Save localgroup properties

    On Error GoTo ErrorHandler
    oLocalGroup.Path = txtPath.Text
    oLocalGroup.UserString = txtUserString.Text
    oLocalGroup.IsKeptFolder = CBool(chkIsKeptFolder.Value)
Exit Sub
ErrorHandler:
    MsgBox Err.Description, , "Error in saving LocalGroup"
End Sub

Private Sub cmdClose_Click()
'dismiss the dialog

    mbCancelAdd = True
    Unload Me
End Sub

Private Sub cmdExploreFolder_Click()
    frmLocRecs.mnuGroupExplore_Click
End Sub

Private Sub cmdLocalFolderProps_Click()
    frmLocRecs.mnuGroupFolderProps_Click
End Sub

Private Sub cmdOpenFolder_Click()
    frmLocRecs.mnuGroupOpenFolder_Click
End Sub

Private Sub cmdRefresh_Click()
'reload the values of the LocalGroup properties
    Form_Load
End Sub

Private Sub cmdSave_Click()
    cmdApply_Click
    Unload Me
End Sub

Private Sub Form_Load()
'Load LocalGroup properties into the form

    chkIsKeptFolder.Value = Abs(oLocalGroup.IsKeptFolder)
    txtPath = oLocalGroup.Path
    txtUserString = oLocalGroup.UserString
    lblInfo.Caption = "This Group contains " + CStr(oLocalGroup.Groups.Count) + " Groups and " + CStr(oLocalGroup.Records.Count) + " Records"
End Sub


