VERSION 5.00
Begin VB.Form frmLocRec 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Local Record"
   ClientHeight    =   5655
   ClientLeft      =   7725
   ClientTop       =   3780
   ClientWidth     =   5175
   Icon            =   "LocRec.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   5175
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   375
      Left            =   2640
      TabIndex        =   20
      Top             =   4680
      Width           =   1095
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      Height          =   375
      Left            =   3840
      TabIndex        =   21
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton cmdFileNetProps 
      Caption         =   "Document Properties..."
      Height          =   375
      Left            =   1800
      TabIndex        =   23
      Top             =   5160
      Width           =   1935
   End
   Begin VB.CommandButton cmdFileProps 
      Caption         =   "File Properties..."
      Height          =   375
      Left            =   120
      TabIndex        =   22
      Top             =   5160
      Width           =   1575
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   18
      Top             =   4680
      Width           =   1215
   End
   Begin VB.Frame fraFrame 
      Caption         =   "Local Record Properties"
      Height          =   4455
      Left            =   120
      TabIndex        =   26
      Top             =   120
      Width           =   4935
      Begin VB.TextBox txtDate 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   39
         Top             =   720
         Width           =   1815
      End
      Begin VB.TextBox txtSize 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   3480
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox txtLibraryLabel 
         Height          =   285
         Left            =   1080
         TabIndex        =   3
         Top             =   1080
         Width           =   1815
      End
      Begin VB.TextBox txtId 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3480
         TabIndex        =   2
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox txtLibraryId 
         Height          =   285
         Left            =   1080
         TabIndex        =   4
         Top             =   1440
         Width           =   3735
      End
      Begin VB.TextBox txtVersion 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4080
         TabIndex        =   6
         Top             =   2160
         Width           =   735
      End
      Begin VB.TextBox txtUser 
         Height          =   285
         Left            =   1080
         TabIndex        =   7
         Top             =   2160
         Width           =   2055
      End
      Begin VB.TextBox txtPath 
         Height          =   285
         Left            =   1080
         TabIndex        =   0
         Top             =   360
         Width           =   3735
      End
      Begin VB.TextBox txtUserString 
         Height          =   285
         Left            =   1080
         TabIndex        =   8
         Top             =   2520
         Width           =   3735
      End
      Begin VB.TextBox txtTitle 
         Height          =   285
         Left            =   1080
         TabIndex        =   5
         Top             =   1800
         Width           =   3735
      End
      Begin VB.TextBox txtLinkID 
         Height          =   285
         Left            =   1080
         TabIndex        =   9
         Top             =   2880
         Width           =   3735
      End
      Begin VB.CheckBox chkCheckedOut 
         Caption         =   "Is Checked Out"
         Height          =   195
         Left            =   2700
         TabIndex        =   16
         Top             =   3840
         Width           =   1575
      End
      Begin VB.CheckBox chkRendition 
         Caption         =   "Is Rendition"
         Height          =   195
         Left            =   2700
         TabIndex        =   15
         Top             =   3600
         Width           =   1335
      End
      Begin VB.CheckBox chkModified 
         Enabled         =   0   'False
         Height          =   195
         Left            =   1080
         TabIndex        =   11
         Top             =   3840
         Width           =   255
      End
      Begin VB.OptionButton optIS 
         Caption         =   "IS"
         Height          =   255
         Left            =   1860
         TabIndex        =   12
         Top             =   3480
         Width           =   495
      End
      Begin VB.OptionButton optDS 
         Caption         =   "DS"
         Height          =   255
         Left            =   1140
         TabIndex        =   10
         Top             =   3480
         Width           =   615
      End
      Begin VB.CheckBox chkKeptFile 
         Caption         =   "Is Kept File"
         Height          =   195
         Left            =   2700
         TabIndex        =   17
         Top             =   4080
         Width           =   1455
      End
      Begin VB.CheckBox chkLocalFile 
         Enabled         =   0   'False
         Height          =   195
         Left            =   1080
         TabIndex        =   25
         Top             =   4080
         Width           =   255
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000003&
         X1              =   120
         X2              =   4800
         Y1              =   3360
         Y2              =   3360
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Date"
         Height          =   195
         Index           =   4
         Left            =   630
         TabIndex        =   38
         Top             =   765
         Width           =   345
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Size"
         Height          =   195
         Index           =   2
         Left            =   3075
         TabIndex        =   37
         Top             =   720
         Width           =   300
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Path"
         Height          =   195
         Index           =   3
         Left            =   645
         TabIndex        =   36
         Top             =   405
         Width           =   330
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "LibraryLabel"
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   35
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Id"
         Height          =   195
         Index           =   5
         Left            =   3240
         TabIndex        =   34
         Top             =   1125
         Width           =   135
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "LibraryId"
         Height          =   195
         Index           =   7
         Left            =   375
         TabIndex        =   33
         Top             =   1485
         Width           =   600
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Version"
         Height          =   195
         Index           =   9
         Left            =   3360
         TabIndex        =   32
         Top             =   2205
         Width           =   525
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "User"
         Height          =   195
         Index           =   10
         Left            =   645
         TabIndex        =   31
         Top             =   2205
         Width           =   330
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "UserString"
         Height          =   195
         Index           =   11
         Left            =   240
         TabIndex        =   30
         Top             =   2565
         Width           =   735
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Title"
         Height          =   195
         Index           =   8
         Left            =   675
         TabIndex        =   29
         Top             =   1845
         Width           =   300
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "LinkID"
         Height          =   195
         Index           =   1
         Left            =   510
         TabIndex        =   28
         Top             =   2925
         Width           =   465
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   " Is LocalFile"
         Height          =   195
         Index           =   12
         Left            =   1320
         TabIndex        =   14
         Top             =   4080
         Width           =   840
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   " Is Modified"
         Height          =   195
         Index           =   0
         Left            =   1320
         TabIndex        =   13
         Top             =   3840
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Library Type:"
         Height          =   195
         Left            =   120
         TabIndex        =   27
         Top             =   3510
         Width           =   915
      End
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1440
      TabIndex        =   19
      Top             =   4680
      Width           =   1095
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Update"
      Height          =   375
      Left            =   3840
      TabIndex        =   24
      Top             =   5160
      Width           =   1215
   End
End
Attribute VB_Name = "frmLocRec"
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
' $Revision:   1.4  $
' $Date:   15 Nov 1999 18:13:02  $
' $Author:   vfridman  $
' $Workfile:   LocRec.frm  $
'----------------------------------------------------------------------
' All Rights Reserved.  Copyright (c) 1988,1999 FileNet Corp.
'----------------------------------------------------------------------

Option Explicit
Public oEditedRecord As LocalRecord

Private Sub cmdApply_Click()
'save record properties
'update the fields of the local record
    
    On Error GoTo ErrorHandler
    
    
    oEditedRecord.Path = txtPath.Text
    oEditedRecord.ID = txtId.Text
    oEditedRecord.LibraryLabel = txtLibraryLabel.Text
    oEditedRecord.LibraryId = txtLibraryId.Text
    oEditedRecord.Title = txtTitle.Text
    oEditedRecord.Version = txtVersion.Text
    oEditedRecord.User = txtUser.Text
    oEditedRecord.UserString = txtUserString.Text
    oEditedRecord.IsCheckedOut = chkCheckedOut.Value
    oEditedRecord.IsRendition = chkRendition.Value
    oEditedRecord.IsKeptFile = chkKeptFile.Value
    oEditedRecord.LinkId = txtLibraryId.Text
    
    If (optDS = True) Then
        oEditedRecord.LibrarySystemType = idmSysTypeDS
    Else
        oEditedRecord.LibrarySystemType = idmSysTypeIS
    End If
Exit Sub
ErrorHandler:
    MsgBox Err.Description, , "Saving record properties"
    Resume Next
End Sub

Private Sub cmdFileNetProps_Click()
    frmLocRecs.mnuRecordFileNETProps_Click
End Sub

Private Sub cmdFileProps_Click()
    frmLocRecs.mnuRecordFileProps_Click
End Sub

Private Sub cmdRefresh_Click()
'Reloads record properties
    Form_Load
End Sub

Private Sub cmdSave_Click()
'save properties and dismiss the form
    
    cmdApply_Click
    
    ' unload the form
    Unload Me

End Sub
Private Sub cmdClose_Click()
    ' unload the form
    Unload Me
End Sub

Private Sub cmdUpdate_Click()
    ' refresh the contents of the local record, updating date and time
    oEditedRecord.Update
    
    ' load the updated values into the form
    Form_Load
End Sub

Private Sub Form_Load()
    'load properties of the local record into the form
    
    On Error GoTo ErrorHandler
    txtPath = oEditedRecord.Path
    txtDate = oEditedRecord.Date
    txtSize = oEditedRecord.Size
    txtId = oEditedRecord.ID
    txtLibraryLabel = oEditedRecord.LibraryLabel
    txtLibraryId = oEditedRecord.LibraryId
    txtTitle = oEditedRecord.Title
    txtVersion = oEditedRecord.Version
    txtUser = oEditedRecord.User
    txtUserString = oEditedRecord.UserString
    txtLinkID = oEditedRecord.LinkId
    chkCheckedOut = Abs(oEditedRecord.IsCheckedOut)
    chkLocalFile = Abs(oEditedRecord.IsLocalFile)
    chkRendition = Abs(oEditedRecord.IsRendition)
    chkKeptFile = Abs(oEditedRecord.IsKeptFile)
    chkModified = Abs(oEditedRecord.IsModified)
    
    If oEditedRecord.LibrarySystemType = idmSysTypeDS Then
        optDS = True
        optIS = False
    ElseIf oEditedRecord.LibrarySystemType = idmSysTypeIS Then
        optDS = False
        optIS = True
    Else
        optDS = False
        optIS = False
    End If
    

Exit Sub
ErrorHandler:
    MsgBox Err.Description, , "Loading record properties"
    Resume Next
End Sub


