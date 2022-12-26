VERSION 5.00
Object = "{39533EA1-2D83-11D2-BB36-006008161DBB}#3.0#0"; "FnActionGrid.ocx"
Begin VB.Form ActionForm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Action Form"
   ClientHeight    =   3810
   ClientLeft      =   645
   ClientTop       =   2925
   ClientWidth     =   7425
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3810
   ScaleWidth      =   7425
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox ProbDesc 
      Enabled         =   0   'False
      Height          =   495
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   3240
      Width           =   3975
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   5880
      TabIndex        =   2
      Top             =   3240
      Width           =   1455
   End
   Begin VB.CommandButton cmdExecute 
      Caption         =   "&Execute"
      Height          =   495
      Left            =   4200
      TabIndex        =   1
      Top             =   3240
      Width           =   1455
   End
   Begin IDMActionGrid.IDMActionGrid ActionGrid 
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7215
      _Version        =   196608
      _ExtentX        =   12726
      _ExtentY        =   5318
      _StockProps     =   228
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   1
      EnableContextMenuPlugIn=   0   'False
      ShowTreeViewLines=   -1  'True
      ShowRowHeaders  =   -1  'True
   End
End
Attribute VB_Name = "ActionForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' This program is an example of how to use the new foundation objects
' for compound documents

'----------------------------------------------------------------------
'   Disclaimer: FileNET provides programming examples for illustration
'   only, without warranty either expressed or implied,  including, but
'   not limited to, the implied warranties of merchantability and/or
'   fitness for a particular purpose. This sample assumes that you are
'   familiar with the programming language being demonstrated and the
'   tools used to create and debug procedures.
'----------------------------------------------------------------------

'----------------------------------------------------------------------
' $Revision:     $
' $Date:     $
' $Author:     $
' $Workfile:     $
'----------------------------------------------------------------------
' All Rights Reserved.  Copyright (c) 1988,1999 FileNet Corp.
'----------------------------------------------------------------------

Option Explicit

Private Sub ActionGrid_ItemSelectChange(ByVal action As Object, ByVal Key As Long, ByVal Selected As Boolean)
    ' update the problem description text when the user modifies the control
    Dim oAction As IDMObjects.action
    Set oAction = action
    ProbDesc.Text = oAction.ProblemDescription.Description
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdExecute_Click()
    
    On Error GoTo ErrHandler
    
    ' get the root action from the action grid control (first item is the root)
    Dim oRootAction As IDMObjects.action
    Set oRootAction = ActionGrid.GetItem(1)
   
    ' execute the action and check the result
    Dim success As Boolean
    success = oRootAction.Execute(idmPromptOnFailure, False)
    If (Not success) Then
        MsgBox "There were problems executing the action for the compound document", vbCritical, "Execute Command"
        cmdCancel.Caption = "&Cancel"
        Exit Sub
    End If
    cmdCancel.Caption = "&Done"
    
    Exit Sub
    
ErrHandler:
    MsgBox Err.Description, vbCritical, "Execute Command"
End Sub

Private Sub Form_Load()
    Call CenterForm(Me)
    
    ' setup the columns we want to see in the control (status icon, file name, command and version)
    ActionGrid.AddColumnHeader idmActionGridColumnStatus
    ActionGrid.AddColumnHeader idmActionGridColumn.idmActionGridColumnCommand
    ActionGrid.AddColumnHeader idmActionGridColumn.idmActionGridColumnVersion
    ActionGrid.AddColumnHeader idmActionGridColumn.idmActionGridColumnFileName
    
    cmdCancel.Caption = "&Cancel"
End Sub

