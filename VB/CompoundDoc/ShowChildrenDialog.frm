VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form ShowChildrenDialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Children of Compound Document"
   ClientHeight    =   3675
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6060
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3675
   ScaleWidth      =   6060
   ShowInTaskbar   =   0   'False
   Begin ComctlLib.ListView ListView1 
      Height          =   2775
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   4895
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Label"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "ID"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Level"
         Object.Width           =   1587
      EndProperty
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   3120
      Width           =   1215
   End
End
Attribute VB_Name = "ShowChildrenDialog"
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
' $Date:    $
' $Author:   $
' $Workfile:   l$
'----------------------------------------------------------------------
' All Rights Reserved.  Copyright (c) 1988,1999 FileNet Corp.
'----------------------------------------------------------------------

Option Explicit
Public oParentDocument As IDMObjects.Document

Private Sub AddChildrenOfParent(oDocument As IDMObjects.Document, level As Integer)
    On Error GoTo ErrorHandler
    
    ' first: add the document to the list at the current level
    Dim label As String
    label = String((level - 1) * 4, " ") & oDocument.label  ' add some spacing based on level
    Dim oItem As ListItem
    Set oItem = ListView1.ListItems.Add(, , label)
    oItem.SubItems(1) = oDocument.ID
    oItem.SubItems(2) = level
    
    ' second: check if the document has children and add each child recursively
    If oDocument.GetState(idmDocHasChild) Then
        Dim childIndex As Integer
        For childIndex = 1 To oDocument.Compound.Children.Count
            ' get the link object between the parent and each child
            Dim oParentToChildLink As IDMObjects.Link
            Set oParentToChildLink = oDocument.Compound.Children.Item(childIndex)
            
            ' get the child document object from the link object
            Dim oChildDocument As IDMObjects.Document
            Set oChildDocument = oParentToChildLink.Child
            
            ' call ourself to add the children of each child
            Call AddChildrenOfParent(oChildDocument, level + 1)
            
        Next childIndex
    End If

Exit Sub
ErrorHandler:
    MsgBox Err.Description, , "Error while retrieving children"

End Sub

Private Sub RefreshList()
    
    ' show the children of the current parent document
    ListView1.ListItems.Clear
    
    ' initiate the recursive calls to add children of the parent
    AddChildrenOfParent oParentDocument, 1
    
End Sub

Private Sub Form_Load()
    Call CenterForm(Me)
    RefreshList
End Sub

Private Sub OKButton_Click()
    Unload Me
End Sub
