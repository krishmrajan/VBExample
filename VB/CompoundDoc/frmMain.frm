VERSION 5.00
Object = "{2A601BA0-B880-11CF-8185-444553540000}#2.0#0"; "FnList.ocx"
Object = "{54E6FE61-B93A-11CF-8185-444553540000}#2.0#0"; "FnTree.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form MainForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FileNET Compound Document Sample"
   ClientHeight    =   5700
   ClientLeft      =   1395
   ClientTop       =   2505
   ClientWidth     =   8310
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5700
   ScaleWidth      =   8310
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5640
      Top             =   5040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin IDMListView.IDMListView IDMListView1 
      Height          =   4455
      Left            =   2880
      TabIndex        =   1
      Top             =   600
      Width           =   5295
      _Version        =   131072
      _ExtentX        =   9340
      _ExtentY        =   7858
      _StockProps     =   239
      BackColor       =   -2147483643
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
      _ColumnHeaders  =   "frmMain.frx":0000
   End
   Begin IDMTreeView.IDMTreeView IDMTreeView1 
      Height          =   4455
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   2655
      _Version        =   131072
      _ExtentX        =   4683
      _ExtentY        =   7858
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
   End
   Begin ComctlLib.StatusBar sbrStatusbar 
      Align           =   2  'Align Bottom
      Height          =   390
      Left            =   0
      TabIndex        =   4
      Top             =   5310
      Width           =   8310
      _ExtentX        =   14658
      _ExtentY        =   688
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   4
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   7646
            Text            =   "Status"
            TextSave        =   "Status"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Object.Width           =   3704
            MinWidth        =   3704
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   6
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1588
            MinWidth        =   1587
            TextSave        =   "5/25/99"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1588
            MinWidth        =   1587
            TextSave        =   "5:23 PM"
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblFullName 
      Caption         =   "Label2"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3000
      TabIndex        =   3
      Top             =   240
      Width           =   5145
   End
   Begin VB.Label lblDir 
      Caption         =   "All Folders"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   2535
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuFolder 
      Caption         =   "Fo&lder"
      Begin VB.Menu mnuAdd 
         Caption         =   "Add Document"
      End
   End
   Begin VB.Menu mnuDocument 
      Caption         =   "&Document"
      Begin VB.Menu mnuDisplayHierarchy 
         Caption         =   "Display Hierarchy"
      End
      Begin VB.Menu mnuSep7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOpenApp 
         Caption         =   "Open"
      End
      Begin VB.Menu mnusep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCheckout 
         Caption         =   "Check Out"
      End
      Begin VB.Menu mnuCheckin 
         Caption         =   "Check In"
      End
      Begin VB.Menu mnuSep8 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "Print"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuShowUI 
         Caption         =   "Show User Interface"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuLargeIcons 
         Caption         =   "Large Icons"
      End
      Begin VB.Menu mnuSmallIcons 
         Caption         =   "Small Icons"
      End
      Begin VB.Menu mnulist 
         Caption         =   "List"
      End
      Begin VB.Menu mnuDetail 
         Caption         =   "Detail"
      End
      Begin VB.Menu mnusep12 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRefresh 
         Caption         =   "Refresh"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "MainForm"
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
' $Revision:   1.1  $
' $Date:   25 May 1999 17:24:06  $
' $Author:   chockenberry  $
' $Workfile:   frmMain.frm  $
'----------------------------------------------------------------------
' All Rights Reserved.  Copyright (c) 1988,1999 FileNet Corp.
'----------------------------------------------------------------------

Option Explicit

Public oNeighborhood As New IDMObjects.Neighborhood
Public oLibrary As IDMObjects.Library
Public oDocument As IDMObjects.Document
Public oFolder As IDMObjects.Folder

Dim showUI As Boolean

Private Sub UpdateMenus()
    
    ' set the state of menus and commands that require a document object
    If (oDocument Is Nothing) Then
        mnuDocument.Enabled = False
    Else
        mnuDocument.Enabled = True
    End If
    
    ' set the state of commands that require a folder object
    If (oFolder Is Nothing) Then
       mnuFolder.Enabled = False
    Else
       mnuFolder.Enabled = True
    End If

End Sub

Private Sub Form_Load()
    Call CenterForm(Me)
    
    ' initialize the tree view
    IDMTreeView1.AddRootItem oNeighborhood, True
    
    ' start out with no document or folder
    Set oDocument = Nothing
    Set oFolder = Nothing
    
    ' set the state of the menus
    UpdateMenus
    
    ' default to showing the UI during library operations
    showUI = True
    
End Sub

Private Sub IDMListView1_DblClick()
    ' check if the double-click item is a folder
    If IDMListView1.SelectedItem.ObjectType = idmObjTypeFolder Then
        ' save the folder object and open it in the tree view
        Dim oFolder As IDMObjects.Folder
        Set oFolder = IDMListView1.SelectedItem
        IDMTreeView1.SelectChildItem oFolder.Name
    End If
    
End Sub

Private Sub IDMListView1_ItemSelectChange(ByVal Item As Object, ByVal Key As Long, ByVal Selected As Boolean, ByVal ObjType As IDMListView.idmObjectType)
    ' check if the selected item is a document, otherwise it's a folder
    If (ObjType = idmObjTypeDocument) Then
        ' save the selected document object
        Set oDocument = Item
        
        ' update the menu items
        UpdateMenus
        
        ' get the status for the selected document
        Dim hasChildren As Boolean
        hasChildren = oDocument.GetState(idmDocHasChild)
        Dim isChild As Boolean
        isChild = oDocument.GetState(idmDocIsChild)
        
        ' update the status bar (both panels)
        Dim statusString1 As String
        Dim statusString2 As String
        If (hasChildren Or isChild) Then
            statusString1 = "Compound"
            If hasChildren Then
                Dim childCount As Integer
                childCount = oDocument.Compound.Children.Count
                statusString2 = "Parent with " & childCount
                ' let's be anal retentive
                If (childCount = 1) Then
                    statusString2 = statusString2 & " child"
                Else
                    statusString2 = statusString2 & " children"
                End If
            End If
            If isChild Then
                statusString2 = "Child"
            End If
        Else
            statusString1 = "Normal"
            statusString2 = ""
        End If
        sbrStatusbar.Panels(1).Text = statusString1 & " document selected"
        sbrStatusbar.Panels(2).Text = statusString2
        
    ElseIf (ObjType = idmObjTypeFolder) Then
        ' save the selected folder object
        Set oFolder = Item
        
        ' no document object is selected
        Set oDocument = Nothing
        
        ' update the menu items
        UpdateMenus
        
        ' update the status bar (both panels)
        sbrStatusbar.Panels(1).Text = "Folder selected"
        sbrStatusbar.Panels(2).Text = ""
        
    Else
        ' no document or folder object is selected
        Set oFolder = Nothing
        Set oDocument = Nothing
        
        ' update the status bar (both panels)
        sbrStatusbar.Panels(1).Text = "Unknown selected"
        sbrStatusbar.Panels(2).Text = ""
        
    End If
    
End Sub

Private Sub IDMTreeView1_ItemSelectChange(ByVal Item As Object, ByVal ObjType As IDMTreeView.idmObjectType)
    
    On Error Resume Next
    
    ' clear out the list view
    IDMListView1.ClearItems
    
    If (ObjType = idmObjTypeFolder) Then
        ' update the document and folder objects
        Set oDocument = Nothing
        Set oFolder = Item
        
        ' update list view
        IDMListView1.AddItems oFolder.SubFolders, -1
        IDMListView1.AddItems oFolder.GetContents(idmFolderContentDocument), -1
        'IDMListView1.AddItems oFolder.GetContents(idmFolderContentStoredSearch), -1
        
        ' update status bar (first panel)
        sbrStatusbar.Panels(1).Text = "Folder selected"
        
        ' update label above list view
        lblFullName.Caption = "Contents of '" & oFolder.label & "'"
    
    Else
        ' update the document and folder objects
        Set oDocument = Nothing
        Set oFolder = Nothing
    
        ' update status bar (first panel)
        Dim objTypeString As String
        objTypeString = "Unknown"
        Select Case ObjType
               Case idmObjTypeLibrary
                    objTypeString = "Library"
               Case idmObjTypeNeighborhood
                    objTypeString = "Neighborhood"
               Case idmObjTypeEntireNetwork
                    objTypeString = "Entire Network"
        End Select
        sbrStatusbar.Panels(1).Text = objTypeString & " selected"
        
        ' update label above list view
        lblFullName.Caption = " "
        
    End If
        
    ' update the menu items
    UpdateMenus
        
    ' update status bar (second panel)
    sbrStatusbar.Panels(2).Text = ""

End Sub


Private Sub mnuAdd_Click()
    If (Not oFolder Is Nothing) Then
        ' get a filename from the local file system
        Dim filePath As String
        CommonDialog1.Filter = "All Files (*.*)"
        CommonDialog1.ShowOpen
        filePath = CommonDialog1.fileName
        
        ' add the file to the library in the currently selected folder
        ' NOTE: you may need to change "General" to a document class that exists on your system
        Call AddDocument(oFolder, filePath, "General", showUI)
        
        ' refresh the list view (it won't happen automatically)
        Call RefreshListView
    End If
End Sub

Private Sub mnuDisplayHierarchy_Click()
    ' make sure that a document has been selected in the list view
    If Not oDocument Is Nothing Then
        ' display the dialog using the current document as a parent
        Dim oDialog As New ShowChildrenDialog
        Set oDialog.oParentDocument = oDocument
        oDialog.Show vbModal, Me
        
    End If

End Sub

Private Sub mnuOpenApp_Click()
    If (Not oDocument Is Nothing) Then
        ' open the currently selected document
        Call OpenDocument(oDocument, showUI)
    End If
End Sub

Private Sub mnuCheckout_Click()
    If (Not oDocument Is Nothing) Then
        ' check out the currently selected document
        Call CheckoutDocument(oDocument, showUI)
        
        ' refresh the list view
        Call RefreshListView
    End If
End Sub

Private Sub mnuCheckin_Click()
    If (Not oDocument Is Nothing) Then
        ' check in the currently selected document
        Call CheckinDocument(oDocument, showUI)
        
        ' refresh the list view
        Call RefreshListView
    End If
End Sub


Private Sub mnuAbout_Click()
    ' show some really useful information
    AboutForm.Show vbModeless, Me
End Sub

Private Sub mnuExit_Click()
   Unload Me
   End
End Sub


Private Sub mnuLargeIcons_Click()
   IDMListView1.View = idmViewIcon
End Sub

Private Sub mnuSmallIcons_Click()
    IDMListView1.View = idmViewSmallIcon
End Sub

Private Sub mnulist_Click()
    IDMListView1.View = idmViewList
End Sub

Private Sub mnuDetail_Click()
   IDMListView1.View = idmViewReport
End Sub

Private Sub mnuRefresh_Click()
   Call RefreshListView
End Sub


Private Sub mnuShowUI_Click()
    ' toggle the state of the variable that controls the UI during library operations
    showUI = Not showUI
    If (showUI) Then
        mnuShowUI.Checked = True
    Else
        mnuShowUI.Checked = False
    End If
    
End Sub

Public Sub RefreshListView()
   If (Not oFolder Is Nothing) Then
        oFolder.Refresh idmFolderRefreshAll
        IDMListView1.ClearItems
        IDMListView1.AddItems oFolder.SubFolders, -1
        IDMListView1.AddItems oFolder.GetContents(idmFolderContentDocument), -1
        'IDMListView1.AddItems oFolder.GetContents(idmFolderContentStoredSearch), -1
    End If
End Sub

