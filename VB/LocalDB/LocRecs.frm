VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmLocRecs 
   Caption         =   "Local Database Explorer"
   ClientHeight    =   5325
   ClientLeft      =   4185
   ClientTop       =   1500
   ClientWidth     =   9195
   Icon            =   "LocRecs.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5325
   ScaleWidth      =   9195
   Begin MSComctlLib.StatusBar stbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   3
      Top             =   5040
      Width           =   9195
      _ExtentX        =   16219
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10583
            Text            =   "Root\asdafda\dfasdfasdf"
            TextSave        =   "Root\asdafda\dfasdfasdf"
            Key             =   "Path"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "Groups:"
            TextSave        =   "Groups:"
            Key             =   "Groups"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "Records:"
            TextSave        =   "Records:"
            Key             =   "Records"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView trvTree 
      Height          =   3855
      Left            =   0
      TabIndex        =   1
      Top             =   885
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   6800
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   353
      PathSeparator   =   "?"
      Style           =   7
      Appearance      =   1
      OLEDropMode     =   1
   End
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   8520
      Top             =   1560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imgImages 
      Left            =   8520
      Top             =   900
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   14
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "LocRecs.frx":0442
            Key             =   "group"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "LocRecs.frx":0894
            Key             =   "LocalFolders"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "LocRecs.frx":0CE6
            Key             =   "LocalFolders_opened"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "LocRecs.frx":1138
            Key             =   "hood"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "LocRecs.frx":2282
            Key             =   "LocalFiles"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "LocRecs.frx":26D4
            Key             =   "group_opened"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "LocRecs.frx":2B26
            Key             =   "record"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "LocRecs.frx":2C80
            Key             =   "record_checked_out"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "LocRecs.frx":2DDA
            Key             =   "Delete All"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "LocRecs.frx":332C
            Key             =   "Exit"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "LocRecs.frx":387E
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "LocRecs.frx":3990
            Key             =   "Properties"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "LocRecs.frx":3DE2
            Key             =   "Refresh"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "LocRecs.frx":4234
            Key             =   "add"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lstRecs 
      Height          =   3855
      Left            =   2160
      TabIndex        =   0
      Top             =   885
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   6800
      View            =   3
      Arrange         =   1
      Sorted          =   -1  'True
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      OLEDropMode     =   1
      AllowReorder    =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      OLEDropMode     =   1
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Path"
         Object.Width           =   3149
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Title"
         Object.Width           =   2116
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Library"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Doc ID"
         Object.Width           =   1746
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Version"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Checked out date"
         Object.Width           =   2999
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "User"
         Object.Width           =   1764
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbToolBar 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   9195
      _ExtentX        =   16219
      _ExtentY        =   1058
      ButtonWidth     =   1693
      ButtonHeight    =   953
      AllowCustomize  =   0   'False
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   10
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Refresh"
            Key             =   "Refresh"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Add Record"
            Key             =   "Add Record"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Add Group"
            Key             =   "Add Group"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Properties"
            Key             =   "Properties"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Delete"
            Key             =   "Delete"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Delete All"
            Key             =   "Delete All"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Exit"
            Key             =   "Exit"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblTree 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Local Groups"
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   600
      Width           =   2085
   End
   Begin VB.Label lblList 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Local Groups and Local Records"
      Height          =   255
      Left            =   2160
      TabIndex        =   4
      Top             =   600
      Width           =   6120
   End
   Begin VB.Image imgSplitter 
      Height          =   4305
      Left            =   2040
      MousePointer    =   9  'Size W E
      Top             =   600
      Width           =   135
   End
   Begin VB.Menu mnuRecord 
      Caption         =   "Record"
      Visible         =   0   'False
      Begin VB.Menu mnuRecordProps 
         Caption         =   "LocalDb Record Properties..."
      End
      Begin VB.Menu mnuSeparator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRecordFileNETProps 
         Caption         =   "FileNET Document Properties..."
      End
      Begin VB.Menu mnuSeparator2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRecordFileProps 
         Caption         =   "Local File Properties..."
      End
      Begin VB.Menu mnuRecordOpenFile 
         Caption         =   "Open Local File"
      End
      Begin VB.Menu mnuSeparator3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRecordDelete 
         Caption         =   "Delete Record"
      End
   End
   Begin VB.Menu mnuGroup 
      Caption         =   "Group"
      Visible         =   0   'False
      Begin VB.Menu mnuGroupProps 
         Caption         =   "LocalDb Group Properties..."
      End
      Begin VB.Menu mnuSeparator4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGroupFolderProps 
         Caption         =   "Local Folder Properties..."
      End
      Begin VB.Menu mnuGroupOpenFolder 
         Caption         =   "Open Local Folder"
      End
      Begin VB.Menu mnuGroupExplore 
         Caption         =   "Explore Local Folder"
      End
      Begin VB.Menu mnuSeparator5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGroupDelete 
         Caption         =   "Delete Group"
      End
   End
End
Attribute VB_Name = "frmLocRecs"
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
' $Revision:   1.1  $
' $Date:   15 Nov 1999 18:13:12  $
' $Author:   vfridman  $
' $Workfile:   LocRecs.frm  $
'----------------------------------------------------------------------
' All Rights Reserved.  Copyright (c) 1988,1999 FileNet Corp.
'----------------------------------------------------------------------

Option Explicit

Dim oLocalDB As New LocalDb, oCurrentGroup As LocalGroup, oListItem As ListItem, _
    oCurrentGroups As LocalGroups, oCurrentRecords As LocalRecords
    
Public mbMoving As Boolean
Const sglSplitLimit = 1500

Function FindChildNode(oParentNode As Node, sText As String) As Node
    Dim oTempNode As Node
    For Each oTempNode In trvTree.Nodes
        If oTempNode.Parent Is oParentNode And oTempNode.Text = sText Then
            Set FindChildNode = oTempNode
            Exit For
        End If
    Next
End Function

Sub RefreshList()
    ' shows all the local records and groups in the current group
    On Error GoTo ErrorHandler
    Dim intCounter, oRec As LocalRecord, oGroup As LocalGroup
    
    lstRecs.ListItems.Clear
    
    If Not oCurrentGroups Is Nothing Then
        For intCounter = 1 To oCurrentGroups.Count
            Set oGroup = oCurrentGroups.item(intCounter)
            With lstRecs.ListItems.Add(, , oGroup.Path, , "group")
                .Tag = "group"
            End With
        Next intCounter
        stbStatusBar.Panels("Groups") = "Groups: " + CStr(oCurrentGroups.Count)
        tbToolBar.Buttons("Add Group").Enabled = True
    Else
        tbToolBar.Buttons("Add Group").Enabled = False
        stbStatusBar.Panels("Groups") = "Groups: N/A"
    End If
    
    
    If Not oCurrentRecords Is Nothing Then
        For intCounter = 1 To oCurrentRecords.Count
            Set oRec = oCurrentRecords.item(intCounter)
            With lstRecs.ListItems.Add(, , oRec.Path)
                .SubItems(1) = oRec.Title
                .SubItems(2) = oRec.LibraryLabel
                .SubItems(3) = oRec.ID
                .SubItems(4) = oRec.Version
                .SubItems(5) = oRec.Date
                .SubItems(6) = oRec.User
                If oRec.IsCheckedOut Then
                    .SmallIcon = "record_checked_out"
                Else
                    .SmallIcon = "record"
                End If
            End With
        Next intCounter
        stbStatusBar.Panels("Records") = "Records: " + CStr(oCurrentRecords.Count)
        tbToolBar.Buttons("Add Record").Enabled = True
    Else
        stbStatusBar.Panels("Records") = "Records: N/A"
        tbToolBar.Buttons("Add Record").Enabled = False
    End If
    
    
Exit Sub
ErrorHandler:
    MsgBox Err.Description, , "Error while retrieving records"
    Resume Next
End Sub


Sub ConstructTree()
    Dim oRoot_Node As Node, oLocalFiles_Node As Node, oLocalFolders_Node As Node
    With trvTree.Nodes
        .Clear
        '''' add LocalFiles
        Set oLocalFiles_Node = .Add(, , "LocalFiles", "Local Files", "LocalFiles")
        '''' add LocalFolders
        Set oLocalFolders_Node = .Add(, , "LocalFolders", "Local Folders", "LocalFolders", "LocalFolders_opened")
        ''''' add LocalDb.Root
         Set oRoot_Node = .Add(, , "root", "LocalDB Root", "hood")
        oRoot_Node.Expanded = True
        AddGroupsToNode oRoot_Node, oLocalDB.Root
    End With
End Sub

Sub AddGroupsToNode(oParentNode As Node, oGroups As LocalGroups)
    On Error GoTo ErrorHandler
    Dim oTempGroup As LocalGroup, oCurrentNode As Node, n As Integer
    
    
    For n = 1 To oGroups.Count
        Set oTempGroup = oGroups.item(n)
        Set oCurrentNode = trvTree.Nodes.Add(oParentNode, tvwChild, , oTempGroup.Path, "group", "group_opened")
        oCurrentNode.Expanded = True
        If oTempGroup.Groups.Count <> 0 Then
            'recursivly add more to the current node
            AddGroupsToNode oCurrentNode, oTempGroup.Groups
            'trvTree.Nodes.Add oCurrentNode, tvwChild, , ""
        End If
    Next
    
Exit Sub
ErrorHandler:
    MsgBox Err.Description, , "Adding Groups to Node in the tree view"
End Sub
Sub SetImages()
    
    Set tbToolBar.ImageList = imgImages
    Set trvTree.ImageList = imgImages
    Set lstRecs.SmallIcons = imgImages
    
    
    With tbToolBar.Buttons
        .item("Refresh").Image = "Refresh"
        .item("Add Record").Image = "add"
        .item("Add Group").Image = "group"
        .item("Properties").Image = "Properties"
        .item("Delete").Image = "Delete"
        .item("Delete All").Image = "Delete All"
        .item("Exit").Image = "Exit"
    End With

End Sub


Private Sub Form_Load()
'
    SetImages
    
    ConstructTree
    trvTree_NodeClick trvTree.Nodes("LocalFiles")
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.Width < 7230 Then
        Me.Width = 7230
    Else
        SizeControls imgSplitter.Left, True
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub imgSplitter_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    mbMoving = True
End Sub

Private Sub imgSplitter_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim sglPos As Single

    If mbMoving Then
        sglPos = x + imgSplitter.Left
        If sglPos < sglSplitLimit Then
            imgSplitter.Left = sglSplitLimit
        ElseIf sglPos > Me.Width - sglSplitLimit Then
            imgSplitter.Left = Me.Width - sglSplitLimit
        Else
            imgSplitter.Left = sglPos
        End If
        SizeControls imgSplitter.Left, False
    End If

End Sub

Private Sub imgSplitter_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    'If mbMoving Then SizeControls imgSplitter.Left, False
    mbMoving = False
End Sub


Sub SizeControls(x As Single, bFormResize As Boolean)
    On Error Resume Next

    
    lstRecs.Left = x + 40
    lblList.Left = x + 40
    
    lstRecs.Width = Me.Width - trvTree.Width - 160
    lblList.Width = lstRecs.Width
    
    trvTree.Width = x
    lblTree.Width = x

    If bFormResize Then
        'set the height - we don't need to do that when moving the splitter
        trvTree.Height = Me.ScaleHeight - 1180
        lstRecs.Height = trvTree.Height
        imgSplitter.Height = trvTree.Height + 400
    End If
End Sub

Private Sub lstRecs_AfterLabelEdit(Cancel As Integer, NewString As String)
    On Error GoTo ErrorHandler
    Set oListItem = lstRecs.SelectedItem
    
    If oListItem.Tag = "group" Then
        oCurrentGroups.item(oListItem.Text).Path = NewString
        
        If trvTree.SelectedItem.Key <> "LocalFolders" Then 'if not in LocalFolders
            UpdateTree
            
        Else 'if in localfolders
            ConstructTree
            trvTree_NodeClick trvTree.Nodes("LocalFolders")
        End If
    
    Else
        oCurrentRecords.item(oListItem.Text).Path = NewString
    End If
    
Exit Sub
ErrorHandler:
    MsgBox Err.Description
    Cancel = True
End Sub

Sub UpdateTree()
    ClearChildNodes trvTree.SelectedItem
    AddGroupsToNode trvTree.SelectedItem, oCurrentGroups
End Sub


Private Sub lstRecs_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'Sort the list view

    lstRecs.SortOrder = IIf(lstRecs.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    lstRecs.SortKey = ColumnHeader.Index - 1
    
End Sub

Private Sub lstRecs_DblClick()
'If double clicked on a Group, change to the selected group in the tree view
'if double clicked on a Record, show record properties

    Dim oNode As Node
    If Not lstRecs.SelectedItem Is Nothing Then
        If lstRecs.SelectedItem.Tag = "group" Then
            If trvTree.SelectedItem.Key <> "LocalFolders" Then
                Set oNode = FindChildNode(trvTree.SelectedItem, lstRecs.SelectedItem.Text)
                If Not oNode Is Nothing Then
                    trvTree_NodeClick oNode
                End If
            End If
        Else
            ShowProperties
        End If
    End If
    

End Sub



Private Sub lstRecs_KeyDown(KeyCode As Integer, Shift As Integer)
'makes the list view respond to standard windows keys..
    
    If KeyCode = vbKeyF5 Then
        RefreshList
        
    ElseIf KeyCode = vbKeyReturn Then
        lstRecs_DblClick
    
    ElseIf KeyCode = vbKeyDelete Then
        DeleteFromList
    
    End If

End Sub

Private Sub lstRecs_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
'displays pop-up menu on right click

    If Button = 2 Then
        If Not lstRecs.SelectedItem Is Nothing Then
            If lstRecs.SelectedItem.Tag = "group" Then
                PopupMenu mnuGroup, vbPopupMenuRightButton, , , mnuGroupProps
            Else
                PopupMenu mnuRecord, vbPopupMenuRightButton, , , mnuRecordProps
            End If
        End If
    End If
    
End Sub



Private Sub mnuGroupDelete_Click()
'handle group delete from pop-up menu

    If ActiveControl Is trvTree Then
        DeleteFromTree
    ElseIf ActiveControl Is lstRecs Then
        DeleteFromList
    End If
    
End Sub

Public Sub mnuGroupExplore_Click()
'Opens explorer at the current folder, if the path exists
    
    On Error GoTo ErrorHandler
    Dim sPath As String, oGroup As LocalGroup
    
    If ActiveControl Is trvTree Then
        Set oGroup = FindGroupByNode(trvTree.SelectedItem)
    Else
        Set oGroup = oCurrentGroups.item(lstRecs.SelectedItem.Text)
    End If
    
    If oGroup Is Nothing Then
        MsgBox "Invalid Operation.  Please select a LocalGroup"
    Else
        sPath = oGroup.Path
        If Dir(sPath, vbDirectory) <> "" Then 'file exists
            ShellExecute Me.hwnd, "explore", sPath, vbNullChar, vbNullChar, 5
        Else
            MsgBox "Folder '" + sPath + "' does not exist"
        End If
    End If
    
Exit Sub
ErrorHandler:
    MsgBox Err.Description
End Sub

Public Sub mnuGroupFolderProps_Click()
'displays windows property dialog for the folder, if it exists

    On Error GoTo ErrorHandler
    Dim sPath As String, oGroup As LocalGroup
    
    If ActiveControl Is trvTree Then
        Set oGroup = FindGroupByNode(trvTree.SelectedItem)
    Else
        Set oGroup = oCurrentGroups.item(lstRecs.SelectedItem.Text)
    End If
    
    If oGroup Is Nothing Then
        MsgBox "Invalid Operation.  Please select a LocalGroup"
    Else
    
        sPath = oGroup.Path
        If Dir(sPath, vbDirectory) <> "" Then 'file exists
            ShowFileProperties sPath
        Else
            MsgBox "Folder '" + sPath + "' does not exist"
        End If
    
    End If
Exit Sub
ErrorHandler:
    MsgBox Err.Description
End Sub

Public Sub mnuGroupOpenFolder_Click()
'opens the folder, if it exists

    On Error GoTo ErrorHandler
    Dim sPath As String, oGroup As LocalGroup
    
    If ActiveControl Is trvTree Then
        Set oGroup = FindGroupByNode(trvTree.SelectedItem)
    Else
        Set oGroup = oCurrentGroups.item(lstRecs.SelectedItem.Text)
    End If
    
    If oGroup Is Nothing Then
        MsgBox "Invalid Operation.  Please select a LocalGroup"
    Else
        sPath = oGroup.Path
        If Dir(sPath, vbDirectory) <> "" Then 'file exists
            ShellExecute Me.hwnd, "open", sPath, vbNullChar, vbNullChar, 5
        Else
            MsgBox "Folder '" + sPath + "' does not exist"
        End If
    
    End If
Exit Sub
ErrorHandler:
    MsgBox Err.Description
End Sub

Private Sub mnuGroupProps_Click()
'shows properties for the selected LocalGroup

    ShowProperties
End Sub

Private Sub mnuRecordDelete_Click()
'handles "delete" from the record context menu

    DeleteFromList
End Sub

Public Sub mnuRecordFileNETProps_Click()
'displays filenet properties for the selected record

    On Error GoTo ErrorHandler
    Dim oLibrary As Library, oRecord As LocalRecord, oDocument As Document, oHood As Neighborhood
    Set oHood = New Neighborhood
    Set oRecord = oCurrentRecords.item(lstRecs.SelectedItem.Text)
    Set oLibrary = New Library
    oLibrary.Name = oRecord.LibraryId
    oLibrary.SystemType = idmSysTypeDS
    
    'Set oLibrary = oHood.Libraries(oRecord.LibraryId)
 
    If oLibrary.Logon(, , , idmLogonOptWithUI) Then
        Set oDocument = oLibrary.GetObject(idmObjTypeDocument, oRecord.ID)
        If oDocument.ShowPropertiesDialog = idmDialogExitOK Then
            oDocument.Save
        End If
    End If
    
Exit Sub
ErrorHandler:
    MsgBox "Unable to display FileNET Document Properties for the selected record", , "LocalDb Tool"
End Sub

Public Sub mnuRecordFileProps_Click()
'displays file properties for the selcted record

    On Error GoTo ErrorHandler
    Dim sPath As String, oRecord As LocalRecord
    Set oRecord = oCurrentRecords(lstRecs.SelectedItem.Text)
    
    sPath = oRecord.Path
    
    If Dir(sPath) <> "" Then 'file exists
        ShowFileProperties sPath
    Else
        MsgBox "File '" + sPath + "' does not exist"
    End If
Exit Sub
ErrorHandler:
    MsgBox Err.Description

End Sub

Private Sub mnuRecordOpenFile_Click()
'Open local file item on Record context menu

    On Error GoTo ErrorHandler
    Dim sPath As String, oRecord As LocalRecord
    Set oRecord = oCurrentRecords(lstRecs.SelectedItem.Text)
    
    sPath = oRecord.Path
    
    If Dir(sPath) <> "" Then 'file exists
        ShowFileProperties sPath, "open"
    Else
        MsgBox "File '" + sPath + "' does not exist"
    End If
Exit Sub
ErrorHandler:
    MsgBox Err.Description

End Sub

Private Sub mnuRecordProps_Click()
'Record properties context menu
    ShowProperties
End Sub

Sub Refresh_All()
'Refresh tree view or list view

    On Error GoTo ErrorHandler
    Dim oCurrentNode As Node
    Set oCurrentNode = trvTree.SelectedItem
    If ActiveControl Is trvTree Then
    
        If oCurrentNode.Key = "root" Then
            ConstructTree
            trvTree_NodeClick trvTree.Nodes("root")
            
        ElseIf oCurrentNode.Key = "LocalFiles" Then
            RefreshList
            
        ElseIf oCurrentNode.Key = "LocalFolders" Then
            ConstructTree
            trvTree_NodeClick trvTree.Nodes("LocalFolders")
        Else
            ClearChildNodes oCurrentNode
            AddGroupsToNode oCurrentNode, FindGroupByNode(oCurrentNode).Groups
            trvTree_NodeClick oCurrentNode
        End If
        
    ElseIf ActiveControl Is lstRecs Then
    
        RefreshList
        
        
    End If
Exit Sub
ErrorHandler:
    MsgBox Err.Description, , "Refresh_All"
End Sub

Sub ShowProperties()
'show properties for the selected item, either in list view or tree view

    On Error GoTo ErrorHandler
    Dim oListItem As ListItem
    
    If ActiveControl Is trvTree Then
        DisplayGroup FindGroupByNode(trvTree.SelectedItem)

    Else
        Set oListItem = lstRecs.SelectedItem
        If Not oListItem Is Nothing Then
            If oListItem.Tag = "group" Then
                DisplayGroup oCurrentGroups.item(oListItem.Text)
                If trvTree.SelectedItem.Key <> "LocalFolders" Then
                    UpdateTree
                Else 'if in localfolders
                    ConstructTree
                    trvTree_NodeClick trvTree.Nodes(2)
                End If
            Else 'if a record
                DisplayRecord oCurrentRecords.item(oListItem.Text)
            End If
            RefreshList
        End If
        
    End If
Exit Sub
ErrorHandler:
    MsgBox Err.Description, , "Show Properties"
    Refresh_All
    
End Sub



Private Sub tbToolBar_ButtonClick(ByVal Button As MSComctlLib.Button)
'Handles ToolBar clicks

    On Error GoTo ErrorHandler
    
    If Button.Caption = "Refresh" Then
        Refresh_All
    
    ElseIf Button.Caption = "Add Group" Then
        AddNewGroup
    
    ElseIf Button.Caption = "Add Record" Then
        AddNewRecord
                
    ElseIf Button.Caption = "Properties" Then
        ShowProperties
        
    ElseIf Button.Caption = "Delete" Then
    
        If ActiveControl Is trvTree Then
            DeleteFromTree
        ElseIf ActiveControl Is lstRecs Then
            DeleteFromList
        End If
        
    ElseIf Button.Caption = "Delete All" Then
        Delete_All
        
    ElseIf Button.Caption = "Exit" Then
        End
    
    End If

Exit Sub
ErrorHandler:
        MsgBox Err.Description, vbCritical, "Error"
End Sub

Sub Delete_All()
'Delete all items in the list view

    Dim i As Integer
    
    If (lstRecs.ListItems.Count > 0) Then
        
        If (MsgBox("Are you sure you want to delete " & lstRecs.ListItems.Count & " items?", vbYesNo, "Delete") = vbYes) Then
            
            ' delete all items
            If Not oCurrentRecords Is Nothing Then
                For i = 1 To oCurrentRecords.Count
                    oCurrentRecords.item(1).IsCheckedOut = False
                    oCurrentRecords.item(1).IsKeptFile = True
                    oCurrentRecords.ClearItem 1
                Next
            End If
            
            If Not oCurrentGroups Is Nothing Then
                For i = 1 To oCurrentGroups.Count
                    oCurrentGroups.item(1).IsKeptFolder = True
                    oCurrentGroups.ClearItem 1
                Next
            End If
            
            If trvTree.SelectedItem.Key <> "LocalFiles" And trvTree.SelectedItem.Key <> "LocalFolders" Then
                ClearChildNodes trvTree.SelectedItem
                AddGroupsToNode trvTree.SelectedItem, oCurrentGroups
            ElseIf trvTree.SelectedItem.Key = "LocalFolders" Then
                ConstructTree
                trvTree_NodeClick trvTree.Nodes("LocalFolders")
            End If
            
            RefreshList
            
        End If
    End If
    
End Sub

Sub AddNewGroup()
'Add new group to current Groups collection

    On Error GoTo ErrorHandler
    Dim oNewGroup As LocalGroup
    Set oNewGroup = New LocalGroup
    DisplayGroup oNewGroup
    If Not mbCancelAdd Then
        oCurrentGroups.AddItem oNewGroup
        
        If trvTree.SelectedItem.Key <> "LocalFolders" Then
            ClearChildNodes trvTree.SelectedItem
            AddGroupsToNode trvTree.SelectedItem, oCurrentGroups
            RefreshList
            
        Else
            
            ConstructTree
            trvTree_NodeClick trvTree.Nodes("LocalFolders") 'click Local Folders node
        End If
    End If
Exit Sub
ErrorHandler:
    MsgBox Err.Description, , "Error when adding new group"
End Sub

Function FindGroupByNode(oNode As MSComctlLib.Node) As LocalGroup
'this function finds a group by Node
'it does that by getthing the full path of the node, splitting it into
'an array, and going from the top to the buttom or the array
    On Error GoTo ErrorHandler
    
    If oNode.Key = "root" Or oNode.Key = "LocalFiles" Or oNode.Key = "LocalFolders" Then
        Set FindGroupByNode = Nothing
    Else
        Dim oTempGroup As LocalGroup, sPaths() As String, i As Integer
        sPaths() = Split(oNode.FullPath, "?")
        Set oTempGroup = oLocalDB.Root.item(sPaths(1))
        For i = 2 To UBound(sPaths)
            Set oTempGroup = oTempGroup.Groups.item(sPaths(i))
        Next
        Set FindGroupByNode = oTempGroup
    End If

Exit Function
ErrorHandler:
    MsgBox Err.Description + vbCrLf + "Selected group could have been deleted from another application", , "FindGroupByNode Function"
End Function



Private Sub trvTree_AfterLabelEdit(Cancel As Integer, NewString As String)
'renaming group in th tree view

    On Error GoTo ErrorHandler
    oCurrentGroup.Path = NewString
Exit Sub
ErrorHandler:
    MsgBox Err.Description, , "Can't rename this group"
    Cancel = True
End Sub

Private Sub trvTree_BeforeLabelEdit(Cancel As Integer)
'prevent the user from renaming "LocalFiles", "Root" or "LocalFolders"

    Dim sKey As String
    sKey = trvTree.SelectedItem.Key
    
    If sKey = "root" Or sKey = "LocalFiles" Or sKey = "LocalFolders" Then
        Cancel = True
    End If
End Sub

Private Sub trvTree_Collapse(ByVal Node As MSComctlLib.Node)
'when collapsing a node, set focus to it
    trvTree_NodeClick Node
End Sub

Private Sub trvTree_KeyDown(KeyCode As Integer, Shift As Integer)
'handle keyboard in tree view
    If KeyCode = vbKeyF5 Then
        Refresh_All
    ElseIf KeyCode = vbKeyDelete Then
        DeleteFromTree
    End If
End Sub

Private Sub trvTree_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
'Display context menu for tree view

    Dim sKey As String
    sKey = trvTree.SelectedItem.Key
    If Button = 2 Then
        If sKey <> "root" And sKey <> "LocalFiles" And sKey <> "LocalFolders" Then
            PopupMenu mnuGroup, vbPopupMenuRightButton, , , mnuGroupProps
        End If
    End If
End Sub

Private Sub trvTree_NodeClick(ByVal Node As MSComctlLib.Node)
'Find current Records and Groups collections, based on the selection in the tree view
    
    On Error GoTo ErrorHandler
    Set oCurrentGroups = Nothing
    Set oCurrentRecords = Nothing
    Set oCurrentGroup = Nothing
    
    If Node.Key = "root" Then 'Clicked on "LocalDB Root"
        Set oCurrentGroups = oLocalDB.Root
        
    ElseIf Node.Key = "LocalFiles" Then
        Set oCurrentRecords = oLocalDB.LocalFiles
        
    ElseIf Node.Key = "LocalFolders" Then
        Set oCurrentGroups = oLocalDB.LocalFolders
        
    Else 'clicked on a normal group
        Set oCurrentGroup = FindGroupByNode(Node)
        If Not oCurrentGroup Is Nothing Then
            Set oCurrentGroups = oCurrentGroup.Groups
            Set oCurrentRecords = oCurrentGroup.Records
        End If
    End If
    
    Node.Selected = True

    stbStatusBar.Panels("Path") = Replace(Node.FullPath, "?", " -> ")
    
    RefreshList

Exit Sub
ErrorHandler:
    MsgBox Err.Description, , "Node Click"
    Resume Next
End Sub

Sub DisplayRecord(oLocalRecord As LocalRecord)
'Displays LocalRecord properties form

    If Not oLocalRecord Is Nothing Then
        Dim frmRecord As New frmLocRec
        Set frmRecord.oEditedRecord = oLocalRecord
        mbCancelAdd = False
        frmRecord.Show vbModal
    End If
End Sub

Sub DisplayGroup(oLocalGroup As LocalGroup)
'Displays LocalGroup properties form
    If Not oLocalGroup Is Nothing Then
        Dim frmGroup As New frmLocGroup
        Set frmGroup.oLocalGroup = oLocalGroup
        mbCancelAdd = False
        frmGroup.Show vbModal
    End If
End Sub

Sub ClearChildNodes(oNode As MSComctlLib.Node)
'clear all children  nodes oNode

    Do Until oNode.Children = 0
        trvTree.Nodes.Remove (oNode.Child.Index)
    Loop
End Sub

Sub AddNewRecord()
'Adds new record to the current collection
    
    On Error GoTo ErrorHandler
    Dim oNewRecord As LocalRecord
    ' select a file with common dialog
    CommonDialog.FileName = ""
    CommonDialog.ShowOpen
    If CommonDialog.FileName <> "" Then
        
        ' initialize the new record
        Set oNewRecord = New LocalRecord
        oNewRecord.Path = CommonDialog.FileName
        
        ' add the new record to the current collection
        oCurrentRecords.AddItem oNewRecord
        
        ' allow the user to edit the other properties of the record
        DisplayRecord oNewRecord
    End If
    RefreshList
Exit Sub
ErrorHandler:
    MsgBox Err.Description, , "Adding Record"
End Sub

Sub DeleteFromList()
'Deletes selected items from list view

    On Error GoTo ErrorHandler
    
    Dim oDeletedRecord As LocalRecord
    Dim oDeletedGroup As LocalGroup
    Dim nSelectedItems As Integer
        
    ' count the selected items
    For Each oListItem In lstRecs.ListItems
        If oListItem.Selected Then
            nSelectedItems = nSelectedItems + 1
        End If
    Next
    
    If nSelectedItems > 0 Then
        
        If (MsgBox("Are you sure you want to delete " & nSelectedItems & " items?", vbYesNo, "Delete") = vbYes) Then
            ' delete the selected items
            For Each oListItem In lstRecs.ListItems
                If oListItem.Selected Then
                    If oListItem.Tag = "group" Then
                        Set oDeletedGroup = oCurrentGroups.item(oListItem.Text)
                        oDeletedGroup.IsKeptFolder = True
                        oCurrentGroups.ClearItem oDeletedGroup
                    Else
                        Set oDeletedRecord = oCurrentRecords.item(oListItem.Text)
                        oDeletedRecord.IsCheckedOut = False
                        oDeletedRecord.IsKeptFile = True
                        oCurrentRecords.ClearItem oDeletedRecord
                    End If
                 End If
            Next

            If trvTree.SelectedItem.Key = "LocalFolders" Then
                ConstructTree
                trvTree_NodeClick trvTree.Nodes("LocalFolders")
            ElseIf trvTree.SelectedItem.Key = "LocalFiles" Then
                'dont need to do anything
            ElseIf trvTree.SelectedItem.Key = "root" Then
                ConstructTree
                trvTree_NodeClick trvTree.Nodes("root")
            Else 'on normal group
                UpdateTree
            End If

            RefreshList

        End If
    End If
Exit Sub
ErrorHandler:
    MsgBox Err.Description, , "Delete"
End Sub
Sub DeleteFromTree()
'delete a group from tree view

    On Error GoTo ErrorHandler
    Dim oDeletedGroup As LocalGroup, oGroups As LocalGroups, oNode As Node, oParentNode As Node
    Set oNode = trvTree.SelectedItem
    If oNode.Key = "root" Or oNode.Key = "LocalFiles" Or oNode.Key = "LocalFolders" Then
        MsgBox "You cannot delete this item"
    Else
        If MsgBox("Are you sure you want to delte this group?", vbYesNo, "Delete Group Confirmation") = vbYes Then
            If oNode.Parent.Text = "LocalDB Root" Then
                Set oGroups = oLocalDB.Root
                Set oDeletedGroup = FindGroupByNode(oNode)
                oGroups.ClearItem oDeletedGroup
                ConstructTree
                trvTree_NodeClick trvTree.Nodes("root")
            Else
                Set oParentNode = oNode.Parent
                Set oGroups = FindGroupByNode(oParentNode).Groups
                Set oDeletedGroup = FindGroupByNode(oNode)
                oGroups.ClearItem oDeletedGroup
                            
                ClearChildNodes oParentNode
                AddGroupsToNode oParentNode, oGroups
                trvTree_NodeClick oParentNode
                Set oDeletedGroup = Nothing
            End If
        End If
    End If
    
Exit Sub
ErrorHandler:
    MsgBox Err.Description, , "Delete from TreeView"
End Sub


