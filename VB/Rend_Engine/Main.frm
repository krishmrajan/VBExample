VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComctl.ocx"
Begin VB.Form frmMain 
   Caption         =   "FileNET Rendition Services for HTML"
   ClientHeight    =   6180
   ClientLeft      =   4035
   ClientTop       =   2895
   ClientWidth     =   7350
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6180
   ScaleWidth      =   7350
   Begin MSComctlLib.StatusBar stbStatus 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   5925
      Width           =   7350
      _ExtentX        =   12965
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7241
            MinWidth        =   1764
            Key             =   "text"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   1773
            MinWidth        =   1764
            Text            =   "Running"
            TextSave        =   "Running"
            Key             =   "status"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            Object.Width           =   1588
            MinWidth        =   1587
            TextSave        =   "6:04 PM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            Object.Width           =   1773
            MinWidth        =   1764
            TextSave        =   "12/1/99"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView trvTree 
      Height          =   2175
      Left            =   0
      TabIndex        =   1
      Top             =   600
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   3836
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   441
      Style           =   7
      Appearance      =   1
   End
   Begin MSComctlLib.Toolbar tlbToolBar 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   7350
      _ExtentX        =   12965
      _ExtentY        =   1058
      ButtonWidth     =   1588
      ButtonHeight    =   953
      Wrappable       =   0   'False
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   10
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Pause"
            Key             =   "Pause"
            Object.ToolTipText     =   "Pauses or resumes services"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Add Library"
            Key             =   "Add Library"
            Object.ToolTipText     =   "Adds a new library to the rendition services"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Add ST"
            Key             =   "Add ST"
            Object.ToolTipText     =   "Adds new style template to the selected library"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Properties"
            Key             =   "Properties"
            Object.ToolTipText     =   "Displays properties for the selected item"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Remove"
            Key             =   "Remove"
            Object.ToolTipText     =   "Removes selected item from rendition services"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Exit"
            Key             =   "Exit"
            Object.ToolTipText     =   "Exits the application"
         EndProperty
      EndProperty
   End
   Begin VB.Timer tmrTimer 
      Interval        =   2000
      Left            =   6600
      Top             =   2760
   End
   Begin MSComctlLib.ImageList imgImages 
      Left            =   5760
      Top             =   2640
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":0442
            Key             =   "template"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":059C
            Key             =   "library"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":09EE
            Key             =   "pause"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":0F40
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":1052
            Key             =   "delete_all"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":15A4
            Key             =   "exit"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":1AF6
            Key             =   "delete"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":1C08
            Key             =   "engine"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":205A
            Key             =   "refresh"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":24AC
            Key             =   "add"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":28FE
            Key             =   "hood"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":3A48
            Key             =   "properties"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":3F9A
            Key             =   "start"
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtMessages 
      Height          =   2775
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   3120
      Width           =   7335
   End
   Begin VB.Label lblStatusMessages 
      AutoSize        =   -1  'True
      Caption         =   "Status Messages"
      Height          =   195
      Left            =   0
      TabIndex        =   3
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Menu mnuLibrary 
      Caption         =   "&Library"
      Begin VB.Menu mnuAddLibrary 
         Caption         =   "Add Library..."
      End
      Begin VB.Menu mnuModifyLibrarySettings 
         Caption         =   "Modify Library Settings..."
      End
      Begin VB.Menu mnuRemoveLibrary 
         Caption         =   "Remove Library"
      End
      Begin VB.Menu mnuRemoveRendServices 
         Caption         =   "Remove Rendition Services"
      End
      Begin VB.Menu mnuSeparator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAddTemplate 
         Caption         =   "Add Template..."
      End
      Begin VB.Menu mnuModifyTemplate 
         Caption         =   "Modify Template..."
      End
      Begin VB.Menu mnuRemoveTemplate 
         Caption         =   "Remove Template"
      End
      Begin VB.Menu mnuSeparator2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewRefresh 
         Caption         =   "Refresh"
      End
      Begin VB.Menu mnuViewClearMessages 
         Caption         =   "Clear Messages"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpTopics 
         Caption         =   "Help Topics"
      End
      Begin VB.Menu mnuSeparator9 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "About Rendition Engine..."
      End
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "POPUP"
      Visible         =   0   'False
      Begin VB.Menu mnuProperties 
         Caption         =   "Properties..."
      End
      Begin VB.Menu mnuPopupAddSTyleTemplate 
         Caption         =   "Add Style Template..."
      End
      Begin VB.Menu mnuSeparator3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRemove 
         Caption         =   "Remove"
      End
   End
End
Attribute VB_Name = "frmMain"
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
' Workfile:   frmMain.frm

'----------------------------------------------------------------------
' All Rights Reserved.  Copyright (c) 1999 FileNet Corp.
'----------------------------------------------------------------------

Option Explicit
Private oCurServedLib As ServicedLib
Private oCurrentTemplate As StyleTemplate

Private Sub Form_Load()
'Here's what this application does when it first starts up.
'1. Creates 'Published' and 'Source' sub folders, if they don't exists
'2. Attempts to load it's settings from the ini file, file name is defined in INI_FILE constand
'3. Sets up pictures on the controls
'4. Draws the configuration in the tree view, displaying libraries and templates
'5. Centers the form

    Randomize Timer
    
    Set oServicedLibraries = New Collection
    Set oFS = New FileSystemObject
    Set oHood = New Neighborhood

    CreateDirectories
    
    LoadLibrarySettings
    
    'set up controls' images
    Set trvTree.ImageList = imgImages
    Set tlbToolBar.ImageList = imgImages
    tlbToolBar.Buttons("Exit").Image = "exit"
    tlbToolBar.Buttons("Properties").Image = "properties"
    tlbToolBar.Buttons("Pause").Image = "pause"
    tlbToolBar.Buttons("Remove").Image = "delete"
    tlbToolBar.Buttons("Add Library").Image = "library"
    tlbToolBar.Buttons("Add ST").Image = "template"
    
    LoadTree
    
    CenterForm Me
    
End Sub

Private Sub Form_Resize()
'Resize the controls properly when the form is resized

    On Error GoTo ErrorHandler

    txtMessages.Width = Me.Width - 120
    txtMessages.Height = Me.ScaleHeight - 3390
    trvTree.Width = Me.Width - 120
Exit Sub
ErrorHandler:
    'An error could happen when making the window too small, but there is really
    'nothing worth doing here...
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Save settins and Logoff all libraries when the Application is terminated
    Dim oTempSL As ServicedLib, i As Integer
    
    If Not bCantQuitNow Then
        SaveLibrarySettings
        
        'this will destroy all ServicedLib objects
        For i = 1 To oServicedLibraries.Count
            Set oTempSL = oServicedLibraries(1)
            oServicedLibraries.Remove 1
            Set oTempSL = Nothing
        Next
        
        Debug.Print "The program has ended"
        
        End 'end the program!
    Else
        MsgBox "Job is in progress, you cannot quit now!  Press Pause button, and wait till the job is finished."
        Cancel = True
    End If
    
Exit Sub
ErrorHandler:
    ShowError "Form_Unload", True
End Sub




Private Sub mnuAddLibrary_Click()
'Displays "Add Library" dialog

    Dim oNewLib As ServicedLib
    
    If oServicedLibraries.Count = oHood.Libraries.Count Then
        MsgBox "You don't have any more libraries you can add.  Use Configure App to add more libraries"
    Else
        Set oNewLib = New ServicedLib
        If oNewLib.ShowPropertiesDialog(True) = idmDialogExitOK Then
            oNewLib.InitializeRenditions
            If oNewLib.IsWorkingOK Then
                oServicedLibraries.Add oNewLib
                SaveLibrarySettings
                If oNewLib.StyleTemplates.Count = 0 Then
                    MsgBox "You should  add at least one Style Template to be able to publish documents using this Rendition Engine"
                    oNewLib.AddNewStyleTemplate
                End If
                LoadTree
                
            End If
        Else
            Set oNewLib = Nothing
        End If
    End If

End Sub

Private Sub mnuAddTemplate_Click()
'Adds style template
    If Not oCurServedLib Is Nothing Then
        oCurServedLib.AddNewStyleTemplate
        LoadTree
    End If
End Sub


Private Sub mnuExit_Click()
    Form_Unload 1
End Sub

Private Sub mnuHelpAbout_Click()
'Shows About dialog box
    frmAbout.Show vbModal
End Sub

Private Sub mnuHelpTopics_Click()
    MsgBox "Sorry, help is not implemented."
End Sub

Private Sub mnuModifyLibrarySettings_Click()
'Displays Library Settings dialog
    If Not oCurServedLib Is Nothing Then
        If oCurServedLib.ShowPropertiesDialog(False) = idmDialogExitOK Then
            oCurServedLib.InitializeRenditions
            LoadTree
        End If
    End If
End Sub

Private Sub mnuModifyTemplate_Click()
'Displays Style Template properties dialog
    If Not oCurrentTemplate Is Nothing Then
        DisplayStyleTemplate oCurrentTemplate
        trvTree.SelectedItem.Text = oCurrentTemplate.Name
    End If
End Sub

Private Sub mnuPopupAddSTyleTemplate_Click()
'Adds style template
    mnuAddTemplate_Click
End Sub

Private Sub mnuProperties_Click()
'Displays properties dialog for currently selected item
    tlbToolBar_ButtonClick tlbToolBar.Buttons("Properties")
End Sub

Private Sub mnuRemove_Click()
'Removes selected item
    tlbToolBar_ButtonClick tlbToolBar.Buttons("Remove")
End Sub

Private Sub mnuRemoveLibrary_Click()
'Remove library
    If MsgBox("Are you sure you want to remove this library from rendition services?", vbYesNo + vbQuestion, "Warning") = vbYes Then
        If Not oCurServedLib Is Nothing Then
            oServicedLibraries.Remove trvTree.SelectedItem.Tag
            Set oCurServedLib = Nothing
            LoadTree
            SaveLibrarySettings
        End If
    End If
End Sub

Private Sub mnuRemoveRendServices_Click()
'Remove library and delete all Rendition objects (RenditionEngine and StyleTemplate(s))
    If MsgBox("Are you sure you want to remove rendition services objects from this library?", vbYesNo + vbQuestion, "Warning") = vbYes Then
        If Not oCurServedLib Is Nothing Then
            oCurServedLib.RemoveRenditionServices
            oServicedLibraries.Remove trvTree.SelectedItem.Tag
            Set oCurServedLib = Nothing
            LoadTree
            SaveLibrarySettings
        End If
    End If
End Sub

Private Sub mnuRemoveTemplate_Click()
'Removes style template
    If Not oCurrentTemplate Is Nothing Then
        If MsgBox("You sure you want to remove this Style Template ?", vbYesNo + vbQuestion) = vbYes Then
            oCurrentTemplate.Delete
            LoadTree
        End If
    End If
End Sub

Private Sub mnuViewClearMessages_Click()
'Clears Status Messages text box
    txtMessages.Text = ""
    DoEvents
End Sub

Sub PAUSE_SERVICES()
    tmrTimer.Enabled = False
    tlbToolBar.Buttons("Pause").Caption = "Start"
    tlbToolBar.Buttons("Pause").Image = "start"
    stbStatus.Panels("status").Text = "Paused"
End Sub

Sub START_SERVICES()
    tmrTimer.Enabled = True
    tlbToolBar.Buttons("Pause").Caption = "Pause"
    tlbToolBar.Buttons("Pause").Image = "pause"
    stbStatus.Panels("status").Text = "Running"
End Sub

Private Sub mnuViewRefresh_Click()
'Refreshes the tree view
    LoadTree
End Sub

Private Sub tlbToolBar_ButtonClick(ByVal Button As MSComctlLib.Button)
'Handle toolbar button clicks

    If Button.Key = "Pause" Then
    
        If tmrTimer.Enabled Then
            PAUSE_SERVICES
        Else
            START_SERVICES
        End If
    
    ElseIf Button.Key = "Exit" Then
        Form_Unload 1
    
    ElseIf Button.Key = "Add Library" Then
        mnuAddLibrary_Click
        
    ElseIf Button.Key = "Add ST" Then
        mnuAddTemplate_Click
    
    ElseIf Button.Key = "Properties" Then
    
        If Not trvTree.SelectedItem Is Nothing Then
            If trvTree.SelectedItem.Image = "template" Then
                mnuModifyTemplate_Click
            ElseIf trvTree.SelectedItem.Image = "library" Then
                mnuModifyLibrarySettings_Click
            End If
        End If
        
    ElseIf Button.Key = "Remove" Then
    
        If Not trvTree.SelectedItem Is Nothing Then
            If trvTree.SelectedItem.Image = "template" Then
                mnuRemoveTemplate_Click
            ElseIf trvTree.SelectedItem.Image = "library" Then
                mnuRemoveLibrary_Click
            End If
        End If

    Else
        MsgBox "This button is not handled"
    
    End If
    
End Sub

Private Sub tmrTimer_Timer()
'CheckLibraryQueue for every library, when timer event occurs

    Dim oServedLib As ServicedLib
    For Each oServedLib In oServicedLibraries
        oServedLib.CheckLibraryQueue
    Next
End Sub

Sub LoadTree()
'Constucts the tree view

    On Error GoTo ErrorHandler
    Dim oLibNode As Node, i As Integer, oTempServedLib As ServicedLib
    
    trvTree.Nodes.Clear
    trvTree.Nodes.Add(, , "root", "FileNET Rendition Services", "hood").Expanded = True
    
    For i = 1 To oServicedLibraries.Count
        Set oTempServedLib = oServicedLibraries(i)
        Set oLibNode = trvTree.Nodes.Add("root", tvwChild, , oTempServedLib.sLibraryName, "library")
        oLibNode.Tag = i
        oLibNode.Expanded = True
        AddStyleTemplates oLibNode, oTempServedLib
    Next
    
    If oServicedLibraries.Count = 0 Then
        trvTree.Nodes.Add "root", tvwChild, , "No Libraries Configured", "delete"
    End If
    
    trvTree_NodeClick trvTree.Nodes("root")
    trvTree.Nodes("root").Selected = True
    
Exit Sub
ErrorHandler:
    ShowError "LOAD TREE"
    Resume Next
End Sub

Sub AddStyleTemplates(oNode As Node, oServedLibrary As ServicedLib)
'Adds style templates for the given library to the tree view node
    Dim oST As StyleTemplate
    For Each oST In oServedLibrary.StyleTemplates
        trvTree.Nodes.Add(oNode, tvwChild, , oST.Name, "template").Tag = oST.ID
    Next
End Sub


Private Sub trvTree_AfterLabelEdit(Cancel As Integer, NewString As String)
'Enables renaming style templates in the tree view
    On Error GoTo ErrorHandler
    oCurrentTemplate.Name = NewString
    oCurrentTemplate.Save
Exit Sub
ErrorHandler:
    Cancel = True
    ShowError "Renaming Style Template", True
End Sub

Private Sub trvTree_BeforeLabelEdit(Cancel As Integer)
'Prevents nodes other then templates from being renamed

    If oCurrentTemplate Is Nothing Then
        Cancel = True
    End If
End Sub


Private Sub trvTree_DblClick()
'Displays Style Template properties on double click event
    If trvTree.SelectedItem.Image = "template" Then
        mnuProperties_Click
    End If
End Sub

Private Sub trvTree_KeyDown(KeyCode As Integer, Shift As Integer)
'Handles delete and F5 keys.  They do what you'd expect.

    If KeyCode = vbKeyDelete Then
        tlbToolBar_ButtonClick tlbToolBar.Buttons("Remove")
    ElseIf KeyCode = vbKeyF5 Then
        LoadTree
    End If

End Sub

Private Sub trvTree_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
'Displays pop-up menu

    If Button = 2 Then
        If trvTree.SelectedItem.Image <> "hood" Then
            If trvTree.SelectedItem.Image = "library" Then
                mnuPopupAddSTyleTemplate.Visible = True
            Else
                mnuPopupAddSTyleTemplate.Visible = False
            End If
            
            PopupMenu mnuPopup, vbPopupMenuRightButton
        End If
    End If
End Sub

Private Sub trvTree_NodeClick(ByVal Node As MSComctlLib.Node)
'Sets active library or style template, disables or enables toolbar buttons accordinly
    
    On Error GoTo ErrorHandler
    Set oCurrentTemplate = Nothing
    Set oCurServedLib = Nothing
    
    If Node.Image = "template" Then
        Set oCurServedLib = oServicedLibraries(Node.Parent.Tag)
        Set oCurrentTemplate = oCurServedLib.StyleTemplates(Node.Tag)
        ButtonsEnabled = True
    ElseIf Node.Image = "library" Then
        Set oCurServedLib = oServicedLibraries(Node.Tag)
        ButtonsEnabled = True
    Else
        ButtonsEnabled = False
    End If
    
Exit Sub
ErrorHandler:
    ShowError "Tree View Click Event"
    Set oCurServedLib = Nothing
    Set oCurrentTemplate = Nothing
End Sub

Property Let ButtonsEnabled(bEnable As Boolean)
'Enable or disable toolbar buttons
    tlbToolBar.Buttons("Remove").Enabled = bEnable
    tlbToolBar.Buttons("Add ST").Enabled = bEnable
    tlbToolBar.Buttons("Properties").Enabled = bEnable
End Property
