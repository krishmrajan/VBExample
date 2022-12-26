VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{2A601BA0-B880-11CF-8185-444553540000}#3.0#0"; "fnlist.ocx"
Object = "{54E6FE61-B93A-11CF-8185-444553540000}#3.0#0"; "fntree.ocx"
Object = "{A9983B40-CE52-11CF-AE75-00A0248802BA}#3.0#0"; "fnviewer.ocx"
Begin VB.Form frmMain 
   Caption         =   "Sample Explorer Application"
   ClientHeight    =   5985
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   6180
   LinkTopic       =   "Form1"
   ScaleHeight     =   5985
   ScaleWidth      =   6180
   StartUpPosition =   3  'Windows Default
   Tag             =   "01900"
   Begin ComctlLib.Toolbar tbrToolBar 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   6180
      _ExtentX        =   10901
      _ExtentY        =   741
      ButtonWidth     =   635
      ButtonHeight    =   572
      Wrappable       =   0   'False
      Appearance      =   1
      ImageList       =   "imlIcons"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   14
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Back"
            Object.ToolTipText     =   "1036"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Forward"
            Object.ToolTipText     =   "1037"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Cut"
            Object.ToolTipText     =   "1038"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Copy"
            Object.ToolTipText     =   "1039"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Paste"
            Object.ToolTipText     =   "1040"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Delete"
            Object.ToolTipText     =   "1041"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Properties"
            Object.ToolTipText     =   "1042"
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "ViewLarge"
            Object.ToolTipText     =   "1043"
            Object.Tag             =   ""
            ImageIndex      =   8
            Style           =   2
         EndProperty
         BeginProperty Button12 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "ViewSmall"
            Object.ToolTipText     =   "1044"
            Object.Tag             =   ""
            ImageIndex      =   9
            Style           =   2
         EndProperty
         BeginProperty Button13 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "ViewList"
            Object.ToolTipText     =   "1045"
            Object.Tag             =   ""
            ImageIndex      =   10
            Style           =   2
         EndProperty
         BeginProperty Button14 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "ViewDetails"
            Object.ToolTipText     =   "1046"
            Object.Tag             =   ""
            ImageIndex      =   11
            Style           =   2
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picSplitterH 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      FillColor       =   &H00808080&
      Height          =   72
      Left            =   0
      ScaleHeight     =   32.658
      ScaleMode       =   0  'User
      ScaleWidth      =   55754.54
      TabIndex        =   6
      Top             =   3480
      Visible         =   0   'False
      Width           =   5364
   End
   Begin VB.PictureBox picSplitter 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      FillColor       =   &H00808080&
      Height          =   4812
      Left            =   2205
      ScaleHeight     =   2096.691
      ScaleMode       =   0  'User
      ScaleWidth      =   780
      TabIndex        =   5
      Top             =   720
      Visible         =   0   'False
      Width           =   72
   End
   Begin IDMViewerCtrl.IDMViewerCtrl vwrViewer 
      Height          =   1905
      Left            =   2310
      TabIndex        =   10
      Top             =   3570
      Width           =   2955
      _Version        =   196608
      _ExtentX        =   5212
      _ExtentY        =   3360
      _StockProps     =   161
      Appearance      =   1
      SystemType      =   -9940
   End
   Begin IDMListView.IDMListView ilvIDMListView 
      Height          =   2430
      Left            =   2310
      TabIndex        =   9
      Top             =   735
      Width           =   2955
      _Version        =   196608
      _ExtentX        =   5212
      _ExtentY        =   4286
      _StockProps     =   239
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
   Begin IDMTreeView.IDMTreeView itvIDMTreeView 
      Height          =   2430
      Left            =   0
      TabIndex        =   8
      Top             =   735
      Width           =   2010
      _Version        =   196608
      _ExtentX        =   3545
      _ExtentY        =   4286
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
   Begin VB.PictureBox picTitles 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   6180
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   420
      Width           =   6180
      Begin VB.Label lblTitle 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " ListView:"
         Height          =   276
         Index           =   1
         Left            =   2040
         TabIndex        =   4
         Tag             =   "1048"
         Top             =   0
         Width           =   3216
      End
      Begin VB.Label lblTitle 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " TreeView:"
         Height          =   276
         Index           =   0
         Left            =   0
         TabIndex        =   3
         Tag             =   "1047"
         Top             =   0
         Width           =   2016
      End
   End
   Begin ComctlLib.StatusBar sbrStatusBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   5715
      Width           =   6180
      _ExtentX        =   10901
      _ExtentY        =   476
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   4
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   3440
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
            TextSave        =   "5/26/99"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1588
            MinWidth        =   1587
            TextSave        =   "15:06"
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   5568
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picPicture 
      Height          =   1905
      Left            =   0
      ScaleHeight     =   1845
      ScaleWidth      =   1845
      TabIndex        =   7
      Top             =   3570
      Width           =   1905
   End
   Begin VB.Image imgSplitterH 
      Height          =   150
      Left            =   0
      MousePointer    =   7  'Size N S
      Top             =   3255
      Width           =   5355
   End
   Begin VB.Image imgSplitter 
      Height          =   4788
      Left            =   2040
      MousePointer    =   9  'Size W E
      Top             =   720
      Width           =   156
   End
   Begin ComctlLib.ImageList imlIcons 
      Left            =   5568
      Top             =   1368
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   11
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":0018
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":056A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":0ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":100E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":1560
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":1AB2
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":2004
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":2556
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":2AA8
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":2FFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":354C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "1000"
      Begin VB.Menu mnuFileOpen 
         Caption         =   "1001"
      End
      Begin VB.Menu mnuFileFind 
         Caption         =   "1002"
      End
      Begin VB.Menu mnuFileBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSendTo 
         Caption         =   "1003"
      End
      Begin VB.Menu mnuFileBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileNew 
         Caption         =   "1004"
      End
      Begin VB.Menu mnuFileBar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileDelete 
         Caption         =   "1005"
      End
      Begin VB.Menu mnuFileRename 
         Caption         =   "1006"
      End
      Begin VB.Menu mnuFileProperties 
         Caption         =   "1007"
      End
      Begin VB.Menu mnuFileBar4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileMRU 
         Caption         =   ""
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileMRU 
         Caption         =   ""
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileMRU 
         Caption         =   ""
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileMRU 
         Caption         =   ""
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileBar5 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "1008"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "1009"
      Begin VB.Menu mnuEditUndo 
         Caption         =   "1010"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuEditBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditCut 
         Caption         =   "1011"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "1012"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "1013"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuEditPasteSpecial 
         Caption         =   "1014"
      End
      Begin VB.Menu mnuEditBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditSelectAll 
         Caption         =   "1015"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEditInvertSelection 
         Caption         =   "1016"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "1017"
      Begin VB.Menu mnuViewToolbar 
         Caption         =   "1018"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewStatusBar 
         Caption         =   "1019"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuListViewMode 
         Caption         =   "1020"
         Index           =   0
      End
      Begin VB.Menu mnuListViewMode 
         Caption         =   "1021"
         Index           =   1
      End
      Begin VB.Menu mnuListViewMode 
         Caption         =   "1022"
         Index           =   2
      End
      Begin VB.Menu mnuListViewMode 
         Caption         =   "1023"
         Index           =   3
      End
      Begin VB.Menu mnuViewBar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewArrangeIcons 
         Caption         =   "1024"
         Begin VB.Menu mnuVAIByDate 
            Caption         =   "1025"
         End
         Begin VB.Menu mnuVAIByName 
            Caption         =   "1026"
         End
         Begin VB.Menu mnuVAIByType 
            Caption         =   "1027"
         End
         Begin VB.Menu mnuVAIBySize 
            Caption         =   "1028"
         End
      End
      Begin VB.Menu mnuViewLineUpIcons 
         Caption         =   "1029"
      End
      Begin VB.Menu mnuViewBar4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewRefresh 
         Caption         =   "1030"
      End
      Begin VB.Menu mnuViewOptions 
         Caption         =   "1031"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "1032"
      Begin VB.Menu mnuHelpContents 
         Caption         =   "1033"
      End
      Begin VB.Menu mnuHelpSearch 
         Caption         =   "1034"
      End
      Begin VB.Menu mnuHelpBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "1035"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const NAME_COLUMN = 0
Const TYPE_COLUMN = 1
Const SIZE_COLUMN = 2
Const DATE_COLUMN = 3
Private Declare Function OSWinHelp% Lib "user32" Alias "WinHelpA" (ByVal hwnd&, ByVal HelpFile$, ByVal wCommand%, dwData As Any)
  
Dim mbMoving As Boolean
Const sglSplitLimit = 500
Dim lsOriginalFormCaption As String

Private Sub Form_Load()

Dim loLibrary As IDMObjects.Library

On Error GoTo ErrorHandler

    gbSuccess = False

    LoadResStrings Me
    lsOriginalFormCaption = Me.Caption
    
    Me.Left = GetSetting(App.Title, "Settings", "MainLeft", 1000)
    Me.Top = GetSetting(App.Title, "Settings", "MainTop", 1000)
    Me.Width = GetSetting(App.Title, "Settings", "MainWidth", 6500)
    Me.Height = GetSetting(App.Title, "Settings", "MainHeight", 6500)
    'Next 3 lines are for splitter bars and list view view
    imgSplitter.Left = GetSetting(App.Title, "Settings", "VerticalSplit", 1500)
    imgSplitterH.Top = GetSetting(App.Title, "Settings", "HorizontalSplit", 1500)
    ilvIDMListView.View = Val(GetSetting(App.Title, "Settings", "ViewMode", "0"))
    
    For Each loLibrary In goNeighborhood.Libraries
        If loLibrary.GetState(idmLibraryLoggedOn) Then
            gbSuccess = AddToIDMTreeView(itvIDMTreeView, loLibrary, True)
            If gbSuccess = False Then
                MsgBox LoadResString(GI_ERR_UNABLE_TO_ADD_TO_TREEVIEW), vbExclamation, LoadResString(GI_ERR_ERROR)
            End If
        End If
    Next
        
Exit Sub

ErrorHandler:

    'Display Error message
    DisplayErrorMessage ("frmMain - Form_Load")

    'Cleanup Error values
    CleanupErrorCodes
    
    Resume Next
        
End Sub

Private Sub Form_Paint()

    tbrToolBar.Buttons(ilvIDMListView.View + LISTVIEW_BUTTON).Value = tbrPressed
    mnuListViewMode(ilvIDMListView.View).Checked = True

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer

    'close all sub forms
    For i = Forms.Count - 1 To 1 Step -1
        Unload Forms(i)
    Next
    If Me.WindowState <> vbMinimized Then
        SaveSetting App.Title, "Settings", "MainLeft", Me.Left
        SaveSetting App.Title, "Settings", "MainTop", Me.Top
        SaveSetting App.Title, "Settings", "MainWidth", Me.Width
        SaveSetting App.Title, "Settings", "MainHeight", Me.Height
        'Next 2 lines save splitter bar settings
        SaveSetting App.Title, "Settings", "VerticalSplit", imgSplitter.Left
        SaveSetting App.Title, "Settings", "HorizontalSplit", imgSplitterH.Top
    End If
    SaveSetting App.Title, "Settings", "ViewMode", ilvIDMListView.View

    'Because this is the main form, call AppTerminate to clean up
    AppTerminate True
    
End Sub

Private Sub ilvIDMListView_DblClick()

On Error Resume Next

    Select Case giCurSelListObjType
        Case idmObjTypeDocument
            'It is a document object
            'Tell it to display its contents in the Viewer control window
            vwrViewer.Visible = True
            vwrViewer.Document = goCurSelListItem
        Case idmObjTypeFolder
            'It is a folder object
            'Set selection on TreeView Control to this same item that was
            'double-clicked on in the ListView Control
            
            itvIDMTreeView.SelectChildItem goCurSelListItem.Label
    End Select
    
End Sub

Private Sub ilvIDMListView_DoBackgroundDragDrop(ByVal Source As Variant, Handled As Boolean)
'This subroutine handles a background drag and drop.
    
On Error GoTo ErrorHandler

    If IsObject(Source) Then
        Dim loItem As Object
        Set loItem = Source
        ' if item is a folder, then find parent folder (or library) on TreeView and
        ' "Copy" the folder into the parent object.
        If loItem.ObjectType = idmObjTypeFolder Then
            If giCurSelTreeObjType = idmObjTypeFolder Then
                MouseWait
                loItem.Copy goCurSelTreeItem
                Handled = True
                ' There are some refresh problems remaining here!
                Call ForceTreeRefresh(itvIDMTreeView, goCurSelTreeItem.Label)
                goCurSelTreeItem.Refresh idmFolderRefreshAll
                itvIDMTreeView.Expand
                MouseNormal
            Else
                MsgBox LoadResString(GI_ERR_UNABLE_TO_COPY), vbExclamation, LoadResString(GI_ERR_ERROR)
            End If
        Else
            ' The dragged object is a document
            If giCurSelTreeObjType = idmObjTypeFolder Then
                MouseWait
                goCurSelTreeItem.File loItem
                Handled = True
                goCurSelTreeItem.Refresh idmFolderRefreshAll
                MouseNormal
            Else
                MsgBox LoadResString(GI_ERR_UNABLE_TO_COPY), vbExclamation, LoadResString(GI_ERR_ERROR)
            End If
        End If
    ElseIf TypeName(Source) = "String" Then
        Dim loNewDocument As IDMObjects.Document
        Set loNewDocument = goCurSelTreeItem.Library.CreateObject(idmObjTypeDocument, GS_DOC_CLASS)
        loNewDocument.SaveNew Source, idmDocSaveNewWithUIWizard
        goCurSelTreeItem.File loNewDocument
        itvIDMTreeView.Refresh
        Call FillIDMListView(ilvIDMListView, goCurSelTreeItem, giCurSelTreeObjType)
        ilvIDMListView.Refresh
    End If

    Exit Sub

ErrorHandler:
    
    'Display Error Message - pass the name of this subroutine/function
    DisplayErrorMessage ("frmMain.ilvIDMListView_DoBackgroundDragDrop")
    
    'Cleanup Error values
    CleanupErrorCodes

    Resume Next

End Sub

Private Sub ilvIDMListView_DoBackgroundDragOver(ByVal Source As Variant, ByVal State As Integer, AllowDrop As Boolean)
'This routine handles events for items that are dragged over the IDMListView.
' State = 0 - DragEnter, 1 - DragLeave, or 2 - DragOver, this is
' similar to VB's DragOver state definition.
    
' On DragEnter state, check currently selected object in TreeView to see if it's
' a folder.  If it is, then a drop is allowed.  If it's a library, then a drop is
' allowed only if the source item is a folder type.  Documents cannot be dropped
' into a library.
    
    If State = 0 Then  'DragEnter
        If giCurSelTreeObjType = idmObjTypeFolder Then
            AllowDrop = True
        ElseIf giCurSelTreeObjType = idmObjTypeLibrary Then
            If IsObject(Source) Then
                Dim loItem As Object
                Set loItem = Source
                If loItem.ObjectType = idmObjTypeFolder Then
                    AllowDrop = True
                Else
                    AllowDrop = False
                End If
            ElseIf TypeName(Source) = "String" Then
                AllowDrop = True
            End If
        End If
    ' ElseIf State = 1 Or State = 2 then
    ' Do other actions here if necessary.  However, changing the AllowDrop flag
    ' in these states will not have any effect on the drop effect.
    End If

End Sub
Private Sub ilvIDMListView_DoItemDelete(ByVal Item As Object, ByVal Key As Long, Handled As Boolean)
'This subroutine processes the deletion of the selected item in the List View.
'But if it is a document on an IDMIS or IDMDS, we will just unfile it - not delete it!
'The UI shows the user deleting the object, but the foundation object
'must be updated as well.
    
Dim liUserResponse As Integer
Dim loItem As Object
Dim loMyFolder As IDMObjects.Folder

On Error GoTo ErrorHandler

    Set loItem = Item
    If loItem.ObjectType = idmObjTypeFolder Then
        'Check to see if user has security to delete
        Set loMyFolder = Item
        If loItem.GetState(idmFolderCanDelete) = True Then
            liUserResponse = MsgBox("'" & loMyFolder.Label & "' " & _
                LoadResString(GI_CONFIRM_FDELETE), vbQuestion + vbYesNo, LoadResString(GI_CONFIRM))
            Select Case liUserResponse
                Case vbNo
                    'User cancelled
                    Handled = False
                Case vbYes
                    MouseWait
                    ' We actually want to treat this as an 'unfile' operation, because we don't
                    ' want to delete docs this easily; so call recursive routine to clean out
                    ' docs from this and all child folders...
                    Call CleanOutFolder(loMyFolder)
                    Item.Delete
                    Handled = True
                    Call ForceTreeRefresh(itvIDMTreeView, _
                        goCurSelTreeItem.Label)
                    ilvIDMListView.Refresh
                    itvIDMTreeView.Refresh
                    MouseNormal
            End Select
        Else
            MsgBox LoadResString(GI_ERR_UNABLE_TO_DELETE), vbExclamation, LoadResString(GI_ERR_ERROR)
        End If
    ElseIf loItem.ObjectType = idmObjTypeDocument Then
        'Check to see if user has security to delete
        If loItem.GetState(idmDocCanDelete) = True Then
            liUserResponse = MsgBox(LoadResString(GI_CONFIRM_DDELETE), vbQuestion + vbYesNo, LoadResString(GI_CONFIRM))
            Select Case liUserResponse
                Case vbNo
                    'User cancelled
                    Handled = False
                Case vbYes
                    MouseWait
                    If giCurSelTreeObjType = idmObjTypeFolder Then
                        Set loMyFolder = goCurSelTreeItem
                        Call loMyFolder.Unfile(Item)
                    End If
                    Handled = True
                    ilvIDMListView.Refresh
                    MouseNormal
            End Select
        Else
            MsgBox LoadResString(GI_ERR_UNABLE_TO_DELETE), vbExclamation, LoadResString(GI_ERR_ERROR)
        End If
    Else
        MsgBox LoadResString(GI_ERR_UNABLE_TO_DELETE), vbExclamation, LoadResString(GI_ERR_ERROR)
    End If
    
    Exit Sub
    
ErrorHandler:

    'Display Error message
    DisplayErrorMessage ("frmMain.ilvIDMListView_DoItemDelete")

    'Cleanup Error values
    CleanupErrorCodes
    
    Handled = False

    Resume Next

End Sub

Private Sub ilvIDMListView_DoLabelEdit(ByVal Item As Object, ByVal Key As Long, ByVal Label As String, Handled As Boolean)
'This subroutine processes the name change of a list view item
'only for folders.
'The UI shows the user changing the label, but the foundation object
'must be updated as well.

Dim loFolder As IDMObjects.Folder
    
On Error GoTo ErrorHandler

    If Item.ObjectType = idmObjTypeFolder Then
        'a folder was selected so rename it
        Set loFolder = Item
        'Check to see if user has security to modify
        If loFolder.GetState(idmFolderCanModify) = True Then
            If loFolder.Name <> Label And Label <> "" And ilvIDMListView.LabelExist(Label) = False Then
                ilvIDMListView.RenameItem Key, Label
                loFolder.Name = Label
                loFolder.Save
                ilvIDMListView.Refresh
                goCurSelTreeItem.Refresh idmFolderRefreshSubFolders
                itvIDMTreeView.Refresh
                Handled = True
            End If
        Else
            MsgBox LoadResString(GI_ERR_UNABLE_TO_RENAME), vbExclamation, LoadResString(GI_ERR_ERROR)
        End If
    Else
        MsgBox LoadResString(GI_ERR_UNABLE_TO_RENAME), vbExclamation, LoadResString(GI_ERR_ERROR)
    End If
    
    Exit Sub
    
ErrorHandler:

    'Display Error message
    DisplayErrorMessage ("frmMain.ilvIDMListView_DoLabelEdit")

    'Cleanup Error values
    CleanupErrorCodes
    
    Handled = False

    Resume Next

End Sub

Private Sub ilvIDMListView_ItemSelectChange(ByVal Item As Object, ByVal Key As Long, ByVal Selected As Boolean, ByVal ObjType As IDMListView.idmObjectType)
    
On Error Resume Next
    
Dim losSelectedItems As ObjectSet

    If Selected Then
    
        Set losSelectedItems = ilvIDMListView.SelectedItems
        If losSelectedItems.Count = 1 Then
                
            'Set global variable to the latest item selected
            Set goCurSelListItem = Item
            giCurSelListObjType = ObjType
                
            Select Case ObjType
                Case idmObjTypeDocument
                'Item clicked is a document
                ' report ID of document being displayed in status bar
                gsMsg = LoadResString(GI_DOCUMENT) & ": " & CStr(goCurSelListItem.ID)
                UpdateStatusBar sbrStatusBar, gsMsg
                    
                Case idmObjTypeFolder
                'Item clicked is a folder
                gsMsg = LoadResString(GI_FOLDER) & ": " & CStr(Item.Name)
                UpdateStatusBar sbrStatusBar, gsMsg
            End Select
        Else
            Set goCurSelListItem = Nothing
            giCurSelListObjType = 0
                
            UpdateStatusBar sbrStatusBar, "Multiple items selected"
            UpdateStatusBar sbrStatusBar, "", 2
        End If
    End If
    
End Sub

Private Sub ilvIDMListView_ViewChanged(ByVal View As idmView)

    'Uncheck the current type
    mnuListViewMode(ilvIDMListView.View).Checked = False
    'Check the new type
    mnuListViewMode(View).Checked = True
    'Set the toolbar to the same new type
    tbrToolBar.Buttons(View + LISTVIEW_BUTTON).Value = tbrPressed

End Sub
' Just as an example, capture the event that says a folder has been copied
' to the clipboard
Private Sub itvIDMTreeView_BeforeInvokeCommand(ByVal Command As String)
If Command = "copy" Then
    gbFolderOnClipboard = True
End If
End Sub
' Pasting a folder from the clipboard creates some repaint problems in the
' Treeview; so as an example of trapping events, let's give him some help
' on the paste event...
Private Sub itvIDMTreeView_InvokeCommand(ByVal Command As String)
If Command = "paste" Then
    Dim lsMyName As String
    Dim loThisFolder As IDMObjects.Folder
    If gbFolderOnClipboard Then
        Set loThisFolder = itvIDMTreeView.SelectedItem
        ' Work around the repaint problems with a forced refresh
        Call ForceTreeRefresh(itvIDMTreeView, loThisFolder.Label)
        itvIDMTreeView.Expand
        ' The following is not strictly true, but it's the best we can do for now
        gbFolderOnClipboard = False
    End If
ElseIf Command = "logoff" Then
    ilvIDMListView.ClearItems
End If
End Sub


Private Sub mnuHelpAbout_Click()
    'To Do
    MsgBox "About Box Code goes here!"
End Sub

Private Sub mnuViewOptions_Click()
        
    frmOptions.Show vbModal

End Sub

Private Sub mnuViewStatusBar_Click()
    If mnuViewStatusBar.Checked Then
        sbrStatusBar.Visible = False
        mnuViewStatusBar.Checked = False
    Else
        sbrStatusBar.Visible = True
        mnuViewStatusBar.Checked = True
    End If
    SizeControls 0, imgSplitterH.Top
    SizeControls imgSplitter.Left, 0
End Sub

Private Sub mnuViewToolbar_Click()
    If mnuViewToolbar.Checked Then
        tbrToolBar.Visible = False
        mnuViewToolbar.Checked = False
    Else
        tbrToolBar.Visible = True
        mnuViewToolbar.Checked = True
    End If
    SizeControls 0, imgSplitterH.Top
    SizeControls imgSplitter.Left, 0
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.Width < 3000 Then Me.Width = 3000
    If Me.Height < 3000 Then Me.Height = 3000
    SizeControls 0, imgSplitterH.Top
    SizeControls imgSplitter.Left, 0
End Sub

Private Sub imgSplitter_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    With imgSplitter
        picSplitter.Move .Left, .Top, .Width \ 2, .Height - 20
    End With
    picSplitter.Visible = True
    mbMoving = True
End Sub
Private Sub imgSplitterH_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Added this subroutine to manage horizontal splitter
    With imgSplitterH
        picSplitterH.Move .Left, .Top, .Width - 20, .Height \ 2
    End With
    picSplitterH.Visible = True
    mbMoving = True
End Sub

Private Sub imgSplitter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim sglPos As Single

    If mbMoving Then
        sglPos = X + imgSplitter.Left
        If sglPos < sglSplitLimit Then
            picSplitter.Left = sglSplitLimit
        ElseIf sglPos > Me.Width - sglSplitLimit Then
            picSplitter.Left = Me.Width - sglSplitLimit
        Else
            picSplitter.Left = sglPos
        End If
    End If
End Sub
Private Sub imgSplitterH_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Added this subroutine to manage horizontal splitter
    Dim sglPos As Single

    If mbMoving Then
        sglPos = Y + imgSplitterH.Top
        If sglPos < (itvIDMTreeView.Top + sglSplitLimit) Then
            picSplitterH.Top = itvIDMTreeView.Top + sglSplitLimit
        ElseIf sglPos > (Me.ScaleHeight - 1000) Then
            picSplitterH.Top = Me.ScaleHeight - 1000
        Else
            picSplitterH.Top = sglPos
        End If
    End If
End Sub

Private Sub imgSplitter_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SizeControls picSplitter.Left, 0
    picSplitter.Visible = False
    mbMoving = False
End Sub

Private Sub imgSplitterH_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Added this subroutine to manage horizontal splitter
    SizeControls 0, picSplitterH.Top
    picSplitterH.Visible = False
    mbMoving = False
End Sub
Sub SizeControls(X As Single, Y As Single)
'Added parameter Y and modified this routine to manage horizontal splitter
    On Error Resume Next
    
    If X <> 0 Then
        'set the width
        If X < 1500 Then X = 1500
        If X > (Me.Width - 1500) Then X = Me.Width - 1500
        itvIDMTreeView.Width = X
        imgSplitter.Left = X
        ilvIDMListView.Left = X + 40
        ilvIDMListView.Width = Me.ScaleWidth - (itvIDMTreeView.Width + 40)
        lblTitle(0).Width = itvIDMTreeView.Width
        lblTitle(1).Left = ilvIDMListView.Left + 20
        lblTitle(1).Width = ilvIDMListView.Width - 40
    End If
    
    picPicture.Left = 20
    picPicture.Width = Me.ScaleWidth - 40
    
    'set the top
    If tbrToolBar.Visible Then
        itvIDMTreeView.Top = tbrToolBar.Height + picTitles.Height
    Else
        itvIDMTreeView.Top = picTitles.Height
    End If
    ilvIDMListView.Top = itvIDMTreeView.Top
        
    'Added code below
    'Set the height
    If Y <> 0 Then
        If Y < 1500 Then Y = 1500
        If Y > (Me.Height - 1500) Then Y = Me.Height - 1500
    
        If sbrStatusBar.Visible Then
            itvIDMTreeView.Height = Me.ScaleHeight - itvIDMTreeView.Top + (Y - (Me.Height - itvIDMTreeView.Top)) - 30
        Else
            itvIDMTreeView.Height = Me.ScaleHeight - itvIDMTreeView.Top + (Y - (Me.Height - itvIDMTreeView.Top))
        End If
    End If
    ilvIDMListView.Height = itvIDMTreeView.Height
    
    'position other controls
    imgSplitter.Top = picTitles.Top + 25
    imgSplitter.Height = itvIDMTreeView.Height + picTitles.Height - 25
    
    imgSplitterH.Top = itvIDMTreeView.Top + itvIDMTreeView.Height
    imgSplitterH.Left = 20
    imgSplitterH.Width = Me.Width - 40
    imgSplitterH.Visible = True
    
    picSplitterH.Top = itvIDMTreeView.Top + itvIDMTreeView.Height
    picSplitterH.Left = 20
    picSplitterH.Width = Me.Width - 40
    picSplitterH.Visible = False
    
    'Modify this for whatever the bottom controls are
    picPicture.Top = imgSplitterH.Top + 40
    If sbrStatusBar.Visible Then
        picPicture.Height = Me.ScaleHeight - picPicture.Top - sbrStatusBar.Height
    Else
        picPicture.Height = Me.ScaleHeight - picPicture.Top
    End If
    
    vwrViewer.Top = picPicture.Top
    vwrViewer.Left = picPicture.Left
    vwrViewer.Height = picPicture.Height
    vwrViewer.Width = picPicture.Width
    
End Sub

Private Sub tbrToolBar_ButtonClick(ByVal Button As ComctlLib.Button)

    Select Case Button.Key

        Case "Back"
            'To Do
            MsgBox "Back Code goes here!"
        Case "Forward"
            'To Do
            MsgBox "Forward Code goes here!"
        Case "Cut"
            mnuEditCut_Click
        Case "Copy"
            mnuEditCopy_Click
        Case "Paste"
            mnuEditPaste_Click
        Case "Delete"
            mnuFileDelete_Click
        Case "Properties"
            mnuFileProperties_Click
        Case "ViewLarge"
            mnuListViewMode_Click lvwIcon
        Case "ViewSmall"
            mnuListViewMode_Click lvwSmallIcon
        Case "ViewList"
            mnuListViewMode_Click lvwList
        Case "ViewDetails"
            mnuListViewMode_Click lvwReport
    End Select
End Sub

Private Sub mnuHelpContents_Click()

    Dim nRet As Integer

    'if there is no helpfile for this project display a message to the user
    'you can set the HelpFile for your application in the
    'Project Properties dialog
    If Len(App.HelpFile) = 0 Then
        MsgBox "Unable to display Help Contents. There is no Help associated with this project.", vbInformation, Me.Caption
    Else
        On Error Resume Next
        nRet = OSWinHelp(Me.hwnd, App.HelpFile, 3, 0)
        If Err Then
            MsgBox Err.Description
        End If
    End If
End Sub

Private Sub mnuHelpSearch_Click()

    Dim nRet As Integer
    
    'if there is no helpfile for this project display a message to the user
    'you can set the HelpFile for your application in the
    'Project Properties dialog
    If Len(App.HelpFile) = 0 Then
        MsgBox "Unable to display Help Contents. There is no Help associated with this project.", vbInformation, Me.Caption
    Else
        On Error Resume Next
        nRet = OSWinHelp(Me.hwnd, App.HelpFile, 261, 0)
        If Err Then
            MsgBox Err.Description
        End If
    End If
End Sub

Private Sub mnuVAIByDate_Click()
    'To Do
    'Set ilvIDMListView SortOrder as necessary
End Sub

Private Sub mnuVAIByName_Click()
    'To Do
    'Set ilvIDMListView SortOrder as necessary
End Sub

Private Sub mnuVAIBySize_Click()
    'To Do
    'Set ilvIDMListView SortOrder as necessary
End Sub

Private Sub mnuVAIByType_Click()
    'To Do
    'Set ilvIDMListView SortOrder as necessary
End Sub

Private Sub mnuListViewMode_Click(Index As Integer)
    
    'uncheck the current type
    mnuListViewMode(ilvIDMListView.View).Checked = False
    'set the listview mode
    ilvIDMListView.View = Index
    'check the new type
    mnuListViewMode(Index).Checked = True
    'set the toolbar to the same new type
    tbrToolBar.Buttons(Index + LISTVIEW_BUTTON).Value = tbrPressed
    
End Sub

Private Sub mnuViewLineUpIcons_Click()
    'To Do
    ilvIDMListView.Arrange = idmArrangeTop
End Sub

Private Sub mnuViewRefresh_Click()
    'To Do
    MsgBox "Refresh Code goes here!"
End Sub

Private Sub mnuEditCopy_Click()
    
    'This copies a folder not a document
    frmBrowse.Show vbModal
    itvIDMTreeView.Refresh

End Sub

Private Sub mnuEditCut_Click()
    'To Do
    MsgBox "Cut Code goes here!"
End Sub

Private Sub mnuEditSelectAll_Click()
    'To Do
    MsgBox "Select All Code goes here!"
End Sub

Private Sub mnuEditInvertSelection_Click()
    'To Do
    MsgBox "Invert Selection Code goes here!"
End Sub

Private Sub mnuEditPaste_Click()
    'To Do
    MsgBox "Paste Code goes here!"
End Sub

Private Sub mnuEditPasteSpecial_Click()
    'To Do
    MsgBox "Paste Special Code goes here!"
End Sub

Private Sub mnuEditUndo_Click()
    'To Do
    MsgBox "Undo Code goes here!"
End Sub

Private Sub mnuFileOpen_Click()
    'To Do
    MsgBox "Open Code goes here!"
End Sub

Private Sub mnuFileFind_Click()
    'To Do
    MsgBox "Find Code goes here!"
End Sub

Private Sub mnuFileSendTo_Click()
    'To Do
    MsgBox "Send To Code goes here!"
End Sub

Private Sub mnuFileNew_Click()
'This subroutine adds a new sub-folder to the currently
'selected folder in the TreeView.

Dim lsFolderName As String
    
On Error GoTo ErrorHandler

    lsFolderName = InputBox(LoadResString(GI_FOLDER_NAME), LoadResString(GI_FOLDER))
    If lsFolderName <> "" Then
        Dim loFolder As IDMObjects.Folder
        If giCurSelTreeObjType = idmObjTypeFolder Then
            Set loFolder = goCurSelTreeItem.CreateSubFolder(lsFolderName)
            loFolder.SaveNew
        ElseIf giCurSelTreeObjType = idmObjTypeLibrary Then
            Set loFolder = goCurSelTreeItem.CreateObject(idmObjTypeFolder, "notused")
            loFolder.Name = lsFolderName
            loFolder.SaveNew
        End If
            
        goCurSelTreeItem.Refresh idmFolderRefreshAll
        gbSuccess = FillIDMListView(ilvIDMListView, goCurSelTreeItem, giCurSelTreeObjType)
        If gbSuccess = False Then
            MsgBox LoadResString(GI_ERR_UNABLE_TO_FILL_LISTVIEW), vbExclamation, LoadResString(GI_ERR_ERROR)
        End If
        itvIDMTreeView.RefreshItem
      
        'Update UI
        'Clear status bar
        UpdateStatusBar sbrStatusBar, ""
        'Update status bar panel 2
        gsMsg = " " & giFolderCount & " " & LoadResString(GI_FOLDERS) & ", " & giDocumentCount & " " & LoadResString(GI_DOCUMENTS) & " "
        UpdateStatusBar sbrStatusBar, gsMsg, 2
        'Update label
        lblTitle(1) = " " & LoadResString(GI_CONTENTS_OF) & " '" & goCurSelTreeItem.Label & "'"
        'Update form caption
        If giCurSelTreeObjType = idmObjTypeFolder Then
            Me.Caption = lsOriginalFormCaption & " - " & goCurSelTreeItem.PathName
        ElseIf giCurSelTreeObjType = idmObjTypeLibrary Then
            Me.Caption = lsOriginalFormCaption & " - " & goCurSelTreeItem.Label
        End If
        
    End If
    
    Exit Sub
    
ErrorHandler:
    
    'Display Error Message - pass the name of this subroutine/function
    DisplayErrorMessage ("frmMain.mnuFileNew_Click")
    
    'Cleanup Error values
    CleanupErrorCodes

    Resume Next
            
End Sub

Private Sub mnuFileDelete_Click()
    'To Do
    MsgBox "Delete Code goes here!"
End Sub

Private Sub mnuFileRename_Click()
    'To Do
    MsgBox "Rename Code goes here!"
End Sub

Private Sub mnuFileProperties_Click()
    
    'To Do
    'Need to add code to handle if user clicks OK or Cancel
    If Screen.ActiveControl.Name = "itvIDMTreeView" Then
        goCurSelTreeItem.ShowPropertiesDialog
    ElseIf Screen.ActiveControl.Name = "ilvIDMListView" Then
        goCurSelListItem.ShowPropertiesDialog
    End If
    
End Sub

Private Sub mnuFileMRU_Click(Index As Integer)
    'To Do
    MsgBox "MRU Code goes here!"
End Sub

Private Sub mnuFileClose_Click()
    'unload the form
    Unload Me
End Sub

Private Sub itvIDMTreeView_Click()
        
    'Cleanup UI
    'Clear status bar panel 2
    UpdateStatusBar sbrStatusBar, "", 2
    'Set form caption back
    Me.Caption = lsOriginalFormCaption

End Sub

Private Sub itvIDMTreeView_DoItemDelete(ByVal Item As Object, Handled As Boolean)
'This subroutine processes the deletion of the selected item in the Tree View.
'The UI shows the user deleting the object, but the foundation object
'must be updated as well.
    
Dim liUserResponse As Integer
Dim loCurSelTreeItemParent As Object
Dim loMyFolder As IDMObjects.Folder
Dim liObjType As Integer
Dim loParent As Object
    
On Error GoTo ErrorHandler
    ' Don't assume target item is necessarily selected...
    If TypeOf Item Is IDMObjects.Folder Then
        Set loMyFolder = Item
        liUserResponse = MsgBox("'" & loMyFolder.Label & "' " & _
            LoadResString(GI_CONFIRM_FDELETE), vbExclamation + vbYesNo, LoadResString(GI_CONFIRM))
        Select Case liUserResponse
            Case vbNo
                'User cancelled
                Handled = False
            Case vbYes
                MouseWait
                ' We want to turn this into an 'unfile documents, delete folders' operation,
                ' so call the return to unfile documents in this folder and all of its
                ' children; then call delete
                Call CleanOutFolder(loMyFolder)
                ' Can't assume that deleted node is actually selected yet
                ' Set loCurSelTreeItemParent = goCurSelTreeItem.Parent
                Item.Delete
                Handled = True
                Set goCurSelTreeItem = loMyFolder.Parent
                goCurSelTreeItem.Refresh idmFolderRefreshAll
                MouseNormal
        End Select
    Else
        MsgBox LoadResString(GI_ERR_UNABLE_TO_DELETE), vbExclamation, LoadResString(GI_ERR_ERROR)
    End If
    
    Exit Sub
    
ErrorHandler:

    'Display Error message
    DisplayErrorMessage ("frmMain.itvIDMTreeView_DoItemDelete")

    'Cleanup Error values
    CleanupErrorCodes
    
    Handled = False

    Resume Next

End Sub

Private Sub itvIDMTreeView_DoLabelEdit(ByVal Item As Object, ByVal Label As String, Handled As Boolean)
'This subroutine processes the name change of a tree view item.
'The UI shows the user changing the label, but the foundation object
'must be updated as well.

Dim loFolder As IDMObjects.Folder
    
On Error GoTo ErrorHandler

    Set loFolder = Item
    'Check to see if user has security to modify
    If loFolder.GetState(idmFolderCanModify) = True Then
        If loFolder.Name <> Label And Label <> "" Then
            itvIDMTreeView.RenameItem Label
            loFolder.Name = Label
            loFolder.Save
            Handled = True
        End If
    Else
        MsgBox LoadResString(GI_ERR_UNABLE_TO_RENAME), vbExclamation, LoadResString(GI_ERR_ERROR)
    End If
    
    Exit Sub
    
ErrorHandler:

    'Display Error message
    DisplayErrorMessage ("frmMain.itvIDMTreeView_DoLabelEdit")

    'Cleanup Error values
    CleanupErrorCodes
    
    Handled = False

    Resume Next

End Sub

Private Sub itvIDMTreeView_ItemSelectChange(ByVal Item As Object, ByVal ObjType As IDMTreeView.idmObjectType)
    
    'Set the global variable to know which item is currently selected.
    Set goCurSelTreeItem = Item
    giCurSelTreeObjType = ObjType
    
    itvIDMTreeView.RefreshItem
    
    gbSuccess = FillIDMListView(ilvIDMListView, Item, ObjType)
    
    If gbSuccess = False Then
       MsgBox LoadResString(GI_ERR_UNABLE_TO_FILL_LISTVIEW), vbExclamation, LoadResString(GI_ERR_ERROR)
    End If
    
    'Update UI
    'Clear status bar
    UpdateStatusBar sbrStatusBar, ""
    'Update status bar panel 2
    gsMsg = " " & giFolderCount & " " & LoadResString(GI_FOLDERS) & ", " & giDocumentCount & " " & LoadResString(GI_DOCUMENTS) & " "
    UpdateStatusBar sbrStatusBar, gsMsg, 2
    'Update label
    lblTitle(1) = " " & LoadResString(GI_CONTENTS_OF) & " '" & Item.Label & "'"
    'Update form caption
    If Item.ObjectType = idmObjTypeFolder Then
        Me.Caption = lsOriginalFormCaption & " - " & Item.PathName
    ElseIf Item.ObjectType = idmObjTypeLibrary Then
        Me.Caption = lsOriginalFormCaption & " - " & Item.Label
    End If
    'Hide the viewer until the user selects a document from the list view
    vwrViewer.Visible = False

End Sub
