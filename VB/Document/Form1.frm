VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{54E6FE61-B93A-11CF-8185-444553540000}#3.0#0"; "fntree.ocx"
Object = "{2A601BA0-B880-11CF-8185-444553540000}#3.0#0"; "fnlist.ocx"
Begin VB.Form MainForm 
   Caption         =   "IDM Sample Application - Document and Folder Automation"
   ClientHeight    =   7530
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9990
   LinkTopic       =   "Form1"
   ScaleHeight     =   7530
   ScaleWidth      =   9990
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Query 
      Caption         =   "Query for Documents..."
      Height          =   495
      Left            =   840
      TabIndex        =   16
      Top             =   4680
      Width           =   1935
   End
   Begin IDMListView.IDMListView ListView1 
      Height          =   3735
      Left            =   3720
      TabIndex        =   15
      Top             =   120
      Width           =   6015
      _Version        =   196608
      _ExtentX        =   10610
      _ExtentY        =   6588
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
      _ColumnHeaders  =   "Form1.frx":0000
   End
   Begin IDMTreeView.IDMTreeView TreeView1 
      Height          =   3735
      Left            =   240
      TabIndex        =   14
      Top             =   120
      Width           =   3375
      _Version        =   196608
      _ExtentX        =   5953
      _ExtentY        =   6588
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
   Begin VB.CommandButton ShowFolders 
      Caption         =   "Show Folders Filed In..."
      Enabled         =   0   'False
      Height          =   495
      Left            =   5760
      TabIndex        =   13
      Top             =   5880
      Width           =   1935
   End
   Begin VB.CommandButton ShowAnnotations 
      Caption         =   "Show Annotations..."
      Enabled         =   0   'False
      Height          =   495
      Left            =   3720
      TabIndex        =   12
      Top             =   5880
      Width           =   1935
   End
   Begin VB.CommandButton ShowClass 
      Caption         =   "Show Class..."
      Enabled         =   0   'False
      Height          =   495
      Left            =   7800
      TabIndex        =   11
      Top             =   5880
      Width           =   1935
   End
   Begin VB.CommandButton Done 
      Caption         =   "Done"
      Height          =   375
      Left            =   8520
      TabIndex        =   10
      Top             =   6960
      Width           =   1215
   End
   Begin VB.CommandButton Send 
      Caption         =   "Send..."
      Enabled         =   0   'False
      Height          =   495
      Left            =   7800
      TabIndex        =   9
      Top             =   5280
      Width           =   1935
   End
   Begin VB.CommandButton ShowState 
      Caption         =   "Show State..."
      Enabled         =   0   'False
      Height          =   495
      Left            =   5760
      TabIndex        =   8
      Top             =   5280
      Width           =   1935
   End
   Begin VB.CommandButton OpenNative 
      Caption         =   "Open..."
      Enabled         =   0   'False
      Height          =   495
      Left            =   3720
      TabIndex        =   7
      Top             =   5280
      Width           =   1935
   End
   Begin VB.CommandButton Permissions 
      Caption         =   "Show Permissions..."
      Enabled         =   0   'False
      Height          =   495
      Left            =   7800
      TabIndex        =   6
      Top             =   4680
      Width           =   1935
   End
   Begin VB.CommandButton ShowPropDialog 
      Caption         =   "Show Properties Dialog..."
      Enabled         =   0   'False
      Height          =   495
      Left            =   5760
      TabIndex        =   5
      Top             =   4680
      Width           =   1935
   End
   Begin VB.CommandButton ShowProperties 
      Caption         =   "Show Properties..."
      Enabled         =   0   'False
      Height          =   495
      Left            =   7800
      TabIndex        =   4
      Top             =   4080
      Width           =   1935
   End
   Begin VB.CommandButton Versions 
      Caption         =   "Show Versions..."
      Enabled         =   0   'False
      Height          =   495
      Left            =   5760
      TabIndex        =   3
      Top             =   4080
      Width           =   1935
   End
   Begin VB.CommandButton Checkin 
      Caption         =   "Checkin Document..."
      Enabled         =   0   'False
      Height          =   495
      Left            =   3720
      TabIndex        =   2
      Top             =   4680
      Width           =   1935
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   360
      Top             =   6120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Checkout 
      Caption         =   "Checkout Document..."
      Enabled         =   0   'False
      Height          =   495
      Left            =   3720
      TabIndex        =   1
      Top             =   4080
      Width           =   1935
   End
   Begin VB.CommandButton AddDocument 
      Caption         =   "Add New Document..."
      Enabled         =   0   'False
      Height          =   495
      Left            =   840
      TabIndex        =   0
      Top             =   4080
      Width           =   1935
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public oErrManager As idmError.ErrorManager
Public oNeighborhood As New IDMObjects.Neighborhood
Public oLibraries As IDMObjects.ObjectSet
Public oCurrentLibrary As IDMObjects.Library
Public oDocument As IDMObjects.Document
Public oFolder As IDMObjects.Folder
Public oSelected As Object

Public Sub ShowError()
Dim oErrCollect As idmError.Errors
Dim oError As idmError.Error
Dim iCnt As Integer
Set oErrCollect = oErrManager.Errors
If oErrCollect.Count > 1 Then
    iCnt = 1
    For Each oError In oErrCollect
        MsgBox "Error " & iCnt & ": " & oError.Description & " : " & Hex(oError.Number)
        iCnt = iCnt + 1
    Next
Else
    If oErrCollect.Count = 1 Then
        oErrManager.ShowErrorDialog
    Else
        If Err.Number <> 0 Then
            MsgBox Err.Description & " : " & Err.Number
        End If
    End If
End If
End Sub

Private Sub Done_Click()
    End
End Sub


Private Sub Form_Load()
    TreeView1.AddRootItem oNeighborhood, True
    Set oErrManager = CreateObject("IDMError.ErrorManager")
End Sub

Function Logon() As Boolean
    On Error GoTo ErrorHandler
    Logon = False
    If ListView1.SelectedItem Is Nothing Then
        MsgBox "Select a Library"
    Else
        Set oLibrary = ListView1.SelectedItem
        If Not oLibrary.GetState(idmLibraryLoggedOn) Then
            If oLibrary.Logon(, , , idmLogonOptWithUI) Then
                Logon = True
            Else
                MsgBox "Logon failed, you must logon"
            End If
        End If
    End If
    Exit Function
ErrorHandler:
    ShowError
End Function

Private Sub AddDocument_Click()
    AddDocumentForm.Show 1, MainForm
End Sub

Private Sub Checkin_Click()
    Dim oVersion As IDMObjects.Version
    Dim frmLocs As New frmChkout
    On Error GoTo Errors

    If Not oDocument Is Nothing Then
        If oDocument.GetState(idmDocCanCheckin) Then
            Set oVersion = oDocument.Version
            frmLocs.txtFullpath = oVersion.CheckoutPath
            frmLocs.txtFullpath.Enabled = False
            frmLocs.txtDirpart = ""
            frmLocs.txtFilepart = ""
            frmLocs.Show vbModal
            
            Call oVersion.Checkin(frmLocs.txtDirpart, _
                frmLocs.txtFilepart)
        Unload frmLocs
        End If
    End If
Exit Sub
Errors:
Call ShowError
    
End Sub

Private Sub Checkout_Click()
    Dim oVersion As IDMObjects.Version
    Dim frmLocs As New frmChkout
    Dim sLoc As String
 
    On Error GoTo Errors
    If Not oDocument Is Nothing Then
        If oDocument.GetState(idmDocCanCheckout) Then
            Set oVersion = oDocument.Version
            'Can't specify full path part; just dir and filename
            frmLocs.txtFullpath = ""
            frmLocs.txtFullpath.Enabled = False
            frmLocs.txtDirpart = ""
            frmLocs.txtFilepart = ""
            frmLocs.Show vbModal
            sLoc = frmLocs.txtFullpath
            Call oVersion.Checkout _
              (sLoc, frmLocs.txtDirpart, frmLocs.txtFilepart)
            MsgBox "Checked out to: " & sLoc
        Else ' document selected doesn't support checkout
            MsgBox "Document cannot be checked out"
        Unload frmLocs
        End If
    End If
Exit Sub
Errors:
Call ShowError
End Sub

Private Sub Form_Terminate()
    Set oNeighborhood = Nothing
    Set oLibraries = Nothing
    Set oCurrentLibrary = Nothing
    Set oErrManager = Nothing
End Sub

Private Sub ListView1_DblClick()
    If Not ListView1.SelectedItem Is Nothing Then
        If ListView1.SelectedItem.ObjectType = idmObjTypeFolder Then
            Dim oFolder As IDMObjects.Folder
            Dim oDocs As IDMObjects.ObjectSet
            Set oFolder = ListView1.SelectedItem
            Set oDocs = oFolder.GetContents(idmFolderContentDocument)
            ListView1.ClearItems
            ListView1.AddItems oFolder.SubFolders, -1
            ListView1.AddItems oDocs, -1
            MainForm.Checkin.Enabled = False
            MainForm.Checkout.Enabled = False
            MainForm.Versions.Enabled = False
            MainForm.ShowPropDialog.Enabled = False
            MainForm.ShowProperties.Enabled = False
            MainForm.Permissions.Enabled = False
            MainForm.OpenNative.Enabled = False
            MainForm.ShowState.Enabled = False
            MainForm.Send.Enabled = False
            MainForm.ShowClass.Enabled = False
            MainForm.ShowAnnotations.Enabled = False
            MainForm.ShowFolders.Enabled = False
        End If
    End If
End Sub


Private Function FormatAccessType(ByVal SysType As idmSysTypeOptions, ByVal AccessType As Long) As String
    Select Case AccessType
        Case idmMzAccessAdmin
            FormatAccessType = "Admin"
        Case idmMzAccessAuthor
            If SysType = idmSysTypeDS Then
                FormatAccessType = "Author"
            Else
                FormatAccessType = "Write"
            End If
        Case idmMzAccessNone
            FormatAccessType = "None"
        Case idmMzAccessOwner
            If SysType = idmSysTypeDS Then
                FormatAccessType = "Owner"
            Else
                FormatAccessType = "Append/Execute"
            End If
        Case idmMzAccessViewer
            If SysType = idmSysTypeDS Then
                FormatAccessType = "Viewer"
            Else
                FormatAccessType = "Read"
            End If
    End Select
End Function

Private Sub ListView1_ItemSelectChange(ByVal Item As Object, ByVal Key As Long, ByVal Selected As Boolean, ByVal ObjType As IDMListView.idmObjectType)
    If ObjType = idmObjTypeDocument Then
        Set oDocument = Item
        Set oSelected = Item
        Set oFolder = Nothing
        supports_versions = oCurrentLibrary.Supports(idmSupportsDocVersions)
        MainForm.Checkin.Enabled = supports_versions
        MainForm.Checkout.Enabled = supports_versions
        MainForm.Versions.Enabled = supports_versions
        MainForm.ShowPropDialog.Enabled = True
        MainForm.ShowProperties.Enabled = True
        MainForm.Permissions.Enabled = True
        MainForm.OpenNative.Enabled = True
        MainForm.ShowState.Enabled = True
        MainForm.Send.Enabled = True
        MainForm.ShowClass.Enabled = True
        MainForm.ShowAnnotations.Enabled = oCurrentLibrary.Supports(idmSupportsAnnotations)
        MainForm.ShowFolders.Enabled = True
    ElseIf ObjType = idmObjTypeFolder Then
        Set oFolder = Item
        Set oSelected = Item
        Set oDocument = Nothing
        ' Dim oPd As PropertyDescription
        ' Set oPd = oCurrentLibrary.GetObject(idmObjTypePropDesc, "idmName", idmObjTypeFolder)
        ' MsgBox (oPd.GetState(idmPropSearchable))
        MainForm.Checkin.Enabled = False
        MainForm.Checkout.Enabled = False
        MainForm.Versions.Enabled = False
        MainForm.ShowPropDialog.Enabled = True
        MainForm.ShowProperties.Enabled = True
        MainForm.Permissions.Enabled = True
        MainForm.OpenNative.Enabled = False
        MainForm.ShowState.Enabled = True
        MainForm.Send.Enabled = False
        MainForm.ShowClass.Enabled = True
        MainForm.ShowAnnotations.Enabled = False
        MainForm.ShowFolders.Enabled = False
    End If

End Sub

Private Sub OpenNative_Click()
    If Not oDocument Is Nothing Then
        Screen.MousePointer = vbHourglass
        If oDocument.TypeName = "FileNET IDM Document" Then
            oDocument.Launch
        Else
            oDocument.Launch (idmDocLaunchNativeApplication)
        End If
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub Permissions_Click()
    Dim oPermission As IDMObjects.Permission
    On Error GoTo ErrorHandler
    PermissionsForm.MsList2.ListItems.Clear
    If Not oSelected Is Nothing Then
        Dim itemX As ListItem
        For Each oPermission In oSelected.Permissions
            Set itemX = PermissionsForm.MsList2.ListItems.Add(, , oPermission.GranteeName)
            itemX.SubItems(1) = oPermission.Label
            itemX.SubItems(2) = oPermission.GranteeType
        Next
    End If
    PermissionsForm.Label1 = oSelected.Name
    PermissionsForm.Show 1, MainForm
    Exit Sub
ErrorHandler:
    ShowError
End Sub

Private Sub Query_Click()
    QueryForm.Show 1, MainForm
End Sub


Private Sub Send_Click()
    oDocument.Send , , , , , , , , idmSendWithUI
End Sub

Private Sub ShowAnnotations_Click()
    Dim oAnnos As IDMObjects.ObjectSet
    Screen.MousePointer = vbHourglass
    Set oAnnos = oDocument.Annotations
    If Not oAnnos Is Nothing Then
        AnnoForm.ListView1.ClearItems
        AnnoForm.ListView1.AddItems oAnnos, -1
    End If
    Screen.MousePointer = vbDefault
    AnnoForm.Show 1, MainForm
End Sub

Private Sub ShowClass_Click()
    Dim oClass As IDMObjects.ClassDescription
    Dim oPropDesc As IDMObjects.PropertyDescription
    Set oClass = oSelected.ClassDescription
    ClassForm.Label2 = oClass.Label
    ClassForm.MsList1.ListItems.Clear
    Dim itemX As ListItem
    For Each oPropDesc In oClass.PropertyDescriptions
        Set itemX = ClassForm.MsList1.ListItems.Add(, , oPropDesc.Name)
        itemX.SubItems(1) = FormatDataType(oPropDesc.TypeID)
    Next
    ClassForm.Show 1, MainForm
End Sub

Private Sub ShowFolders_Click()
    Dim oFolders As IDMObjects.ObjectSet
    Dim oTempFolder As IDMObjects.Folder
    Set oFolders = oDocument.FoldersFiledIn
    FoldersFiledForm.List1.Clear
    If Not oFolders Is Nothing Then
        For Each oTempFolder In oFolders
            FoldersFiledForm.List1.AddItem oTempFolder.PathName
        Next
        FoldersFiledForm.Label1 = oDocument.Label
    Else
        FoldersFiledForm.Label1 = oDocument.Label & " not filed in folders"
    End If
    FoldersFiledForm.Show 1, MainForm
End Sub

Private Sub ShowPropDialog_Click()
    If oSelected.ShowPropertiesDialog = idmDialogExitOK Then
        oSelected.Save
    End If
End Sub

Private Sub ShowProperties_Click()
    Dim oProperty As IDMObjects.Property
    Dim itemX As ListItem
    PropertiesForm.MsList.ListItems.Clear
    For Each oProperty In oSelected.Properties
        'Set the first column to the property name
        Set itemX = PropertiesForm.MsList.ListItems.Add(, , oProperty.PropertyDescription.Name)
        itemX.SubItems(1) = oProperty.FormatValue
        itemX.SubItems(2) = FormatDataType(oProperty.PropertyDescription.TypeID)
        itemX.SubItems(3) = oProperty.PropertyDescription.GetState(idmPropSearchable)
        itemX.SubItems(4) = oProperty.PropertyDescription.GetState(idmPropMultiValue)
    Next
    PropertiesForm.Label2 = oSelected.Name
    PropertiesForm.Show 1, MainForm
End Sub

Private Sub ShowState_Click()
    ' show state of the current document
    If oSelected.ObjectType = idmObjTypeDocument Then
        Dim oDoc As IDMObjects.Document
        Set oDoc = MainForm.oDocument
        StateForm.Label1 = "Annotations modified:"
        StateForm.Text1 = oDoc.GetState(idmDocAnnosModified)
        StateForm.Label2 = "Document is annotated:"
        StateForm.Text2 = oDoc.GetState(idmDocAnnotated)
        StateForm.Label3 = "Document can be annotated:"
        StateForm.Text3 = oDoc.GetState(idmDocCanAnnotate)
        StateForm.Label4 = "Can cancel document checkout:"
        StateForm.Text4 = oDoc.GetState(idmDocCanCancelCheckout)
        StateForm.Label5 = "Document can be checked in:"
        StateForm.Text5 = oDoc.GetState(idmDocCanCheckin)
        StateForm.Label6 = "Document can be checked out:"
        StateForm.Text6 = oDoc.GetState(idmDocCanCheckout)
        StateForm.Label7 = "Document can be deleted:"
        StateForm.Text7 = oDoc.GetState(idmDocCanDelete)
        StateForm.Label8 = "Document is checked out:"
        StateForm.Text8 = oDoc.GetState(idmDocCheckedout)
        StateForm.Label9 = "Document is latest version:"
        StateForm.Text9 = oDoc.GetState(idmDocLatestVersion)
        StateForm.Label10 = "Document has been modified:"
        StateForm.Text10 = oDoc.GetState(idmDocModified)
    ' show state of the current folder
    ElseIf oSelected.ObjectType = idmObjTypeFolder Then
        Dim oFolder As IDMObjects.Folder
        Set oFolder = MainForm.oFolder
        StateForm.Label1 = "Folder can be deleted"
        StateForm.Text1 = oFolder.GetState(idmFolderCanDelete)
        StateForm.Label2 = "Folder can be filed in"
        StateForm.Text2 = oFolder.GetState(idmFolderCanFileIn)
        StateForm.Label3 = "Can modify folder properties:"
        StateForm.Text3 = oFolder.GetState(idmFolderCanModify)
        StateForm.Label4 = "Folder is a replica:"
        StateForm.Text4 = oFolder.GetState(idmFolderIsReplica)
        StateForm.Label5 = "Folder has been modified:"
        StateForm.Text5 = oFolder.GetState(idmFolderModified)
        StateForm.Label6 = ""
        StateForm.Text6 = ""
        StateForm.Label7 = ""
        StateForm.Text7 = ""
        StateForm.Label8 = ""
        StateForm.Text8 = ""
        StateForm.Label9 = ""
        StateForm.Text9 = ""
        StateForm.Label10 = ""
        StateForm.Text10 = ""
    End If
    
    ' Show the form
    StateForm.Show 1, MainForm
End Sub

Private Sub TreeView1_ItemSelectChange(ByVal Item As Object, ByVal ObjType As IDMTreeView.idmObjectType)
    Select Case ObjType
        Case idmObjTypeLibrary
            Dim oLibrary As IDMObjects.Library
            Set oLibrary = Item
            On Error Resume Next
            If Not oLibrary.GetState(idmLibraryLoggedOn) Then
                oLibrary.Logon , , , idmLogonOptWithUI
            End If
            ListView1.ClearItems
            ListView1.AddItems oLibrary.TopFolders, -1
            Set oCurrentLibrary = oLibrary
            MainForm.AddDocument.Enabled = True
        Case idmObjTypeFolder
            Dim oFolder As IDMObjects.Folder
            Set oFolder = Item
            ListView1.ClearItems
            ListView1.AddItems oFolder.SubFolders, -1
            ListView1.AddItems oFolder.GetContents(idmFolderContentDocument), -1
            Set oCurrentLibrary = oFolder.Library
            'ListView1.AddItems oFolder.GetContents(idmFolderContentStoredSearch), -1
            MainForm.AddDocument.Enabled = True
            Set oFolder = Item
    End Select
End Sub


Private Sub Versions_Click()
    On Error GoTo ErrorHandler
    ' Check to see if the library supports versions
    VersionsForm.lvVersions.ListItems.Clear
    Dim objVersion As IDMObjects.Document
    Dim lvItem As ListItem
    If Not oDocument.Version.Series Is Nothing Then
        'add each doc version to the listview
        VersionsForm.Label1.Caption = oDocument.Name
        For Each objVersion In oDocument.Version.Series
            Set lvItem = VersionsForm.lvVersions.ListItems.Add(, , objVersion.Label)
            lvItem.SubItems(1) = objVersion.Version.Number ' get ver #
        Next objVersion
        VersionsForm.Show vbModal, Me
    End If
    Exit Sub
ErrorHandler:
    ShowError
    Exit Sub
End Sub

Function FormatDataType(ByVal TypeID As idmTypeID)
    Select Case TypeID
        Case idmTypeArray
            FormatDataType = "Array"
        Case idmTypeBoolean
            FormatDataType = "Boolean"
        Case idmTypeByte
            FormatDataType = "Byte"
        Case idmTypeCharacter
            FormatDataType = "Character"
        Case idmTypeCurrency
            FormatDataType = "Currency"
        Case idmTypeDate
            FormatDataType = "Date"
        Case idmTypeDouble
            FormatDataType = "Double"
        Case idmTypeEmpty
            FormatDataType = "Empty"
        Case idmTypeError
            FormatDataType = "Error"
        Case idmTypeGuid
            FormatDataType = "GUID"
        Case idmTypeLong
            FormatDataType = "Long"
        Case idmTypeNull
            FormatDataType = "NULL"
        Case idmTypeObject
            FormatDataType = "Object"
        Case idmTypeShort
            FormatDataType = "Short"
        Case idmTypeSingle
            FormatDataType = "Single"
        Case idmTypeString
            FormatDataType = "String"
        Case idmTypeUnknown
            FormatDataType = "Unknown"
        Case idmTypeUnsignedLong
            FormatDataType = "Unsigned Long"
        Case idmTypeUnsignedShort
            FormatDataType = "Unsigned Short"
        Case idmTypeVariant
            FormatDataType = "Variant"
    End Select
End Function

