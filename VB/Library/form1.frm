VERSION 5.00
Object = "{2A601BA0-B880-11CF-8185-444553540000}#3.0#0"; "fnlist.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form Form1 
   Caption         =   "Library Sample Application"
   ClientHeight    =   8160
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   10560
   LinkTopic       =   "Form1"
   ScaleHeight     =   8160
   ScaleWidth      =   10560
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtClassInfo 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1560
      TabIndex        =   24
      Top             =   2160
      Width           =   3855
   End
   Begin VB.TextBox txtVersionNumber 
      Height          =   285
      Left            =   8880
      TabIndex        =   23
      Top             =   2880
      Width           =   1575
   End
   Begin VB.TextBox txtSelectedClass 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   285
      Left            =   4080
      TabIndex        =   17
      Top             =   2640
      Width           =   2895
   End
   Begin VB.Frame Frame1 
      Caption         =   "Object Type "
      Height          =   1335
      Left            =   7200
      TabIndex        =   11
      Top             =   480
      Width           =   3255
      Begin VB.OptionButton rbCustomObject 
         Caption         =   "CustomObject"
         Height          =   195
         Left            =   1560
         TabIndex        =   16
         Top             =   600
         Width           =   1455
      End
      Begin VB.OptionButton rbStoredSearch 
         Caption         =   "StoredSearch"
         Height          =   195
         Left            =   1560
         TabIndex        =   15
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton rbAnnotation 
         Caption         =   "Annotation"
         Height          =   315
         Left            =   360
         TabIndex        =   14
         Top             =   960
         Width           =   1455
      End
      Begin VB.OptionButton rbFolder 
         Caption         =   "Folder"
         Height          =   195
         Left            =   360
         TabIndex        =   13
         Top             =   600
         Width           =   1455
      End
      Begin VB.OptionButton rbDocument 
         Caption         =   "Document"
         Height          =   195
         Left            =   360
         TabIndex        =   12
         Top             =   240
         Value           =   -1  'True
         Width           =   1455
      End
   End
   Begin VB.TextBox txtDocClass 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   285
      Left            =   8280
      TabIndex        =   10
      Top             =   7800
      Width           =   2055
   End
   Begin VB.CommandButton btnShowProperties 
      Caption         =   "Show Properties"
      Enabled         =   0   'False
      Height          =   375
      Left            =   7320
      TabIndex        =   8
      Top             =   3600
      Width           =   1455
   End
   Begin VB.TextBox txtObjectID 
      Height          =   285
      Left            =   8040
      TabIndex        =   7
      Top             =   2520
      Width           =   2415
   End
   Begin ComctlLib.ListView ListView3 
      Height          =   3615
      Left            =   7320
      TabIndex        =   5
      Top             =   4080
      Width           =   3375
      _ExtentX        =   5948
      _ExtentY        =   6371
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Value"
         Object.Width           =   2540
      EndProperty
   End
   Begin ComctlLib.ListView ListView2 
      Height          =   4695
      Left            =   3120
      TabIndex        =   4
      Top             =   3000
      Width           =   3855
      _ExtentX        =   6795
      _ExtentY        =   8276
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
         Text            =   "Property Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Data Type"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Required"
         Object.Width           =   2540
      EndProperty
   End
   Begin ComctlLib.ListView ListView1 
      Height          =   4695
      Left            =   120
      TabIndex        =   3
      Top             =   3000
      Width           =   2895
      _ExtentX        =   5101
      _ExtentY        =   8276
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "ClassDescription Name"
         Object.Width           =   5080
      EndProperty
   End
   Begin VB.CommandButton btnShowClasses 
      Caption         =   "Show Classes"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   2520
      Width           =   1575
   End
   Begin IDMListView.IDMListView IDMListView1 
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   6735
      _Version        =   196608
      _ExtentX        =   11880
      _ExtentY        =   1931
      _StockProps     =   239
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   1
      View            =   1
      _ColumnHeaders  =   "form1.frx":0000
   End
   Begin VB.Label Label10 
      Caption         =   "Specific version #:"
      Height          =   255
      Left            =   7200
      TabIndex        =   22
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   7080
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Label Label8 
      Caption         =   "Object Type Selection"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7560
      TabIndex        =   21
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label Label7 
      Caption         =   "Library Selection"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1920
      TabIndex        =   20
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label Label6 
      Caption         =   "Single Object Selection"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7560
      TabIndex        =   19
      Top             =   2160
      Width           =   2775
   End
   Begin VB.Line Line4 
      X1              =   7080
      X2              =   7080
      Y1              =   1560
      Y2              =   0
   End
   Begin VB.Line Line2 
      X1              =   7080
      X2              =   7080
      Y1              =   1560
      Y2              =   8640
   End
   Begin VB.Line Line1 
      X1              =   7080
      X2              =   11880
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Label Label4 
      Caption         =   "Class:"
      Height          =   255
      Left            =   3360
      TabIndex        =   18
      Top             =   2640
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "Class:"
      Height          =   255
      Left            =   7440
      TabIndex        =   9
      Top             =   7800
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Object ID:"
      Height          =   255
      Left            =   7200
      TabIndex        =   6
      Top             =   2520
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Available libraries:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim oHood As New IDMObjects.Neighborhood
Dim oLib As IDMObjects.Library
Dim CurrObjType As IDMObjects.idmObjectType
Dim Class(1) As String
Dim oErrorManager As idmError.ErrorManager
' General purpose error handler - will show stack history
' of errors via MsgBoxes
Public Sub ShowError()
Dim oErrCollect As idmError.Errors
Dim oError As idmError.Error
Dim iCnt As Integer
Set oErrCollect = oErrorManager.Errors
If oErrCollect.Count > 1 Then
    iCnt = 1
    For Each oError In oErrCollect
        MsgBox "Error " & iCnt & ": " & oError.Description & " : " & Hex(oError.Number)
        iCnt = iCnt + 1
    Next
Else
    If oErrCollect.Count = 1 Then
        oErrorManager.ShowErrorDialog
    Else
        If Err.Number <> 0 Then
            MsgBox Err.Description & " : " & Err.Number
        End If
    End If
End If
End Sub
' Utility function for displaying property types
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

'Populates the read-only box to show what type of class info
' is being displayed
Private Sub ShowClassInfo()
Dim sTemp As String
Select Case CurrObjType
    Case idmObjTypeDocument
        sTemp = "Document "
    Case idmObjTypeFolder
        sTemp = "Folder "
    Case idmObjTypeAnnotation
        sTemp = "Annotation "
    Case idmObjTypeStoredSearch
        sTemp = "Stored search "
    Case idmObjTypeCustomObject
        sTemp = "Custom object"
End Select
txtClassInfo = sTemp & " class info for " & oLib.Label
End Sub

Private Sub Form_Load()
    Set oErrorManager = CreateObject("IDMError.ErrorManager")
    IDMListView1.AddItems oHood.Libraries, -1
    CurrObjType = idmObjTypeDocument
    rbDocument.Value = True
    txtVersionId = ""
    ' Disable options until library is selected
    rbDocument.Enabled = False
    rbFolder.Enabled = False
    rbAnnotation.Enabled = False
    rbStoredSearch.Enabled = False
    rbCustomObject.Enabled = False
    btnShowClasses.Enabled = False
    btnShowProperties.Enabled = False
   
End Sub
' Clean up any objects which have been instantiated with new
' or CreateObject
Private Sub Form_Unload(Cancel As Integer)
Set oErrorManager = Nothing
End Sub
' Handle click events in the Library selection listbox
Private Sub IDMListView1_Click()
Call HandleLibrarySelect(IDMListView1.SelectedItem)
End Sub

' Handle selection changes in the Library list
Private Sub HandleLibrarySelect(ByVal item As Object)  ', ByVal Key As Long, ByVal Selected As Boolean, ByVal objType As IDMListView.idmObjectType)
 Dim LoggedOn As Boolean
    Set oLib = item
    On Error GoTo Problems
    ' Make sure we are logged on
    LoggedOn = True
    If Not oLib.GetState(idmLibraryLoggedOn) Then
        LoggedOn = oLib.Logon(, , , idmLogonOptWithUI)
    End If
    ' Get the UI controls initialized
    ListView1.ListItems.Clear
    ListView2.ListItems.Clear
    ListView3.ListItems.Clear
    txtObjectID = ""
    txtDocClass = ""
    btnShowClasses.Enabled = True
    rbDocument.Enabled = True
    rbFolder.Enabled = True
    If LoggedOn Then
        rbAnnotation.Enabled = oLib.Supports(idmSupportsAnnotations)
        rbStoredSearch.Enabled = oLib.Supports(idmSupportsStoredSearch)
        rbCustomObject.Enabled = oLib.Supports(idmSupportsCustomObject)
        txtVersionNumber.Enabled = oLib.Supports(idmSupportsDocVersions)
        ' Fill the read-only info field to establish UI context
        Call ShowClassInfo
        End If
    Exit Sub
Problems:
    Call ShowError
 
End Sub

' Handle selection of a class name in the left-most listbox
Private Sub ListView1_ItemClick(ByVal item As ComctlLib.ListItem)
    Class(0) = item
    ListView2.ListItems.Clear
    ShowClassProperties
    txtSelectedClass = Class(0)
End Sub
' Populate the left-most listbox with class information
Private Sub ShowClasses(ByVal objType As idmObjectType)
    If oLib Is Nothing Then
        MsgBox "Select a Library"
    Else
        Dim oClass As IDMObjects.ClassDescription
        Dim oClasses As IDMObjects.ObjectSet
        Set oClasses = oLib.FilterClassDescriptions(objType)
        For Each oClass In oClasses
            ListView1.ListItems.Add , , oClass.Name
        Next
        CurrObjType = objType
    End If
End Sub
' Populate the center listbox control with properties of a specific
' class
Private Sub ShowClassProperties()
    If Not Class(0) = "" Then
        Dim oClassDesc As IDMObjects.ClassDescription
        Set oClassDesc = oLib.GetObject(idmObjTypeClassDesc, Class(0), CurrObjType)
        Dim oPropDesc As IDMObjects.PropertyDescription
        Dim oPropDescs As IDMObjects.PropertyDescriptions
        Set oPropDescs = oClassDesc.PropertyDescriptions
        For Each oPropDesc In oPropDescs
            Dim item As ListItem
            Set item = ListView2.ListItems.Add(, , oPropDesc.Name)
            item.SubItems(1) = FormatDataType(oPropDesc.TypeID)
            item.SubItems(2) = oPropDesc.GetState(idmPropRequired)
        Next
    End If
End Sub
' Start of section for handling mouse clicks on the radio buttons
Private Sub rbDocument_Click()
    CurrObjType = idmObjTypeDocument
    btnShowClasses.Enabled = True
    txtObjectID.Enabled = True
    Call ShowClassInfo
    ListView1.ListItems.Clear
    ListView2.ListItems.Clear
    ListView3.ListItems.Clear
End Sub

Private Sub rbFolder_Click()
    CurrObjType = idmObjTypeFolder
    btnShowClasses.Enabled = True
    txtObjectID.Enabled = True
    Call ShowClassInfo
    ListView1.ListItems.Clear
    ListView2.ListItems.Clear
    ListView3.ListItems.Clear
End Sub

Private Sub rbAnnotation_Click()
    CurrObjType = idmObjTypeAnnotation
    btnShowClasses.Enabled = True
    txtObjectID.Enabled = False
    btnShowProperties.Enabled = False
    Call ShowClassInfo
    ListView1.ListItems.Clear
    ListView2.ListItems.Clear
    ListView3.ListItems.Clear
End Sub

Private Sub rbStoredSearch_Click()
    CurrObjType = idmObjTypeStoredSearch
    btnShowClasses.Enabled = False
    txtObjectID.Enabled = True
    Call ShowClassInfo
    ListView1.ListItems.Clear
    ListView2.ListItems.Clear
    ListView3.ListItems.Clear
End Sub

Private Sub rbCustomObject_Click()
    CurrObjType = idmObjTypeCustomObject
    btnShowClasses.Enabled = False
    txtObjectID.Enabled = True
    Call ShowClassInfo
    ListView1.ListItems.Clear
    ListView2.ListItems.Clear
    ListView3.ListItems.Clear
End Sub
' Section for handling button clicks
Private Sub btnShowClasses_Click()
    ListView1.ListItems.Clear
    ShowClasses CurrObjType
End Sub

Private Sub btnShowProperties_Click()
    On Error GoTo ErrorHandler
    Dim oIDMObject As Object
    Dim oProp As IDMObjects.Property
    Dim item As ListItem
    Dim ObjId As Variant
    Dim FolderName As String
    Dim check
    
    If txtObjectID = "" Then
        MsgBox ("Must specify an object id")
        Exit Sub
    End If
    If txtVersionNumber <> "" Then
        Dim MyString As String
        MyString = txtObjectID & ":" & txtVersionNumber
        Set oIDMObject = oLib.GetObject(CurrObjType, MyString)
    Else
        check = IsNumeric(txtObjectID.Text)
        If (check = True) And (oLib.SystemType = idmSysTypeIS) Then
            ObjId = CLng(txtObjectID)
        Else
            ObjId = txtObjectID
        End If
        Set oIDMObject = oLib.GetObject(CurrObjType, ObjId)
    End If
    ListView3.ListItems.Clear
    For Each oProp In oIDMObject.Properties
        Set item = ListView3.ListItems.Add(, , oProp.Name)
        If oProp.Value = Null Then
            item.SubItems(1) = "<NULL>"
        Else
            item.SubItems(1) = oProp.FormatValue
        End If
    Next
    If Not oIDMObject.ClassDescription Is Nothing Then
        txtDocClass = oIDMObject.ClassDescription.Name
    Else
        txtDocClass = ""
    End If
    Exit Sub
ErrorHandler:
    Call ShowError
End Sub

Private Sub txtObjectID_Change()
btnShowProperties.Enabled = True
End Sub
