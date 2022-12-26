VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{6FBA474E-43AC-11CE-9A0E-00AA0062BB4C}#1.0#0"; "SYSINFO.OCX"
Begin VB.Form FormMain 
   Caption         =   "Panagon IDM Property Descriptions"
   ClientHeight    =   9240
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13515
   LinkTopic       =   "Form1"
   ScaleHeight     =   9240
   ScaleWidth      =   13515
   StartUpPosition =   3  'Windows Default
   Begin SysInfoLib.SysInfo SysInfo1 
      Left            =   1080
      Top             =   6600
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Frame fmPropDesc 
      Height          =   9015
      Left            =   5040
      TabIndex        =   1
      Top             =   0
      Width           =   8415
      Begin VB.CheckBox chkChoice 
         Caption         =   "Choice"
         Height          =   255
         Left            =   120
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   2040
         Width           =   2295
      End
      Begin VB.CheckBox chkChoicePaging 
         Caption         =   "ChoicePaging"
         Height          =   255
         Left            =   120
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   2280
         Width           =   2295
      End
      Begin VB.CheckBox chkPropBatchTotal 
         Caption         =   "PropBatchTotal"
         Height          =   255
         Left            =   120
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   2520
         Width           =   2295
      End
      Begin VB.CheckBox chkPropCustom 
         Caption         =   "PropCustom"
         Height          =   255
         Left            =   120
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   2760
         Width           =   2295
      End
      Begin VB.CheckBox chkPropDisplayable 
         Caption         =   "PropDisplayable"
         Height          =   255
         Left            =   120
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   3000
         Width           =   2295
      End
      Begin VB.CheckBox chkPropHasCVL 
         Caption         =   "PropHasCVL"
         Height          =   255
         Left            =   120
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   3480
         Width           =   2295
      End
      Begin VB.CheckBox chkPropHasDefault 
         Caption         =   "PropHasDefault"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   3720
         Width           =   2295
      End
      Begin VB.CheckBox chkPropHasEVL 
         Caption         =   "PropHasEVL"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   3960
         Width           =   2295
      End
      Begin VB.CheckBox chkPropKey 
         Caption         =   "PropKey"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   4200
         Width           =   2295
      End
      Begin VB.CheckBox chkPropMultiValue 
         Caption         =   "PropMultiValue"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   4440
         Width           =   2295
      End
      Begin VB.CheckBox chkPropMustPick 
         Caption         =   "PropMustPick"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   4680
         Width           =   2295
      End
      Begin VB.CheckBox chkPropMustSetByID 
         Caption         =   "PropMustSetByID"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   4920
         Width           =   2295
      End
      Begin VB.CheckBox chkPropReadOnly 
         Caption         =   "PropReadOnly"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   5400
         Width           =   2295
      End
      Begin VB.CheckBox chkPropRendezvous 
         Caption         =   "PropRendezvous"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   5640
         Width           =   2295
      End
      Begin VB.CheckBox chkPropRequired 
         Caption         =   "PropRequired"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   5880
         Width           =   2295
      End
      Begin VB.CheckBox chkPropSearchable 
         Caption         =   "PropSearchable"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   6120
         Width           =   2295
      End
      Begin VB.CheckBox chkPropSelectable 
         Caption         =   "PropSelectable"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   6360
         Width           =   2295
      End
      Begin VB.CheckBox chkPropShouldDisplay 
         Caption         =   "PropShouldDisplay"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   6600
         Width           =   2295
      End
      Begin VB.CheckBox chkPropSupportsFPNumber 
         Caption         =   "PropSupportsFPNumber"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   7080
         Width           =   2295
      End
      Begin VB.CheckBox chkPropUnique 
         Caption         =   "PropUnique"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   7320
         Width           =   2295
      End
      Begin VB.CheckBox chkPropUpcased 
         Caption         =   "PropUpcased"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   7560
         Width           =   2295
      End
      Begin VB.CheckBox chkPropVerifyRequired 
         Caption         =   "PropVerifyRequired"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   8040
         Width           =   2295
      End
      Begin VB.CheckBox chkPropVersion 
         Caption         =   "PropVersion"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   8280
         Width           =   2295
      End
      Begin VB.TextBox txtDefaultValue 
         Height          =   285
         Left            =   5280
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   360
         Width           =   2895
      End
      Begin VB.CheckBox chkSynthetic 
         Caption         =   "Synthetic"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   1800
         Width           =   2295
      End
      Begin VB.CheckBox chkShowChoices 
         Caption         =   "Show Choices (first 1000 only)"
         Height          =   255
         Left            =   2640
         TabIndex        =   12
         Top             =   1800
         Width           =   4095
      End
      Begin VB.TextBox txtSize 
         Height          =   285
         Left            =   5280
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   720
         Width           =   2895
      End
      Begin VB.TextBox txtSearchWith 
         Height          =   285
         Left            =   5280
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   1080
         Width           =   2895
      End
      Begin VB.CheckBox chkPropHasAccess 
         Caption         =   "PropHasAccess"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   3240
         Width           =   2295
      End
      Begin VB.CheckBox chkPropSpecialFormatting 
         Caption         =   "PropSpecialFormatting"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   6840
         Width           =   2295
      End
      Begin VB.CheckBox chkPropValueHoldsChoiceID 
         Caption         =   "PropValueHoldsChoiceID"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   7800
         Width           =   2295
      End
      Begin VB.CheckBox chkPropOnlyDate 
         Caption         =   "PropOnlyDate"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   5160
         Width           =   2295
      End
      Begin VB.TextBox txtTypeID 
         Height          =   285
         Left            =   960
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1065
         Width           =   3015
      End
      Begin VB.TextBox txtLabel 
         Height          =   285
         Left            =   960
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   705
         Width           =   3015
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   960
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   360
         Width           =   3015
      End
      Begin VB.TextBox txtMask 
         Height          =   285
         Left            =   960
         TabIndex        =   2
         Top             =   1410
         Width           =   3015
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   6615
         Left            =   2640
         TabIndex        =   13
         Top             =   2040
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   11668
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Label Label1 
         Caption         =   "Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   45
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Label:"
         Height          =   255
         Left            =   120
         TabIndex        =   44
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Type ID:"
         Height          =   255
         Left            =   120
         TabIndex        =   43
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Default value:"
         Height          =   255
         Left            =   4080
         TabIndex        =   42
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Size:"
         Height          =   255
         Left            =   4080
         TabIndex        =   41
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "Search with:"
         Height          =   255
         Left            =   4080
         TabIndex        =   40
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "Mask:"
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   1440
         Width           =   1215
      End
   End
   Begin MSComctlLib.TreeView tvProps 
      Height          =   6375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   11245
      _Version        =   393217
      Indentation     =   529
      LabelEdit       =   1
      Style           =   7
      ImageList       =   "imlSmallIcons"
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList imlSmallIcons 
      Left            =   120
      Top             =   6600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   13
      ImageHeight     =   13
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormMain.frx":0000
            Key             =   "closed"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormMain.frx":0172
            Key             =   "leaf"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormMain.frx":02E4
            Key             =   "open"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FormMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private idmErrManager As IDMError.ErrorManager

' Most of the data in the treeview is held in the tag property.
' All of the tags have one of these:
Const tagNeighborhood = "Neighborhood"
Const tagLibrary = "Library"
Const tagClass = "Class"
Const tagClassType = "ClassType"
Const tagPropDesc = "PropertyDescription"
Const tagPlaceholder = "placeHolder"

' And may have one each of these xml-like tags:
Const tagPopulated = "<pop/>"   ' In tag if this node has been populated. Use SetTagIsPopulated.
Const tagName = "<name>"        ' Stores the object name in <name>obj-name</name>
Const tagNameEnd = "</name>"    ' Use AddNameToTag and GetNameFromTag

' Text to use for the different class types
Const textAnnotationClass = "Annotation Classes"
Const textCustomObjectClass = "Custom Object Classes"
Const textDocumentClass = "Document Classes"
Const textFolderClass = "Folder Classes"
Const textLibraryClass = "Library Classes"
Const textStoredSearchClass = "Stored Search Classes"

' Map the data types to a string
Dim typeIDStrings(20) As String

' The currently selected property description
Private pd As PropertyDescription

' Keep our own list of libraries around.  If we used neighborhood.libraries,
' then we'd lose the logon every time we released the library.
Dim libraries() As IDMObjects.library
Dim numLibraries As Integer

Private Sub Form_Load()
    Set idmErrManager = New IDMError.ErrorManager
    AddLibraries
    
    ' Map idm types to strings
    typeIDStrings(0) = "idmTypeEmpty"
    typeIDStrings(1) = "idmTypeNull"
    typeIDStrings(2) = "idmTypeShort"
    typeIDStrings(3) = "idmTypeLong"
    typeIDStrings(4) = "idmTypeSingle"
    typeIDStrings(5) = "idmTypeDouble"
    typeIDStrings(6) = "idmTypeCurrency"
    typeIDStrings(7) = "idmTypeDate"
    typeIDStrings(8) = "idmTypeString"
    typeIDStrings(9) = "idmTypeObject"
    typeIDStrings(10) = "idmTypeError"
    typeIDStrings(11) = "idmTypeBoolean"
    typeIDStrings(12) = "idmTypeVariant"
    typeIDStrings(13) = "idmTypeUnknown"
    typeIDStrings(14) = "<not valid>"
    typeIDStrings(15) = "<not valid>"
    typeIDStrings(16) = "idmTypeCharacter"
    typeIDStrings(17) = "idmTypeByte"
    typeIDStrings(18) = "idmTypeUnsignedShort"
    typeIDStrings(19) = "idmTypeUnsignedLong"
    
    ' set up choices listview -- widths are just a guess,
    ' trying to avoid an hscrollbar
    Dim widthSoFar As Integer
    widthSoFar = SysInfo1.ScrollBarSize + 60
    Dim colX As ColumnHeader
    Set colX = ListView1.ColumnHeaders.Add
    colX.text = "ID"
    colX.Width = SysInfo1.ScrollBarSize * 2
    widthSoFar = widthSoFar + colX.Width
    Set colX = ListView1.ColumnHeaders.Add
    colX.text = "Value"
    colX.Width = SysInfo1.ScrollBarSize * 6
    widthSoFar = widthSoFar + colX.Width
    Set colX = ListView1.ColumnHeaders.Add
    colX.text = "Formatted Value"
    colX.Width = ListView1.Width - widthSoFar


End Sub

' Add the class types under a library
Private Sub AddClassTypes(theNode As Node, library As library)
    Dim mNode As Node
    
    Set mNode = tvProps.Nodes.Add(theNode, tvwChild, , textAnnotationClass, "closed")
    mNode.tag = tagClassType
    Set mNode = tvProps.Nodes.Add(theNode, tvwChild, , textCustomObjectClass, "closed")
    mNode.tag = tagClassType
    Set mNode = tvProps.Nodes.Add(theNode, tvwChild, , textDocumentClass, "closed")
    mNode.tag = tagClassType
    Set mNode = tvProps.Nodes.Add(theNode, tvwChild, , textFolderClass, "closed")
    mNode.tag = tagClassType
    Set mNode = tvProps.Nodes.Add(theNode, tvwChild, , textLibraryClass, "closed")
    mNode.tag = tagClassType
    Set mNode = tvProps.Nodes.Add(theNode, tvwChild, , textStoredSearchClass, "closed")
    mNode.tag = tagClassType
End Sub

Private Sub AddLibraries()
On Error GoTo errHandler

    'Get the list of libraries from the neighborhood, put them in the treeview and list
    
    Dim neigh As New IDMObjects.Neighborhood
    Dim libList As IDMObjects.ObjectSet
    Dim tmpLib As IDMObjects.library
    Set libList = neigh.libraries
    Dim mNode As Node
    Dim ii As Integer
    
    ' Add FileNET neighborhood as root
    Set mNode = tvProps.Nodes.Add()
    mNode.text = "FileNET Neighborhood"
    mNode.tag = SetTagIsPopulated(tagNeighborhood)  ' Already populated neighborhood
    mNode.Image = "closed"
    
    ' Add libraries, keeping the objects in an array.
    numLibraries = libList.Count
    ReDim libraries(numLibraries)
    ii = 1
    For Each tmpLib In libList
        Set mNode = tvProps.Nodes.Add(1, tvwChild, tmpLib.name, tmpLib.Label, "closed")
        mNode.tag = tagLibrary
        Set libraries(ii) = tmpLib
        ii = ii + 1
        AddClassTypes mNode, tmpLib
    Next
    
    ' Sort and expand top node.
    tvProps.Nodes(1).Sorted = True
    tvProps.Nodes(1).Expanded = True
    Exit Sub
    
errHandler:
    If idmErrManager.Errors.Count > 0 Then
        idmErrManager.ShowErrorDialog
    Else
        MsgBox Err.Description
    End If
End Sub

' Put in the the real property descriptions after removing the dummy node
Private Sub PopulateClass(lib As library, theNode As Node)
On Error GoTo errHandler

    Dim oldPointer As Integer
    oldPointer = MousePointer
    MousePointer = vbHourglass
    
    Dim childNode As Node
    
    ' Remove the dummy node
    Set childNode = theNode.Child
    tvProps.Nodes.Remove childNode.Index
    
    ' Put in the property descriptions for the class
    ' First get this class name
    Dim propDescs As PropertyDescriptions, propDesc As PropertyDescription, classname As String
    classname = GetNameFromTag(theNode.tag)
    
    ' Now get the object type from the parent node
    Dim objType As idmObjectType
    GetObjectTypeFromText theNode.Parent.text, objType
    
    ' An array is REQUIRED here
    Dim classNames(1) As String
    classNames(1) = classname
    Set propDescs = lib.FilterPropertyDescriptions(objType, classNames)
    
    ' Finally, add the property descriptions
    For Each propDesc In propDescs
        Set childNode = tvProps.Nodes.Add(theNode, tvwChild, , GetPropDescLabel(propDesc), "leaf")
        childNode.tag = AddNameToTag(tagPropDesc, propDesc.name)
    Next propDesc
    
    ' And sort them
    theNode.Sorted = True
    
    theNode.tag = SetTagIsPopulated(theNode.tag)
    MousePointer = oldPointer
    Exit Sub
    
errHandler:
    If idmErrManager.Errors.Count > 0 Then
        idmErrManager.ShowErrorDialog
    Else
        MsgBox Err.Description
    End If
    theNode.Expanded = False  ' Problem, unexpand node
    MousePointer = oldPointer
End Sub

'Add the Grandchildren of the library: the classes under the class types
Private Sub PopulateLibrary(lib As library, theNode As Node)
On Error GoTo errHandler

    Dim oldPointer As Integer
    oldPointer = MousePointer
    
    ' Logon
    If Not lib.GetState(idmLibraryLoggedOn) Then
        lib.Logon "", "", "", idmLogonOptWithUI
    End If
    
    MousePointer = vbHourglass
    
    If lib.GetState(idmLibraryLoggedOn) Then
        ' Handle each type of class (doc, anno, etc)
        Dim childNode As Node
        If theNode.Children Then
            Set childNode = theNode.Child
            Do While Not childNode Is Nothing
                ' This is where the classes get added
                PopulateClasses lib, childNode
                Set childNode = childNode.Next
            Loop
        End If
            
        theNode.tag = SetTagIsPopulated(theNode.tag)
    Else
        theNode.Expanded = False  ' Not logged on, unexpand node
    End If
    
    MousePointer = oldPointer
    Exit Sub
    
errHandler:
    If idmErrManager.Errors.Count > 0 Then
        idmErrManager.ShowErrorDialog
    Else
        MsgBox Err.Description
    End If
    theNode.Expanded = False  ' Problem, unexpand node
    MousePointer = oldPointer
End Sub

Private Sub Form_Resize()
    ' tvProps.left sets the border for all controls
    Dim border As Integer
    border = tvProps.Left
    
    ' Fix height of treeview
    If ScaleHeight > border Then
        tvProps.Height = ScaleHeight - border
    End If
    
    ' Give the property description details enough room
    Dim l As Integer
    l = ScaleWidth - fmPropDesc.Width - border
    fmPropDesc.Left = l
    ' And the treeview what's left
    If l > 2 * border Then
        tvProps.Width = l - 2 * border
    End If

End Sub

' Walk up the tree to get a library
Private Function GetLibFromNode(theNode As Node) As library
    Dim tmpLib As library
    ' Find library node.  Walk up tree until we get to a library type
    Dim libNode As Node
    Set libNode = theNode
    Do While GetBaseTypeFromTag(libNode.tag) <> tagLibrary
        Set libNode = libNode.Parent
    Loop
    
    ' Now find the library in the array of libraries
    Dim ii As Integer
    ii = 1
    Do While ii <= numLibraries
        If libraries(ii).name = libNode.key Then
            Set tmpLib = libraries(ii)
            Exit Do
        End If
        ii = ii + 1
    Loop

    Set GetLibFromNode = tmpLib
End Function

' Expand a node; only libraries and classes
Private Sub tvProps_Expand(ByVal theNode As MSComctlLib.Node)
    If GetIsPopulatedFromTag(theNode.tag) = False Then
        Dim tmpLib As library
        Set tmpLib = GetLibFromNode(theNode)
        
        If GetBaseTypeFromTag(theNode.tag) = tagLibrary Then
            PopulateLibrary tmpLib, theNode
        ElseIf GetBaseTypeFromTag(theNode.tag) = tagClass Then
            PopulateClass tmpLib, theNode
        End If
    End If
End Sub

' Set the selected property description and update the displayed state.
Private Sub tvProps_NodeClick(ByVal theNode As MSComctlLib.Node)
On Error GoTo errHandler

    If GetBaseTypeFromTag(theNode.tag) = tagPropDesc Then
        ' Build up the caption
        Dim caption As String
        caption = theNode.text
        Dim parentNode As Node
        Set parentNode = theNode.Parent
        Do While GetBaseTypeFromTag(parentNode.tag) <> tagNeighborhood
            caption = Trim(parentNode.text) & "\" & caption
            Set parentNode = parentNode.Parent
        Loop
        
        ' Set the caption
        Me.fmPropDesc.caption = caption
        
        ' Find out the type of propdesc (document, anno, etc).
        Set parentNode = theNode.Parent
        Do While GetBaseTypeFromTag(parentNode.tag) <> tagClassType
            Set parentNode = parentNode.Parent
        Loop
        Dim objType As idmObjectType
        GetObjectTypeFromText parentNode.text, objType
        
        ' Get the actual property description
        Dim tmpLib As library
        Set tmpLib = GetLibFromNode(theNode)
        Set pd = tmpLib.GetObject(idmObjTypePropDesc, GetNameFromTag(theNode.tag), objType)
        UpdatePropDescDetails
    End If
    Exit Sub
    
errHandler:
    If idmErrManager.Errors.Count > 0 Then
        idmErrManager.ShowErrorDialog
    Else
        MsgBox Err.Description
    End If
End Sub

' Add all of the class in the class type.  Also add property descriptions for each class,
' and for the class type
Private Sub PopulateClasses(lib As library, theNode As Node)
On Error GoTo errHandler
    If GetBaseTypeFromTag(theNode.tag) <> tagClassType Then Exit Sub
    
    Dim classes As IDMObjects.ObjectSet
    Dim objType As idmObjectType
    Dim filterClass As idmObjectType
    Dim propDescs As PropertyDescriptions, propDesc As PropertyDescription
    Dim childNode As Node
    
    GetObjectTypeFromText theNode.text, objType
    ' Not all libraries support all types.
    Err.Clear
    On Error Resume Next
    Set classes = lib.FilterClassDescriptions(objType, idmFilterClassAll)
    If Err.Number = 0 Then
        On Error GoTo errHandler
        
        ' OK, go ahead and add each class
        Dim aClass As ClassDescription
        For Each aClass In classes
            Dim mNode As Node
            Set mNode = tvProps.Nodes.Add(theNode, tvwChild, , GetClassLabel(aClass), "closed")
            mNode.tag = AddNameToTag(tagClass, aClass.name)
            ' We want the class to be expandable.  Add a dummy node
            Set mNode = tvProps.Nodes.Add(mNode, tvwChild, , "")
            mNode.tag = tagPlaceholder
        Next aClass
        
    End If
    
    ' And property descriptions for the class type (e.g. all documents).  Again, just
    ' use an empty object set for libraries that don't support the class type
    Err.Clear
    On Error Resume Next
    Set propDescs = lib.FilterPropertyDescriptions(objType)
    If Err.Number = 0 Then
        On Error GoTo errHandler
    
        For Each propDesc In propDescs
            Set childNode = tvProps.Nodes.Add(theNode, tvwChild, , GetPropDescLabel(propDesc), "leaf")
            childNode.tag = AddNameToTag(tagPropDesc, propDesc.name)
        Next propDesc
    End If
    
    ' Sort the classes and propdescs
    theNode.Sorted = True
    Exit Sub
    
errHandler:
    If idmErrManager.Errors.Count > 0 Then
        idmErrManager.ShowErrorDialog
    Else
        MsgBox Err.Description
    End If
End Sub

'    GetObjectTypeFromText node.text, objType, idmFilterClass
Private Function GetObjectTypeFromText(text As String, objType As idmObjectType)
    If text = textAnnotationClass Then
        objType = idmObjTypeAnnotation
    ElseIf text = textCustomObjectClass Then
        objType = idmObjTypeCustomObject
    ElseIf text = textDocumentClass Then
        objType = idmObjTypeDocument
    ElseIf text = textFolderClass Then
        objType = idmObjTypeFolder
    ElseIf text = textLibraryClass Then
        objType = idmObjTypeLibrary
    ElseIf text = textStoredSearchClass Then
        objType = idmObjTypeStoredSearch
    Else
        MsgBox "Internal error, unknown object type " & text
    End If
End Function

' Strips off any appended state information from the tag
Private Function GetBaseTypeFromTag(tag As String) As String
    Dim lessPos As Integer
    lessPos = InStr(1, tag, "<")
    If lessPos <> 0 Then
        GetBaseTypeFromTag = Left(tag, lessPos - 1)
    Else
        GetBaseTypeFromTag = tag
    End If
End Function

Private Function GetIsPopulatedFromTag(tag As String) As Boolean
    Dim popPos As Integer
    popPos = InStr(1, tag, tagPopulated)
    If popPos <> 0 Then
        GetIsPopulatedFromTag = True
    Else
        GetIsPopulatedFromTag = False
    End If
End Function

Private Function SetTagIsPopulated(tag As String) As String
    SetTagIsPopulated = tag & tagPopulated
End Function

Private Function AddNameToTag(tag As String, name As String) As String
    AddNameToTag = tag & tagName & name & tagNameEnd
End Function

Private Function GetNameFromTag(tag As String) As String
    Dim ret As String
    ret = ""
    Dim namePos As Integer
    namePos = InStr(1, tag, tagName)
    If namePos > 0 Then
        Dim nameEndPos As Integer
        nameEndPos = InStrRev(tag, tagNameEnd)
            If nameEndPos > 0 Then
                ret = Mid(tag, namePos + Len(tagName), nameEndPos - namePos - Len(tagName))
            End If
    End If
    GetNameFromTag = ret
End Function

' Put the name and the label together, if different.
Private Function GetPropDescLabel(propDesc As PropertyDescription) As String
    Dim retVal As String
    If propDesc.Label = propDesc.name Then
        retVal = propDesc.Label
    Else
        retVal = propDesc.name & " (" & propDesc.Label & ")"
    End If
    
    GetPropDescLabel = retVal
End Function

'Put a space on the front so that they get sorted before property descriptions
Private Function GetClassLabel(aClass As ClassDescription) As String
    GetClassLabel = " " & aClass.Label
End Function

Private Function MapState(state As Boolean) As Integer
    If state Then
        MapState = 1
    Else
        MapState = 0
    End If
End Function

' Update checkboxes, etc to show details of the property description
Public Sub UpdatePropDescDetails()
On Error GoTo errHandler
    Dim oldPointer As Integer
    oldPointer = MousePointer
    MousePointer = vbHourglass
    
    Dim tmpStr As String
    
    txtName = pd.name
    txtLabel = pd.Label
    txtTypeID = CStr(pd.typeid) + " - " + GetTypeIDName(pd.typeid)
    txtSize = CStr(pd.Size)
    ' Not all propdescs have a mask
    tmpStr = ""
    On Error Resume Next
    tmpStr = pd.GetExtendedProperty("F_DISPMASK")
    Err.Clear
    On Error GoTo errHandler
    txtMask = tmpStr
    
    chkChoice.Value = MapState(pd.GetState(idmChoice))
    chkChoicePaging.Value = MapState(pd.GetState(idmChoicePaging))
    chkPropBatchTotal = MapState(pd.GetState(idmPropBatchTotal))
    chkPropCustom = MapState(pd.GetState(idmPropCustom))
    chkPropDisplayable = MapState(pd.GetState(idmPropDisplayable))
    chkPropHasAccess = MapState(pd.GetState(idmPropHasAccess))
    chkPropHasCVL = MapState(pd.GetState(idmPropHasCVL))
    chkPropHasDefault = MapState(pd.GetState(idmPropHasDefault))
    chkPropHasEVL = MapState(pd.GetState(idmPropHasEVL))
    chkPropKey = MapState(pd.GetState(idmPropKey))
    chkPropMultiValue = MapState(pd.GetState(idmPropMultiValue))
    chkPropMustPick = MapState(pd.GetState(idmPropMustPick))
    chkPropMustSetByID = MapState(pd.GetState(idmPropMustSetByID))
    chkPropOnlyDate = MapState(pd.GetState(idmPropOnlyDate))
    chkPropReadOnly = MapState(pd.GetState(idmPropReadOnly))
    chkPropRendezvous = MapState(pd.GetState(idmPropRendezvous))
    chkPropRequired = MapState(pd.GetState(idmPropRequired))
    chkPropSearchable = MapState(pd.GetState(idmPropSearchable))
    ' SearchWith... are covered below
    chkPropSelectable = MapState(pd.GetState(idmPropSelectable))
    chkPropShouldDisplay = MapState(pd.GetState(idmPropShouldDisplay))
    chkPropSpecialFormatting = MapState(pd.GetState(idmPropSpecialFormatting))
    chkPropSupportsFPNumber = MapState(pd.GetState(idmPropSupportsFPNumber))
    chkPropUnique = MapState(pd.GetState(idmPropUnique))
    chkPropUpcased = MapState(pd.GetState(idmPropUpcased))
    chkPropValueHoldsChoiceID = MapState(pd.GetState(idmPropValueHoldsChoiceID))
    chkPropVerifyRequired = MapState(pd.GetState(idmPropVerifyRequired))
    chkPropVersion = MapState(pd.GetState(idmPropVersion))
    
    If pd.GetState(idmPropHasDefault) Then
        txtDefaultValue = pd.DefaultValue
    Else
        txtDefaultValue = ""
    End If
    
    'Help says: If GetState is True for idmPropReadOnly and False for both idmPropSearchable and idmPropSelectable, then the property is a synthetic property.
    If pd.GetState(idmPropReadOnly) = True And _
            pd.GetState(idmPropSearchable) = False And _
            pd.GetState(idmPropSelectable) = False Then
        chkSynthetic.Value = 1
    Else
        chkSynthetic.Value = 0
    End If
    
    SetSearchWith pd
    UpdateChoices
    
    MousePointer = oldPointer
    Exit Sub
    
errHandler:
    If idmErrManager.Errors.Count > 0 Then
        idmErrManager.ShowErrorDialog
    Else
        MsgBox Err.Description
    End If
    MousePointer = oldPointer
End Sub

' Update the choice list
Private Sub UpdateChoices()

    On Error GoTo errHandler
    Dim key
    Dim myChoices As IDMObjects.Choices
    Dim myPaging As IDMObjects.Paging
    Dim myChoice As IDMObjects.Choice
    key = Null
    Dim oldPointer As Integer
    oldPointer = MousePointer
    MousePointer = vbHourglass
    
    ' Reset listbox
    ListView1.ListItems.Clear
    ListView1.Sorted = False

    If Not (pd Is Nothing) Then
        ' valid property description has been selected
        If (Not (pd Is Nothing)) And (pd.GetState(idmChoice)) And (chkShowChoices.Value = 1) Then
            Do
                Set myChoices = pd.Choices
                If pd.GetState(idmChoicePaging) Then
                    Set myPaging = myChoices.Paging
                    myPaging.Size = 1000
                    
                    myPaging.NextPage key, idmForward
                End If
                'Process this page of choices
                For Each myChoice In myChoices
                    Dim itemX As ListItem
                    Set itemX = ListView1.ListItems.Add
                    itemX.text = CStr(myChoice.ID)
                    itemX.SubItems(1) = CStr(myChoice.Value)
                    itemX.SubItems(2) = CStr(pd.FormatValue(myChoice.Value))
                Next myChoice
                
                'Set key to last item on page
                If myChoices.Count > 0 Then
                    key = myChoices.Item(myChoices.Count).Value
                End If
            Loop While False 'Assumes paging: myChoices.Paging.Size <= myChoices.Count
        End If
        
    End If 'Not (pd Is Nothing)
    
    MousePointer = oldPointer
    Exit Sub
    
errHandler:
    If idmErrManager.Errors.Count > 0 Then
        idmErrManager.ShowErrorDialog
    Else
        MsgBox Err.Description
    End If
    MousePointer = oldPointer
End Sub

Private Sub chkShowChoices_Click()
    UpdateChoices
End Sub

Private Function GetTypeIDName(typeid As Long) As String
    Dim retVal As String
    If typeid > 0 And typeid <= 19 Then
        retVal = typeIDStrings(typeid)
    ElseIf typeid = 72 Then
        retVal = "idmTypeGuid"
    ElseIf typeid = 8192 Then
        retVal = "idmTypeArray"
    Else
        retVal = "<not valid>"
    End If
    GetTypeIDName = retVal
End Function

Private Sub SetSearchWith(pd As IDMObjects.PropertyDescription)
On Error GoTo errHandler

    Dim ss As String
    If pd.GetState(idmPropSearchWithEqualOperator) Then
        ss = ss + ", ="
    End If
    If pd.GetState(idmPropSearchWithGreaterOperator) Then
        ss = ss + ", >"
    End If
    If pd.GetState(idmPropSearchWithGreaterOrEqualOperator) Then
        ss = ss + ", >="
    End If
    If pd.GetState(idmPropSearchWithLessOperator) Then
        ss = ss + ", <"
    End If
    If pd.GetState(idmPropSearchWithLessOrEqualOperator) Then
        ss = ss + ", <="
    End If
    If pd.GetState(idmPropSearchWithLikeOperator) Then
        ss = ss + ", like"
    End If
    If pd.GetState(idmPropSearchWithNotEqualOperator) Then
        ss = ss + ", <>"
    End If
    If pd.GetState(idmPropSearchWithNotLikeOperator) Then
        ss = ss + ", not like"
    End If
    
    'Remove leading ", "
    If Len(ss) > 0 Then
        ss = Mid(ss, 3)
    End If
    
    txtSearchWith = ss
    Exit Sub
    
errHandler:
    If idmErrManager.Errors.Count > 0 Then
        idmErrManager.ShowErrorDialog
    Else
        MsgBox Err.Description
    End If
End Sub


