VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form FormProperty 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Setting Document Properties"
   ClientHeight    =   6045
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8235
   LinkTopic       =   "Form1"
   ScaleHeight     =   6045
   ScaleWidth      =   8235
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtEdit 
      Height          =   615
      Left            =   360
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   7320
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton BtnQuit 
      Height          =   495
      Left            =   6000
      Picture         =   "FormCommit.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Cancel/Exit"
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton BtnDone 
      Height          =   495
      Left            =   6000
      Picture         =   "FormCommit.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Done"
      Top             =   360
      Width           =   975
   End
   Begin VB.ComboBox CmbClasses 
      Height          =   315
      Left            =   2160
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   1140
      Width           =   2775
   End
   Begin VB.ComboBox CmbLibraries 
      Height          =   315
      Left            =   2160
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   360
      Width           =   3135
   End
   Begin MSFlexGridLib.MSFlexGrid FlexGrid 
      Height          =   4815
      Left            =   360
      TabIndex        =   4
      Top             =   2520
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   8493
      _Version        =   393216
      Rows            =   100
      Cols            =   4
      FixedCols       =   3
      AllowUserResizing=   3
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Height          =   2175
      Left            =   360
      TabIndex        =   6
      Top             =   120
      Width           =   7455
      Begin VB.TextBox txtFolderName 
         Height          =   285
         Left            =   1800
         TabIndex        =   9
         Top             =   1680
         Width           =   2655
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Folder:"
         Height          =   255
         Left            =   720
         TabIndex        =   10
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Classes:"
         Height          =   375
         Left            =   720
         TabIndex        =   8
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Library:"
         Height          =   255
         Left            =   720
         TabIndex        =   7
         Top             =   240
         Width           =   735
      End
   End
End
Attribute VB_Name = "FormProperty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Neighborhood As New IDMObjects.Neighborhood
Dim Libraries As IDMObjects.ObjectSet
Dim CurrentClass As String
Dim CurrentPropDescriptions As IDMObjects.PropertyDescriptions
Dim RequiredProps As Collection

Const ReqdColLoc = 1
Const TypeColLoc = 2
Const InputColLoc = 3
Dim InputColWidth As Integer ' in twips

' Function to handle the squirrely index in TextArray property
Function FaIndex(Row As Integer, col As Integer) As Long
     FaIndex = Row * FlexGrid.Cols + col
End Function

' Start data entry in cell
Sub FlexGrid_KeyPress(KeyAscii As Integer)
If FlexGrid.col = InputColLoc Then
    FlexGridEdit FlexGrid, txtEdit, KeyAscii
End If
End Sub
' Handle double click in a cell
Sub FlexGrid_DblClick()
If FlexGrid.col = InputColLoc Then
    FlexGridEdit FlexGrid, txtEdit, 32 ' Simulate a space.
End If
End Sub
' Fire up edit box for data entry
Sub FlexGridEdit(MSFlexGrid As Control, _
    Edt As Control, KeyAscii As Integer)

    ' Use the character that was typed.
    Select Case KeyAscii

    ' A space means edit the current text.
    Case 0 To 32
        Edt = MSFlexGrid
        Edt.SelStart = 1000

    ' Anything else means replace the current text.
    Case Else
        Edt = Chr(KeyAscii)
        Edt.SelStart = 1
    End Select

    ' Show edit box at the right place.
    Edt.Move MSFlexGrid.Left + MSFlexGrid.CellLeft, MSFlexGrid.Top + MSFlexGrid.CellTop, _
        MSFlexGrid.CellWidth, MSFlexGrid.CellHeight
    Edt.Visible = True

    ' And let it work.
    Edt.SetFocus
End Sub

' Capture input data from edit box and validate
Sub CaptureData()
Dim NewWidth As Integer
Dim Ok As Boolean
Dim TypeLoc As Integer
Ok = False
' Do a quick validation on the data

' Type is located one cell left of the input cell
TypeLoc = FaIndex(FlexGrid.Row, TypeColLoc)
If Len(txtEdit) > 0 Then
    Select Case FlexGrid.TextArray(TypeLoc)
        Case "String"
            Ok = True
        Case "Date"
            Ok = IsDate(txtEdit)
        Case "Short", "Long", "Integer", "Double", "Unsigned Long", _
          "Unsigned Short"
            Ok = IsNumeric(txtEdit)
        Case Else
            Ok = True    ' what the hell
    End Select
Else
    Ok = True
End If
If Ok Then
    FlexGrid = txtEdit
Else
    Beep
    txtEdit.Visible = False
    Exit Sub
End If
txtEdit.Visible = False
NewWidth = TextWidth(txtEdit.Text)
If NewWidth > InputColWidth Then
    ' Width is always a tad small, so add 100 to be safe
    FlexGrid.ColWidth(InputColLoc) = NewWidth + 300
    InputColWidth = FlexGrid.ColWidth(InputColLoc)
End If

End Sub
' Get edited data if there is any
Sub FlexGrid_GotFocus()
If txtEdit.Visible Then
    Call CaptureData
End If
    
End Sub
' Get edited data if there is any
Sub FlexGrid_LeaveCell()
If txtEdit.Visible Then
    Call CaptureData
End If
    
End Sub

' Routines for handling text editing
Sub txtEdit_KeyPress(KeyAscii As Integer)
' Delete returns to get rid of beep.
If Chr(KeyAscii) = vbCr Then KeyAscii = 0
End Sub

Sub txtEdit_KeyDown(KeyCode As Integer, _
Shift As Integer)
    EditKeyCode FlexGrid, txtEdit, KeyCode, Shift
End Sub

' Sub for handling termination of user input
Sub EditKeyCode(MSFlexGrid As Control, Edt As _
Control, KeyCode As Integer, Shift As Integer)

    ' Standard edit control processing.
    Select Case KeyCode

    Case 27 ' ESC: hide, return focus to MSFlexGrid.
        Edt.Visible = False
        MSFlexGrid.SetFocus

    Case 13 ' ENTER return focus to MSFlexGrid.
        MSFlexGrid.SetFocus

    Case 38     ' Up.
        MSFlexGrid.SetFocus
        DoEvents
        If MSFlexGrid.Row > MSFlexGrid.FixedRows Then
            MSFlexGrid.Row = MSFlexGrid.Row - 1
        End If

    Case 40     ' Down.
        MSFlexGrid.SetFocus
        DoEvents
        If MSFlexGrid.Row < MSFlexGrid.Rows - 1 Then
            MSFlexGrid.Row = MSFlexGrid.Row + 1
        End If
    End Select
End Sub
    

Private Sub LibLogon(oLibrary As IDMObjects.Library)
    On Error GoTo Errorhandler  ' Enable error-handling routine.
    
    If Not (oLibrary.GetState(idmLibraryLoggedOn)) Then
        oLibrary.Logon "", "", "", idmLogonOptWithUI
    End If
    Exit Sub        ' Exit to avoid handler.
    
Errorhandler:
    Select Case Err.Number
        Case 64518
            MsgBox "Library not available", vbOKOnly, "Logon Failure: " & oLibrary.Label
        Case Else
            MsgBox Err.Description & Err.Number
    End Select
End Sub

Private Sub BtnCancel_Click()
Me.Tag = "0"    ' Signal user bailed out
Unload Me
End Sub
' Done button - check to be sure required fields are not empty
Private Sub BtnDone_Click()
Dim Row As Variant
Dim TheRow As Integer
Dim ErrSeen As Boolean
ErrSeen = False
FlexGrid.col = 0 ' reset the current col, this seems to complete entry if user didn't press enter
If Online Then
    For Each Row In RequiredProps
        TheRow = Row
        If Len(FlexGrid.TextArray(FaIndex(TheRow, InputColLoc))) = 0 Then
            MsgBox ("Required field is blank...")
            FlexGrid.Row = TheRow
            FlexGrid.col = InputColLoc
            FlexGrid.SetFocus
            ErrSeen = True
            Exit For
        End If
    Next
End If
If Not ErrSeen Then
    If Online Then
        ' Now save all this stuff as Doc properties
        Dim Doc As IDMObjects.Document
        Set Doc = CurrentLib.CreateObject(idmObjTypeDocument, CurrentClass)
        Dim PropDesc As IDMObjects.PropertyDescription
        Dim iRow As Integer
        Dim CellLoc As Integer
        iRow = 1
        For Each PropDesc In CurrentPropDescriptions
            If Not PropDesc.GetState(idmPropReadOnly) Then
                CellLoc = FaIndex(iRow, InputColLoc)
                If Len(FlexGrid.TextArray(CellLoc)) > 0 Then
                    Doc.Properties(PropDesc.Name).Value = _
                        FlexGrid.TextArray(CellLoc)
                End If
                iRow = iRow + 1
            End If
        Next
        Set FinalList(CurrentDocInx) = Doc
        If txtFolderName <> "" Then
            FolderList(CurrentDocInx) = txtFolderName
        End If
    End If
    Me.Tag = "1"      ' Everything completed
    Me.Visible = False
End If
End Sub

Private Sub BtnQuit_Click()
Me.Visible = False
End Sub
' This is the routine for filling the FlexGrid control with
' the property names, data types, and required flags
Private Sub InitGrid(FlexGrid As MSFlexGrid, Online As Boolean)
Dim GridFormat As String
Dim ActualRow As Integer
Dim PropDesc As IDMObjects.PropertyDescription

FlexGrid.Clear
FlexGrid.Rows = 100    ' We dont know yet
FlexGrid.FormatString = "Property Name|Req'd|Type        |Value" + _
    ";"
GridFormat = "Property|Req'd|Type          |Value;"
ActualRow = 0
Set RequiredProps = New Collection
If Online Then
    For Each PropDesc In CurrentPropDescriptions
        If Not PropDesc.GetState(idmPropReadOnly) Then
            ActualRow = ActualRow + 1
            ' Property name
            GridFormat = GridFormat + "|" + PropDesc.Label
            ' Property data type
            FlexGrid.TextMatrix(ActualRow, TypeColLoc) = FormatDataType(PropDesc.TypeID)
            If PropDesc.GetState(idmPropRequired) Then
                FlexGrid.Row = ActualRow
                FlexGrid.col = ReqdColLoc
                FlexGrid.CellPictureAlignment = 0
                Set FlexGrid.CellPicture = LoadPicture _
                    (HomeDirectory + "\Reqd.ico")
                RequiredProps.Add (ActualRow)
            End If
        End If
    Next
Else    ' Offline - fake some stuff
    GridFormat = GridFormat + "|Property 1|Property 2|Property 3|"
End If
FlexGrid.FormatString = GridFormat

' Set the input cursor
If ActualRow > 0 Or Not Online Then
    FlexGrid.Row = 2
    FlexGrid.col = 1
    FlexGrid.Row = 1
    FlexGrid.col = InputColLoc
    If Online Then
        FlexGrid.Rows = ActualRow + 1
    Else
        FlexGrid.Rows = 4
    End If
    FlexGrid.Enabled = True
Else
    ' Maybe there are no modifiable properties
    FlexGrid.Enabled = False
End If
' Stretch the width of the input column
InputColWidth = FlexGrid.Width - FlexGrid.CellLeft
FlexGrid.ColWidth(InputColLoc) = InputColWidth
FlexGrid.ColAlignment(InputColLoc) = 0   ' Force left justify
End Sub
Private Sub CmbClasses_Click()
Dim ClassList(1) As String

CurrentClass = CmbClasses.List(CmbClasses.ListIndex)
ClassList(0) = CmbClasses.List(CmbClasses.ListIndex)
If Online Then
    Set CurrentPropDescriptions = CurrentLib.FilterPropertyDescriptions(idmObjTypeDocument, _
        ClassList)
Else
    Set CurrentPropDescriptions = Nothing
End If
    
' Now set up the grid control for this class
Call InitGrid(FlexGrid, Online)
End Sub

Private Sub CmbLibraries_Click()
CmbClasses.Clear
If Online Then
    Set CurrentLib = Libraries(CmbLibraries.ListIndex + 1)
    'make sure we are logged on to the selected Library
    Call LibLogon(CurrentLib)
    
    If (CurrentLib.GetState(idmLibraryLoggedOn)) Then
    
        ' Fill the combo box with document class description names
        Dim Class As IDMObjects.ClassDescription
        Dim DocClasses As IDMObjects.ObjectSet
        Set DocClasses = CurrentLib.FilterClassDescriptions(idmObjTypeDocument, idmFilterClassAllowsInstance)
        For Each Class In DocClasses
            CmbClasses.AddItem Class.Name
        Next
    Else
        Exit Sub
    End If
Else
    CmbClasses.AddItem "Bogus"
End If
CmbClasses.ListIndex = 0
End Sub

Private Sub Form_Load()
' Get Error Manager running in case we hit problems
Set oErrManager = CreateObject("IDMError.ErrorManager")
If MsgBox("Are you really online?", vbYesNo) = vbYes Then
    Set Libraries = Neighborhood.Libraries
    Dim Library As IDMObjects.Library
    CmbLibraries.Clear
    CmbClasses.Clear
    For Each Library In Libraries
        CmbLibraries.AddItem Library.Label
    Next
    Online = True
Else
    CmbLibraries.AddItem ("BogusLib")
    Online = False
End If
' CmbLibraries.ListIndex = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim Library As IDMObjects.Library
    If Online Then
        For Each Library In Libraries
            If Library.GetState(idmLibraryLoggedOn) Then
                Library.Logoff
            End If
         Next
    End If
    ' Clean up object instances
    Set Neighborhood = Nothing
    Set Libraries = Nothing
    Set CurrentLib = Nothing
    Set CurrentPropDescriptions = Nothing
    Set oErrManager = Nothing

End Sub

