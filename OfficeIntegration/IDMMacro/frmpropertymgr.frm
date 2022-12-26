VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmPropertyMgr 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FileNET Property Manager"
   ClientHeight    =   3075
   ClientLeft      =   3000
   ClientTop       =   3495
   ClientWidth     =   4830
   Icon            =   "frmPropertyMgr.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   4830
   Begin MSFlexGridLib.MSFlexGrid grdPropDisp 
      Height          =   2535
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4825
      _ExtentX        =   8520
      _ExtentY        =   4471
      _Version        =   393216
      Rows            =   1
      Cols            =   3
      FixedCols       =   0
      BackColor       =   16777215
      BackColorBkg    =   16777215
      WordWrap        =   -1  'True
      ScrollTrack     =   -1  'True
      FocusRect       =   2
      GridLines       =   2
      SelectionMode   =   1
      AllowUserResizing=   3
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   400
      Left            =   3950
      TabIndex        =   5
      Top             =   2620
      Width           =   855
   End
   Begin VB.CommandButton cmdReplace 
      Caption         =   "Replace"
      Enabled         =   0   'False
      Height          =   400
      Left            =   1995
      TabIndex        =   4
      Top             =   2620
      Width           =   855
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Enabled         =   0   'False
      Height          =   400
      Left            =   2970
      TabIndex        =   3
      Top             =   2620
      Width           =   855
   End
   Begin VB.CommandButton cmdGoto 
      Caption         =   "Go To "
      Enabled         =   0   'False
      Height          =   400
      Left            =   1015
      TabIndex        =   2
      Top             =   2620
      Width           =   855
   End
   Begin VB.CommandButton cmdInsert 
      Caption         =   "Add..."
      Height          =   400
      Left            =   40
      TabIndex        =   1
      Top             =   2620
      Width           =   855
   End
   Begin VB.Label lblNoProp 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "There are no properties inserted into this document."
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   960
      TabIndex        =   6
      Top             =   600
      Visible         =   0   'False
      Width           =   3015
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmPropertyMgr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim moAppl As Object
Dim miApplType As Integer
Dim msFileName As String
Dim maBookMarkNames() As String
Dim miBookMarkArrayCounter As Integer
Dim msSummaryInfo As String
Dim lParentHWND As Long
Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdDelete_Click()
    Dim vbResult As VbMsgBoxResult
    Dim strName As String
    Dim sPropName As String
    
    On Error GoTo errHandler
    
    'high light the property before deleting
    Call GoToBkPosition
    
    'get the property's name out of the current row of the grid:
    'grdPropDisp.Col = 0
    strName = maBookMarkNames(grdPropDisp.Row)
    sPropName = GetCurrentPropertyName()
    
    vbResult = MsgBox(LoadResString(PROMPT_DELETE_PROPERTY) & sPropName & LoadResString(IDM_QUESTION_MARK), vbYesNo, "Confirm Delete")
    If vbResult = vbYes Then
        Select Case miApplType
            Case APPL_WORD
                Call WordDeleteProperty(moAppl, strName)
            Case APPL_EXCEL
                Call ExcelDeleteProperty(moAppl, strName)
            Case Else
                'no bookmarks in powerpoint
                GoTo Done
        End Select
    End If
    Call EnableButtons
Done:
    Exit Sub
errHandler:
    MsgBox Err.Description, vbCritical, LoadResString(DLG_ERR_FRMPROPMGR_CMDINSERT)
End Sub
Private Sub cmdReplace_Click()
    Dim vbResult As VbMsgBoxResult
    Dim igrdRow As Integer
    Dim sBookMarkName As String
    Dim sFileName As String
    Dim sPropertyName As String
    
    On Error GoTo errHandler
    'high light the propertbefore replace
    Call GoToBkPosition
    
    sPropertyName = GetCurrentPropertyName
    igrdRow = GetGridRowNum
    sBookMarkName = GetBookMarkName(igrdRow)
    sFileName = getFullName(moAppl, miApplType)
    Select Case miApplType
           Case APPL_WORD
                Call InsertMezzProperties(sFileName, miApplType, moAppl, IDMReplace)
           Case APPL_EXCEL
                Call InsertMezzProperties(sFileName, miApplType, moAppl, IDMReplace)
            Case APPL_POWERPOINT
                'no bookmarks in powerpoint
    End Select
    'Call EnableButtons
    Exit Sub
    
errHandler:
    MsgBox Err.Description, vbCritical, LoadResString(DLG_ERR_FRMPROPMGR_CMDREPLACE_CLICK)
End Sub
Private Sub cmdGoto_Click()
    Dim lRowNumber As Long
    On Error GoTo errHandler
    
    lRowNumber = grdPropDisp.Row
    grdPropDisp.Col = 0
    Call GotoProperty(maBookMarkNames(lRowNumber), miApplType, moAppl)
    'Call EnableButtons
    Exit Sub
    
errHandler:
    MsgBox Err.Description, vbCritical, LoadResString(DLG_ERR_FRMPROPMGR_CMDGOTO_CLICK)
End Sub
Private Sub cmdInsert_Click()
    Call InsertMezzProperties(msFileName, miApplType, moAppl, IDMInsert)
    'Call EnableButtons
End Sub
Public Property Let AppObject(ByVal oA As Object)
    Set moAppl = oA
End Property
Public Property Let ApplType(ByVal iAT As Integer)
    miApplType = iAT
End Property
Private Sub Form_Load()
   'Dim lParentHWND As Long
   Dim lResult As Long
   Dim bResult As Boolean
   Dim sString As String
   
   
   miBookMarkArrayCounter = 1
   ReDim maBookMarkNames(10)
   'This sets up the Property Mangager to subclass the native application
   
   lParentHWND = GetForegroundWindow()
   gHW = lParentHWND
   Hook
   bResult = SetWindowPos(frmPropertyMgr.hWnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
   gbPropMgrStatus = True
   Call EnableButtons
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Dim dl&
    Erase maBookMarkNames
    miBookMarkArrayCounter = 1
    gbPropMgrStatus = False
    dl& = SetFocusAPI(gpWnd)
    Unhook
End Sub
Private Sub grdPropDisp_DblClick()
     Call GoToBkPosition
End Sub
Public Property Let FileName(ByVal vNewValue As Variant)
    msFileName = vNewValue
End Property
Public Sub UpdatePropGrid(sBookMarkName As String, sPropName As String, sBookMarkValue As Variant)
  'Funtion: Add items to the display
  'Inputs: sBookMarkName string variable holding Bookmark name to add to grid
  '         sBookMarkValue - variant value that represents the value of the property being inserted
    '           iNumRows - current row count of the grid
    'Constraints: None
    'Dependencies: None
    Dim sSummaryInfo As String
    Dim iRow As Integer
    
    On Error GoTo errHandler:
   
    frmPropertyMgr.grdPropDisp.FormatString = LoadResString(IDM_GRID_HEADER)
    'This if-then deals handles going back to the frmPropertyManager and we now have things
    'to display in the grid
    If grdPropDisp.Visible = False Then
        grdPropDisp.Visible = True
        lblNoProp.Visible = False
    End If
    Select Case miApplType
        Case APPL_WORD
            sSummaryInfo = WordSummaryInfo(sBookMarkName)
        Case APPL_EXCEL
            sSummaryInfo = ExcelSummaryInfo(sBookMarkName)
        Case Else
            Exit Sub
    End Select
    grdPropDisp.AddItem sPropName & vbTab & sBookMarkValue _
            & vbTab & sSummaryInfo
    
'    If sBookMarkValue = LoadResString(TXT_NO_THIS_VALUE) Then
'        iRow = grdPropDisp.Row
'        grdPropDisp.Row = grdPropDisp.Rows - 1
'        grdPropDisp.Col = 1
'        grdPropDisp.CellForeColor = vbRed
'    End If
    
    Call Add_to_Array(sBookMarkName)
    Call EnableButtons
    Exit Sub

errHandler:
    MsgBox Err.Description, vbCritical, LoadResString(MSG_UPDATE_PROP_GRID)
End Sub
Private Sub Add_to_Array(sBookMarkName As String)
    '------------------------------------------------------------
    'Purpose: Adds an item to the array holding the current documents bookmark names
    'Inputs: sBookMark - a string representing the bookmark naem
    'Outputs: None
    'Assumptions: the element number of the input string equals the row number in the
    '             grdPropDisp.
    'Constraints
    'Copyright © 1998 FileNET Corporation
    '------------------------------------------------------------
    Dim iUBound As Long
    Dim iTest As Long
    iUBound = UBound(maBookMarkNames)
    'First check to see it the array is large enough
    If (miBookMarkArrayCounter + 1) >= iUBound Then
        iTest = iUBound + 5
        'we need to enlarge the array
        ReDim Preserve maBookMarkNames(iTest) As String
    End If
    maBookMarkNames(miBookMarkArrayCounter) = sBookMarkName
    miBookMarkArrayCounter = miBookMarkArrayCounter + 1
End Sub
Private Function WordSummaryInfo(sBookMarkName As String) As String
    '------------------------------------------------------------
    'Purpose:Creates a summary string from the information contained in the BookMark
    'Inputs: sBookMark - string representing a Word BookMark
    'Outputs: function returns a string representing summary information about the
    'Assumptions:None
    'Constraints: None
    'Copyright © 1998 FileNET Corporation
    '------------------------------------------------------------
    Dim oBKMark As Object
    Dim oRange As Object
    Dim oCurrentRange As Object
    Dim sPosChar As String
    Dim sSectionNum As String
    Dim sPageNum As String
    Dim sSummaryInfo As String
    Dim lResult As Long
    Dim vSec As Variant
    Dim vPage As Variant
    Dim viewType As Long
    On Error GoTo errHandler
    
    'We do some housekeeping to keep track of customers state
    Set oCurrentRange = moAppl.Selection.Range
    viewType = moAppl.ActiveWindow.ActivePane.View.Type
    
    Set oBKMark = moAppl.ActiveDocument.Bookmarks.Item(sBookMarkName)
    If Not IsObject(oBKMark) Then
        Exit Function
    End If
    'First determine where the property is Header, Footer Body,
    Set oRange = oBKMark.Range
    
    sPosChar = CurrentPosChar(oRange)
    
    lResult = oRange.Information(wdHeaderFooterType)
    'In the case of HEaders and Footers we need to open them to get real information
    'from the .Information call
    If (lResult >= 0) Then
        oBKMark.Select
        With moAppl.ActiveWindow.View
            .Type = wdPageView
            .SeekView = GetViewType(lResult)
        End With
        sSectionNum = Str(moAppl.Selection.Information(wdActiveEndSectionNumber))
        sPageNum = Str(moAppl.Selection.Information(wdActiveEndPageNumber))
    Else
        sSectionNum = Str(oRange.Information(wdActiveEndSectionNumber))
        sPageNum = Str(oRange.Information(wdActiveEndPageNumber))
    
    End If
    
    'Next Determine the Section and page the property it is located in
    lResult = StrComp(LoadResString(IDM_PROP_BODY), sPosChar)
    If lResult = 0 Then
        sSummaryInfo = sPosChar & LoadResString(MSG_ON_PAGE) & sPageNum
    Else
        sSummaryInfo = sPosChar & LoadResString(MSG_IN_SECTION) & sSectionNum & LoadResString(MSG_ON_PAGE) & sPageNum
    End If
    
    msSummaryInfo = sSummaryInfo
    WordSummaryInfo = sSummaryInfo
    
    'Get the users state back like it was
    oCurrentRange.Select
    moAppl.ActiveWindow.ActivePane.View.Type = viewType

    Set oRange = Nothing
    Set oBKMark = Nothing
    Set oCurrentRange = Nothing
    
    Exit Function
    
errHandler:
        MsgBox Err.Description, vbCritical, LoadResString(DLG_ERR_FRMPROPMGR_WORD_SUMMARY)
End Function
Private Function ExcelSummaryInfo(sBookMarkName As String) As String
    '------------------------------------------------------------
    'Purpose:Creates a summary string from the information contained in the BookMark
    'Inputs: sBookMark - string representing a Word BookMark
    'Outputs: function returns a string representing summary information about the
    'Assumptions:None
    'Constraints: None
    'Copyright © 1998 FileNET Corporation
    '------------------------------------------------------------
    Dim oName As Object
    Dim oWorkSheet As Object
    Dim sPosChar As String
    Dim sSectionNum As String
    Dim sPageNum As String
    Dim sSummaryInfo As String
    Dim lResult As Long
    Dim vSec As Variant
    Dim vPage As Variant
    Dim viewType As Long
    
    On Error GoTo errHandler
    Set oName = moAppl.ActiveWorkbook.Names.Item(sBookMarkName)
    If Not IsObject(oName) Then
        Exit Function
    End If
    moAppl.Goto Reference:=sBookMarkName
    sSummaryInfo = moAppl.ActiveSheet.Name & LoadResString(MSG_IN_CELL) & moAppl.ActiveCell.Address
    msSummaryInfo = sSummaryInfo
    ExcelSummaryInfo = sSummaryInfo
    Set oName = Nothing
    
    Exit Function
errHandler:
        MsgBox Err.Description, vbCritical, LoadResString(DLG_ERR_FRMPROPMGR_EXCEL_SUMMARY)
End Function
Public Function CurrentPosChar(oRange As Object) As String
    '------------------------------------------------------------
    'Purpose:Used for the Word Integration this function gets _
             is used to return a string describing the location _
             of the property in the Word Document
    'Inputs:   oRange - Range of the current Bookmark
    'Outputs:
    'Assumptions: To get the correct information we need to open the document
    'Constraints
    'Copyright © 1998 FileNET Corporation
    '------------------------------------------------------------
        Select Case oRange.Information(wdHeaderFooterType)
            Case 0
                CurrentPosChar = LoadResString(IDM_PROP_HEADER_EVEN_PAGE)
            Case 1
                CurrentPosChar = LoadResString(IDM_PROP_HEADER_ODD_PAGE)
            Case 2
                CurrentPosChar = LoadResString(IDM_PROP_FOOTER_EVEN_PAGE)
            Case 3
                CurrentPosChar = LoadResString(IDM_PROP_FOOTER_ODD_PAGE)
            Case 4
                CurrentPosChar = LoadResString(IDM_PROP_HEADER_FIRST_PAGE)
            Case 5
                CurrentPosChar = LoadResString(IDM_PROP_FOOTER_FIRST_PAGE)
            Case Else
                CurrentPosChar = LoadResString(IDM_PROP_BODY)
        End Select
End Function
Public Function GetViewType(lType As Long) As Variant
    '------------------------------------------------------------
    'Purpose:Returns the view currently open
    'Inputs:
    'Outputs:
    'Assumptions:called by WordSummaryInfo
    'Constraints
    'Copyright © 1998 FileNET Corporation
    '------------------------------------------------------------
    Select Case lType
           Case 0
                GetViewType = wdSeekEvenPagesHeader
           Case 1
                GetViewType = wdSeekPrimaryHeader
           Case 2
                GetViewType = wdSeekEvenPagesFooter
           Case 3
                GetViewType = wdSeekPrimaryFooter
           Case 4
                GetViewType = wdSeekFirstPageHeader
           Case 5
                GetViewType = wdSeekFirstPageFooter
           Case Else
                GetViewType = wdSeekMainDocument
    End Select
End Function
Public Sub RemoveGridEntry(sBookMarkName As String)
    '------------------------------------------------------------
    'Purpose: Removes the item from grdPropDisp and the maBookMarkNames Array.
    '          Updates all of the required data members to support new grid and array size
    'Inputs:    sBookMarkName - string representing the bookmark being removed
    'Outputs:
    'Assumptions: Can only be called when the user wants to remove a bookmark of excel
    '               reference.
    'Constraints:
    'Copyright © 1998 FileNET Corporation
    '------------------------------------------------------------
    Dim iLoopCount As Integer
    Dim iRowNum As Integer
    
    On Error GoTo errHandler
    
    'First thing is to update the maBookMarkArray
    For iLoopCount = grdPropDisp.Row To (UBound(maBookMarkNames) - 1)
         maBookMarkNames(iLoopCount) = maBookMarkNames(iLoopCount + 1)
    Next iLoopCount
    
    iRowNum = grdPropDisp.Row
    
    If miBookMarkArrayCounter - 1 > 1 Then
    
        grdPropDisp.RemoveItem (grdPropDisp.Row)
    Else
        With grdPropDisp
            .Col = 0
            .Text = LoadResString(IDM_SPACE)
            .Col = 1
            .Text = LoadResString(IDM_SPACE)
            .Col = 2
            .Text = LoadResString(IDM_SPACE)
            .Visible = False
            grdPropDisp.Rows = 1
        End With
        lblNoProp.Visible = True
    End If
    miBookMarkArrayCounter = miBookMarkArrayCounter - 1
    Exit Sub
errHandler:
    MsgBox Err.Description, vbCritical, LoadResString(DLG_ERR_FRMPROPMGR_EXCEL_SUMMARY)
End Sub
Public Function GetBookMarkName(iRowNum As Integer) As String
    GetBookMarkName = maBookMarkNames(iRowNum)
End Function
Public Function GetGridRowNum() As Integer
    GetGridRowNum = grdPropDisp.Row
    
End Function
Public Function GetCurrentPropertyName() As String
    With grdPropDisp
              .Col = 0
            GetCurrentPropertyName = .Text
    End With
End Function
Public Function GetCurrentPropertyValue() As String
    With grdPropDisp
         .Col = 1
         GetCurrentPropertyValue = .Text
    End With
End Function
Public Function GetCurrentPropertyLocation() As String
    With grdPropDisp
          .Col = 2
          GetCurrentPropertyLocation = .Text
    End With
End Function
Private Sub grdPropDisp_SelChange()
    cmdReplace.Enabled = True
    cmdDelete.Enabled = True
    cmdGoto.Enabled = True
End Sub
Private Sub GoToBkPosition()
    Dim lRowNumber As Long
    lRowNumber = grdPropDisp.Row
    grdPropDisp.Col = 0
    Call GotoProperty(maBookMarkNames(lRowNumber), miApplType, moAppl)
End Sub
Sub EnableButtons()
  If grdPropDisp.Rows > 1 Then
     cmdReplace.Enabled = True
     cmdDelete.Enabled = True
     cmdGoto.Enabled = True
   Else
     cmdReplace.Enabled = False
     cmdDelete.Enabled = False
     cmdGoto.Enabled = False
   End If
End Sub


