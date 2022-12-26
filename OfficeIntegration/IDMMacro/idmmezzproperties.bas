Attribute VB_Name = "IdmMezzProperties"
Option Explicit
Public oSelection As Selection
Sub ShowPropertyMgr(oAppl As Object, iApplType As Integer)
    '------------------------------------------------------------
    'Purpose:Sets up and exposes the frmPropertyManger
    'Inputs: oappl - Calling Application's application object
    '        iApplType - Calling applications name
    'Outputs:
    'Assumptions:
    'Constraints
    'Copyright © 1998 FileNET Corporation
    '------------------------------------------------------------
    Dim lResult As Long
    Dim sFileName As String
    Dim iNumBookMarks As Integer
    Dim lResponse As Long
    Dim bLogin As Boolean
    Dim oDoc As IDMObjects.Document
    
    On Error GoTo errHandler
    
    sFileName = getFullName(oAppl, iApplType)
    
    'to check login user name
    Set oDoc = GetDocObject(sFileName, bLogin)
    If bLogin = False Then GoTo Done

    If GetDocStatus(sFileName) <> DocCheckedout Then
    
        'We need to get the user to add doc to IDMDS so we do a FileNET Save and
        'then get things back like itshould be
        
        lResponse = MsgBox(LoadResString(MSG_MUST_ADD_FIRST), vbExclamation + vbYesNo + vbApplicationModal, LoadResString(MSG_SHOW_PROPERTY_MGR))
        
        If lResponse = vbYes Then
            lResult = fileNetSave(oAppl, iApplType, sFileName)
            
            'check to see if the save went ok
            If lResult <> CIDMOk Then
                Exit Sub
            End If
            'We need to have a tracked file object for later
            sFileName = getFullName(oAppl, iApplType)
        Else
            Exit Sub
        End If
    End If
    'pseudocode
    '1 cruise through all of the bookmarks in the document, get the property
    '  name and value.  If it's an SDM tag, convert the tag to IDM using Eric's
    '  conversion routine
    '2 fill the grid of the Property Manager form with the bookmarks you just nabbed.
    With frmPropertyMgr
            .FileName = getFullName(oAppl, iApplType)
            .ApplType = iApplType
            .AppObject = oAppl
    End With
    Select Case iApplType
           Case APPL_WORD
                Set oSelection = oAppl.Selection
                Load frmPropertyMgr
                iNumBookMarks = WordGetProperties(oAppl, iApplType, bLogin)
                If bLogin = False Then
                    GoTo Done
                End If
           Case APPL_EXCEL
                Load frmPropertyMgr
                iNumBookMarks = ExcelGetProperties(oAppl, iApplType, bLogin)
                If bLogin = False Then
                   GoTo Done
                End If
           Case APPL_POWERPOINT
                'no bookmarks in powerpoint
                GoTo Done
    End Select
    With frmPropertyMgr
    'this code deals with the grid being made visible based on the fact
    'that a inserted property exists.
            If iNumBookMarks = 0 Then
                .lblNoProp.Visible = True
                .grdPropDisp.Visible = False
            Else
                .lblNoProp.Visible = False
                .grdPropDisp.Visible = True
            End If
    End With
    frmPropertyMgr.Show vbModeless
Done:
    Exit Sub
errHandler:
    MsgBox Err.Description, vbCritical, LoadResString(DLG_ERR_SHOW_PROP_MGR)
End Sub
Public Function WordInsertProperty(strBkMrkName As String, strBkMrkValue As String, oAppl As Object, iApplType As Integer) As String
    'Parameters:
    '
    'strBkMrkName:
    '  SDM it was strBkMrkName = "MEZZ_" & property, ie Property Name
    '  IDM it will be "IDM" & oprop.PropertyDescription from Mezzanine
    '
    'strAttrValue:
    '  the property's value
    Dim oBMark As Bookmark
    Dim iPropCnt As Integer
    Dim iPropLength As Integer
    
    On Error GoTo errHandler
    
    iPropCnt = GetMaxBkNumber(strBkMrkName, oAppl, iApplType)
    'build the bookmark's name
    strBkMrkName = LoadResString(IDM_TAG) & LoadResString(IDM_SEPARATOR) & strBkMrkName & LoadResString(IDM_SEPARATOR) & iPropCnt
    
    'Deals with Null values from property
    If IsNull(strBkMrkValue) Then
        strBkMrkValue = " "
    End If
    
   oSelection.TypeText Text:=strBkMrkValue
   iPropLength = Len(strBkMrkValue)

   oSelection.MoveLeft Unit:=wdCharacter, Count:=iPropLength, Extend:=wdExtend
   'If strBkMrkValue = LoadResString(TXT_NO_THIS_VALUE) Then
   '   oSelection.Font.ColorIndex = wdRed
   'End If
    oAppl.ActiveDocument.Bookmarks.Add Range:=oSelection.Range, Name:=strBkMrkName
    oSelection.MoveRight Unit:=wdCharacter, Count:=1 ', Extend:=wdExtend  '???
    ' Insert space so properties do not get combined
    oSelection.TypeText Text:=LoadResString(IDM_SPACE)
    'oSelection.Font.ColorIndex = wdBlack
    WordInsertProperty = strBkMrkName
    
    Exit Function
    
errHandler:
    MsgBox Err.Description, vbCritical, LoadResString(DLG_ERR_WORD_INSERT_PROP)
End Function

Public Function ExcelInsertProperty(strBkMrkName As String, strBkMrkValue As String, oAppl As Object, iApplType As Integer) As String
    'all this routine has to do is insert the bookmark and value into Excel.
    'Parameters:
    'strBkMrkName:
    '  SDM it was bmSearchName = "MEZZ_" & property, ie Property Name
    '  IDM it will be "IDM" & oprop.PropertyDescription from Mezzanine
    '
    'strAttrValue:
    '  the property's value
    Dim strName As String
    Dim namTemp As Name
    Dim strCurrentSheet As String
    Dim strCurrentCell As String
    Dim strMaxNum As String
    Dim strCellName As String
    Dim strPropNum As String
    Dim strSheetName As String
    Dim iPos As Integer
    Dim iProperty As Integer
    Dim iCount As Integer
    Dim iPropUniqID As Integer
    Dim iMax As Integer
    Dim sSDMTag As String
    Dim sSDMNTag As String
    
    On Error GoTo errHandler
    
    'sheet and cell that were active when macro was called
    strCurrentSheet = oAppl.ActiveSheet.Name
    strCurrentCell = oAppl.ActiveCell.Address
    
    'Get number of Names
    iCount = oAppl.ActiveWorkbook.Names.Count
    sSDMTag = LoadResString(SDM_TAG)
    sSDMNTag = LoadResString(SDM_N_TAG)
    'If there are no Names, then you don't have to increment the propertyID
    If (iCount = 0) Then
        GoTo loopExit
    End If
    
    For Each namTemp In oAppl.ActiveWorkbook.Names
        strName = namTemp.Name
        'See if the name is a Mezz Property
        If UCase(Left$(strName, 4)) = sSDMTag Then
            oAppl.GoTo Reference:=strName
            strCellName = oAppl.ActiveCell.Address
            strSheetName = oAppl.ActiveSheet.Name
            'See if it's the same cell
            If (0 = StrComp(strCurrentCell, strCellName)) Then
                If (0 = StrComp(strCurrentSheet, strSheetName)) Then
                    'if it's the same cell, delete it so you can insert the new one.
                    namTemp.Delete
                    GoTo loopExit:
                End If
            End If
        End If
    Next
loopExit:
    'Return active cell
    oAppl.ActiveWorkbook.Worksheets(strCurrentSheet).Select
    oAppl.ActiveSheet.Range(strCurrentCell).Select
    
    'get the last index number
    iMax = 0
    iCount = oAppl.ActiveWorkbook.Names.Count
    If (iCount > 0) Then
        For Each namTemp In oAppl.ActiveWorkbook.Names
            strName = namTemp.Name
                'if the range name from above contains the property number, OR
                'if property already exists AND the beginning of the range name
                'matches the global range name constant:  then
                '   pos = location where uniqueness number (N) starts, and if
                '   pos <> 0 then
            
            'If '(InStr(strName, LoadResString(LIBRARY_ST)) > 0) And
            If UCase(Left$(strName, 4)) = sSDMTag Then
                iPos = InStr(strName, sSDMNTag)
                If iPos <> 0 Then
                    iPropUniqID = Val(Mid(strName, iPos + 1))
                    If (iPropUniqID > iMax) Then
                        iMax = iPropUniqID
                    End If
                End If
            End If
        Next
    End If
    iMax = GetMaxBkNumber(strBkMrkName, oAppl, iApplType)
    strMaxNum = LTrim$(Str$(iMax))
    strBkMrkName = LoadResString(IDM_TAG) & LoadResString(IDM_SEPARATOR) & strBkMrkName & LoadResString(IDM_SEPARATOR) & strMaxNum
    oAppl.ActiveWorkbook.Names.Add Name:=strBkMrkName, RefersToR1C1:=oAppl.ActiveCell
        
    If IsNull(strBkMrkValue) Then
        strBkMrkValue = " "
    End If
    
    'oAppl.ActiveCell.FormulaR1C1 = strBkMrkValue
    oAppl.ActiveCell.Characters.Text = strBkMrkValue
    'oAppl.ActiveCel.Text = strBkMrkValue
    'If strBkMrkValue = LoadResString(TXT_NO_THIS_VALUE) Then
    '  oAppl.ActiveCell.Font.ColorIndex = 3
    'End If
    ExcelInsertProperty = strBkMrkName
Done:
    Exit Function
    
errHandler:
    MsgBox Err.Description, vbCritical, LoadResString(DLG_ERR_EXCEL_INSERT_PROP)
End Function

Sub UpdateMezzProperties(oAppl As Object, iApplType As Integer, sPreferenceName As String)
    Dim iPropCount As Integer
    Dim vPrefVal As Variant
    
    iPropCount = IDMPropCount(oAppl, iApplType) + SDMPropCount(oAppl, iApplType)
    If iPropCount = 0 Then
       Exit Sub
    Else
       If sPreferenceName = LoadResString(STR_UPDATE_PROPERTIES) Then
          Call DoSDMtoIDMConversion(oAppl, iApplType)
          Call DoUpdateMezzProperties(oAppl, iApplType)
          Exit Sub
       End If
            vPrefVal = GetPreferenceValue(sPreferenceName, iApplType)
            Select Case vPrefVal
                   Case LoadResString(TXT_NEVER) '"NEVER"
                        'SDM strings will not be converted
                        'will not update MEZZ properties
                   Case LoadResString(TXT_ALWAYS) '"ALWAYS"
                        Call DoSDMtoIDMConversion(oAppl, iApplType)
                        Call DoUpdateMezzProperties(oAppl, iApplType)
                   Case LoadResString(TXT_PROMPTUSER) '"PROMPTUSER"
                        If MsgBox(LoadResString(MSG_DO_YOU_WANT_TO_UPDATE_PROPERTIES), vbInformation + vbYesNo, LoadResString(MNU_UPDATE_MEZZ_PROP)) = vbYes Then
                            Call DoSDMtoIDMConversion(oAppl, iApplType)
                            Call DoUpdateMezzProperties(oAppl, iApplType)
                        End If
           End Select
    End If
    End Sub
Sub DoSDMtoIDMConversion(oAppl As Object, iApplType As Integer)
       Dim iSDMPropCount As Integer
       Dim vPrefVal As Variant
       iSDMPropCount = SDMPropCount(oAppl, iApplType)
       If iSDMPropCount >= 1 Then
            vPrefVal = GetPreferenceValue(LoadResString(STR_UPDATE_ENBEDDED_PROP), iApplType)
            Select Case vPrefVal
                   Case LoadResString(TXT_NEVER) '"NEVER"
                        'SDM strings will not be converted
                   Case LoadResString(TXT_ALWAYS) '"ALWAYS"
                        Call SDMtoIDMConversion(oAppl, iApplType)
                   Case LoadResString(TXT_PROMPTUSER) '"PROMPTUSER"
                        If MsgBox(LoadResString(MSG_DO_YOU_WANT_TO_CONVERT_SDM_TO_IDM), vbInformation + vbYesNo, LoadResString(MSG_CONVERSION)) = vbYes Then
                            Call SDMtoIDMConversion(oAppl, iApplType)
                        End If
           End Select
       End If
End Sub
Sub DoUpdateMezzProperties(oAppl As Object, iApplType As Integer)
    Select Case iApplType
        Case APPL_WORD
            Call WordUpdateProperties(oAppl, iApplType)
        Case APPL_EXCEL
            Call ExcelUpdateProperties(oAppl, iApplType)
        Case APPL_POWERPOINT
            'no bookmarks in powerpoint
            Exit Sub
    End Select
End Sub
Sub ExcelUpdateProperties(oAppl As Object, iApplType As Integer, Optional oDoc As IDMObjects.Document)
Attribute ExcelUpdateProperties.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim strField As String
    Dim namTemp As Name
    Dim sPropName As String
    Dim sFileName As String
    Dim iLoc As Integer
    Dim sIDMTag As String
    Dim sPropValue As String
    Dim blogon As Boolean
    
    On Error GoTo errHandler
    
    sIDMTag = LoadResString(IDM_TAG) & LoadResString(IDM_SEPARATOR)
    For Each namTemp In oAppl.ActiveWorkbook.Names
        strField = namTemp.Name
        iLoc = InStr(strField, sIDMTag)
        If iLoc > 0 Then
            oAppl.GoTo Reference:=strField
            'read the property value from the bookmark
            sPropName = DeterminePropName(strField)
            sFileName = getFullName(oAppl, iApplType)
            sPropValue = GetPropValue(sPropName, sFileName, blogon)
            If blogon = False Then Exit Sub  'user clicked cancel on logon dialog
            'oAppl.ActiveCell.FormulaR1C1 = sPropValue
            oAppl.ActiveCell.Characters.Text = sPropValue
            'If sPropValue = LoadResString(TXT_NO_THIS_VALUE) Then
            '   oAppl.ActiveCell.Font.ColorIndex = 3
            'End If
        End If
    Next
Exit Sub
errHandler:
    If Err.Number = 1004 Then
           oAppl.ActiveCell.FormulaR1C1 = " "
           oAppl.ActiveCell.Characters.Text = sPropValue
           Resume Next
    End If
     
    MsgBox Err.Description, vbCritical, LoadResString(DLG_ERR_EXCEL_UPDATE_PROP)
End Sub
Sub WordUpdateProperties(oAppl As Object, iApplType As Integer)
    Dim oBMark As Bookmark
    Dim strBkMarkName As String
    Dim iBmCnt As Integer
    Dim iPrimaryCnt As Integer
    Dim oRange As Object
    Dim iLoc As Integer
    Dim iNewLength As Integer
    Dim iCount As Integer
    Dim iStrt As Long
    Dim iEnd As Long
    Dim iCurLength As Integer
    Dim aBMList() As String
    Dim sFileName As String
    Dim sPropName As String
    Dim sNewPropValue As String
    Dim oErrorMgr As ErrorManager
    Dim sIDMTag As String
    Dim blogon As Boolean
    
    On Error GoTo errHandler
    'First setup error handling
    Set oErrorMgr = CreateObject(CREATE_ERR_MGR)
    If oErrorMgr Is Nothing Then
        Err.Raise 1, LoadResString(DLG_ERR_WORD_UPDATE_PROP), LoadResString(MSG_CANNOT_CREATE_ERRMGR)
    End If
    
    sFileName = getFullName(oAppl, iApplType)
    
    iBmCnt = oAppl.ActiveDocument.Bookmarks.Count
    ' if there are not any book marks skip out and enjoy the sun
    If iBmCnt < 1 Then Exit Sub
    
    ReDim aBMList(iBmCnt) As String
    
    If oAppl.ActiveWindow.View.SplitSpecial <> wdPaneNone Then
        oAppl.ActiveWindow.Panes(2).Close
    End If
    If oAppl.ActiveWindow.ActivePane.View.Type <> wdPageView Then
        oAppl.ActiveWindow.ActivePane.View.Type = wdPageView
    End If
    iPrimaryCnt = 0
    
   'Get a list of bookmarks first
   sIDMTag = LoadResString(IDM_TAG) & LoadResString(IDM_SEPARATOR)
   
    For Each oBMark In oAppl.ActiveDocument.Bookmarks
        strBkMarkName = oBMark.Name
            iLoc = InStr(strBkMarkName, sIDMTag)
        If iLoc > 0 Then
            aBMList(iPrimaryCnt) = strBkMarkName
            iPrimaryCnt = iPrimaryCnt + 1
        End If
    Next oBMark
            
     'Now we can update the properties
    For iCount = 0 To iPrimaryCnt - 1      'Update each bookmark
        Set oBMark = oAppl.ActiveDocument.Bookmarks.Item(aBMList(iCount))
        strBkMarkName = oBMark.Name
        'if mezz header bookmark then refresh bookmark data
        iLoc = InStr(strBkMarkName, sIDMTag)
        If iLoc > 0 Then
            Set oRange = oBMark.Range
            iStrt = oRange.Start
            iEnd = oRange.End
            sPropName = DeterminePropName(strBkMarkName)
            sNewPropValue = NewText(GetPropValue(sPropName, sFileName, blogon))
            If blogon = False Then Exit Sub  'user clicked cancel on logon dialog
            iCurLength = Len(Trim(NewText(oRange.Text)))
            'If iCurLength = 0 Then GoTo GoNext
            If Trim(NewText(oRange.Text)) = Trim(sNewPropValue) Then GoTo GoNext
            iNewLength = Len(Trim(sNewPropValue))
            oRange.Text = Trim(sNewPropValue)
        
            If (iCurLength > iNewLength) Then
                oRange.SetRange Start:=iStrt, End:=iEnd - (iCurLength - iNewLength)
            End If
            
            If (iCurLength < iNewLength) Then
                oRange.SetRange Start:=iStrt, End:=iEnd + (iNewLength - iCurLength)
            End If
            'Work around for deleting the Bookmark due to the change of the oRange.Text
            
            'If sNewPropValue = LoadResString(TXT_NO_THIS_VALUE) Then
            '   oRange.Font.ColorIndex = wdRed
            'End If
            'If Not (oAppl.ActiveDocument.Bookmarks.Exists(strBkMarkName)) Then
                oAppl.ActiveDocument.Bookmarks.Add Name:=strBkMarkName, Range:=oRange
            'End If
        End If
GoNext:
    Next iCount
    Exit Sub
errHandler:
    If Err.Number <> 0 And Err.Number <> -2147216381 Then
        MsgBox Err.Description, vbCritical, LoadResString(DLG_ERR_WORD_UPDATE_PROP)
    Else
        If oErrorMgr Is Nothing Then
            MsgBox MSG_ERROR_WITHOUT_ERRMGR, vbCritical, LoadResString(DLG_ERR_WORD_UPDATE_PROP)
            Exit Sub
        Else
            If oErrorMgr.Errors.Count > 0 Then
                oErrorMgr.ShowErrorDialog
            End If
        End If
    End If
End Sub
Public Function WordGetProperties(oAppl As Object, iApplType As Integer, bLogin As Boolean) As Integer
    'Assumptions:
    '   frmPropertyMgr is already LOADED when this function gets called
    '
    Dim oBMark As Bookmark
    Dim oDoc As IDMObjects.Document
    Dim strBkMarkName As String
    Dim strText As String
    Dim iBmCnt As Integer
    Dim iPrimaryCnt As Integer
    Dim iRowCnt As Integer
    Dim oRange As Object
    Dim sPropName As String
    Dim sPropLabel As String
    Dim sFileName As String
    Dim iLoc As Long
    Dim sIDMTag As String
    
    On Error GoTo errHandler
    
    sFileName = getFullName(oAppl, iApplType)   '08/31/99 moved to here because we need the bLogin
    Set oDoc = GetDocObject(sFileName, bLogin)

    If oDoc Is Nothing Then
        Exit Function
    End If
    
    iBmCnt = oAppl.ActiveDocument.Bookmarks.Count
    
    If iBmCnt < 1 Then
        iPrimaryCnt = 0
        GoTo Done
    End If
    
    iPrimaryCnt = 0
    iRowCnt = 1
    sIDMTag = LoadResString(IDM_TAG) & LoadResString(IDM_SEPARATOR)
   
    'do some funky stuff to the Word environment, inherited this from SDM.
    'not quite sure what it does, Eric can you document this sometime?
    If oAppl.ActiveWindow.View.SplitSpecial <> wdPaneNone Then
        oAppl.ActiveWindow.Panes(2).Close
    End If
    If oAppl.ActiveWindow.ActivePane.View.Type <> wdPageView Then
        oAppl.ActiveWindow.ActivePane.View.Type = wdPageView
    End If
    
    For Each oBMark In oAppl.ActiveDocument.Bookmarks
        strBkMarkName = oBMark.Name
        iLoc = InStr(strBkMarkName, sIDMTag)
        If iLoc > 0 Then
            Set oRange = oBMark.Range
            strText = oRange.Text
            sPropName = DeterminePropName(strBkMarkName)
            sPropLabel = GetPropLabel(oDoc, sPropName)
            Call frmPropertyMgr.UpdatePropGrid(strBkMarkName, sPropLabel, strText)
            iPrimaryCnt = iPrimaryCnt + 1
        End If
    Next
Done:
    WordGetProperties = iPrimaryCnt
    Set oDoc = Nothing
    Exit Function

errHandler:
    MsgBox Err.Description, vbCritical, LoadResString(DLG_ERR_WORD_GET_PROP)
End Function
Public Function ExcelGetProperties(oAppl As Object, iApplType As Integer, bLogin As Boolean) As Integer
    Dim strFullName As String
    Dim strAttrValue As String
    Dim strField As String
    Dim namTemp As Name
    Dim iCount As Integer
    Dim strCurrentCell As String
    Dim strCurrentSheet As String
    Dim sPropName As String
    Dim sFileName As String
    Dim sPropLabel As String
    Dim iRowCnt As Integer
    Dim iLoc As Integer
    Dim oDoc As IDMObjects.Document
    Dim sIDMTag As String
    
    On Error GoTo errHandler
    
    'sheet and cell that were active when macro was called
    strCurrentSheet = oAppl.ActiveSheet.Name
    strCurrentCell = oAppl.ActiveCell.Address
    
    'Disable any user input while properties are being updated.
    oAppl.Interactive = False

    'Get Current Filename
    strFullName = oAppl.ActiveWorkbook.FullName
    
    sFileName = getFullName(oAppl, iApplType)    '08/31/99 moved to here from below
    Set oDoc = GetDocObject(sFileName, bLogin)
    'Check on Error message
    If oDoc Is Nothing Then
       Exit Function
    End If
    
    
    'Get number of Names
    iCount = oAppl.ActiveWorkbook.Names.Count     'Get number of defined names in the document

    'If there are no Names, then done
    If (iCount = 0) Then
        GoTo Done
    End If
    iCount = 0
    iRowCnt = 1
    sIDMTag = LoadResString(IDM_TAG) & LoadResString(IDM_SEPARATOR)
    
    'scan through all the Names in the active workbook
    For Each namTemp In oAppl.ActiveWorkbook.Names
        strField = namTemp.Name
        iLoc = InStr(strField, sIDMTag)
        If iLoc > 0 Then
            oAppl.GoTo Reference:=strField
            'read the property value from the bookmark
            strAttrValue = oAppl.ActiveCell.Value 'FormulaR1C1
            sPropName = DeterminePropName(strField)
            sPropLabel = GetPropLabel(oDoc, sPropName)
            Call frmPropertyMgr.UpdatePropGrid(strField, sPropLabel, strAttrValue)
            iCount = iCount + 1
        End If
    Next
    'Place cursor in cell that was active when Update was called.
    ExcelGetProperties = iCount
Done:
    'Return active cell
    oAppl.ActiveWorkbook.Worksheets(strCurrentSheet).Select
    oAppl.ActiveSheet.Range(strCurrentCell).Select
    
    'Enable user input
    oAppl.Interactive = True
    Set oDoc = Nothing
    
    Exit Function

errHandler:
    MsgBox Err.Description, vbCritical, LoadResString(DLG_ERR_EXCEL_GET_PROP)
End Function

Function WordDeleteProperty(oAppl As Object, strBkMarkName As String)
    Dim oErrorMgr As ErrorManager

    On Error GoTo errHandler
    Set oErrorMgr = CreateObject(CREATE_ERR_MGR)
    If oErrorMgr Is Nothing Then
        Err.Raise 1, LoadResString(MSG_DELETE_PROP_WORD), LoadResString(MSG_CANNOT_CREATE_ERRMGR)
    End If
    
    If oAppl.ActiveWindow.View.SplitSpecial <> wdPaneNone Then
        oAppl.ActiveWindow.Panes(2).Close
    End If
    If oAppl.ActiveWindow.ActivePane.View.Type <> wdPageView Then
        oAppl.ActiveWindow.ActivePane.View.Type = wdPageView
    End If
    'oAppl.ActiveDocument.Bookmarks(strBkMarkName).Select
    oAppl.ActiveDocument.Bookmarks(strBkMarkName).Delete
    oAppl.Selection.Delete
    frmPropertyMgr.RemoveGridEntry (strBkMarkName)
    
    oAppl.ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
    frmPropertyMgr.cmdDelete.Enabled = False
Done:
    Set oErrorMgr = Nothing
    Exit Function

errHandler:
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbCritical, LoadResString(DLG_ERR_WORD_DEL_PROP)
    Else
        If oErrorMgr Is Nothing Then
            MsgBox LoadResString(MSG_ERROR_WITHOUT_ERRMGR), vbCritical, LoadResString(DLG_ERR_WORD_DEL_PROP)
            GoTo Done
        Else
            If oErrorMgr.Errors.Count > 0 Then
                oErrorMgr.ShowErrorDialog
            End If
        End If
    End If
    Exit Function
End Function

Function GotoProperty(strBkMarkName As String, iApplType As Integer, oAppl As Object)
    '------------------------------------------------------------
    'Purpose:
    'Inputs:
    'Outputs:
    'Assumptions:
    'Constraints
    'Copyright © 1998 FileNET Corporation
    '------------------------------------------------------------
    On Error GoTo errHandler
    
    Select Case iApplType
        Case APPL_WORD
            Call WordGotoProperty(oAppl, iApplType, strBkMarkName)
            GoTo Done
        Case APPL_EXCEL
            oAppl.GoTo Reference:=strBkMarkName
            GoTo Done
        Case Else
            'no bookmarks in powerpoint
            Exit Function
    End Select
Done:
    frmPropertyMgr.cmdReplace.Enabled = True
    frmPropertyMgr.cmdDelete.Enabled = True
    Exit Function

errHandler:
    MsgBox Err.Description, vbCritical, LoadResString(DLG_ERR_GOTO_PROP)

End Function

Function ExcelDeleteProperty(oAppl As Object, strName As String)
    Dim namBkMark As Name
    
    On Error GoTo errHandler
    
    For Each namBkMark In oAppl.ActiveWorkbook.Names
        If namBkMark.Name = strName Then
            namBkMark.Delete
            oAppl.ActiveCell.Value = "" ' Delete
            frmPropertyMgr.RemoveGridEntry (strName)
        End If
    Next
    Exit Function

errHandler:
    MsgBox Err.Description, vbCritical, LoadResString(DLG_ERR_EXCEL_DEL_PROP)

End Function
Sub showProperties(oAppl As Object, iApplType As Integer)
    'Name
    '   idmShowProperties
    'Funciton
    '   The IDMmacro Class method to show the properties
    '   dialog.
    '   Get the pathname of the application object
    '   Call ShowProperitesDialog
    'Inputs
    '   oAppl           The Application Object
    '   iApplType       The Application Type Number
    'Outputs
    '   ?
    Dim strFileFullname As String
   
    On Error GoTo errHandler
    
    If oAppl Is Nothing Then
        'Bail right away if no documents are open!!
        If DocCount(oAppl, iApplType) = 0 Then
            GoTo Done
        End If
    End If
    
    strFileFullname = getFullName(oAppl, iApplType)
    
    Call ShowPropertiesDialog(oAppl, iApplType, strFileFullname)
    Exit Sub
errHandler:
    ' Check for Automation Error from TrackFile->GetObject when user hits Cancel on Logon dialog
    ' and don't display any errors, just finish.
    If Err.Number <> 0 And Err.Number <> -2147216381 Then
        MsgBox Err.Description, vbCritical, LoadResString(DLG_ERR_SHOWPROP)
    End If
    
Done:

End Sub

Function ShowPropertiesDialog(oAppl As Object, iApplType As Integer, strFullPath As String)
    'Name
    '   ShowPropertiesDialog
    'Function
    '   The method to show the properties dialog
    '   and save it if there are any changes.
    'Inputs
    '   strFullPath - The full path of the file
    'Outputs
    '   ?
    
    Dim oDoc As IDMObjects.Document
    Dim oErrorMgr As ErrorManager
    Dim sFileName As String
    Dim blogon As Boolean
    
    On Error GoTo errHandler
    Set oErrorMgr = CreateObject(CREATE_ERR_MGR)
    If oErrorMgr Is Nothing Then
        Err.Raise 1, OPTIONS_DIALOG_SHOW_PROPS, MSG_CANNOT_CREATE_ERRMGR
    End If
    If DocCount(oAppl, iApplType) = 0 Then
        MsgBox LoadResString(MSG_FILE_NOT_EXIST), vbCritical, LoadResString(DLG_ERR_SHOW_PROP)
        GoTo Done
    End If
    sFileName = getFullName(oAppl, iApplType)
    If GetDocStatus(sFileName) <> DocCheckedout Then
        MsgBox sFileName + LoadResString(MSG_NOT_CHECKED_OUT), vbInformation, LoadResString(DLG_ERR_SHOW_PROP)
        GoTo Done
    End If
  
    Set oDoc = GetDocObject(sFileName, blogon)

    If blogon = False Then
        Exit Function
    End If
    
    If oDoc Is Nothing Then
        Err.Raise 1, LoadResString(OPTIONS_DIALOG_SHOW_PROPS), LoadResString(MSG_NO_DOC)
    End If
    If oDoc.ShowPropertiesDialog = idmDialogExitOK Then
        If oDoc.GetState(idmDocModified) Then
            oDoc.Save
            'update mezz properites here
        End If
    End If
    Exit Function
errHandler:
    ' Check for Automation Error from TrackFile->GetObject when user hits Cancel on Logon dialog
    ' and don't display any errors, just finish.
    If Err.Number <> 0 And Err.Number <> -2147216381 Then
        MsgBox Err.Description, vbCritical, LoadResString(OPTIONS_DIALOG_SHOW_PROPS)
    Else
        If oErrorMgr Is Nothing Then
            MsgBox MSG_ERROR_WITHOUT_ERRMGR, vbCritical, LoadResString(OPTIONS_DIALOG_SHOW_PROPS)
            GoTo Done
        Else
            If oErrorMgr.Errors.Count > 0 Then
                oErrorMgr.ShowErrorDialog
            End If
        End If
    End If
    
Done:
    Set oErrorMgr = Nothing
End Function
Function InsertMezzProperties(ByVal sFileName As String, miApplType As Integer, moAppl As Object, eOperation As AppIntOperationEnum)
    Dim oDoc As IDMObjects.Document
    Dim oVer As IDMObjects.Version
    Dim oProp As IDMObjects.Property
    Dim iOp As idmOperation
    Dim sBookMarkName As String
    Dim sPropValue As String
    Dim bInRange As String
    Dim blogon As Boolean
    
    On Error GoTo errHandler
    'We need to verify the user is not trying to put a property in a location already
    'containing the property
    
    Select Case miApplType
        Case APPL_WORD
            bInRange = InBookmark(moAppl)
         Case APPL_EXCEL
            bInRange = InPropertyCell(moAppl)
        Case Else
            bInRange = False
        End Select
    If (bInRange = True) And (eOperation = IDMInsert) Then
        MsgBox LoadResString(MSG_INSERT_POINT_ALREADY_PROPERTY), vbOKOnly
        Exit Function
    End If
    Set oDoc = GetDocObject(sFileName, blogon)
    
    If blogon = False Then
        GoTo Done
    End If
    
    If oDoc Is Nothing Then
        Err.Raise 0, LoadResString(DLG_ERR_INSERT_PROP), LoadResString(MSG_NO_DOC)
    End If
    
    Set oVer = oDoc.Version
    If oVer Is Nothing Then
        Err.Raise 0, LoadResString(DLG_ERR_INSERT_PROP), LoadResString(MSG_CANNOT_GET_VERSION)
    End If
        
    If goCmnDlg Is Nothing Then
        Set goCmnDlg = CreateObject(CREATE_COMMON_DLG)
    End If
    
    With goCmnDlg
        .Title = LoadResString(MSG_INSERT_PROP)
        .hWnd = GetActiveWindow
        .OpenAsDefault = idmOpenAsCheckout
        .Options = idmSelectShowTrackedFiles + idmSelectHideOpenAsView
    End With
    
    Call goCmnDlg.SelectProperty(oDoc, oProp, iOp)
    
    If iOp <> idmOperationCancel Then
        
        If IsNull(oProp) Or IsNull(oProp.Value) Then
            sPropValue = "" ' LoadResString(TXT_NO_THIS_VALUE)
        Else
            If IsNumeric(oProp.Value) Then
               sPropValue = oProp.FormatValue
            Else
               sPropValue = oProp.Value
            End If
            sPropValue = NewText(sPropValue)
        End If
        If (eOperation = IDMReplace) Then
            Call ReplaceProperty(moAppl, miApplType, oDoc, oProp)
            Exit Function
        End If
        
        Select Case miApplType
            Case APPL_WORD
                sBookMarkName = WordInsertProperty(oProp.PropertyDescription, sPropValue, moAppl, miApplType)
                
                If sBookMarkName <> "" Then
                    Call frmPropertyMgr.UpdatePropGrid(sBookMarkName, oProp.Label, sPropValue)
                End If
            Case APPL_EXCEL
                sBookMarkName = ExcelInsertProperty(oProp.PropertyDescription, sPropValue, moAppl, miApplType)
                
                Call frmPropertyMgr.UpdatePropGrid(sBookMarkName, oProp.Label, sPropValue)
            Case Else
                'no bookmarks in powerpoint
                Exit Function
        End Select
        Call GotoProperty(sBookMarkName, miApplType, moAppl)
    End If
Done:
    Exit Function
    
errHandler:
    MsgBox Err.Description, vbCritical, LoadResString(DLG_ERR_INSERT_PROP)
End Function
Sub ReplaceProperty(oAppl As Object, iApplType As Integer, oDoc As IDMObjects.Document, oProp As IDMObjects.Property)
    '------------------------------------------------------------
    'Purpose:Handles the replacement of the property
    'Inputs:oappl - application object
    '       iApplType - Type of Application
    '       oDoc - IDM Doc Object
    '       oProp - IDM Property Object
    'Outputs:
    'Assumptions:Called by user clicking the cmdReplace Button on the frmPropertyMgr
    'Constraints
    'Copyright © 1998 FileNET Corporation
    '------------------------------------------------------------
    Dim igrdRow As Integer
    Dim sBookMarkName As String
    Dim vbResult As VbMsgBoxResult
    Dim sCurrentPropName As String
    Dim sCurrentPropValue As String
    Dim SCurrentPropLocation As String
    Dim sPropValue As String
    
    On Error GoTo errHandler
    
     igrdRow = frmPropertyMgr.GetGridRowNum
     sCurrentPropName = frmPropertyMgr.GetCurrentPropertyName
     sCurrentPropValue = frmPropertyMgr.GetCurrentPropertyValue
     SCurrentPropLocation = frmPropertyMgr.GetCurrentPropertyLocation
     
     'Message Box verifying users desire replace the property
       vbResult = MsgBox(LoadResString(PROMPT_REPLACE_PROPERTY) & Chr(13) & sCurrentPropName & Chr(13) _
       & LoadResString(MSG_WITH) & Chr(13) & oProp.Label & Chr(13), _
       vbInformation + vbYesNo, LoadResString(MSG_CONFIRM_REPLACE))
        
    If vbResult = vbYes Then
        sBookMarkName = frmPropertyMgr.GetBookMarkName(igrdRow)
        If IsNull(oProp) Then
           sPropValue = ""
        ElseIf IsNumeric(oProp.Value) Then
            sPropValue = oProp.FormatValue
        Else
            sPropValue = oProp.Value
        End If
        If IsNull(sPropValue) Then
            sPropValue = " "
        Else
            sPropValue = NewText(sPropValue)
        End If
    
         Select Case iApplType
                Case APPL_WORD
                     sBookMarkName = WordReplaceProperty(sBookMarkName, oProp.PropertyDescription, sPropValue, oAppl, iApplType)
                    'Call WordDeleteProperty(oAppl, sBookMarkName)
                    'sBookMarkName = WordInsertProperty(oProp.PropertyDescription, sPropValue, oAppl, iApplType)
                    Call frmPropertyMgr.UpdatePropGrid(sBookMarkName, oProp.Label, sPropValue)
                Case APPL_EXCEL
                    Call ExcelDeleteProperty(oAppl, sBookMarkName)
                    sBookMarkName = ExcelInsertProperty(oProp.PropertyDescription, sPropValue, oAppl, iApplType)
                    Call frmPropertyMgr.UpdatePropGrid(sBookMarkName, oProp.Label, sPropValue)
                Case Else
                    'no bookmarks in powerpoint
                    Exit Sub
         End Select
         Call GotoProperty(sBookMarkName, iApplType, oAppl)
    End If
  Exit Sub
errHandler:
  MsgBox Err.Description, vbCritical, LoadResString(DLG_ERR_REPLACE_PROP)
End Sub
Public Function DeterminePropName(sBookMarkName As String) As String
'------------------------------------------------------------
'Purpose:Takes apart a Bookmark name and returns the property name
'Inputs: sBookMarkName - bookmark name
'Outputs: Property Label
'Assumptions:
'Constraints
'Copyright © 1998 FileNET Corporation
'------------------------------------------------------------
    Dim sString As String
    Dim sIDMTag As String
    Dim iLoc As Integer
    Dim iLength As Integer
    Dim iLengthIDMTAG As Integer

    sIDMTag = LoadResString(IDM_TAG) & LoadResString(IDM_SEPARATOR)
    iLengthIDMTAG = Len(sIDMTag)
    
    iLength = Len(sBookMarkName)
    iLoc = InStr((iLengthIDMTAG + 1), sBookMarkName, LoadResString(IDM_SEPARATOR))
    
    sString = Left(sBookMarkName, (iLoc - 1))  'err (iLoc - 1)<0 11/19/98
    
    iLength = Len(sString)
    
    DeterminePropName = Right(sString, (iLength - iLengthIDMTAG))
End Function

Public Function GetPropValue(sPropName As String, sFileName As String, Optional blogon As Boolean) As String
   '------------------------------------------------------------
   'Purpose:Returns the value of a property based on the property's name
   'Inputs: sPropName - Name of Property
            'sFileName - Active documents name
   'Outputs: String representing the Property's value
   'Assumptions:
   'Constraints
   'Copyright © 1998 FileNET Corporation
   '------------------------------------------------------------
    Dim oDoc As IDMObjects.Document
    Dim oProp As IDMObjects.Property
    Dim oErrorMgr As ErrorManager
    On Error GoTo errHandler
    
    Set oErrorMgr = CreateObject(CREATE_ERR_MGR)
    If oErrorMgr Is Nothing Then
        Err.Raise 1, LoadResString(DLG_ERR_GET_PROP_PROP), LoadResString(MSG_CANNOT_CREATE_ERRMGR)
    End If
    If gdoc Is Nothing Then
        Set oDoc = GetDocObject(sFileName, blogon)
    Else
        Set oDoc = gdoc
        blogon = True
    End If
    If oDoc Is Nothing Then
        GetPropValue = " " ' LoadResString(TXT_NO_THIS_VALUE)
        Exit Function
    End If
        
    Set oProp = oDoc.GetExtendedProperty(sPropName)
    If IsEmpty(oProp) Then
        Err.Raise 1, LoadResString(MSG_CANNOT_GET_PROP_OBJECT), LoadResString(MSG_CANNOT_CREATE_ERRMGR)
    End If

    If Not IsNull(oProp.Value) Then
        If IsNumeric(oProp.Value) Then
               GetPropValue = oProp.FormatValue
        Else
               GetPropValue = oProp.Value
        End If
    Else
        GetPropValue = " " ' LoadResString(TXT_NO_THIS_VALUE)
    End If

    Set oProp = Nothing
    Set oDoc = Nothing
    Exit Function
    
errHandler:
    If Err.Number <> 0 And Err.Number <> -2147216381 Then
    Else
        If oErrorMgr Is Nothing Then
            MsgBox LoadResString(MSG_ERROR_WITHOUT_ERRMGR), vbCritical, LoadResString(DLG_ERR_GET_PROP_PROP)
            GetPropValue = " " ' LoadResString(TXT_NO_THIS_VALUE)
            Exit Function
        Else
            If oErrorMgr.Errors.Count > 0 Then
                oErrorMgr.ShowErrorDialog
            End If
        End If
    End If
    GetPropValue = " " ' LoadResString(TXT_NO_THIS_VALUE)
End Function
Public Sub WordGotoProperty(oAppl As Object, iApplType As Integer, sBookMarkName As String)
    Dim oBKMark As Object
    Dim oRange As Object
    Dim lResult As Long
    
    On Error GoTo errHandler
     
    If oAppl.ActiveWindow.View.SplitSpecial <> wdPaneNone Then
            oAppl.ActiveWindow.Panes(2).Close
    End If
    
    Set oBKMark = oAppl.ActiveDocument.Bookmarks.Item(sBookMarkName)
    If Not IsObject(oBKMark) Then
        Exit Sub
    End If
    'First determine where the property is Header, Footer Body,
    Set oRange = oBKMark.Range
    lResult = oRange.Information(wdHeaderFooterType)
    
    If (lResult >= 0) Then
        
        With oAppl.ActiveWindow.View
            .Type = wdPageView
            .SeekView = frmPropertyMgr.GetViewType(lResult)
        End With
    End If
    
    'Next Determine the Section and page the property it is located in
    oRange.Select
    Set oRange = Nothing
    Set oBKMark = Nothing
    
    Exit Sub
    
errHandler:
    MsgBox Err.Description, vbCritical, LoadResString(DLG_ERR_WORD_GOTO_PROP)
End Sub

Public Function InBookmark(oAppl As Object) As Boolean
'------------------------------------------------------------
'Purpose:Checks to see if the current property insertion _
         point is in the range of a previously inserted property
'Inputs: oAppl - Word Application item
'Outputs:
'Assumptions:
'Constraints
'Copyright © 1998 FileNET Corporation
'------------------------------------------------------------
    Dim oSelection As Object
    Dim oBMark As Object
    Dim oSelectRange As Object
    Dim iLoc As Integer
    Dim BMID As Long
    Dim sIDMTag As String
     
    On Error GoTo errHandler
    
    Set oSelection = oAppl.Selection
    Set oSelectRange = oAppl.Selection.Range
    'et oBookMarks = oAppl.ActiveDocument.Bookmarks
    
    InBookmark = False
    BMID = oSelectRange.BookmarkID
    
    If BMID > 0 Then
        sIDMTag = LoadResString(IDM_TAG) & LoadResString(IDM_SEPARATOR)
            For Each oBMark In oSelectRange.Bookmarks
                iLoc = InStr(oBMark.Name, sIDMTag)
                If iLoc > 0 Then
                    'we have at least one IDM bookmark so lets get out
                    InBookmark = True
                    Exit Function
                Else
                    InBookmark = False
                End If
            Next
    End If
    Exit Function
errHandler:
    MsgBox Err.Description, vbCritical, LoadResString(DLG_ERR_INBOOKMARK)
    InBookmark = True
End Function
Public Function InPropertyCell(oAppl As Object) As Boolean
        Dim namTemp As Object
        Dim oRange As Object
        Dim sCurrentSheet As String
        Dim sCurrentCell As String
        Dim iResult As Integer
        Dim sRangeAddress As String
        Dim sCellName As String
        Dim sRangeWorkSheetName As String
        Dim sIDMTag As String
        
        On Error GoTo errHandler
        
        InPropertyCell = False
        
        sCurrentSheet = oAppl.ActiveSheet.Name
        sCurrentCell = oAppl.ActiveCell.Address
        sIDMTag = LoadResString(IDM_TAG) & LoadResString(IDM_SEPARATOR)
        
        For Each namTemp In oAppl.ActiveWorkbook.Names
            sCellName = namTemp.Name
            
            iResult = InStr(sCellName, sIDMTag)
            If iResult > 0 Then
                Set oRange = oAppl.Range(sCellName)
                sRangeAddress = oRange.Address
                sRangeWorkSheetName = oRange.Worksheet.Name
                If StrComp(sRangeAddress, sCurrentCell) = 0 And StrComp(sRangeWorkSheetName, sCurrentSheet) = 0 Then
                    InPropertyCell = True
                    Exit Function
                Else
                    InPropertyCell = False
                End If
            End If
        Next
        Exit Function
    
errHandler:
    MsgBox Err.Description, vbCritical, LoadResString(DLG_ERR_INBOOKMARK)
    InPropertyCell = True
End Function
Public Sub SDMtoIDMConversion(oAppl As Object, iApplType As Integer)
   Dim oBMark As Bookmark
   Dim sNewBkName As String
   Dim nName As Name
   Dim sNewName As String
   
   On Error GoTo errHandler
   
   Select Case iApplType
          Case APPL_WORD:
                For Each oBMark In oAppl.ActiveDocument.Bookmarks
                    If (InStr(1, oBMark.Name, LoadResString(STR_MEZZ), 1) = 1) Then
                       sNewBkName = ConvertSDMtoIDMTag(oBMark.Name)
                       If sNewBkName = LoadResString(TXT_CANNOT_CONVERT_STRING) Then
                           MsgBox LoadResString(STR_THE_BOOKMARK) & oBMark.Name & LoadResString(STR_CAN_NOT_BE_CONVERT), vbInformation, LoadResString(MSG_LOAD_IDM_STRING)
                       Else
                           sNewBkName = sNewBkName & "_" & AddUniqueNumber(sNewBkName, oAppl, iApplType)
                           oAppl.ActiveDocument.Bookmarks.Add Name:=sNewBkName, Range:=oBMark.Range
                           oBMark.Delete
                       End If
                    End If
              Next
          Case APPL_EXCEL:
              For Each nName In oAppl.ActiveWorkbook.Names
                    If (InStr(1, nName.Name, LoadResString(STR_MEZZ), 1) = 1) Then
                       sNewName = ConvertSDMtoIDMTag(nName.Name)
                       If sNewName = LoadResString(TXT_CANNOT_CONVERT_STRING) Then
                          MsgBox LoadResString(STR_THE_BOOKMARK) & oBMark.Name & LoadResString(STR_CAN_NOT_BE_CONVERT), vbInformation, LoadResString(MSG_LOAD_IDM_STRING)
                       Else
                          sNewName = sNewName & "_" & AddUniqueNumber(sNewName, oAppl, iApplType)
                          oAppl.ActiveWorkbook.Names.Add Name:=sNewName, RefersToR1C1:=nName.RefersToR1C1
                          nName.Delete
                       End If
                    End If
              Next
   End Select
   Exit Sub

errHandler:
   MsgBox Err.Description, vbCritical, LoadResString(MSG_LOAD_IDM_STRING)
End Sub
Public Function ConvertSDMtoIDMTag(sSDMTag As String) As String
    Dim sPropertyNumber As String
    Dim sIDMTag As String

    On Error GoTo errHandler:
    
    If (InStr(1, sSDMTag, LoadResString(STR_MEZZ), 1) <> 1) Then
        ConvertSDMtoIDMTag = LoadResString(TXT_CANNOT_CONVERT_STRING) '"Nothing"
        GoTo Done
    End If
    sPropertyNumber = Mid(sSDMTag, 6, 3)
    sIDMTag = LoadResString(STR_IDM) & LoadResString(CInt(sPropertyNumber)) ' & "_"
    ConvertSDMtoIDMTag = sIDMTag

    Exit Function
errHandler:
    If (Err.Number = 326) Or (Err.Number = 13) Then
        ConvertSDMtoIDMTag = LoadResString(TXT_CANNOT_CONVERT_STRING) ' LoadResString(TXT_NO_THIS_VALUE) ' "Nothing"
    End If
    MsgBox Err.Description, vbCritical, LoadResString(MSG_CONVERSION)
Done:
End Function

Public Function AddUniqueNumber(sBkMarkName As String, oAppl As Object, iApplType As Integer) As String
    Dim oBMark As Bookmark
    Dim iPropCnt As Integer
    Dim iPos As Integer
    Dim nName As Name
    Dim sTempName As String
    
    On Error GoTo errHandler:
    
    iPropCnt = 0
    Select Case iApplType
           Case APPL_WORD:
                For Each oBMark In oAppl.ActiveDocument.Bookmarks
                    'to find 2nd '_' position
                    iPos = InStr(6, oBMark.Name, "_")
                    'compare IDM_PropName_ to converted name, if the IDM_PropName_ exists, increase the numbeer
                    sTempName = Left(oBMark.Name, iPos - 1)
                    If StrComp(sTempName, sBkMarkName) = 0 Then
                       iPropCnt = iPropCnt + 1
                    End If
                Next
           Case APPL_EXCEL:
                For Each nName In oAppl.ActiveWorkbook.Names
                    'to find 2nd '_' position
                    iPos = InStr(6, nName.Name, "_")
                    sTempName = Left(nName.Name, iPos)
                    If InStr(sTempName, sBkMarkName) Then
                       iPropCnt = iPropCnt + 1
                    End If
                Next
    End Select
    'Increment property count by 1 for next property.
    iPropCnt = iPropCnt + 1
    AddUniqueNumber = iPropCnt
    Exit Function
    
errHandler:
    MsgBox Err.Description, vbCritical, LoadResString(MSG_GET_UNIQUE_NUMBER)
End Function
Public Function SDMPropCount(oAppl As Object, iAppType As Integer) As Integer
   Dim oBMark As Bookmark
   Dim iCnt As Integer
   Dim nName As Name

   On Error Resume Next
   iCnt = 0
   Select Case iAppType
          Case APPL_WORD:
                For Each oBMark In oAppl.ActiveDocument.Bookmarks
                    If (InStr(1, oBMark.Name, LoadResString(STR_MEZZ), 1) = 1) Then
                       iCnt = iCnt + 1
                    End If
                 Next
          Case APPL_EXCEL:
              For Each nName In oAppl.ActiveWorkbook.Names
                    If (InStr(1, nName.Name, LoadResString(STR_MEZZ), 1) = 1) Then
                       iCnt = iCnt + 1
                    End If
              Next
   End Select
   SDMPropCount = iCnt
End Function
Public Function IDMPropCount(oAppl As Object, iAppType As Integer) As Integer
   Dim oBMark As Bookmark
   Dim iCnt As Integer
   Dim nName As Name

   On Error Resume Next
   iCnt = 0
   Select Case iAppType
          Case APPL_WORD:
                For Each oBMark In oAppl.ActiveDocument.Bookmarks
                    If (InStr(1, oBMark.Name, LoadResString(STR_IDM), 1) = 1) Then  '"IDM_"
                       iCnt = iCnt + 1
                    End If
                 Next
          Case APPL_EXCEL:
              For Each nName In oAppl.ActiveWorkbook.Names
                    If (InStr(1, nName.Name, LoadResString(STR_IDM), 1) = 1) Then  '"IDM_"
                       iCnt = iCnt + 1
                    End If
              Next
   End Select
   IDMPropCount = iCnt
End Function
Public Function IDMSDMPropCount(oAppl As Object, iAppType As Integer) As Integer
   Dim oBMark As Bookmark
   Dim iCnt As Integer
   Dim nName As Name

   On Error Resume Next
   iCnt = 0
   Select Case iAppType
          Case APPL_WORD:
                For Each oBMark In oAppl.ActiveDocument.Bookmarks
                    iCnt = iCnt + 1
                Next
          Case APPL_EXCEL:
              For Each nName In oAppl.ActiveWorkbook.Names
                   iCnt = iCnt + 1
              Next
   End Select
   IDMSDMPropCount = iCnt
End Function
Private Function NewText(oldText As String) As String
    Dim iPos As Integer
    'to remove return chars in a string
    NewText = ""
    Do Until InStr(oldText, vbCrLf) = 0 Or InStr(oldText, vbCrLf) = Null
       iPos = InStr(oldText, vbCrLf)
       NewText = NewText & Left(oldText, iPos - 1) & " "
       oldText = Right(oldText, Len(oldText) - iPos - 1)
    Loop
    NewText = NewText & oldText
End Function
Public Function WordReplaceProperty(sOldBkName As String, sNewBkName As String, sNewBkValue As String, oAppl As Object, iApplType As Integer) As String
    Dim oBKMark As Bookmark
    Dim oRange As Object 'Range
    Dim iStart As Integer
    Dim iEnd As Integer
    Dim iOldLength As Integer
    Dim iNewLength As Integer
    Dim iPropCnt As Integer
    
    On Error GoTo errHandler
     
    If oAppl.ActiveWindow.View.SplitSpecial <> wdPaneNone Then
            oAppl.ActiveWindow.Panes(2).Close
    End If
    'to get old bookmark information
    Set oBKMark = oAppl.ActiveDocument.Bookmarks.Item(sOldBkName)
    If Not IsObject(oBKMark) Then
        Exit Function
    End If
    Set oRange = oBKMark.Range
    iStart = oRange.Start
    iEnd = oRange.End
    iOldLength = Len(oRange.Text)
    'to get new bookmark information
    iNewLength = Len(sNewBkValue)
    oRange.Text = Trim(sNewBkValue)
    
    iPropCnt = GetMaxBkNumber(sNewBkName, oAppl, iApplType)
    'build the bookmark's name
    sNewBkName = LoadResString(IDM_TAG) & LoadResString(IDM_SEPARATOR) & sNewBkName & LoadResString(IDM_SEPARATOR) & iPropCnt
    
    'replace the text
    If (iOldLength > iNewLength) Then
        oRange.SetRange Start:=iStart, End:=iEnd - (iOldLength - iNewLength)
    End If
    If (iOldLength < iNewLength) Then
        oRange.SetRange Start:=iStart, End:=iEnd + (iNewLength - iOldLength)
    End If
    'Add the new bookmark
    oAppl.ActiveDocument.Bookmarks.Add Name:=sNewBkName, Range:=oRange
    'delete the old bookmark if exists
    Call RemoveBookMark(sOldBkName, 1, oAppl)
    frmPropertyMgr.RemoveGridEntry (sOldBkName)
    frmPropertyMgr.cmdDelete.Enabled = False
    WordReplaceProperty = sNewBkName
    Exit Function
errHandler:
  MsgBox Err.Description, vbCritical, LoadResString(DLG_ERR_REPLACE_PROP)
End Function
Private Function GetMaxBkNumber(sNewBkName As String, oAppl As Object, iApplType As Integer) As Integer
    Dim oBKMark As Bookmark
    Dim namTemp As Name
    Dim iPos As Integer
    Dim iLen As Integer
    Dim iMaxNum As Integer
    Dim iPropCnt As Integer
    iPropCnt = 0
    Select Case iApplType
           Case APPL_WORD:
                For Each oBKMark In oAppl.ActiveDocument.Bookmarks
                    If InStr(oBKMark.Name, sNewBkName) Then
                        iLen = Len(oBKMark.Name)
                        iPos = InStr(5, oBKMark.Name, "_")
                        iMaxNum = Val(Right(oBKMark.Name, (iLen - iPos)))
                        If iMaxNum > iPropCnt Then
                           iPropCnt = iMaxNum
                        End If
                    End If
                Next
           Case APPL_EXCEL:
                For Each namTemp In oAppl.ActiveWorkbook.Names
                    If InStr(namTemp.Name, sNewBkName) Then
                       iLen = Len(namTemp.Name)
                        iPos = InStr(5, namTemp.Name, "_")
                        iMaxNum = Val(Right(namTemp.Name, (iLen - iPos)))
                        If iMaxNum > iPropCnt Then
                           iPropCnt = iMaxNum
                        End If
                    End If
                Next
    End Select
    GetMaxBkNumber = iPropCnt + 1
End Function

Public Sub Hook()
'------------------------------------------------------------
'Purpose:Replaces the Native Applications Window Process function with our own
'Inputs: None
'Outputs: none
'Assumptions:None
'Constraints: None
'Copyright © 1999 FileNET Corporation
'------------------------------------------------------------

   lpPrevWndProc = SetWindowLong(gHW, GWL_WNDPROC, _
   AddressOf WindowProc)
End Sub

Public Sub Unhook()
'------------------------------------------------------------
'Purpose:Rehooks up the native application with its own Window Process function
'Inputs: None
'Outputs: None
'Assumptions: None
'Constraints: None
'Copyright © 1999 FileNET Corporation
'------------------------------------------------------------

   Dim temp As Long
   temp = SetWindowLong(gHW, GWL_WNDPROC, _
   lpPrevWndProc)
End Sub

Function WindowProc(ByVal hw As Long, ByVal uMsg As _
Long, ByVal wParam As Long, ByVal lParam As Long) As _
Long
 '------------------------------------------------------------
 'Purpose:Message Filter of the Native Application.  We use this to keep the Property Manager
 '          Dialog on top when necessary
 'Inputs: hw - handle to window
 '          uMsg - Window Message number
 '          wParam - Primary Window message parameter
 '          lParam - secondary Window Message parameter
 'Outputs: None
 'Assumptions: None
 'Constraints
 'Copyright © 1999 FileNET Corporation
 '------------------------------------------------------------
   Dim bResult As Boolean
     
    If (uMsg = WM_SETFOCUS) Then
        If wParam = lParentHWND Then
            bResult = SetWindowPos(frmPropertyMgr.hWnd, HWND_NOTOP, 0, 0, 0, 0, FLAGS)
        End If
   End If
   
   If (uMsg = WM_ACTIVATEAPP) Then
   'the wParam = 0 is checking for deactivation
       If wParam = False Then
          bResult = SetWindowPos(frmPropertyMgr.hWnd, HWND_NOTOP, 0, 0, 0, 0, FLAGS)
       Else
          bResult = SetWindowPos(frmPropertyMgr.hWnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
       End If
   End If
   
   If (uMsg = WM_ACTIVATE) Then
   'the wParam = 0 checks to see if the window is being deactivated
     If wParam = 0 And lParam = lParentHWND Then
        bResult = SetWindowPos(frmPropertyMgr.hWnd, HWND_NOTOP, 0, 0, 0, 0, FLAGS)
     End If
   End If
   
   If (uMsg = WM_SHOWWINDOW) Then
      bResult = SetWindowPos(frmPropertyMgr.hWnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
   End If
   
   'If (uMsg = WM_KILLFOCUS) Then
        'bResult = SetWindowPos(frmPropertyMgr.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
   'End If
   
   If (uMsg = WM_SIZE) Then
        Select Case wParam
        Case 0
            frmPropertyMgr.Visible = True
            bResult = SetWindowPos(frmPropertyMgr.hWnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
        Case 1
            frmPropertyMgr.Visible = False
            bResult = SetWindowPos(frmPropertyMgr.hWnd, HWND_TOP, 0, 0, 0, 0, FLAGS)
        Case 2
            frmPropertyMgr.Visible = True
            bResult = SetWindowPos(frmPropertyMgr.hWnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
        Case 3
            frmPropertyMgr.Visible = True
            bResult = SetWindowPos(frmPropertyMgr.hWnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
        Case 4
            frmPropertyMgr.Visible = False
            bResult = SetWindowPos(frmPropertyMgr.hWnd, HWND_TOP, 0, 0, 0, 0, FLAGS)
        End Select
   End If
   WindowProc = CallWindowProc(lpPrevWndProc, hw, uMsg, wParam, lParam)
End Function
Public Sub RemoveBookMark(sBkName As String, iApplType As Integer, oAppl As Object)
   Dim oBMark As Bookmark
   Dim nName As Name
    
   Select Case iApplType
          Case APPL_WORD:
                For Each oBMark In oAppl.ActiveDocument.Bookmarks
                    If InStr(oBMark.Name, sBkName) Then
                       oBMark.Delete
                    End If
               Next
          Case APPL_EXCEL:   'currently we do not the part
                For Each nName In oAppl.ActiveWorkbook.Names
                    If InStr(nName.Name, sBkName) Then
                       nName.Delete
                    End If
                Next
   End Select
End Sub
