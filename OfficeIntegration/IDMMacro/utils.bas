Attribute VB_Name = "Utils"
Option Explicit
'DTS134785 - Raja - 12/07/04
Public clpData
Public iClpBrdFormat As Integer
'DTS134785 - Raja - 12/07/04
Public Sub BkpClpBrddata()
On Error Resume Next

If (Clipboard.GetFormat(ClipBoardConstants.vbCFText)) Then
    clpData = Clipboard.GetText(ClipBoardConstants.vbCFText)
    iClpBrdFormat = ClipBoardConstants.vbCFText
ElseIf (Clipboard.GetFormat(ClipBoardConstants.vbCFBitmap)) Then
    Set clpData = Clipboard.GetData(ClipBoardConstants.vbCFBitmap)
    iClpBrdFormat = ClipBoardConstants.vbCFBitmap
ElseIf (Clipboard.GetFormat(ClipBoardConstants.vbCFDIB)) Then
    Set clpData = Clipboard.GetData(ClipBoardConstants.vbCFDIB)
    iClpBrdFormat = ClipBoardConstants.vbCFDIB
'ElseIf (Clipboard.GetFormat(ClipBoardConstants.vbCFEMetafile)) Then
 '   Set clpData = Clipboard.GetData(ClipBoardConstants.vbCFEMetafile)
'    iClpBrdFormat = ClipBoardConstants.vbCFEMetafile
'ElseIf (Clipboard.GetFormat(ClipBoardConstants.vbCFFiles)) Then
    'Set clpData = Clipboard.GetData(ClipBoardConstants.vbCFFiles)
   ' sClpData = Clipboard.GetText(ClipBoardConstants.vbCFFiles)
   ' iClpBrdFormat = ClipBoardConstants.vbCFFiles
ElseIf (Clipboard.GetFormat(ClipBoardConstants.vbCFLink)) Then
    clpData = Clipboard.GetText(ClipBoardConstants.vbCFLink)
    iClpBrdFormat = ClipBoardConstants.vbCFLink
ElseIf (Clipboard.GetFormat(ClipBoardConstants.vbCFMetafile)) Then
    Set clpData = Clipboard.GetData(ClipBoardConstants.vbCFMetafile)
    iClpBrdFormat = ClipBoardConstants.vbCFMetafile
'ElseIf (Clipboard.GetFormat(ClipBoardConstants.vbCFPalette)) Then
 '   Set clpData = Clipboard.GetData(ClipBoardConstants.vbCFMetafile)
 '   iClpBrdFormat = ClipBoardConstants.vbCFMetafile
'ElseIf (Clipboard.GetFormat(ClipBoardConstants.vbCFRTF)) Then
 '   clpData = Clipboard.GetText(ClipBoardConstants.vbCFRTF)
'    iClpBrdFormat = ClipBoardConstants.vbCFRTF
End If
End Sub
'DTS134785 - Raja - 12/07/04
Public Sub RestoreClpBrddata()
On Error Resume Next

Select Case iClpBrdFormat
Case ClipBoardConstants.vbCFText:
    Clipboard.SetText clpData, ClipBoardConstants.vbCFText
Case ClipBoardConstants.vbCFBitmap:
    Clipboard.SetData clpData, ClipBoardConstants.vbCFBitmap
Case ClipBoardConstants.vbCFDIB:
    Clipboard.SetData clpData, ClipBoardConstants.vbCFDIB
'Case ClipBoardConstants.vbCFEMetafile:
'    Clipboard.SetData clpData, ClipBoardConstants.vbCFEMetafile
'Case ClipBoardConstants.vbCFFiles:
 '   Clipboard.SetData clpData, ClipBoardConstants.vbCFFiles
Case ClipBoardConstants.vbCFLink:
    Clipboard.SetText clpData, ClipBoardConstants.vbCFLink
Case ClipBoardConstants.vbCFMetafile:
    Clipboard.SetData clpData, ClipBoardConstants.vbCFMetafile
'Case ClipBoardConstants.vbCFPalette:
'    Clipboard.SetData clpData, ClipBoardConstants.vbCFPalette
'Case ClipBoardConstants.vbCFRTF:
'    Clipboard.SetText clpData, ClipBoardConstants.vbCFRTF
End Select
End Sub
Public Sub CheckPref_Print_Checkin(oAppl As Object, iApplType As Integer)
     Dim PrefVal As Variant
     If iApplType = APPL_POWERPOINT Then
        Exit Sub
     End If
     PrefVal = GetPreferenceValue(LoadResString(STR_PRINT_DOC_ON_ADD), iApplType) '"PrintDocOnAdd"
     Select Case PrefVal
            Case LoadResString(TXT_NEVER) '"NEVER"
            Case LoadResString(TXT_ALWAYS) '"ALWAYS"
                 Call PrintOnCheckin(oAppl, iApplType)
            Case LoadResString(TXT_PROMPTUSER) '"PROMPTUSER"
                 If MsgBox(LoadResString(MSG_DO_YOU_WANT_TO_PRINT_FILE), vbInformation + vbYesNo, LoadResString(MSG_PRINT)) = vbYes Then
                     Call PrintOnCheckin(oAppl, iApplType)
                 End If
     End Select
End Sub

Function PrintOnCheckin(oAppl As Object, iApplType As Integer)
    On Error GoTo ErrorHandler
    Select Case iApplType
        Case APPL_WORD
            oAppl.Dialogs(wdDialogFilePrint).Show
        Case APPL_EXCEL
            oAppl.Dialogs(xlDialogPrint).Show
        Case APPL_POWERPOINT
    End Select
    Exit Function
ErrorHandler:
    If Err.Number = 5217 Then
       MsgBox LoadResString(PRINTER_NOT_INSTALL), vbInformation, LoadResString(MSG_PRINT)
    Else
       MsgBox Err.Description, vbCritical, LoadResString(MSG_PRINT)
    End If
End Function
Function saveChanges(oAppl As Object, iApplType As Integer, strFileFullname As String, strinroutine As String) As Long
    Dim vbResult As VbMsgBoxResult
    Dim oActiveDoc As Object
    Dim bSaved As Boolean
    Dim strfilename As String
    Dim CallingOperation As AddCheckinEnum

    On Error GoTo errHandler

    'Bail right away if no documents are open!!
    If DocCount(oAppl, iApplType) = 0 Then
       If iApplType = APPL_POWERPOINT Then
           MsgBox LoadResString(MSG_FILE_NOT_EXIST), vbOKOnly + vbCritical, strinroutine
       End If
        saveChanges = CIDMCancel
        GoTo Done
    End If
    
    strfilename = getName(oAppl, iApplType)

    If DocIsNew(oAppl, iApplType) = True Then
        If GetPreferenceValue(LoadResString(STR_PRT_SAVE_ADD), iApplType) = "1" Then
             vbResult = MsgBox(LoadResString(STR_DO_YOU_WANT_TO_SAVE_THE_DOC) & LoadResString(STR_P_LEFT) & strfilename & LoadResString(STR_P_RIGHT), vbOKCancel + vbQuestion, strinroutine)
        Else
             vbResult = vbOK
        End If
        Select Case vbResult
            Case vbOK
                strFileFullname = DEFAULT_CHECKOUT_PATH & "\" & strfilename
                If DocSaveDialog(oAppl, strFileFullname, iApplType, False) = vbCancel Then
                    saveChanges = CIDMCancel
                    GoTo Done
                Else
                    strFileFullname = getFullName(oAppl, iApplType)
   '                 If iApplType = APPL_POWERPOINT Then
   '                    Call GetActiveDocObj(oAppl, iApplType, oActiveDoc)
   '                 End If
                End If
                
            Case vbCancel
                    saveChanges = CIDMCancel
                GoTo Done
                ' Do not close Document
        
        End Select
        vbResult = vbIgnore
    Else
        strFileFullname = getFullName(oAppl, iApplType)
        '11/01/99 move to here from idmwizard.cls
        If (DocIsSaved(oAppl, iApplType) = False) Then
                If GetDocStatus(strFileFullname) <> DocCheckedout Then
                    CallingOperation = idmAdd
                Else
                    CallingOperation = idmCheckin
                End If
                Select Case promptsave(CallingOperation, strFileFullname, iApplType)
                    Case vbYes
                         Call DocSave(oAppl, iApplType) '###
                    Case vbNo
                         Call fileRevertOffice(oAppl, iApplType, strFileFullname)
                    '    Party On
                    Case vbCancel
                         saveChanges = CIDMCancel
                         GoTo Done
                         ' Do not close Document
                End Select
        End If
    End If
    saveChanges = CIDMOk
    strFileFullname = getFullName(oAppl, iApplType)
GoTo Done
errHandler:
    saveChanges = CIDMError
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbCritical, strinroutine
    End If
Done:
    Set oActiveDoc = Nothing
End Function
Function DocClose(oAppl As Object, iApplType As Integer)
    Dim bEvents As Boolean
    
    On Error Resume Next
    
    Select Case iApplType
        Case APPL_WORD:
            oAppl.ActiveDocument.Close False
       
        Case APPL_EXCEL:
            bEvents = oAppl.EnableEvents
            oAppl.EnableEvents = False
            oAppl.ActiveWorkbook.Close False
            oAppl.EnableEvents = bEvents
        
        Case APPL_POWERPOINT:
            'oAppl.ActivePresentation.Saved = True
            oAppl.ActivePresentation.Close
            'These two line are here to get the active
            'presentation to close. It opens up a new
            'presentation and closes it
            oAppl.Presentations.Add
            oAppl.ActivePresentation.Close
     End Select
End Function

Function DocIsNew(oAppl As Object, iApplType As Integer) As Boolean
    DocIsNew = False
    Select Case iApplType
        Case APPL_WORD
            'If oAppl.ActiveDocument.Path = "" Then
            If ActiveDocument.Path = "" Then   'for OfficeXP case
                DocIsNew = True
            End If
        Case APPL_EXCEL
            If oAppl.ActiveWorkbook.Path = "" Then
                DocIsNew = True
            End If
        Case APPL_POWERPOINT
            If oAppl.ActivePresentation.Path = "" Then
                DocIsNew = True
            End If
    End Select
End Function

Function DocIsSaved(oAppl As Object, iApplType As Integer) As Boolean
'Description:
' returns true if active doc does not have any unsaved changes
' returns false if active doc has changes that need to be saved
    Select Case iApplType
        Case APPL_WORD
            DocIsSaved = oAppl.ActiveDocument.Saved
        Case APPL_EXCEL
            DocIsSaved = oAppl.ActiveWorkbook.Saved
        Case APPL_POWERPOINT
            DocIsSaved = oAppl.ActivePresentation.Saved
        Case Else
            DocIsSaved = False
    End Select

End Function
Function DocContained(oAppl As Object, iApplType As Integer) As ContainerTypeEnum
Dim contain As Object
Dim sdirectory As String
Dim sFilePath As String
Dim cbcParentMenu As CommandBarControl
Dim cbcChildMenu As CommandBarControl
On Error GoTo errorexit
    If oAppl Is Nothing Then
        DocContained = NoContainer
        Exit Function
    Else
        If (DocCount(oAppl, iApplType) = 0) Then
            DocContained = NoContainer
            Exit Function
        End If
        Select Case iApplType
            Case APPL_WORD:
                Dim iCount As Integer
                iCount = ActiveDocument.Shapes.Count
                If (iCount > 0) Then   'added this one for solving a XP case
                     Set contain = ActiveDocument.Container
                Else
                    DocContained = NoContainer
                    Exit Function
                End If
            Case APPL_EXCEL:
                Set contain = oAppl.ActiveWorkbook.Container
            Case APPL_POWERPOINT:
                Set contain = oAppl.ActivePresentation.Container
        End Select
        DocContained = ContainerInt
        Exit Function
    End If
errorexit:
    If Err.Number = 4248 Then
        DocContained = NoContainer
        Exit Function
    End If
    If Err.Number = 91 Then
        DocContained = ContainerInt
        Exit Function
    End If
    sFilePath = getFullName(oAppl, iApplType)
    If sFilePath = "" Then
        DocContained = ContainerAtt
        Exit Function
    End If
    Set cbcParentMenu = oAppl.CommandBars.FindControl(ID:=MB_FILE)
    For Each cbcChildMenu In cbcParentMenu.Controls
        If cbcChildMenu.BuiltIn = True Then
            If (cbcChildMenu.ID = 106) And (InStr(1, cbcChildMenu.Caption, LoadResString(MSG_RETURN), vbTextCompare)) Then
                DocContained = ContainerInt
                Exit Function
            End If
        End If
    Next cbcChildMenu
    If (DocIsNew(oAppl, iApplType) = True) Then
        If InStr(1, sFilePath, LoadResString(STR_IN), vbTextCompare) > 0 Then    '" in "
            DocContained = ContainerInt
            Exit Function
        End If
    End If
    sdirectory = String(260, " ")
    GetTempPath 260, sdirectory
    sdirectory = RTrim(sdirectory)
    If (Mid(sdirectory, 1, Len(sdirectory) - 1) = Mid(sFilePath, 1, Len(sdirectory) - 1)) Then
        DocContained = ContainerAtt
    ElseIf InStr(UCase(sFilePath), UCase(GetPreferenceValue(LoadResString(STR_CACHE_DIRECTORY), DIRECTORIES_FILES))) Then
        DocContained = ContainerAtt
    Else
        DocContained = NoContainer
    End If
End Function

Public Function ConvertWord6_95ToCurrentWordVersion(oAppl As Object, iApplType As Integer)
    
    On Error Resume Next
    
    Select Case iApplType
        Case APPL_WORD:
            If oAppl.ActiveDocument.SaveFormat = 18 Then
                'Word 6.0/ 95 convert to word 10 or current doc format
                If CBool(GetPreferenceValue(LoadResString(STR_CONVERT_WORD6_95_TO_CURRENT_WORD_VERSION), iApplType)) = True Then
                   ActiveDocument.SaveAs FileFormat:=wdFormatDocument
                End If
            End If
   End Select
             
End Function

Function DocSave(oAppl As Object, iApplType As Integer)
    On Error Resume Next
    Select Case iApplType
        Case APPL_WORD:
            If oAppl.ActiveDocument.SaveFormat = 18 Then
                'Word 6.0/ 95 convert to word 10 or current doc format
                If CBool(GetPreferenceValue(LoadResString(STR_CONVERT_WORD6_95_TO_CURRENT_WORD_VERSION), iApplType)) = True Then
                   ActiveDocument.SaveAs FileFormat:=wdFormatDocument
                End If
             Else
                oAppl.ActiveDocument.Save
             End If
        Case APPL_EXCEL:
            oAppl.ActiveWorkbook.Save
        Case APPL_POWERPOINT:
            oAppl.ActivePresentation.Save
    End Select
End Function

Function DocSaveAs(oAppl As Object, iApplType As Integer, FileName As String)
    Select Case iApplType
        Case APPL_WORD:
            oAppl.ActiveDocument.SaveAs FileName
        Case APPL_EXCEL:
            oAppl.ActiveWorkbook.SaveAs FileName
        Case APPL_POWERPOINT:
            oAppl.ActivePresentation.SaveAs FileName
    End Select
End Function

Private Function DocSaveBeforeAdd(oAppl As Object, iApplType As Integer, strSavePath)
    Dim bEvents As Boolean
    Dim strfilename As String

    Select Case iApplType
        Case APPL_WORD:
            strfilename = oAppl.ActiveDocument.Name
            killFile (strSavePath & "\" & strfilename)
            oAppl.ActiveDocument.SaveAs (strSavePath & "\" & strfilename)
       
        Case APPL_EXCEL:
            strfilename = oAppl.ActiveWorkbook.Name
            bEvents = oAppl.EnableEvents
            oAppl.EnableEvents = False
            killFile (strSavePath & "\" & strfilename)
            oAppl.ActiveWorkbook.SaveAs (strSavePath & "\" & strfilename)
            oAppl.EnableEvents = bEvents
        
        Case APPL_POWERPOINT:
            strfilename = oAppl.ActivePresentation.Name
            killFile (strSavePath & "\" & strfilename)
            oAppl.ActivePresentation.SaveAs (strSavePath & "\" & strfilename)
    
    End Select

End Function

Private Function killFile(strFullName As String)
    If Dir$(strFullName) <> "" Then
        Kill strFullName
    End If
End Function

Function DisplayCommonDialog(oAppl As Object, iApplType As Integer, strfilename As String, bIncActive As Boolean) As VbMsgBoxResult
    
    Dim vbResult As VbMsgBoxResult
    Dim strSaveAsName As String
    
    DisplayCommonDialog = vbIgnore
    
    On Error Resume Next
    Err.Number = 0
    
    frmCommonDialogs.SaveDialogSetup iApplType      'initalize common dialog form

    If strfilename <> "" Then
        'stuff auto filename into save dialog
        frmCommonDialogs.CommonDialog1.FileName = strfilename
    End If
    
    Do
        Load frmCommonDialogs
        frmCommonDialogs.Left = (Screen.Width - frmCommonDialogs.Width) / 2
        frmCommonDialogs.Top = (Screen.Height - frmCommonDialogs.Height) / 2
        
        frmCommonDialogs.CommonDialog1.ShowSave
        If Err.Number <> 0 Then
            DisplayCommonDialog = vbCancel
            GoTo errHandler
        Else
            DisplayCommonDialog = vbOK
            vbResult = vbOK
        End If
        strSaveAsName = frmCommonDialogs.CommonDialog1.FileName
        ' Check to see if document exists as a local file
        If Dir(strSaveAsName) <> "" Then
            If GetDocStatus(strSaveAsName) = DocCheckedout Then
                vbResult = MsgBox(strSaveAsName & LoadResString(MSG_FILE_SAVEAS_CHECKEDOUT), vbOKCancel + vbExclamation, LoadResString(MSG_SAVEAS))
                If vbResult = vbCancel Then
                    Exit Do
                End If
            Else
                If DocIsOpen(oAppl, iApplType, strSaveAsName, bIncActive) = True Then
                ' Check to see if document is allready open in application
                    vbResult = MsgBox(strSaveAsName & LoadResString(MSG_FILE_SAVEAS_OVERWRITE), vbOKCancel + vbExclamation, LoadResString(MSG_SAVEAS))
                    If vbResult = vbCancel Then
                        Exit Do
                    End If
                Else
                    vbResult = MsgBox(strSaveAsName & LoadResString(MSG_FILE_EXISTS_OVERWRITE), vbYesNoCancel + vbQuestion, LoadResString(MSG_SAVEAS))
                    If vbResult = vbCancel Then
                        Exit Do
                    ElseIf vbResult = vbYes Then
                        vbResult = vbOK
                        Exit Do
                    End If
                End If
            End If
        Else
            Exit Do
        End If
    Loop
            
    strfilename = strSaveAsName
    DisplayCommonDialog = vbResult
    Exit Function
errHandler:
    Unload frmCommonDialogs
End Function

Function DocSaveDialog(oAppl As Object, strfilename As String, iApplType As Integer, bIncActive As Boolean) As VbMsgBoxResult
    'Call save as dialog
    'default newfilename if provided
    '-------------------------------------
    'returns True if successful
    'returns False if fails
    '-----------------------------------
    
    Dim vbResult As VbMsgBoxResult
    Dim lResult As Long
    Dim sNewFileName As String
    Dim iResp As Integer
    Dim sCheckedOut As String
    Dim ppfiletype As PpSaveAsFileType
    Dim hand As Long
    Dim tempname As String
    Dim saveformat1 As Long
    DocSaveDialog = vbIgnore

    Select Case iApplType
        Case APPL_WORD
            If strfilename = "" Then
                hand = GetFocus()
                With oAppl.Dialogs(wdDialogFileSaveAs)
                    lResult = .Display
                    If lResult Then
                        tempname = .Name
                        saveformat1 = .Format
                        .Update
                        .Name = tempname
                        .Format = saveformat1
                        .Execute
                    End If
                End With
                SetFocus (hand)
            Else
                hand = GetFocus()
                With oAppl.Dialogs(wdDialogFileSaveAs)
                    .Name = strfilename
                    lResult = .Display
                    If lResult Then
                        tempname = .Name
                        saveformat1 = .Format
                        .Update
                        .Name = tempname
                        .Format = saveformat1
                        .Execute
                    End If
                End With
                SetFocus (hand)
            End If
            If lResult = True Then
                'saved
                DocSaveDialog = vbOK
            Else
                DocSaveDialog = vbCancel
            End If

        Case APPL_EXCEL
            'excel save dialog
            If strfilename = "" Then
                lResult = oAppl.Dialogs(xlDialogSaveAs).Show
            Else
                lResult = oAppl.Dialogs(xlDialogSaveAs).Show(arg1:=strfilename)
            End If
            If lResult = True Then
                'saved
                DocSaveDialog = vbOK
            Else
                DocSaveDialog = vbCancel
            End If
        Case APPL_POWERPOINT
            vbResult = DisplayCommonDialog(oAppl, iApplType, strfilename, bIncActive)
            If (vbResult = vbOK) Then
                Select Case frmCommonDialogs.CommonDialog1.FilterIndex
                    Case 1
                        ppfiletype = ppSaveAsPresentation
                    Case 2
                         ppfiletype = ppSaveAsRTF
                    Case 3
                         ppfiletype = ppSaveAsTemplate
                    Case 4
                        ppfiletype = ppSaveAsShow
                    Case 5
                        ppfiletype = ppSaveAsPresentation
                    Case 6
                        ppfiletype = ppSaveAsPowerPoint7
                    Case 7
                        ppfiletype = ppSaveAsPowerPoint4
                    Case 8
                        ppfiletype = ppSaveAsPowerPoint3
                    Case 9
                        ppfiletype = ppSaveAsAddIn
                        
                End Select
                oAppl.ActivePresentation.SaveAs strfilename, ppfiletype
            Else
                DocSaveDialog = vbCancel
            End If
    End Select
        
End Function

Function DocSaveLocalDialog(oAppl As Object, strfilename As String, iApplType As Integer, bIncActive As Boolean) As VbMsgBoxResult
    'Call save as dialog
    'default newfilename if provided
    '-------------------------------------
    'returns True if successful
    'returns False if fails
    '-----------------------------------
    
    Dim vbResult As VbMsgBoxResult
    Dim strSaveAsName As String
    Dim sStripedFileName As String
    Dim sExt As String
    Dim sExtpp As String
    Dim sName As String
    Dim sDir As String
    Dim bRemove As Boolean
    Dim sFilter As String
    Dim sOldExt As String
    
    DocSaveLocalDialog = False
    On Error Resume Next
    Err.Number = 0
    
    Call GetFileNameAndExt(strfilename, sStripedFileName, sExt)
    sOldExt = sExt
    
    sFilter = ProcessFilterList(sExt)
    Call frmCommonDialogs.SaveLocalDialogSetup(VBA.Mid(sExt, 1, Len(sExt) - 1), sFilter) 'initalize common dialog from
    If strfilename <> "" Then
        'stuff auto filename into save dialog
        frmCommonDialogs.CommonDialog1.FileName = strfilename
    End If
    'Do
    Load frmCommonDialogs
    frmCommonDialogs.Left = (Screen.Width - frmCommonDialogs.Width) / 2
    frmCommonDialogs.Top = (Screen.Height - frmCommonDialogs.Height) / 2
'ShowSaveDialogAgain:
    Do
        frmCommonDialogs.CommonDialog1.ShowSave
        If Err.Number <> 0 Then
            DocSaveLocalDialog = vbCancel
            GoTo errHandler
        End If
        strSaveAsName = frmCommonDialogs.CommonDialog1.FileName
        
        'to handle ..do/..sl/..pp case
        Call ResetFileFullName(strSaveAsName, sOldExt)
        
        Call GetFileNameAndExt(strSaveAsName, sStripedFileName, sExt)
        
        If Len(sExt) = 0 Then
            strSaveAsName = strSaveAsName & "." & frmCommonDialogs.CommonDialog1.DefaultExt
        End If
        '==================
        Dim iLenLimit As Integer
        Select Case iApplType
               Case APPL_EXCEL
                    iLenLimit = 218
               Case Else
                    iLenLimit = 256
        End Select
        bRemove = False
        Do Until (Len(strSaveAsName) <= iLenLimit And bRemove = True)
           If Len(strSaveAsName) > iLenLimit Then
                MsgBox LoadResString(MSG_FILE_PATH_AND_NAME_TOO_LONG) & vbCrLf & strSaveAsName, vbInformation, LoadResString(STR_CHECK_OR_COPY_TO)
                frmCommonDialogs.CommonDialog1.ShowSave
                If Err.Number <> 0 Then
                   DocSaveLocalDialog = vbCancel
                   GoTo errHandler
                End If
                strSaveAsName = frmCommonDialogs.CommonDialog1.FileName
                'to handle ..do/..sl/..pp case
                Call ResetFileFullName(strSaveAsName, sOldExt)
           End If
           idmGetDirectoryAndFileName strSaveAsName, sDir, sName
           If bRemove = False Then
                Dim sIllegalChar As String
                If FindIllegalChar(sName, sExt, sIllegalChar) = True Then
                   If (MsgBox(LoadResString(MSG_FILE_NAME_CANNOT_CONTAIN_CHARACTER) & sIllegalChar & vbCrLf & LoadResString(MSG_DELETE_CHARACTER), vbInformation + vbYesNo, LoadResString(STR_CHECK_OR_COPY_TO))) = vbYes Then
                      Call RemoveIllegalChars(strSaveAsName, sDir, sName, sExt)
                      bRemove = True
                      'Call GetFileNameAndExt(strSaveAsName, sStripedFileName, sExt)
                      'strSaveAsName = sStripedFileName
                   Else
                      frmCommonDialogs.CommonDialog1.FileName = strSaveAsName
                      frmCommonDialogs.CommonDialog1.ShowSave
                      If Err.Number <> 0 Then
                      DocSaveLocalDialog = vbCancel
                         GoTo errHandler
                      End If
                      strSaveAsName = frmCommonDialogs.CommonDialog1.FileName
                      'to handle ..do/..sl/..pp case
                      Call ResetFileFullName(strSaveAsName, sOldExt)
                      bRemove = False
                   End If
                 Else
                    bRemove = True
                 End If
            End If
        Loop
        'If iApplType = APPL_POWERPOINT Then
        '   Call GetFileNameAndExt(strSaveAsName, sStripedFileName, sExtpp)
        '   If sExtpp = ".ppt" Then
        '      strSaveAsName = sStripedFileName
        '   End If
        'End If
        
        'check if checkout or copy path exists
        idmGetDirectoryAndFileName strSaveAsName, sDir, sName
        If Dir(sDir, vbDirectory) = "" Then
           MsgBox strfilename & vbCrLf & LoadResString(STR_PATH_DOES_NOT_EXIST) & vbCrLf & LoadResString(STR_VERIFY_PATH), vbInformation, LoadResString(STR_CHECK_OR_COPY_TO)
        ' Check to see if document exists as a local file
        ElseIf Dir(strSaveAsName) <> "" Then
            If GetDocStatus(strSaveAsName) = DocCheckedout Then
                vbResult = MsgBox("The <" & strSaveAsName & LoadResString(MSG_FILE_SAVEAS_CHECKEDOUT), vbOKCancel + vbExclamation, LoadResString(MSG_SAVEAS))
                If vbResult = vbCancel Then
                    Exit Do
                End If
            Else
                If DocIsOpen(oAppl, iApplType, strSaveAsName, bIncActive) = True Then
                    ' Check to see if document is allready open in application
                    vbResult = MsgBox(strSaveAsName & LoadResString(MSG_FILE_SAVEAS_OVERWRITE), vbOKCancel + vbExclamation, LoadResString(STR_WARNING))
                    If vbResult = vbCancel Then
                        Exit Do
                    End If
                Else
                    vbResult = MsgBox(strSaveAsName & LoadResString(MSG_FILE_EXISTS_OVERWRITE), vbYesNoCancel + vbInformation, LoadResString(STR_WARNING))
                    If vbResult <> vbNo Then
                        Exit Do
                    End If
                End If
            End If
        Else
            Exit Do
        End If
    Loop
    strfilename = strSaveAsName   'file path and name
    DocSaveLocalDialog = vbResult
    
errHandler:

    Unload frmCommonDialogs

End Function
Function DocSavedByfileName(oAppl As Object, iApplType As Integer, strfilepath As String) As Boolean
    
    'returns True  if document is already open in application (dependency on bIncActive parameter)
    'returns False if not open

    Dim strTempFilePath As String
    Dim lCount As Long
    Dim lIndex As Long
    
 '   On Error GoTo Errorout:
    
    Select Case iApplType
       Case APPL_WORD:
            lCount = oAppl.Documents.Count
            lIndex = 1
            Do While lIndex <= lCount
                strTempFilePath = oAppl.Documents.Item(lIndex).FullName
                If strfilepath = strTempFilePath Then
                    DocSavedByfileName = oAppl.Documents.Item(lIndex).Saved
                    Exit Function
                End If
                lIndex = lIndex + 1
            Loop
       
        Case APPL_EXCEL:
            lCount = oAppl.Workbooks.Count
            lIndex = 1
            Do While lIndex <= lCount
                strTempFilePath = oAppl.Workbooks.Item(lIndex).FullName
                If strfilepath = strTempFilePath Then
                    DocSavedByfileName = oAppl.Workbooks.Item(lIndex).Saved
                    Exit Function
                End If
                lIndex = lIndex + 1
            Loop
        
        Case APPL_POWERPOINT:
            lCount = oAppl.Presentations.Count
            lIndex = 1
            Do While lIndex <= lCount
                strTempFilePath = idmGetFullPathName(oAppl.Presentations.Item(lIndex))
                If strfilepath = strTempFilePath Then
                     DocSavedByfileName = oAppl.Presentations.Item(lIndex).Saved
                     Exit Function
                End If
                lIndex = lIndex + 1
            Loop
    End Select
    
    Exit Function
'Errorout:
'     If Err.Number <> 9 Then
'        MsgBox Err.Description
'     End If
End Function
Sub DocMakeActive(oAppl As Object, iApplType As Integer, strfilepath As String)
    
    'returns True  if document is already open in application (dependency on bIncActive parameter)
    'returns False if not open

    Dim strTempFilePath As String
    Dim lCount As Long
    Dim lIndex As Long
    
 '   On Error GoTo Errorout:
    
    Select Case iApplType
       Case APPL_WORD:
            lCount = oAppl.Documents.Count
            lIndex = 1
            Do While lIndex <= lCount
                strTempFilePath = oAppl.Documents.Item(lIndex).FullName
                If strfilepath = strTempFilePath Then
                    oAppl.Documents.Item(lIndex).Activate
                    Exit Sub
                End If
                lIndex = lIndex + 1
            Loop
       
        Case APPL_EXCEL:
            lCount = oAppl.Workbooks.Count
            lIndex = 1
            Do While lIndex <= lCount
                strTempFilePath = oAppl.Workbooks.Item(lIndex).FullName
                If strfilepath = strTempFilePath Then
                    oAppl.Workbooks.Item(lIndex).Activate
                    Exit Sub
                End If
                lIndex = lIndex + 1
            Loop
        
        Case APPL_POWERPOINT:
 '           lCount = oAppl.Presentations.Count
 '           lIndex = 1
 '           Do While lIndex <= lCount
 '               strTempFilePath = idmGetFullPathName(oAppl.Presentations.Item(lIndex))
 '               If strfilepath = strTempFilePath Then
 '                    oAppl.Presentations.Item(lIndex).Activate
 '                    Exit Sub
 '               End If
 '               lIndex = lIndex + 1
 '           Loop
    End Select
    
    Exit Sub
'Errorout:
'     If Err.Number <> 9 Then
'        MsgBox Err.Description
'     End If
End Sub
Sub DocCloseByfileName(oAppl As Object, iApplType As Integer, strfilepath As String)
    
    'returns True  if document is already open in application (dependency on bIncActive parameter)
    'returns False if not open

    Dim strTempFilePath As String
    Dim lCount As Long
    Dim lIndex As Long
    
 '   On Error GoTo Errorout:
    
    Select Case iApplType
       Case APPL_WORD:
            lCount = oAppl.Documents.Count
            lIndex = 1
            Do While lIndex <= lCount
                strTempFilePath = oAppl.Documents.Item(lIndex).FullName
                If UCase(strfilepath) = UCase(strTempFilePath) Then
                    oAppl.Documents.Item(lIndex).Save
                    oAppl.Documents.Item(lIndex).Close False
                    Exit Sub
                End If
                lIndex = lIndex + 1
            Loop
       
        Case APPL_EXCEL:
            lCount = oAppl.Workbooks.Count
            lIndex = 1
            Do While lIndex <= lCount
                strTempFilePath = oAppl.Workbooks.Item(lIndex).FullName
                If strfilepath = strTempFilePath Then
                    oAppl.Workbooks.Item(lIndex).Close True
                    Exit Sub
                End If
                lIndex = lIndex + 1
            Loop
        
        Case APPL_POWERPOINT:
            lCount = oAppl.Presentations.Count
            lIndex = 1
            Do While lIndex <= lCount
                strTempFilePath = idmGetFullPathName(oAppl.Presentations.Item(lIndex))
                If strfilepath = strTempFilePath Then
                    oAppl.Presentations.Item(lIndex).Save
                    oAppl.Presentations.Item(lIndex).Close
                    'These two line are here to get the active
                    'presentation to close. It opens up a new
                    'presentation and closes it
                    oAppl.Presentations.Add
                    oAppl.ActivePresentation.Close
                    Exit Sub
                End If
                lIndex = lIndex + 1
            Loop
    End Select
    
    Exit Sub
'Errorout:
'     If Err.Number <> 9 Then
'        MsgBox Err.Description
'     End If
End Sub
Function DocIsOpen(oAppl As Object, iApplType As Integer, strfilepath As String, bIncActive As Boolean) As Boolean
    
    'returns True  if document is already open in application (dependency on bIncActive parameter)
    'returns False if not open

    Dim oActiveDoc As Object
    Dim strTempFilePath As String
    Dim lCount As Long
    Dim lIndex As Long
    
    DocIsOpen = False

    Select Case iApplType
       Case APPL_WORD:
            lCount = oAppl.Documents.Count
            lIndex = 1
            Do While lIndex <= lCount
                strTempFilePath = oAppl.Documents.Item(lIndex).FullName
                If UCase(strfilepath) = UCase(strTempFilePath) Then
                    If bIncActive = False Then
                        'Only set DocIsOpen to True when document is not the active document
                        If UCase(strfilepath) <> UCase(oAppl.ActiveDocument.FullName) Then
                            DocIsOpen = True
                            Exit Do
                        End If
                    Else    ' Document is open in the application, don't care about active
                        DocIsOpen = True
                        Exit Do
                    End If
                End If
                lIndex = lIndex + 1
            Loop
       
        Case APPL_EXCEL:
            lCount = oAppl.Workbooks.Count
            lIndex = 1
            Do While lIndex <= lCount
                strTempFilePath = oAppl.Workbooks.Item(lIndex).FullName
                ' to handle NNC problem for comp Doc
                Dim bDocActive As Boolean
                bDocActive = CompareFilePath(strTempFilePath, strfilepath)
                
                If ((strfilepath = strTempFilePath) Or bDocActive = True) Then
                    If bIncActive = False Then
                        'Only set DocIsOpen to True when document is not the active document
                        If strfilepath <> oAppl.ActiveWorkbook.FullName Then
                            DocIsOpen = True
                            Exit Do
                        End If
                    Else    ' Document is open in the application, don't care about active
                        DocIsOpen = True
                        Exit Do
                    End If
                End If
                lIndex = lIndex + 1
                bDocActive = False
            Loop
        
        Case APPL_POWERPOINT:
            lCount = oAppl.Presentations.Count
            lIndex = 1
            Do While lIndex <= lCount
                strTempFilePath = idmGetFullPathName(oAppl.Presentations.Item(lIndex))
                If strfilepath = strTempFilePath Then
                    If bIncActive = False Then
                        'Only set DocIsOpen to True when document is not the active document
                        Call GetActiveDocObj(oAppl, iApplType, oActiveDoc)
                        If strfilepath <> oAppl.ActivePresentation.FullName Then
                            DocIsOpen = True
                            Exit Do
                        End If
                    Else    ' Document is open in the application, don't care about active
                        DocIsOpen = True
                        Exit Do
                    End If
                End If
                lIndex = lIndex + 1
            Loop
        Case APPL_WORDPRO
            DocIsOpen = False
    End Select
    
End Function

Public Sub GetActiveDocObj(oAppl As Object, iApplType As Integer, oActiveDoc As Object)

    Select Case iApplType
        Case APPL_WORD
            Set oActiveDoc = oAppl.ActiveDocument
        Case APPL_EXCEL
            Set oActiveDoc = oAppl.ActiveWorkbook
        Case APPL_POWERPOINT
            Set oActiveDoc = oAppl.ActivePresentation
    End Select

End Sub

Function QueryValue(sKeyName As String, sValueName As String) As Variant
    Dim lRetVal As Long        'result of the API functions
    Dim hKey As Long           'handle of opened key
    Dim vValue As Variant      'setting of queried value
    lRetVal = RegOpenKeyEx(HKEY_CURRENT_USER, sKeyName, 0, KEY_ALL_ACCESS, hKey)
    lRetVal = QueryValueEx(hKey, sValueName, vValue)
               
    RegCloseKey (hKey)
    QueryValue = vValue
End Function

Private Function QueryValueEx(ByVal lhKey As Long, ByVal szValueName As String, vValue As Variant) As Long
    Dim cch As Long
    Dim lrc As Long
    Dim lType As Long
    Dim lvalue As Long
    Dim sValue As String
    On Error GoTo QueryValueExError
       
    ' Determine the size and type of data to be read
    lrc = RegQueryValueExNULL(lhKey, szValueName, 0&, lType, 0&, cch)
    If lrc <> ERROR_NONE Then
        Error 5
    End If
    
    Select Case lType
        
        ' For strings
        Case REG_SZ:
            sValue = String(cch, 0)
            lrc = RegQueryValueExString(lhKey, szValueName, 0&, lType, sValue, cch)
            If lrc = ERROR_NONE Then
                vValue = Left$(sValue, cch - 1)
            Else
                vValue = Empty
            End If
            
        ' For DWORDS
        Case REG_DWORD:
            lrc = RegQueryValueExLong(lhKey, szValueName, 0&, lType, lvalue, cch)
            If lrc = ERROR_NONE Then
                vValue = lvalue
            End If
                
        Case Else
            'all other data types not supported
            lrc = -1
        
       End Select
       
QueryValueExExit:
    QueryValueEx = lrc
    Exit Function

QueryValueExError:
    Resume QueryValueExExit
End Function

Function idmGetFullPathName(oActiveDoc As Object) As String
'
' Gets filename from window and calls kernel Dll to get the PathName if filename is not in 8.3 format
'
    Dim strShortName As String * 256
    Dim strFilePart As String * 256
    Dim newLength%
    
    On Error Resume Next
    newLength% = GetFullPathName(oActiveDoc.FullName, 256, strShortName, strFilePart)
    
    If newLength% = 0 Then
        idmGetFullPathName = oActiveDoc.FullName
    Else
        idmGetFullPathName = Left(strShortName, newLength%)
    End If
    
End Function

Function idmGetDirectoryAndFileName(strfilepath As String, strDirectory As String, strfilename As String)

   Dim lIndex As Long
   
   For lIndex = Len(strfilepath) To 1 Step -1
        If Mid$(strfilepath, lIndex, 1) = "\" Then
            strfilename = Right$(strfilepath, Len(strfilepath) - lIndex)
            strDirectory = Left$(strfilepath, lIndex - 1) ' Remove last Slash character
            Exit For
        End If
    Next lIndex

End Function

Function CenterForm(objFrm As Form)
    objFrm.Left = (Screen.Width - objFrm.Width) / 2
    objFrm.Top = (Screen.Height - objFrm.Height) / 2
End Function
Function initializeVars(iApplType As Integer)

    Select Case iApplType
        Case APPL_WORD
            gMNU_CANCEL_CHECKOUT = TXT_CANCEL_CHECKOUT1
        
        Case APPL_EXCEL
            gMNU_CANCEL_CHECKOUT = TXT_CANCEL_CHECKOUT1
        
        Case APPL_POWERPOINT
            gMNU_CANCEL_CHECKOUT = TXT_CANCEL_CHECKOUT2
    
    End Select
    
    'this is necessary because the acclerator key for CancelCheckout is different in PPT
    g_FN_CANCEL(APPL_WORD) = LoadResString(TXT_CANCEL_CHECKOUT1)
    g_FN_CANCEL(APPL_EXCEL) = LoadResString(TXT_CANCEL_CHECKOUT1)
    g_FN_CANCEL(APPL_POWERPOINT) = LoadResString(TXT_CANCEL_CHECKOUT2)

End Function

Function getName(oAppl As Object, iApplType As Integer) As String
 
 On Error GoTo errHandler
    
    Select Case iApplType
        Case APPL_WORD:
            getName = ActiveDocument.Name
        
        Case APPL_EXCEL:
            getName = oAppl.ActiveWorkbook.Name
        
        Case APPL_POWERPOINT:
            getName = oAppl.ActivePresentation.Name
    
    End Select
    
    Exit Function

errHandler:
    getName = ""
    
End Function

Function getFullName(oAppl As Object, iApplType As Integer) As String
    On Error GoTo errorexit
    Select Case iApplType
        Case APPL_WORD:
            getFullName = ActiveDocument.FullName
           
        Case APPL_EXCEL:
            getFullName = oAppl.ActiveWorkbook.FullName
        
        Case APPL_POWERPOINT:
            getFullName = oAppl.ActivePresentation.FullName
    
    End Select
    Exit Function
errorexit:
    getFullName = ""
End Function
Function fileOpenOffice(oAppl As Object, iApplType As Integer, strfilepath As String)
Dim vbResult As String
    
    Select Case iApplType
        Case APPL_WORD:
            Call oAppl.Documents.Open(FileName:=strfilepath, AddToRecentFiles:=True)
       
        Case APPL_EXCEL:
            Call oAppl.Workbooks.Open(FileName:=strfilepath, AddToMRU:=True)
            oAppl.Visible = True
        
        Case APPL_POWERPOINT:
            Call oAppl.Presentations.Open(strfilepath)
            oAppl.Visible = True
            If PowerPointLinkExsists(oAppl) Then
                vbResult = MsgBox(LoadResString(MSG_PRESENTATION) & strfilepath & LoadResString(MSG_CONTAINSLINKS), vbOKCancel + vbQuestion, LoadResString(MSG_OPEN))
                If vbResult = vbOK Then
                    Call oAppl.ActivePresentation.UpdateLinks
                End If
            End If
    End Select

End Function

Function fileRevertOffice(oAppl As Object, iApplType As Integer, strfilepath As String)
      
    On Error Resume Next
    
    Select Case iApplType
        Case APPL_WORD:
            Call oAppl.Documents.Open(FileName:=strfilepath, AddToRecentFiles:=True, Revert:=True)
       
        Case APPL_EXCEL:
            Call oAppl.ActiveWorkbook.Close(False)
            Call oAppl.Workbooks.Open(FileName:=strfilepath, AddToMRU:=True)
            oAppl.Visible = True
        
        Case APPL_POWERPOINT:
            Call oAppl.ActivePresentation.Close
            Call oAppl.Presentations.Open(strfilepath)
            oAppl.Visible = True
    
    End Select

End Function
Function readDefaultSavePath(iApplType As Integer)
    Dim varReturn As Variant

    Select Case iApplType
        Case APPL_WORD:
            varReturn = QueryValue(LoadResString(REG_KEY_WORD), LoadResString(DOC_PATH))

        Case APPL_EXCEL:
            varReturn = QueryValue(LoadResString(REG_KEY_EXCELL), LoadResString(DEFAULT_PATH))
        
        Case APPL_POWERPOINT:
            varReturn = QueryValue(LoadResString(REG_KEY_POWERPOINT), LoadResString(Default))
    
    End Select
    readDefaultSavePath = CStr(varReturn)

End Function

Function DocCount(oAppl As Object, iApplType As Integer) As Integer
    
'Description:
' returns number of documents in application
    Select Case iApplType
        Case APPL_WORD
            DocCount = oAppl.Documents.Count
        Case APPL_EXCEL
            DocCount = oAppl.Workbooks.Count
        Case APPL_POWERPOINT
            DocCount = oAppl.Presentations.Count
        Case Else
            DocCount = 0
    End Select

End Function

Function idmGetShortPathName(oActiveDoc As Object) As String
    Dim newShortName As String * 128
    Dim newLength As Integer
    
    On Error Resume Next
    
    newLength = GetShortPathName(oActiveDoc.FullName, newShortName, 128)
    If newLength = 0 Then
        idmGetShortPathName = oActiveDoc.FullName
    Else
        idmGetShortPathName = Left(newShortName, newLength%)
    End If
    
End Function

Function ShowToolbar(oAppl As Object, iApplType As Integer)
'*************************************************************
'* Function: showToolBar
'*************************************************************
'* Description: Creates and displays the custom FileNET toolbar
'* in an Office app.  If the toolbar already exists, it gets
'* deleted, then created again.
'*
'* After the toolbar gets created, it adds numerous buttons that perform
'* standard AppInt functions.  I came up with a special trick so we could
'* use custom bitmaps on the toolbar: I created a function called copyPicToClipboard
'* which takes a bitmap name, looks it up in the resource file, and copies the bitmap
'* to the clipboard.  Finally, the clipboard contents are copied to the buttonface
'* with the Pasteface method.
'*
'* The set of buttons is hardwired since the spec doesn't call for any customization.
'*
'* Input: oAppl which is the Application object from Word
'*
'* Output: none.
'*
'*************************************************************
    
    Dim cb As CommandBar
    Dim graphBtn As CommandBarButton
    Dim strName As String
    Dim bExists As Boolean
    Dim IconID As Long
   
    On Error Resume Next
    
    'Set graphBtn = oAppl.CommandBars.FindControl(Tag:=LoadResString(MSG_OPEN))
    'If Not (graphBtn Is Nothing) Then
    '   GoTo Done
    'End If
    'DTS134785 - Raja - 12/07/04
    Call BkpClpBrddata
    strName = LoadResString(MSG_FILENET)
    bExists = False
    
    Set cb = oAppl.CommandBars(strName)
    If Not (cb Is Nothing) Then
       Set FnBtnAdd = cb.FindControl(Tag:=LoadResString(MSG_ADD))
       Set FnBtnCheckin = cb.FindControl(Tag:=LoadResString(MSG_CHECKIN))
       Set FnBtnCancelCheckout = cb.FindControl(Tag:=LoadResString(MSG_CANCEL_CHECKOUT))
       Set FnBtnSave = cb.FindControl(Tag:=LoadResString(MSG_SAVE))
       Set FnBtnShowProperty = cb.FindControl(Tag:=LoadResString(MSG_SHOW_PROP))
       Set FnBtnUpdateProperty = cb.FindControl(Tag:=LoadResString(MSG_UPDATE_PROPERTIES))
       Set FnBtnInsertProperty = cb.FindControl(Tag:=LoadResString(MSG_INSERT_PROPERTIES))
       GoTo Done
    End If
            
    Set cb = oAppl.CommandBars.Add(Temporary:=True)
    With cb
        .Name = strName
        '.Visible = True
        .Position = msoBarFloating ' msoBarTop
    End With
    
    On Error GoTo errHandler
   
    Set graphBtn = cb.Controls.Add(Type:=msoControlButton, Temporary:=True)
    With graphBtn
        .OnAction = "idmFileOpen"
        .ToolTipText = LoadResString(MSG_OPEN)
        .Tag = LoadResString(MSG_OPEN)
    End With
    Call LoadMsoDIBToClipboard(2, oAppl)
    IconID = 2
    graphBtn.PasteFace
    Set FnBtnAdd = graphBtn
    
    Set graphBtn = cb.Controls.Add(Type:=msoControlButton, Temporary:=True)
    With graphBtn
        .OnAction = "idmFileAdd"
        .ToolTipText = LoadResString(MSG_ADD)
        .Tag = LoadResString(MSG_ADD)
    End With
    Call LoadMsoDIBToClipboard(1, oAppl)
    IconID = 1
    graphBtn.PasteFace
    Set FnBtnAdd = graphBtn
    
    Set graphBtn = cb.Controls.Add(Type:=msoControlButton, Temporary:=True)
    With graphBtn
        .OnAction = "idmFileCheckin"
        .ToolTipText = LoadResString(MSG_CHECKIN)
        .Tag = LoadResString(MSG_CHECKIN)
    End With
    Call LoadMsoDIBToClipboard(3, oAppl)
    IconID = 3
    graphBtn.PasteFace
    Set FnBtnCheckin = graphBtn
    
    Set graphBtn = cb.Controls.Add(Type:=msoControlButton, Temporary:=True)
    With graphBtn
        .OnAction = "idmFileCancelCheckout"
        .ToolTipText = LoadResString(MSG_CANCEL_CHECKOUT)
        .Tag = LoadResString(MSG_CANCEL_CHECKOUT)
    End With
    
    Call LoadMsoDIBToClipboard(4, oAppl)
    IconID = 4
    graphBtn.PasteFace
    Set FnBtnCancelCheckout = graphBtn

    Set graphBtn = cb.Controls.Add(Type:=msoControlButton, Temporary:=True)
    With graphBtn
        .OnAction = "idmFileSave"
        .ToolTipText = LoadResString(MSG_SAVE)
        .Tag = LoadResString(MSG_SAVE)
    End With

    Call LoadMsoDIBToClipboard(5, oAppl)
    IconID = 5
    graphBtn.PasteFace
    Set FnBtnSave = graphBtn
    
    'power point does not have any mezz properties
    If iApplType = APPL_WORD Or iApplType = APPL_EXCEL Then
        Set graphBtn = cb.Controls.Add(Type:=msoControlButton, Temporary:=True)
        With graphBtn
            .OnAction = "idmInsertIDMProperty"
            .ToolTipText = LoadResString(MNU_INSERT_MEZZ_PROP) ' LoadResString(MSG_INSERT_PROPERTIES)
            .Tag = LoadResString(MSG_INSERT_PROPERTIES)
        End With

        Call LoadMsoDIBToClipboard(6, oAppl)
        IconID = 6
        graphBtn.PasteFace
        Set FnBtnInsertProperty = graphBtn
        
        Set graphBtn = cb.Controls.Add(Type:=msoControlButton, Temporary:=True)
        With graphBtn
            .OnAction = "idmUpdateIDMProperty"
            .ToolTipText = LoadResString(MNU_UPDATE_MEZZ_PROP) ' LoadResString(MSG_UPDATE_PROPERTIES)
            .Tag = LoadResString(MSG_UPDATE_PROPERTIES)
        End With

         Call LoadMsoDIBToClipboard(7, oAppl)
         IconID = 7
         graphBtn.PasteFace
         Set FnBtnUpdateProperty = graphBtn
      
    End If
    
    Set graphBtn = cb.Controls.Add(Type:=msoControlButton, Temporary:=True)
    With graphBtn
        .OnAction = "idmProperties"
        .ToolTipText = LoadResString(MSG_SHOW_PROPERTY_MGR)
        .Tag = LoadResString(MSG_SHOW_PROP)
    End With

    Call LoadMsoDIBToClipboard(8, oAppl)
    IconID = 8
    graphBtn.PasteFace
    Set FnBtnShowProperty = graphBtn
        
    If cb.Position <> msoBarTop Then
       cb.Position = msoBarTop
    End If

Done:
    cb.Visible = True
    Set cb = Nothing
    Set graphBtn = Nothing
    Call ClearClipboard
    'DTS134785 - Raja - 12/07/04
    Call RestoreClpBrddata
    Exit Function

errHandler:
    If Err.Number = -2147467259 Then
        DoEvents
        Call LoadMsoDIBToClipboard(IconID, oAppl)
        graphBtn.PasteFace
        Resume Next
    Else
       MsgBox Err.Description, vbCritical, LoadResString(MSG_SHOW_TOOLBAR)
    End If
End Function
Function copyPicToClipboard(strPic As String) As Boolean
'*************************************************************
'* Function: copyPicToClipboard
'*************************************************************
'* Description: Copies a bitmap out of the resource file to the clipboard
'*
'* This function is needed in order to be able to use custom icons on the
'* FileNET custom toolbar in Office.  Ordinarily, you're only able to choose
'* from the standard MSFT bitmaps.  But with this function you can copy a
'* bitmap from the project's resource file onto the clipboard, and from there
'* use the CommandBarButton.PasteFace method to paste from the clipboard onto
'* the toolbar.
'*
'* Input: string name of bitmap in the resource file, in this case use idmopen:
'*          idmopen     BITMAP idmopen.bmp
'*
'* Output: Boolean indicating success/failure
'*
'*************************************************************
    Dim ipdTemp As IPictureDisp
    
    On Error GoTo errHandler

    Set ipdTemp = LoadResPicture(strPic, vbResBitmap)

    Clipboard.SetData ipdTemp, vbCFBitmap
    
    copyPicToClipboard = True
    Exit Function
    
errHandler:
    MsgBox Err.Description, vbCritical, LoadResString(MSG_COPY_TO_CLIPBOARD)
    copyPicToClipboard = False

End Function
Public Sub ClearClipboard()
    Clipboard.Clear
End Sub
Public Function GetPropLabel(oDoc As IDMObjects.Document, sPropName As String) As String
'------------------------------------------------------------
'Purpose:Returns the label of a property
'Inputs:PropName - String representing the Property Object Name
'       sFileName - String representing the currently active document
'Outputs: returns the string representing the Property Object label
'Assumptions:
'Constraints
'Copyright © 1998 FileNET Corporation
'------------------------------------------------------------
    Dim oProp As IDMObjects.Property
    Dim oErrorMgr As ErrorManager

    On Error GoTo errHandler
    GetPropLabel = ""
    
    Set oErrorMgr = CreateObject(CREATE_ERR_MGR)
    If oErrorMgr Is Nothing Then
        Err.Raise 1, OPTIONS_DIALOG_SHOW_PROPS, MSG_CANNOT_CREATE_ERRMGR
    End If
    
    Set oProp = oDoc.GetExtendedProperty(sPropName)
    If oProp Is Nothing Then
       GetPropLabel = ""
       Exit Function
    End If
    If IsEmpty(oProp) Then
        GetPropLabel = ""
        Exit Function
    End If
    
    GetPropLabel = oProp.Label
        
    Set oProp = Nothing
    
    Exit Function
    
errHandler:
    If Err.Number <> 0 And Err.Number <> -2147216381 And Err.Number <> -2147157420 Then
        MsgBox Err.Description, vbCritical, LoadResString(DLG_ERR_GET_OBJECT)
    ElseIf Err.Number = -2147157420 Then
        Resume Next
    Else
        If oErrorMgr Is Nothing Then
            MsgBox MSG_ERROR_WITHOUT_ERRMGR, vbCritical, LoadResString(MSG_GET_PROP_LABEL)
            Exit Function
        Else
            If oErrorMgr.Errors.Count > 0 Then
                oErrorMgr.ShowErrorDialog
            End If
        End If
    End If
End Function
Function ShowPreferences(oAppl As Object, iApplType As Integer)

    Call ShowPreferences1
    If (iApplType <> APPL_OUTLOOK) And (iApplType <> APPL_WORDPRO) Then
        If CBool(GetPreferenceValue(LoadResString(STR_SHOW_TOOLBAR), iApplType)) Then
            Call ShowToolbar(oAppl, iApplType)
        Else
            Call HideToolbar(oAppl, iApplType)
        End If
    End If

End Function
Function HideToolbar(oAppl As Object, iApplType As Integer)

'*************************************************************
'* Function: hideToolBar
'*************************************************************
'*
'* Description: Deletes custom FileNET toolbar
'*
'* The set of buttons is hardwired since the spec doesn't call for any customization.
'*
'* Input: oAppl which is the Application object from Word
'*
'* Output: none.
'*
'*************************************************************

    Dim cb As CommandBar
    Dim strName As String
    Dim bExists As Boolean
    Dim graphBtn As CommandBarButton
    
    On Error GoTo errHandler
    
    Set graphBtn = oAppl.CommandBars.FindControl(Tag:=LoadResString(MSG_OPEN))
    If (graphBtn Is Nothing) Then
       GoTo Done:
    End If
    strName = LoadResString(MSG_FILENET)
    bExists = False
   
    For Each cb In oAppl.CommandBars
        If cb.Name = strName Then
            bExists = True
            'EXIT THE LOOP!!!
            Exit For
        End If
    Next cb
    
    If bExists = True Then
        cb.Visible = False
        cb.Delete
    End If

    Set cb = Nothing
Done:
   
    Exit Function
    
    
errHandler:
    MsgBox Err.Description, vbCritical, LoadResString(MSG_SHOW_TOOLBAR)

End Function

Function modifyControls(oAppl As Object, iApplType As Integer, cb As CommandBar, ByVal iMenuTag As Integer, ByVal iControlTag As Integer, ByVal bEnabled As Boolean) ', ByVal bShowToolbar)
    
    Dim cbc As CommandBarControl
    On Error GoTo errHandler

    'update menuitem
    Set cbc = oAppl.CommandBars.FindControl(Tag:=LoadResString(iMenuTag))
    If Not (cbc Is Nothing) Then
        cbc.Enabled = bEnabled
    End If
    
    'update toolbar
    If CBool(GetPreferenceValue(LoadResString(STR_SHOW_TOOLBAR), iApplType)) = False Then
    'If (bShowToolbar = False) Then
       GoTo Done
    End If
    
    If iControlTag = 0 Then GoTo Done
    
    Set cbc = cb.FindControl(Tag:=LoadResString(iControlTag))
    If Not (cbc Is Nothing) Then
        cbc.Enabled = bEnabled
    End If
   
    Exit Function
    
errHandler:
    If Err.Number = 91 Then
       Resume Next
    Else
       MsgBox Err.Description, vbCritical, LoadResString(STR_MODIFY_CONTROLS) ' "Modify Controls"
    End If
Done:
End Function

Public Function GetActiveFileName(iApplType As Integer, oAppl As Object) As String
    'Use the Name property to return the file name without the path
    'use the FullName property to return the file name and the path together.
    On Error GoTo errHandler
    Select Case iApplType
           Case APPL_WORD:
                GetActiveFileName = oAppl.ActiveDocument.FullName
           Case APPL_EXCEL:
                GetActiveFileName = oAppl.ActiveWorkbook.FullName
           Case APPL_POWERPOINT:
         
           Case Else
    End Select
    Exit Function
errHandler:
    MsgBox Err.Description, vbCritical, LoadResString(MSG_GET_ACTIVE_FILE)
End Function
Public Sub CheckPropMgrStatus()
    On Error GoTo errHandler
    
    If gbPropMgrStatus = True Then
       Unload frmPropertyMgr
    End If
    
    Exit Sub
    
errHandler:
    MsgBox Err.Description, vbCritical, LoadResString(MSG_CHECK_PROP_MGR)
End Sub

Public Function saveTemplate(oAppl As Object)
    Dim wrdTemplate As Template
    Dim sFilePath As String
    Dim lReturn As Long
    Dim iReturn As Integer
    
    On Error Resume Next
    If (InStr(1, oAppl.Version, "8", vbTextCompare)) Then
        For Each wrdTemplate In oAppl.Templates
            With wrdTemplate
                sFilePath = .FullName
                iReturn = isFileInUse(sFilePath)
                If (iReturn <> -1) Then
                    lReturn = GetFileAttributes(sFilePath)
                    '*******Raja - DTS134075 - 10/27/04******
                    If (lReturn <> 33 And lReturn <> -1) Then  ' FILE_ATTRIBUTE_READONLY Then
                        'If the user templates are configured to some network drive and is already open then we would not be able to save
                        'the customizations made to the toolbar.
                        If (InStr(1, sFilePath, "\\") = 0) And Not (InStr(1, sFilePath, "Microsoft") = 0) Then
                            If .Saved = False And .Name <> LoadResString(IDM_WORD_TEMPLATE_FILENAME) Then
                                .Saved = True
                            End If
                        End If
                    End If
                End If
            End With
        Next wrdTemplate
    Else
        sFilePath = NormalTemplate.FullName
        iReturn = isFileInUse(sFilePath)
        If (iReturn <> -1) Then
            lReturn = GetFileAttributes(sFilePath)
            '*******Raja - DTS134075 - 10/27/04******
            If (lReturn <> 33 And lReturn <> -1) Then   ' FILE_ATTRIBUTE_READONLY Then
                'If the user templates are configured to some network drive and is already open then we would not be able to save
                'the customizations made to the toolbar.
                If (InStr(1, sFilePath, "\\") = 0) And Not (InStr(1, sFilePath, "Microsoft") = 0) Then
                    NormalTemplate.Save
                Else
                    NormalTemplate.Saved = True
                End If
            Else
                NormalTemplate.Saved = True
            End If
        Else
            NormalTemplate.Saved = True
        End If
    End If
End Function
Function isFileInUse(ByVal sFileName As String) As Integer
  'If the file is already opened by another process and the specified type of access
  'is not allowed then Open operation fails and an error occurs.
  On Error Resume Next
  Dim nFileNum As Integer
  
  nFileNum = FreeFile()
  Open sFileName For Binary Access Write As nFileNum
  Close nFileNum
  'If an error occurs the file is already open or we have run out of File handles or the file does not exist
  isFileInUse = (Err > 0)
End Function

Public Sub ShowPreferences1()

    Dim oPM As New IDMPreferences.Manager
    Dim oSubSystem As IDMPreferences.SubSystem
    Dim oCat As IDMPreferences.Category
    Dim PrefPath As String
    
    oPM.UserType = idmPoUserCurrent
    Set oSubSystem = New IDMPreferences.SubSystem
    oSubSystem.Name = LoadResString(STR_APP_INTEGRATION) '"ApplicationIntegration"
    oSubSystem.UserType = idmPoUserCurrent
    Call oPM.Add(oSubSystem)
    PrefPath = oSubSystem.PathName
    Set oSubSystem = Nothing
    
    oPM.ShowPreferences PrefPath
    
End Sub

Public Function GetPreferenceValue(PreferenceName As String, ApplType As Integer) As Variant

    Dim varPrefValue As Variant
    Dim sKeyName As String
    Dim sValueName As String
    
    sKeyName = LoadResString(STR_REG_PREF_PATH) '"Software\FileNET\IDM\Preferences"
    sValueName = LoadResString(STR_VALUE)       '"Value"
    
    On Error Resume Next
    'On Error GoTo ErrorHandler
    
    Select Case ApplType
            Case APPL_WORD:
                sKeyName = sKeyName & "\" & LoadResString(STR_APP_INTEGRATION)    ' "ApplicationIntegration"
                sKeyName = sKeyName & "\" & LoadResString(STR_WORD_INTEGRATION)   ' "WordIntegration"
           Case APPL_EXCEL:
                sKeyName = sKeyName & "\" & LoadResString(STR_APP_INTEGRATION)    ' "ApplicationIntegration"
                sKeyName = sKeyName & "\" & LoadResString(STR_EXCEL_INTEGRATION)  ' "ExcelIntegration"
           Case APPL_POWERPOINT:
                sKeyName = sKeyName & "\" & LoadResString(STR_APP_INTEGRATION)    ' "ApplicationIntegration"
                sKeyName = sKeyName & "\" & LoadResString(STR_PP_INTEGRATION)     ' "PowerPointIntegration"
           Case APPL_OUTLOOK
                sKeyName = sKeyName & "\" & LoadResString(STR_APP_INTEGRATION)     ' "ApplicationIntegration"
                sKeyName = sKeyName & "\" & LoadResString(STR_OUTLOOK_INTEGRATION) ' "OutlookIntegration"
           Case DIRECTORIES_FILES
                sKeyName = sKeyName & "\" & LoadResString(STR_DIRECTORIES_AND_FILES)
                sKeyName = sKeyName & "\" & LoadResString(STR_LOCAL_CACHING)
           Case Else
                sKeyName = sKeyName & "\" & LoadResString(STR_DOCUMENT)             ' "Documents"
                sKeyName = sKeyName & "\" & LoadResString(STR_ADD_CHECKIN_RETRIEVE) ' "AddCheckinRetrieve"
    End Select
    
    sKeyName = sKeyName & "\" & PreferenceName
    
    varPrefValue = QueryValue(sKeyName, sValueName)
    If (varPrefValue = "") Or (varPrefValue = Empty) Then
        varPrefValue = 0
    End If
    GetPreferenceValue = UCase(varPrefValue)
    
    'Exit Function
    
'ErrorHandler:
    'MsgBox Err.Description, vbCritical, LoadResString(DLG_ERROR_TITLE)
End Function
Public Function GetDirectory_new(DirType As String) As Variant

    Dim varPrefValue As Variant
    Dim sKeyName As String
    Dim sValueName As String
    
    sKeyName = "Software\FileNet\IDM\Preferences"
    sValueName = "Value"
    
    On Error GoTo ErrorHandler
   
    'Initializes the category object
    sKeyName = sKeyName & "\" & LoadResString(STR_DIRECTORIES_AND_FILES) ' "DirectoriesAndFiles" 'SubSystemName
    sKeyName = sKeyName & "\" & LoadResString(STR_CHECKOUTS_AND_COPIES)  '("CheckoutsAndCopies")
    sKeyName = sKeyName & "\" & DirType
    
    varPrefValue = QueryValue(sKeyName, sValueName)
    
    GetDirectory_new = FixDefaultDirectory(varPrefValue)
    Exit Function
    
ErrorHandler:
    MsgBox Err.Description, vbCritical, LoadResString(DLG_ERROR_TITLE)
    
End Function
Public Function GetPreferenceValue_old(PreferenceName As String, ApplType As Integer) As Variant
    Dim oSubSystem As New IDMPreferences.SubSystem
    Dim oCategory As New IDMPreferences.Category
    Dim oCategories As New IDMPreferences.Categories
    Dim oPreference As IDMPreferences.Preference
    Dim oValue As IDMPreferences.Option
    Dim varPrefValue As Variant
    Dim CategoryName As String
    
    On Error GoTo ErrorHandler
    Select Case ApplType
            Case APPL_WORD:
                oSubSystem.Name = LoadResString(STR_APP_INTEGRATION) ' "ApplicationIntegration"
                CategoryName = LoadResString(STR_WORD_INTEGRATION)   '"WordIntegration"
           Case APPL_EXCEL:
                oSubSystem.Name = LoadResString(STR_APP_INTEGRATION) ' "ApplicationIntegration"
                CategoryName = LoadResString(STR_EXCEL_INTEGRATION)  '"ExcelIntegration"
           Case APPL_POWERPOINT:
                oSubSystem.Name = LoadResString(STR_APP_INTEGRATION) '"ApplicationIntegration"
                CategoryName = LoadResString(STR_PP_INTEGRATION)     ' "PowerPointIntegration"
           Case APPL_OUTLOOK
                oSubSystem.Name = LoadResString(STR_APP_INTEGRATION) ' "ApplicationIntegration"
                CategoryName = LoadResString(STR_OUTLOOK_INTEGRATION) '"OutlookIntegration"
           Case DIRECTORIES_FILES
                oSubSystem.Name = LoadResString(STR_DIRECTORIES_AND_FILES)
                CategoryName = LoadResString(STR_LOCAL_CACHING)
           Case Else
                oSubSystem.Name = LoadResString(STR_DOCUMENT)         ' "Documents"
                CategoryName = LoadResString(STR_ADD_CHECKIN_RETRIEVE) ' "AddCheckinRetrieve"
    End Select
    'Initializes the category object
    oSubSystem.UserType = idmPoUserCurrent
    Set oCategory = oSubSystem.GetCategory(CategoryName)
    'Gets the preference
    Set oPreference = oCategory.GetPreference(PreferenceName)
    Set oValue = oPreference.Value
    varPrefValue = oValue.Value
    GetPreferenceValue_old = UCase(varPrefValue)
    Exit Function
    
ErrorHandler:
    MsgBox Err.Description, vbCritical, LoadResString(DLG_ERROR_TITLE)
End Function
Public Function GetDirectory(DirType As String) As Variant
    Dim oSubSystem As New IDMPreferences.SubSystem
    Dim oCategory As New IDMPreferences.Category
    Dim oCategories As New IDMPreferences.Categories
    Dim oPreference As IDMPreferences.Preference
    Dim oValue As IDMPreferences.Option
    
    'On Error GoTo ErrorHandler
    On Error Resume Next
   
    'Initializes the category object
    oSubSystem.Name = LoadResString(STR_DIRECTORIES_AND_FILES)                      ' "DirectoriesAndFiles" 'SubSystemName
    oSubSystem.UserType = idmPoUserCurrent
    Set oCategory = oSubSystem.GetCategory(LoadResString(STR_CHECKOUTS_AND_COPIES)) '("CheckoutsAndCopies")
    'Gets the preference
    Set oPreference = oCategory.GetPreference(DirType)
    Set oValue = oPreference.Value
    GetDirectory = FixDefaultDirectory(oValue.Value)
    'Exit Function
    
'ErrorHandler:
    'MsgBox Err.Description, vbCritical, LoadResString(DLG_ERROR_TITLE)
End Function
Public Function GetLocalDBKey() As Variant
    ' the function is not used any more
    Dim oSubSystem As New IDMPreferences.SubSystem
    Dim oCategory As New IDMPreferences.Category
    Dim oCategories As New IDMPreferences.Categories
    Dim oPreference As IDMPreferences.Preference
    Dim oValue As IDMPreferences.Option
    
    On Error Resume Next
    'On Error GoTo ErrorHandler
   
    'Initializes the category object
    oSubSystem.Name = LoadResString(STR_OTHER)  '"Other" 'SubSystemName
    oSubSystem.UserType = idmPoUserCurrent
    Set oCategory = oSubSystem.GetCategory(LoadResString(STR_LOCALDB)) '("LocalDb")
    'Gets the preference
    Set oPreference = oCategory.GetPreference(LoadResString(STR_LOCALDB_KEY)) '("LocalDbKey")
    Set oValue = oPreference.Value
    GetLocalDBKey = oValue.Value
    'Exit Function
    
'ErrorHandler:
    'MsgBox Err.Description, vbCritical, LoadResString(DLG_ERROR_TITLE)
End Function
Public Function GetDocObject(sFileName As String, Optional blogon As Boolean) As IDMObjects.Document
    Dim oLocalFiles As LocalRecords
    Dim oalocalfiles As LocalRecords
    Dim oLocalFile As LocalRecord
    Dim sDocInfo As tDocInfo
    
    Dim oLocalDB As IDMObjects.LocalDb
    Set oLocalDB = New IDMObjects.LocalDb
    
    On Error GoTo errHandler
    
    'using LocalFiles collection to get a LocalFile item
    oLocalDB.LocalFiles.Refresh

    Set oLocalFiles = oLocalDB.LocalFiles
    If oLocalFiles Is Nothing Then
       Exit Function
    End If
    Set oalocalfiles = oLocalFiles.FindByPath(sFileName)    'returns a LocalFiles collection
    Set oLocalFile = oalocalfiles.Item(1)
    sDocInfo.eSystemType = oLocalFile.LibrarySystemType
    'sDocInfo.eSystemType = idmSysTypeDS
    sDocInfo.sLibraryName = oLocalFile.LibraryId
    sDocInfo.vDocId = oLocalFile.ID
        
    Set GetDocObject = ConvertDocObject(sDocInfo, blogon, oLocalFile.User)
        
    Set oalocalfiles = Nothing
    Set oLocalFile = Nothing
    Set oLocalFiles = Nothing
    Exit Function
errHandler:
    If Err.Number = -2147467259 And Err.Description = LoadResString(STR_INDEX_OUT_RANGE) Then
        blogon = True
    Else
         MsgBox Err.Description, vbCritical, LoadResString(MSG_GET_DOC_STATUS)
    End If
    Set oalocalfiles = Nothing
    Set oLocalFile = Nothing
    Set oLocalFiles = Nothing
    Set oLocalDB = Nothing
End Function
Public Sub GetDocInfo(sFileName As String, sLibraryName As String, sDocID As String)
    Dim oLocalFiles As LocalRecords
    Dim oalocalfiles As LocalRecords
    Dim oLocalFile As LocalRecord
    Dim sDocInfo As tDocInfo
    
    Dim oLocalDB As IDMObjects.LocalDb
    Set oLocalDB = New IDMObjects.LocalDb
    'using LocalFiles collection to get a LocalFile item
    oLocalDB.LocalFiles.Refresh

    Set oLocalFiles = oLocalDB.LocalFiles
    If oLocalFiles Is Nothing Then
       Exit Sub
    End If
    Set oalocalfiles = oLocalFiles.FindByPath(sFileName)    'returns a LocalFiles collection
    Set oLocalFile = oalocalfiles.Item(1)
    'sDocInfo.eSystemType = oLocalFile.LibrarySystemType
    sLibraryName = oLocalFile.LibraryId
    sDocID = CStr(oLocalFile.ID)
           
    Set oalocalfiles = Nothing
    Set oLocalFile = Nothing
    Set oLocalFiles = Nothing
    Set oLocalDB = Nothing
End Sub
Public Function GetDocStatus(sFileName As String) As DocStatusEnum
    Dim oLocalFiles As LocalRecords
    Dim oLocalFile As LocalRecord
    Dim oalocalfiles As LocalRecords
    Dim bReturn As Boolean
    
    On Error GoTo errHandler
    'using LocalFiles collection to get a LocalFile item
    Dim oLocalDB As IDMObjects.LocalDb
    Set oLocalDB = New IDMObjects.LocalDb
    
    oLocalDB.LocalFiles.Refresh
    Set oLocalFiles = oLocalDB.LocalFiles
    If oLocalFiles Is Nothing Then
 '      MsgBox LoadResString(MSG_CANNOT_CREATE_LDB), vbCritical, LoadResString(MSG_GET_DOC_STATUS)
       Exit Function
    End If
    '==========================================================================================
    '01/16/03 to handle compDoc cases with UNC or map drive,
    'if yes, we need to convert file path to UNC
    'if no, do not convert file path to UNC
    Dim sUNCPath As String
    Dim iCount As Integer
    Dim bCompDoc As Boolean
    Set oalocalfiles = oLocalFiles.FindByPath(sFileName)    'returns a LocalFiles collection
    iCount = oalocalfiles.Count
    If (iCount = 0) Then
         'check if the document is a compound document
         'bCompDoc = IsCompoundDocument(sUNCPath)
         bCompDoc = False
         bCompDoc = IsCompoundDocument(sFileName)
         If (bCompDoc) Then
            'check unc path
             sUNCPath = ConvertPathToUNC(sFileName)
             If sUNCPath <> "" Then
                  sFileName = sUNCPath
             Else
                GoTo labelelse
             End If
         End If
    End If
    '==========================================================================================
    If (oLocalFiles.FindByPath(sFileName).Count <> 0) Then
        Set oalocalfiles = oLocalFiles.FindByPath(sFileName)    'returns a LocalFiles collection
        Set oLocalFile = oalocalfiles.Item(1)
    
        'using LocalFile property to get document status
        bReturn = CBool(oLocalFile.IsCheckedOut)
        Select Case bReturn
               Case True
                    GetDocStatus = DocCheckedout
               Case False
                    GetDocStatus = DocCopied
        End Select
    Else
labelelse:
       GetDocStatus = docnew
    End If
    Set oalocalfiles = Nothing
    Set oLocalFile = Nothing
    Set oLocalFiles = Nothing
    Set oLocalDB = Nothing
    Exit Function
errHandler:
    If Err.Number = -2147467259 Then
        GoTo labelelse
    End If
    MsgBox Err.Description, vbCritical, LoadResString(MSG_GET_DOC_STATUS)
    Set oalocalfiles = Nothing
    Set oLocalFile = Nothing
    Set oLocalFiles = Nothing
    Set oLocalDB = Nothing
End Function

Public Function ConvertDocObject(sDocInfo As tDocInfo, blogon As Boolean, Optional sUser As String) As IDMObjects.Document
    Dim sLogonUserName As String
    On Error GoTo errHandler
    If golib Is Nothing Then
        Set golib = New IDMObjects.Library
        golib.SystemType = sDocInfo.eSystemType
        golib.Name = sDocInfo.sLibraryName
    ElseIf ((golib.SystemType <> sDocInfo.eSystemType) Or (golib.Name <> sDocInfo.sLibraryName)) Then
        Set golib = Nothing
        Set golib = New IDMObjects.Library
        golib.SystemType = sDocInfo.eSystemType
        golib.Name = sDocInfo.sLibraryName
    End If
    blogon = False
    If (golib.GetState(idmLibraryLoggedOn) = False) Then
        blogon = golib.Logon(, , , idmLogonOptWithUI)
        sLogonUserName = golib.ActiveUser.Name
        
        If ((IsNull(sUser) = False Or IsEmpty(sUser) = False) And sUser <> "") Then
           If sUser <> sLogonUserName Then
              blogon = False
              golib.Logoff
              MsgBox LoadResString(MSG_INSUFFICIENT_RIGHT), vbInformation, LoadResString(MSG_GET_DOC_STATUS)
           End If
        Else
           'to handle empty user in local db 01/18/02 *SS FR24847
           blogon = True
        End If
    Else
        blogon = True
    End If
    If blogon = True Then
        Set ConvertDocObject = golib.GetObject(idmObjTypeDocument, sDocInfo.vDocId)
    Else
        Set ConvertDocObject = Nothing
    End If
    
    Exit Function
errHandler:
    If Err.Number = -2147215821 Then
       blogon = False
       MsgBox Err.Description, vbInformation, LoadResString(MSG_GET_DOC_STATUS)
       golib.Logoff
    ElseIf Err.Number = -2147207677 And sLogonUserName = "" Then
       'cancel logon
       Exit Function
    Else
       MsgBox Err.Number & Err.Description, vbCritical, LoadResString(MSG_GET_DOC_STATUS)
    End If
End Function
Public Function ResetDocStatus(sFileName As String) As DocStatusEnum
    Dim oLocalFiles As LocalRecords
    Dim oLocalFile As LocalRecord
    Dim oalocalfiles As LocalRecords
    Dim bReturn As Boolean
    
    On Error GoTo errHandler
    'using LocalFiles collection to get a LocalFile item
    Dim oLocalDB As IDMObjects.LocalDb
    Set oLocalDB = New IDMObjects.LocalDb
    
    oLocalDB.LocalFiles.Refresh
    Set oLocalFiles = oLocalDB.LocalFiles
    If oLocalFiles Is Nothing Then
 '      MsgBox LoadResString(MSG_CANNOT_CREATE_LDB), vbCritical, LoadResString(MSG_GET_DOC_STATUS)
       Exit Function
    End If
    If (oLocalFiles.FindByPath(sFileName).Count <> 0) Then
        Set oalocalfiles = oLocalFiles.FindByPath(sFileName)    'returns a LocalFiles collection
        Set oLocalFile = oalocalfiles.Item(1)
    
        'change status to false
        oLocalFile.IsCheckedOut = False   'copy
    End If
    Set oalocalfiles = Nothing
    Set oLocalFile = Nothing
    Set oLocalFiles = Nothing
    Set oLocalDB = Nothing
    Exit Function
errHandler:
    MsgBox Err.Number & " " & Err.Description, vbCritical, LoadResString(MSG_GET_DOC_STATUS)
    Set oalocalfiles = Nothing
    Set oLocalFile = Nothing
    Set oLocalFiles = Nothing
    Set oLocalDB = Nothing
End Function
Public Function GetTitle(vTitle As Variant) As Variant
    Dim iPos As Integer
    iPos = InStrRev(vTitle, ".", -1) ' InStr(vTitle, ".")
    If iPos = 0 Then
       GetTitle = vTitle
    Else
       GetTitle = Left(vTitle, (iPos - 1))
    End If
End Function
Public Function FixDefaultDirectory(vDir As Variant) As Variant
    Dim iLen As Integer
    If Right(vDir, 1) = "\" Then
       iLen = Len(vDir)
       FixDefaultDirectory = Left(vDir, (iLen - 1))
    Else
       FixDefaultDirectory = vDir
    End If
End Function
Public Function IDMUpdateMenu(iApplType As Integer, oAppl As Object, Optional vLastDocClosed As Variant)
    Dim oErrorMgr As ErrorManager
    Dim bExists As Boolean
    Dim iDocCount As Integer
    Dim bGetPrefResults As Boolean
    Dim iInStrResults As Integer
    Dim DocStatus As DocStatusEnum
    Dim sCaption As String
    Dim sFile As String
    Dim bNoContainer As Boolean
    Dim bShowToolbar As Boolean
    
    bNoContainer = False
    
    On Error Resume Next
    
    'If iApplType = APPL_POWERPOINT Then
    '    Exit Function
    'End If
    
    bShowToolbar = CBool(GetPreferenceValue(LoadResString(STR_SHOW_TOOLBAR), iApplType))
    
    bIsUpdateMenuToolbar = CBool(GetPreferenceValue(LoadResString(STR_UPDATE_MENU_TOOLBAR), iApplType))
    If bIsUpdateMenuToolbar = False Then
        Call ResetFnMenuAndToolbar(True, True, True, True, True, True, True, True, bShowToolbar)
        Call ResetMSSaveAsMenuItem(oAppl, iApplType, True)
        Call Resetoffice97Save(oAppl, MB_FILE, MB_FILE_SAVE, True)
        GoTo Done
    End If
    
    On Error GoTo errHandler
    
    If DocContained(oAppl, iApplType) <> NoContainer Then
        If IsMissing(vLastDocClosed) Then
          'to handle the container caused the error for the file in the temp folder
           bNoContainer = True
        End If
        'vLastDocClosed = True
    End If

    Set oErrorMgr = CreateObject(CREATE_ERR_MGR)
    If oErrorMgr Is Nothing Then
        Err.Raise 1, LoadResString(IDM_UPDATE_MENU), LoadResString(MSG_CANNOT_CREATE_ERRMGR)
    End If

    iDocCount = DocCount(oAppl, iApplType)
  
    'this case for Excel close all workbooks
    If iApplType = APPL_EXCEL Then
        If Not IsMissing(vLastDocClosed) Then
            If oAppl.Windows.Count <= 1 Then
                If vLastDocClosed = True Then iDocCount = 0
            End If
        End If
    End If
    
    If iDocCount = 0 Then
            Call ResetFnMenuAndToolbar(False, False, False, False, False, False, False, False, bShowToolbar)
                 
            If InStr(sFile, LoadResString(STR_CACHE)) > 0 Then 'And vAttr = ReadOnly Then
               Call ResetMSSaveAsMenuItem(oAppl, iApplType, True)
               Call Resetoffice97Save(oAppl, MB_FILE, MB_FILE_SAVE, True)
            Else
               If bNoContainer = True Then
                  Call ResetMSSaveAsMenuItem(oAppl, iApplType, True)
                  Call Resetoffice97Save(oAppl, MB_FILE, MB_FILE_SAVE, True)
               Else
                  Call ResetMSSaveAsMenuItem(oAppl, iApplType, False)
                  Call Resetoffice97Save(oAppl, MB_FILE, MB_FILE_SAVE, False)
               End If
               
            End If
        
            GoTo Done
    End If
    
    Call ResetMSSaveAsMenuItem(oAppl, iApplType, True)
    Call Resetoffice97Save(oAppl, MB_FILE, MB_FILE_SAVE, True)
   
    DocStatus = GetDocStatus(getFullName(oAppl, iApplType))
    If DocStatus <> DocCheckedout Then
        Call ResetFnMenuAndToolbar(True, False, False, True, False, True, False, True, bShowToolbar)
        
        If bRename = True Then
           oAppl.ActiveWindow.Caption = ResetCaption(oAppl, iApplType, 0)
           bRename = False
           GoTo Done
        End If
        'check preference and show doc status
        If iApplType <> APPL_POWERPOINT Then
            bGetPrefResults = CBool(GetPreferenceValue(LoadResString(STR_SHOW_DOC_STATUS), iApplType))
            iInStrResults = InStr(oAppl.ActiveWindow.Caption, LoadResString(STR_COPIED))
            
            If (bGetPrefResults = True And DocStatus = DocCopied) Then
                If (iInStrResults = 0) Then
                    oAppl.ActiveWindow.Caption = ResetCaption(oAppl, iApplType, Len(LoadResString(STR_COPIED)) + 1) & " " & LoadResString(STR_COPIED)
                End If
            ElseIf bGetPrefResults = False And iInStrResults <> 0 Then
                oAppl.ActiveWindow.Caption = ResetCaption(oAppl, iApplType, 0)
            Else
                oAppl.ActiveWindow.Caption = ResetCaption(oAppl, iApplType, 0)
            End If
        End If
    Else
        Call ResetFnMenuAndToolbar(False, True, True, True, True, True, True, True, bShowToolbar)
        
        'check preference and show doc status
        If iApplType <> APPL_POWERPOINT Then
            bGetPrefResults = CBool(GetPreferenceValue(LoadResString(STR_SHOW_DOC_STATUS), iApplType))
            iInStrResults = InStr(oAppl.ActiveWindow.Caption, LoadResString(STR_CHECKEDOUT))
            If (bGetPrefResults = True And iInStrResults = 0) Then
                oAppl.ActiveWindow.Caption = ResetCaption(oAppl, iApplType, Len(LoadResString(STR_CHECKEDOUT)) + 1) & " " & LoadResString(STR_CHECKEDOUT)
            ElseIf (bGetPrefResults = False And iInStrResults <> 0) Then
                oAppl.ActiveWindow.Caption = ResetCaption(oAppl, iApplType, 0)
            End If
        End If
    End If
    GoTo Done

errHandler:
    ' Check for Automation Error from TrackFile->GetObject when user hits Cancel on Logon dialog
    ' and don't display any errors, just finish.
    If Err.Number <> 0 And Err.Number <> -2147216381 And Err.Number <> 4248 Then
        If Err.Number = 91 Or Err.Number = 1004 Then
           'suppress the error message for localization
        Else
            MsgBox Err.Description, vbCritical, LoadResString(IDM_UPDATE_MENU)
        End If
    Else
        If oErrorMgr Is Nothing Then
            MsgBox LoadResString(MSG_ERROR_WITHOUT_ERRMGR), vbCritical, LoadResString(IDM_UPDATE_MENU) '"FileNET IDMUpdateMenu"
            GoTo Done
        Else
            If oErrorMgr.Errors.Count > 0 Then
                oErrorMgr.ShowErrorDialog
            End If
        End If
    End If

Done:
    If iApplType = APPL_WORD Then
        'since we have just changed normal.dot we need to reset .saved
        'so will not ask us to save to normal.dot when we exit
        Call saveTemplate(oAppl)
    End If
    
End Function

Public Function ResetCaption(oAppl As Object, iApplType As Integer, Optional iLen As Integer) As String
   Dim sCaption As String
   Dim iDiff As Integer
   sCaption = getName(oAppl, iApplType)
   If iApplType = 2 Then 'Excel
      If oAppl.ActiveWorkbook.ReadOnly Then
          sCaption = sCaption & "  " & LoadResString(READ_ONLY) '"[Read-Only]"
      End If
   End If
   iDiff = Len(sCaption) + iLen - 200
   If iApplType = 2 Then 'Excel
        If iDiff > 0 Then
           ResetCaption = Left(sCaption, 200 - iLen - 3) & "..."
        Else
           ResetCaption = sCaption
        End If
   Else
        ResetCaption = sCaption
   End If
End Function
Public Function SetPreferenceInWizard(sSubSystemName As String, sCategoryName As String, sPreferenceName As String, vVal As Variant) As IDMPreferences.Preference
    Dim oSubSystem As New IDMPreferences.SubSystem
    Dim oCategory As New IDMPreferences.Category
    Dim oPreference As IDMPreferences.Preference
    Dim oOption As IDMPreferences.Option
    
    oSubSystem.Name = sSubSystemName ' LoadResString(STR_DOCUMENT)
    oSubSystem.UserType = idmPoUserCurrent
    Set oCategory = oSubSystem.GetCategory(sCategoryName) 'LoadResString(STR_ADD_CHECKIN_RETRIEVE))
    Set oPreference = oCategory.GetPreference(sPreferenceName)
    Set oOption = oPreference.Value
    oOption.Value = vVal
    Set SetPreferenceInWizard = oPreference
End Function

Public Sub GetHelp(iApplType As Integer)
    Dim lContextnumber As Long
    On Error GoTo ErrorHandler
    Select Case iApplType
        Case APPL_WORD
             lContextnumber = 3040
        Case APPL_EXCEL
             lContextnumber = 3040
        Case APPL_POWERPOINT
             lContextnumber = 3040
    End Select
        Call WinHelp(GetFocus(), LoadResString(STR_HELP_FILE), HELP_CONTEXT, lContextnumber)
    Exit Sub
ErrorHandler:
       MsgBox Err.Description, vbCritical, LoadResString(STR_FN_HELP)
End Sub

Public Function PowerPointLinkExsists(oAppl As Object) As Boolean
    Dim oSlide As PowerPoint.Slide
    Dim oShape As PowerPoint.Shape
    PowerPointLinkExsists = False
    
    For Each oSlide In oAppl.ActivePresentation.Slides
        For Each oShape In oSlide.Shapes
            If oShape.Type = msoLinkedOLEObject Then
                PowerPointLinkExsists = True
            End If
        Next
    Next
End Function
Public Function BlockPowerPoint(oAppl As Object, iApplType As Integer) As Boolean
    Dim oPresentation As PowerPoint.Presentation
    Dim oSlide As PowerPoint.Slide
    Dim oShape As PowerPoint.Shape
    Dim oobject As Excel.OLEObject
    Dim osheet As Excel.Worksheet
    Dim sAppName As String
    BlockPowerPoint = False
    
    If (iApplType = APPL_POWERPOINT) Then
       If (InStr(1, oAppl.Version, "8", vbTextCompare)) Then
            For Each oSlide In oAppl.ActivePresentation.Slides
                For Each oShape In oSlide.Shapes
                    If oShape.Type = msoLinkedOLEObject Then
                        If InStr(1, oShape.LinkFormat.SourceFullName, LoadResString(STR_PPT), vbTextCompare) Then '".ppt"
                            BlockPowerPoint = True
                        End If
                    End If
                Next
            Next
        End If
    ElseIf (iApplType = APPL_EXCEL) Then
        If (InStr(1, oAppl.Version, "8", vbTextCompare)) Then
            For Each osheet In oAppl.ActiveWorkbook.Worksheets
                For Each oobject In osheet.OLEObjects
                    If oobject.OLEType = xlOLELink Then
                        sAppName = Mid(oobject.SourceName, 1, InStr(1, oobject.SourceName, "|", vbTextCompare) - 1)
                        If InStr(1, sAppName, LoadResString(STR_POWERPOINT), vbTextCompare) Then   '"powerpoint"
                            BlockPowerPoint = True
                        End If
                    End If
                Next
            Next
        End If
    End If
End Function
Public Sub GetFileNameAndExt(sFileFullName As String, sFileName As String, sExt As String)
    Dim iPos As Integer
    iPos = InStrRev(sFileFullName, ".", -1)
    If iPos = 0 Then
       sFileName = sFileFullName
    Else
       sFileName = Left(sFileFullName, (iPos - 1))
       sExt = Right(sFileFullName, Len(sFileFullName) - iPos + 1)
    End If
End Sub
Public Sub LoadMsoDIBToClipboard(nResID As Long, Optional oAppl As Object)
   Static cfBtnFace As Long
   Static cfBtnMask As Long
   
   Dim hDIB As Long
   Dim lpData As Long
   Dim cbDIBFaceSize As Long
   Dim cbDIBMaskSize As Long
   
   Dim yResData() As Byte
   Dim iCheck As Integer
   
   On Error Resume Next

 ' Load the resource data and exit if not valid...
   yResData = LoadResData(nResID, LoadResString(STR_BUTTON)) '"Buttons"
   If Err Or UBound(yResData) < 12 Then Exit Sub
   
   CopyMemory iCheck, yResData(0), 2
   If iCheck <> &H1111 Then Exit Sub
   
   CopyMemory cbDIBFaceSize, yResData(2), 4
   CopyMemory cbDIBMaskSize, yResData(6), 4

 ' Open the clipboard...
   If CBool(OpenClipboard(0&)) Then
    ' If we haven't already done so, grab the cf for button face and mask...
      If cfBtnFace = 0 Then
      
         If (InStr(1, oAppl.Version, "8", vbTextCompare)) Then
             cfBtnFace = RegisterClipboardFormat(LoadResString(STR_TOOLBAR_BUTTON_FACE_97)) '"Toolbar Button Face"
             cfBtnMask = RegisterClipboardFormat(LoadResString(STR_TOOLBAR_BUTTON_MASK_97)) '"Toolbar Button Mask"
         ElseIf (InStr(1, oAppl.Version, "9", vbTextCompare)) Then
             cfBtnFace = RegisterClipboardFormat(LoadResString(STR_TOOLBAR_BUTTON_FACE_2000))
             cfBtnMask = RegisterClipboardFormat(LoadResString(STR_TOOLBAR_BUTTON_MASK_2000))
         Else
             cfBtnFace = RegisterClipboardFormat(LoadResString(STR_TOOLBAR_BUTTON_FACE))
             cfBtnMask = RegisterClipboardFormat(LoadResString(STR_TOOLBAR_BUTTON_MASK))
         End If
         
      End If
      
    ' Clear the current contents...
      EmptyClipboard
      
    ' Allocate a buffer and copy the button face as CF_DIB...
      hDIB = GlobalAlloc(GMEM_MOVEABLE Or GMEM_DDESHARE, cbDIBFaceSize)
      If hDIB <> 0 Then
         lpData = GlobalLock(hDIB)
         CopyMemory ByVal lpData, yResData(10), cbDIBFaceSize
         GlobalUnlock hDIB

         If SetClipboardData(CF_DIB, hDIB) = 0 Then
            GlobalFree hDIB
         End If
      End If
    
    ' Allocate a second buffer and copy the button face as cfBtnFace...
      hDIB = GlobalAlloc(GMEM_MOVEABLE Or GMEM_DDESHARE, cbDIBFaceSize)
      If hDIB <> 0 Then
         lpData = GlobalLock(hDIB)
         CopyMemory ByVal lpData, yResData(10), cbDIBFaceSize
         GlobalUnlock hDIB

         If SetClipboardData(cfBtnFace, hDIB) = 0 Then
            GlobalFree hDIB
         End If
      End If

    ' Allocate another buffer and copy the button mask...
      hDIB = GlobalAlloc(GMEM_MOVEABLE Or GMEM_DDESHARE, cbDIBMaskSize)
      If hDIB <> 0 Then
         lpData = GlobalLock(hDIB)
         CopyMemory ByVal lpData, yResData(10 + cbDIBFaceSize), cbDIBMaskSize
         GlobalUnlock hDIB

         If SetClipboardData(cfBtnMask, hDIB) = 0 Then
            GlobalFree hDIB
         End If
      End If
    
      CloseClipboard
   End If

End Sub
Public Function RemoveIllegalChars(sFilePathName As String, sDir As String, sFileName As String, sExt As String) As String
   Dim sNewFileName As String
   Dim sChar As String
   Dim iLen As Integer
   Dim i As Integer
   iLen = Len(sFileName)
   sNewFileName = ""
   For i = 1 To iLen
        sChar = Mid(sFileName, i, 1)
        Select Case sChar
               Case "/", "\", ":", "*", "?", "<", ">", "|", "[", "]"
                    'MsgBox "A file name cannot contain the following character(s) : " & sChar & vbCrLf & sFilePathName & sExt, vbInformation
               'Case "[", "]"
                    'for power point case to keep "[,]"
                    'If iAppl = APPL_POWERPOINT Then
                    '   sNewFileName = sNewFileName & sChar
                    'End If
               Case Else
                    sNewFileName = sNewFileName & sChar
        End Select
   Next i
   sFilePathName = sDir & "\" & sNewFileName
   RemoveIllegalChars = sFilePathName
End Function
Public Function FindIllegalChar(sFileName As String, sExt As String, sIllegalChar As String) As Boolean
   Dim sChar As String
   Dim iLen As Integer
   Dim i As Integer
   
   FindIllegalChar = False
   
   sIllegalChar = ""
   iLen = Len(sFileName)
   For i = 1 To iLen
        sChar = Mid(sFileName, i, 1)
        Select Case sChar
               Case "/", "\", ":", "*", "?", "<", ">", "|", "[", "]"
                    If sIllegalChar <> "" Then
                       sIllegalChar = sIllegalChar & " , " & sChar
                    Else
                       sIllegalChar = sIllegalChar & sChar
                    End If
                    FindIllegalChar = True
        End Select
   Next
End Function

Public Function IsDocSaved(sDocName As String, iApplType As Integer, oAppl As Object) As Boolean
'Description:
' returns true if active doc does not have any unsaved changes
' returns false if active doc has changes that need to be saved
    Select Case iApplType
        Case APPL_WORD
            IsDocSaved = oAppl.Documents(sDocName).Saved
        Case APPL_EXCEL
            IsDocSaved = oAppl.Workbooks(sDocName).Saved
        Case APPL_POWERPOINT
            IsDocSaved = oAppl.Presentations(sDocName).Saved
        Case Else
            IsDocSaved = False
    End Select

End Function
Public Sub SaveDoc(sDocName As String, iApplType As Integer, oAppl As Object)
    
    On Error Resume Next
    
    Select Case iApplType
        Case APPL_WORD
             
             If oAppl.Documents(sDocName).SaveFormat = 18 Then
                'Word 6.0/ 95 convert to word 10 or current doc format
                If CBool(GetPreferenceValue(LoadResString(STR_CONVERT_WORD6_95_TO_CURRENT_WORD_VERSION), iApplType)) = True Then
                   oAppl.Documents(sDocName).SaveAs FileFormat:=wdFormatDocument
                End If
             Else
                oAppl.Documents(sDocName).Save
             End If
             
        Case APPL_EXCEL
            oAppl.Workbooks(sDocName).Save
        Case APPL_POWERPOINT
            oAppl.Presentations(sDocName).Save
    End Select
End Sub
Public Sub CloseDoc(sDocName As String, iApplType As Integer, oAppl As Object)
 
 On Error Resume Next
 
 Select Case iApplType
        Case APPL_WORD
            oAppl.Documents(sDocName).Close
        Case APPL_EXCEL
            oAppl.Workbooks(sDocName).Close
        Case APPL_POWERPOINT
            oAppl.Presentations(sDocName).Close
  End Select

End Sub

Public Function ProcessFilterList(ext As String) As String
    Dim strExtName As String
    Dim retvalue As Long
    Dim hkeyOpenKey As Long
    Dim lType As Long
    Dim strKeytoExtDesc As String
    Dim strExtDesc As String
    Dim llen As Long
    Dim strFilter As String
    
    On Error Resume Next
    
    llen = 511
    strKeytoExtDesc = String(llen + 1, " ")
    strExtDesc = String(llen + 1, " ")
    strFilter = ""
    strExtName = ext
    Call VBA.LCase(strExtName)
    
    retvalue = RegOpenKeyEx(HKEY_CLASSES_ROOT, strExtName, 0, KEY_READ, hkeyOpenKey)
    If (retvalue = ERROR_SUCCESS) Then
        retvalue = RegQueryValueExString(hkeyOpenKey, "", 0, lType, strKeytoExtDesc, llen)
        RegCloseKey (hkeyOpenKey)
        hkeyOpenKey = 0
        strKeytoExtDesc = VBA.RTrim(strKeytoExtDesc)
        If (retvalue = ERROR_SUCCESS) Then
            retvalue = RegOpenKeyEx(HKEY_CLASSES_ROOT, strKeytoExtDesc, 0, KEY_READ, hkeyOpenKey)
            If (retvalue = ERROR_SUCCESS) Then
                llen = 511
                retvalue = RegQueryValueExString(hkeyOpenKey, "", 0, lType, strExtDesc, llen)
                If (retvalue = ERROR_SUCCESS) Then
                    strExtDesc = VBA.Mid(strExtDesc, 1, llen - 1)
                    strFilter = strExtDesc
                End If
                RegCloseKey (hkeyOpenKey)
                hkeyOpenKey = 0
            End If
        End If
    End If
    If (retvalue = ERROR_SUCCESS) Then
        strFilter = strFilter & " (*"
        strFilter = strFilter & strExtName
        strFilter = strFilter & ")|"
        strFilter = strFilter & "*"
        strFilter = strFilter & strExtName
        strFilter = strFilter & "|"
    End If
    strFilter = strFilter & LoadResString(STR_ALLFILES) '"All Files (*.*)|*.*"
    ProcessFilterList = strFilter
    
End Function
Public Function DSinstalled() As Boolean
    Dim retvalue As Long
    Dim hkeyOpenKey As Long
    Dim lType As Long
    Dim Value As Long
    Dim llen As Long
    Dim strFilter As String
    
    On Error Resume Next
    llen = 4
    retvalue = RegOpenKeyEx(HKEY_LOCAL_MACHINE, LoadResString(STR_SubKey_IDM_Install), 0, KEY_READ, hkeyOpenKey)
    If (retvalue = ERROR_SUCCESS) Then
        retvalue = RegQueryValueExLong(hkeyOpenKey, LoadResString(STR_Mezzanine_Installed), 0, lType, Value, llen)
       ' MsgBox llen
        RegCloseKey (hkeyOpenKey)
       ' MsgBox value
        If Value = 0 Then
            DSinstalled = False
        ElseIf Value = 1 Then
            DSinstalled = True
        End If
    Else
        DSinstalled = True
    End If

End Function

Public Function promptsave(CallingOperation As AddCheckinEnum, FileName As String, iApplType As Integer) As VbMsgBoxResult
    Dim spreference As String
    Dim strInSubroutine As String
    If CallingOperation = idmAdd Then
        spreference = LoadResString(STR_PRT_SAVE_ADD)
        strInSubroutine = LoadResString(MSG_ADD)
    Else
        spreference = LoadResString(STR_PRT_SAVE_CHECKIN)
        strInSubroutine = LoadResString(MSG_CHECKIN)
    End If
    If GetPreferenceValue(spreference, iApplType) = "1" Then
        promptsave = MsgBox(LoadResString(STR_DO_YOU_WANT_TO_SAVE_THE_DOC) & LoadResString(STR_P_LEFT) & FileName & LoadResString(STR_P_RIGHT), vbYesNoCancel + vbQuestion, strInSubroutine)
    Else
        promptsave = vbYes
    End If
End Function
Public Sub TruncateFileName(sFilePath As String, iLen As Integer)
    Dim sDir As String
    Dim sFileName As String
    Dim sStripedFileName As String
    Dim sExt As String
    Dim iL As Integer
    Dim sNewFilePath As String
    Call idmGetDirectoryAndFileName(sFilePath, sDir, sFileName)
    If iL = (Len(sFileName) - iLen) > 0 Then
        Call GetFileNameAndExt(sFileName, sStripedFileName, sExt)
        sFileName = Left(sFileName, Len(sFileName) - iL) & sExt
    End If
    sNewFilePath = sDir & "\" & sFileName
    'rename the old filename
    Name sFilePath As sNewFilePath
    'return a new filepath
    sFilePath = sNewFilePath
End Sub
Public Sub Walklink_old(oAppl As Object, iApplType As Integer, CallingOperation As AddCheckinEnum)
Dim oobject As Excel.OLEObject
Dim osheet As Excel.Worksheet
Dim sAppName As String
Dim sName As String
Dim oApplil   As Object
On Error Resume Next
    If iApplType = APPL_EXCEL Then
        For Each osheet In oAppl.ActiveWorkbook.Worksheets
            For Each oobject In osheet.OLEObjects
            If oobject.OLEType = xlOLELink Then
                sAppName = Mid(oobject.SourceName, 1, InStr(1, oobject.SourceName, "|", vbTextCompare) - 1)
                sName = Mid(oobject.SourceName, InStr(1, oobject.SourceName, "|", vbTextCompare) + 1, InStr(1, oobject.SourceName, "!", vbTextCompare) - InStr(1, oobject.SourceName, "|", vbTextCompare) - 1)
                If InStr(1, sAppName, LoadResString(STR_WORD), vbTextCompare) Then   '"word"
                    Set oApplil = GetObject(, "Word.Application") '"Word.Application"
                    If Err.Number <> 0 Then
                        Err.Clear
                        Set oApplil = GetObject(sName)
                        oApplil.Activate
                        oApplil.Close
                        oApplil.Application.Quit
                    ElseIf oApplil.Visible = False Then
                        Set oApplil = GetObject(sName)
                        oApplil.Activate
                        oApplil.Close
                    Else
                        Set oApplil = GetObject(sName)
                        oApplil.Activate
                        If (oApplil.Saved = False) Then
                            Select Case promptsave(CallingOperation, sName, iApplType)
                            Case vbYes
                                oApplil.Save
                            Case vbNo
                               '    Party On
                            End Select
                        End If
                        oApplil.Close False
                    End If
                    Set oApplil = Nothing
                ElseIf InStr(1, sAppName, LoadResString(STR_POWERPOINT), vbTextCompare) Then '"powerpoint"
                    Set oApplil = GetObject(, "PowerPoint.Application")
                    If Err.Number <> 0 Then
                        Err.Clear
                         Set oApplil = GetObject(sName)
                         oApplil.Close
                         oApplil.Application.Quit
                         Set oApplil = Nothing
                    Else
                         Set oApplil = GetObject(sName)
                         If (oApplil.Saved = False) Then
                            Select Case promptsave(CallingOperation, sName, iApplType)
                            Case vbYes
                                oApplil.Save
                            Case vbNo
                            End Select
                         End If
                         oApplil.Close
                    End If
                    Set oApplil = Nothing
                End If
            End If
            Next
        Next
    End If
End Sub
Public Sub Walklink(oAppl As Object, iApplType As Integer, CallingOperation As AddCheckinEnum)
Dim oApplil   As Object
Dim aLinks As Variant
Dim i As Integer
Dim sName As String

If iApplType = APPL_EXCEL Then
    'paste link and object link
    aLinks = oAppl.ActiveWorkbook.LinkSources(xlOLELinks)
    If Not IsEmpty(aLinks) Then
            For i = 1 To UBound(aLinks)
                sName = CStr(aLinks(i))
                Call CloseChildDocument(sName, oAppl, iApplType, CallingOperation)
           Next i
     End If
End If
End Sub
Public Sub CloseChildDocument(sLinkString As String, oAppl As Object, iApplType As Integer, CallingOperation As AddCheckinEnum)
Dim oobject As Excel.OLEObject
Dim osheet As Excel.Worksheet
Dim sAppName As String
Dim sName As String
Dim oApplil  As Object
               sAppName = Mid(sLinkString, 1, InStr(1, sLinkString, "|", vbTextCompare) - 1)
               sName = Mid(sLinkString, InStr(1, sLinkString, "|", vbTextCompare) + 1, InStr(1, sLinkString, "!", vbTextCompare) - InStr(1, sLinkString, "|", vbTextCompare) - 1)
                If InStr(1, sAppName, LoadResString(STR_WORD), vbTextCompare) Then  '"word"
                    Set oApplil = GetObject(, "Word.Application")
                    If Err.Number <> 0 Then
                        Err.Clear
                        Set oApplil = GetObject(sName)
                        oApplil.Activate
                        oApplil.Close
                        oApplil.Application.Quit
                    ElseIf oApplil.Visible = False Then
                        Set oApplil = GetObject(sName)
                        oApplil.Activate
                        oApplil.Close
                    Else
                        Set oApplil = GetObject(sName)
                        oApplil.Activate
                        If (oApplil.Saved = False) Then
                            Select Case promptsave(CallingOperation, sName, iApplType)
                            Case vbYes
                                oApplil.Save
                            Case vbNo
                               '    Party On
                            End Select
                        End If
                        oApplil.Close False
                    End If
                    Set oApplil = Nothing
                ElseIf InStr(1, sAppName, LoadResString(STR_POWERPOINT), vbTextCompare) Then '"powerpoint"
                    Set oApplil = GetObject(, "PowerPoint.Application")
                    If Err.Number <> 0 Then
                        Err.Clear
                         Set oApplil = GetObject(sName)
                         oApplil.Close
                         oApplil.Application.Quit
                         Set oApplil = Nothing
                    Else
                         Set oApplil = GetObject(sName)
                         If (oApplil.Saved = False) Then
                            Select Case promptsave(CallingOperation, sName, iApplType)
                            Case vbYes
                                oApplil.Save
                            Case vbNo
                            End Select
                         End If
                         oApplil.Close
                    End If
                    Set oApplil = Nothing
                End If
End Sub
Public Sub ResetMSMenuItem(oAppl As Object, id1 As Variant, id2 As Variant, bStatus As Boolean)
    
    Dim cbcParentMenu As CommandBarControl
    Dim cbcChildMenu As CommandBarControl
    
    On Error Resume Next
    
    Set cbcParentMenu = oAppl.CommandBars.FindControl(ID:=id1)
    For Each cbcChildMenu In cbcParentMenu.Controls
        If (cbcChildMenu.ID = id2) Then
            cbcChildMenu.Enabled = bStatus
        End If
    Next cbcChildMenu
    Set cbcChildMenu = Nothing
    
End Sub

Public Sub ResetMSSaveAsMenuItem(oAppl As Object, iApplType As Integer, bStatus As Boolean)
    Dim oMenu As Object
    
    Set oMenu = oAppObject.CommandBars.FindControl(ID:=MB_FILE_SAVEAS)
    oMenu.Enabled = bStatus
    
    'Select Case iApplType
       'Case APPL_WORD:
       '     oAppl.CommandBars("Menu bar").Controls("File").Controls("Save As...").Enabled = bStatus
       'Case APPL_EXCEL:
       '     oAppl.CommandBars("Worksheet Menu bar").Controls("File").Controls("Save As...").Enabled = bStatus
    'End Select
    
End Sub

Public Sub ResetFnMenuAndToolbar(bAdd As Boolean, bCheckin As Boolean, bCancelCheckout As Boolean, _
bSave As Boolean, bShowProperty As Boolean, bInsertProperety As Boolean, bUpdateProperty As Boolean, _
bInsertFile As Boolean, bShowToolbar As Boolean)
    
    On Error Resume Next
   
    Dim FileMenu As Object
    Dim InsertMenu As Object
    
    
    Select Case giApplType
        Case APPL_WORD:
             'Set FileMenu = oAppObject.CommandBars("Menu bar").Controls("File")
             'Set InsertMenu = oAppObject.CommandBars("Menu bar").Controls("Insert")
             'Set FileMenu = oAppObject.CommandBars.FindControl(ID:=MB_FILE)
             Set InsertMenu = oAppObject.CommandBars.FindControl(ID:=MB_INSERT)
        Case APPL_EXCEL:
             'Set FileMenu = oAppObject.CommandBars("Worksheet Menu bar").Controls("File")
             'Set InsertMenu = oAppObject.CommandBars("Worksheet Menu bar").Controls("Insert")
             'Set FileMenu = oAppObject.CommandBars.FindControl(ID:=MB_FILE)
             Set InsertMenu = oAppObject.CommandBars.FindControl(ID:=MB_INSERT)
        'Case APPL_POWERPOINT:
             'Set FileMenu = oAppObject.CommandBars.FindControl(ID:=MB_FILE)
    End Select
    Set FileMenu = oAppObject.CommandBars.FindControl(ID:=MB_FILE)
    'file menu
    Set FnMenuAdd = FileMenu.Controls(LoadResString(MNU_FN_ADD))
    Set FnMenuFnCheckin = FileMenu.Controls(LoadResString(MNU_FN_CHECKIN))
    Set FnMenuCancelCheckout = FileMenu.Controls(g_FN_CANCEL(giApplType))
    Set FnMenuSave = FileMenu.Controls(LoadResString(MNU_FN_SAVE))
    Set FnMenuShowProperty = FileMenu.Controls(LoadResString(MNU_FN_PROPERTIES))
    
    
    FnMenuAdd.Enabled = bAdd
    FnMenuFnCheckin.Enabled = bCheckin
    FnMenuCancelCheckout.Enabled = bCancelCheckout
    FnMenuSave.Enabled = bSave
    FnMenuShowProperty.Enabled = bShowProperty
    
    'insert menu
    If giApplType <> APPL_POWERPOINT Then
        Set FnMenuInsertProperty = InsertMenu.Controls(LoadResString(MNU_INSERT_MEZZ_PROP))
        Set FnMenuUpdateProperty = InsertMenu.Controls(LoadResString(MNU_UPDATE_MEZZ_PROP))
        
        FnMenuInsertProperty.Enabled = bInsertProperety
        FnMenuUpdateProperty.Enabled = bUpdateProperty
        
        If (giApplType = APPL_WORD) Then
          Set FnMenuInsertFile = InsertMenu.Controls(LoadResString(MNU_INSERT_FILE))
          FnMenuInsertFile.Enabled = bInsertFile
        End If
    End If
    
    'update toolbar
    If (bShowToolbar = True) Then
    
        FnBtnAdd.Enabled = bAdd
        FnBtnCheckin.Enabled = bCheckin
        FnBtnCancelCheckout.Enabled = bCancelCheckout
        FnBtnSave.Enabled = bSave
        
        FnBtnShowProperty.Enabled = bShowProperty
        
        If giApplType <> APPL_POWERPOINT Then
           FnBtnInsertProperty.Enabled = bInsertProperety
           FnBtnUpdateProperty.Enabled = bUpdateProperty
        End If
    End If
    
   
End Sub

Public Sub Resetoffice97Save(oAppl As Object, id1 As Variant, id2 As Variant, bStatus As Boolean)
    
    On Error Resume Next
    
    'Office97 Word Save menu and toolbar save
    If oAppl.Name = LoadResString(STR_MS_WORD) Then
        If (InStr(1, oAppl.Version, "8", vbTextCompare)) Then
             Call ResetMSMenuItem(oAppl, MB_FILE, MB_FILE_SAVE, bStatus)
             Call ResetMSToolbarButton(oAppl, "Standard", MB_FILE_SAVE, bStatus)
        End If
    End If
    'office 97 and 2000 Excel Save and toolbar save
    If oAppl.Name = "Microsoft Excel" Then
        'If (InStr(1, oAppl.Version, "9", vbTextCompare)) Then
             Call ResetMSMenuItem(oAppl, MB_FILE, MB_FILE_SAVE, bStatus)
             Call ResetMSToolbarButton(oAppl, "Standard", MB_FILE_SAVE, bStatus)
        'End If
    End If
    
End Sub

Public Sub ResetMSToolbarButton(oAppl As Object, ToolbarName As String, ID As Integer, bStatus As Boolean)
     
     Dim oBtn As CommandBarControl
     Dim OStdToolbar As CommandBar
     
     On Error Resume Next
           
     Set OStdToolbar = oAppl.CommandBars(ToolbarName)

     For Each oBtn In OStdToolbar.Controls
         
         'If oBtn.BuiltIn = True Then
             If oBtn.Caption = "&Save" And oBtn.ID = 3 Then
                  oBtn.Enabled = bStatus
             End If
         'End If
         
     Next oBtn

End Sub

Private Sub ResetFileFullName(sFileFullName As String, sOldExt)
    'to fix a MS common dialog bug
    Dim sExt As String
    Dim sStripedFileFullName As String
    
    If InStr(sFileFullName, "..") Then
       Call GetFileNameAndExt(sFileFullName, sStripedFileFullName, sExt)
       If Right(sStripedFileFullName, 1) = "." And InStr(sOldExt, sExt) Then
          sStripedFileFullName = Left(sStripedFileFullName, (Len(sStripedFileFullName) - 1))
       End If
       sFileFullName = sStripedFileFullName & sOldExt
    End If
    
End Sub
Public Function CheckReplicaAndDisableMenuitem(iApplType As Integer, oAppl As Object) As Boolean
    Dim cb As CommandBar
    Dim blogon As Boolean
    Dim oDoc As IDMObjects.Document
    Dim sFilePath As String
    
    CheckReplicaAndDisableMenuitem = False
    
    For Each cb In oAppl.CommandBars
        If cb.Name = LoadResString(MSG_FILENET) Then
            Exit For
        End If
    Next cb
    sFilePath = getFullName(oAppl, iApplType)
    Set oDoc = GetDocObject(sFilePath, blogon)
    
    If oDoc Is Nothing Then
        GoTo Done
    End If
    
    If oDoc.GetState(idmDocIsReplica) = True Then
        Call modifyControls(oAppl, iApplType, cb, MNU_FN_SAVE, MSG_SAVE, False)
        CheckReplicaAndDisableMenuitem = True
    End If
    
    Exit Function

Done:

End Function
Public Function FormatMsg(Msg As String, msg2 As String) As String
   FormatMsg = Replace(Msg, "^0", msg2)
End Function

Public Function SetValueEx(hKey As Long, sValueName As String, vValue As Variant)
    Dim lRetVal As Long
    Dim lType As Long
    Dim lVal As Long
    
    If IsNumeric(vValue) = True Then
       lType = REG_DWORD
       lVal = CLng(vValue)
    Else
       lType = REG_SZ
       
    End If
    
    Select Case lType
         Case REG_SZ        'string
             lRetVal = RegSetValueExString(hKey, sValueName, 0&, lType, vValue, Len(vValue))
         Case REG_DWORD     'DWORD
             lRetVal = RegSetValueEx(hKey, sValueName, 0&, lType, lVal, Len(lVal))
    End Select
    
End Function
Public Sub CreateLogFile(Value As Variant)

    Open App.Path & "idmmacro.txt" For Append As #1
    Print #1, Value
    Close #1

End Sub
Function IsFileOpen(FileName As String) As Boolean
       Dim filenum As Integer, errnum As Integer

       On Error Resume Next   ' Turn error checking off.
       filenum = FreeFile()   ' Get a free file number.
       ' Attempt to open the file and lock it.
       Open FileName For Input Lock Read As #filenum
       Close filenum          ' Close the file.
       errnum = Err           ' Save the error number that occurred.
       On Error GoTo 0        ' Turn error checking back on.

       ' Check to see which error occurred.
       Select Case errnum

           ' No error occurred.
           ' File is NOT already open by another user.
           Case 0
               IsFileOpen = False

           ' Error number for "Permission Denied."
           ' File is already opened by another user.
           Case 70
               IsFileOpen = True

           ' Another error occurred.
           Case Else
               Error errnum
       End Select
   End Function
Public Function ConvertPathToUNC(FilePath As String) As String
   
   Dim sDrive As String
   Dim sFilePath As String
   Dim sUNC As String
   Dim lRet As Long
   Dim sRemoteName As String
   Dim lSizeOfRemoteName As Long
   Dim i As Integer
   
   'ConvertPathToUNC = ""

   lSizeOfRemoteName = lBuffer_size
   sUNC = ""
   sDrive = Left(FilePath, 2) 'such as y:
   If sDrive = "\\" Then
      sUNC = FilePath
      ConvertPathToUNC = sUNC ' no change
      Exit Function
   End If
   sRemoteName = sRemoteName & Space(lBuffer_size)
   
   'return unc path (\\server\share)
   lRet = WNetGetConnection(sDrive, sRemoteName, lSizeOfRemoteName)
   
   If lRet = No_ERROR Then
   
        If Trim(sRemoteName) <> "" Then
           i = InStr(1, sRemoteName, ChrW(0), vbTextCompare)
           sUNC = Left(sRemoteName, i - 1) & Right(FilePath, (Len(FilePath) - 2))
           'sUNC = Left(sRemoteName, Len(Trim(sRemoteName)) - 1) & Right(FilePath, (Len(FilePath) - 2))
        Else
            sUNC = ""
        End If
   Else
        sUNC = ""
   End If
  
   ConvertPathToUNC = sUNC
End Function

Private Function IsCompoundDocument(FilePath As String) As Boolean
    'oAppObject--Application Object
    'giApplType--Application Type
    Dim oAppl As Object
    Dim iCount As Integer
    Dim i As Integer
    Dim sFileFullName As String
    
    iCount = 0
    
    Set oAppl = oAppObject

    On Error Resume Next
    
    Select Case giApplType
        Case APPL_WORD:
        
            iCount = CountWordCompoundDocuments
            
        Case APPL_EXCEL:
             
             Dim vLinks As Variant
             vLinks = oAppl.ActiveWorkbook.LinkSources(xlOLELinks)
            
             If Not IsEmpty(vLinks) Then
             
                For i = 1 To UBound(vLinks)
                    sFileFullName = CStr(vLinks(i))
                    If sFileFullName <> "" Then
                        iCount = iCount + 1
                    End If
                Next i
                
             End If
             
             
        Case APPL_POWERPOINT:
             
             Dim oSlide As PowerPoint.Slide
             Dim oShapes As PowerPoint.Shapes
             Dim oShape As PowerPoint.Shape
             Dim oLinkFormat As PowerPoint.LinkFormat

             'loop through the slides in a activePresentation
             For Each oSlide In oAppl.ActivePresentation.Slides
                 'loop through shape in the slide
                 For Each oShape In oSlide.Shapes
                     Set oLinkFormat = oShape.LinkFormat
                     sFileFullName = oLinkFormat.SourceFullName
                     If sFileFullName <> "" Then
                        iCount = iCount + 1
                    End If

                 Next
             Next
             
     End Select
     
     If iCount > 0 Then
         IsCompoundDocument = True
     Else
         IsCompoundDocument = False
     End If


End Function

Private Function CountWordCompoundDocuments() As Integer
    
    On Error GoTo errHandler
    
    Dim sFileFullName As String
    Dim iCount As Integer
    Dim oLinkFormat As Word.LinkFormat
    
    iCount = 0
    
    'case 1 --- inlineshapes
    Dim oDocument As Word.Document
    Dim oInlineSharps As Word.InlineShapes
    Dim oInlineSharp As Word.InlineShape
    Dim i As Integer

'    Set oInlineSharps = Word.ActiveDocument.InlineSharps
'    For Each oInlineSharp In oInlineSharps
'        Set oLinkFormat = oInlineSharp.LinkFormat
'        sFileFullName = oLinkFormat.SourceFullName
'        If sFileFullName <> "" Then
'            iCount = iCount + 1
'        End If
'    Next

    'case 2 --- shapes
    Dim oShapes As Word.Shapes
    Dim oShape  As Word.Shape
    
    Set oShapes = Word.ActiveDocument.Shapes
    For Each oShape In oShapes
        Set oLinkFormat = oShape.LinkFormat
        sFileFullName = oLinkFormat.SourceFullName
        If sFileFullName <> "" Then
            iCount = iCount + 1
        End If
    Next
    
    'case 3 --- fields
    Dim oFields As Word.Fields
    Dim oField As Word.Field
    Set oFields = Word.ActiveDocument.Fields
    For Each oField In oFields
        Set oLinkFormat = oField.LinkFormat
        sFileFullName = oLinkFormat.SourceFullName
        If sFileFullName <> "" Then
            iCount = iCount + 1
        End If
    Next
    
    CountWordCompoundDocuments = iCount
    
    Exit Function
errHandler:

    CountWordCompoundDocuments = 0

End Function

Public Function CompareFilePath(ByVal FilePath1 As String, ByVal FilePath2 As String) As Boolean
   
   On Error Resume Next
   
   Dim sFilePath1 As String
   Dim sFilePath2 As String
   sFilePath1 = FilePath1
   sFilePath2 = FilePath2
   
   If (Left(sFilePath1, 2) = "\\") Then
   
      sFilePath2 = ConvertPathToUNC(sFilePath2)
      
   ElseIf (Left(sFilePath2, 2) = "\\") Then
     
      sFilePath1 = ConvertPathToUNC(sFilePath1)
      
   End If
   
   If (sFilePath1 = sFilePath2) Then
        CompareFilePath = True
   Else
        CompareFilePath = False
   End If
      
End Function

