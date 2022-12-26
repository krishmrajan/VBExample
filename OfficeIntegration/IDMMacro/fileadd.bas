Attribute VB_Name = "fileAddbas"
Option Explicit

Public Function fileAdd(oAppl As Object, iApplType As Integer, strFileFullname As String) As Long
    
    Dim oDoc As IDMObjects.Document
    Dim lReturn As Long
       
    On Error GoTo errHandler
     If DocCount(oAppl, iApplType) > 0 Then
'    If (iAppltype = APPL_POWERPOINT) Then
        If (BlockPowerPoint(oAppl, iApplType) = True) Then
            MsgBox LoadResString(MSG_POWERPOINT_BLOCK), vbExclamation, LoadResString(MSG_ADD)
            fileAdd = CIDMCancel
            GoTo Done
        End If
 '   End If
     End If
    If Not (oAppl Is Nothing) Then
       
        lReturn = saveChanges(oAppl, iApplType, strFileFullname, LoadResString(MSG_ADD))
        fileAdd = lReturn
        If lReturn <> CIDMOk Then
            GoTo Done
        End If

        If iApplType = 2 Then 'Excel
            oAppl.ActiveWindow.Caption = ResetCaption(oAppl, iApplType, 0)
        End If
              
    End If
    
    'AddItem closes the document
    lReturn = AddItem(oAppl, iApplType, strFileFullname, oDoc)
    fileAdd = lReturn
        
    GoTo Done
    
errHandler:
    fileAdd = CIDMError
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbCritical, LoadResString(MSG_ADD)
    End If
Done:
    Set oDoc = Nothing

End Function

Public Function AddItem(oAppl As Object, iApplType As Integer, strFileFullnameInput As String, oDoc As IDMObjects.Document, Optional enuSaveCheckin As AddCheckinEnum) As Long
    
    Dim vbResult As VbMsgBoxResult
    Dim oCmnDlg As IDMObjects.CommonDialogs
    Dim intWizard As idmwizard
    Dim oAddTo As Object
    Dim oLib As IDMObjects.Library
    Dim oProp As IDMObjects.Property
    Dim oErrorMgr As ErrorManager
    Dim lOperation As Long
    Dim bEvents As Boolean
    Dim vTitle As Variant
    Dim vOption As Variant
    Dim vDocClass As Variant
    Dim newfullname$
    Dim strfilepath As String
    Dim strfilename As String
    Dim oSubSystem As New IDMPreferences.SubSystem
    Dim oCategory As New IDMPreferences.Category
    Dim oPreference As IDMPreferences.Preference
    Dim oOption As IDMPreferences.Option
    Dim oPreferences As New IDMPreferences.Preferences
    Dim eWizardStatus As IDMObjects.idmDocumentWizardExecuteStatus
    Dim strFileFullname As String
    
    On Error GoTo errHandler
    Set oErrorMgr = CreateObject(CREATE_ERR_MGR)
    If oErrorMgr Is Nothing Then
        Err.Raise 1, LoadResString(MSG_ADD), LoadResString(MSG_CANNOT_CREATE_ERRMGR)
    End If
    
    Dim bOutlookMemoAdd As Boolean
    Dim bOutlookAttachmentAdd As Boolean
    bOutlookMemoAdd = False
    bOutlookAttachmentAdd = False
            
    If Right(strFileFullnameInput, 10) = "<<<read>>>" Then
        bOutlookMemoAdd = True
        strFileFullname = Left(strFileFullnameInput, Len(strFileFullnameInput) - 10)
    ElseIf Right(strFileFullnameInput, 10) = "<<<atta>>>" Then
        bOutlookAttachmentAdd = True
        strFileFullname = Left(strFileFullnameInput, Len(strFileFullnameInput) - 10)
    Else
        strFileFullname = strFileFullnameInput
    End If
    
    Call idmGetDirectoryAndFileName(strFileFullname, strfilepath, strfilename)
        
    'Check Tracked Files object to see if active has been checked out. Since this is an ADD,
    'force the user to do a local SaveAs before doing the add (instead of just continuing which
    'will cause problems with the TrackedFiles object if local copy is removed).
    If GetDocStatus(strFileFullname) = DocCheckedout Then
        'vbResult = MsgBox(strFileFullname & LoadResString(MSG_FILE_ADD_CHECKEDOUT), vbOKOnly + vbExclamation, LoadResString(MSG_ADD))
        ' May consider placing the user in the SaveAs dialog instead of just returning
           Dim sLibraryName As String
           Dim sDocID As String
           Call GetDocInfo(strFileFullname, sLibraryName, sDocID)
           If iApplType = APPL_POWERPOINT Then
              If oAppl.ActivePresentation.FullName = strFileFullname Then
                 vbResult = MsgBox(LoadResString(MSG_FILE_ADD_CHECKEDOUT) & strFileFullname & (MSG_FILE_ADD_CHECKEDOUT2) & vbCrLf & LoadResString(MSG_THE_CHECKEDOUT_FILE_INFO) & strFileFullname & vbCrLf & LoadResString(MSG_DOCID) & sDocID & vbCrLf & LoadResString(MSG_LIBRARY) & sLibraryName & vbCrLf & LoadResString(MSG_WOULD_YOU_LIKE_TO_CONTINE_ADD), vbYesNo + vbInformation, LoadResString(MSG_ADD))
              End If
           Else
              vbResult = MsgBox(LoadResString(MSG_FILE_ADD_CHECKEDOUT) & strFileFullname & LoadResString(MSG_FILE_ADD_CHECKEDOUT1) & vbCrLf & LoadResString(MSG_THE_CHECKEDOUT_FILE_INFO) & strFileFullname & vbCrLf & LoadResString(MSG_DOCID) & sDocID & vbCrLf & LoadResString(MSG_LIBRARY) & sLibraryName & vbCrLf & LoadResString(MSG_WOULD_YOU_LIKE_TO_CONTINE_ADD), vbYesNo + vbExclamation, LoadResString(MSG_ADD))
           End If
           If vbResult = vbNo Then
              AddItem = CIDMCancel
              GoTo Done
           End If
           Call ResetDocStatus(strFileFullname)
    End If
    
    If goCmnDlg Is Nothing Then
        Set goCmnDlg = CreateObject(CREATE_COMMON_DLG)
    End If
    If goCmnDlg Is Nothing Then
        Err.Raise 1, LoadResString(MSG_ADD), LoadResString(MSG_CANNOT_CREATE_CMNDIALOG)
    End If
    
    With goCmnDlg
        .Title = LoadResString(ADD_FN_DOC) 'should be a constant
        .Options = idmSelectHideDrives
        .hWnd = GetActiveWindow
    End With
    '
    If DSinstalled = True Then
        If CBool(GetPreferenceValue(LoadResString(STR_FOLDER_REQ_WHEN_ADD), 6)) = True Then
           vOption = idmSelectFoldersOnly ' idmSelectFolderWriteOnly Or idmSelectFoldersOnly
        End If
    End If
    
    'get default doc title
    If DSinstalled = True Then
        If CBool(GetPreferenceValue(LoadResString(STR_SET_DOC_TITLE_TO_FILENAME), 6)) = True Then
            vTitle = GetTitle(strfilename)
        Else
            vOption = vOption + idmSelectFolderHideTitle
        End If
    End If
    
    Call goCmnDlg.SelectFolder(oAddTo, lOperation, vOption, vTitle, vDocClass)

    If lOperation = idmOperationCancel Then
        AddItem = CIDMCancel
        GoTo Done
    End If
      
    Set oLib = CreateObject(CREATE_LIBRARY)
    If oLib Is Nothing Then
        Err.Raise 1, LoadResString(MSG_ADD), LoadResString(MSG_CANNOT_CREATE_LIBRARY)
    End If
    
    oLib.SystemType = oAddTo.SystemType
    If oAddTo.ObjectType = idmObjTypeFolder Then
        oLib.Name = oAddTo.Library.Name
        Set oDoc = oLib.CreateObject(idmObjTypeDocument, vDocClass, oAddTo)
    Else
        oLib.Name = oAddTo.Name
        Set oDoc = oLib.CreateObject(idmObjTypeDocument, vDocClass)
    End If
    
    If oDoc Is Nothing Then
        Err.Raise 1, LoadResString(MSG_ADD), LoadResString(MSG_CANNOT_CREATE_DOC)
    End If
    '==================
    Call gIdmEvent.AddFooter(oDoc.ID, LoadResString(STR_DIALOG)) '"Dialog")
    '==================
    ' If system Type is DS, then try to set the idmName property
    If oLib.SystemType = idmSysTypeDS Then
        Set oProp = oDoc.GetExtendedProperty("idmName")
        If oProp Is Nothing Then
            Err.Raise 1, LoadResString(MSG_ADD), LoadResString(MSG_CANNOT_GET_PROPERTY)
        End If
        If CBool(GetPreferenceValue(LoadResString(STR_FOLDER_REQ_WHEN_ADD), 6)) = True Then
            If vOption = idmSelectFoldersOnly + idmSelectFolderHideTitle Then
               oProp.Value = oDoc.ID
            Else
               oProp.Value = vTitle
            End If
        Else
            If vOption = idmSelectFolderHideTitle Then
               oProp.Value = oDoc.ID
            Else
               oProp.Value = vTitle
            End If
        End If
    End If
    
    Set intWizard = New idmwizard
    Set intWizard.oWizard = New IDMObjects.DocumentWizard
    
    If intWizard.oWizard Is Nothing Then
        Err.Raise 1, LoadResString(MSG_ADD), LoadResString(MSG_CANNOT_GET_VERSION)
    End If
    If DSinstalled Then
        If iApplType = APPL_OUTLOOK Then
            oSubSystem.Name = LoadResString(STR_DOCUMENT)
            oSubSystem.UserType = idmPoUserCurrent
            Set oCategory = oSubSystem.GetCategory(LoadResString(STR_ADD_CHECKIN_RETRIEVE))
            Set oPreference = oCategory.GetPreference(LoadResString(STR_DCDC))
            Set oOption = oPreference.Value
            oOption.Value = "0"
            oPreferences.Add oPreference
            Set oPreference = oCategory.GetPreference(LoadResString(STR_SDCDC))
            Set oOption = oPreference.Value
            oOption.Value = "0"
            oPreferences.Add oPreference
            If bOutlookAttachmentAdd = True Then
                Set oPreference = SetPreferenceInWizard(LoadResString(STR_DOCUMENT), LoadResString(STR_ADD_CHECKIN_RETRIEVE), LoadResString(STR_KEEP_LOCAL_COPY), 0&)
                oPreferences.Add oPreference
                Set oPreference = SetPreferenceInWizard(LoadResString(STR_DOCUMENT), LoadResString(STR_ADD_CHECKIN_RETRIEVE), LoadResString(STR_SHOW_KEEP_LOCAL_COPY), 0&)
                oPreferences.Add oPreference
            End If
        End If
        If enuSaveCheckin = idmSaveAdd Then
            Set oPreference = SetPreferenceInWizard(LoadResString(STR_DOCUMENT), LoadResString(STR_ADD_CHECKIN_RETRIEVE), LoadResString(STR_KEEP_LOCAL_COPY), 0&)
            oPreferences.Add oPreference
            Set oPreference = SetPreferenceInWizard(LoadResString(STR_DOCUMENT), LoadResString(STR_ADD_CHECKIN_RETRIEVE), LoadResString(STR_SHOW_KEEP_LOCAL_COPY), 0&)
            oPreferences.Add oPreference
        End If
        intWizard.oWizard.Preferences = oPreferences
    End If
    'this part for outlook and IS library only
    If oLib.SystemType = idmSysTypeIS And iApplType = APPL_OUTLOOK Then
       Call TruncateFileName(strFileFullname, 200)
    End If
    intWizard.oWizard.FilePath = strFileFullname
    Let intWizard.oWizard.Parent = oDoc
    Set intWizard.oAppl = oAppl
    intWizard.iApplType = iApplType
    intWizard.oWizard.Operation = idmDocumentWizardOperationAdd
    intWizard.strInSubroutine = LoadResString(MSG_ADD)
    intWizard.CallingOperation = idmAdd
    intWizard.SysType = oLib.SystemType
    giApplType = iApplType
    Set goAppl = oAppl
    gActionType = idmAdd
    Call Walklink(oAppl, iApplType, idmAdd)
    intWizard.oWizard.Callback = New clsRecognizer
    intWizard.enuSaveCheckinAction = enuSaveCheckin
    
    'to handle an add wizard bug  10/29/02
    If oLib.SystemType = idmSysTypeIS Then
       Call DocClose(oAppl, iApplType)
    End If
    '===============================
    If intWizard.oWizard.Show(idmDocumentWizardModeExecute, eWizardStatus) = idmDialogExitCancel Then
        'to handle an add wizard bug  10/29/02
        If oLib.SystemType = idmSysTypeIS Then
           Call fileRevertOffice(oAppl, iApplType, strFileFullname)
        End If
        '=============================
        AddItem = CIDMCancel
        GoTo Done
    Else
        If oLib.SystemType = idmSysTypeIS And eWizardStatus = idmDocumentWizardSucceeded Then
           If oAddTo.ObjectType = idmObjTypeFolder Then
              Call oAddTo.File(oDoc)
           End If
        End If
        If iApplType = APPL_OUTLOOK And eWizardStatus = idmDocumentWizardSucceeded Then
            If (Dir$(intWizard.oWizard.FilePath) = "") And (bOutlookAttachmentAdd = False) Then
                ''' Outlook memo, not Outlook attachment was added; file was deleted
                Dim aOutlook As New Outlook.Application
                Dim oItem As Object
                Dim oInspector As Outlook.Inspector
                Dim oExplorer As Outlook.Explorer
                
                Dim myNameSpace As Outlook.NameSpace
                Set myNameSpace = aOutlook.GetNamespace("MAPI")

                Dim oFolder As Outlook.MAPIFolder
                Set oFolder = myNameSpace.GetDefaultFolder(olFolderDeletedItems)
                'MsgBox oFolder.Name
                
                If bOutlookMemoAdd = True Then
                    Set oInspector = Outlook.ActiveInspector
                    If Not oInspector Is Nothing Then
                        Set oItem = oInspector.CurrentItem
                        If Not oItem Is Nothing Then
                            Call oInspector.Close(olDiscard)
                            On Error Resume Next
                            Call oItem.Move(oFolder)
                            On Error GoTo errHandler
                        End If
                    End If
                Else
                    Set oExplorer = aOutlook.ActiveExplorer
                    If Not oExplorer Is Nothing Then
                        Set oItem = oExplorer.Selection(1)
                        If Not oItem Is Nothing Then
                            On Error Resume Next
                            Call oItem.Move(oFolder)
                            On Error GoTo errHandler
                        End If
                    End If
                End If
            End If
        End If
    End If

 '   End If
    
 '   If oLib.SystemType = idmSysTypeIS Then
        ' Just use the Add wizard cause we are not supporting insert properties on IS Documents
  '      Call oDoc.SaveNew(strfilepath & "\" & strfilename, idmDocSaveNewWithUIWizard)
       
 '       If oAddTo.ObjectType = idmObjTypeFolder Then
 '           Call oAddTo.File(oDoc)
 '       End If
 '   End If
    
    AddItem = CIDMOk
    GoTo Done
    
errHandler:
    AddItem = CIDMError
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbCritical, LoadResString(MSG_ADD)
    Else
        If oErrorMgr Is Nothing Then
            MsgBox LoadResString(MSG_ERROR_WITHOUT_ERRMGR), vbCritical, LoadResString(MSG_ADD)
            GoTo Done
        Else
            If oErrorMgr.Errors.Count > 0 Then
                oErrorMgr.ShowErrorDialog
            End If
        End If
    End If
    
Done:
    Set oErrorMgr = Nothing
    Set oProp = Nothing
    Set oLib = Nothing
    Set oAddTo = Nothing
    Set intWizard = Nothing
    Set goAppl = Nothing
End Function

