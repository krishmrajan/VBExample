Attribute VB_Name = "Checkinbas"
Option Explicit

Public Function Checkin(oAppl As Object, iApplType As Integer, strFileFullname As String, oDoc As IDMObjects.Document, Optional enuSaveCheckin As AddCheckinEnum) As Long
    Dim vbResult As VbMsgBoxResult
    Dim oErrorMgr As ErrorManager
    Dim oVer As IDMObjects.Version
    Dim intWizard As idmwizard
    Dim strfilepath As String
    Dim strfilename As String
    Dim eWizardStatus As IDMObjects.idmDocumentWizardExecuteStatus
    Dim blogon As Boolean
    Dim oPreference As IDMPreferences.Preference
    Dim oPreferences As New IDMPreferences.Preferences
     
    On Error GoTo errHandler
    Set oErrorMgr = CreateObject(CREATE_ERR_MGR)
    If oErrorMgr Is Nothing Then
        Err.Raise 0, LoadResString(MSG_CHECKIN), LoadResString(MSG_CANNOT_CREATE_ERRMGR)
    End If
    
    Call idmGetDirectoryAndFileName(strFileFullname, strfilepath, strfilename)

    If GetDocStatus(strFileFullname) <> DocCheckedout Then
        vbResult = MsgBox(strFileFullname & LoadResString(MSG_NOT_CHECKOUT), vbExclamation, LoadResString(MSG_FILECHECKIN))
        Checkin = CIDMCancel
        GoTo Done
    End If
        
    Set oDoc = GetDocObject(strFileFullname, blogon)
    If blogon = False Then  '08/18/99 added
       Checkin = CIDMCancel
       Exit Function
    End If
    If oDoc Is Nothing Then
        Err.Raise 0, LoadResString(MSG_CHECKIN), LoadResString(MSG_NO_DOC)
    End If
    If oDoc.GetState(idmDocCanCheckin) = False Then
       MsgBox strFileFullname & LoadResString(MSG_NOT_CHECKOUT), vbInformation, LoadResString(MSG_FILECHECKIN)
       'update localDB change checkout to copy
       Call ResetDocStatus(strFileFullname)
       'call update menu
       Call IDMUpdateMenu(iApplType, oAppl)
       Checkin = CIDMCancel
       GoTo Done
    End If
    
    Set intWizard = New idmwizard
    Set intWizard.oWizard = New IDMObjects.DocumentWizard
    If intWizard.oWizard Is Nothing Then
        Err.Raise 0, LoadResString(MSG_CHECKIN), LoadResString(MSG_CANNOT_GET_VERSION)
    End If
    If enuSaveCheckin = idmSaveCheckin Then
        Set oPreference = SetPreferenceInWizard(LoadResString(STR_DOCUMENT), LoadResString(STR_ADD_CHECKIN_RETRIEVE), LoadResString(STR_KEEP_LOCAL_COPY), 0&)
        oPreferences.Add oPreference
        Set oPreference = SetPreferenceInWizard(LoadResString(STR_DOCUMENT), LoadResString(STR_ADD_CHECKIN_RETRIEVE), LoadResString(STR_SHOW_KEEP_LOCAL_COPY), 0&)
        oPreferences.Add oPreference
    End If
    intWizard.oWizard.Preferences = oPreferences
    intWizard.oWizard.FilePath = strFileFullname
    Let intWizard.oWizard.Parent = oDoc
    intWizard.oWizard.Operation = idmDocumentWizardOperationCheckin
    Set intWizard.oAppl = oAppl
    intWizard.iApplType = iApplType
    intWizard.strInSubroutine = LoadResString(MSG_CHECKIN)
    intWizard.CallingOperation = idmCheckin
    If oDoc.GetState(idmDocHasChild) = True Then
        giApplType = iApplType
        Set goAppl = oAppl
        gActionType = idmCheckin
        Call Walklink(oAppl, iApplType, idmCheckin)
        intWizard.oWizard.Callback = New clsRecognizer
        Set intWizard.oDoc = oDoc
        intWizard.enuSaveCheckinAction = enuSaveCheckin
    End If
    If intWizard.oWizard.Show(idmDocumentWizardModeExecute, eWizardStatus) = idmDialogExitCancel Then
        Checkin = CIDMCancel
        GoTo Done
    End If
     
    Checkin = CIDMOk
    
    GoTo Done
    
errHandler:
    ' Check for Automation Error from TrackFile->GetObject when user hits Cancel on Logon dialog
    ' and don't display any errors, just finish.
    Checkin = CIDMError
    If Err.Number <> 0 And Err.Number <> -2147216381 Then
        MsgBox Err.Description, vbCritical, LoadResString(MSG_CHECKIN)
    Else
        If oErrorMgr Is Nothing Then
            MsgBox LoadResString(MSG_ERROR_WITHOUT_ERRMGR), vbCritical, LoadResString(MSG_CHECKIN)
            GoTo Done
        Else
            If oErrorMgr.Errors.Count > 0 Then
                oErrorMgr.ShowErrorDialog
            End If
        End If
    End If
    
Done:
    Set oErrorMgr = Nothing
    Set oVer = Nothing
    Set intWizard = Nothing
    Set goAppl = Nothing
End Function
