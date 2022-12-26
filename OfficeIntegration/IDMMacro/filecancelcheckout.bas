Attribute VB_Name = "fileCancelCheckoutbas"
Option Explicit
Dim ActiveName As String

Public Function fileCancelCheckout(oAppl As Object, iApplType As Integer, strFileFullname As String) As Long
    Dim oVer As IDMObjects.Version
    Dim oCom As IDMObjects.Compound
    Dim oDoc As IDMObjects.Document
    Dim oErrorMgr As ErrorManager
    Dim vbResult As VbMsgBoxResult
    Dim menuCtrl As CommandBarControl
    Dim oBeh As IDMObjects.Behavior
    Dim oCmdData As IDMObjects.CommandCancelCheckoutData
    Dim oAction As IDMObjects.action
    Dim bLogin As Boolean

    On Error GoTo errHandler
    Set oErrorMgr = CreateObject(CREATE_ERR_MGR)
    If oErrorMgr Is Nothing Then
        Err.Raise 1, LoadResString(MSG_CANCEL_CHECKOUT), LoadResString(MSG_CANNOT_CREATE_ERRMGR)
    End If

    If Not (oAppl Is Nothing) Then
        Set menuCtrl = oAppl.CommandBars.FindControl(Tag:=g_FN_CANCEL(iApplType))
        'Bail right away if no documents are open!!
        If DocCount(oAppl, iApplType) = 0 Then
            If iApplType <> APPL_POWERPOINT Then
                menuCtrl.Enabled = False
            Else
                MsgBox LoadResString(MSG_FILE_NOT_CHECKOUT), vbOKOnly + vbCritical, LoadResString(MSG_CANCEL_CHECKOUT)
            End If
            GoTo Done
        End If
        strFileFullname = getFullName(oAppl, iApplType)
    End If

    If GetDocStatus(strFileFullname) <> DocCheckedout Then
        vbResult = MsgBox(strFileFullname & LoadResString(MSG_NOT_CHECKOUT), vbInformation, LoadResString(MSG_CANCEL_CHECKOUT))
        'handle word update menu, in excel the routine is called in the template
        If iApplType = APPL_WORD Then
           Call IDMUpdateMenu(iApplType, oAppl)
        End If
        fileCancelCheckout = CIDMCancel
        GoTo Done
    End If

    Set oDoc = GetDocObject(strFileFullname, bLogin)
    
    If bLogin = False Then
        GoTo Done
    End If
    
    If oDoc Is Nothing Then
        Err.Raise 1, LoadResString(MSG_CANCEL_CHECKOUT), LoadResString(MSG_NO_DOC)
    End If

    If oDoc.GetState(idmDocCanCancelCheckout) = False Then
       MsgBox strFileFullname & LoadResString(MSG_NOT_CHECKOUT), vbInformation, LoadResString(MSG_CANCEL_CHECKOUT)
       Call ResetDocStatus(strFileFullname)
       'call update menu
       Call IDMUpdateMenu(iApplType, oAppl)
       fileCancelCheckout = CIDMCancel
       GoTo Done
    End If
    
    vbResult = MsgBox(LoadResString(MSG_KEEP_LOCAL_COPY) & strFileFullname & "'?", vbYesNoCancel + vbQuestion + vbDefaultButton2, LoadResString(MSG_CANCEL_CHECKOUT))
    If vbResult = vbCancel Then
       fileCancelCheckout = CIDMCancel
       GoTo Done
    End If
    
    If (oDoc.Library.GetState(idmLibrarySupportsCompoundDocuments) = True) Then
        Set oCom = oDoc.Compound
        If oCom Is Nothing Then
            Err.Raise 1, LoadResString(MSG_CANCEL_CHECKOUT), LoadResString(MSG_CANNOT_GET_VERSION)
        End If
        Set oBeh = oDoc.Compound.Behavior
        If Not oBeh Is Nothing Then
            Set oCmdData = New IDMObjects.CommandCancelCheckoutData
            Let oCmdData.Document = oDoc
            Set oAction = oBeh.CreateRootAction(oCmdData)
            ActiveName = getFullName(oAppl, iApplType)
            Call TransversAction(oAppl, iApplType, oAction)
        End If
        Call CloseDocument(oAppl, iApplType)
        If vbResult = vbNo Then
           If oCom.CancelCheckout(idmCDCancelCheckoutWithUI, Nothing) = True Then
              fileCancelCheckout = CIDMOk
           Else
              Call fileOpenOffice(oAppl, iApplType, strFileFullname)
              fileCancelCheckout = CIDMCancel
           End If
        Else
           If oCom.CancelCheckout(idmCancelCheckoutKeep Or idmCDCancelCheckoutWithUI, Nothing) = True Then
              fileCancelCheckout = CIDMOk
           Else
              Call fileOpenOffice(oAppl, iApplType, strFileFullname)
              fileCancelCheckout = CIDMCancel
           End If
        End If
    Else
        Set oVer = oDoc.Version
        If oVer Is Nothing Then
            Err.Raise 1, LoadResString(MSG_CANCEL_CHECKOUT), LoadResString(MSG_CANNOT_GET_VERSION)
        End If
        If vbResult = vbNo Then
           oVer.CancelCheckout idmCancelCheckoutKeep
        Else
           oVer.CancelCheckout idmCancelCheckoutKeep
        End If
        Call CloseDocument(oAppl, iApplType)

        If vbResult = vbNo Then
            Kill strFileFullname
        End If

        fileCancelCheckout = CIDMOk
    End If
    GoTo Done

errHandler:
    ' Check for Automation Error from TrackFile->GetObject when user hits Cancel on Logon dialog
    ' and don't display any errors, just finish.
    fileCancelCheckout = CIDMError
    If Err.Number <> 0 And Err.Number <> -2147216381 Then
        MsgBox Err.Description, vbCritical, LoadResString(MSG_CANCEL_CHECKOUT)
    Else
        If oErrorMgr Is Nothing Then
            MsgBox LoadResString(MSG_ERROR_WITHOUT_ERRMGR), vbCritical, LoadResString(MSG_CANCEL_CHECKOUT)
            GoTo Done
        Else
            If oErrorMgr.Errors.Count > 0 Then
                oErrorMgr.ShowErrorDialog
            End If
        End If
    End If

Done:
    Set oVer = Nothing
    Set oDoc = Nothing
    Set oCom = Nothing
    Set menuCtrl = Nothing
    Set oBeh = Nothing
    Set oCmdData = Nothing
    Set oAction = Nothing
End Function
Private Sub CloseDocument(oAppl As Object, iApplType As Integer)
     Dim menuCtrl As CommandBarControl
     If Not (oAppl Is Nothing) Then
           Call DocClose(oAppl, iApplType)
           If DocCount(oAppl, iApplType) = 0 Then
             If iApplType <> APPL_POWERPOINT Then
                 Set menuCtrl = oAppl.CommandBars.FindControl(Tag:=g_FN_CANCEL(iApplType))
                 menuCtrl.Enabled = False
             End If
        End If
     End If
End Sub
Private Sub TransversAction(oAppl As Object, iApplType As Integer, action As IDMObjects.action)
    Dim SubAction As IDMObjects.action
    If (DocIsOpen(oAppl, iApplType, action.TargetObjectFilePath, False) = True) Then
        DocCloseByfileName oAppl, iApplType, action.TargetObjectFilePath
    End If
    DocMakeActive oAppl, iApplType, ActiveName
    For Each SubAction In action.SubActions
        TransversAction oAppl, iApplType, SubAction
    Next
End Sub
