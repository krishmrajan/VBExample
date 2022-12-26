Attribute VB_Name = "fileSavebas"
Option Explicit

Public Function fileNetSave(oAppl As Object, iApplType As Integer, strFileFullname As String) As Long
'*************************************************************
'* Function: fileSave
'*************************************************************
'* Description: Performs a FileNet Save on the open document
'*
'* For documents that are NEW, the user is asked to SAVE the
'* document, then an AddItem is done on the document which
'* creates an initial version of the doc in Panagon.
'*
'* For documents that are not new, the user is asked if they
'* want to save the document, then the document is checked
'* in.
'*
'* Finally, the document is checked back out, because in both
'* cases the document was closed by Panagon.
'*
'* Input: Object (handle to application)
'*        Integer (represents application type)
'*
'* Output: None
'*************************************************************
    
    Dim lReturn As Long
    
    On Error GoTo errHandler
    
    gbFileNETSave = False
    
     If DocCount(oAppl, iApplType) > 0 Then
'    If (iAppltype = APPL_POWERPOINT) Then
        If (BlockPowerPoint(oAppl, iApplType) = True) Then
            MsgBox LoadResString(MSG_POWERPOINT_BLOCK), vbExclamation, LoadResString(MSG_CHECKIN)
            fileNetSave = CIDMCancel
            GoTo Done
        End If
'    End If
     End If
    
    If Not (oAppl Is Nothing) Then
        lReturn = saveChanges(oAppl, iApplType, strFileFullname, LoadResString(MSG_SAVEAS))
        If lReturn <> CIDMOk Then
            fileNetSave = lReturn
            GoTo Done
        End If
    End If
   
    'Put save function in here
    lReturn = SaveItem(oAppl, iApplType, strFileFullname)
    fileNetSave = lReturn
    'now, reopen the doc if user doed not cancel
    If lReturn = CIDMOk Then
       Call fileOpenOffice(oAppl, iApplType, strFileFullname)
    End If
    gbFileNETSave = True
    GoTo Done
    
errHandler:
    fileNetSave = CIDMError
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbCritical, LoadResString(MSG_FILENETSAVE)
    End If
Done:
   
End Function

Public Function SaveItem(oAppl As Object, iApplType As Integer, strFileFullname As String) As Long
    Dim strfilepath As String
    Dim strfilename As String
    Dim strDirectory As String
    Dim oErrorMgr As ErrorManager
    Dim oDoc As IDMObjects.Document
    Dim oVer As IDMObjects.Version
    Dim oCom As IDMObjects.Compound
    Dim lResult As Long
    
    On Error GoTo errHandler
    Set oErrorMgr = CreateObject(CREATE_ERR_MGR)
    If oErrorMgr Is Nothing Then
        Err.Raise 1, LoadResString(MSG_FILENETSAVE), LoadResString(MSG_CANNOT_CREATE_ERRMGR)
    End If
        
    'Check Tracked Files object to see if active has been checked out. Since this is an ADD,
    'force the user to do a local SaveAs before doing the add (instead of just continuing which
    'will cause problems with the TrackedFiles object if local copy is removed).
        
    If GetDocStatus(strFileFullname) <> DocCheckedout Then
        'this means the doc was NOT checked out, so add it
        lResult = AddItem(oAppl, iApplType, strFileFullname, oDoc, idmSaveAdd)
    Else
         'if a doc is replica then stop
        If CheckReplicaAndDisableMenuitem(iApplType, oAppl) = True Then
             MsgBox LoadResString(STR_REPLICA), vbInformation + vbOKOnly, LoadResString(MSG_FILENETSAVE)
             SaveItem = CIDMError
             GoTo Done
        End If
        'this means it WAS checked out, so close it then check it in
        lResult = Checkin(oAppl, iApplType, strFileFullname, oDoc, idmSaveCheckin)
    End If
        
    If lResult = CIDMOk Then
        If oDoc.Library.SystemType = idmSysTypeIS Then
           strFileFullname = oDoc.GetCachedFile(1, , idmDocGetOriginalFileName)
           SaveItem = CIDMOk
           GoTo Done
        End If
        If oDoc.GetState(idmDocCanCheckout) = False Then
            SaveItem = CIDMError
            If oErrorMgr Is Nothing Then
                MsgBox LoadResString(MSG_ERR_MANAGER_NOT_INITIALIZED), vbCritical, LoadResString(MSG_OPEN)
                GoTo Done
            Else
                If oErrorMgr.Errors.Count > 0 Then
                    MsgBox oErrorMgr.Errors.Item(1).Description, vbExclamation, LoadResString(MSG_OPEN)
                Else
                    MsgBox LoadResString(MSG_CANNOT_CHECKOUT) & strfilepath, vbExclamation, LoadResString(MSG_OPEN)
                End If
            End If
            GoTo Done
        End If
    
        Call idmGetDirectoryAndFileName(strFileFullname, strDirectory, strfilename)
    
        'need to check the document right back out so it appears to have "stayed open"
        If (oDoc.Library.GetState(idmLibrarySupportsCompoundDocuments) = True) Then
            Set oCom = oDoc.Compound
            If oCom Is Nothing Then
                Err.Raise 1, LoadResString(MSG_OPEN), LoadResString(MSG_CANNOT_GET_VERSION)
            End If
            If (oCom.Checkout(strfilepath, strDirectory, strfilename, idmCDCheckoutWithUI, Nothing) = True) Then
               strFileFullname = strfilepath
               SaveItem = CIDMOk
            Else
               SaveItem = CIDMCancel
            End If
        Else
            Set oVer = oDoc.Version
            If oVer Is Nothing Then
                Err.Raise 1, LoadResString(MSG_OPEN), LoadResString(MSG_CANNOT_GET_VERSION)
            End If
            oVer.Checkout strfilepath, strDirectory, strfilename
            strFileFullname = strfilepath
            SaveItem = CIDMOk
        End If
        'SaveItem = CIDMOk
    Else
        SaveItem = CIDMCancel
    End If

GoTo Done

errHandler:
    SaveItem = CIDMError
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbCritical, LoadResString(MSG_FILENETSAVE)
    Else
        If oErrorMgr Is Nothing Then
            MsgBox LoadResString(MSG_ERROR_WITHOUT_ERRMGR), vbCritical, LoadResString(MSG_FILENETSAVE)
            GoTo Done
        Else
            If oErrorMgr.Errors.Count > 0 Then
                oErrorMgr.ShowErrorDialog
            End If
        End If
    End If
Done:
    Set oErrorMgr = Nothing
    Set oDoc = Nothing
    Set oVer = Nothing
    Set oCom = Nothing
 
End Function
