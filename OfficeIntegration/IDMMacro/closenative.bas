Attribute VB_Name = "closeNativebas"
Option Explicit

Function closeNative(oAppl As Object, iApplType As Integer) As VbMsgBoxResult
    Dim strfilepath As String
    Dim strFileFullname As String
    Dim strfilename As String
    Dim strRemoveFile As String
    Dim vbResult As VbMsgBoxResult
    Dim bSaved As Boolean
    
    On Error GoTo errHandler
    
    closeNative = vbYes
    
    strfilename = getName(oAppl, iApplType)
    bSaved = DocIsSaved(oAppl, iApplType)
    
    If DocIsNew(oAppl, iApplType) = True Then
        vbResult = MsgBox(LoadResString(MSG_FILE_SAVE_CHANGES) & strfilename & "?", vbYesNoCancel + vbExclamation, LoadResString(MSG_CLOSE))
        Select Case vbResult
            Case vbYes
                DEFAULT_SAVE_PATH = readDefaultSavePath(iApplType)
                strFileFullname = DEFAULT_SAVE_PATH & strfilename
                If DocSaveDialog(oAppl, strFileFullname, iApplType, LoadResString(MSG_FILE_EXISTS_OVERWRITE)) = vbCancel Then
                    closeNative = vbCancel
                    GoTo Done
                Else
                    DocSaveAs oAppl, strFileFullname, iApplType
                    Call DocClose(oAppl, iApplType)
                End If
            Case vbNo
                Call DocClose(oAppl, iApplType)
                GoTo Done
            Case vbCancel
                closeNative = vbCancel
                GoTo Done
        End Select
    Else
        ' The document exist on local drive, just ask if they want to save before closing
        If bSaved = False Then
            vbResult = MsgBox(LoadResString(MSG_FILE_SAVE_CHANGES) & strfilename & "?", vbYesNoCancel + vbExclamation, LoadResString(MSG_CLOSE))
            Select Case vbResult
                Case vbYes
                    DEFAULT_SAVE_PATH = readDefaultSavePath(iApplType)
                    strFileFullname = DEFAULT_SAVE_PATH & strfilename
                    If DocSaveDialog(oAppl, strFileFullname, iApplType, LoadResString(MSG_FILE_EXISTS_OVERWRITE)) = vbCancel Then
                        closeNative = vbCancel
                        GoTo Done
                    Else
                        DocSaveAs oAppl, strFileFullname, iApplType
                        Call DocClose(oAppl, iApplType)
                    End If
                Case vbNo
                    Call DocClose(oAppl, iApplType)
                    GoTo Done
                Case vbCancel
                    closeNative = vbCancel
                    GoTo Done
            End Select
        Else
            Call DocClose(oAppl, iApplType)
        End If
    End If
    GoTo Done
    
Done:
    Exit Function

errHandler:
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbCritical, LoadResString(MSG_CLOSE)
    End If

End Function

