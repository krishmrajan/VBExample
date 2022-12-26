Attribute VB_Name = "fileCheckinbas"
Option Explicit

Public Function fileCheckin(oAppl As Object, iApplType As Integer, strFileFullname As String) As Long
    Dim oDoc As IDMObjects.Document
    Dim lReturn As Long
    Dim vbResult As VbMsgBoxResult
    On Error GoTo errHandler
      
    If DocCount(oAppl, iApplType) = 0 Then
          MsgBox LoadResString(MSG_FILE_NOT_CHECKOUT), vbCritical, LoadResString(MSG_CHECKIN)
          fileCheckin = CIDMCancel
          GoTo Done
    End If
    strFileFullname = getFullName(oAppl, iApplType)
    If GetDocStatus(strFileFullname) <> DocCheckedout Then
        vbResult = MsgBox(strFileFullname & LoadResString(MSG_NOT_CHECKOUT), vbInformation, LoadResString(MSG_FILECHECKIN))
        GoTo Done
    End If
    
 '   If (iAppltype = APPL_POWERPOINT) Then
        If (BlockPowerPoint(oAppl, iApplType) = True) Then
            MsgBox LoadResString(MSG_POWERPOINT_BLOCK), vbExclamation, LoadResString(MSG_CHECKIN)
            fileCheckin = CIDMCancel
            GoTo Done
        End If
'    End If
    
    If Not (oAppl Is Nothing) Then
        lReturn = saveChanges(oAppl, iApplType, strFileFullname, LoadResString(MSG_CHECKIN))
        fileCheckin = lReturn
        If lReturn <> CIDMOk Then
            GoTo Done
        End If
    End If
    lReturn = Checkin(oAppl, iApplType, strFileFullname, oDoc)
    fileCheckin = lReturn
       
    GoTo Done
    
errHandler:
    fileCheckin = CIDMError
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbCritical, LoadResString(MSG_FILECHECKIN)
    End If
   
Done:
    Set oDoc = Nothing
End Function
