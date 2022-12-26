VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRecognizer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private iApplType As Integer
Private oAppl As Object
Dim ActiveName As String
Private CallingOperation As AddCheckinEnum
Implements IFnRecognizerCallback

Private Sub Class_Initialize()
 Set oAppl = goAppl
 iApplType = giApplType
 CallingOperation = gActionType
End Sub

Private Sub IFnRecognizerCallback_AfterRecognition(ByVal FilePath As String)

End Sub

Private Sub IFnRecognizerCallback_BeforeRecognition(ByVal FilePath As String)
    
    On Error Resume Next

    ActiveName = getFullName(oAppl, iApplType)
    Call SaveAndCloseChildDoc(FilePath)
    
    Select Case iApplType
           Case APPL_WORD
                oAppl.Documents(ActiveName).Activate
           Case APPL_EXCEL
                oAppl.Workbooks.Item(ActiveName).Activate
    End Select
    
End Sub

Private Sub SaveAndCloseChildDoc(sFileFullName As String)
      Dim sDir As String
      Dim sFileName As String
      Dim vbResult As VbMsgBoxResult
      
      If (DocIsOpen(oAppl, iApplType, sFileFullName, False) = True) Then
            idmGetDirectoryAndFileName sFileFullName, sDir, sFileName
            If IsDocSaved(sFileName, iApplType, oAppl) = False Then
               Select Case gActionType
                      Case idmAdd
                           If GetPreferenceValue(LoadResString(STR_PRT_SAVE_ADD), iApplType) = "1" Then
                               vbResult = MsgBox(LoadResString(STR_DO_YOU_WANT_TO_SAVE_THE_DOC) & LoadResString(STR_P_LEFT) & sFileName & LoadResString(STR_P_RIGHT), vbYesNo + vbQuestion, LoadResString(MSG_ADD))
                           Else
                               vbResult = vbYes
                           End If
                      Case idmCheckin
                           If GetPreferenceValue(LoadResString(STR_PRT_SAVE_CHECKIN), iApplType) = "1" Then
                               vbResult = MsgBox(LoadResString(STR_DO_YOU_WANT_TO_SAVE_THE_DOC) & LoadResString(STR_P_LEFT) & sFileName & LoadResString(STR_P_RIGHT), vbYesNo + vbQuestion, LoadResString(MSG_CHECKIN))
                           Else
                               vbResult = vbYes
                           End If
               End Select
               If vbResult = vbYes Then
                  SaveDoc sFileName, iApplType, oAppl
               Else
                  Call fileRevertOffice(oAppl, iApplType, sFileFullName)
               End If
            End If
            CloseDoc sFileName, iApplType, oAppl
      End If
      
End Sub

