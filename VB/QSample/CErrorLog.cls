VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CErrorLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Used for Sample Applications"
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Const CLASS_NAME = "CErrorLog"
Const SRC_FILE_NAME = "CErrorLog.cls"

Private m_oFNErrManager As idmError.ErrorManager
Private m_sErrorFile As String
Private m_bShowNativeFNErr As Boolean
Private m_oTraceLog As New CTraceLog

Public Enum ErrorType
    errFatal = vbCritical
    errWarning = vbExclamation
    errInfo = vbInformation
End Enum


Public Function logError(ByVal iErrType As ErrorType, ByVal sTechErrorMsg As String, Optional ByVal sUserErrorMsg As String = "") As Boolean

    Dim lErr As Long
    Dim sErrDesc As String
    
    'Store a reference to the current error properties (VB resets it in subsequent calls
    lErr = Err.Number
    sErrDesc = Err.Description
    
    m_oTraceLog.traceFunctionEntry CLASS_NAME & ".logError", SRC_FILE_NAME
    
    If (sUserErrorMsg <> "") Then
        MsgBox sUserErrorMsg & " - " & sErrDesc, iErrType, "Error in " & App.EXEName
    End If
 
    logError = writeToErrorFile("ERROR: " & sTechErrorMsg & vbCrLf & " Native ERROR: " & sErrDesc & " (" & CStr(lErr) & ")")
    
    m_oTraceLog.traceFunctionExit CLASS_NAME & ".logError", SRC_FILE_NAME
    
End Function

Public Function logFNError(ByVal iErrType As ErrorType, ByVal sTechErrorMsg As String, Optional ByVal sUserErrorMsg As String = "") As Boolean

    Dim i As Integer
    Dim oCurrErr As idmError.Error
    Dim lErr As Long
    Dim sErrDesc As String
    
    'Store a reference to the current error properties (VB resets it in subsequent calls
    lErr = Err.Number
    sErrDesc = Err.Description
    
    m_oTraceLog.traceFunctionEntry CLASS_NAME & ".logFNError", SRC_FILE_NAME
    
    'If there is a user error, show it
    If (sUserErrorMsg <> "") Then
        MsgBox sUserErrorMsg & " - " & sErrDesc, iErrType, AppName
    End If
    
    'Make sure we have error in the FileNet error stack, otherwise show native error
    If (m_oFNErrManager.Errors.Count > 0) Then
                
        If m_bShowNativeFNErr Then m_oFNErrManager.ShowErrorDialog
        
        sTechErrorMsg = "ERROR: " & sTechErrorMsg
        
        For Each oCurrErr In m_oFNErrManager.Errors

            sTechErrorMsg = sTechErrorMsg & vbCrLf & " FileNet ERROR: " & _
                oCurrErr.Description & " (Number: " & CStr(oCurrErr.Number) & _
                ")" & vbCrLf & vbTab & "Error Class: " & CStr(oCurrErr.Class) & _
                vbCrLf & vbTab & "Error Code: " & CStr(oCurrErr.Code) & _
                vbCrLf & vbTab & "Native Error: " & CStr(oCurrErr.NativeError) & _
                vbCrLf & vbTab & "Error Source: " & oCurrErr.Source
                
        Next 'i
        Set oCurrErr = Nothing
         
        logFNError = writeToErrorFile(sTechErrorMsg)
        
    Else
    
        logFNError = writeToErrorFile("ERROR: " & sTechErrorMsg & vbCrLf & " Native ERROR: " & sErrDesc & " (" & CStr(lErr) & ")")
       
    End If
    
    m_oTraceLog.traceFunctionExit CLASS_NAME & ".logFNError", SRC_FILE_NAME

    
End Function

Private Sub Class_Initialize()

    m_oTraceLog.initialize
    m_oTraceLog.traceClassCreate CLASS_NAME
    
End Sub

Private Sub Class_Terminate()

    'Cleanup memory
    Set m_oFNErrManager = Nothing
    
    m_oTraceLog.traceClassDestroy CLASS_NAME
    Set m_oTraceLog = Nothing
    
End Sub

Public Function initialize(Optional ByVal sErrorFile As String = "C:\FNError.txt", Optional ByVal bShowNative As Boolean = True)

    m_oTraceLog.traceFunctionEntry CLASS_NAME & ".initialize", SRC_FILE_NAME
    
    On Error GoTo Init_Error
    
    m_sErrorFile = sErrorFile
    m_bShowNativeFNErr = bShowNative
    
    'Create our reference to the ErrorManager
    Set m_oFNErrManager = CreateObject("IDMError.ErrorManager")
    
    initialize = True
        
    m_oTraceLog.traceFunctionExit "CErrorLog.initialize", SRC_FILE_NAME
    
    Exit Function

Init_Error:

    'Failed to create error manager
    writeToErrorFile ("Failed to create the IDMError.ErrorManager object.")
    
    initialize = False
    
    m_oTraceLog.traceFunctionExit "CErrorLog.initialize", SRC_FILE_NAME, "Failed"

End Function

Private Function writeToErrorFile(sMsg As String) As Boolean

    Dim iFileNum As Integer
    Dim bFileOpen As Boolean
    
    m_oTraceLog.traceFunctionEntry CLASS_NAME & ".writeToErrorFile", SRC_FILE_NAME
    
    bFileOpen = False
    
    On Error GoTo Write_ErrorFile_Err
    
    'Open file
    iFileNum = FreeFile
    Open m_sErrorFile For Append As #iFileNum
    bFileOpen = True
    
    'Write to file
    Print #iFileNum, CStr(Now) & " - " & sMsg
    Print #iFileNum, "******************************************************************"
    Print #iFileNum, ""
    
    Close #iFileNum
    bFileOpen = False
    
    writeToErrorFile = True
    
    m_oTraceLog.traceFunctionExit CLASS_NAME & ".writeToErrorFile", SRC_FILE_NAME
 
    Exit Function
    
Write_ErrorFile_Err:

    MsgBox ("Unable to write to error file (" & m_sErrorFile & ") Error: " & sMsg)
    If (bFileOpen) Then
        Close #iFileNum
    End If

    writeToErrorFile = False
    
    m_oTraceLog.traceFunctionExit CLASS_NAME & ".writeToErrorFile", SRC_FILE_NAME, "Failed"
   
End Function
