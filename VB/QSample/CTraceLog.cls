VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CTraceLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_oFNLogServer As IDMLogging.FnLogServer
Private m_sAppName As String

Public Function initialize() As Boolean

    On Error GoTo Init_Error
    
    'Create our reference to the LogServer
    Set m_oFNLogServer = CreateObject("IDMLogging.FnLogServer")
    
    m_sAppName = App.Path & "\" & App.EXEName & ".exe"
    
    initialize = True
        
    Exit Function

Init_Error:

    'Failed to create error manager
    initialize = False
    
End Function

Private Sub Class_Terminate()

    'Cleanup memory
    Set m_oFNLogServer = Nothing
    
End Sub

Public Sub traceFunctionEntry(ByVal sFunctionName As String, Optional ByVal sSrcFileName As String = "", Optional ByVal sTechMsg As String = "")

    Dim oRec As New IDMLogging.FnTraceFuncEntryRecord
    Dim aDesc(0) As String
    
    On Error GoTo Trace_FuncI_Err
    
    'Set data
    oRec.ModuleName = m_sAppName
    oRec.Function = sFunctionName
    oRec.filename = sSrcFileName
    
    aDesc(0) = sTechMsg
    oRec.Description = aDesc
    
    'Log record
    m_oFNLogServer.Log oRec, False
    
Trace_FuncI_Err:

    Set oRec = Nothing

End Sub

Public Sub traceFunctionExit(ByVal sFunctionName As String, Optional ByVal sSrcFileName As String = "", Optional ByVal sTechMsg As String = "")

    Dim oRec As New IDMLogging.FnTraceFuncExitRecord
    Dim aDesc(0) As String
    
    On Error GoTo Trace_FuncO_Err
    
    'Set data
    oRec.ModuleName = m_sAppName
    oRec.Function = sFunctionName
    oRec.filename = sSrcFileName
    
    aDesc(0) = sTechMsg
    oRec.Description = aDesc
     
    'Log record
    m_oFNLogServer.Log oRec, False
    
Trace_FuncO_Err:

    Set oRec = Nothing

End Sub

Public Sub traceClassCreate(ByVal sClassName As String, Optional ByVal sTechMsg As String = "")

    Dim oRec As New IDMLogging.FnTraceCtorRecord
    Dim aDesc(0) As String
    
    On Error GoTo Trace_ClassI_Err
    
    'Set data
    oRec.ModuleName = m_sAppName
    oRec.ClassName = sClassName
    
    aDesc(0) = sTechMsg
    oRec.Description = aDesc
    
    'Log record
    m_oFNLogServer.Log oRec, False
    
Trace_ClassI_Err:

    Set oRec = Nothing

End Sub
Public Sub traceClassDestroy(ByVal sClassName As String, Optional ByVal sTechMsg As String = "")

    Dim oRec As New IDMLogging.FnTraceDtorRecord
    Dim aDesc(0) As String
  
    On Error GoTo Trace_ClassO_Err
    
    'Set data
    oRec.ModuleName = m_sAppName
    oRec.ClassName = sClassName
    
    aDesc(0) = sTechMsg
    oRec.Description = aDesc
    
    'Log record
    m_oFNLogServer.Log oRec, False
    
Trace_ClassO_Err:

    Set oRec = Nothing

End Sub

Public Sub traceModuleLoad(Optional ByVal sTechMsg As String = "")

    Dim oRec As New IDMLogging.FnTraceModuleLoadRecord
    Dim aDesc(0) As String
  
    On Error GoTo Trace_ModI_Err
    
    'Set data
    oRec.ModuleName = m_sAppName
    
    aDesc(0) = sTechMsg
    oRec.Description = aDesc
    
    'Log record
    m_oFNLogServer.Log oRec, False
    
Trace_ModI_Err:

    Set oRec = Nothing

End Sub
Public Sub traceModuleUnload(Optional ByVal sTechMsg As String = "")

    Dim oRec As New IDMLogging.FnTraceModuleUnloadRecord
    Dim aDesc(0) As String
  
    On Error GoTo Trace_ModI_Err
    
    'Set data
    oRec.ModuleName = m_sAppName
    
    aDesc(0) = sTechMsg
    oRec.Description = aDesc
    
    'Log record
    m_oFNLogServer.Log oRec, False
    
Trace_ModI_Err:

    Set oRec = Nothing

End Sub

