Attribute VB_Name = "dbGlobals"
Option Explicit

Public ds As New ADODB.Connection
Public systemType As idmSysTypeOptions
Public systemName As String
Public oErrManager As idmError.ErrorManager

Public Sub ShowError()
Dim oErrCollect As idmError.Errors
Dim oError As idmError.Error
Dim iCnt As Integer
Set oErrCollect = oErrManager.Errors
If oErrCollect.Count > 1 Then
    iCnt = 1
    For Each oError In oErrCollect
        MsgBox "Error " & iCnt & ": " & oError.Description & " : " & Hex(oError.Number)
        iCnt = iCnt + 1
    Next
Else
    If oErrCollect.Count = 1 Then
        oErrManager.ShowErrorDialog
    Else
        If Err.Number <> 0 Then
            MsgBox Err.Description & " : " & Err.Number
        End If
    End If
End If
End Sub


