Attribute VB_Name = "Module1"
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

Public Sub Main()
Set oErrManager = CreateObject("IDMError.ErrorManager")
Form1.Show vbModal
End Sub
