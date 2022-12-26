VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRegPersist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Sub SaveSettings(ByVal sAppName As String, _
    ByVal sSection As String, ByVal fFrm As Form)
Dim oCtrl As Control
For Each oCtrl In fFrm
    If TypeOf oCtrl Is TextBox Then
        Call SaveSetting(sAppName, sSection, oCtrl.Name, oCtrl.Text)
    End If
Next
End Sub
Public Sub GetSettings(ByVal sAppName As String, _
    ByVal sSection As String, ByRef fFrm As Form)
Dim oCtrl As Control
Dim sTemp As String
For Each oCtrl In fFrm
    If TypeOf oCtrl Is TextBox Then
        sTemp = GetSetting(sAppName, sSection, _
            oCtrl.Name)
        fFrm.Controls(oCtrl.Name) = sTemp
    End If
Next
End Sub
Public Sub DeleteSettings(ByVal sAppName As String, _
    ByVal sSection As String)
On Error Resume Next
Call DeleteSetting(sAppName, sSection)
End Sub


