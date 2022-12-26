VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 1  'vbDataSource
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Database"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim WithEvents adoRS As Recordset
Attribute adoRS.VB_VarHelpID = -1

Private Sub Class_Initialize()
    On Error Resume Next
    Dim db As Connection
    Set db = New Connection
    db.CursorLocation = adUseClient
    db.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & App.Path & "\CustomCVL.mdb;"
    
    Set adoRS = New Recordset
    adoRS.Open strSQL, db, adOpenStatic, adLockOptimistic
    
    DataMembers.Add "Primary"
    Exit Sub
End Sub

Private Sub Class_GetDataMember(DataMember As String, Data As Object)
  Select Case DataMember
  Case "Primary"
    Set Data = adoRS
  End Select
End Sub

Public Property Get RecordCount() As Long
    'Number of Records
    RecordCount = adoRS.RecordCount
End Property

Public Function CorrespondingData() As String
    'This function will be called by PropertyDescriptionPlugin_Item
    'One record will be returned at a time
    
    On Error Resume Next
    CorrespondingData = adoRS.Fields(0).Value
    If Not adoRS.EOF Then
        adoRS.MoveNext
    Else
        adoRS.MoveFirst
    End If
    Exit Function
End Function





