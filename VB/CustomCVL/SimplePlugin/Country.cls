VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 1  'vbDataSource
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Country"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim WithEvents adoCountryRS As Recordset
Attribute adoCountryRS.VB_VarHelpID = -1

Private Sub Class_Initialize()
    On Error GoTo ErrHandler
    Dim db As Connection
    Set db = New Connection
    db.CursorLocation = adUseClient
    db.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & App.Path & "\Country.mdb;"
    
    Set adoCountryRS = New Recordset
    'Get Country list
    adoCountryRS.Open "SELECT Country.CountryName FROM Country " & _
    "ORDER BY Country.CountryName", db, adOpenForwardOnly, adLockOptimistic
    
    DataMembers.Add "Primary"
    
    Exit Sub
ErrHandler:
    MsgBox Err.Description
End Sub

Private Sub Class_GetDataMember(DataMember As String, Data As Object)
  Select Case DataMember
  Case "Primary"
    Set Data = adoCountryRS
  End Select
End Sub

Public Property Get RecordCount() As Long
    'Number of Countries
    RecordCount = adoCountryRS.RecordCount
End Property

Public Function CorrespondingCountry() As String
    'This function will be called by IFnPropDescPlugin_Item
    'One record will be returned at a time
    
    On Error GoTo ErrHandler
    CorrespondingCountry = adoCountryRS.Fields(0).Value
    If Not adoCountryRS.EOF Then
        adoCountryRS.MoveNext
    Else
        adoCountryRS.MoveFirst
    End If
    Exit Function
ErrHandler:
    MsgBox Err.Description
End Function



