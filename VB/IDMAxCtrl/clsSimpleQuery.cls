VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSimpleQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' This class is useful for generating background document
' queries where there isn't a need for using one of the
' FNQuery ActiveX controls.  This class encapsulates all the
' ADO details while providing control over the filter
' conditions in the query.

' Instructions for use:
'    1.  Call BindToLib passing a Library object and a
'        collection of  headings you want to see in the
'        IDMListView.  An empty collection implies that no
'        column headings will be displayed.
'    2.  Call ExecQuery, passing in the where clause, the
'        folder constraints, the max number of rows, and
'        the IDMListView control you want to populate.  A
'        reminder on the where clause - literal string values
'        must be bracketed with quotes, e.g. AccountName = 'Bruce'
         
Option Explicit

Private oRS As ADODB.Recordset
Private sConnect As String
Private sQuery As String
Private oQueryLib As IDMObjects.Library
Private oPropDescs As IDMObjects.PropertyDescriptions
Dim cColHeadings As New Collection


' If we keep the library as a global variable, we can
' cache data like property descriptions and column headings
Public Sub BindToLib(ByVal oNewLib As IDMObjects.Library, _
    pColHeadings As Collection)
Dim sClasses(1) As String
Dim oTmp As Variant

sClasses(0) = ""
Set oQueryLib = oNewLib
Set oPropDescs = oQueryLib.FilterPropertyDescriptions(idmObjTypeDocument, sClasses)
For Each oTmp In pColHeadings
    ' Weed out any bogus labels the caller passed us
    On Error Resume Next
    If Not (oPropDescs(oTmp) Is Nothing) Then
        cColHeadings.Add oTmp
    End If
Next
End Sub

' Private subroutine for building up IDMListView
Private Sub ShowResults(IDMLView As IDMListView.IDMListView)
Dim oTmp As Variant

' Do basic IDMLView initialization
IDMLView.DefaultLibrary = oQueryLib
IDMLView.ClearItems
' Now do the column header stuff - client told us
' what to use; empty collection => don't do them
If cColHeadings.Count > 0 Then
    Call IDMLView.ClearColumnHeaders(oQueryLib)
    On Error Resume Next
    For Each oTmp In cColHeadings
        Call IDMLView.AddColumnHeader(oQueryLib, oPropDescs(oTmp))
    Next
    Call IDMLView.SwitchColumnHeaders(oQueryLib)
    IDMLView.View = idmViewReport
Else
    IDMLView.View = idmViewList
End If
' Now for the easy part - slam in the actual items
If oRS.RecordCount > 0 Then
    IDMLView.AddItems oRS.Fields("ObjSet").Value, 1
End If

End Sub

' Executes query using passed params, places results in
' passed IDMListView control
' Calls must be preceded by a BindToLib
Public Sub ExecQuery(IDMLView As IDMListView.IDMListView, _
    ByVal sWhereClause As String, ByVal sFolderName As String, _
    ByVal iMaxRows As Integer)

If Not (oQueryLib Is Nothing) Then
    ' Build the string necessary to bind to the database connection
    sConnect = "provider=FnDBProvider;data source=" _
        & oQueryLib.Name _
        & ";Prompt=4;SystemType=" & oQueryLib.SystemType _
        & ";"
    ' Build the query string
    sQuery = "SELECT * FROM FnDocument "
    If Len(sWhereClause) > 0 Then
        sQuery = sQuery & "WHERE " & sWhereClause
    End If
    ' Set up the properties on the record set
    If Not (oRS Is Nothing) Then
        Set oRS = Nothing
    End If
    Set oRS = New ADODB.Recordset
    oRS.ActiveConnection = sConnect
    oRS.Properties("SupportsObjSet") = True
    If iMaxRows > 0 Then
        oRS.MaxRecords = iMaxRows
    End If
    oRS.Properties("SearchFolderName") = sFolderName
    ' All set up - pull the trigger
    oRS.Open sQuery, , adOpenKeyset
    Call ShowResults(IDMLView)
Else
    MsgBox ("Must set library!")
End If
End Sub
