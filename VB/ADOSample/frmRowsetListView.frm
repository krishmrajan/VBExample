VERSION 5.00
Object = "{2A601BA0-B880-11CF-8185-444553540000}#3.0#0"; "fnlist.ocx"
Begin VB.Form frmRowsetListView 
   Caption         =   "Results"
   ClientHeight    =   4260
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   7188
   LinkTopic       =   "Form1"
   ScaleHeight     =   4260
   ScaleWidth      =   7188
   StartUpPosition =   3  'Windows Default
   Begin IDMListView.IDMListView lstRowset 
      Height          =   3615
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   6015
      _Version        =   196608
      _ExtentX        =   10610
      _ExtentY        =   6376
      _StockProps     =   239
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.81
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   1
      Sorted          =   0   'False
      MainColumnLabel =   "     Name"
      ShowAnnotations =   -1  'True
      _ColumnHeaders  =   "frmRowsetListView.frx":0000
   End
End
Attribute VB_Name = "frmRowsetListView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Functions for manipulating a recordset, extracting data,
' and placing it in an IDMListView

Public rs As New ADODB.Recordset

Private Sub Form_Load()
On Error GoTo Handle
    Dim objSet As Object
    If Not rs Is Nothing Then
        If rs.EOF Then
            MsgBox "No records selected."
        Else
            Screen.MousePointer = vbHourglass
            CreateColumnHeaders
            Set objSet = rs.Fields("objset").Value
            lstRowset.AddItems objSet, 1
            Screen.MousePointer = vbDefault
        End If
    End If
    Exit Sub
Handle:
    ShowError
End Sub
' Pull the field names from the recordset; get the IDM
' properties for each one, and build the column headers
' of the IDMListView
Private Sub CreateColumnHeaders()
On Error GoTo Handle
    Dim lib As New IDMObjects.Library
    lib.Name = dbGlobals.systemName
    lib.systemType = dbGlobals.systemType
    Dim field As ADODB.field
    Dim iColumn As Integer
    lstRowset.DefaultLibrary = lib
    For iColumn = 1 To rs.Fields.Count - 1
        Set field = rs.Fields(iColumn - 1)
        Dim propDesc As IDMObjects.PropertyDescription
        ' field name is case sensitive here.  IDMDS SQL is case
        ' insensitive.  If you enter a lower-case column name, you
        ' will get error in this statement.
        Set propDesc = lib.GetObject(idmObjTypePropDesc, field.Name, idmObjTypeDocument)
        lstRowset.AddColumnHeader lib, propDesc
    Next 'iColumn
    lstRowset.SwitchColumnHeaders lib
    lstRowset.View = idmViewReport
    Exit Sub
Handle:
    ShowError
End Sub

