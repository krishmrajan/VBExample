VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmRowset 
   Caption         =   "Query Rowset"
   ClientHeight    =   3840
   ClientLeft      =   7368
   ClientTop       =   4032
   ClientWidth     =   8520
   LinkTopic       =   "Form1"
   ScaleHeight     =   3840
   ScaleWidth      =   8520
   Begin VB.CommandButton cmdGetData 
      Caption         =   "Get &Data"
      Height          =   375
      Left            =   6960
      TabIndex        =   1
      Top             =   3240
      Width           =   1215
   End
   Begin ComctlLib.ListView LstView 
      Height          =   2775
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   8055
      _ExtentX        =   14203
      _ExtentY        =   4890
      Arrange         =   1
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   ""
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Menu mnuProperties 
      Caption         =   "ADO &Properties"
   End
End
Attribute VB_Name = "frmRowset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Logic for manipulating a record set and displaying
' rows in a basic MS listview.

Public rs As New ADODB.Recordset
Dim once As Boolean

Private Sub cmdGetData_Click()
    If once And Not (rs Is Nothing) Then
        CreateColumnHeaders
        If rs.EOF Then
            MsgBox "No records selected."
        Else
            Screen.MousePointer = vbHourglass
            rs.MoveFirst
            LoadResultsPage
            Screen.MousePointer = vbDefault
        End If
        once = False
    End If
End Sub

Private Sub Form_Load()
    once = True
End Sub

Private Sub Form_Resize()
    LstView.Width = frmRowset.ScaleWidth * 0.9
    LstView.Height = frmRowset.ScaleHeight - 360 - 720
    cmdGetData.Top = frmRowset.ScaleHeight - (375 * 2)
    cmdGetData.Left = frmRowset.ScaleWidth - (1275 + 500)
End Sub


Private Sub mnuProperties_Click()
    Set frmConnectionProperties.obj = rs
    frmConnectionProperties.Show vbModal, Me
End Sub
' Pull the field names out of the record set and put
' them in the listview columns
Private Sub CreateColumnHeaders()
    Dim field As ADODB.field
    Dim iColumn As Integer
    For iColumn = 1 To rs.Fields.Count
        Set field = rs.Fields(iColumn - 1)
        LstView.ColumnHeaders.Add iColumn, , field.Name, LstView.Width / rs.Fields.Count
    Next 'iColumn
    LstView.HideColumnHeaders = False
    LstView.View = lvwReport
End Sub
' Pull the data out of the recordset and put it in
' the listview
Private Sub LoadResultsPage()
On Error GoTo Handle
    Dim iColumn As Integer
    Dim oListItem As ListItem
    Dim strValue As String
    Do While Not rs.EOF
        Set oListItem = LstView.ListItems.Add
        For iColumn = 0 To rs.Fields.Count - 1
            If IsNull(rs.Fields(iColumn).Value) Then
                strValue = TypeName(rs.Fields(iColumn).Value)
            Else
                If rs.Fields(iColumn).Name = "F_DOCNUMBER" Then 'FR 25963
                    If rs.Fields(iColumn).Value < 0 Then
                        strValue = CStr(4294967296# + rs.Fields(iColumn).Value)
                    Else
                        strValue = CStr(rs.Fields(iColumn))
                    End If
                Else
                    strValue = CStr(rs.Fields(iColumn))
                End If
            End If
            If iColumn = 0 Then
                oListItem.Text = strValue
            Else
                oListItem.SubItems(iColumn) = strValue
            End If 'iColumn = 0
        Next 'iColumn
        rs.MoveNext
    Loop
    Exit Sub
Handle:
    ShowError
End Sub

