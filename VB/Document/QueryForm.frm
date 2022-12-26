VERSION 5.00
Object = "{00727125-0202-11D1-9BEB-00A0241E626D}#2.0#0"; "fnqsmpl.ocx"
Begin VB.Form QueryForm 
   Caption         =   "Find Documents"
   ClientHeight    =   5100
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7845
   LinkTopic       =   "Form1"
   ScaleHeight     =   5100
   ScaleWidth      =   7845
   StartUpPosition =   3  'Windows Default
   Begin IDMSimpleQuery.IDMSimpleQuery IDMSimpleQuery1 
      Height          =   4335
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   7575
      _Version        =   131072
      _ExtentX        =   13361
      _ExtentY        =   7646
      _StockProps     =   0
   End
   Begin VB.CommandButton Cancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6360
      TabIndex        =   1
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton FindNow 
      Caption         =   "Find "
      Height          =   375
      Left            =   4920
      TabIndex        =   0
      Top             =   4560
      Width           =   1215
   End
End
Attribute VB_Name = "QueryForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cancel_Click()
    QueryForm.Hide
End Sub

Private Sub FindNow_Click()
    Dim oResults As ADODB.Recordset
    'Dim oDocuments As IDMObjects.ObjectSet
    Dim oDocuments As Object
    Set oResults = IDMSimpleQuery1.Execute
    If oResults Is Nothing Then
        Exit Sub
    Else
        If oResults.RecordCount > 0 Then
            Set oDocuments = oResults.Fields("ObjSet").Value
            MainForm.ListView1.ClearItems
            MainForm.ListView1.AddItems oDocuments, -1
            Dim Scope As Object
            Set Scope = IDMSimpleQuery1.SearchScope
            If Scope.ObjectType = idmObjTypeLibrary Then
                Set MainForm.oCurrentLibrary = Scope
            ElseIf Scope.ObjectType = idmObjTypeFolder Then
                Set MainForm.oCurrentLibrary = Scope.Library
            End If
            QueryForm.Hide
        Else
            MsgBox "No documents found"
        End If
    End If
End Sub

Private Sub Form_Activate()
    If Not MainForm.oFolder Is Nothing Then
        IDMSimpleQuery1.SearchScope = MainForm.oFolder
    ElseIf Not MainForm.oCurrentLibrary Is Nothing Then
        IDMSimpleQuery1.SearchScope = MainForm.oCurrentLibrary
    End If
End Sub
