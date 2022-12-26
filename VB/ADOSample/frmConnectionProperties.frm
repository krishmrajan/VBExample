VERSION 5.00
Begin VB.Form frmConnectionProperties 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Connection Properties"
   ClientHeight    =   4416
   ClientLeft      =   4920
   ClientTop       =   3972
   ClientWidth     =   6528
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4416
   ScaleWidth      =   6528
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Default         =   -1  'True
      Height          =   375
      Left            =   4800
      TabIndex        =   8
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton cmdSetValue 
      Caption         =   "&Set Value"
      Height          =   375
      Left            =   3360
      TabIndex        =   7
      Top             =   3600
      Width           =   1215
   End
   Begin VB.TextBox txtAttributes 
      Height          =   375
      Left            =   2760
      TabIndex        =   6
      Top             =   2880
      Width           =   3255
   End
   Begin VB.TextBox txtType 
      Height          =   375
      Left            =   2760
      TabIndex        =   4
      Top             =   1920
      Width           =   3255
   End
   Begin VB.TextBox txtValue 
      Height          =   735
      Left            =   2760
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   600
      Width           =   3255
   End
   Begin VB.ListBox lstProperties 
      Height          =   3504
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "Attributes"
      Height          =   255
      Left            =   2880
      TabIndex        =   5
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Type"
      Height          =   255
      Left            =   2760
      TabIndex        =   3
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Va 
      Caption         =   "Value:"
      Height          =   255
      Left            =   2760
      TabIndex        =   2
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "frmConnectionProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Form for viewing and setting connection properties

Public obj As Object
Dim once As Boolean
Dim strPropName As String

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdSetValue_Click()
On Error GoTo Handle
    obj.Properties(strPropName).Value = txtValue
    Exit Sub
Handle:
    ShowError
End Sub

Private Sub Form_Activate()
    If once And Not (obj Is Nothing) Then
        Dim objProperty As ADODB.Property
        Dim i As Integer
        Dim strRow As String
        i = 0
        For Each objProperty In obj.Properties
            strRow = objProperty.Name
            lstProperties.AddItem strRow, i
            i = i + 1
        Next objProperty
        once = False
        ' ADO doesn't allow setting properties of connection
        ' or recordset objects after they are opened.  we just
        ' disable the SetValue button for these cases.
        If (TypeOf obj Is ADODB.Recordset Or _
            TypeOf obj Is ADODB.Connection) Then
            cmdSetValue.Enabled = False
        End If
    End If
End Sub

Private Sub Form_Load()
    once = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set obj = Nothing
End Sub

Private Sub lstProperties_Click()
    Dim i As Integer
    Dim objProp As ADODB.Property
    i = lstProperties.ListIndex
    strPropName = lstProperties.List(i)
    
    Set objProp = obj.Properties(i)
    txtAttributes = objProp.Attributes
    txtType = objProp.Type
    If objProp.Type <> ADODB.adIDispatch And objProp.Type <> ADODB.adIUnknown Then
        txtValue = obj.Properties(i).Value
    Else
        txtValue = "Object"
    End If
    Dim bitsAttr As Long
    
    ' Command object allows setting properties after it is opened.
    ' But we still want to disable SetValue button for ReadOnly
    ' properties.
    bitsAttr = objProp.Attributes
    If (TypeOf obj Is ADODB.Command And _
        (bitsAttr And adPropWrite) > 0) Then
        cmdSetValue.Enabled = True
    Else
        cmdSetValue.Enabled = False
    End If
End Sub

