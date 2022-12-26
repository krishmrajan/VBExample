VERSION 5.00
Begin VB.Form frmFind 
   Caption         =   "Form1"
   ClientHeight    =   6420
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6840
   LinkTopic       =   "Form1"
   ScaleHeight     =   6420
   ScaleWidth      =   6840
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cmbValues 
      Height          =   315
      Index           =   8
      Left            =   2760
      TabIndex        =   19
      Text            =   "Combo1"
      Top             =   4800
      Width           =   3375
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   3960
      TabIndex        =   17
      Top             =   5640
      Width           =   975
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "Find"
      Height          =   495
      Left            =   2160
      TabIndex        =   16
      Top             =   5640
      Width           =   975
   End
   Begin VB.ComboBox cmbValues 
      Height          =   315
      Index           =   7
      Left            =   2760
      TabIndex        =   15
      Text            =   "Combo1"
      Top             =   4215
      Width           =   3375
   End
   Begin VB.ComboBox cmbValues 
      Height          =   315
      Index           =   6
      Left            =   2760
      TabIndex        =   14
      Text            =   "Combo1"
      Top             =   3630
      Width           =   3375
   End
   Begin VB.ComboBox cmbValues 
      Height          =   315
      Index           =   5
      Left            =   2760
      TabIndex        =   13
      Text            =   "Combo1"
      Top             =   3045
      Width           =   3375
   End
   Begin VB.ComboBox cmbValues 
      Height          =   315
      Index           =   4
      Left            =   2760
      TabIndex        =   12
      Text            =   "Combo1"
      Top             =   2460
      Width           =   3375
   End
   Begin VB.ComboBox cmbValues 
      Height          =   315
      Index           =   3
      Left            =   2760
      TabIndex        =   11
      Text            =   "Combo1"
      Top             =   1875
      Width           =   3375
   End
   Begin VB.ComboBox cmbValues 
      Height          =   315
      Index           =   2
      Left            =   2760
      TabIndex        =   10
      Text            =   "Combo1"
      Top             =   1290
      Width           =   3375
   End
   Begin VB.ComboBox cmbValues 
      Height          =   315
      Index           =   1
      Left            =   2760
      TabIndex        =   9
      Text            =   "Combo1"
      Top             =   705
      Width           =   3375
   End
   Begin VB.ComboBox cmbValues 
      Height          =   315
      Index           =   0
      Left            =   2760
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   120
      Width           =   3375
   End
   Begin VB.Label Labels 
      Caption         =   "Labels"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   360
      TabIndex        =   18
      Top             =   4920
      Width           =   1935
   End
   Begin VB.Label Labels 
      Caption         =   "Labels"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   360
      TabIndex        =   8
      Top             =   4335
      Width           =   1935
   End
   Begin VB.Label Labels 
      Caption         =   "Labels"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   360
      TabIndex        =   7
      Top             =   3750
      Width           =   1935
   End
   Begin VB.Label Labels 
      Caption         =   "Labels"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   360
      TabIndex        =   6
      Top             =   3165
      Width           =   1935
   End
   Begin VB.Label Labels 
      Caption         =   "Labels"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   360
      TabIndex        =   5
      Top             =   2580
      Width           =   1935
   End
   Begin VB.Label Labels 
      Caption         =   "Labels"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   360
      TabIndex        =   4
      Top             =   1995
      Width           =   1935
   End
   Begin VB.Label Labels 
      Caption         =   "Labels"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   360
      TabIndex        =   3
      Top             =   1410
      Width           =   1935
   End
   Begin VB.Label Labels 
      Caption         =   "Labels"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   360
      TabIndex        =   2
      Top             =   825
      Width           =   1935
   End
   Begin VB.Label Labels 
      Caption         =   "Labels"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   1935
   End
End
Attribute VB_Name = "frmFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim iTermLimit As Integer



Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdFind_Click()
Dim sWhere As String
Dim bCondition As Boolean
Dim iInx As Integer
Dim sClass As String
bCondition = False
sClass = gfSettings.txtResDocClass
' Construct the where clause from the form fields
For iInx = 0 To iTermLimit - 1
    If cmbValues(iInx) <> "" Then
        If bCondition Then
            sWhere = sWhere & " AND "
        End If
        sWhere = sWhere & gcPropNames(iInx + 1) & " = '" & _
            cmbValues(iInx) & "'"
        bCondition = True
    End If
Next
' Now launch the query and populate the ListView
If bCondition Then
    Dim oQuery As New clsSimpleQuery
    Call oQuery.BindToLib(goISLib, gcPropNames, sClass)
    MousePointer = vbHourglass
    Call oQuery.ExecQuery(ResumeForm.ListView1, sWhere, "", 20)
    MousePointer = vbArrow
    ResumeForm.ViewerCtrl1.Clear
    Me.Hide
End If

End Sub

Private Sub Form_Load()
    Dim iInx As Integer
    Dim sTmp As Variant
    Dim oCtrl As Control
    Dim oPropDesc As IDMObjects.PropertyDescription
    iTermLimit = 0
    For Each oCtrl In Me.Controls
        If TypeOf oCtrl Is ComboBox Then
            iTermLimit = iTermLimit + 1
        End If
        If TypeOf oCtrl Is ComboBox Or _
            TypeOf oCtrl Is Label Then
            oCtrl.Visible = False
        End If
    Next
    iInx = 0
    For Each sTmp In gcHeadings
        Me.Labels(iInx) = sTmp
        Me.Labels(iInx).Visible = True
        Me.cmbValues(iInx).Visible = True
        Me.cmbValues(iInx).Text = ""
        Set oPropDesc = goPropDescs(gcPropNames(iInx + 1))
        If oPropDesc.GetState(idmChoice) Then
            Dim oChoice As IDMObjects.Choice
            Dim oChoices As IDMObjects.Choices
            Set oChoices = oPropDesc.Choices
            If oPropDesc.GetState(idmChoicePaging) Then
                oChoices.Paging.NextPage (0)
            End If
            For Each oChoice In oChoices
                cmbValues(iInx).AddItem oChoice.Value
            Next
        End If
        iInx = iInx + 1
        If iInx = iTermLimit Then
            Exit For
        End If
    Next
    If iTermLimit > gcPropNames.Count Then
        iTermLimit = gcPropNames.Count
    End If
    ' Fix this stuff!!
    'Set oPropDesc = goISLib.GetObject(idmObjTypePropDesc, gcHeadings(4))
    'If oPropDesc.GetState(idmChoice) Then
    '    Dim oChoice As IDMObjects.Choice
    '    Dim oChoices As IDMObjects.Choices
    '    Set oChoices = oPropDesc.Choices
    '    oChoices.Paging.NextPage (0)
    '    For Each oChoice In oChoices
    '        Value4.AddItem oChoice.Value
     '   Next
    'End If
End Sub

