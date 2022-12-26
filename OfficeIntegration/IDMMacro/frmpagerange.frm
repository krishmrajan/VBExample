VERSION 5.00
Begin VB.Form frmPageRange 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "FileNET Document Page Range"
   ClientHeight    =   2595
   ClientLeft      =   1170
   ClientTop       =   4695
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2595
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin VB.OptionButton PagesOption 
      Caption         =   "&Pages"
      Height          =   372
      Left            =   360
      TabIndex        =   4
      Top             =   1200
      Width           =   972
   End
   Begin VB.OptionButton AllOption 
      Caption         =   "&All"
      Height          =   372
      Left            =   360
      TabIndex        =   3
      Top             =   600
      Value           =   -1  'True
      Width           =   972
   End
   Begin VB.Frame Frame1 
      Caption         =   "Page range"
      Height          =   1692
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   5652
      Begin VB.TextBox TextLastPage 
         Height          =   285
         Left            =   4200
         TabIndex        =   9
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox TextFirstPage 
         Height          =   285
         Left            =   2280
         TabIndex        =   8
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label LabelFirst 
         Caption         =   "Last Page"
         Height          =   252
         Left            =   3360
         TabIndex        =   7
         Top             =   1200
         Width           =   852
      End
      Begin VB.Label LabelLast 
         Caption         =   "First Page"
         Height          =   372
         Left            =   1320
         TabIndex        =   6
         Top             =   1200
         Width           =   972
      End
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4560
      TabIndex        =   1
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3240
      TabIndex        =   0
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   372
      Left            =   2520
      TabIndex        =   5
      Top             =   1440
      Width           =   972
   End
End
Attribute VB_Name = "frmPageRange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim m_FirstPage As Integer
Dim m_LastPage As Integer
Dim m_PageTotal As Integer
Dim m_CanceledFlag As Boolean
'

Private Sub AllOption_Click()
    TextFirstPage.Enabled = False
    TextLastPage.Enabled = False
    LabelFirst.Enabled = False
    LabelLast.Enabled = False
    TextFirstPage = 1
    TextLastPage = m_PageTotal
End Sub

Private Sub CancelButton_Click()
    m_CanceledFlag = True
    Me.Hide
End Sub

Private Sub Form_Load()
    m_CanceledFlag = False
    TextFirstPage.Text = m_FirstPage
    TextLastPage.Text = m_LastPage
    m_PageTotal = m_LastPage
    TextFirstPage.Enabled = False
    TextLastPage.Enabled = False
    LabelFirst.Enabled = False
    LabelLast.Enabled = False
End Sub

Private Sub Frame1_DragDrop(Source As Control, X As Single, Y As Single)
    Me.Hide
End Sub

Private Sub OKButton_Click()
    ' if "all pages" is selected, set the page range appropriately
    If AllOption.Value = True Then
        m_FirstPage = 1
        m_LastPage = m_PageTotal
    Else
        m_FirstPage = Val(TextFirstPage.Text)
        m_LastPage = Val(TextLastPage.Text)
    End If
    ' verify page range is legal
    If m_FirstPage = 0 Then
        Call MsgBox(LoadResString(DLG_ERR_PAGE_ZERO), vbExclamation + vbOKOnly, LoadResString(MSG_OPEN))
    ElseIf m_FirstPage > m_LastPage Then
        Call MsgBox(LoadResString(DLG_ERR_PAGE_FIRST), vbExclamation + vbOKOnly, LoadResString(MSG_OPEN))
    ElseIf m_LastPage > m_PageTotal Then
        Call MsgBox(LoadResString(DLG_ERR_PAGE_LAST), vbExclamation + vbOKOnly, LoadResString(MSG_OPEN))
        TextLastPage.Text = m_PageTotal
    Else
        Me.Hide
    End If
End Sub

Public Property Get FirstPage() As Integer
    FirstPage = m_FirstPage
End Property

Public Property Let FirstPage(New_FirstPage As Integer)
    m_FirstPage = New_FirstPage
End Property

Public Property Get LastPage() As Integer
    LastPage = m_LastPage
End Property

Public Property Let LastPage(New_LastPage As Integer)
    m_LastPage = New_LastPage
End Property

Public Property Get CanceledFlag() As Integer
    CanceledFlag = m_CanceledFlag
End Property

Public Property Let CanceledFlag(New_CanceledFlag As Integer)
    m_CanceledFlag = New_CanceledFlag
End Property

Private Sub Option1_Click()

End Sub

Private Sub PagesOption_Click()
    TextFirstPage.Enabled = True
    TextLastPage.Enabled = True
    LabelFirst.Enabled = True
    LabelLast.Enabled = True
End Sub

Private Sub TextFirstPage_GotFocus()
    TextFirstPage.SelStart = 0
    TextFirstPage.SelLength = Len(TextFirstPage.Text)
End Sub

Private Sub TextFirstPage_KeyPress(KeyAscii As Integer)
    If ((KeyAscii < 48) Or (KeyAscii > 57)) And (Not (KeyAscii = 8)) Then
        KeyAscii = 0
    End If
End Sub

Private Sub TextLastPage_GotFocus()
    TextLastPage.SelStart = 0
    TextLastPage.SelLength = Len(TextLastPage.Text)
End Sub

Private Sub TextLastPage_KeyPress(KeyAscii As Integer)
    If ((KeyAscii < 48) Or (KeyAscii > 57)) And (Not (KeyAscii = 8)) Then
        KeyAscii = 0
    End If
End Sub
