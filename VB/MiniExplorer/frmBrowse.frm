VERSION 5.00
Object = "{54E6FE61-B93A-11CF-8185-444553540000}#3.0#0"; "fntree.ocx"
Begin VB.Form frmBrowse 
   Caption         =   "Select destination folder for folder copy"
   ClientHeight    =   4095
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   5085
   LinkTopic       =   "Form1"
   ScaleHeight     =   4095
   ScaleWidth      =   5085
   StartUpPosition =   3  'Windows Default
   Begin IDMTreeView.IDMTreeView itvIDMTreeView 
      Height          =   3132
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4812
      _Version        =   196608
      _ExtentX        =   8488
      _ExtentY        =   5524
      _StockProps     =   228
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   1
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2640
      TabIndex        =   1
      Top             =   3480
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1200
      TabIndex        =   0
      Top             =   3480
      Width           =   1200
   End
End
Attribute VB_Name = "frmBrowse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()

    Unload Me
    
End Sub

Private Sub cmdOK_Click()

Dim loCurSelTreeItem As Object
Dim liCurSelTreeObjType As Integer
        
On Error GoTo ErrorHandler

        Set loCurSelTreeItem = itvIDMTreeView.SelectedItem
        liCurSelTreeObjType = loCurSelTreeItem.ObjectType
        
        If liCurSelTreeObjType = idmObjTypeFolder Then
            Dim loLibrary As Object
            Set loLibrary = goCurSelTreeItem.Library
            Dim loDestFolder As Object
            Set loDestFolder = loLibrary.GetObject(idmObjTypeFolder, loCurSelTreeItem.PathName)
            goCurSelTreeItem.Copy loDestFolder
        Else
            MsgBox LoadResString(GI_ERR_UNABLE_TO_COPY), vbExclamation, LoadResString(GI_ERR_ERROR)
        End If

    Unload Me
    
    Exit Sub
    
ErrorHandler:
    
    'Display Error Message - pass the name of this subroutine/function
    DisplayErrorMessage ("frmBrowse.cmdOK_Click")
    
    'Cleanup Error values
    CleanupErrorCodes

    Resume Next
    
End Sub

Private Sub Form_Load()

Dim loLibrary As IDMObjects.Library

    Set loLibrary = goCurSelTreeItem.Library
    
    MouseWait
    gbSuccess = AddToIDMTreeView(itvIDMTreeView, loLibrary, True)
    If gbSuccess = False Then
        MsgBox LoadResString(GI_ERR_UNABLE_TO_ADD_TO_TREEVIEW), vbExclamation, LoadResString(GI_ERR_ERROR)
    End If
    MouseNormal

End Sub
