VERSION 5.00
Object = "{2A601BA0-B880-11CF-8185-444553540000}#3.0#0"; "fnlist.ocx"
Begin VB.Form frmAddLibrary 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add Library"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2775
   Icon            =   "frmAddLibrary.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   2775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancelAdd 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1440
      TabIndex        =   8
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton cmdAddLibrary 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   3240
      Width           =   1215
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      Left            =   960
      TabIndex        =   3
      Top             =   2760
      Width           =   1695
   End
   Begin VB.TextBox txtLoginId 
      Height          =   285
      Left            =   960
      TabIndex        =   2
      Text            =   "Admin"
      Top             =   2400
      Width           =   1695
   End
   Begin IDMListView.IDMListView lstLibs 
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   2535
      _Version        =   196608
      _ExtentX        =   4471
      _ExtentY        =   2778
      _StockProps     =   239
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   1
      MultiSelect     =   0   'False
      _ColumnHeaders  =   "frmAddLibrary.frx":1272
   End
   Begin VB.Label lblLibSelected 
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   2040
      Width           =   2535
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Password"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   2760
      Width           =   690
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Login ID"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   2400
      Width           =   600
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Available Libraries"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1275
   End
End
Attribute VB_Name = "frmAddLibrary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAddLibrary_Click()
    
    If Not lstLibs.SelectedItem Is Nothing Then
        Dim oNewServedLib As ServicedLib
        Set oNewServedLib = New ServicedLib
        With oNewServedLib
            .sLibraryName = lstLibs.SelectedItem.Name
            .sLoginID = txtLoginId.Text
            .sPassword = txtPassword.Text
            .Initialize
        End With
        
        oServedLibs.Add oNewServedLib
        Unload Me
    End If
    
End Sub

Private Sub cmdCancelAdd_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    LoadAvailableLibraries
End Sub

Sub LoadAvailableLibraries()
    Dim oTempLib As Library
    For Each oTempLib In oHood.Libraries
        If Not IsLibraryServiced(oTempLib) Then
            lstLibs.AddItem oTempLib, -1
        End If
    Next
End Sub


Function IsLibraryServiced(oLib As Library) As Boolean
    Dim oServLib As ServicedLib
    IsLibraryServiced = False
    For Each oServLib In oServedLibs
        If oLib.Name = oServLib.oMyLibrary.Name Then
            IsLibraryServiced = True
            Exit For
        End If
    Next
End Function

Private Sub lstLibs_ItemSelectChange(ByVal Item As Object, ByVal Key As Long, ByVal Selected As Boolean, ByVal ObjType As IDMListView.idmObjectType)
    lblLibSelected = Item.Name
End Sub
