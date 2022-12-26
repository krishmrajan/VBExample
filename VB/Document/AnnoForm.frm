VERSION 5.00
Object = "{2A601BA0-B880-11CF-8185-444553540000}#1.0#0"; "fnlist.ocx"
Begin VB.Form AnnoForm 
   Caption         =   "Annotations"
   ClientHeight    =   5748
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   7416
   LinkTopic       =   "Form1"
   ScaleHeight     =   5748
   ScaleWidth      =   7416
   StartUpPosition =   3  'Windows Default
   Begin IDMListView.IDMListView ListView1 
      Height          =   4575
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   6855
      _Version        =   65536
      _ExtentX        =   12091
      _ExtentY        =   8070
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
      _ColumnHeaders  =   "AnnoForm.frx":0000
   End
   Begin VB.CommandButton Done 
      Caption         =   "Done"
      Height          =   375
      Left            =   5880
      TabIndex        =   0
      Top             =   5280
      Width           =   1215
   End
End
Attribute VB_Name = "AnnoForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Done_Click()
    AnnoForm.Hide
End Sub
