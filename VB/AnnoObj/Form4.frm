VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "comctl32.ocx"
Begin VB.Form Form4 
   Caption         =   "Annotation Properties"
   ClientHeight    =   4365
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7545
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   ScaleHeight     =   4365
   ScaleWidth      =   7545
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton OK 
      Caption         =   "OK"
      Height          =   375
      Left            =   6240
      TabIndex        =   1
      Top             =   3840
      Width           =   1095
   End
   Begin ComctlLib.ListView ListView1 
      Height          =   3495
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   6165
      View            =   3
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Propery Name"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Property Value"
         Object.Width           =   7056
      EndProperty
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
    Dim objProp As idmobjects.Property
    Dim itemX As ListItem
    ListView1.ListItems.Clear
    For Each objProp In Form1.objAnno.Properties
        'Set the first column to the property name
        Set itemX = ListView1.ListItems.Add(, , objProp.PropertyDescription.Name)
        'set the second column to the property name
        itemX.SubItems(1) = objProp.FormatValue
    Next
End Sub

Private Sub OK_Click()
    Form4.Hide
End Sub
