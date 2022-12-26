VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "comctl32.ocx"
Begin VB.Form PropertiesForm 
   Caption         =   "Properties"
   ClientHeight    =   5640
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   8772
   LinkTopic       =   "Form2"
   ScaleHeight     =   5640
   ScaleWidth      =   8772
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Details 
      Caption         =   "Details..."
      Enabled         =   0   'False
      Height          =   375
      Left            =   5640
      TabIndex        =   4
      Top             =   5160
      Width           =   1095
   End
   Begin VB.CommandButton Done 
      Caption         =   "Done"
      Height          =   375
      Left            =   7440
      TabIndex        =   3
      Top             =   5160
      Width           =   1095
   End
   Begin ComctlLib.ListView MsList 
      Height          =   4335
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   8295
      _ExtentX        =   14626
      _ExtentY        =   7641
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327680
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      MouseIcon       =   "PropertiesForm.frx":0000
      NumItems        =   6
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Property Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Value"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Data Type"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Checked Out"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   4
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Has Annotations"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   5
         Key             =   ""
         Object.Tag             =   ""
         Text            =   ""
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   8295
   End
   Begin VB.Label Label1 
      Caption         =   "This dialog shows how to access Document properties"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   5160
      Width           =   4575
   End
End
Attribute VB_Name = "PropertiesForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' oForObject is the IDM object we are showing properties for
Dim oForObject As Object
' PropName is the name of the currently selected property
Public PropName As String

Public Sub ShowPropertiesFor(oIDMObject As Object, CurrentForm As Form, ButtonCaption As String, LabelText As String)
    Set oForObject = oIDMObject
    Dim oProperty As IDMObjects.Property
    Dim itemX As ListItem
    MsList.ListItems.Clear
    For Each oProperty In oIDMObject.Properties
        'Set the first column to the property name
        Set itemX = PropertiesForm.MsList.ListItems.Add(, , oProperty.PropertyDescription.Name)
        If IsNull(oProperty.Value) Then
            itemX.SubItems(1) = "<<IS NULL>>"
        Else
            itemX.SubItems(1) = oProperty.FormatValue
        End If
        itemX.SubItems(2) = MainForm.FormatDataType(oProperty.PropertyDescription.TypeID)
        itemX.SubItems(3) = oProperty.PropertyDescription.GetState(idmPropSearchable)
        itemX.SubItems(4) = oProperty.PropertyDescription.GetState(idmPropMultiValue)
    Next
    Label2 = LabelText
    Done.Caption = ButtonCaption
    Show 1, CurrentForm
End Sub

Private Sub Details_Click()
    Dim oPropDesc As IDMObjects.PropertyDescription
    Set oPropDesc = MainForm.oCurrentLibrary.GetObject(idmObjTypePropDesc, PropertiesForm.PropName, idmObjTypeDocument)
    Dim PropVal As Boolean
    
    ' show the property related state information in the PropDetailForm
    PropDetailForm.Label13 = PropName
    
    PropDetailForm.Label1 = "Has a choice list:"
    PropVal = oPropDesc.GetState(idmChoice)
    PropDetailForm.Text1 = PropVal
    If PropVal Then
        PropDetailForm.BtnChoices.Enabled = True
    Else
        PropDetailForm.BtnChoices.Enabled = False
    End If
    PropDetailForm.Label2 = "Supports paging:"
    PropDetailForm.Text2 = oPropDesc.GetState(idmChoicePaging)
    PropDetailForm.Label3 = "Custom property:"
    PropDetailForm.Text3 = oPropDesc.GetState(idmPropCustom)
    PropDetailForm.Label4 = "Has default:"
    PropDetailForm.Text4 = oPropDesc.GetState(idmPropHasDefault)
    PropDetailForm.Label5 = "Displayable:"
    PropDetailForm.Text5 = oPropDesc.GetState(idmPropDisplayable)
    PropDetailForm.Label6 = "Key:"
    PropDetailForm.Text6 = oPropDesc.GetState(idmPropKey)
    PropDetailForm.Label7 = "Must pick from choices:"
    PropDetailForm.Text7 = oPropDesc.GetState(idmPropMustPick)
    PropDetailForm.Label8 = "Required:"
    PropDetailForm.Text8 = oPropDesc.GetState(idmPropRequired)
    PropDetailForm.Label9 = "Use to query:"
    PropDetailForm.Text9 = oPropDesc.GetState(idmPropSelectable)
    PropDetailForm.Label10 = "Version property:"
    PropDetailForm.Text10 = oPropDesc.GetState(idmPropVersion)
    
    ' set the property value
    Dim oProp As IDMObjects.Property
    Set oProp = oForObject.Properties(PropName)
    If oPropDesc.GetState(idmPropMultiValue) Then
        PropDetailForm.Label11 = "Multiple Values:"
        Dim oMulti As IDMObjects.MultipleValues
        Set oMulti = oProp.Value
        ' Protect against mv props with no values set
        If oMulti.Count = 0 Then
            PropDetailForm.Combo1.AddItem (" ")
        Else
            ii = 1
            While ii <= oMulti.Count
                PropDetailForm.Combo1.AddItem oMulti(ii)
                ii = ii + 1
            Wend
        End If
    Else
        PropDetailForm.Label11 = "Single Value property"
        PropDetailForm.Combo1.AddItem oProp.FormatValue
    End If
    PropDetailForm.Combo1.ListIndex = 0
    
    PropDetailForm.Show 1, PropertiesForm
End Sub

Private Sub Done_Click()
    PropertiesForm.Hide
End Sub

Private Sub Form_Load()
    MsList.ColumnHeaders.Clear
    MsList.ColumnHeaders.Add , , "Property Name"
    MsList.ColumnHeaders.Add , , "Value"
    MsList.ColumnHeaders.Add , , "Data Type"
    MsList.ColumnHeaders.Add , , "Search Criteria"
    MsList.ColumnHeaders.Add , , "Has Multiple Values"
    MsList.View = lvwReport
End Sub

Private Sub MsList_ItemClick(ByVal Item As ComctlLib.ListItem)
    PropName = Item.Text
    PropertiesForm.Details.Enabled = False
    'Set oForObject = Item
End Sub
