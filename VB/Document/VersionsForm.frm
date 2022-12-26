VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form VersionsForm 
   Caption         =   "Versions"
   ClientHeight    =   4845
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6165
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   4845
   ScaleWidth      =   6165
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.ListView lvVersions 
      Height          =   3375
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   5953
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton cmdShowProps 
      Caption         =   "Show Properties..."
      Enabled         =   0   'False
      Height          =   375
      Left            =   3120
      TabIndex        =   2
      Top             =   4320
      Width           =   1455
   End
   Begin VB.CommandButton Done 
      Caption         =   "Done"
      Height          =   375
      Left            =   4800
      TabIndex        =   1
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5895
   End
End
Attribute VB_Name = "VersionsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Note: used to use a filenet listview here but to make the version
'retrieval code more interesting, using a standard listview

Private Sub cmdShowProps_Click()
    Dim oSelected As IDMObjects.Document
    Dim oProperty As IDMObjects.Property
    Dim itemX As ListItem
    Dim objID As String
    'get the version (1st item is latest ver, 2nd ver is next version, etc.)
    'note: an item in a version series is just a document object
    Set oSelected = MainForm.oDocument.Version.Series(MainForm.oDocument.Version.Series.Count - lvVersions.SelectedItem.SubItems(1) + 1)
    PropertiesForm.MsList.ListItems.Clear
    For Each oProperty In oSelected.Properties
        'Set the first column to the property name
        Set itemX = PropertiesForm.MsList.ListItems.Add(, , oProperty.PropertyDescription.Name)
        itemX.SubItems(1) = oProperty.FormatValue
        itemX.SubItems(2) = MainForm.FormatDataType(oProperty.PropertyDescription.TypeID)
        itemX.SubItems(3) = oProperty.PropertyDescription.GetState(idmPropSearchable)
        itemX.SubItems(4) = oProperty.PropertyDescription.GetState(idmPropMultiValue)
    Next
    PropertiesForm.Label2 = oSelected.Name
    PropertiesForm.Show vbModal, Me

End Sub

Private Sub Done_Click()
    VersionsForm.Hide
End Sub

Private Sub Form_Load()
    'add column headers
    lvVersions.ColumnHeaders.Clear
    lvVersions.ColumnHeaders.Add , , "Name"
    lvVersions.ColumnHeaders.Add , , "Version Number"
End Sub



Private Sub lvVersions_ItemClick(ByVal Item As ComctlLib.ListItem)
    cmdShowProps.Enabled = True
End Sub
