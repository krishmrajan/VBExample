VERSION 5.00
Begin VB.Form frmLogon 
   Caption         =   "Form2"
   ClientHeight    =   1995
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6060
   LinkTopic       =   "Form2"
   ScaleHeight     =   1995
   ScaleWidth      =   6060
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   495
      Left            =   2520
      TabIndex        =   2
      Top             =   1320
      Width           =   1095
   End
   Begin VB.ComboBox CmbLibraries 
      Height          =   315
      Left            =   840
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   840
      Width           =   4335
   End
   Begin VB.Label Label1 
      Caption         =   "Select library:"
      Height          =   255
      Left            =   840
      TabIndex        =   1
      Top             =   480
      Width           =   2055
   End
End
Attribute VB_Name = "frmLogon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Neighborhood As New IDMObjects.Neighborhood
Dim Libraries As IDMObjects.ObjectSet
Dim LibIndices() As Integer

Private Sub Command1_Click()
For Each oLib In Libraries
    If oLib.Label = CmbLibraries.Text Then
        Exit For
    End If
Next

MousePointer = vbHourglass
'make sure we are logged on to the selected Library
oLib.Logon "", "", "", idmLogonOptWithUI
MousePointer = vbArrow
Unload Me
End Sub

Private Sub Form_Load()
Set Libraries = Neighborhood.Libraries
Dim Library As IDMObjects.Library
ReDim LibIndices(Libraries.Count)
CmbLibraries.Clear
' Just show document libraries...
For Each Library In Libraries
    If Library.SystemType = idmSysTypeDS Then
        CmbLibraries.AddItem Library.Label
    End If
Next
If CmbLibraries.ListCount = 0 Then
    MsgBox "You cannot run the application because no server libraries are available."
    Unload Form1
Else
    CmbLibraries.ListIndex = 0
End If
End Sub

