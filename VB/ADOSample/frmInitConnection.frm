VERSION 5.00
Begin VB.Form frmInitialize 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Initialize ADO Connection"
   ClientHeight    =   2250
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2250
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox CmbLibraries 
      Height          =   315
      Left            =   600
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   600
      Width           =   3735
   End
   Begin VB.CommandButton cmdInit 
      Caption         =   "I&nitialize"
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Libraries:"
      Height          =   255
      Left            =   600
      TabIndex        =   2
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "frmInitialize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim oHood As New IDMObjects.Neighborhood
Dim oLibraries As IDMObjects.ObjectSet
Dim oCurrentLib As IDMObjects.Library

Public bInitialized As Boolean

Dim strDS As String
Dim strUser As String
Dim strPwd As String
Dim strPrompt As String
Dim strSystemType As String
Const strProvider = "provider=FnDBProvider;"
' Logic to handle Library selection and logon
Private Sub CmbLibraries_Click()
Dim iTmp As Integer
Set oCurrentLib = oLibraries(CmbLibraries.ListIndex + 1)
If oCurrentLib.LogonId = 0 Then
    Call oCurrentLib.Logon("", "", "", idmLogonOptWithUI)
End If
' Now get the ADO string stuff set up for this library
strDS = "data source=" & oCurrentLib.Name & ";"
strUser = "" '"user id=admin;"
strPwd = "" '"password=;"
strPrompt = "Prompt=1;"
iTmp = oCurrentLib.systemType    ' Force type coercion
dbGlobals.systemType = iTmp
dbGlobals.systemName = oCurrentLib.Name
strSystemType = "SystemType=" & iTmp & ";"
strDS = "data source=" & oCurrentLib.Name & ";"

Exit Sub
End Sub
' Handle the Initialize button command
Private Sub cmdInit_Click()
On Error GoTo Handle
Dim strConnect As String
Dim strGet As String
' If we are restarting, close the ADO connection
If bInitialized Then
    dbGlobals.ds.Close
End If
' Construct the connect string with the correct name
' and system type
strConnect = strDS & strSystemType '& "Group=General Users;"
dbGlobals.ds.ConnectionString = strConnect
dbGlobals.ds.Open
    
bInitialized = True
Hide
Exit Sub

Handle:
ShowError
End Sub

' Fill the libraries combobox and get rolling
Private Sub Form_Load()
Dim oLibrary As IDMObjects.Library
CmbLibraries.Clear
Set oLibraries = oHood.Libraries
For Each oLibrary In oLibraries
    CmbLibraries.AddItem (oLibrary.Name)
Next
' CmbLibraries.ListIndex = 0
bInitialized = False
End Sub

