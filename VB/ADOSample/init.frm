VERSION 5.00
Begin VB.Form frmInitialize 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Initialize ADO Connection"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdUnInit 
      Caption         =   "&Uninitialize"
      Height          =   375
      Left            =   2520
      TabIndex        =   4
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton cmdInit 
      Caption         =   "I&nitialize"
      Height          =   375
      Left            =   840
      TabIndex        =   3
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox txtCatName 
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   720
      Width           =   3615
   End
   Begin VB.OptionButton btnMez 
      Caption         =   "&Mezzanine"
      Height          =   495
      Left            =   1800
      TabIndex        =   1
      Top             =   1680
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.OptionButton btnIMS 
      Caption         =   "&IMS"
      Height          =   495
      Left            =   1800
      TabIndex        =   0
      Top             =   1320
      Width           =   1215
   End
End
Attribute VB_Name = "frmInitialize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public bInitialized As Boolean

Dim strProvider As String
Dim strDS As String
Dim strUser As String
Dim strPwd As String
Dim strPrompt As String
Dim strSystemType As String

Dim strDSName As String

Private Sub InitConnectSubStrings()
    strProvider = "provider=FnDBProvider;"
    'strDS = "data source=QALib^QAMezz;"
    strDS = "data source=newportqa;"
    
    strUser = "" '"user id=admin;"
    strPwd = "" '"password=;"
    strPrompt = "Prompt=1;"
    ' strSystemType = "SystemType=" & idmSysTypeDS & ";"
    ' strDSName = "QALib^QAMezz"
    strSystemType = "SystemType=" & idmSysTypeIS & ";"
    strDSName = "newportqa"
    
End Sub

Private Sub SetDSAndSysType()
    If txtCatName <> "" Then
        strDSName = txtCatName.Text
        dbGlobals.systemName = txtCatName.Text
        strDS = "data source=" & strDSName & ";"
    End If
    
    If btnIMS.Value = True Then
        dbGlobals.systemType = idmSysTypeIS
        strSystemType = "SystemType=" & idmSysTypeIS & ";"
    ElseIf btnMez.Value = True Then
        dbGlobals.systemType = idmSysTypeDS
        strSystemType = "SystemType=" & idmSysTypeDS & ";"
    End If
End Sub


Private Sub cmdInit_Click()
On Error GoTo Handle
    Dim strConnect As String
    Dim strGet As String
    
    SetDSAndSysType
    
    'strGet = dbGlobals.ds.Properties(0).Value
        
    strConnect = strDS & strSystemType '& "Group=General Users;"
    dbGlobals.ds.ConnectionString = strConnect
    dbGlobals.ds.Open
    
    bInitialized = True
    cmdInit.Enabled = False
    cmdUnInit.Enabled = True
    Hide
    Exit Sub
Handle:
    ShowError
End Sub

Private Sub cmdUnInit_Click()
    dbGlobals.ds.Close
    
    bInitialized = False
    cmdUnInit.Enabled = False
    cmdInit.Enabled = True
    Hide
End Sub

Private Sub Form_Load()
    bInitialized = False
    cmdInit.Enabled = True
    cmdUnInit.Enabled = False
    
    InitConnectSubStrings
    txtCatName.Text = strDSName
End Sub

