VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "ADO Test Application"
   ClientHeight    =   1440
   ClientLeft      =   132
   ClientTop       =   588
   ClientWidth     =   4320
   LinkTopic       =   "Form1"
   ScaleHeight     =   1440
   ScaleWidth      =   4320
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mnuConnection 
      Caption         =   "&Connection"
      Begin VB.Menu mnuInit 
         Caption         =   "&Re-initialize"
      End
      Begin VB.Menu mnuProperties 
         Caption         =   "&Properties"
      End
   End
   Begin VB.Menu mnuNewCommand 
      Caption         =   "&NewCommand"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    mnuInit.Enabled = True
    mnuProperties.Enabled = True
    mnuNewCommand.Enabled = False
    
    'We need to set this here to have any hope of talking
    'to the right provider before the connection is opened.
    dbGlobals.ds.Provider = "FnDBProvider"
    Set oErrManager = CreateObject("IDMError.ErrorManager")
    ' The first time around, take the user directly to the
    ' connection init logic
    mnuInit_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set oErrManager = Nothing
End Sub

Private Sub mnuInit_Click()
    frmInitialize.Show vbModal, Me
    If frmInitialize.bInitialized Then
        mnuProperties.Enabled = True
        mnuNewCommand.Enabled = True
    End If
End Sub
' Bring up the right command dialog box based on type
' of library
Private Sub mnuNewCommand_Click()
    If dbGlobals.systemType = idmSysTypeIS Then
        frmCommandIDMIS.Show vbModal, Me
    Else
        frmCommandIDMDS.Show vbModal, Me
    End If
End Sub
' Show the properties of the basic ADO connection
Private Sub mnuProperties_Click()
    Set frmConnectionProperties.obj = dbGlobals.ds
    frmConnectionProperties.Show vbModal, Me
End Sub
