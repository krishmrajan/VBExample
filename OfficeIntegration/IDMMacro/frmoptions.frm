VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1725
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   3270
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1725
   ScaleWidth      =   3270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CheckBox TransferCheck 
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   600
      Value           =   1  'Checked
      Width           =   2500
   End
   Begin VB.CheckBox SaveChangesCheck 
      Caption         =   "&Save Changes"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Value           =   1  'Checked
      Width           =   2500
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public vbButton As VbMsgBoxResult

Function OptionsDialog(iDlgType As Integer, iSaveCheckValue As Integer, iTransferCheckValue As Integer) As VbMsgBoxResult
    vbButton = vbCancel
    
    Select Case iDlgType
        Case OPTIONS_DIALOG_ADD_TYPE
            frmOptions.Caption = OPTIONS_DIALOG_ADD_TITLE
            frmOptions.TransferCheck.Caption = OPTIONS_DIALOG_TRANS_CHECK_ADD_CAPTION
        Case OPTIONS_DIALOG_CHECKIN_TYPE
            frmOptions.Caption = OPTIONS_DIALOG_CHECKIN_TITLE
            frmOptions.TransferCheck.Caption = OPTIONS_DIALOG_TRANS_CHECK_CHECKIN_CAPTION
        Case Else
            OptionsDialog = vbCancel
            Exit Function
    End Select
    
    frmOptions.SaveChangesCheck.Value = iSaveCheckValue
    frmOptions.TransferCheck.Value = iTransferCheckValue
    
    frmOptions.Show 1
    
    iSaveCheckValue = frmOptions.SaveChangesCheck.Value
    iTransferCheckValue = frmOptions.TransferCheck.Value
    OptionsDialog = vbButton
    
    Unload Me
    
End Function

Private Sub CancelButton_Click()
    frmOptions.Hide
    vbButton = vbCancel
End Sub

Private Sub OKButton_Click()
    frmOptions.Hide
    vbButton = vbOK
End Sub


