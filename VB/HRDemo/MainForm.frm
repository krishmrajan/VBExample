VERSION 5.00
Begin VB.Form MainForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Human Resources"
   ClientHeight    =   8715
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9300
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "MainForm.frx":0000
   ScaleHeight     =   8715
   ScaleWidth      =   9300
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   810
      Left            =   0
      Picture         =   "MainForm.frx":9C844
      ScaleHeight     =   780
      ScaleWidth      =   9750
      TabIndex        =   9
      Top             =   0
      Width           =   9780
   End
   Begin VB.CommandButton btnStock 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Stock Purchase Plan"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3120
      MaskColor       =   &H008080FF&
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6960
      Width           =   4695
   End
   Begin VB.CommandButton btn401K 
      BackColor       =   &H00C0E0FF&
      Caption         =   "401K"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3120
      MaskColor       =   &H008080FF&
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5240
      Width           =   4695
   End
   Begin VB.CommandButton btnBenefits 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Benefits"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3120
      MaskColor       =   &H008080FF&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3480
      Width           =   4695
   End
   Begin VB.PictureBox Picture4 
      Height          =   1575
      Left            =   480
      Picture         =   "MainForm.frx":B5508
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   4
      Top             =   6720
      Width           =   1575
   End
   Begin VB.PictureBox Picture3 
      AutoSize        =   -1  'True
      Height          =   1560
      Left            =   480
      Picture         =   "MainForm.frx":B805C
      ScaleHeight     =   1500
      ScaleWidth      =   1500
      TabIndex        =   3
      Top             =   4920
      Width           =   1560
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      Height          =   1560
      Left            =   480
      Picture         =   "MainForm.frx":BABB0
      ScaleHeight     =   1500
      ScaleWidth      =   1500
      TabIndex        =   2
      Top             =   3120
      Width           =   1560
      Begin VB.Frame Frame1 
         Caption         =   "Frame1"
         Height          =   1575
         Left            =   1560
         TabIndex        =   5
         Top             =   0
         Width           =   15
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1530
      Left            =   480
      Picture         =   "MainForm.frx":BD704
      ScaleHeight     =   1500
      ScaleWidth      =   1500
      TabIndex        =   1
      Top             =   1440
      Width           =   1530
   End
   Begin VB.CommandButton btnResumes 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Resumes"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3120
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1800
      Width           =   4695
   End
   Begin VB.Line Line4 
      BorderColor     =   &H000080FF&
      BorderWidth     =   3
      X1              =   2160
      X2              =   8760
      Y1              =   6720
      Y2              =   6720
   End
   Begin VB.Line Line3 
      BorderColor     =   &H000080FF&
      BorderWidth     =   3
      X1              =   2160
      X2              =   8760
      Y1              =   4920
      Y2              =   4920
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000080FF&
      BorderWidth     =   3
      X1              =   2160
      X2              =   8760
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000080FF&
      BorderWidth     =   3
      X1              =   2160
      X2              =   8760
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Options"
      Visible         =   0   'False
      Begin VB.Menu mnuChangeOptions 
         Caption         =   "Change options"
      End
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btn401K_Click()
MsgBox ("You're too young to retire - go back to work...")
End Sub

' Fire up the benefits UI
Private Sub btnBenefits_Click()
    BenefitForm.Show 1, MainForm
End Sub
' Fire up the resumes UI
Private Sub btnResumes_Click()
    ResumeForm.Show 1, MainForm
End Sub
' Routine for restoring Registry settings dealing with
' IDMLibrary information
Private Sub RestoreSettings(ByVal bForceUi As Boolean)
If bForceUi Then
    gfSettings.Show vbModal
Else
    Call goPersist.GetSettings(gsAppName, gsSectionName, gfSettings)
End If
End Sub
' Logon to both libs if necessary
Private Function ConnectToLibraries() As Boolean
On Error GoTo ErrorHandler

RestoreSettings (False)
' Make sure we have some valid library ID's
If gfSettings.txtIMSLibName = "" Or gfSettings.txtMZLibName = "" Then
    RestoreSettings (True)
End If
If gfSettings.txtIMSLibName = "" Or gfSettings.txtMZLibName = "" Then
    MsgBox ("You must first set the runtime parameters!")
    ConnectToLibraries = False
    Exit Function
End If
' We may have been here before, so clean up any old library
' connections
If gbDSLogOff Then
    goDSLib.Logoff
End If
Set goDSLib = Nothing
If gbISLogOff Then
    goISLib.Logoff
End If
Set goISLib = Nothing
' Hook up to the Mezzanine library
goDSLib.SystemType = idmSysTypeDS
goDSLib.Name = gfSettings.txtMZLibName
If Not goDSLib.GetState(idmLibraryLoggedOn) Then
    goDSLib.Logon gfSettings.txtMZUser, gfSettings.txtMZPassword, , idmLogonOptNoUI
    gbDSLogOff = True
Else
    gbDSLogOff = False
End If
' Hook up to the IMS library
goISLib.SystemType = idmSysTypeIS
goISLib.Name = gfSettings.txtIMSLibName

If Not goISLib.GetState(idmLibraryLoggedOn) Then
    goISLib.Logon gfSettings.txtIMSUser, gfSettings.txtIMSPassword, , idmLogonOptNoUI
    gbISLogOff = True
Else
    gbISLogOff = False
End If
ConnectToLibraries = True
Exit Function
ErrorHandler:
    MsgBox Err.Description
    ConnectToLibraries = False
End Function


Private Sub btnStock_Click()
MsgBox ("Buy some - make a fortune...")
End Sub

Private Sub Form_Load()
gbDSLogOff = False
gbISLogOff = False
If Not ConnectToLibraries Then
    RestoreSettings (True)
    End
End If
End Sub
' Handle the right mouse button context menu
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then  ' Check if right mouse button
                                ' was clicked.
    PopupMenu mnuOptions   ' Display the Options menu as a
                                ' pop-up menu.
    ConnectToLibraries
End If

End Sub
' Before termination, log off the libraries
Private Sub Form_Unload(Cancel As Integer)
If goISLib.GetState(idmLibraryLoggedOn) Then
    goISLib.Logoff
End If
If goDSLib.GetState(idmLibraryLoggedOn) Then
    goDSLib.Logoff
End If
End
End Sub
' Right mouse button menu action
Private Sub mnuOptions_Click()
Call goPersist.GetSettings(gsAppName, gsSectionName, gfSettings)
gfSettings.Show vbModal
End Sub
