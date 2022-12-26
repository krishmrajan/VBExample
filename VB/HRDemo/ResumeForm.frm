VERSION 5.00
Begin VB.Form ResumeForm 
   Caption         =   "Resumes"
   ClientHeight    =   6435
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7980
   LinkTopic       =   "Form2"
   ScaleHeight     =   6435
   ScaleWidth      =   7980
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Interviews this week..."
      Height          =   495
      Left            =   2640
      TabIndex        =   2
      Top             =   3720
      Width           =   2895
   End
   Begin VB.CommandButton AddResume 
      Caption         =   "Add a Resume..."
      Height          =   495
      Left            =   2640
      TabIndex        =   1
      Top             =   2280
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Browse Resumes..."
      Height          =   495
      Left            =   2640
      TabIndex        =   0
      Top             =   840
      Width           =   2895
   End
End
Attribute VB_Name = "ResumeForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim oLib As New IDMObjects.Library
Private Sub Command1_Click()
    oLib.Name = "DefaultIMS:AIIM2:FileNet"
    oLib.SystemType = idmSysTypeIMS
    oLib.Logon , , , idmLogonOptWithUI
    Dim oFolder As IDMObjects.Folder
    Set oFolder = oLib.GetObject(idmObjTypeFolder, "/Lauren")
    BrowseNewResumeForm.ListView1.AddItems oFolder.GetContents(idmFolderContentDocument), -1
    BrowseNewResumeForm.Show 1, ResumeForm
End Sub

