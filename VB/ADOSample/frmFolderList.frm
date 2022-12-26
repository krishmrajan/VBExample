VERSION 5.00
Begin VB.Form frmFolderList 
   Caption         =   "Results"
   ClientHeight    =   6675
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5715
   LinkTopic       =   "Form1"
   ScaleHeight     =   6675
   ScaleWidth      =   5715
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lbFolderNames 
      Height          =   5520
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   5052
   End
   Begin VB.CommandButton btnCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2760
      TabIndex        =   1
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton btnOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   1440
      TabIndex        =   0
      Top             =   6000
      Width           =   1215
   End
End
Attribute VB_Name = "frmFolderList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Form for browsing folders and selecting folders

Public strFolderName As String
Public strFolderId As Long
Dim oLib As New IDMObjects.Library
Dim vFolderIds() As Variant

Private Sub btnCancel_Click()
    strFolderName = ""
    Hide
End Sub

Private Sub btnOK_Click()
    Dim inx As Integer
    Dim oLib As New IDMObjects.Library
    oLib.Name = dbGlobals.systemName
    oLib.systemType = dbGlobals.systemType
    
    inx = lbFolderNames.ListIndex
    If inx < 0 Then
        strFolderName = ""
        strFolderId = -1
    Else
        strFolderName = lbFolderNames.List(inx)
        ' The following isn't working
        ' Dim oFolder As IDMObjects.folder
        ' Set oFolder = oLib.GetObject(idmObjTypeFolder, strFolderName)
        ' strFolderId = oFolder.ID
        strFolderId = vFolderIds(inx)
            
    End If
    Hide
End Sub

Private Sub Form_Load()
    Dim folders As IDMObjects.ObjectSet
    Dim folder  As IDMObjects.folder
    Dim folderNode As Node
    Dim iCnt As Integer
    oLib.Name = dbGlobals.systemName
    oLib.systemType = dbGlobals.systemType
    Set folders = oLib.TopFolders
    ReDim vFolderIds(folders.Count)
    iCnt = 0
    ' We need to keep track of the name->id mapping here
    ' to work around a temporary bug; this will go away
    Screen.MousePointer = vbHourglass
    For Each folder In folders
        lbFolderNames.AddItem "/" + folder.Name
        vFolderIds(iCnt) = folder.ID
        iCnt = iCnt + 1
    Next folder
    Screen.MousePointer = vbDefault
End Sub



Private Sub lbFolderNames_DblClick()
    btnOK_Click
End Sub
