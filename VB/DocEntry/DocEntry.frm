VERSION 5.00
Object = "{A9983B40-CE52-11CF-AE75-00A0248802BA}#1.0#0"; "fnviewer.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "comdlg32.ocx"
Begin VB.Form FormMain 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Browsing Local Document Files"
   ClientHeight    =   6345
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6735
   LinkTopic       =   "Form1"
   ScaleHeight     =   6345
   ScaleWidth      =   6735
   StartUpPosition =   3  'Windows Default
   Begin IDMViewerCtrl.IDMViewerCtrl ViewerCtrl1 
      Height          =   6135
      Left            =   3240
      TabIndex        =   12
      Top             =   480
      Width           =   6015
      _Version        =   65536
      _ExtentX        =   10610
      _ExtentY        =   10821
      _StockProps     =   161
      BackColor       =   -2147483633
      BackColor       =   -2147483633
      SystemType      =   -24748
   End
   Begin VB.CommandButton BtnCancel 
      Height          =   615
      Left            =   1800
      Picture         =   "DocEntry.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Cancel/Exit"
      Top             =   5070
      Width           =   855
   End
   Begin VB.CommandButton BtnNext 
      Height          =   615
      Left            =   1800
      Picture         =   "DocEntry.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Next Doc"
      Top             =   1800
      Width           =   855
   End
   Begin VB.CommandButton BtnPrevious 
      Height          =   615
      Left            =   1800
      Picture         =   "DocEntry.frx":074C
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Previous Doc"
      Top             =   2610
      Width           =   855
   End
   Begin VB.CommandButton BtnCommit 
      Height          =   615
      Left            =   1800
      Picture         =   "DocEntry.frx":0B8E
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Commit"
      Top             =   3435
      Width           =   855
   End
   Begin VB.CommandButton BtnRestart 
      Height          =   615
      Left            =   1800
      Picture         =   "DocEntry.frx":0FD0
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Restart"
      Top             =   4245
      Width           =   855
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   10680
      Top             =   8160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327681
      DialogTitle     =   "Select local files for committal..."
   End
   Begin VB.CommandButton BtnRotate 
      Height          =   615
      Left            =   5280
      Picture         =   "DocEntry.frx":1412
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7200
      Width           =   735
   End
   Begin VB.CommandButton BtnDone 
      Height          =   615
      Left            =   1800
      Picture         =   "DocEntry.frx":171C
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Finish"
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton BtnZoomIn 
      Height          =   615
      Left            =   6600
      Picture         =   "DocEntry.frx":1B5E
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Zoom In"
      Top             =   7200
      Width           =   735
   End
   Begin VB.CommandButton BtnZoomOut 
      Height          =   615
      Left            =   7920
      Picture         =   "DocEntry.frx":1E68
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Zoom Out"
      Top             =   7200
      Width           =   735
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Page"
      Height          =   1095
      Left            =   3240
      TabIndex        =   4
      Top             =   6960
      Width           =   6015
      Begin VB.CommandButton BtnRotateLeft 
         Height          =   615
         Left            =   840
         Picture         =   "DocEntry.frx":2172
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Document"
      Height          =   5175
      Left            =   1320
      TabIndex        =   10
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Image ImgCommitState 
      BorderStyle     =   1  'Fixed Single
      Height          =   780
      Left            =   1320
      Picture         =   "DocEntry.frx":247C
      Stretch         =   -1  'True
      Top             =   480
      Width           =   900
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   780
      Left            =   2160
      Picture         =   "DocEntry.frx":2BBE
      Stretch         =   -1  'True
      Top             =   480
      Width           =   900
   End
End
Attribute VB_Name = "FormMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim NavForward As Boolean
Dim RotateAmount As Integer
Dim CommCount As Integer       ' Count of docs marked for committal


Public Sub LoadFiles()
    Dim NullOffset As Integer
    Dim FileNameList As String
    Dim Pathpart As String
    Dim inx As Integer
    ' Trap cancel as an error
    CommonDialog1.CancelError = True
    On Error GoTo ErrHandler
    ' Set flags
    CommonDialog1.Flags = cdlOFNHideReadOnly Or cdlOFNExplorer Or cdlOFNAllowMultiselect
    ' Set filters, we only demo jpg files
    CommonDialog1.Filter = "JPEG files (*.jpg)|*.jpg"
                
    ' Specify default filter
    CommonDialog1.FilterIndex = 1
    ' Display the Open dialog box
    CommonDialog1.ShowOpen
    ' Display name of selected file
    FileNameList = CommonDialog1.FileName
    ' We may have multiple names, separated by null
    NullOffset = InStr(FileNameList, Chr(0))
    inx = 0
    TotalDocs = 0
    CommCount = 0
    If NullOffset > 0 Then
        Pathpart = Left(FileNameList, NullOffset - 1) + "\"
        FileNameList = Mid(FileNameList, NullOffset + 1)
        While Len(FileNameList) > 0
            NullOffset = InStr(FileNameList, Chr(0))
            If NullOffset > 0 Then
                DocList(inx).FileName = Pathpart + Left(FileNameList, NullOffset - 1)
                DocList(inx).CommitFlag = UnDecided
                TotalDocs = TotalDocs + 1
                FileNameList = Mid(FileNameList, NullOffset + 1)
            Else
                DocList(inx).FileName = Pathpart + FileNameList
                DocList(inx).CommitFlag = UnDecided
                TotalDocs = TotalDocs + 1
                FileNameList = ""
            End If
            If inx = ArraySz Then
                ArraySz = 2 * ArraySz
                ReDim Preserve DocList(ArraySz)
                ReDim FinalList(ArraySz)
                ReDim FolderList(ArraySz)
            End If
            inx = inx + 1
        Wend
    Else
        TotalDocs = TotalDocs + 1
        DocList(0).FileName = FileNameList
        DocList(0).CommitFlag = UnDecided
    End If
        
    Exit Sub
    
ErrHandler:
    'User pressed the Cancel button
    'He may already have Docs loaded up, so leave them
    ' alone
    Exit Sub
End Sub

Private Sub AdjustButtons(ByVal DocInx As Integer)
If ViewerCtrl1.IsOperationSupported(idmOpZoomInOut) Then
    BtnZoomIn.Enabled = True
    BtnZoomOut.Enabled = True
Else
    BtnZoomIn.Enabled = False
    BtnZoomOut.Enabled = False
End If
If ViewerCtrl1.IsOperationSupported(idmOpRotation) Then
    BtnRotate.Enabled = True
Else
    BtnRotate.Enabled = False
End If
If CommCount > 0 Then
    BtnDone.Enabled = True
Else
    BtnDone.Enabled = False
End If
If DocList(DocInx).CommitFlag = Commit Then
    ' BtnCommit.Caption = "Decommit"
    ImgCommitState.Picture = LoadPicture(HomeDirectory + "\Checkmrk.ico")
Else
    If DocList(DocInx).CommitFlag = DontCommit Then
        ' BtnCommit.Caption = "Commit"
        ImgCommitState.Picture = LoadPicture(HomeDirectory + "\X.ico")
    Else
        ' BtnCommit.Caption = "Commit"
        ImgCommitState.Picture = LoadPicture(HomeDirectory + "\Question.ico")
    End If
End If

End Sub

Private Sub LoadDocument(ByVal inx As Integer)
ViewerCtrl1.DocumentFilename = DocList(inx).FileName
RotateAmount = 0
ViewerCtrl1.Rotation = 0
ViewerCtrl1.Brightness = idmBrightnessEnhance
ViewerCtrl1.Refresh
Call AdjustButtons(inx)
End Sub

Private Sub BtnCommit_Click()
If DocList(CurrentDocInx).CommitFlag = Commit Then
    ' Changed our mind - decommit
    DocList(CurrentDocInx).CommitFlag = DontCommit
    ' Clean up doc and properties
    Set FinalList(CurrentDocInx) = Nothing
    FolderList(CurrentDocInx) = ""
    CommCount = CommCount - 1
Else
    PropertyForm.Tag = "0"
    PropertyForm.Show vbModal
    If PropertyForm.Tag = "1" Then
        DocList(CurrentDocInx).CommitFlag = Commit
        CommCount = CommCount + 1
    Else
        ' the user cancelled out of the property stuff
        DocList(CurrentDocInx).CommitFlag = UnDecided
    End If
End If
Call AdjustButtons(CurrentDocInx)
If TotalDocs > 1 Then
    If NavForward Then
        Call BtnNext_Click
    Else
        Call BtnPrevious_Click
    End If
End If
End Sub

Private Sub BtnCancel_Click()
Unload Me
End Sub

Private Sub Pause(ByVal Secs As Integer)
Dim Start As Variant
Start = Timer
Do While Timer < Start + Secs
    DoEvents
Loop
End Sub
Private Sub BtnDone_Click()
Dim inx As Integer
Dim oFolder As IDMObjects.Folder

Me.Visible = False

' Try out the animation
' CommitForm.Animation1.Open ("d:\devstudio\vb\graphics\avis\filecopy.avi")
CommitForm.Animation1.Open (HomeDirectory + "\commit.avi")
CommitForm.Show
CommitForm.ProgressBar1.Min = 0
CommitForm.txtMin.Caption = 0
CommitForm.txtMax.Caption = CommCount
CommitForm.ProgressBar1.Max = CommCount
CommitForm.SetFocus
DoEvents
On Error GoTo Errs
For inx = 0 To TotalDocs - 1
    If DocList(inx).CommitFlag = Commit Then
        CommitForm.ProgressBar1.Value = _
            CommitForm.ProgressBar1.Value + 1
        CommitForm.Text1.Text = DocList(inx).FileName
        CommitForm.SetFocus
        CommitForm.Animation1.Play
        If Online Then
            Call FinalList(inx).SaveNew(DocList(inx).FileName, idmDocSaveNewKeep)
            '  Call FinalList(inx).SaveNew(DocList(inx).FileName)
            If FolderList(inx) <> "" Then
                Set oFolder = CurrentLib.GetObject(idmObjTypeFolder, _
                    FolderList(inx))
                Call oFolder.File(FinalList(inx))
            End If
              
        Else
            Call Pause(10)
        End If
        CommitForm.Animation1.Stop
        Set FinalList(inx) = Nothing
    End If
Next
MsgBox ("Operation Complete")
Errs:
Call ShowError
Me.Visible = True
Unload CommitForm
Unload Me
End Sub

Private Sub BtnNext_Click()
CurrentDocInx = CurrentDocInx + 1
If CurrentDocInx > TotalDocs - 2 Then
    BtnNext.Enabled = False
    NavForward = False
Else
    NavForward = True
End If
BtnPrevious.Enabled = True
Call LoadDocument(CurrentDocInx)
End Sub

Private Sub BtnPrevious_Click()
CurrentDocInx = CurrentDocInx - 1
If CurrentDocInx = 0 Then
    BtnPrevious.Enabled = False
    NavForward = True
Else
    NavForward = False
End If
BtnNext.Enabled = True
Call LoadDocument(CurrentDocInx)
End Sub

Private Sub BtnRestart_Click()
Call LoadFiles
If TotalDocs > 0 Then
    CurrentDocInx = 0
    Call LoadDocument(CurrentDocInx)
    BtnPrevious.Enabled = False
    NavForward = True
    If TotalDocs > 1 Then
        BtnNext.Enabled = True
    Else
        BtnNext.Enabled = False
    End If
End If
End Sub

Private Sub BtnRotate_Click()
RotateAmount = RotateAmount + 90
If RotateAmount = 360 Then
    RotateAmount = 0
End If
ViewerCtrl1.Rotation = RotateAmount
End Sub

Private Sub BtnRotateLeft_Click()
RotateAmount = RotateAmount - 90
If RotateAmount < 0 Then
    RotateAmount = 360 + RotateAmount
End If
ViewerCtrl1.Rotation = RotateAmount
End Sub

Private Sub BtnZoomIn_Click()
ViewerCtrl1.ZoomIn
End Sub

Private Sub BtnZoomOut_Click()
ViewerCtrl1.ZoomOut
End Sub

Private Sub Form_Load()
Me.WindowState = vbMaximized
DocList(0).FileName = ""
TotalDocs = 0
CommCount = 0
Call BtnRestart_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload PropertyForm
Set oErrManager = Nothing
End Sub
