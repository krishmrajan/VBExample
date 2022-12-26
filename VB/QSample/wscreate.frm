VERSION 5.00
Begin VB.Form WSCreate 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Create Workspace"
   ClientHeight    =   3810
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8940
   LinkTopic       =   "WSCreate"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3810
   ScaleWidth      =   8940
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   5040
      TabIndex        =   12
      Top             =   3360
      Width           =   1815
   End
   Begin VB.CommandButton btnSave 
      Caption         =   "&Save"
      Default         =   -1  'True
      Height          =   375
      Left            =   2040
      TabIndex        =   11
      Top             =   3360
      Width           =   1815
   End
   Begin VB.Frame frPermissions 
      Caption         =   "Permissions"
      Height          =   1935
      Left            =   240
      TabIndex        =   4
      Top             =   1200
      Width           =   8415
      Begin VB.TextBox txtRead 
         Height          =   285
         Left            =   960
         MaxLength       =   82
         TabIndex        =   6
         Top             =   360
         Width           =   7215
      End
      Begin VB.TextBox txtWrite 
         Height          =   285
         Left            =   960
         MaxLength       =   82
         TabIndex        =   8
         Top             =   840
         Width           =   7215
      End
      Begin VB.TextBox txtAX 
         Height          =   285
         Left            =   960
         MaxLength       =   82
         TabIndex        =   10
         Top             =   1320
         Width           =   7215
      End
      Begin VB.Label lblAX 
         Caption         =   "AX"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label lblWrite 
         Caption         =   "Write"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   840
         Width           =   735
      End
      Begin VB.Label lblRead 
         Caption         =   "Read"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.TextBox txtDescription 
      Height          =   495
      Left            =   1320
      MaxLength       =   800
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   600
      Width           =   7215
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   1320
      MaxLength       =   14
      TabIndex        =   1
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label lblDescription 
      Caption         =   "Description:"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   975
   End
   Begin VB.Label lblName 
      Caption         =   "Name:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "WSCreate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public isNew As Boolean    ' If is a new workspace
Dim newWS As IDMObjects.QueueWorkspace

Private Sub btnCancel_Click()
    If isNew Then
        QMaint.MainStatusBar.SimpleText = "Workspace creation cancelled."
    Else
        QMaint.MainStatusBar.SimpleText = "Workspace modification cancelled."
    End If
    QMaint.MainStatusBar.Refresh
    Set newWS = Nothing
    Unload Me
End Sub

Private Sub btnSave_Click()
    On Error GoTo ErrorHandler
    Screen.MousePointer = vbHourglass
    QMaint.MainStatusBar.SimpleText = "Saving workspace..."
    QMaint.MainStatusBar.Refresh
    If isNew Then
        newWS.Name = txtName
        Dim str As String
        str = txtDescription
        newWS.Description = txtDescription.Text
        newWS.Permissions.Item(1).GranteeName = txtRead
        newWS.Permissions.Item(2).GranteeName = txtWrite
        newWS.Permissions.Item(3).GranteeName = txtAX
        newWS.SaveNew
        QMaint.cmbWorkspace.AddItem (txtName)
        Set newWS = Nothing
    Else
        Dim oldName As String
        
        oldName = QMaint.oWorkspace.Name
        QMaint.oWorkspace.Name = txtName
        QMaint.oWorkspace.Description = txtDescription
        QMaint.oWorkspace.Permissions.Item(1).GranteeName = txtRead
        QMaint.oWorkspace.Permissions.Item(2).GranteeName = txtWrite
        QMaint.oWorkspace.Permissions.Item(3).GranteeName = txtAX
        QMaint.oWorkspace.Save
        If txtName <> oldName Then
            ' Handle renames
            Dim lstCount As Integer
            Dim inx As Integer
            
            lstCount = QMaint.cmbWorkspace.listCount
            For inx = 0 To lstCount - 1
                If QMaint.cmbWorkspace.List(inx) = oldName Then
                    QMaint.cmbWorkspace.RemoveItem (inx)
                    Exit For
                End If
            Next inx
            QMaint.cmbWorkspace.AddItem (txtName)
            QMaint.cmbWorkspace.Text = txtName
        End If
    End If
    QMaint.MainStatusBar.SimpleText = "Workspace saved."
    QMaint.MainStatusBar.Refresh
    Screen.MousePointer = vbDefault
    Unload Me
    Exit Sub
ErrorHandler:
    oErrorLog.logFNError errWarning, "Error saving workspace.", "Error saving workspace."
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    If isNew Then
        Me.Caption = "Create Workspace"
        Set newWS = QMaint.oLibrary.CreateObject(idmObjTypeQueueWorkspace, "")
        txtRead = newWS.Permissions.Item(1).GranteeName
        txtWrite = newWS.Permissions.Item(2).GranteeName
        txtAX = newWS.Permissions.Item(3).GranteeName
    Else
        Me.Caption = "Modify Workspace"
        txtName = QMaint.oWorkspace.Name
        txtDescription = QMaint.oWorkspace.Description
        txtRead = QMaint.oWorkspace.Permissions.Item(1).GranteeName
        txtWrite = QMaint.oWorkspace.Permissions.Item(2).GranteeName
        txtAX = QMaint.oWorkspace.Permissions.Item(3).GranteeName
    End If
    btnSave.Enabled = (txtName <> "")
End Sub

Private Sub txtName_Change()

    txtName.Text = Trim(txtName.Text)
    btnSave.Enabled = (txtName.Text <> "")
    
End Sub
