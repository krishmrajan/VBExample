VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form QMaint 
   Caption         =   "QSample"
   ClientHeight    =   7935
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   7935
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.StatusBar MainStatusBar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   23
      Top             =   7560
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   661
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton QueryButton 
      Caption         =   "Query"
      Enabled         =   0   'False
      Height          =   375
      Left            =   8520
      TabIndex        =   10
      Top             =   6840
      Width           =   1455
   End
   Begin VB.CommandButton UnbusyButton 
      Caption         =   "Reset busy"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1800
      TabIndex        =   6
      Top             =   6840
      Width           =   1455
   End
   Begin VB.CommandButton BusyButton 
      Caption         =   "Set busy"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   6840
      Width           =   1455
   End
   Begin VB.CommandButton DeleteButton 
      Caption         =   "Delete Entry"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6840
      TabIndex        =   9
      Top             =   6840
      Width           =   1455
   End
   Begin VB.Frame QueueInfoFrame 
      Caption         =   "Queue Information"
      Height          =   5535
      Left            =   120
      TabIndex        =   18
      Top             =   1200
      Width           =   11655
      Begin MSFlexGridLib.MSFlexGrid grdQueueData 
         Height          =   4815
         Left            =   120
         TabIndex        =   17
         Top             =   600
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   8493
         _Version        =   393216
         Rows            =   1
         Cols            =   1
         FixedRows       =   0
         FixedCols       =   0
         BackColorBkg    =   12632256
         FocusRect       =   0
         SelectionMode   =   1
         AllowUserResizing=   1
      End
      Begin VB.Label QueueInfoLabel 
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   5415
      End
   End
   Begin VB.CommandButton InsertButton 
      Caption         =   "Insert Entry"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3480
      TabIndex        =   7
      Top             =   6840
      Width           =   1455
   End
   Begin VB.CommandButton EditButton 
      Caption         =   "Edit Entry"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5160
      TabIndex        =   8
      Top             =   6840
      Width           =   1455
   End
   Begin VB.Frame QueueFrame 
      Caption         =   "Queue Selection"
      Height          =   975
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   11655
      Begin VB.Frame frQueryResult 
         Caption         =   "Workspace Query Results"
         Height          =   750
         Left            =   8460
         TabIndex        =   21
         Top             =   150
         Width           =   3015
         Begin VB.ComboBox cmbQueues 
            Enabled         =   0   'False
            Height          =   315
            Left            =   630
            Style           =   2  'Dropdown List
            TabIndex        =   22
            Top             =   270
            Width           =   1935
         End
      End
      Begin VB.ComboBox cmbQueue 
         Height          =   315
         Left            =   3840
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   600
         Width           =   1935
      End
      Begin VB.ComboBox cmbWorkspace 
         Height          =   315
         Left            =   3840
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   1935
      End
      Begin VB.ComboBox cmbLibraries 
         Height          =   315
         Left            =   720
         TabIndex        =   0
         Top             =   240
         Width           =   1935
      End
      Begin VB.CheckBox SystemCheck 
         Alignment       =   1  'Right Justify
         Caption         =   "Show System Fields"
         Height          =   255
         Left            =   6120
         TabIndex        =   4
         Top             =   600
         Width           =   1935
      End
      Begin VB.TextBox maxRetText 
         Height          =   285
         Left            =   7440
         TabIndex        =   3
         Text            =   "500"
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "Library:"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Workspace:"
         Height          =   255
         Left            =   2760
         TabIndex        =   15
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Queue:"
         Height          =   255
         Left            =   3120
         TabIndex        =   14
         Top             =   600
         Width           =   615
      End
      Begin VB.Label maxRetLabel 
         Caption         =   "Max to retrieve:"
         Height          =   255
         Left            =   6120
         TabIndex        =   13
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.CommandButton ViewButton 
      Caption         =   "View Document"
      Enabled         =   0   'False
      Height          =   375
      Left            =   10200
      TabIndex        =   11
      Top             =   6840
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "Description:"
      Height          =   255
      Left            =   960
      TabIndex        =   16
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Menu mnuWS 
      Caption         =   "&Workspace"
      Begin VB.Menu mnuWSNew 
         Caption         =   "&New..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuWSModify 
         Caption         =   "&Modify..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuWSQuery 
         Caption         =   "&Query..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuWSCopy 
         Caption         =   "&Copy..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuWSRename 
         Caption         =   "&Rename"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuWSDelete 
         Caption         =   "&Delete..."
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuQueue 
      Caption         =   "&Queue"
      Begin VB.Menu mnuQueueNew 
         Caption         =   "&New..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuQueueModify 
         Caption         =   "&Modify..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuQueueQuery 
         Caption         =   "&Query..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuQueueCopy 
         Caption         =   "&Copy..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuQueueRename 
         Caption         =   "&Rename..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuQueueDelete 
         Caption         =   "&Delete..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuQueuePrint 
         Caption         =   "&Print..."
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About QSample"
      End
   End
End
Attribute VB_Name = "QMaint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public oHood As New IDMObjects.Neighborhood
Public oLibrary As New IDMObjects.Library       ' the library for access to the IMS
Public oWorkspace As IDMObjects.QueueWorkspace  ' queue workspace we are working on
Public oQueue As IDMObjects.queue               ' queue we are working on
Public iDocIndex As Integer                     ' the index of document ID column in the GridControl
Public iSelectedRow As Integer                  ' the row selected in the grid control
Public cQueueEntries As New Collection          ' a collection of the QueueEntries retrieved
Public iFormHeight As Integer                   ' the height of this form
Public iFormWidth As Integer                        ' the width of this form
Public iMinFormHeight As Integer                ' the minimun height of this form
Public oLocalCache As New IDMObjects.LocalCache ' the local cache of documents to display
Public oServerCache As IDMObjects.ServerCache   ' The server cache of documents to display
Public vSession As Variant                      ' the session for the local cache
Public oQueueQuerySpec As IDMObjects.QueueQuerySpecification   ' the queue query spec object
Public sPrevWSName As String                    ' the previous queue workspace name
Const iMinFormWidth = 12000                     ' the minimum width of this form
Const F_DELAY_DEFAULT = "12/31/1899"            ' the default time for F_Delay system field
Const SRC_FILE_NAME = "QMAINT.FRM"

'Handle the user's library selection, and get logged on
Private Sub cmbLibraries_Click()
Dim oLib As IDMObjects.Library
On Error GoTo ErrorHandler
Screen.MousePointer = vbHourglass
Me.cmbWorkspace.Enabled = False
Me.cmbQueue.Enabled = False
Me.mnuWSNew.Enabled = False
Me.mnuWSDelete.Enabled = False
Me.mnuWSModify.Enabled = False
Me.mnuWSQuery.Enabled = False
Me.mnuWSCopy.Enabled = False
Me.mnuWSRename.Enabled = False
Me.mnuQueueNew.Enabled = False
Me.mnuQueueDelete.Enabled = False
Me.mnuQueuePrint.Enabled = False
Me.mnuQueueModify.Enabled = False
Me.mnuQueueQuery.Enabled = False
Me.mnuQueueCopy.Enabled = False
Me.mnuQueueRename.Enabled = False
sPrevWSName = ""
For Each oLib In oHood.Libraries
    If oLib.Label = cmbLibraries Then
        oLib.Logon , , , idmLogonOptWithUI
        Exit For
    End If
Next
If oLib.GetState(idmLibraryLoggedOn) Then
    Set oLibrary = oLib
    Dim wsList As IDMObjects.ObjectSet
    Dim ws As IDMObjects.QueueWorkspace
    Set wsList = oLibrary.FilterQueueWorkspaces("")
    Me.cmbWorkspace.Clear
    For Each ws In wsList
        Me.cmbWorkspace.AddItem ws.Name
    Next
    Set wsList = Nothing
    Me.cmbWorkspace.Enabled = True
    Me.mnuWSNew.Enabled = True
    Set oServerCache = oLibrary.GetObject(idmObjTypeServerCache, "")
End If
Screen.MousePointer = vbDefault
Exit Sub
ErrorHandler:
    oErrorLog.logFNError errWarning, "Error opening library.", "Error opening library."
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmbQueues_Click()
    Dim iIndex As Integer
    
    Dim oQueue As IDMObjects.queue                              ' queue we are working on
    Dim oPropDescs As IDMObjects.PropertyDescriptions           ' the object containing the property descriptions
    Dim lMaxRet As Long                                         ' the maximum number of queue entries to return
    Dim oBrowseSet As IDMObjects.QueueBrowseSet                 ' the browse set
    Dim bDocIdPresent As Boolean                                ' was a document id found in the queue
    Dim lCount As Long                                          ' Approximate count of queue entries
    
    On Error GoTo ErrorHandler
    
    oTraceLog.traceFunctionEntry "QMaint.cmbQueues_Click", SRC_FILE_NAME
    
    Me.MousePointer = vbHourglass
    
    ' Make sure we have a valid selection
    If cmbQueues.ListIndex > -1 Then
        iIndex = cmbQueues.ItemData(cmbQueues.ListIndex)
        Set oPropDescs = gQueue(iIndex).PropertyDescriptions
    
        ' set the Queue Info Frame information
        QueueInfoFrame.Caption = "Queue Search Results For: " & gQueue(iIndex).PathName
    
        BusyButton.Enabled = False
        UnbusyButton.Enabled = False
        InsertButton.Enabled = False
        EditButton.Enabled = False
        DeleteButton.Enabled = False
        QueryButton.Enabled = False
        ViewButton.Enabled = False
    
        oLocalCache.ClearPrefetchCandidates (vSession)
    
        ' Initialize all the stuff dealing with queue contents
        InitializeContent
    
        ' Initialize the grid headings and DocID locator variables
        Call InitializeGrid(oPropDescs, iDocIndex, bDocIdPresent)
        Set oPropDescs = Nothing

        ' Query the queue and display the contents
        ' Limit the #rows searched in queue...
        lMaxRet = CLng(maxRetText.Text)
        If lMaxRet = 0 Then
            lMaxRet = 50
        End If
    
        ' Build the browse set
        Call gBuildBrowseSet(oBrowseSet, gQueue(iIndex), gWSQQuerySpec(iIndex), _
                             lMaxRet, lCount)
    
        Call DisplayQContents(oBrowseSet, lMaxRet, bDocIdPresent, iDocIndex, lCount)
        Set oBrowseSet = Nothing
    
    
        MainStatusBar.SimpleText = "Retrieval Complete.  " & lMaxRet & " rows retrieved."
        MainStatusBar.Refresh

        grdQueueData.LeftCol = 0
        If lMaxRet > 0 Then
            grdQueueData.TopRow = 1
        Else
            grdQueueData.TopRow = 0
        End If
        grdQueueData.Redraw = True
        
    End If
        
    Me.MousePointer = vbDefault
    oTraceLog.traceFunctionExit "QMaint.cmbQueues_Click", SRC_FILE_NAME
    
    Exit Sub

ErrorHandler:
    
    oErrorLog.logFNError errWarning, "Error getting Queue data", "Unable to retrieve queue data"
    
    Me.MousePointer = vbDefault
    oTraceLog.traceFunctionExit "QMaint.cmbQueues_Click", SRC_FILE_NAME
    grdQueueData.Redraw = True


End Sub

'Handle the user's workspace selection
Private Sub cmbWorkspace_Click()
On Error GoTo ErrorHandler

' Reset the Workspace Query Results combo box as those entries will lock the queues.
Dim iIndex As Integer
If QMaint.cmbQueues.listCount > 0 Then
    For iIndex = QMaint.cmbQueues.listCount - 1 To 0 Step -1
        QMaint.cmbQueues.RemoveItem (iIndex)
        Set gWSQQuerySpec(iIndex) = Nothing
        Set gQueue(iIndex) = Nothing
    Next iIndex

    QMaint.cmbQueues.Enabled = False
End If

Dim sWSName As String               ' the name of the workspace containing the queue

sWSName = cmbWorkspace.Text
If sWSName <> sPrevWSName Then
    sPrevWSName = sWSName
    Screen.MousePointer = vbHourglass
    Me.cmbQueue.Enabled = False
    Me.mnuWSDelete.Enabled = False
    Me.mnuWSModify.Enabled = False
    Me.mnuWSQuery.Enabled = False
    Me.mnuWSCopy.Enabled = False
    Me.mnuWSRename.Enabled = False
    Me.mnuQueueNew.Enabled = False
    Me.mnuQueueDelete.Enabled = False
    Me.mnuQueuePrint.Enabled = False
    Me.mnuQueueModify.Enabled = False
    Me.mnuQueueQuery.Enabled = False
    Me.mnuQueueCopy.Enabled = False
    Me.mnuQueueRename.Enabled = False
    
    Set oWorkspace = Nothing
    Set oQueue = Nothing
    Set oWorkspace = oLibrary.GetObject(idmObjTypeQueueWorkspace, cmbWorkspace.Text)
    If Not (oWorkspace Is Nothing) Then
        Dim queueList As IDMObjects.ObjectSet
        Dim queue As IDMObjects.queue
        Set queueList = oWorkspace.FilterQueues("")
        Me.cmbQueue.Clear
        For Each queue In queueList
            Me.cmbQueue.AddItem queue.Name
        Next
        Set queueList = Nothing
        Me.cmbQueue.Enabled = True
        Me.mnuWSDelete.Enabled = True
        Me.mnuWSModify.Enabled = True
        Me.mnuWSQuery.Enabled = True
        Me.mnuWSCopy.Enabled = True
        Me.mnuWSRename.Enabled = True
        Me.mnuQueueNew.Enabled = True
    End If
    Screen.MousePointer = vbDefault
End If

Exit Sub
ErrorHandler:
    oErrorLog.logFNError errWarning, "Error opening workspace.", "Error opening workspace."
    Screen.MousePointer = vbDefault
End Sub
'Handle the user's queue selection
Public Sub cmbQueue_Click()

On Error GoTo ErrorHandler

' Reset the Workspace Query Results combo box as those entries will lock the queues.
Dim iIndex As Integer
If QMaint.cmbQueues.listCount > 0 Then
    For iIndex = QMaint.cmbQueues.listCount - 1 To 0 Step -1
        QMaint.cmbQueues.RemoveItem (iIndex)
        Set gWSQQuerySpec(iIndex) = Nothing
        Set gQueue(iIndex) = Nothing
    Next iIndex

    QMaint.cmbQueues.Enabled = False
End If

If cmbQueue.ListIndex > -1 Then
    MainStatusBar.SimpleText = "Opening the Queue..."
    MainStatusBar.Refresh
    Screen.MousePointer = vbHourglass
    Me.InsertButton.Enabled = False
    Me.QueryButton.Enabled = False
    Me.mnuQueueDelete.Enabled = False
    Me.mnuQueuePrint.Enabled = False
    Me.mnuQueueModify.Enabled = False
    Me.mnuQueueQuery.Enabled = False
    Me.mnuQueueCopy.Enabled = False
    Me.mnuQueueRename.Enabled = False

    Set oQueue = Nothing
    Set oQueueQuerySpec = Nothing ' clear any previous query specification
    Set oQueue = oWorkspace.GetQueue(cmbQueue.Text)
    ' set the Queue Info Frame information
    QueueInfoFrame.Caption = "Queue Entries For: " & oQueue.PathName
    QueueInfoLabel.Caption = oQueue.Description
    Set oQueueQuerySpec = oQueue.CreateQuerySpecification()
     
    ' Set default options to get all enteries
    ' This sets up the query specification to retrieve everything, which is
    ' normal for maintenance operations but not correct for most production
    ' queue usages.  In production applications setting Filters and CacheSize
    ' would be normal, but not the remainder of the following properties.
    oQueueQuerySpec.CheckUser = False
    oQueueQuerySpec.EvenIfDelayed = True
    oQueueQuerySpec.Incomplete = idmIncompleteOK
    oQueueQuerySpec.MinPriority = 0
    oQueueQuerySpec.Status = idmBusyOK
    
    If Not (oQueue Is Nothing) Then
        Me.mnuQueueDelete.Enabled = True
        Me.mnuQueuePrint.Enabled = True
        Me.mnuQueueModify.Enabled = True
        Me.mnuQueueQuery.Enabled = True
        Me.mnuQueueCopy.Enabled = True
        Me.mnuQueueRename.Enabled = True
    End If
    Screen.MousePointer = vbDefault
    Call RefreshGrid
End If
Exit Sub
ErrorHandler:
    oErrorLog.logFNError errWarning, "Error opening queue.", "Error opening queue."
    Screen.MousePointer = vbDefault
End Sub
' Delete the selected queue entry
Private Sub DeleteButton_Click()
    Dim iAnswer As Integer
    Dim oQueueEntry As IDMObjects.QueueEntry
    
    ' Always confirm a delete
    iAnswer = MsgBox("Are you sure you want to delete the entry?", vbYesNo + vbQuestion, AppName)
    If iAnswer = vbYes Then
        Set oQueueEntry = cQueueEntries(iSelectedRow)
        If oQueueEntry.EntryId = "" Or IsNull(oQueueEntry.EntryId) Then
            MsgBox "Unable to delete this Queue Entry.  Re-retrieve the queue and try again.", vbExclamation, AppName
        Else
            If oQueueEntry.Properties("F_Busy").Value Then
                If Not FetchQueueEntry(oQueueEntry, iSelectedRow, True) Then
                    MsgBox "Entry could not be fetched to delete...", vbExclamation, AppName
                    Exit Sub
                End If
            Else
                oQueueEntry.MakeReadWrite
            End If
            oQueueEntry.Delete
            cQueueEntries.Remove (iSelectedRow)
            If grdQueueData.Rows > 2 Then
                grdQueueData.RemoveItem iSelectedRow
                
            Else
                grdQueueData.Rows = 1
                UnbusyButton.Enabled = False
                DeleteButton.Enabled = False
                EditButton.Enabled = False
                QueryButton.Enabled = False
            End If
            ' Bug in MSFlexGrid control; row above is
            ' selected, but change event not triggered
            iSelectedRow = grdQueueData.RowSel
            
        End If
    End If
    
End Sub
' Bring up the edit form to modify a queue entry
Private Sub EditButton_Click()

    ' set the indexnumber to the number of the selected row in the gridcontrol
    ' and open the QEditEntry
    
    MainStatusBar.SimpleText = "Opening Edit form..."
    MainStatusBar.Refresh
    QEditEntry.iIndexNum = iSelectedRow
    QEditEntry.Show vbModal, Me
    
End Sub
' Busy out the selected queue item by doing a fetch on it
Private Sub BusyButton_Click()
Dim oQueueEntry As IDMObjects.QueueEntry
If FetchQueueEntry(oQueueEntry, iSelectedRow, False) Then
    MsgBox "Entry now marked as busy...", vbInformation, AppName
Else
    MsgBox "Could not busy this entry...", vbExclamation, AppName
End If
' This could be optimized - but we need to rebuild the queue entry
' in the cQueueEntries collection and update the grid display
Call RefreshGrid
End Sub

Private Sub mnuAbout_Click()
    AboutBox.Show vbModal, Me
End Sub

Private Sub mnuQueuePrint_Click()

    Dim OldOrientation As Integer ' Original orientation of the printer
    Dim tppx As Integer  ' alias TwipsPerPixelX
    Dim tppy As Integer  ' alias TwipsPerPixelY
    tppx = Printer.TwipsPerPixelX
    tppy = Printer.TwipsPerPixelY
    Dim Col As Integer   ' index to grid columns
    Dim Row As Integer   ' index to grid rows
    Dim x0 As Single     ' upper left corner
    Dim y0 As Single     '   "
    Dim x1 As Single     ' position of text
    Dim y1  As Single    '   "
    Dim x2  As Single    ' position of grid lines
    Dim y2  As Single    '   "
    Dim CurrentRow As Integer
    Dim MoreData As Boolean
      
    On Error GoTo ErrorHandler
    
    ' Get setup to print
    Screen.MousePointer = vbHourglass
    QMaint.MainStatusBar.SimpleText = "Printing Queue Contents for: " & oQueue.Name & "..."
    QMaint.MainStatusBar.Refresh

    ' Save and setup page orientation
    OldOrientation = Printer.Orientation
    Printer.Orientation = vbPRORLandscape
    
    ' Print the header
    Printer.Print "Library = " + oLibrary.Name + "  /  Workspace = " + oWorkspace.Name + "  /  Queue = " + oQueue.Name
    
    ' Add a little space after the header
    Printer.CurrentY = Printer.CurrentY + 50

    CurrentRow = 0
    MoreData = True
    ' Loop for each page of data to print
    Do While MoreData
    
       ' Set upper left corner
       x0 = Printer.CurrentX
       y0 = Printer.CurrentY
    
       ' Draw the text in the grid
       x1 = x0
       For Col = 0 To grdQueueData.Cols - 1
          ' Skip non-visible columns
          If Col >= grdQueueData.FixedCols And Col < grdQueueData.LeftCol Then
             Col = grdQueueData.LeftCol
          End If
          ' Stop if outside grid
          If x1 + grdQueueData.ColWidth(Col) >= Printer.Width Then Exit For
          y1 = y0
          For Row = CurrentRow To grdQueueData.Rows - 1
             ' Stop if outside grid
             If y1 + grdQueueData.RowHeight(Row) >= Printer.ScaleHeight Then
                ' Set flag to indicate there is another page of data to print
                MoreData = True
                Exit For
             End If
             MoreData = False
             ' Set position to print the cell
             Printer.CurrentX = x1 + tppx * 2 + 20
             Printer.CurrentY = y1 + tppy + 50
             ' Print cell text
             grdQueueData.Col = Col
             grdQueueData.Row = Row
             Printer.Print grdQueueData.Text
             
             ' Advance to next row
             y1 = y1 + grdQueueData.RowHeight(Row)
             If grdQueueData.GridLines Then
                y1 = y1 + tppy
             End If
          Next
          ' Advance to next column
          x1 = x1 + grdQueueData.ColWidth(Col)
          If grdQueueData.GridLines Then
             x1 = x1 + tppx
          End If
       Next

       ' Draw grid lines
       If grdQueueData.GridLines Then
          x2 = x0
          y2 = y0
          For Col = 0 To grdQueueData.Cols - 1
             ' Skip non-visible columns
             If Col >= grdQueueData.FixedCols And Col < grdQueueData.LeftCol Then
                Col = grdQueueData.LeftCol
             End If
             ' Stop if outside grid
             If x2 >= Printer.Width Then Exit For
             
             ' y0 represents the twips for the header
             Printer.Line (x2, y0)-Step(0, y1 - tppy - y0)
             x2 = x2 + grdQueueData.ColWidth(Col)
             x2 = x2 + tppx
             If Col = grdQueueData.Cols - 1 Then
                 Printer.Line (x2, y0)-Step(0, y1 - tppy - y0)
             End If
          Next
          
          For Row = CurrentRow To grdQueueData.Rows - 1
             ' Stop if outside grid
             If y2 >= Printer.ScaleHeight Then Exit For
             
             Printer.Line (x0, y2)-Step(x1 - tppx, 0)
             y2 = y2 + grdQueueData.RowHeight(Row)
             y2 = y2 + tppy
             If Row = grdQueueData.Rows - 1 Then
                 Printer.Line (x0, y2)-Step(x1 - tppx, 0)
             End If
          Next
          
       End If
    
       ' Back up one
       CurrentRow = Row - 1
      
       If MoreData Then
           Printer.NewPage
       End If
    Loop

    ' Force the print
    Printer.EndDoc
    
    ' Restore the original orientation
    Printer.Orientation = OldOrientation
    
    ' Reset the grid
    If grdQueueData.Rows = 1 Then
        grdQueueData.Row = 0
    Else
        grdQueueData.Row = 1
    End If
    grdQueueData.Col = 0
    
    ' Restore mouse pointer
    Screen.MousePointer = vbDefault
    QMaint.MainStatusBar.SimpleText = "Completed Printing Queue Contents for: " & oQueue.Name & "..."
    QMaint.MainStatusBar.Refresh

    Exit Sub
    
ErrorHandler:
    oErrorLog.logFNError errWarning, "Error printing queue contents.", "Error printing queue contents."
    Screen.MousePointer = vbDefault

End Sub
Private Sub mnuQueueCopy_Click()
    QCopy.Show vbModal, Me
End Sub

Private Sub mnuQueueDelete_Click()
    DelConf.Show vbModal, Me
End Sub

Private Sub mnuQueueModify_Click()
    QCreate.isNew = False
    QCreate.Show vbModal, Me
    If Not (oQueue Is Nothing) Then
        QueueInfoLabel.Caption = oQueue.Description
    End If
End Sub

Private Sub mnuQueueNew_Click()
    QCreate.isNew = True
    QCreate.Show vbModal, Me
End Sub

Private Sub mnuQueueQuery_Click()
    QQuery.Show vbModal, Me
End Sub

Private Sub mnuQueueRename_Click()
    QRename.Show vbModal, Me
End Sub

Private Sub mnuWSDelete_Click()
    Dim listCount As Integer
    Dim inx As Integer
    Dim res As VbMsgBoxResult
    
    On Error GoTo ErrorHandler
    res = MsgBox("Are you sure you wish to delete workspace " & oWorkspace.Name & "?", _
                 vbOKCancel + vbQuestion, AppName)
    If res <> vbOK Then
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    MainStatusBar.SimpleText = "Deleting workspace..."
    MainStatusBar.Refresh
    oWorkspace.Delete
    listCount = cmbWorkspace.listCount
    For inx = 0 To listCount - 1
        If cmbWorkspace.List(inx) = oWorkspace.Name Then
            cmbWorkspace.RemoveItem (inx)
            Exit For
        End If
    Next inx
    MainStatusBar.SimpleText = "Workspace deleted."
    MainStatusBar.Refresh
    mnuWSDelete.Enabled = False
    mnuWSModify.Enabled = False
    mnuWSQuery.Enabled = False
    mnuWSCopy.Enabled = False
    mnuWSRename.Enabled = False
    mnuQueueNew.Enabled = False
    Set oWorkspace = Nothing
    sPrevWSName = ""
    cmbQueue.Enabled = False
    InsertButton.Enabled = False
    QueryButton.Enabled = False
    mnuQueueDelete.Enabled = False
    mnuQueuePrint.Enabled = False
    mnuQueueModify.Enabled = False
    mnuQueueQuery.Enabled = False
    mnuQueueCopy.Enabled = False
    mnuQueueRename.Enabled = False

    Set oQueue = Nothing
    Screen.MousePointer = vbDefault
    Exit Sub
ErrorHandler:
    oErrorLog.logFNError errWarning, "Error deleting workspace.", "Error deleting workspace."
    Screen.MousePointer = vbDefault
End Sub
Private Sub mnuWSModify_Click()
    WSCreate.isNew = False
    WSCreate.Show vbModal, Me
End Sub

Private Sub mnuWSCopy_Click()
    WSCopy.Show vbModal, Me
End Sub

Private Sub mnuWSNew_Click()
    WSCreate.isNew = True
    WSCreate.Show vbModal, Me
End Sub

Private Sub mnuWSQuery_Click()
    WSQuery.Show vbModal, Me
End Sub

Private Sub mnuWSRename_Click()
    WSRename.Show vbModal, Me
End Sub

Private Sub QueryButton_Click()

    MainStatusBar.SimpleText = "Opening Query form..."
    MainStatusBar.Refresh
    QQuery.Show vbModal, Me
    
End Sub

Private Sub SystemCheck_Click()

    Call cmbQueue_Click
    
End Sub

Private Sub UnbusyButton_Click()
Dim oQueueEntry As IDMObjects.QueueEntry

' Always confirm an override of the busy status
If MsgBox("Are you sure you want to make this item not busy?", vbYesNo + vbQuestion, AppName) _
              = vbNo Then
    Exit Sub
End If
If FetchQueueEntry(oQueueEntry, iSelectedRow, True) Then
    oQueueEntry.Properties("F_Busy").Value = Empty
    On Error GoTo Problems
    oQueueEntry.Update (False)
    MsgBox "Entry now marked as not busy...", vbInformation, AppName
Else
    MsgBox "Entry could not be fetched to make not busy...", vbExclamation, AppName
End If
' This could be optimized - but we need to rebuild the queue entry
' in the cQueueEntries collection and update the grid display
Call RefreshGrid
Exit Sub
Problems:
MsgBox "Queue update failed: " & Err.Description, vbExclamation, AppName
End Sub
' Insert a new entry in the queue
Private Sub InsertButton_Click()

    ' Set the iIndexNum to zero to indicate an insert and open the QEditEntry
    
    MainStatusBar.SimpleText = "Opening Insert form..."
    MainStatusBar.Refresh

    QEditEntry.iIndexNum = 0
    QEditEntry.Show vbModal, Me
    ' This could be optimized - but we need to rebuild the queue entry
    ' in the cQueueEntries collection and update the grid display
    Call RefreshGrid

End Sub
' Bring up the Viewer on the Doc referenced in the queue entry
Private Sub ViewButton_Click()
    Dim oDocument As IDMObjects.Document
    Dim sDocId As String
    Dim iRow As Integer
          
    ' create a document object with the sDocId and launch it
    
    On Error GoTo ErrorHandler
    
    MainStatusBar.SimpleText = "Opening Image Viewer..."
    MainStatusBar.Refresh

    iRow = grdQueueData.RowSel
    sDocId = grdQueueData.TextMatrix(iRow, iDocIndex)
    Set oDocument = oLibrary.GetObject(idmObjTypeDocument, CDbl(sDocId))
    oDocument.Launch idmDocLaunchIDMViewer
    Set oDocument = Nothing

    MainStatusBar.SimpleText = ""
    MainStatusBar.Refresh

    Exit Sub
    
ErrorHandler:
    
    oErrorLog.logFNError errWarning, "Error opening document (id=" & sDocId & ")", "Unable to open the document."
    
    Me.MousePointer = vbDefault
    
End Sub

' Function to fetch and set busy the queue entry matching the pass row # in grid
Public Function FetchQueueEntry(oFinalQueueEntry As IDMObjects.QueueEntry, _
    ByVal iSelectedRow As Integer, bEvenIfBusy As Boolean) As Boolean
Dim oQueueEntry As IDMObjects.QueueEntry
Dim sEntryId As String

On Error GoTo Problems

Set oQueueEntry = cQueueEntries.Item(iSelectedRow)

sEntryId = oQueueEntry.EntryId
            
' get the entry
' EvenIfBusy and EvenIfIncomplete would normally be False in production applications
Set oFinalQueueEntry = oQueue.GetEntryFromId(sEntryId, bEvenIfBusy, True, True)
FetchQueueEntry = True
Exit Function

Problems:
FetchQueueEntry = False
End Function
Private Sub Form_Load()
    Dim oLib As IDMObjects.Library
    
    ' initialize some things
    
    oTraceLog.initialize        ' Initialize the trace log object
    oTraceLog.traceModuleLoad   ' Trace the load of the module
    oErrorLog.initialize        ' Initialize the error log object
    
    oLocalCache.PrefetchThreadPriority = idmThreadPriorityIdle
    vSession = oLocalCache.CreatePrefetchSession    ' initialize the session for the local cache
    
    Set oErrManager = CreateObject("IDMError.ErrorManager")
    
    ' Populate the libraries combo box and let the user
    ' make a choice
    For Each oLib In oHood.Libraries
        If oLib.SystemType = idmSysTypeIS Then
            cmbLibraries.AddItem oLib.Label
        End If
    Next
    ' Disable the other controls until user has
    ' logged on to a library
    Me.cmbWorkspace.Enabled = False
    Me.cmbQueue.Enabled = False
    Me.InsertButton.Enabled = False
    Me.QueryButton.Enabled = False
    Me.cmbQueues.Enabled = False
    iFormHeight = Me.Height
    iFormWidth = Me.Width
    iMinFormHeight = Me.Height / 2
'    iMinFormWidth = 12000

    sPrevWSName = ""
    
End Sub
' Handle form resizing event
Private Sub Form_Resize()

    Dim iHeightChange As Integer
    Dim iWidthChange As Integer
    Dim oTmpCtrl As Control
    
    ' reposition the objects when the form is resized
    oTraceLog.traceFunctionEntry "QMaint.Form_Resize", SRC_FILE_NAME
    
    If Me.WindowState = vbMinimized Then
        Exit Sub
    End If
    
    If Me.Height < iMinFormHeight Then
        Me.Height = iMinFormHeight
    End If
    
    If Me.Width < iMinFormWidth Then
        Me.Width = iMinFormWidth
    End If

    iHeightChange = Me.Height - iFormHeight
    iWidthChange = Me.Width - iFormWidth
    
    ' Move the buttons below the grid
    For Each oTmpCtrl In Me.Controls
        If TypeOf oTmpCtrl Is CommandButton And _
         oTmpCtrl.Name <> "frQueryResult" Then
            oTmpCtrl.Top = oTmpCtrl.Top + iHeightChange
            oTmpCtrl.Left = oTmpCtrl.Left + iWidthChange
        End If
    Next
    
    grdQueueData.Height = grdQueueData.Height + iHeightChange
    grdQueueData.Width = grdQueueData.Width + iWidthChange
    
    QueueInfoFrame.Height = QueueInfoFrame.Height + iHeightChange
    QueueInfoFrame.Width = QueueInfoFrame.Width + iWidthChange
    
    QueueFrame.Width = QueueFrame.Width + iWidthChange
    frQueryResult.Left = frQueryResult.Left + iWidthChange
    SystemCheck.Left = SystemCheck.Left + iWidthChange
    maxRetText.Left = maxRetText.Left + iWidthChange
    maxRetLabel.Left = maxRetLabel.Left + iWidthChange
    
    iFormHeight = Me.Height
    iFormWidth = Me.Width
    oTraceLog.traceFunctionExit "QMaint.Form_Resize", SRC_FILE_NAME

End Sub
' Cleanup on termination
Private Sub Form_Unload(Cancel As Integer)

    oLocalCache.ClearPrefetchCandidates (vSession)
    oTraceLog.traceModuleUnload
    
    Set oLibrary = Nothing
    Set oWorkspace = Nothing
    Set oQueue = Nothing
    Set oLocalCache = Nothing
    Set oErrorLog = Nothing
    Set oTraceLog = Nothing
    
    Set oErrManager = Nothing
    
End Sub
' Internal routine for initializing the contents grid and
' the collection of queue entry data
Private Sub InitializeContent()
Dim iCounter As Integer
' Clear the grid control
grdQueueData.Rows = 0
grdQueueData.Redraw = False
' Clear our local collection of queue data
For iCounter = 1 To cQueueEntries.Count
    cQueueEntries.Remove (1)
Next
        
End Sub
' Initialize the grid control with the right property headings;
' See if there is a DocId field in this queue schema
Private Sub InitializeGrid(oPropDescs As IDMObjects.PropertyDescriptions, _
    ByRef iDocIndex As Integer, bDocIdPresent As Boolean)
Dim iColumns As Integer
Dim iCounter As Integer
Dim sRowData As String
Dim oPropDesc As IDMObjects.PropertyDescription

MainStatusBar.SimpleText = "Create the headings for the grid..."
MainStatusBar.Refresh
bDocIdPresent = False
iDocIndex = -1   ' Column location for DocId field

' if the user has asked for system fields, show the Entry ID field
If SystemCheck.Value = 1 Then
    sRowData = "Entry_id"
    iColumns = 1
Else
    sRowData = ""
    iColumns = 0
End If
        
' Create headings for each field to display
    
For Each oPropDesc In oPropDescs
    ' See if this property is a DocId field
    If oPropDesc.GetExtendedProperty("F_QUEUETYPEID") = idmQueueTypeDocument And _
            oPropDesc.GetState(idmPropCustom) And Not bDocIdPresent Then
        iDocIndex = iColumns
        bDocIdPresent = True
    End If
    If oPropDesc.GetState(idmPropCustom) Or SystemCheck.Value = 1 Then
        If iColumns > 0 Then
            sRowData = sRowData & vbTab
        End If
        sRowData = sRowData & oPropDesc.Name
        iColumns = iColumns + 1
    End If
Next oPropDesc
    
grdQueueData.Cols = iColumns
grdQueueData.AddItem sRowData
grdQueueData.Rows = 1

' Set up field alignments
If SystemCheck.Value = 1 Then
    ' Align EntryId column
    grdQueueData.ColAlignment(0) = 1    ' flexAlignLeftCenter
    grdQueueData.FixedAlignment(0) = 4  ' flexAlignCenterCenter
    grdQueueData.ColWidth(0) = -1
    iCounter = 1
Else
    iCounter = 0
End If
For Each oPropDesc In oPropDescs
    If oPropDesc.GetState(idmPropCustom) Or SystemCheck.Value = 1 Then
        Select Case oPropDesc.TypeID
        Case idmTypeDouble, idmTypeLong, idmTypeShort, idmTypeObject
            grdQueueData.ColAlignment(iCounter) = flexAlignRightCenter
        Case Else
            grdQueueData.ColAlignment(iCounter) = flexAlignLeftCenter
        End Select
        grdQueueData.FixedAlignment(iCounter) = flexAlignCenterCenter
        grdQueueData.ColWidth(iCounter) = -1
        iCounter = iCounter + 1
    End If
Next oPropDesc

Set oPropDesc = Nothing

End Sub
' Extract data from the browse set and populate the grid control
' As we encounter DocIds, tell the ServerCache about them...
Private Sub DisplayQContents(oBrowseSet As IDMObjects.QueueBrowseSet, _
    ByRef lMaxRet As Long, ByVal bDocIdPresent As Boolean, _
    ByVal iDocIndex As Integer, lCount As Long)
Dim oProperties As IDMObjects.Properties                    ' the properties object
Dim iPropCount As Integer   ' Counter for columns displayed
Dim iCounter As Integer     ' another counter
Dim sRowData As String
Dim sValue As String        ' string representation of value
Dim iDispWidth As Integer   ' display width of sValue
Dim sDocId As String
Dim oProperty As IDMObjects.Property

If oBrowseSet Is Nothing Then
    lMaxRet = 0
End If
If lMaxRet < lCount Then
    lCount = lMaxRet
End If
' Don't use lCount as an upper bound for the number of records to get,
' since it is only an approximation (i.e. do not use in For ... Next).
' Also don't use RecordCount because it interferes with the cacheing
' mechanism in a QueueBrowseSet and copies all queue entries to the local
' workstation.
iCounter = 1
Do While iCounter <= lMaxRet
    Set oProperties = oBrowseSet.Entry.Properties
    
    ' If user wants "system values", also give him the entry id
    If SystemCheck.Value = 1 Then
        sRowData = oBrowseSet.Entry.EntryId
        ' Auto size column width
        iDispWidth = Len(oBrowseSet.Entry.EntryId) * grdQueueData.CellFontSize * 12
        If iDispWidth > grdQueueData.ColWidth(0) Then
            grdQueueData.ColWidth(0) = iDispWidth
        End If
        iPropCount = 1
    Else
        sRowData = ""
        iPropCount = 0
    End If
    
    For Each oProperty In oProperties
        If oProperty.PropertyDescription.GetState(idmPropCustom) Or _
                SystemCheck.Value = 1 Then
            If iPropCount > 0 Then
                sRowData = sRowData & vbTab
            End If
            If IsEmpty(oProperty.Value) Or IsNull(oProperty.Value) Or _
                    Not oProperty.PropertyDescription.GetState(idmPropSupportsFPNumber) Then
                ' Clear the default value
                If oProperty.Name = "F_Delay" And oProperty.Value = F_DELAY_DEFAULT Then
                    oProperty.Value = Null
                End If
                sValue = CStr(IIf(IsNull(oProperty.Value), "", oProperty.Value))
            Else
                ' Get string from FnFPNumber, so no loss of precision happens
                ' in a conversion to and from a double
                sValue = oProperty.FnFPNumber.ValueAsString
            End If
            sRowData = sRowData & sValue
            ' Auto size column width
            iDispWidth = Len(sValue) * grdQueueData.CellFontSize * 12
            If iDispWidth > grdQueueData.ColWidth(iPropCount) Then
                grdQueueData.ColWidth(iPropCount) = iDispWidth
            End If
            
            ' if this is the doc id, cache the doc on the server
            
            If bDocIdPresent And (iPropCount = iDocIndex) Then
                If Not IsNull(oProperty.Value) Then
                    sDocId = oProperty.Value    ' Type coercion
                    ServerCacheDocument (sDocId)
                End If
            End If
            iPropCount = iPropCount + 1
        End If
            
    Next oProperty
    grdQueueData.AddItem sRowData
    cQueueEntries.Add oBrowseSet.Entry
    If oBrowseSet.MoreResults Then
        oBrowseSet.MoveNext
    Else
        lMaxRet = iCounter
    End If
    
    MainStatusBar.SimpleText = "Retrieving Queue Entry " & iCounter & " of " & lCount & "..."
    MainStatusBar.Refresh
    DoEvents
    
    iCounter = iCounter + 1
    
Loop
' Tweak the grid control some more
' Make the first row the column headings
If lMaxRet > 0 Then
    grdQueueData.FixedRows = 1
End If
' Move the focus to the first row
If lMaxRet > 0 Then
    grdQueueData.Row = 1
Else
    grdQueueData.Row = 0
End If

Set oProperties = Nothing
Set oProperty = Nothing

End Sub
' Handle the button click to show queue contents
Public Sub RefreshGrid()
    
    Dim oPropDescs As IDMObjects.PropertyDescriptions           ' the object containing the property descriptions
    Dim lMaxRet As Long                                         ' the maximum number of queue entries to return
    Dim oBrowseSet As IDMObjects.QueueBrowseSet                 ' the browse set
    Dim bDocIdPresent As Boolean                                ' was a document id found in the queue
    Dim lCount As Long                                          ' Approximate count of queue entries
    
    On Error GoTo ErrorHandler
    
    oTraceLog.traceFunctionEntry "QMaint.RefreshGrid", SRC_FILE_NAME
    
    Me.MousePointer = vbHourglass
    
    Set oPropDescs = oQueue.PropertyDescriptions
    
    ViewButton.Enabled = False
    EditButton.Enabled = False
    DeleteButton.Enabled = False
    BusyButton.Enabled = False
    UnbusyButton.Enabled = False
    QueryButton.Enabled = False
    
    oLocalCache.ClearPrefetchCandidates (vSession)
    
    ' Initialize all the stuff dealing with queue contents
    InitializeContent
    ' Initialize the grid headings and DocID locator variables
    Call InitializeGrid(oPropDescs, iDocIndex, bDocIdPresent)
    Set oPropDescs = Nothing

    ' Query the queue and display the contents
    ' Limit the #rows searched in queue...
    lMaxRet = CLng(maxRetText.Text)
    If lMaxRet = 0 Then
        lMaxRet = 50
    End If
    InsertButton.Enabled = True
    QueryButton.Enabled = True
    Call gBuildBrowseSet(oBrowseSet, oQueue, oQueueQuerySpec, lMaxRet, lCount)
    Call DisplayQContents(oBrowseSet, lMaxRet, bDocIdPresent, _
        iDocIndex, lCount)
    Set oBrowseSet = Nothing
    
    
    MainStatusBar.SimpleText = "Retrieval Complete.  " & lMaxRet & " rows retrieved."
    MainStatusBar.Refresh

    grdQueueData.LeftCol = 0
    If lMaxRet > 0 Then
        grdQueueData.TopRow = 1
    Else
        grdQueueData.TopRow = 0
    End If
    grdQueueData.Redraw = True
    
    Me.MousePointer = vbDefault
    
    oTraceLog.traceFunctionExit "QMaint.RefreshGrid", SRC_FILE_NAME
    
    Exit Sub

ErrorHandler:
    
    oErrorLog.logFNError errWarning, "Error getting Queue data", "Unable to retrieve queue data"
    
    Me.MousePointer = vbDefault
    oTraceLog.traceFunctionExit "QMaint.RefreshGrid", SRC_FILE_NAME
    grdQueueData.Redraw = True

End Sub
' Handle selection changes in the grid
Private Sub grdQueueData_Click()
    Dim lRow As Long
    Dim sDocId As String
    
    On Error GoTo ErrorHandler
    
    oTraceLog.traceFunctionEntry "QMaint.grdQueueData_selChange", SRC_FILE_NAME
    lRow = grdQueueData.RowSel
    
    If lRow > 0 Then                ' if a queue entry is selected, allow the user to edit the entry
        iSelectedRow = lRow
        EditButton.Enabled = True
        DeleteButton.Enabled = True
        BusyButton.Enabled = True
        QueryButton.Enabled = True
    Else                            ' otherwise, turn off the EditButton
        iSelectedRow = 0
        EditButton.Enabled = False
        DeleteButton.Enabled = False
        BusyButton.Enabled = False
        QueryButton.Enabled = False
    End If
    
    If lRow > 0 Then
        MainStatusBar.SimpleText = "Queue Entry (ID=" & cQueueEntries(lRow).EntryId & ") selected."
        MainStatusBar.Refresh
        If cQueueEntries.Item(lRow).Properties("F_Busy").Value Then
            UnbusyButton.Enabled = True
            BusyButton.Enabled = False
        Else
            UnbusyButton.Enabled = False
            BusyButton.Enabled = True
        End If
        
    Else
        MainStatusBar.SimpleText = "No Queue Entry selected."
        MainStatusBar.Refresh
    End If
    If iDocIndex < 0 Or lRow <= 0 Then           'if there is no document id field in the queue
        ViewButton.Enabled = False
    Else                            'if there is a document id field, is it filled in?
        sDocId = grdQueueData.TextMatrix(lRow, iDocIndex)
        If (sDocId <> "") Then
            ViewButton.Enabled = True
            Call LocalCacheDocument(sDocId)
        Else
            ViewButton.Enabled = False
        End If
    End If
    
    oTraceLog.traceFunctionExit "QMaint.grdQueueData_selChange", SRC_FILE_NAME
    Exit Sub
    
ErrorHandler:
    
    oErrorLog.logFNError errWarning, "Error selecting Queue Entry", "Error selecting Queue Entry"
    oTraceLog.traceFunctionExit "QMaint.grdQueueData_selChange", SRC_FILE_NAME
    
End Sub
' Add a document to the local (client) cache
Private Function LocalCacheDocument(ByVal sDocId As String)
    Dim oDocument As IDMObjects.Document
    
    On Error GoTo ErrorHandler
 
    If sDocId <> "" Then
        Set oDocument = oLibrary.GetObject(idmObjTypeDocument, CDbl(sDocId))
        oLocalCache.AddPrefetchCandidate vSession, oDocument, 1, oDocument.PageCount
        Set oDocument = Nothing
    End If
    
    Exit Function
    
ErrorHandler:
    
    ' If the Error is -2147220985 ("Document Not Found") or -2147215871
    ' ("Invalid document identifier"), or -2147215870 ("Document doesn't exist"),
    ' do not report it now.
    ' It will be reported when they try to view the document
    
    If Err.Number <> -2147220985 And Err.Number <> -2147215871 And _
            Err.Number <> -2147215870 Then
        oErrorLog.logFNError errWarning, "Error writing document to local cache.", "Error writing document to local cache."
    End If
    
    Me.MousePointer = vbDefault
        
    
End Function
' Tell the server that we will probably want to view this document
Private Function ServerCacheDocument(ByVal sDocId As String)
    Dim oDocument As IDMObjects.Document
    
    On Error GoTo ErrorHandler
    
    If sDocId <> "" Then
        Set oDocument = oLibrary.GetObject(idmObjTypeDocument, CDbl(sDocId))
        oServerCache.Prefetch oDocument, idmPrefetchPriorityLow
        Set oDocument = Nothing
    End If
    
    Exit Function
    
ErrorHandler:
    
    ' If the Error is -2147220985 ("Document Not Found") or
    '                 -2147215871 ("Invalid document identifier") or
    '                 -2147215870 ("Document doesn't exist") or
    '                 -2147214748 ("Retrieve document content failed")
    ' Do not report it now. It will be reported when they try to view the document
    
    If Err.Number <> -2147220985 And Err.Number <> -2147215871 And _
       Err.Number <> -2147215870 And Err.Number <> -2147214748 Then
        oErrorLog.logFNError errWarning, "Error writing document to local cache.", "Error writing document to local cache."
    End If
    
    Me.MousePointer = vbDefault
        
 End Function
