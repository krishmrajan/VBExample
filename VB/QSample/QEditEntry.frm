VERSION 5.00
Begin VB.Form QEditEntry 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Edit Queue Entry"
   ClientHeight    =   6825
   ClientLeft      =   3375
   ClientTop       =   825
   ClientWidth     =   5775
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   5775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3480
      TabIndex        =   3
      Top             =   6240
      Width           =   1455
   End
   Begin VB.CommandButton SaveButton 
      Caption         =   "Save"
      Default         =   -1  'True
      Height          =   375
      Left            =   840
      TabIndex        =   2
      Top             =   6240
      Width           =   1455
   End
   Begin VB.TextBox DataField 
      Height          =   285
      Index           =   0
      Left            =   2040
      TabIndex        =   0
      Top             =   240
      Width           =   3015
   End
   Begin VB.VScrollBar sbProperties 
      Height          =   3930
      Left            =   5280
      TabIndex        =   4
      Top             =   240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label DataLabel 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   1
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "QEditEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public iIndexNum As Integer                         ' the index number of the QueueEntry to edit; passed
                                                    ' in by caller
Public oQueueEntry As IDMObjects.QueueEntry         ' the QueueEntry to edit
Public bInsert As Boolean                           ' Is this an insert? otherwise, it's an update
Public bBusy As Boolean                             ' is the QueueEntry busy?
Public cProperties As Collection                    ' the collection of properties for this QueueEntry
Public iTotalFields As Integer                      ' the total number of fields displayed on the form
Private Const visibleProps = 20                     ' the number of visible properties on form
Private Const iSpacing = 120                        ' the vertical spacing between the datafields

' date order constants
Private Const DATEORDER_MMDDYY = 0
Private Const DATEORDER_DDMMYY = 1
Private Const DATEORDER_YYMMDD = 2
' date type (long or short)
Private Const LOCALE_SHORTDATE = 33
Private Const LOCALE_LONGDATE = 34


' Cancel out of edit
Private Sub CancelButton_Click()
    If bInsert Then
        QMaint.MainStatusBar.SimpleText = "Insert cancelled at user's request."
    Else
        QMaint.MainStatusBar.SimpleText = "Update cancelled at user's request."
    End If
    
    Set oQueueEntry = Nothing
    QMaint.MainStatusBar.Refresh
    Unload Me
End Sub

Private Sub InitQueueEntry(oQueueEntry As IDMObjects.QueueEntry, _
    ByVal iIndexNum As Integer, ByVal bInsert As Boolean)

If bInsert Then   'No previous EntryID.  This is an insert
    
    ' create the new entry
    Set oQueueEntry = QMaint.oQueue.CreateEmptyEntry
    QEditEntry.Caption = "Insert Queue Entry"
    SaveButton.Caption = "Insert"
        
Else    'This is an update
        
        ' get the entry from the cQueueEntries collection
    Set oQueueEntry = QMaint.cQueueEntries.Item(iIndexNum)
    QEditEntry.Caption = "Edit Queue Entry"
    SaveButton.Caption = "Update"
        
End If

End Sub

Private Sub Form_Load()

    Dim oProperties As IDMObjects.Properties        ' a collection of properties for a QueueEntry
    Dim oProperty As IDMObjects.Property            ' a property of a QueueEntry
    Dim iCounter As Integer                         ' an integer counter
    Dim iHeight As Integer                          ' the height of the standard DataField field
    Dim sQueueName As String                        ' the name of the queue being edited
    Dim sWSName As String                           ' the name of the workspace being edited
    Dim bShowSystem As Boolean                      ' should the edit form show the system fields?
    Dim iVisibleFields As Integer                   ' the number of visible fields

On Error GoTo ErrorHandler

    bShowSystem = (QMaint.SystemCheck.Value = 1)
     
    ' Get the QueueEntry initialized
    If iIndexNum = 0 Then
        bInsert = True
    Else
        bInsert = False
    End If
    Call InitQueueEntry(oQueueEntry, iIndexNum, bInsert)
    
    ' get the properties from the Queue Entry
    Set oProperties = oQueueEntry.Properties
    
    iCounter = 0
    iHeight = DataField(0).Height
    
    ' for each property (Name/Value), add a field and label to the form.  Fill the field and label
    ' DataLabel and DataField are declared as controls arrays on the form
    bBusy = False
    For Each oProperty In oProperties
        If oProperty.PropertyDescription.GetState(idmPropCustom) Or _
                bShowSystem Then
            If iCounter >= 1 Then
                Load DataLabel(iCounter)
                Load DataField(iCounter)
                If iCounter >= visibleProps Then
                    sbProperties.Max = iCounter - visibleProps + 1
                    sbProperties.Visible = True
                End If
            End If
            DataLabel(iCounter).Caption = oProperty.Name & ":"
        
            If IsEmpty(oProperty.Value) Or IsNull(oProperty.Value) Or _
               Not oProperty.PropertyDescription.GetState(idmPropSupportsFPNumber) Then
                    DataField(iCounter).Text = CStr(IIf(IsNull(oProperty.Value), "", oProperty.Value))
            Else
                    ' Get string from FnFPNumber, so no loss of precision happens
                    ' in a conversion to and from a double.
                    DataField(iCounter).Text = oProperty.FnFPNumber.ValueAsString
            End If

            If iCounter < visibleProps Then
                DataLabel(iCounter).Top = DataLabel(0).Top + (iHeight + iSpacing) * iCounter
                DataField(iCounter).Top = DataField(0).Top + (iHeight + iSpacing) * iCounter
                Call setVisible(iCounter, True)
            Else
                Call setVisible(iCounter, False)
            End If
            iCounter = iCounter + 1
        End If
        If oProperty.Name = "F_Busy" And oProperty.Value = True Then
            bBusy = True
        End If
    Next
    
    iTotalFields = iCounter
    If iTotalFields <= visibleProps Then
        iVisibleFields = iCounter
    Else
        iVisibleFields = visibleProps
    End If
    
    ' move the savebutton and cancelbutton below the last field and resize the form to fit everything
    ' adjust scroll bar
    sbProperties.TabIndex = 2 * iCounter + 2
    CancelButton.TabIndex = 2 * iCounter + 1
    SaveButton.TabIndex = 2 * iCounter
    SaveButton.Top = DataField(0).Top + (iHeight + iSpacing) * iVisibleFields + iSpacing
    CancelButton.Top = SaveButton.Top
    sbProperties.Height = visibleProps * (iHeight + iSpacing) - iSpacing
    Me.Height = SaveButton.Top + SaveButton.Height + (5 * iSpacing)
    
Exit Sub

ErrorHandler:
    ShowError
    
End Sub

' Update the properties in the queue entry, then call the update method
Private Sub UpdateQueueEntry(oQueueEntry As IDMObjects.QueueEntry, _
    oProperties As IDMObjects.Properties, _
    ByVal iTotalFields As Integer, ByVal bInsert As Boolean)

Dim oProperty As IDMObjects.Property
Dim iCounter As Integer
Dim ConvertDate As Date
Dim ErrMsg As String
' for each Property, find its associated DataLabel/DataField pair
' cast the data to the correct type and write it to the value field
    
On Error GoTo ErrorHandler

For Each oProperty In oProperties
    For iCounter = 0 To iTotalFields - 1
        If DataLabel(iCounter).Caption = oProperty.Name & ":" Then
            If DataField(iCounter).Text = "" Then
                If oProperty.PropertyDescription.GetState(idmPropCustom) Or _
                        oProperty.TypeID = idmTypeDate Then
                    oProperty.Value = Null
                End If
            Else
                Select Case oProperty.PropertyDescription.TypeID
                Case idmTypeBoolean
                    oProperty.Value = CBool(DataField(iCounter).Text)
                Case idmTypeByte
                    oProperty.Value = CByte(DataField(iCounter).Text)
                Case idmTypeCurrency
                    oProperty.Value = CCur(DataField(iCounter).Text)
                Case idmTypeDate
                    ConvertDate = CDate(DataField(iCounter).Text)
                    If (Not VerifyDate(DataField(iCounter).Text, ConvertDate)) Then
                        ' force error here
                        ErrMsg = "Invalid input date """ + DataField(iCounter).Text + """"
                        Err.Raise vbObjectError + 513, , ErrMsg
                        'oProperty.Value = NULL
                    Else
                        oProperty.Value = ConvertDate
                    End If
                Case idmTypeObject
                    If oProperty.PropertyDescription.GetState(idmPropSupportsFPNumber) Then
                        ' Produce FnFPNumber from string, so no loss of precision
                        ' happens in a conversion to and from a double
                        If IsEmpty(oProperty.Value) Or IsNull(oProperty.Value) Then
                            Dim tempFP As New IDMObjects.FnFPNumber
                            tempFP.ValueAsString = DataField(iCounter).Text
                            oProperty.Value = tempFP
                            Set tempFP = Nothing
                        Else
                            oProperty.FnFPNumber.ValueAsString = DataField(iCounter).Text
                        End If
                    Else
                        oProperty.Value = DataField(iCounter).Text
                    End If
                Case idmTypeLong
                    oProperty.Value = CLng(DataField(iCounter).Text)
                Case idmTypeUnsignedLong
                    oProperty.Value = CLng(DataField(iCounter).Text)
                Case idmTypeShort
                    oProperty.Value = CInt(DataField(iCounter).Text)
                Case idmTypeUnsignedShort
                    oProperty.Value = CInt(DataField(iCounter).Text)
                Case idmTypeString
                    oProperty.Value = DataField(iCounter).Text
                Case Else
                    oProperty.Value = DataField(iCounter).Text
                End Select
            End If
        End If
    Next
Next
    
oProperties("F_Busy").Value = Empty
   
If bInsert Then     ' Do the insert
    oQueueEntry.Insert
    QMaint.MainStatusBar.SimpleText = "Insert Successful."
Else                ' Do the update
    oQueueEntry.Update False    ' Don't update F_Busy (it will become False)
    QMaint.MainStatusBar.SimpleText = "Update Successful."
End If
oProperties("F_Busy").Value = False
QMaint.MainStatusBar.Refresh
Exit Sub
ErrorHandler:
     ShowError
        
End Sub
Private Sub UpdateGrid(oQueueEntry As IDMObjects.QueueEntry, _
    oProperties As IDMObjects.Properties, ByRef iIndexNum As Integer, _
    ByVal bInsert As Boolean)
Dim oProperty As IDMObjects.Property
Dim iCounter As Integer
Dim sValue As String        ' string representation of value
Dim iDispWidth As Integer   ' display width of sValue

If bInsert Then
    QMaint.grdQueueData.AddItem ""
    iIndexNum = QMaint.grdQueueData.Rows - 1
    QMaint.cQueueEntries.Add oQueueEntry
End If
If QMaint.SystemCheck.Value = 1 Then
    QMaint.grdQueueData.TextMatrix(iIndexNum, iCounter) = oQueueEntry.EntryId
    ' Auto size column width
    iDispWidth = Len(oQueueEntry.EntryId) * QMaint.grdQueueData.CellFontSize * 12
    If iDispWidth > QMaint.grdQueueData.ColWidth(0) Then
        QMaint.grdQueueData.ColWidth(0) = iDispWidth
    End If
    iCounter = 1
Else
    iCounter = 0
End If
        
For Each oProperty In oProperties
    If oProperty.PropertyDescription.GetState(idmPropCustom) Or _
            QMaint.SystemCheck.Value = 1 Then
        If IsEmpty(oProperty.Value) Or IsNull(oProperty.Value) Or _
                Not oProperty.PropertyDescription.GetState(idmPropSupportsFPNumber) Then
            sValue = CStr(IIf(IsNull(oProperty.Value), "", oProperty.Value))
        Else
            ' Get string from FnFPNumber, so no loss of precision happens
            ' in a conversion to and from a double
            sValue = oProperty.FnFPNumber.ValueAsString
        End If
                QMaint.grdQueueData.TextMatrix(iIndexNum, iCounter) = sValue
        ' Auto size column width
        iDispWidth = Len(sValue) * QMaint.grdQueueData.CellFontSize * 12
        If iDispWidth > QMaint.grdQueueData.ColWidth(iCounter) Then
            QMaint.grdQueueData.ColWidth(iCounter) = iDispWidth
        End If
        iCounter = iCounter + 1
    End If
Next

End Sub
    
Private Sub SaveButton_Click()
    Dim oProperties As IDMObjects.Properties                ' a collection of properties for a QueueEntry
    Dim iCounter As Integer                                 ' an integer Counter
    
    Set oProperties = oQueueEntry.Properties
  
    If Not bInsert Then         ' Update a queue entry
        If bBusy Then
            ' Always confirm overriding someone else's busy state
            If MsgBox("This queue entry is busy.  Save Changes anyway?", vbYesNo + vbQuestion, AppName) _
              = vbNo Then
                Unload Me
                Exit Sub
            Else
            ' If it's busy, we must do a fetch on it, overriding the busy
                If QMaint.FetchQueueEntry(oQueueEntry, iIndexNum, True) Then
                    Set oProperties = oQueueEntry.Properties
                Else
                    ' If you somehow did not retrieve the correct item, do not do the update
                    MsgBox "Unable to unbusy this Queue Entry", vbExclamation, AppName
                    Unload Me
                    Exit Sub
                End If
            End If
        Else                    'Update a non-busy queue entry
            oQueueEntry.MakeReadWrite
        End If
    End If
    
    ' Actually update/insert the entry in the queue
    Call UpdateQueueEntry(oQueueEntry, oProperties, iTotalFields, bInsert)
    
    ' Add the modified data back into the datagrid
    Call UpdateGrid(oQueueEntry, oProperties, iIndexNum, bInsert)
    
    Set oQueueEntry = Nothing
    Set oProperties = Nothing
    
    Unload Me
    
End Sub

Private Sub setVisible(inx As Integer, visibility As Boolean)
    DataLabel(inx).Visible = visibility
    DataField(inx).Visible = visibility
End Sub

Private Sub sbProperties_Change()
    Dim inx As Integer
    Dim pos As Integer
    Dim iHeight As Integer                          ' the height of the standard DataField field
    
    iHeight = DataField(0).Height
    pos = sbProperties.Value
    For inx = 0 To iTotalFields - 1
        If inx >= pos And inx < pos + visibleProps Then
            DataLabel(inx).Top = DataLabel(0).Top + _
                                 (iHeight + iSpacing) * (inx - pos)
            DataField(inx).Top = DataField(0).Top + _
                                 (iHeight + iSpacing) * (inx - pos)
            Call setVisible(inx, True)
        Else
            Call setVisible(inx, False)
        End If
    Next inx
End Sub


Private Function VerifyDate(ByVal InputDateStr$, ByVal ConvertedDate As Date) As Boolean
    Dim convertedyr As Integer, inputyr As Integer
    Dim dformat As Integer
    Dim lcinfo As Long, lctype As Long
    Dim retStr As String, slen As Long
    Dim ConvertDateStr As String
    
    VerifyDate = True
    
    ConvertDateStr = CStr(ConvertedDate)
    If (ConvertDateStr = InputDateStr) Then
        Exit Function                   ' get out early if conversion is ok.
    End If
    
    ' get year from converted date
    convertedyr = DatePart("yyyy", ConvertedDate)
    
    ' get date format from system
    lcinfo = GetThreadLocale()
    lctype = LOCALE_SHORTDATE      ' for now
    ' initialize buffer
    slen = 20
    retStr = String(slen, " ")
    slen = GetLocaleInfo(lcinfo, lctype, retStr, slen)
    
    ' get input year
    inputyr = ScanInputYear(InputDateStr, CInt(retStr))
    If (inputyr < 100) Then
        If (convertedyr >= 100) Then
            convertedyr = convertedyr Mod 100
        End If
    End If
    
    ' now compare if MS switches the year from what was input
    If (convertedyr <> inputyr) Then
        VerifyDate = False             ' bad date return false.
    End If

End Function

Private Function ScanInputYear(ByVal DateStr$, ByVal DateFormat As Integer) As Integer
    Dim tstr1, tstr2 As String, tch As String
    Dim slen As Integer, i As Integer
    Dim Year As Integer, rlen As Integer
        
    slen = Len(DateStr)
    rlen = slen
    tstr1 = DateStr
    tstr2 = " "
    Do While (rlen > 0)
        If (DateFormat = DATEORDER_DDMMYY) Or (DateFormat = DATEORDER_MMDDYY) Then
            If (IsNumeric(Right(tstr1, 1))) Then
                tstr2 = Right(tstr1, 1) + tstr2
                rlen = rlen - 1
                If (rlen > 0) Then
                    tstr1 = Left(tstr1, rlen)
                End If
            Else
                Exit Do
            End If
        Else
            If (IsNumeric(Left(tstr1, 1))) Then
                tstr2 = tstr2 + Left(tstr1, 1)
                rlen = rlen - 1
                 If (rlen > 0) Then
                    tstr1 = Right(tstr1, rlen)
                End If
           Else
                Exit Do
            End If
       End If
    Loop
    
    Year = CInt(tstr2)
    ScanInputYear = Year
End Function
