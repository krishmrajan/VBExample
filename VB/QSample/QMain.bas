Attribute VB_Name = "FnQMaintGlobals"
Option Explicit
Declare Function GetThreadLocale& Lib "kernel32" ()
Declare Function GetLocaleInfo& Lib "kernel32" Alias "GetLocaleInfoA" (ByVal locale As Long, _
                                                                    ByVal lctype As Long, _
                                                                    ByVal outstr As String, _
                                                                    ByVal strlen As Long)
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Const AppName = "Queue Sample"
Public Const iMaxQueues = 1000
Public gWSQQuerySpec(iMaxQueues) As QueueQuerySpecification
Public gQueue(iMaxQueues) As queue
Public oErrorLog As New CErrorLog
Public oTraceLog As New CTraceLog
Public oErrManager As idmError.ErrorManager

Public Sub ShowError()
    Dim oErrCollect As idmError.Errors
    Dim oError As idmError.Error
    Dim iCnt As Integer
    Set oErrCollect = oErrManager.Errors
    If oErrCollect.Count > 1 Then
        iCnt = 1
        For Each oError In oErrCollect
            MsgBox "Error " & iCnt & ": " & oError.Description & " : " & Hex(oError.Number), vbExclamation, AppName
            iCnt = iCnt + 1
        Next
    Else
        If oErrCollect.Count = 1 Then
            oErrManager.ShowErrorDialog
        Else
            If Err.Number <> 0 Then
                MsgBox Err.Description & " : " & Err.Number, vbExclamation, AppName
            End If
        End If
    End If
End Sub

' Build the browse set by querying the queue
Public Sub gBuildBrowseSet(oBrowseSet As IDMObjects.QueueBrowseSet, _
    oQueue As IDMObjects.queue, oQueueQuerySpec As IDMObjects.QueueQuerySpecification, ByVal lMaxRet As Long, lCount As Long)

On Error GoTo ErrorHandler

    QMaint.MainStatusBar.SimpleText = "Create the specification..."
    QMaint.MainStatusBar.Refresh

    ' Get an approximate count of the number of entries
    ' Note that this is only approximate since other operations on the queue from
    ' other workstations could occur before the browse set gets its snapshot
    oQueueQuerySpec.CacheSize = lMaxRet
    lCount = oQueueQuerySpec.Count

    ' Retrieve the records
    QMaint.MainStatusBar.SimpleText = "Browsing the Queue..."
    QMaint.MainStatusBar.Refresh
    
    Set oBrowseSet = oQueueQuerySpec.Browse()
    
Exit Sub
ErrorHandler:
    If Err.Number <> -2147208689 Then  ' No more entries
        oErrorLog.logFNError errWarning, "Error browsing queue", "Unable to retrieve queue data"
    End If
    Set oBrowseSet = Nothing
End Sub

Public Function gCopyQDefinition(oSourceQueue As IDMObjects.queue, oDestQueue As IDMObjects.queue) As Boolean

    Dim oSourcePropDesc As IDMObjects.PropertyDescription       ' The source property description
    Dim oDestPropDesc As IDMObjects.PropertyDescription         ' The dest property description
    Dim bTemp As Boolean

    On Error GoTo ErrorHandler

    QMaint.MainStatusBar.SimpleText = "Copying Queue Definitions Only for: " & oSourceQueue.Name
    QMaint.MainStatusBar.Refresh
        
    oDestQueue.Description = oSourceQueue.Description
    oDestQueue.DefinitionPermissions(idmISAccessRead).GranteeName = oSourceQueue.DefinitionPermissions(idmISAccessRead).GranteeName
    oDestQueue.DefinitionPermissions(idmISAccessWrite).GranteeName = oSourceQueue.DefinitionPermissions(idmISAccessWrite).GranteeName
    oDestQueue.DefinitionPermissions(idmISAccessAX).GranteeName = oSourceQueue.DefinitionPermissions(idmISAccessAX).GranteeName
    oDestQueue.ContentPermissions(idmISAccessRead).GranteeName = oSourceQueue.ContentPermissions(idmISAccessRead).GranteeName
    oDestQueue.ContentPermissions(idmISAccessWrite).GranteeName = oSourceQueue.ContentPermissions(idmISAccessWrite).GranteeName
    oDestQueue.ContentPermissions(idmISAccessAX).GranteeName = oSourceQueue.ContentPermissions(idmISAccessAX).GranteeName

    For Each oSourcePropDesc In oSourceQueue.PropertyDescriptions
        If oSourcePropDesc.GetState(idmPropCustom) Then
            Set oDestPropDesc = oDestQueue.AddPropertyDescription(oSourcePropDesc)
            oDestPropDesc.GetExtendedProperty("F_QUEUETYPEID") = oSourcePropDesc.GetExtendedProperty("F_QUEUETYPEID")
            oDestPropDesc.GetExtendedProperty("F_FIELDSIZE") = oSourcePropDesc.GetExtendedProperty("F_FIELDSIZE")
            oDestPropDesc.GetExtendedProperty("F_ISKEY") = oSourcePropDesc.GetExtendedProperty("F_ISKEY")
            oDestPropDesc.GetExtendedProperty("F_ISUNIQUE") = oSourcePropDesc.GetExtendedProperty("F_ISUNIQUE")
            oDestPropDesc.GetExtendedProperty("F_ISREQUIRED") = oSourcePropDesc.GetExtendedProperty("F_ISREQUIRED")
            bTemp = oSourcePropDesc.GetExtendedProperty("F_ISRENDEZVOUS")  ' Work around
            oDestPropDesc.GetExtendedProperty("F_ISRENDEZVOUS") = bTemp
            oDestPropDesc.GetExtendedProperty("F_SHOULDDISPLAY") = oSourcePropDesc.GetExtendedProperty("F_SHOULDDISPLAY")
            oDestPropDesc.GetExtendedProperty("F_PRECISION") = oSourcePropDesc.GetExtendedProperty("F_PRECISION")
            oDestPropDesc.GetExtendedProperty("F_SCALE") = oSourcePropDesc.GetExtendedProperty("F_SCALE")
        End If
    Next
        
    oDestQueue.SaveNew
    gCopyQDefinition = True
    Exit Function
    
ErrorHandler:
    oErrorLog.logFNError errWarning, "Error copying queue definition.", "Error copying queue definition."
    Screen.MousePointer = vbDefault
    gCopyQDefinition = False
End Function

' Extract data from the browse set and populate the grid control
' As we encounter DocIds, tell the ServerCache about them...
Public Function gCopyQContents(oBrowseSet As IDMObjects.QueueBrowseSet, oSourceQueue As IDMObjects.queue, _
    oDestQueue As IDMObjects.queue, lCount As Long) As Boolean
    
    Dim oDestQueueEntry As IDMObjects.QueueEntry        ' The dest QueueEntry to edit
    Dim oDestProperties As IDMObjects.Properties        ' The dest collection of properties for a QueueEntry
    Dim oSourceProperties As IDMObjects.Properties      ' The source collection of properties for a QueueEntry
    Dim oSourceProperty As IDMObjects.Property          ' The source property

    Dim iRecordCount As Long
    Dim iField As Integer
    Dim sFieldName As String
    Dim bMoreRecords As Boolean
    
    On Error GoTo ErrorHandler

    If oBrowseSet Is Nothing Then
        bMoreRecords = False
    Else
        bMoreRecords = True
    End If
    ' Don't use lCount as an upper bound for the number of records to get,
    ' since it is only an approximation (i.e. do not use in For ... Next).
    ' Also don't use RecordCount because it interferes with the cacheing
    ' mechanism in a QueueBrowseSet and copies all queue entries to the local
    ' workstation.
    iRecordCount = 1

    ' Loop through all the entries
    Do While bMoreRecords
        
        QMaint.MainStatusBar.SimpleText = "Copying Queue Entry " & iRecordCount & " of " & lCount & "..."
        QMaint.MainStatusBar.Refresh
        DoEvents
        
        Set oDestQueueEntry = oDestQueue.CreateEmptyEntry
        Set oDestProperties = oDestQueueEntry.Properties
        Set oSourceProperties = oBrowseSet.Entry.Properties
    
        iField = 1
        For Each oSourceProperty In oSourceProperties
            sFieldName = oSourceProperties(iField).Name             ' Source Field Name
            oDestProperties(sFieldName).Value = oSourceProperties(sFieldName).Value
            iField = iField + 1
        Next
    
        oDestQueueEntry.Insert              ' Insert the entery
           
        If oBrowseSet.MoreResults Then
            oBrowseSet.MoveNext             ' Get the next record
        Else
            bMoreRecords = False
        End If
            
        iRecordCount = iRecordCount + 1
    Loop
    
    gCopyQContents = True
    
    Exit Function

ErrorHandler:
    oErrorLog.logFNError errWarning, "Error copying queue data.", "Error copying queue data."
    Screen.MousePointer = vbDefault
    gCopyQContents = False
End Function
