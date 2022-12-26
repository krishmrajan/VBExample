VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "idmFunction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Public Enum eIDMReturnStatus
    CIDMOk = 0
    CIDMError = 1
    CIDMDriveSelection = 2
    CIDMSelection = 3
    CIDMCancel = 4
End Enum

Public Enum eIDMFlags
    CIDMCopy = 1
    CIDMReference = 2
    CIDMNoOpenFromDrive = 4
    CIDMTrackedFiles = 8
    CIDMNoSaveAsDialog = 16
    CIDMInsert = 32
    CIDMShortCut = 64
End Enum

Implements CIDMMacroAPI
Public oEvent As New idmEvent
Private bRename As Boolean
Private bFirstTime As Boolean
Private bWordDocOpened As Boolean

Public Function IDMMenu(iApplType As Integer, oAppl As Object)

    Dim cbcParentMenu As CommandBarControl
    Dim cbcChildMenu As CommandBarControl
    Dim iMenuPos As Integer
    Dim ChildMenu As CommandBarControl
    
    On Error Resume Next
    'On Error GoTo errHandler
    'copy FileNET logo to clipbaard
    'DTS134785 - Raja - 12/07/04
    Call BkpClpBrddata
    Call LoadMsoDIBToClipboard(9, oAppl)
    
    iMenuPos = Getposition(iApplType, oAppl)
    
    Set cbcParentMenu = oAppl.CommandBars.FindControl(ID:=MB_FILE)
    
    'add "Open FileNET Document"
    Set cbcChildMenu = oAppl.CommandBars.FindControl(Tag:=LoadResString(MNU_FN_OPEN))
    'Set cbcChildMenu = cbcParentMenu.Controls(LoadResString(MNU_FN_OPEN))
    If cbcChildMenu Is Nothing Then
        Set cbcChildMenu = cbcParentMenu.Controls.Add(Type:=msoControlButton, Before:=iMenuPos, Temporary:=True)
        cbcChildMenu.Caption = LoadResString(MNU_FN_OPEN)
        cbcChildMenu.Tag = LoadResString(MNU_FN_OPEN)
        cbcChildMenu.OnAction = "idmFileOpen"
        cbcChildMenu.OLEUsage = msoControlOLEUsageNeither
        cbcChildMenu.BeginGroup = True ' Add seperator bar before FileNET Open...
        cbcChildMenu.PasteFace
    Else
        GoTo NextStep
    End If
    Set cbcChildMenu = Nothing
    
    'add "Add FileNET Document"
    Set cbcChildMenu = oAppl.CommandBars.FindControl(Tag:=LoadResString(MNU_FN_ADD))
    'Set cbcChildMenu = cbcParentMenu.Controls(LoadResString(MNU_FN_ADD))
    If cbcChildMenu Is Nothing Then
        Set cbcChildMenu = cbcParentMenu.Controls.Add(Type:=msoControlButton, Before:=iMenuPos + 1, Temporary:=True)
        cbcChildMenu.Caption = LoadResString(MNU_FN_ADD)
        cbcChildMenu.Tag = LoadResString(MNU_FN_ADD)
        cbcChildMenu.OLEUsage = msoControlOLEUsageNeither
        cbcChildMenu.OnAction = "idmFileAdd"
        cbcChildMenu.PasteFace
    End If
    Set cbcChildMenu = Nothing

    'add "Checkin FileNET Document"
    Set cbcChildMenu = oAppl.CommandBars.FindControl(Tag:=LoadResString(MNU_FN_CHECKIN))
    'Set cbcChildMenu = cbcParentMenu.Controls(LoadResString(MNU_FN_CHECKIN))
    If cbcChildMenu Is Nothing Then
        Set cbcChildMenu = cbcParentMenu.Controls.Add(Type:=msoControlButton, Before:=iMenuPos + 2, Temporary:=True)
        cbcChildMenu.Caption = LoadResString(MNU_FN_CHECKIN)
        cbcChildMenu.Tag = LoadResString(MNU_FN_CHECKIN)
        cbcChildMenu.OLEUsage = msoControlOLEUsageNeither
        cbcChildMenu.OnAction = "idmFileCheckin"
        cbcChildMenu.PasteFace
    End If
    Set cbcChildMenu = Nothing

    'add "Cancel Checkout FileNET Document"
    Set cbcChildMenu = oAppl.CommandBars.FindControl(Tag:=g_FN_CANCEL(iApplType))
    'Set cbcChildMenu = cbcParentMenu.Controls(g_FN_CANCEL(giApplType))
    If cbcChildMenu Is Nothing Then
        Set cbcChildMenu = cbcParentMenu.Controls.Add(Type:=msoControlButton, Before:=iMenuPos + 3, Temporary:=True)
        cbcChildMenu.Caption = g_FN_CANCEL(iApplType)
        cbcChildMenu.Tag = g_FN_CANCEL(iApplType)
        cbcChildMenu.OLEUsage = msoControlOLEUsageNeither
        cbcChildMenu.OnAction = "idmFileCancelCheckout"
        cbcChildMenu.PasteFace
    End If
    Set cbcChildMenu = Nothing

    'add "FileNet Save"
    Set cbcChildMenu = oAppl.CommandBars.FindControl(Tag:=LoadResString(MNU_FN_SAVE))
    'Set cbcChildMenu = cbcParentMenu.Controls(LoadResString(MNU_FN_SAVE))
    If cbcChildMenu Is Nothing Then
        Set cbcChildMenu = cbcParentMenu.Controls.Add(Type:=msoControlButton, Before:=iMenuPos + 4, Temporary:=True)
        cbcChildMenu.Caption = LoadResString(MNU_FN_SAVE)
        cbcChildMenu.Tag = LoadResString(MNU_FN_SAVE)
        cbcChildMenu.OLEUsage = msoControlOLEUsageNeither
        cbcChildMenu.OnAction = "idmFileSave"
        cbcChildMenu.PasteFace
    End If
    Set cbcChildMenu = Nothing

    'add Properties
    Set cbcChildMenu = oAppl.CommandBars.FindControl(Tag:=LoadResString(MNU_FN_PROPERTIES))
    'Set cbcChildMenu = cbcParentMenu.Controls(LoadResString(MNU_FN_PROPERTIES))
    If cbcChildMenu Is Nothing Then
        Set cbcChildMenu = cbcParentMenu.Controls.Add(Type:=msoControlButton, Before:=iMenuPos + 5, Temporary:=True)
        cbcChildMenu.Caption = LoadResString(MNU_FN_PROPERTIES)
        cbcChildMenu.Tag = LoadResString(MNU_FN_PROPERTIES)
        cbcChildMenu.OLEUsage = msoControlOLEUsageNeither
        cbcChildMenu.OnAction = "idmProperties"
        cbcChildMenu.PasteFace
    End If
    Set FnMenuShowProperty = cbcChildMenu
NextStep:
    Set cbcParentMenu = oAppl.CommandBars.FindControl(ID:=MB_INSERT)

    'power point does not have any mezz properties
    If iApplType = APPL_WORD Or iApplType = APPL_EXCEL Then
        ' add Insert Mezzine Properties menu item to INSERT menu
        Set cbcChildMenu = oAppl.CommandBars.FindControl(Tag:=LoadResString(MNU_INSERT_MEZZ_PROP))
        'Set cbcChildMenu = cbcParentMenu.Controls(LoadResString(MNU_INSERT_MEZZ_PROP))
        If cbcChildMenu Is Nothing Then
            Set cbcChildMenu = cbcParentMenu.Controls.Add(Type:=msoControlButton, Before:=9, Temporary:=True)
            cbcChildMenu.Caption = LoadResString(MNU_INSERT_MEZZ_PROP)
            cbcChildMenu.Tag = LoadResString(MNU_INSERT_MEZZ_PROP)
            cbcChildMenu.OLEUsage = msoControlOLEUsageNeither
            cbcChildMenu.OnAction = "idmInsertIDMProperty"
            cbcChildMenu.PasteFace
        End If
        Set cbcChildMenu = Nothing


        Set cbcChildMenu = oAppl.CommandBars.FindControl(Tag:=LoadResString(MNU_UPDATE_MEZZ_PROP))
        'Set cbcChildMenu = cbcParentMenu.Controls(LoadResString(MNU_UPDATE_MEZZ_PROP))
        If cbcChildMenu Is Nothing Then
            Set cbcChildMenu = cbcParentMenu.Controls.Add(Type:=msoControlButton, Before:=10, Temporary:=True)
            cbcChildMenu.Caption = LoadResString(MNU_UPDATE_MEZZ_PROP)
            cbcChildMenu.Tag = LoadResString(MNU_UPDATE_MEZZ_PROP)
            cbcChildMenu.OLEUsage = msoControlOLEUsageNeither
            cbcChildMenu.OnAction = "idmUpdateIDMProperty"
            cbcChildMenu.PasteFace
        End If
        Set cbcChildMenu = Nothing
    End If

    If iApplType = APPL_WORD Then
        Set cbcChildMenu = oAppl.CommandBars.FindControl(Tag:=LoadResString(MNU_INSERT_FILE))
        'Set cbcChildMenu = cbcParentMenu.Controls(LoadResString(MNU_INSERT_FILE))
        If cbcChildMenu Is Nothing Then
            Set cbcChildMenu = cbcParentMenu.Controls.Add(Type:=msoControlButton, Before:=16, Temporary:=True)
            cbcChildMenu.Caption = LoadResString(MNU_INSERT_FILE)
            cbcChildMenu.Tag = LoadResString(MNU_INSERT_FILE)
            cbcChildMenu.OLEUsage = msoControlOLEUsageNeither
            cbcChildMenu.OnAction = "idmInsertFile"
            cbcChildMenu.PasteFace
        End If
        Set cbcChildMenu = Nothing
    End If

    'add Preferences
    Set cbcParentMenu = oAppl.CommandBars.FindControl(ID:=MB_TOOLS)
    Set cbcChildMenu = oAppl.CommandBars.FindControl(Tag:=LoadResString(MNU_FN_PREFERENCES))
    'Set cbcChildMenu = cbcParentMenu.Controls(LoadResString(MNU_FN_PREFERENCES))
    If cbcChildMenu Is Nothing Then
        Select Case iApplType
'               Case APPL_WORD
'                    iMenuPos = 16
               Case APPL_EXCEL
                    iMenuPos = 14
'               Case APPL_POWERPOINT
'                    iMenuPos = 12 ' office 2000
'                    iMenuPos = 14 'office 97
                Case Else
                     Set ChildMenu = oAppl.CommandBars.FindControl(ID:=522) 'Options
                     iMenuPos = ChildMenu.Index
        End Select
        Set cbcChildMenu = cbcParentMenu.Controls.Add(Type:=msoControlButton, Before:=iMenuPos, Temporary:=True)
        cbcChildMenu.Caption = LoadResString(MNU_FN_PREFERENCES)
        cbcChildMenu.Tag = LoadResString(MNU_FN_PREFERENCES)
        cbcChildMenu.OLEUsage = msoControlOLEUsageNeither
        cbcChildMenu.OnAction = "idmPreferences"
        cbcChildMenu.PasteFace
    End If
    Set cbcChildMenu = Nothing

    'add idmHelp item
    Set cbcParentMenu = oAppl.CommandBars.FindControl(ID:=MB_HELP)
    Set cbcChildMenu = oAppl.CommandBars.FindControl(Tag:=LoadResString(MNU_FN_HELP))
    'Set cbcChildMenu = cbcParentMenu.Controls(LoadResString(MNU_FN_HELP))
    If cbcChildMenu Is Nothing Then
        iMenuPos = 2
        Set cbcChildMenu = cbcParentMenu.Controls.Add(Type:=msoControlButton, Before:=iMenuPos, Temporary:=True)
        cbcChildMenu.Caption = LoadResString(MNU_FN_HELP)
        cbcChildMenu.Tag = LoadResString(MNU_FN_HELP)
        cbcChildMenu.OLEUsage = msoControlOLEUsageNeither
        cbcChildMenu.OnAction = "idmHelp"
        cbcChildMenu.PasteFace
    End If
    Set cbcChildMenu = Nothing

'NextStep:

    'the following part to handle the native stuff
    If iApplType = APPL_EXCEL Then
        Dim iCount As Integer
        iCount = 0
        Set cbcParentMenu = oAppl.CommandBars.FindControl(ID:=MB_FILE)
        For Each cbcChildMenu In cbcParentMenu.Controls
          'If cbcChildMenu.BuiltIn = True Then
                If (cbcChildMenu.ID = MB_FILE_SAVE) Then
                     cbcChildMenu.OnAction = "msSave"
                     iCount = iCount + 1
                End If
                If (cbcChildMenu.ID = MB_FILE_SAVEAS) Then
                     cbcChildMenu.OnAction = "msSaveAs"
                     iCount = iCount + 1
                End If
                If iCount > 1 Then Exit For
          'End If
        Next cbcChildMenu
        
        'For Each cbcChildMenu In cbcParentMenu.Controls
          'If cbcChildMenu.BuiltIn = True Then
                'If (cbcChildMenu.ID = MB_FILE_SAVE) Then
                '     cbcChildMenu.OnAction = "msSave"
                '     Exit For
                'End If
          'End If
        'Next cbcChildMenu
    
    Set cbcChildMenu = Nothing
        'Modify Excel 97/2000 Native Save/SaveAs
        'oAppl.CommandBars.FindControl(ID:=MB_FILE_SAVE).OnAction = "msSave"
        'cbcChildMenu.OnAction = "msSave"
        'oAppl.CommandBars.FindControl(ID:=MB_FILE_SAVEAS).OnAction = "msSaveas"
        'cbcChildMenu.OnAction = "msSaveas"
        'oAppl.CommandBars("Worksheet Menu bar").Controls("File").Controls("Save").OnAction = "msSave"
        'oAppl.CommandBars("Worksheet Menu bar").Controls("File").Controls("Save As...").OnAction = "msSaveAs"
        
        'modify Excel 97/2000 native save toolbar button
        'oAppl.CommandBars("Standard").Controls("Save").OnAction = "msSave"
        Dim oStandardToolbar As CommandBar
        Dim oBtn As CommandBarControl
        Set oStandardToolbar = oAppl.CommandBars("Standard")
        For Each oBtn In oStandardToolbar.Controls
            If (oBtn.BuiltIn = True And oBtn.ID = 3) Then
                oBtn.OnAction = "msSave"
                Exit For
            End If
        Next oBtn
        
        'enable close event in Excel so ANY workbook close will call idmfileclose
        'this is likely to change, probably going to switch idmFileClose to
        'work off the menu instead of the event.
        oAppl.EnableEvents = True
    End If

    If iApplType = APPL_WORD Then
      
        'modify word native SaveAs menuitems in word
        
        oAppl.CommandBars.FindControl(ID:=MB_FILE_SAVEAS).OnAction = "msSaveas"
        'oAppl.CommandBars("Menu bar").Controls("File").Controls("Save As...").OnAction = "msSaveas"
        
        '3.3 we do not support office97
        'If (InStr(1, oAppl.Version, "8", vbTextCompare)) Then

            'modify word 97 native save menuitems in word
            'oAppl.CommandBars("Menu bar").Controls("File").Controls("Save").OnAction = "msSave"

            'modify word 97 native save toolbar button
            'oAppl.CommandBars("Standard").Controls("Save").OnAction = "msSave"

        'End If
        
        
        'modify save menuitems in word
        'since we have just changed normal.dot we need to reset .saved
        'so will not ask us to save to normal.dot when we exit
        Call saveTemplate(oAppl)
        
        bWordDocOpened = True

    End If
    Set cbcParentMenu = Nothing

    Call ClearClipboard
    'DTS134785 - Raja - 12/07/04
    Call RestoreClpBrddata

    'Exit Function

'errHandler:
    'If Err.Number <> -2147467259 Then
    '    If (Err.Number = 9 And iApplType = APPL_POWERPOINT) Then
    '       'Index out of range
    '    Else
    '       MsgBox Err.Description, vbCritical, LoadResString(DLG_ERROR_TITLE)
    '       Resume Next
    '    End If
    'End If
    'Call ClearClipboard
End Function

Private Function IDMStart(iApplType As Integer, oAppl As Object)
    Dim bShowToolbar As Boolean

    On Error Resume Next

    bFirstTime = True
    bShowToolbar = CBool(GetPreferenceValue(LoadResString(STR_SHOW_TOOLBAR), iApplType))
    If bShowToolbar Then
       Call ShowToolbar(oAppl, iApplType)
    End If

    Select Case iApplType
        Case APPL_WORD
            If bShowToolbar Then
                Call saveTemplate(oAppl)
            End If
        Case APPL_EXCEL
            oAppl.EnableEvents = True

        Case APPL_POWERPOINT
            'no way to set at this point
    End Select
    
End Function

Public Function CIDMMacroAPI_IDMCancelCheckout(iApplType As Integer, strPathName As String, Optional oAppl As Object) As Long
    Dim whandle As Long
    
    If (DocContained(oAppl, iApplType) = NoContainer) Then
       CIDMMacroAPI_IDMCancelCheckout = fileCancelCheckout(oAppl, iApplType, strPathName)
    Else
       whandle = GetActiveWindow()
       MessageBox whandle, LoadResString(MSG_OPERATION_NOT_DEFINED), LoadResString(DLG_WARNING), MB_OK Or MB_ICONEXCLAMATION
    End If
End Function
 
Public Function CIDMMacroAPI_IDMDocStatus(strPathName As String) As Integer
    CIDMMacroAPI_IDMDocStatus = GetDocStatus(strPathName)
End Function

Public Function CIDMMacroAPI_IDMFileAdd(iApplType As Integer, strPathName As String, Optional oAppl As Variant) As Integer
    
    Dim oTemp As Object
    Dim whandle As Long
    If (Not IsMissing(oAppl)) And (CStr(oAppl) <> "Outlook") Then
       Set oTemp = oAppl
    End If
    If (DocContained(oTemp, iApplType) = NoContainer) Then
        CIDMMacroAPI_IDMFileAdd = fileAdd(oTemp, iApplType, strPathName)
    Else
       whandle = GetActiveWindow()
       MessageBox whandle, LoadResString(MSG_OPERATION_NOT_DEFINED), LoadResString(DLG_WARNING), MB_OK Or MB_ICONEXCLAMATION
    End If
End Function

Public Function CIDMMacroAPI_IDMFileCheckin(iApplType As Integer, strPathName As String, Optional oAppl As Object) As Long
    
    Dim whandle As Long
    If (DocContained(oAppl, iApplType) = NoContainer) Then
       CIDMMacroAPI_IDMFileCheckin = fileCheckin(oAppl, iApplType, strPathName)
    Else
       whandle = GetActiveWindow()
       MessageBox whandle, LoadResString(MSG_OPERATION_NOT_DEFINED), LoadResString(DLG_WARNING), MB_OK Or MB_ICONEXCLAMATION
    End If
End Function

Public Function CIDMMacroAPI_IDMFileOpen(iApplType As Integer, strPathName As String, Optional oAppl As Variant, Optional eFilter As Variant, Optional eflag = 0, Optional vPathNames As Variant) As Integer
    Dim oTemp As Object
    Dim iflag As Integer
    Dim sFilter() As String
    Dim Index As Integer
    Dim whandle As Long
    
    If (Not IsMissing(eFilter)) Then
        ReDim sFilter(UBound(eFilter))
        For Index = LBound(sFilter) To UBound(sFilter)
            sFilter(Index) = eFilter(Index)
        Next Index
    End If
    iflag = eflag
    If (Not IsMissing(oAppl)) And (CStr(oAppl) <> "Outlook") Then
       Set oTemp = oAppl
    End If
    If (DocContained(oTemp, iApplType) <> ContainerInt) Then
        CIDMMacroAPI_IDMFileOpen = fileOpen(oTemp, iApplType, strPathName, sFilter, iflag, vPathNames)
    Else
       whandle = GetActiveWindow()
       MessageBox whandle, LoadResString(MSG_OPERATION_NOT_DEFINED), LoadResString(DLG_WARNING), MB_OK Or MB_ICONEXCLAMATION
    End If
End Function

Public Function CIDMMacroAPI_IDMGetHelp(iApplType As Integer) As Integer
    GetHelp iApplType
End Function

Public Function CIDMMacroAPI_IDMGetMezzPropertyValue(iApplType As Integer, oAppl As Object, strPropertyName As String, strPropertyValue As String) As Long

    'first get an active file name by calling GetActiveFileName(iApplType, oAppl)
    
    strPropertyValue = GetPropValue(strPropertyName, GetActiveFileName(iApplType, oAppl))
    Select Case strPropertyValue
           Case " " 'LoadResString(TXT_NO_THIS_VALUE) '"NoValue"
                 CIDMMacroAPI_IDMGetMezzPropertyValue = CIDMError
           Case Else
                 CIDMMacroAPI_IDMGetMezzPropertyValue = CIDMOk
    End Select
End Function

Public Function CIDMMacroAPI_IDMInsertMezzProperties(iApplType As Integer, oAppl As Object) As Long
Dim whandle As Long
    If (DocContained(oAppl, iApplType) = NoContainer) Then
        'To get a win handler
        gpWnd = GetFocus
        'Close Property Manager
        Call CheckPropMgrStatus
        If DocCount(oAppl, iApplType) > 0 Then
            'only if there are documents
            Call ShowPropertyMgr(oAppl, iApplType)
        End If
    Else
       whandle = GetActiveWindow()
       MessageBox whandle, LoadResString(MSG_OPERATION_NOT_DEFINED), LoadResString(DLG_WARNING), MB_OK Or MB_ICONEXCLAMATION
    End If
End Function

Public Function CIDMMacroAPI_IDMMenu(iApplType As Integer, oAppl As Object) As Long
   
   On Error Resume Next
   
   Call IDMMenu(iApplType, oAppl)
   Set oAppObject = oAppl
   giApplType = iApplType
   
End Function

Public Function CIDMMacroAPI_IDMSave(iApplType As Integer, strPathName As String, Optional oAppl As Object) As Long
    Dim whandle As Long
    If (DocContained(oAppl, iApplType) = NoContainer) Then
        CIDMMacroAPI_IDMSave = fileNetSave(oAppl, iApplType, strPathName)
    Else
       whandle = GetActiveWindow()
       MessageBox whandle, LoadResString(MSG_OPERATION_NOT_DEFINED), LoadResString(DLG_WARNING), MB_OK Or MB_ICONEXCLAMATION
    End If
End Function

Public Function CIDMMacroAPI_IDMShowPreferences(iApplType As Integer, oAppl As Variant) As Integer
    Dim oTemp As Object
    Dim whandle As Long
    Dim vDocClose As Variant
    
    On Error Resume Next
    
    If (Not IsMissing(oAppl)) And (CStr(oAppl) <> "Outlook") Then
       Set oTemp = oAppl
    End If
    
    
    Call ShowPreferences(oTemp, iApplType)
    Call IDMUpdateMenu(iApplType, oTemp, vDocClose)
    CIDMMacroAPI_IDMShowPreferences = CIDMOk
    
End Function

Public Function CIDMMacroAPI_IDMGetPreferenceValue(strPreference As String, iApplType As Integer, varPreferenceValue As Variant) As Integer

    varPreferenceValue = GetPreferenceValue(strPreference, iApplType)

End Function

Public Function CIDMMacroAPI_IDMShowProperties(iApplType As Integer, strPathName As String, Optional oAppl As Object) As Long
    Dim whandle As Long
    
    If (DocContained(oAppl, iApplType) = NoContainer) Then
       Call showProperties(oAppl, iApplType)
    Else
       whandle = GetActiveWindow()
       MessageBox whandle, LoadResString(MSG_OPERATION_NOT_DEFINED), LoadResString(DLG_WARNING), MB_OK Or MB_ICONEXCLAMATION
    End If
End Function

Public Function CIDMMacroAPI_IDMStart(iApplType As Integer, oAppl As Object) As Long
    
    On Error Resume Next
    
    initializeVars (iApplType)

    Call IDMStart(iApplType, oAppl)
    
    oAppl.ActiveDocument.Saved = True
       
End Function
Public Function CIDMMacroAPI_IDMUpdateMenu(iApplType As Integer, oAppl As Object, Optional vDocClose As Variant) As Long
    
    On Error Resume Next
    
    If ((CBool(GetPreferenceValue("UpdateMenuToolbar", iApplType)) = False) And (bFirstTime = True)) Then
    
    Else
        Call IDMUpdateMenu(iApplType, oAppl, vDocClose)
    End If
    
    bFirstTime = False
    
    If gbFileNETSave = False Then
       Call CheckPropMgrStatus
    End If
    gbFileNETSave = False
    
    If bWordDocOpened = True Then
        oAppl.ActiveDocument.Saved = True
        bWordDocOpened = False
    End If
    
End Function

Public Function CIDMMacroAPI_IDMSaveDialog(iApplType As Integer, strPathName As String, Optional oAppl As Object) As Long

  If (DocSaveDialog(oAppl, strPathName, iApplType, False) = vbCancel) Then
      CIDMMacroAPI_IDMSaveDialog = CIDMCancel
  Else
      CIDMMacroAPI_IDMSaveDialog = CIDMOk
  End If

End Function

Public Function CIDMMacroAPI_IDMGetResource(strResourceName As String) As String
 
    Select Case strResourceName
        Case "Default Checkout Path":
            CIDMMacroAPI_IDMGetResource = DEFAULT_CHECKOUT_PATH
        Case "Default Copy Path"
            CIDMMacroAPI_IDMGetResource = DEFAULT_COPY_PATH
        Case "Msg File Save Change":
            CIDMMacroAPI_IDMGetResource = LoadResString(MSG_FILE_SAVE_CHANGES)
    End Select

End Function


Public Function CIDMMacroAPI_IDMUpdateMezzProperties(iApplType As Integer, oAppl As Object) As Long
    Dim whandle As Long
    
    If (DocContained(oAppl, iApplType) = NoContainer) Then
        Call UpdateMezzProperties(oAppl, iApplType, LoadResString(STR_UPDATE_PROPERTIES))
        CIDMMacroAPI_IDMUpdateMezzProperties = CIDMOk
    Else
       whandle = GetActiveWindow()
       MessageBox whandle, LoadResString(MSG_OPERATION_NOT_DEFINED), LoadResString(DLG_WARNING), MB_OK Or MB_ICONEXCLAMATION
    End If
End Function

Private Sub Class_Initialize()
    
    On Error Resume Next
    
    setDefaults
    Set gIdmEvent = oEvent
    
    bRename = False
    Set gdoc = Nothing

End Sub
Private Function setDefaults()

    On Error Resume Next
    
    Dim varReturn As Variant
    Dim sdirectory As String
    
    sdirectory = String(260, vbNullChar)
    GetTempPath 260, sdirectory
    sdirectory = Left(sdirectory, InStr(1, sdirectory, vbNullChar) - 2)
    
    If DSinstalled = True Then
        varReturn = GetDirectory(LoadResString(STR_COPYDIR)) '"CopyDir")
        DEFAULT_COPY_PATH = CStr(varReturn)
        varReturn = GetDirectory(LoadResString(STR_CHECKOUTDIR)) '"CheckoutDir")
        DEFAULT_CHECKOUT_PATH = CStr(varReturn)
    Else
        DEFAULT_COPY_PATH = sdirectory
        DEFAULT_CHECKOUT_PATH = sdirectory
    End If
End Function

Public Function IDMReadDefaultSavePath(iApplType As Integer) As String

    IDMReadDefaultSavePath = readDefaultSavePath(iApplType)

End Function

Public Sub msSaveas(iApplType As Integer, oAppl As Object)
    Dim sName As String
    Dim lSaveformat As Long
    Dim lResults As Long
  
    On Error GoTo errHandler
    Select Case iApplType
           Case APPL_WORD
                oAppl.Dialogs(wdDialogFileSaveAs).Show
                
           Case APPL_EXCEL
                oAppl.Dialogs(xlDialogSaveAs).Show
    
    End Select
    bRename = True
    
    ' to handle OfficeXP case
    If iApplType = APPL_WORD Then
        If (oAppl.Version <> "10.0") Then
            'refresh the caption
            Dim sCaption As String
            sCaption = getName(oAppl, iApplType)
            If sCaption <> "" Then
                oAppl.ActiveWindow.Caption = ResetCaption(oAppl, iApplType, 0)
            End If
         End If
    End If
    Exit Sub
errHandler:
    If Err.Number = 4198 Then
       bRename = False
       Exit Sub
    End If
End Sub
Public Sub msSave(iApplType As Integer, oAppl As Object)
    Dim sName As String
    Dim lSaveformat As Long
    Dim lResults As Long
    Dim sFilePath As String
    Dim sDir As String
    
    'If DocCount(oAppl, iApplType) = 0 Then
    '   Exit Sub
    'End If
    
    On Error Resume Next
    
    sFilePath = getFullName(oAppl, iApplType)
    idmGetDirectoryAndFileName sFilePath, sDir, sName
    If sDir <> "" Then
        Select Case iApplType
               Case APPL_WORD
                    oAppl.Documents(sFilePath).Save
               Case APPL_EXCEL
                    oAppl.ActiveWorkbook.Save
       End Select
       
       GoTo Done
    
    End If
    
    Select Case iApplType
           Case APPL_WORD
                With oAppl.Dialogs(wdDialogFileSaveAs)
                    lResults = .Display
                    If lResults Then
                        sName = .Name
                        lSaveformat = .Format
                        .Update
                        .Name = sName
                        .Format = lSaveformat
                        .Execute
                    End If
                End With
            
         Case APPL_EXCEL
                oAppl.Dialogs(xlDialogSaveAs).Show
    End Select
    
Done:
    
    'refresh the caption
    Dim sCaption As String
    sCaption = getName(oAppl, iApplType)
    If sCaption <> "" Then
        oAppl.ActiveWindow.Caption = ResetCaption(oAppl, iApplType, 0)
    End If
    
End Sub
Private Sub Class_Terminate()
     
     On Error Resume Next
     
     Set oEvent = Nothing
     Set goCmnDlg = Nothing
     If oAppObject.Name = LoadResString(STR_MS_WORD) Then
        'If (InStr(1, oAppObject.Version, "9", vbTextCompare)) Then
            Call saveTemplate(oAppObject)
        'End If
     End If
     'Call ResetMSMenuItem(oAppObject, MB_FILE, MB_FILE_SAVEAS, True)
     'Call Resetoffice97Save(oAppObject, MB_FILE, MB_FILE_SAVE, True)
        
     'Office 2000/97 Excel case
     If (giApplType <> APPL_POWERPOINT) Then
        
        Call ResetMenuItem(oAppObject)
        
     End If
     
     'office 2000 case-word
     If oAppObject.Name = LoadResString(STR_MS_WORD) Then
        'If (InStr(1, oAppObject.Version, "9", vbTextCompare)) Then
            Call saveTemplate(oAppObject)
        'End If
     End If
     
End Sub
Private Function Getposition(iApplType As Integer, oAppl As Object) As Integer
   Dim cbcParentMenu As CommandBarControl
   Dim cbcChildMenu As CommandBarControl
   Select Case iApplType
          Case APPL_WORD
               Getposition = 8
          Case APPL_EXCEL
               Set cbcChildMenu = oAppl.CommandBars.FindControl(ID:=30255)
               Getposition = cbcChildMenu.Index - 1
          Case APPL_POWERPOINT
               Getposition = 8
          Case Else
               GoTo Done
   End Select
Done:
End Function

Private Sub ResetMenuItem(oAppl As Object)
     
     Dim cbcMenu As CommandBarControl
     Dim cbcParentMenu As CommandBarControl
     Dim cbcChildMenu As CommandBarControl

     Dim OStdToolbar As CommandBar
     
     On Error Resume Next
          
     'reset file menu
     '******Start - DTS 119816 - 07/29/2004 - Raja******
     'Set cbcMenu = oAppl.CommandBars.FindControl(ID:=MB_FILE)
     Set cbcMenu = oAppl.CommandBars.FindControl(Tag:=LoadResString(MNU_FN_OPEN))
     cbcMenu.Delete
     Set cbcMenu = oAppl.CommandBars.FindControl(Tag:=LoadResString(MNU_FN_ADD))
     cbcMenu.Delete
     Set cbcMenu = oAppl.CommandBars.FindControl(Tag:=LoadResString(MNU_FN_CHECKIN))
     cbcMenu.Delete
     Set cbcMenu = oAppl.CommandBars.FindControl(Tag:=g_FN_CANCEL(2))
     cbcMenu.Delete
     Set cbcMenu = oAppl.CommandBars.FindControl(Tag:=LoadResString(MNU_FN_SAVE))
     cbcMenu.Delete
     Set cbcMenu = oAppl.CommandBars.FindControl(Tag:=LoadResString(MNU_FN_PROPERTIES))
     cbcMenu.Delete
     
     Dim iCount As Integer
     iCount = 0
     Set cbcParentMenu = oAppl.CommandBars.FindControl(ID:=MB_FILE)
     For Each cbcChildMenu In cbcParentMenu.Controls
        If (cbcChildMenu.ID = MB_FILE_SAVE) Then
             cbcChildMenu.Reset
             cbcChildMenu.Enabled = True
             iCount = iCount + 1
        End If
        If (cbcChildMenu.ID = MB_FILE_SAVEAS) Then
             cbcChildMenu.Reset
             iCount = iCount + 1
        End If
        If iCount > 1 Then Exit For
     Next cbcChildMenu
    
     Set cbcChildMenu = Nothing
    
     'cbcMenu.Reset
          
     '******End - DTS 119816 - 07/29/2004 - Raja******
     
     'reset insert menu
     Set cbcMenu = oAppl.CommandBars.FindControl(Tag:=LoadResString(MNU_INSERT_MEZZ_PROP))
     cbcMenu.Delete
     Set cbcMenu = oAppl.CommandBars.FindControl(Tag:=LoadResString(MNU_UPDATE_MEZZ_PROP))
     cbcMenu.Delete
     Set cbcMenu = oAppl.CommandBars.FindControl(Tag:=LoadResString(MNU_INSERT_FILE))
     cbcMenu.Delete
     
     'tools menu
     Set cbcMenu = oAppl.CommandBars.FindControl(Tag:=LoadResString(MNU_FN_PREFERENCES))
     cbcMenu.Delete
     
     'reset help menu
     Set cbcMenu = oAppl.CommandBars.FindControl(ID:=MB_HELP)
     cbcMenu.Controls(LoadResString(MNU_FN_HELP)).Delete
          
     'reset Standard toolbar
     '****** DTS 135721 - Raja - We donot need to reset the standard toolbar as we are not making any changes to it.
     'Set OStdToolbar = oAppl.CommandBars("Standard")
     'OStdToolbar.Reset

End Sub
