Attribute VB_Name = "fileOpenbas"
Option Explicit
Public Function fileOpen(oAppl As Object, iApplType As Integer, strfilepath As String, strExtString() As String, eflag As Integer, Optional vPathNames As Variant) As Long
    Dim oVer As IDMObjects.Version
    Dim oCom As IDMObjects.Compound
    Dim oDoc As IDMObjects.Document
    Dim oStoredSearch As IDMObjects.StoredSearch
    Dim oRetObject As Object
    Dim oProp As IDMObjects.Property
    Dim oErrorMgr As ErrorManager
    Dim lOperation As Long
    Dim lReturn As Long
    Dim strfilename As String
    Dim strFullName As String
    Dim strDirectory As String
    Dim lvalue As Long
    Dim sTempPath As String * 20
    Dim DEFAULT_INSERT_PATH As String
    Dim PageRange As frmPageRange
    Dim i As Integer
    Dim PageFirst, PageLast As Integer
    Dim sPathNames() As String
    Dim vbResult As Variant
               
    On Error GoTo errHandler
    Set oErrorMgr = CreateObject(CREATE_ERR_MGR)
    If oErrorMgr Is Nothing Then
        Err.Raise 1, LoadResString(MSG_OPEN), LoadResString(MSG_CANNOT_CREATE_ERRMGR)
    End If
    
    ' always set Page params to 0 unless user opens an IS doc
    PageFirst = 0
    PageLast = 0
    
    Select Case iApplType
        Case APPL_WORD:
            If (CIDMInsert = (CIDMInsert And eflag)) Then  'FileNET insert a file
                ReDim strExtString(4)
                strExtString(1) = LoadResString(EXT_WORD3) ' "Document Templates (*.dot)"
                strExtString(2) = LoadResString(EXT_WORD2) '"Word Documents (*.doc)"
                strExtString(3) = LoadResString(EXT_WORD4) '"Rich Text Format (*.rtf)"
                strExtString(4) = LoadResString(EXT_WORD0) '"Text Files (*.txt)"
            Else  'open a file
                ReDim strExtString(17)
                strExtString(1) = LoadResString(EXT_WORD1) '"All Files (*.*)"
                strExtString(2) = LoadResString(EXT_WORD2) ' "Word Documents (*.doc)"
                strExtString(3) = LoadResString(EXT_WORD3) ' "Document Templates (*.dot)"
                strExtString(4) = LoadResString(EXT_WORD4) ' "Rich Text Format (*.rtf)"
                strExtString(5) = LoadResString(EXT_WORD5) ' "Text Files/Unicode Text Files (*.txt)"
                strExtString(6) = LoadResString(EXT_WORD6) '"Lotus 1-2-3 (*.wk1;*.wk3;*.wk4)"
                strExtString(7) = LoadResString(EXT_WORD7) '"MS-Doc Text with Layout (*.asc)"
                strExtString(8) = LoadResString(EXT_WORD8) '"Personal Address Book (*.pab)"
                strExtString(9) = LoadResString(EXT_WORD9) '"Outlook Address Book (*.olk)"
                strExtString(10) = LoadResString(EXT_WORD10) ' "Schedule + Contacts (*.scd)"
                strExtString(11) = LoadResString(EXT_WORD11) '"Text with Layout (*.ans)"
                strExtString(12) = LoadResString(EXT_WORD12) '"HTML Document (*.htm;*.html;*.htx*)"
                strExtString(13) = LoadResString(EXT_WORD13) '"Microsoft Excel Worksheet (*.xls;*.xlw)"
                strExtString(14) = LoadResString(EXT_WORD14) '"Word 4.0-5.1 for Macintosh (*.mcw)"
                strExtString(15) = LoadResString(EXT_WORD15) '"Works 3.0-4.0 for Windows (*.wps)"
                strExtString(16) = LoadResString(EXT_WORD16) '"WordPerfect 6.x (*.doc;*.wpd)"
                strExtString(17) = LoadResString(EXT_WORD17) '"WordPerfect 5.x (*.doc)"
            End If
 
        Case APPL_EXCEL:
            ReDim strExtString(19)
            strExtString(1) = LoadResString(EXT_EXCEL1) '"All Files (*.*)"
            strExtString(2) = LoadResString(EXT_EXCEL2) ' "Microsoft Excel Files (*.xls;*.xla)"
            strExtString(3) = LoadResString(EXT_EXCEL3) '"Text Files (*.prn;*.txt;*.csv)"
            strExtString(4) = LoadResString(EXT_EXCEL4) '"Lotus 1-2-3 Files (*.wk1;*.wk3;*.wk4)"
            strExtString(5) = LoadResString(EXT_EXCEL5) '"Quatro Pro/Dos Files (*.wq1)"
            strExtString(6) = LoadResString(EXT_EXCEL6) '"Microsoft Works 2.0 Files (*.wks)"
            strExtString(7) = LoadResString(EXT_EXCEL7) '"dBase Files (*.dbf)"
            strExtString(8) = LoadResString(EXT_EXCEL8) '"Microsoft Excel 4.0 Macros (*.xlm;*.xla)"
            strExtString(9) = LoadResString(EXT_EXCEL9) '"Microsoft Excel 4.0 Charts (*.xlc)"
            strExtString(10) = LoadResString(EXT_EXCEL10) '"Microsoft Excel 4.0 Workbooks (*.xlw)"
            strExtString(11) = LoadResString(EXT_EXCEL11) '"Worksheets (*.xls)"
            strExtString(12) = LoadResString(EXT_EXCEL12) '"Workspace (*.xlw)"
            strExtString(13) = LoadResString(EXT_EXCEL13) '"Templates (*.xlt)"
            strExtString(14) = LoadResString(EXT_EXCEL14) '"Add-Ins (*.xla;*.xll)"
            strExtString(15) = LoadResString(EXT_EXCEL15) '"Toolbars (*.xlb)"
            strExtString(16) = LoadResString(EXT_EXCEL16) '"SYLK Files (*.slk)"
            strExtString(17) = LoadResString(EXT_EXCEL17) ' "Data Interchange Format (*.dif)"
            strExtString(18) = LoadResString(EXT_EXCEL18) '"Backup Files (*.xlk;*.bak)"
            strExtString(19) = LoadResString(EXT_EXCEL19) '"HTML Documents (*.html;*.htm)"
      
        Case APPL_POWERPOINT:
            ReDim strExtString(6)
            strExtString(1) = LoadResString(EXT_PP1) '"All Files (*.*)"
            strExtString(2) = LoadResString(EXT_PP2) '"Presentations and Shows (*.ppt;*.pps)"
            strExtString(3) = LoadResString(EXT_PP3) '"Presentation Templates (*.pot)"
            strExtString(4) = LoadResString(EXT_PP4) '"All Outlines (*.txt;*.rtf;*.doc;*.wpd;*.dot;*.xls;*.xlw;*.htm;*.html;*.htx;*.otm)"
            strExtString(5) = LoadResString(EXT_PP5) '"HTML Documents (*.html;*.htm;*.htx;*.otm)"
            strExtString(6) = LoadResString(EXT_PP6) '"PowerPoint Add-Ins (*.ppa)"
          
'        Case APPL_OUTLOOK:
'            strExtString(1) = "All Files (*.*)"
'            strExtString(2) = "Mail Messages (*.msg)"
'            strExtString(3) = "Document Templates (*.dot)"
'            strExtString(4) = "Rich Text Format (*.rtf)"
'            strExtString(5) = "Text Files (*.txt)"
            
'       Case APPL_WORDPRO:
'            strExtString1(1) = "All Files (*.*)"
'            strExtString1(2) = "Word Documents (*.doc)"
'            strExtString1(3) = "Document Templates (*.dot)"
'            strExtString1(4) = "Rich Text Format (*.rtf)"
'            strExtString1(5) = "Text Files (*.txt)"
        
    End Select
        
    If goCmnDlg Is Nothing Then
        Set goCmnDlg = CreateObject(CREATE_COMMON_DLG)
    End If
    
    With goCmnDlg
        .Title = LoadResString(OPEN_FN_DOC) 'needs to be a constant
        .Extensions = strExtString
        .ExtensionDefault = 2
        If CIDMShortCut = (CIDMShortCut And eflag) Then
           .OpenMode = idmOpenReturnStoredSearch
        Else
           .OpenMode = idmOpenExecuteStoredSearch
        End If
        .ShowObjectType = idmSelectDocuments Or idmSelectStoredSearches
        .hWnd = GetActiveWindow
        .OpenAsDefault = idmOpenAsCheckout
        'open button title
        If (CIDMInsert = (CIDMInsert And eflag)) Then   'FileNET insert File for word
           .OpenButtonText = LoadResString(TXT_INSERT) ' "Insert"
           If iApplType = APPL_OUTLOOK Then
             .Title = LoadResString(INSERT_FN_ATTACHMENT)
           Else
             .Title = LoadResString(INSERT_FN_FILE)
           End If
        ElseIf (CIDMShortCut = (CIDMShortCut And eflag)) Or _
                 (CIDMReference = (CIDMReference And eflag)) Then  'File insert Attachment for Outlook
           .OpenButtonText = LoadResString(TXT_INSERT) ' "Insert"
           .Title = LoadResString(INSERT_FN_ATTACHMENT)
        Else
           .OpenButtonText = LoadResString(TXT_OPEN) ' "Open"
        End If
        'lvalue stuff

        lvalue = idmSelectShowFileNameAndType + idmSelectShowAnnotations + idmSelectHideVersionsTab
         ' + idmSelectHideOpenAsView
        If (CIDMNoOpenFromDrive = (CIDMNoOpenFromDrive And eflag)) Then
            lvalue = lvalue + idmSelectHideDrives
        End If
        If (CIDMCopy = (CIDMCopy And eflag)) Or _
        (CIDMReference = (CIDMReference And eflag)) Or _
        (CIDMInsert = (CIDMInsert And eflag)) Or _
        (CIDMShortCut = (CIDMShortCut And eflag)) Then
            If (CIDMCopy = (CIDMCopy And eflag)) And (iApplType = APPL_OUTLOOK) Then
               ' special case for chris open
               lvalue = lvalue + idmSelectHideOpenAsCheckout
            Else
               lvalue = lvalue + idmSelectHideOpenAs ' idmSelectHideOpenAsCheckout ' idmSelectHideOpenAs + idmSelectHideOpenAsView
            End If
        Else
            lvalue = lvalue + idmSelectShowTrackedFiles
        End If
        .Options = lvalue
    End With
    
    Call goCmnDlg.SelectDocument(oRetObject, lOperation)
    If oRetObject Is Nothing Then
       'GoTo NextStep
    ElseIf oRetObject.ObjectType = idmObjTypeDocument Then
        Set oDoc = oRetObject
    ElseIf oRetObject.ObjectType = idmObjTypeStoredSearch Then
         Set oStoredSearch = oRetObject
         strfilepath = oStoredSearch.CreateShortcut(strfilepath)
         fileOpen = CIDMOk
         GoTo Done
    End If
    If oDoc Is Nothing Then
         Select Case lOperation
            Case IDMObjects.idmOperationCancel
               fileOpen = CIDMCancel
               GoTo Done
            
            Case IDMObjects.idmOperationDrives
                'this means they want to use a Local drive
                Select Case iApplType
                    Case APPL_WORD:
                        lReturn = oAppl.Dialogs(wdDialogFileOpen).Show
                        'don't do anything with the return value right now, maybe later.
                   
                    Case APPL_EXCEL:
                        'Display File Open dialog and get back file selected.
                        'strfilepath = oAppl.GetOpenFilename
                        'If strfilepath <> "False" Then
                        On Error Resume Next
                        Err.Number = 0
                        frmCommonDialogs.OpenDialogSetup (APPL_EXCEL) 'initalize common dialog form
                        frmCommonDialogs.CommonDialog1.InitDir = DEFAULT_CHECKOUT_PATH
                        frmCommonDialogs.CommonDialog1.ShowOpen
                        If Err.Number = 0 Then
                            strfilepath = frmCommonDialogs.CommonDialog1.FileName
                            oAppl.Workbooks.Open (strfilepath)
                            oAppl.Visible = True
                        End If
                        GoTo Done
                    
                    Case APPL_POWERPOINT:
                        'Display File Open dialog and get back file selected.
                        'On Error Resume Next
                        'Err.Number = 0
                        frmCommonDialogs.OpenDialogSetup (APPL_POWERPOINT) 'initalize common dialog form
                        frmCommonDialogs.CommonDialog1.InitDir = DEFAULT_CHECKOUT_PATH
                        frmCommonDialogs.CommonDialog1.ShowOpen
                        If Err.Number = 0 Then
                            strfilepath = frmCommonDialogs.CommonDialog1.FileName
                            oAppl.Presentations.Open (strfilepath)
                            oAppl.Visible = True
                        End If
                
                End Select
                fileOpen = CIDMDriveSelection
            
            Case IDMObjects.idmOperationOpenTrackedFile
                'it will return a handle to a document, which you can use to
                'look up the file path in the tracked files object.
                strfilepath = goCmnDlg.FilePath
                Call fileOpenOffice(oAppl, iApplType, strfilepath)
                'fileOpen = CIDMOk
                
            Case Else
                fileOpen = CIDMCancel
                GoTo Done
        
        End Select
    Else
        If oDoc.TypeName = LoadResString(STR_EXTERNAL_DOC) Then ' "External Document"
           If oDoc.ShowPropertiesDialog = idmDialogExitOK Then
                If oDoc.GetState(idmDocModified) Then
                    oDoc.Save
                    'update mezz properites here
                End If
           End If
           fileOpen = CIDMCancel
           GoTo Done
        End If
        Select Case lOperation
            Case IDMObjects.idmOperationOpen
            
                ' This can be an open from an IMS Library, The actual file name will be mangled
                ' and don't need to show SaveAs dialog cause IS the GetCachedFile() call below
                ' will automatically put the file into the local cache
                If oDoc.GetState(idmDocCanSendCopy) = False Then
                    MsgBox LoadResString(MSG_CANOT_COPY_DOC), vbInformation, LoadResString(MSG_OPEN)
                    fileOpen = CIDMCancel
                    GoTo Done
                End If
                'check library type IS or DS
                If oDoc.Library.SystemType = idmSysTypeIS Then
                    ' special stuff if outlook is the caller and it is inserting an attachment
                    ' (but NOT a shortcut!)
                    If ((iApplType = APPL_OUTLOOK) And (CIDMInsert = (CIDMInsert And eflag))) Then
                        Set PageRange = New frmPageRange
                        If (oDoc.PageCount <= 1) Then
                            ' set both first and last page to 1
                            PageFirst = 1
                            PageLast = 1
                            fileOpen = CIDMOk
                        Else
                            ' ask user which pages of the IS document should be inserted as attachments
                            PageRange.FirstPage = 1
                            PageRange.LastPage = oDoc.PageCount
                            PageRange.Show vbModal
                            If (PageRange.CanceledFlag = False) Then
                                PageFirst = PageRange.FirstPage
                                PageLast = PageRange.LastPage
                                fileOpen = CIDMOk
                            Else
                                fileOpen = CIDMCancel
                            End If
                        End If
                        If (fileOpen = CIDMOk) Then
                            ' Note, we always return the first entry in sPathNames
                            ' equal to the original filename.  The rest of the entries
                            ' are the ones the caller asked for.  If the caller asked for
                            ' page 1 then it will be duplicated...
                            ReDim sPathNames(PageFirst - 1 To PageLast) As String
                            strfilepath = oDoc.GetCachedFile(1, , idmDocGetOriginalFileName)
                            sPathNames(PageFirst - 1) = strfilepath
                            ' get the requested page range.
                            For i = PageFirst To PageLast
                                sPathNames(i) = oDoc.GetCachedFile(i, , 0) 'idmDocGetOriginalFileName)
                            Next i
                            vPathNames = sPathNames
                        End If
                        Set PageRange = Nothing
                    Else
                        If (CIDMShortCut = (CIDMShortCut And eflag)) Then 'outlook insert shortcut
                            strfilepath = oDoc.CreateShortcut(strfilepath)
                        Else
                            strfilepath = oDoc.GetCachedFile(1, , idmDocGetOriginalFileName)
                            Call fileOpenOffice(oAppl, iApplType, strfilepath)
                        End If
                        fileOpen = CIDMOk
                    End If
                    GoTo Done
                End If
                'DS LIbrary, in these cases users will not see the openas group and the copy radio button.
                'these cases are word insert file, outlook open copy, insert shortcut and attachment.
                If (CIDMShortCut = (CIDMShortCut And eflag)) Then 'outlook insert shortcut
                    strfilepath = oDoc.CreateShortcut(strfilepath)
                    GoTo Done
                ElseIf (CIDMReference = (CIDMReference And eflag)) Then  'outlook reference
                     Set oProp = oDoc.GetExtendedProperty("idmVerFileName")
                     If oProp Is Nothing Then
                         Err.Raise 1, LoadResString(MSG_OPEN), LoadResString(MSG_CANNOT_GET_PROPERTY)
                     End If
                     strfilepath = strfilepath & "\" & oProp.Value & LoadResString(STR_FNI) '".fni"
                     oDoc.CreateReferenceFile (strfilepath)
                     
                Else  'word and outlook insert document and outlook copy cases
                    Set oProp = oDoc.GetExtendedProperty("idmVerFileName")
                    Set oVer = oDoc.Version
                    If oVer Is Nothing Then
                        Err.Raise 1, LoadResString(MSG_OPEN), LoadResString(MSG_CANNOT_GET_VERSION)
                    End If
                    If (CIDMInsert = (CIDMInsert And eflag)) And iApplType = APPL_OUTLOOK Then
                        'outlook insert document as attachment
                        strfilename = oProp.Value
                        strDirectory = strfilepath
                        strfilepath = FixDefaultDirectory(strfilepath) & "\" & oProp.Value
                    ElseIf (CIDMCopy = (CIDMCopy And eflag)) Then 'outlook open copy
                        ' [6/24/99 cmarsh] commented out the 3 lines below and
                        ' added the codelet that follows so that Outlook functionality
                        ' will display a saveas dialog during a "filenet open" operation
                        ' and properly handle the CIDMNoSaveAsDialog flag
                        
                        'strfilename = oProp.Value
                        'strDirectory = DEFAULT_COPY_PATH
                        'strfilepath = DEFAULT_COPY_PATH & "\" & oProp.Value
                        If (CIDMNoSaveAsDialog = (CIDMNoSaveAsDialog And eflag)) Then
                            strfilepath = FixDefaultDirectory(strfilepath) & "\" & oProp.Value
                        Else
                            strfilepath = DEFAULT_COPY_PATH & "\" & oProp.Value
                            If oDoc.GetState(idmDocHasChild) = False Then
                                If DocSaveLocalDialog(oAppl, strfilepath, iApplType, True) = vbCancel Then
                                    fileOpen = CIDMCancel
                                    GoTo Done
                                End If
                            End If
                        End If
    
                        idmGetDirectoryAndFileName strfilepath, strDirectory, strfilename
                    Else 'word insert case
                        If oDoc.GetState(idmDocHasChild) = True Then
                           MsgBox LoadResString(MSG_CAN_NOT_INSERT_COMP_DOC), vbInformation, LoadResString(MSG_OPEN)
                           fileOpen = CIDMCancel
                           Exit Function
                        End If
                        Call GetTempPath(20, sTempPath)
                        DEFAULT_INSERT_PATH = FixDefaultDirectory(Trim(sTempPath))
                        strfilename = oProp.Value
                        strfilepath = DEFAULT_INSERT_PATH & "\" & strfilename
                        strDirectory = DEFAULT_INSERT_PATH
                    End If
                    oVer.Copy strfilepath, strDirectory, strfilename
                End If
                
            Case IDMObjects.idmOperationOpenCopy
                 If oDoc.GetState(idmDocCanSendCopy) = False Then
                    MsgBox LoadResString(MSG_CANOT_COPY_DOC), vbInformation, LoadResString(MSG_OPEN)
                    fileOpen = CIDMCancel
                    GoTo Done
                 End If
                If (CIDMShortCut = (CIDMShortCut And eflag)) Then
                      Set oProp = oDoc.GetExtendedProperty("idmVerFileName")
                     If oProp Is Nothing Then
                         Err.Raise 1, LoadResString(MSG_OPEN), LoadResString(MSG_CANNOT_GET_PROPERTY)
                     End If
                     strfilepath = strfilepath
                     '& "\" & oProp.Value & ".lnk"
                     strfilepath = oDoc.CreateShortcut(strfilepath)
                ElseIf (CIDMReference = (CIDMReference And eflag)) Then
                     Set oProp = oDoc.GetExtendedProperty("idmVerFileName")
                     If oProp Is Nothing Then
                         Err.Raise 1, LoadResString(MSG_OPEN), LoadResString(MSG_CANNOT_GET_PROPERTY)
                     End If
                     strfilepath = strfilepath & "\" & oProp.Value & LoadResString(STR_FNI) '".fni"
                              
                     oDoc.CreateReferenceFile (strfilepath)
                Else  'copy case
                    Set oProp = oDoc.GetExtendedProperty("idmVerFileName")
                    If oProp Is Nothing Then
                        Err.Raise 1, LoadResString(MSG_OPEN), LoadResString(MSG_CANNOT_GET_PROPERTY)
                    End If
'                    If (CIDMInsert = (CIDMInsert And eflag)) And iApplType <> APPL_OUTLOOK Then
'                            Call GetTempPath(20, sTempPath)
'                            DEFAULT_COPY_PATH = FixDefaultDirectory(Trim(sTempPath))  'for show dialog case
'                            strfilepath = DEFAULT_COPY_PATH 'for no dialog case
'                    End If
                    
                    If (CIDMNoSaveAsDialog = (CIDMNoSaveAsDialog And eflag)) Then
                        strfilepath = FixDefaultDirectory(strfilepath) & "\" & oProp.Value
                    Else
                        strfilepath = DEFAULT_COPY_PATH & "\" & oProp.Value
                        If oDoc.GetState(idmDocHasChild) = False Then
                           If GetDirectory("ShowSaveAsUI") = 1 Then    'get preference value in Directories and Files
SaveDialog1:
                              If DocSaveLocalDialog(oAppl, strfilepath, iApplType, True) = vbCancel Then
                                 fileOpen = CIDMCancel
                                 GoTo Done
                              End If
                           Else
                                If Dir(strfilepath, vbNormal) <> "" Then
                                     If DocIsOpen(oAppl, iApplType, strfilepath, True) = True Then
                                     ' Check to see if document is allready open in application
                                           vbResult = MsgBox(strfilepath & LoadResString(MSG_FILE_SAVEAS_OVERWRITE), vbOKCancel + vbExclamation, LoadResString(STR_WARNING))
                                           If vbResult = vbCancel Then
                                               GoTo Done
                                           ElseIf vbResult = vbOK Then
                                               GoTo SaveDialog1
                                           End If
                                    Else
                                           vbResult = MsgBox(strfilepath & LoadResString(MSG_FILE_EXISTS_OVERWRITE), vbYesNoCancel + vbInformation, LoadResString(STR_WARNING))
                                           If vbResult = vbCancel Then
                                              GoTo Done
                                           ElseIf vbResult = vbNo Then
                                              GoTo SaveDialog1
                                           End If
                                    End If
                                End If
                           
                           
                           End If
                        End If
                    End If

                    idmGetDirectoryAndFileName strfilepath, strDirectory, strfilename
                    
                    If (oDoc.Library.GetState(idmLibrarySupportsCompoundDocuments) = True) And _
                       ((iApplType <> APPL_OUTLOOK) And (iApplType <> APPL_WORDPRO)) And (eflag <> 52) Then
                        Set oCom = oDoc.Compound
                        If oCom Is Nothing Then
                            Err.Raise 1, LoadResString(MSG_OPEN), LoadResString(MSG_CANNOT_GET_VERSION)
                        End If
                        If oCom.Copy(strfilepath, strDirectory, strfilename, idmCDCheckoutWithUI, Nothing) = False Then
                           fileOpen = CIDMCancel
                           GoTo Done
                        End If
                    Else
                        Set oVer = oDoc.Version
                        If oVer Is Nothing Then
                            Err.Raise 1, LoadResString(MSG_OPEN), LoadResString(MSG_CANNOT_GET_VERSION)
                        End If
                        oVer.Copy strfilepath, strDirectory, strfilename
                    End If
                    
                End If
                
            Case IDMObjects.idmOperationOpenCheckout
                
                If oDoc.GetState(idmDocCanCheckout) = False Then
                    fileOpen = CIDMError
                    If oErrorMgr Is Nothing Then
                        MsgBox LoadResString(MSG_ERROR_WITHOUT_ERRMGR), vbCritical, LoadResString(MSG_OPEN)
                        GoTo Done
                    Else
                        If oErrorMgr.Errors.Count > 0 Then
                            MsgBox oErrorMgr.Errors.Item(1).Description, vbExclamation, LoadResString(MSG_OPEN)
                        Else
                            MsgBox LoadResString(MSG_CANNOT_CHECKOUT) & strfilepath, vbExclamation, LoadResString(MSG_OPEN)
                        End If
                    End If
                    GoTo Done
                End If
                
                Set oProp = oDoc.GetExtendedProperty("idmVerFileName")
                If oProp Is Nothing Then
                    Err.Raise 1, LoadResString(MSG_OPEN), LoadResString(MSG_CANNOT_GET_PROPERTY)
                End If
                
                strfilepath = DEFAULT_CHECKOUT_PATH & "\" & oProp.Value
                If oDoc.GetState(idmDocHasChild) = False Then
                   '08/02/01 block out the if ...end for a *SS FR#29795 according to Kristin , always let users see  a UI and worning if the doc is in the loca drive.
                   'the preference(Prompt for filename on check out or copy) will become useless
                   If GetDirectory("ShowSaveAsUI") = 1 Then 'get preference value in Directories and Files
SaveDialog2:
                      If DocSaveLocalDialog(oAppl, strfilepath, iApplType, True) = vbCancel Then
                            fileOpen = CIDMCancel
                            GoTo Done
                      End If
                   Else
                      If Dir(strfilepath, vbNormal) <> "" Then
                          If DocIsOpen(oAppl, iApplType, strfilepath, True) = True Then
                          ' Check to see if document is allready open in application
                                vbResult = MsgBox(strfilepath & LoadResString(MSG_FILE_SAVEAS_OVERWRITE), vbOKCancel + vbExclamation, LoadResString(STR_WARNING))
                                If vbResult = vbCancel Then
                                    GoTo Done
                                ElseIf vbResult = vbOK Then
                                    GoTo SaveDialog2
                                End If
                         Else
                                vbResult = MsgBox(strfilepath & LoadResString(MSG_FILE_EXISTS_OVERWRITE), vbYesNoCancel + vbInformation, LoadResString(STR_WARNING))
                                If vbResult = vbCancel Then
                                   GoTo Done
                                ElseIf vbResult = vbNo Then
                                   GoTo SaveDialog2
                                End If
                         End If
                      End If
                      
                   End If
                End If
                Call idmGetDirectoryAndFileName(strfilepath, strDirectory, strfilename)
                
                If (oDoc.Library.GetState(idmLibrarySupportsCompoundDocuments) = True) Then
                    Set oCom = oDoc.Compound
                    If oCom Is Nothing Then
                        Err.Raise 1, LoadResString(MSG_OPEN), LoadResString(MSG_CANNOT_GET_VERSION)
                    End If
                    If oCom.Checkout(strfilepath, strDirectory, strfilename, idmCDCheckoutWithUI, Nothing) = False Then
                       fileOpen = CIDMCancel
                       GoTo Done
                    End If
                Else
                    Set oVer = oDoc.Version
                    If oVer Is Nothing Then
                        Err.Raise 1, LoadResString(MSG_OPEN), LoadResString(MSG_CANNOT_GET_VERSION)
                    End If
                    oVer.Checkout strfilepath, strDirectory, strfilename
                End If
                Call fileOpenOffice(oAppl, iApplType, strfilepath)
                Call UpdateMezzProperties(oAppl, iApplType, LoadResString(STR_UPDATE_PROP_AFTER_CHECKOUT))
                GoTo Done

            Case IDMObjects.idmOperationOpenView
                    If oDoc.GetState(idmDocHasChild) = False Then
                       strfilepath = oDoc.GetCachedFile(1, , idmDocGetOriginalFileName)
                    Else
                       strfilepath = oDoc.Compound.GetCachedFile(idmCDGetChildrenWithUI)
                    End If
                    If (iApplType <> APPL_OUTLOOK) Then
                        Call fileOpenOffice(oAppl, iApplType, strfilepath)
                    End If
                    fileOpen = CIDMOk
                    GoTo Done
           
            Case Else
                GoTo Done
            
        End Select
        If (CIDMInsert = (CIDMInsert And eflag)) Then
            If (iApplType = APPL_WORD) Then
                oAppl.Selection.InsertFile strfilepath
                Kill strfilepath
            End If
        Else
            Call fileOpenOffice(oAppl, iApplType, strfilepath)
        End If
        fileOpen = CIDMOk
    End If
    
    GoTo Done
 
errHandler:
    fileOpen = CIDMError
    If Err.Number <> 0 Then
        If Err.Number = 4198 Then    ' Word - "Command failed"
            GoTo Done
        End If
        If Err.Number = 1004 Then    ' Excel - "Open method of Workbooks class failed"
            ' Microsoft used the same error number for two different error conditions.
            ' We need to filter out one of those conditions and display a message for the other.
            If Err.Description = LoadResString(MSG_OPEN_METHOD_FAILED) Then
                GoTo Done
            End If
        End If
        MsgBox Err.Description, vbCritical, LoadResString(MSG_OPEN)
        If (Err.Number = -2147467259 And GetDocStatus(strfilepath) = DocCheckedout) Then
            If MsgBox(LoadResString(MSG_WOULD_YOU_LIKE_TO_CANCEL_CHECKOUT), vbYesNo + vbInformation, LoadResString(MSG_OPEN)) = vbYes Then
                If (oDoc.Library.GetState(idmLibrarySupportsCompoundDocuments) = True) Then
                    oCom.CancelCheckout idmCDCancelCheckoutWithUI, Nothing
                Else
                    oVer.CancelCheckout idmCancelCheckoutKeep
                End If
                If Dir(strfilepath) <> "" Then
                   Kill strfilepath
                End If
            End If
        End If
    Else
        If oErrorMgr Is Nothing Then
            MsgBox LoadResString(MSG_ERROR_WITHOUT_ERRMGR), vbCritical, LoadResString(MSG_OPEN)
            GoTo Done
        Else
            If oErrorMgr.Errors.Count > 0 Then
                oErrorMgr.ShowErrorDialog
            End If
        End If
    End If
    
Done:
    Set oVer = Nothing
    Set oProp = Nothing
    Set oDoc = Nothing
    Set oCom = Nothing
    Set oErrorMgr = Nothing
    Set oRetObject = Nothing
End Function
