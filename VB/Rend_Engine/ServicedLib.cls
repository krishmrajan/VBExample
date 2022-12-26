VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ServicedLib"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' This program is an example which uses Publishing foundation objects
'----------------------------------------------------------------------
'   Disclaimer: FileNET provides programming examples for illustration
'   only, without warranty either expressed or implied,  including, but
'   not limited to, the implied warranties of merchantability and/or
'   fitness for a particular purpose. This sample assumes that you are
'   familiar with the programming language being demonstrated and the
'   tools used to create and debug procedures.
'----------------------------------------------------------------------

' Revision:   1.1
' Date:       November 19, 1999 12:35:54
' Author:     Vladimir Fridman
' Workfile:   ServicedLib.cls

'----------------------------------------------------------------------
' All Rights Reserved.  Copyright (c) 1999 FileNet Corp.
'----------------------------------------------------------------------

Option Explicit

'settings
Public sLibraryName As String
Public sLoginID As String
Public sPassword As String

Public bVersionChildren As Boolean
Public bFileWithParent As Boolean
Public nRequestsEachTime As Integer


'private variables
Private oRenditionEngine As RenditionEngine
Private oLibrary As Library

Private oWordApp As Word.Application, oWordDoc As Word.Document
Private oExcelApp As Excel.Application, oExcelDoc As Excel.Workbook
Private oPowerPointApp As PowerPoint.Application, oPPDoc As PowerPoint.Presentation


'statistics variables
Public nQueries As Long
Public nDocsPublishedOK As Long
Public nPreventedCollisions As Long
Public nDocumentsAdded As Long
Public nVersionsAdded As Long




Private Function SetPEStatus(oPE As PublishedElement, nStatus As idmPublishedElementStatus, Optional nError As idmPublishedElementErrorCategory, Optional sError As String, Optional sReturnedErrorMessage As String) As Boolean
    On Error GoTo ErrorHandler
    oPE.SetStatus nStatus, nError, sError
    sReturnedErrorMessage = ""
    SetPEStatus = True
Exit Function
ErrorHandler:
    sReturnedErrorMessage = Err.Description
    SetPEStatus = False
End Function

Public Property Get StyleTemplates() As ObjectSet
    If Not oRenditionEngine Is Nothing Then
        Set StyleTemplates = oRenditionEngine.StyleTemplates
    Else
        Set StyleTemplates = New ObjectSet
    End If
End Property

Sub AddNewStyleTemplate()
    Dim oNewTemplate As StyleTemplate
    
    Dim frmDlg As New frmStyleTemplate
    Set frmDlg.oStyleTemplate = oNewTemplate
    Set frmDlg.oRenditionEngine = oRenditionEngine
    frmDlg.bAddingMode = True
    frmDlg.Show vbModal

End Sub


Public Property Get IsWorkingOK() As Boolean
    If Not oLibrary Is Nothing Then
        If oLibrary.GetState(idmLibraryLoggedOn) Then
            If Not oRenditionEngine Is Nothing Then
                IsWorkingOK = True
            End If
        End If
    End If
End Property

Private Function GetFirstUnPublishedElement() As PublishedElement
    On Error GoTo ErrorHandler
    Dim oPE As PublishedElement, oQueue As ObjectSet
    Set oQueue = oLibrary.FilterUnPublishedElements(idmLibraryPEInQueue, REND_ENG_ID)
    nQueries = nQueries + 1
    If oQueue.Count > 0 Then
        Set oPE = oQueue.Item(1)
    End If
    Set GetFirstUnPublishedElement = oPE
Exit Function
ErrorHandler:
    GetFirstUnPublishedElement = Nothing
End Function

Public Property Get RequestsInQueue() As Long
    On Error GoTo ErrorHandler
    RequestsInQueue = oLibrary.FilterUnPublishedElements(idmLibraryPEInQueue, REND_ENG_ID).Count
Exit Property
ErrorHandler:
    RequestsInQueue = 0
End Property




Function ShowPropertiesDialog(Optional bAddLibraryMode As Boolean = False) As idmDialogExit
    
    Dim oPropDlg As frmLibrarySettings
    Set oPropDlg = New frmLibrarySettings
    Set oPropDlg.oServicedLibrary = Me
    oPropDlg.bAddLibraryMode = bAddLibraryMode
    oPropDlg.Show vbModal
    
    
    If oPropDlg.bPressedOK Then
        ShowPropertiesDialog = idmDialogExitOK
    Else
        ShowPropertiesDialog = idmDialogExitCancel
    End If

End Function

Sub WriteSettings(nFileNumber As Integer)
'Write settings to the file
    
    Print #nFileNumber, "=================================="
    Print #nFileNumber, sLibraryName
    Print #nFileNumber, sLoginID
    Print #nFileNumber, sPassword
    Print #nFileNumber, nRequestsEachTime
    Print #nFileNumber, bFileWithParent
    Print #nFileNumber, bVersionChildren
End Sub

Sub ReadSettings(nFileNumber As Integer)
'Read settings from the file
    
    Dim sTemp As String
    
    Input #nFileNumber, sTemp 'empty line
    Input #nFileNumber, sLibraryName
    Input #nFileNumber, sLoginID
    Input #nFileNumber, sPassword
    Input #nFileNumber, nRequestsEachTime
    
    Input #nFileNumber, sTemp
    bFileWithParent = CBool(sTemp)
    
    Input #nFileNumber, sTemp
    bVersionChildren = CBool(sTemp)
    
    InitializeRenditions
End Sub


Sub CheckLibraryQueue()
'Checks if there are requests in the Publishing queue
'gets the first element, tries to set it to "InWork"
'if successfull, processes theat published element

    On Error GoTo ErrorHandler
    Dim oWorkPE As PublishedElement, i As Integer
    
    For i = 1 To nRequestsEachTime
        Set oWorkPE = GetFirstUnPublishedElement()
        If Not oWorkPE Is Nothing Then
        
            'even though this is the first element in the queue, another
            'rendition engine could have already changed the status to "InWork"
            'it did, PublishedElement.Setstatus would fail, and
            'SetPEStatus would return False,  and we'll skip this Published Element
        
            If SetPEStatus(oWorkPE, idmPEInWork) Then

                bCantQuitNow = True
                ProcessPublishedElement oWorkPE
                bCantQuitNow = False
            Else
                nPreventedCollisions = nPreventedCollisions + 1
            End If
        Else
            
        End If
    Next i

Exit Sub
ErrorHandler:
    ShowError "CheckLibraryQueue"
End Sub

Sub InitializeRenditions()
'will logon and find or create oRenditionEngine in the library
    On Error GoTo ErrorHandler:
    
    Set oLibrary = oHood.Libraries(sLibraryName)
    If Not oLibrary.GetState(idmLibraryLoggedOn) Then
        If oLibrary.Logon(sLoginID, sPassword, , idmLogonOptNoUI) Then
            Set oRenditionEngine = FindEngine(REND_ENG_ID)
        Else
            MsgBox "Error loggin into " + sLibraryName, vbCritical
        End If
    Else
        Set oRenditionEngine = FindEngine(REND_ENG_ID)
    End If
    
    
Exit Sub
ErrorHandler:
    ShowError "Initializing " + sLibraryName, True
End Sub





Sub RemoveRenditionServices()
'Delete all style templates and the Rendition Engine object from the library

    Dim oST As StyleTemplate
    On Error GoTo ErrorHandler
    For Each oST In oRenditionEngine.StyleTemplates
        oST.Delete
    Next
    oRenditionEngine.Delete
    Set oRenditionEngine = Nothing
Exit Sub
ErrorHandler:
    ShowError "RemoveRenditionServices", True
End Sub


Private Function FindEngine(sEngineID As String) As RenditionEngine
    On Error GoTo ErrorHandler
    Set FindEngine = oLibrary.GetObject(idmObjTypeRenditionEngine, sEngineID)
Exit Function
ErrorHandler:
    Set FindEngine = CreateEngine()
End Function


Private Function CreateEngine() As RenditionEngine
'Creates RenditionEngine object

    On Error GoTo ErrorHandler
    Dim oTempEngine As RenditionEngine
    Set oTempEngine = oLibrary.CreateObject(idmObjTypeRenditionEngine, REND_ENG_ID)
    oTempEngine.Name = REND_ENG_NAME
    oTempEngine.SupportsCD = True
    oTempEngine.SupportsCrossLibraryRequest = False
    oTempEngine.Save
    Set CreateEngine = oTempEngine
Exit Function
ErrorHandler:
    MsgBox Err.Description, , "Could not create rendition engine"
    Set CreateEngine = Nothing
End Function


Private Sub ProcessPublishedElement(oWorkPE As PublishedElement)
'Renders and adds new published document
'oWorkPE.status should be "inWork"
    
    On Error GoTo ErrorHandler

    Dim sTargetFile As String, sFullSourcePath As String, oStyleTemplate As StyleTemplate, _
        sErrDescr As String, oSourceDoc As IDMObjects.Document, sSetStatusError As String
        
    AddMessage oLibrary.Name + " - Starting job on document " + oWorkPE.SourceDocID
    StatusBarText "Converting document..."
    
    ClearTemporaryFolders
    
    Set oSourceDoc = oWorkPE.SourceDocument
    oSourceDoc.TrackInLocalDb = False
    
    Set oStyleTemplate = oLibrary.GetObject(idmObjTypeStyleTemplate, oWorkPE.StyleTemplateID)
    
    oSourceDoc.Version.Copy sFullSourcePath, sSourceDir

    'if rendering succedes and
    'if adding new version to published document succedes
    'set status
    'if rendering fails or adding contents fails, set status to InError
    
    If RenderDocument(sFullSourcePath, sTargetFile, oStyleTemplate, sErrDescr) Then
        If AddPublishedDocumentContents(oWorkPE.PublishedDocument, sTargetFile) Then
            If SetPEStatus(oWorkPE, idmPEPublished, , , sSetStatusError) Then
                AddMessage oLibrary.Name + " - Job is complete.  ID is " + oWorkPE.PublishedDocID
                AddMessage
                nDocsPublishedOK = nDocsPublishedOK + 1
            Else
                AddMessage "Error in PublishedElement.SetStatus: " + sSetStatusError
                AddMessage "Setting status to 'Published' failed.  Will try to set status to 'In Error'"
                FailJob oWorkPE, idmPEErrorRenderingFailed, "Could not set status to 'Published'"
            End If
        Else
            FailJob oWorkPE, idmPEErrorCannotAddTargRend, "Error when adding contents to the published document"
        End If
    Else
        FailJob oWorkPE, idmPEErrorRenderingFailed, sErrDescr
    End If
    
    ClearTemporaryFolders 'delete source and published files
    StatusBarText ""
    
Exit Sub
ErrorHandler:
    ShowError "ProcessPublishedElement", False
    oWorkPE.SetStatus idmPEInError, idmPEErrorRenderingFailed, Err.Description
    ClearTemporaryFolders
End Sub

Sub FailJob(oWorkPE As PublishedElement, nErrorCode As idmPublishedElementErrorCategory, sErrDescr As String)
    Dim sSetStatusError As String
    If SetPEStatus(oWorkPE, idmPEInError, nErrorCode, sErrDescr, sSetStatusError) Then
        AddMessage "Status was set to 'In Error'.  Job failed - " + sErrDescr
        AddMessage
    Else
        AddMessage "Error in PublishedElement.SetStatus: " + sSetStatusError
        AddMessage "Could not set status to 'In Error'.  It is now stuck in 'In Work'.  Contact your Administrator to delete this publish request"
        AddMessage
    End If
End Sub


Function GetChildDocuments(oParentDoc As Document) As ObjectSet
'finds all available children for the parent document, to be versioned if needed

    On Error GoTo ErrorHandler
    Dim oChildDocs As ObjectSet, oExistingLinks As Links, i As Integer
    Set oChildDocs = New ObjectSet

    Set oExistingLinks = oParentDoc.Compound.GetChildren(0)
    
    
    'got a collection of children
    For i = 1 To oExistingLinks.Count
        'this will fail if a child doc is deleted
        'in that case it will just go to the next child
        oChildDocs.Add oExistingLinks.Item(i).Child
    Next i
    
    Set GetChildDocuments = oChildDocs
    
Exit Function
ErrorHandler:
    Resume Next
End Function


Private Function AddPublishedDocumentContents(oPubDoc As Document, sTargetFile As String) As Boolean
'Same as AddPublishedDocumentContents, but versions existing children

    On Error GoTo ErrorHandler:
    
    Dim oChildDocs As ObjectSet
    Dim oDocChild As Document, i As Integer
    Dim oFolders As IDMObjects.ObjectSet
    Dim oChildFiles As Collection, sChildPath As String, sChildRelativePath As String
    Dim nNewDocsNeeded As Integer, nChildrenToVersion As Integer
    
    If bFileWithParent Then
        Set oFolders = oPubDoc.FoldersFiledIn
    End If
    
    oPubDoc.Version.CheckoutNoCopy
    
    'got a collection of children for the published document
    Set oChildDocs = IIf(bVersionChildren, GetChildDocuments(oPubDoc), New ObjectSet)
    
    
    
    oPubDoc.Compound.Children.Clear 'clear existing links
    
    'Files we need to add, either by versioning existing children or by adding new children
    Set oChildFiles = FindAllFiles(sTargetFile)
    
    
    nNewDocsNeeded = Abs(oChildFiles.Count - oChildDocs.Count)
    
    If nNewDocsNeeded = 0 Then 'if there are no new documents needed, ie we got enough children
        nChildrenToVersion = oChildFiles.Count 'then the number of children we need to version is the number of files we have
    Else 'if not, then version all of the children documents
        nChildrenToVersion = oChildDocs.Count
    End If
    
    
    'checkout and check in the new version of existing children
    For i = 1 To nChildrenToVersion
        sChildPath = oChildFiles(i)
        sChildRelativePath = GetRelativePath(sTargetFile, sChildPath)
        Set oDocChild = oChildDocs(i)
        AddMessage "Versioning child '" + CStr(oDocChild.Name) + "', " + CStr(i) + " of " + CStr(oChildFiles.Count) + " - " + sChildRelativePath
        
        VersionChild oPubDoc, oDocChild, sChildPath, sChildRelativePath, oFolders
    Next i
    
    
    'add new children for all nNewDocsNeeded
    For i = nChildrenToVersion + 1 To nNewDocsNeeded + nChildrenToVersion
        sChildPath = oChildFiles(i)
        sChildRelativePath = GetRelativePath(sTargetFile, sChildPath)
        AddMessage "Adding child " + CStr(i) + " of " + CStr(oChildFiles.Count) + " - " + sChildRelativePath
        
        AddNewChild oPubDoc, sChildPath, sChildRelativePath, oFolders
    Next

    CheckinParent oPubDoc, sTargetFile
    
    AddPublishedDocumentContents = True
        
Exit Function
ErrorHandler:
    ShowError "Adding Contents to Published Document"
    AddPublishedDocumentContents = False

End Function



Private Sub FileInFolders(oDocument As IDMObjects.Document, oFolders As ObjectSet)
'File oDocument in every folder in oFolders object set

    On Error GoTo ErrorHandler:
    Dim oFolder As Folder
    For Each oFolder In oFolders
        oFolder.File oDocument
    Next
Exit Sub
ErrorHandler:
    'nothing worth doing here
End Sub

Private Sub CopyRequiredProperties(oSourceDoc As IDMObjects.Document, oTargetDoc As IDMObjects.Document)
'this subprocedure copies all required properties from the source document
'to the target.  Source and target documents MUST have the same doc class..
    On Error GoTo ErrorHandler
    Dim i As Integer, oProperty As IDMObjects.Property
    
    For i = 1 To oSourceDoc.Properties.Count
        Set oProperty = oSourceDoc.Properties(i)
        
        If oProperty.PropertyDescription.GetState(idmPropRequired) And _
           oProperty.PropertyDescription.GetState(idmPropReadOnly) = False And _
           oProperty.TypeID <> idmTypeNull Then
                
                oTargetDoc.Properties(i).Value = oProperty.Value
        
        End If
    Next
Exit Sub
ErrorHandler:
    ShowError "CopyRequiredProperties"
End Sub


Private Sub CheckinParent(oPubDoc As IDMObjects.Document, sTargetFile)
'Checks in the parent document

    oPubDoc.Compound.BehaviorID = "IDMObjects.OLEBehavior"
    AddMessage "Adding main document - .\" + oFS.GetFileName(sTargetFile)
    oPubDoc.Version.Checkin oFS.GetParentFolderName(sTargetFile), oFS.GetFileName(sTargetFile), idmCheckinNoKeep
    
    nVersionsAdded = nVersionsAdded + 1
    
End Sub
Private Sub LinkDocuments(oDocParent As IDMObjects.Document, oDocChild As IDMObjects.Document, sChildPath As String, sChildRelativePath As String)
'Creates a link between Parent and Child documents

    On Error GoTo ErrorHandler
    Dim oLink As Link
    Set oLink = New Link
    
    oLink.Parent = oDocParent
    oLink.Child = oDocChild
    
    oLink.ClassID = idmDynamicLink
    oLink.Properties("idmLinkUsesRelativePath") = idmFileLinkRelative
    oLink.Properties("idmLinkAbsolutePath") = sChildPath
    oLink.Properties("idmLinkRelativePath") = sChildRelativePath
    oLink.Properties("idmLinkStrength") = idmStrongLink

Exit Sub
ErrorHandler:
    ShowError "LinkDocuments"

End Sub

Private Sub AddNewChild(oDocParent As IDMObjects.Document, sChildPath As String, sChildRelativePath As String, oFoldersToFileIn As IDMObjects.ObjectSet)
'Adds new child document

    On Error GoTo ErrorHandler
    Dim oDocChild As IDMObjects.Document
    
    Set oDocChild = oLibrary.CreateObject(idmObjTypeDocument, oDocParent.ClassDescription.Name)
    CopyRequiredProperties oDocParent, oDocChild
    
    oDocChild.Properties("idmName") = CStr(oDocParent.Version.Series.Count + 1) + "_" + oFS.GetBaseName(sChildPath)
    oDocChild.Compound.BehaviorID = "IDMObjects.OLEBehavior"
    
    LinkDocuments oDocParent, oDocChild, sChildPath, sChildRelativePath
    
    oDocChild.SaveNew sChildPath
    
    FileInFolders oDocChild, oFoldersToFileIn
    
    nDocumentsAdded = nDocumentsAdded + 1
    
Exit Sub
ErrorHandler:
    ShowError "AddNewChild"
    Resume Next
End Sub


Private Sub VersionChild(oDocParent As IDMObjects.Document, oDocChild As IDMObjects.Document, sChildPath As String, sChildRelativePath As String, oFoldersToFileIn As IDMObjects.ObjectSet)
'Version a child of a published document
    
    On Error GoTo ErrorHandler

    oDocChild.Version.CheckoutNoCopy '<-- checkout child document
    oDocChild.Properties("idmName") = CStr(oDocParent.Version.Series.Count + 1) + "_" + oFS.GetBaseName(sChildPath)
    oDocChild.Compound.BehaviorID = "IDMObjects.OLEBehavior"
    
    LinkDocuments oDocParent, oDocChild, sChildPath, sChildRelativePath
    
    'checkin child
    Dim sChildDir As String, sChildFileName As String
    sChildDir = oFS.GetParentFolderName(sChildPath)
    sChildFileName = oFS.GetFileName(sChildPath)
    
    oDocChild.Version.Checkin sChildDir, sChildFileName, idmCheckinNoKeep
    
    FileInFolders oDocChild, oFoldersToFileIn

    nVersionsAdded = nVersionsAdded + 1
    
Exit Sub
ErrorHandler:
    ShowError "VersionChild"
End Sub
Public Property Get UserStore() As String
'Returns the value of UserStore property of RenditionEngine object.  This is not used in this application
    UserStore = oRenditionEngine.UserStore
End Property

Private Sub Class_Initialize()
'Prints the time in VB Debug window
    Debug.Print "ServicedLib constructor " + CStr(Now())
End Sub

Private Sub Class_Terminate()
'Logs of the library, prints the time and library name in VB Debug window
    If Not oLibrary Is Nothing Then
        oLibrary.Logoff
    End If
    Debug.Print "ServicedLib destructor - Lib " + sLibraryName + " " + CStr(Now())
End Sub

Private Function RenderDocument(sSourcePath As String, sTargetFile As String, oStyleTemplate As StyleTemplate, sErrorDescription As String) As Boolean
'This is the function that actually converts files into HTML
'This function uses Office 2000
'Will not work with Office 97 because of differences in .SaveAs method
'Note: This should handle password protected documents (by failing the request)
    
    On Error GoTo ErrorHandler
    Dim sExtension As String
    
    sExtension = oFS.GetExtensionName(sSourcePath)
    sTargetFile = sTargetDir + oFS.GetBaseName(sSourcePath) + ".html"
    
    If sExtension = "xls" Then 'use excel
    
        Set oExcelApp = New Excel.Application
        oExcelApp.DisplayAlerts = False
        Set oExcelDoc = oExcelApp.Workbooks.Open(sSourcePath, False, True, , "password")
        FormatExcelDoc oExcelDoc, oStyleTemplate
        oExcelDoc.SaveAs sTargetFile, xlHtml
        oExcelDoc.Close
    
    ElseIf sExtension = "ppt" Then 'use power point
        
        Set oPowerPointApp = New PowerPoint.Application
        Set oPPDoc = oPowerPointApp.Presentations.Open(sSourcePath, msoTrue, msoTrue, msoFalse)
        FormatPPDoc oPPDoc, oStyleTemplate
        oPPDoc.SaveAs sTargetFile, ppSaveAsHTML
        oPPDoc.Close
    
    Else 'use Word for everything else
        
        Set oWordApp = New Word.Application
        oWordApp.DisplayAlerts = wdAlertsNone
        Set oWordDoc = oWordApp.Documents.Open(sSourcePath, False, True, False, _
           "password", , , , , wdOpenFormatAuto)
        FormatWordDoc oWordDoc, oStyleTemplate
        oWordDoc.SaveAs sTargetFile, wdFormatHTML
        oWordDoc.Close
        
    End If
    
    CloseRunningApps
    
    RenderDocument = True
       
Exit Function
ErrorHandler:
    sErrorDescription = Err.Description
    RenderDocument = False
    CloseRunningApps
End Function


Sub FormatExcelDoc(oWorkBook As Excel.Workbook, oST As StyleTemplate)
'changes background to color specified in StyleTemplate.Definition
    On Error GoTo ErrorHandler
    Dim oSheet As Excel.Worksheet, lColor As Long
    
    If IsNumeric(oST.TemplateDefinition) Then
        lColor = CLng(oST.TemplateDefinition)
        For Each oSheet In oWorkBook.Sheets
            oSheet.Cells.Interior.Color = lColor
        Next
    End If
    
Exit Sub
ErrorHandler:
    ShowError "FormatExcelDoc", False
End Sub

Sub FormatWordDoc(oDocument As Word.Document, oST As StyleTemplate)
'changes background to color specified in StyleTemplate.Definition
    If IsNumeric(oST.TemplateDefinition) Then
        oDocument.Background.Fill.ForeColor.RGB = CLng(oST.TemplateDefinition)
        oDocument.Background.Fill.Visible = msoTrue
        oDocument.Background.Fill.Solid
    End If
End Sub


Sub FormatPPDoc(oPresentation As PowerPoint.Presentation, oST As StyleTemplate)
'This function TRIES to change background to color specified in StyleTemplate.Definition
'It does not work with all presentations, especially with pictures as the background
    
    If IsNumeric(oST.TemplateDefinition) Then
        Dim lColor As Long
        lColor = CLng(oST.TemplateDefinition)
        
        If oPresentation.HasTitleMaster Then
            With oPresentation.TitleMaster.Background
                .Fill.Visible = msoTrue
                .Fill.ForeColor.RGB = lColor
                .Fill.Transparency = 0#
                .Fill.Solid
            End With
        End If
        
        With oPresentation.SlideMaster.Background
            .Fill.Visible = msoTrue
            .Fill.ForeColor.RGB = lColor
            .Fill.Transparency = 0#
            .Fill.Solid
        End With
        
        With oPresentation.Slides.Range
            .FollowMasterBackground = msoTrue
            .DisplayMasterShapes = msoTrue
        End With
        
    End If
End Sub



Sub CloseRunningApps()
'Close any of the office apps that are running...

    If Not oWordApp Is Nothing Then
        oWordApp.Quit
        Set oWordApp = Nothing
    End If
    If Not oExcelApp Is Nothing Then
        oExcelApp.Quit
        Set oExcelApp = Nothing
    End If
    If Not oPowerPointApp Is Nothing Then
        oPowerPointApp.Quit
        Set oPowerPointApp = Nothing
    End If
End Sub
