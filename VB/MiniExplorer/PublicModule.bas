Attribute VB_Name = "PublicModule"
Option Explicit

'NOTE:  The next 2 constants should only be set for FileNET IDMIS demo systems
'       A real user should not have SysAdmin privileges
Public Const GS_SYSADMIN_LOGON_NAME = "SysAdmin"
Public Const GS_SYSADMIN_PASSWORD = "SysAdmin"
'NOTE:  The next 2 constants should only be set for FileNET IDMDS demo systems
'       A real user should not have admin privileges
Public Const GS_ADMIN_LOGON_NAME = "admin"
Public Const GS_ADMIN_PASSWORD = ""
Public Const GS_ADMIN_GROUP = ""
'Boolean to track whether or not to use SysAdmin/admin for logon
Public gbSettingUseSysAdminLogon As Boolean
'Need to set this to some Document class on IDMIS for drag-and-drop
'file commit wizard to work correctly.
Public Const GS_DOC_CLASS = "general"

'Global FileNET objects
Public goNeighborhood As New IDMObjects.Neighborhood
Public goDefaultLibrary As IDMObjects.Library
Public goLibrary As IDMObjects.Library
Public goLastLoggedOnLibrary As IDMObjects.Library

'Global variable used to fill a MsgBox
Public gsMsg As String
'Global variable for trapping error codes not handled by VB
Public glErrorCode As Long

Public gbSuccess As Boolean
Public gbLoggedOn As Boolean
Public gsUserName As String
Public gsPassword As String
Public gbFolderOnClipboard As Boolean

'Global constants for custom error messages in DisplayErrorMessage procedure.
'Define in .RC file
Public Const GI_ERR_FATAL_ERROR = 32000
Public Const GI_ERR_ERROR_HAS_OCCURRED = 32001
Public Const GI_ERR_IN_THE_ROUTINE = 32002
Public Const GI_ERR_IN_THE_PROGRAM = 32003
Public Const GI_ERR_VERSION = 32004
Public Const GI_ERR_THE_ERROR_MESSAGE_IS = 32005
Public Const GI_ERR_ERROR = 32006
Public Const GI_ERR_CUSTOM_ERROR_MESSAGE = 32007
Public Const GI_ERR_TERMINATING_APPLICATION = 32008
Public Const GI_ERR_LOGON_UNSUCCESSFUL = 32010
Public Const GI_ERR_UNABLE_TO_ADD_TO_TREEVIEW = 32020
Public Const GI_ERR_UNABLE_TO_FILL_LISTVIEW = 32030
Public Const GI_ERR_LOAD_GLOBAL_SETTINGS = 32040
Public Const GI_ERR_SAVE_GLOBAL_SETTINGS = 32041
Public Const GI_ERR_UNABLE_TO_DELETE = 32042
Public Const GI_ERR_UNABLE_TO_COPY = 32043
Public Const GI_ERR_UNABLE_TO_RENAME = 32044

'If an error occurs trying to display an error, these constants will be used
'in the error message.  Typically this is caused by a resource string for one
'of the above errors not being specified and/or available.
Public Const GS_ERR_ERROR = "Error"
Public Const GS_ERR_ERROR_ON_ERROR_DISPLAY = "An error has occurred while trying to display a previous error."

'Global constant for indenting
Public Const GS_TAB = "     "

Public Const GI_DOCUMENT = 31000
Public Const GI_DOCUMENTS = 31001
Public Const GI_FOLDER = 31002
Public Const GI_FOLDERS = 31003
Public Const GI_CONTENTS_OF = 31004
Public Const GI_CONFIRM = 31005
Public Const GI_CONFIRM_FDELETE = 31006
Public Const GI_FOLDER_NAME = 31007
Public Const GI_CONFIRM_DDELETE = 31008

Public Const LISTVIEW_BUTTON = 11

'For tree view/list view
Public goCurSelTreeItem As Object
Public giCurSelTreeObjType As IDMTreeView.idmObjectType
Public goCurSelListItem As Object
Public giCurSelListObjType As IDMListView.idmObjectType

'For status bar counts
Public giFolderCount As Integer
Public giDocumentCount As Integer

Public gcLibraryCollection As New Collection


Sub Main()
        
On Error GoTo ErrorHandler

    gbSuccess = False
    
    'Call AppInitialize to connect all FileNET and any other global settings
    'needed for this application.
    gbSuccess = AppInitialize
    
    'Test gbSuccess for success/fail
    If gbSuccess = False Then  'there was an error
        'If AppInitialize fails, then the connections to some pieces
        'of the FileNET system(s) are invalid.  This would be considered
        'a "Fatal" error in initalizing the application.
        
        'Setup error messsage text
        gsMsg = LoadResString(GI_ERR_FATAL_ERROR) & " " & LoadResString(GI_ERR_TERMINATING_APPLICATION)
        'Display error
        MsgBox gsMsg, vbCritical, LoadResString(GI_ERR_ERROR)
        
        'Call AppTerminate cleanup routine to try to free resources,
        'terminate, and shutdown/end the program without saving settings
        AppTerminate False
    End If
    
    frmLogon.Show vbModal
    
    frmMain.Show
   
    Exit Sub
    
ErrorHandler:
    
    'Display Error Message - pass the name of this subroutine/function
    DisplayErrorMessage ("PublicModule - Main")
    
    'Cleanup Error values
    CleanupErrorCodes

    Resume Next
        
End Sub


Sub LoadResStrings(frm As Form)

On Error Resume Next
    
Dim lcControl As Control
Dim loObject As Object
Dim loFont As Object
Dim lsControlType As String
Dim liValue As Integer

    'Set the form's caption
    liValue = 0
    liValue = Val(frm.Tag)
    If liValue > 0 Then
        frm.Caption = LoadResString(CInt(liValue))
    End If
    
    'Set the font - you must specify the fonts in the .RC file
    'to correspond with these numbered entries here.
    Set loFont = frm.Font
    loFont.Name = LoadResString(20)
    loFont.Size = CInt(LoadResString(21))
    
    'Set the controls' captions using the caption
    'property for menu items and the Tag property
    'for all other controls.
    For Each lcControl In frm.Controls
        lsControlType = TypeName(lcControl)
        If lsControlType = "Label" Then
            Set lcControl.Font = loFont
            lcControl.Caption = LoadResString(CInt(lcControl.Tag))
        ElseIf lsControlType = "Menu" Then
            'Test if it's a separator or placeholder
            If lcControl.Caption <> "-" And lcControl.Caption <> "" Then
                lcControl.Caption = LoadResString(CInt(lcControl.Caption))
            End If
        ElseIf lsControlType = "TabStrip" Then
            For Each loObject In lcControl.Tabs
                loObject.Caption = LoadResString(CInt(loObject.Tag))
                loObject.ToolTipText = LoadResString(CInt(loObject.ToolTipText))
            Next
        ElseIf lsControlType = "Toolbar" Then
            For Each loObject In lcControl.Buttons
                'Test if it's a separator or placeholder
                If loObject.Style <> tbrSeparator And loObject.Style <> tbrPlaceholder Then
                    loObject.ToolTipText = LoadResString(CInt(loObject.ToolTipText))
                End If
            Next
        ElseIf lsControlType = "ListView" Then
            For Each loObject In lcControl.ColumnHeaders
                Set loObject.Font = loFont
                loObject.Text = LoadResString(CInt(loObject.Tag))
            Next
        ElseIf lsControlType = "TextBox" Then
            liValue = 0
            liValue = Val(lcControl.Tag)
            If liValue > 0 Then
                lcControl.Text = LoadResString(CInt(liValue))
            End If
            liValue = 0
            liValue = Val(lcControl.ToolTipText)
            If liValue > 0 Then
                lcControl.ToolTipText = LoadResString(CInt(liValue))
            End If
        ElseIf lsControlType = "CommonDialog" Then
            'do nothing - invisible control
        ElseIf lsControlType = "ImageList" Then
            'do nothing - invisible control
        Else
            liValue = 0
            liValue = Val(lcControl.Tag)
            If liValue > 0 Then
                lcControl.Caption = LoadResString(liValue)
            End If
            liValue = 0
            liValue = Val(lcControl.ToolTipText)
            If liValue > 0 Then
                lcControl.ToolTipText = LoadResString(liValue)
            End If
        End If
    Next

    Exit Sub

ErrorHandler:

    DisplayErrorMessage ("PublicModule - LoadResStrings")

    CleanupErrorCodes

    Resume Next

End Sub

Public Function AppInitialize() As Boolean
'This function is called the first time the program starts up.
'It is used to perform all the general overhead processing to initialize
'this application (e.g. getting system-wide information, connecting to
'databases and/or systems, etc.)
'Put any code here that you want to execute during startup.
'No parameters are passed, but a false must be returned if there
'is an error (otherwise, return true - i.e. no errors)

On Error GoTo ErrorHandler

    'Initialize any FileNET global stuff
    Set goDefaultLibrary = goNeighborhood.DefaultLibrary
    
    'Initialize any other application settings
    gbSuccess = LoadGlobalSettings
    If gbSuccess = False Then
        MsgBox LoadResString(GI_ERR_LOAD_GLOBAL_SETTINGS), vbInformation, LoadResString(GI_ERR_ERROR)
    End If
    
    AppInitialize = True
    
    Exit Function

ErrorHandler:

    'Display Error Message - pass the name of this subroutine/function
    DisplayErrorMessage ("PublicModule - AppInitialize")
    
    'Cleanup Error values
    CleanupErrorCodes

    'Return false to calling routine
    AppInitialize = False

End Function

Public Sub AppTerminate(pbSaveSettings As Boolean)
'This function is used to cleanup everything when your application
'terminates/shuts-down (either normally or when called during error
'handling).  Put all code here that you need to cleanup global memory/resources
' pbSaveSettings - do you want to save any global settings
'                  set to true for normal shutdown, set to false for abnormal shutdown

Dim loLibrary As IDMObjects.Library

On Error GoTo ErrorHandler

    'Save any application settings if requested
    If pbSaveSettings = True Then
        gbSuccess = SaveGlobalSettings
        If gbSuccess = False Then
            MsgBox LoadResString(GI_ERR_SAVE_GLOBAL_SETTINGS), vbInformation, LoadResString(GI_ERR_ERROR)
        End If
    End If

    'Cleanup global FileNET settings
    If gbLoggedOn = True Then
        For Each loLibrary In goNeighborhood.Libraries
            loLibrary.Logoff
        Next
    End If
    
    Set goNeighborhood = Nothing
    Set goDefaultLibrary = Nothing
        
    'End the whole program
    End
    
ErrorHandler:

    DisplayErrorMessage ("PublicModule - AppTerminate")

    'Cleanup Error values
    CleanupErrorCodes
    
    Resume Next

End Sub

Public Sub CleanupErrorCodes()
'This procedure cleans up global variable/objects for error handling

On Error Resume Next
    
    'Clear up Error values
    Err.Clear
    gsMsg = ""
    Screen.MousePointer = vbDefault
    glErrorCode = 0

End Sub

Public Sub DisplayErrorMessage(psProcedureName As String)
'This procedure is used to display an error message to the user.
'By default it uses a MsgBox.  You can customize this routine to display
'an error in whichever manner you want (e.g. show an error form that you design).

' psProcedureName - the name of the procedure where the error occurred.

Dim lsAppName As String
Dim liOriginalErrorNumber As Long
Dim lsOriginalErrorDescription As String

    'Store the original Err values because "On Error" resets the Err object
    liOriginalErrorNumber = Err.Number
    lsOriginalErrorDescription = Err.Description

On Error Resume Next

    'Set the name of this application in the VB project's properties,
    'or else this will display just the project name.
    If App.ProductName = "" Then
        lsAppName = App.Title
    Else
        lsAppName = App.ProductName
    End If
    
    'Try to build error message text here.
    'NOTE:  Make sure you have these global constant values corresponding to
    '       the correct entries in your .RC file, and you have compiled the .RC
    '       file into a .RES file using the Resource Compiler (RC.EXE),
    '       or else you will get a runtime error here.
    gsMsg = LoadResString(GI_ERR_ERROR_HAS_OCCURRED) & " " & LoadResString(GI_ERR_IN_THE_ROUTINE) _
            & " '" & psProcedureName & "' " & LoadResString(GI_ERR_IN_THE_PROGRAM) & vbCrLf _
            & lsAppName & " - " & LoadResString(GI_ERR_VERSION) & " " & App.Major & "." & App.Minor & vbCrLf _
            & LoadResString(GI_ERR_THE_ERROR_MESSAGE_IS) & "..." & vbCrLf & vbCrLf & GS_TAB _
            & LoadResString(GI_ERR_ERROR) & " #" & Str(liOriginalErrorNumber) & vbCrLf _
            & GS_TAB & lsOriginalErrorDescription & vbCrLf & vbCrLf _
            & LoadResString(GI_ERR_CUSTOM_ERROR_MESSAGE)  'Set this global constant to add your own custom error message
    
    If Err <> 0 Then
    'An error occurred preparing the message for the message box.  Typically this
    'is because the LoadResString call failed above due to a problem accessing the
    'correct resource value in the .RC/.RES file.
    'NOTE:  Set your global constant for GS_ERR_ERROR_ON_ERROR_DISPLAY and
    '       GS_ERR_ERROR in VB code rather than .RC file to display the
    '       following message.
        gsMsg = GS_ERR_ERROR_ON_ERROR_DISPLAY & vbCrLf & vbCrLf _
                & Err.Number & " - " & Err.Description & vbCrLf _
                & liOriginalErrorNumber & " - " & lsOriginalErrorDescription
    
        'Use GS_ERR_ERROR = text in global constant in VB code
        MsgBox gsMsg, vbExclamation, GS_ERR_ERROR, Err.HelpFile, Err.HelpContext
    Else
        'Use GI_ERR_ERROR = text from .RC/.RES file
        MsgBox gsMsg, vbExclamation, LoadResString(GI_ERR_ERROR), Err.HelpFile, Err.HelpContext
    End If

End Sub

Public Function LogonToLibrary(psLibraryLabel As String, psUserName As String, psPassword As String, psGroup As String) As Boolean
'This function attempts to logon to the library identified by the label sent in
'psLibraryLabel using the UserName passed in psUserName and the Password sent in
'psPassword.
'It returns whether or not it successfully logged on.

Dim loLibrary As IDMObjects.Library

On Error GoTo ErrorHandler
  
    gbSuccess = False
    
    'Search through all libraries in the neighborhood to find
    'the one that matches the label that the user selected.
    'NOTE:  If you don't get the library object from the neighborhood, then
    '       you must set the SystemType and Name or an error will occur
    '       when you try to Logon.  If you need to supply the SystemType, then
    '       you need to identify whether this is an IDMIS or IDMDS library.
    '       By going through the Neighborhood, you can avoid needing to
    '       determine whether it is an IDMIS or IDMDS server at this time.
    For Each loLibrary In goNeighborhood.Libraries
        If loLibrary.Label = psLibraryLabel Then
            Exit For  'we have found the library we want to logon to
        End If
    Next
    
    gbSuccess = loLibrary.Logon(psUserName, psPassword, psGroup, idmLogonOptNoUI)
        
    If gbSuccess <> False Then  'i.e. it was successful
        gbLoggedOn = True
        gsUserName = psUserName
        gsPassword = psPassword
        Set goLastLoggedOnLibrary = loLibrary
        gcLibraryCollection.Add loLibrary.Name, loLibrary.Label
        LogonToLibrary = True
    Else
        LogonToLibrary = False
    End If
    
    Exit Function

ErrorHandler:

    'Display Error message
    DisplayErrorMessage ("Public Module - LogonToLibrary")

    'Cleanup Error values
    CleanupErrorCodes
    
    Resume Next

End Function

Public Sub MouseWait()

    Screen.MousePointer = vbHourglass

End Sub

Public Sub MouseNormal()

    Screen.MousePointer = vbDefault
    
End Sub
' This little routine is here to deal with some current repaint
' problems in the tree control.  It forces a collapse and rebuild
' of the entire folder hierarchy
Public Sub ForceTreeRefresh(itvTV As IDMTreeView.IDMTreeView, _
    ByVal sSelectTarget As String)
' Hop back up to the parent library
While giCurSelTreeObjType = idmObjTypeFolder
    itvTV.SelectParentItem
Wend
' Force a refresh on the tree object
itvTV.Refresh
' Move down to the target object
itvTV.SelectChildItem (sSelectTarget)
End Sub
' Recursive subroutine for unfiling all the documents in a tree
' of folders, beginning at the passed folder
Public Sub CleanOutFolder(poThisFolder As IDMObjects.Folder)
Dim loDocList As IDMObjects.ObjectSet
Dim loFolderList As IDMObjects.ObjectSet
Dim loDoc As IDMObjects.Document
Dim loChildFolder As IDMObjects.Folder

Set loFolderList = poThisFolder.SubFolders
Set loDocList = poThisFolder.GetContents(idmFolderContentDocument)
For Each loChildFolder In loFolderList
    Call CleanOutFolder(loChildFolder)
Next
For Each loDoc In loDocList
    Call poThisFolder.Unfile(loDoc)
Next
End Sub

Public Function AddToIDMTreeView(pcControl As Control, poIDMObject As Object, pbExpand) As Boolean
'This function adds an IDM object to the IDMTreeView control.
'You pass it the name of the control, the object to add
'(e.g. Neighborhood, Library, Folder), and whether to expand
'the object in the tree view.

On Error GoTo ErrorHandler

    pcControl.AddRootItem poIDMObject, pbExpand
    
    AddToIDMTreeView = True
    
    Exit Function

ErrorHandler:

    'Display Error message
    DisplayErrorMessage ("PublicModule - AddToIDMTreeView - " & poIDMObject)

    'Cleanup Error values
    CleanupErrorCodes
    
    AddToIDMTreeView = False
    
End Function
Public Function FillIDMListView(pcControl As Control, poItem As Object, piObjectType As IDMTreeView.idmObjectType) As Boolean
'This function populates the IDMListView control with the contents
'of the Item passed in poItem.
'Pass the name of the IDMListView control that you want to use,
'the item you want to display the contents of, and the items type.

On Error GoTo ErrorHandler

Dim loContents As Object
    
    'Clears items in the ListView Control to fill it with new items
    pcControl.ClearItems
    
    Select Case piObjectType
        Case idmObjTypeFolder
            'Item clicked is a folder
            'Fill ListView with subfolders and store count...
            giFolderCount = pcControl.AddItems(goCurSelTreeItem.SubFolders, -1)
            
            'Get contents of selected folder
            Set loContents = goCurSelTreeItem.GetContents(idmFolderContentDocument)
            'Fill ListView with folder contents and store count...
            giDocumentCount = pcControl.AddItems(loContents, -1)
            
        Case idmObjTypeLibrary
            On Error Resume Next
            'Item clicked is a library
            'Fill ListView with top-level folders in the catalog
            Dim TreeItem As Object
            Set TreeItem = poItem.TopFolders
            If Not IsNull(TreeItem) Then
                pcControl.AddItems TreeItem, -1
            End If
    End Select
    
    FillIDMListView = True
    
    Exit Function

ErrorHandler:

    'Display Error message
    DisplayErrorMessage ("PublicModule - FillIDMListView")

    'Cleanup Error values
    CleanupErrorCodes
    
    FillIDMListView = False

    Resume Next
    
End Function

Public Sub UpdateStatusBar(pcControl As Control, psStatusMessage As String, Optional piPanel As Integer = 1)
'This routine updates the status bar.  Pass it the name of the
'status bar control, the message, and which panel to display it in
'(the default panel is the first panel).

    pcControl.Panels(piPanel).Text = psStatusMessage

End Sub

Public Function LoadGlobalSettings() As Boolean
'This function loads all global settings from the registry.
    
Dim lbAnyErrors As Boolean

On Error GoTo ErrorHandler

    lbAnyErrors = False

    'Add any global settings you want to load here
    gbSettingUseSysAdminLogon = GetSetting(App.Title, "Settings", "UseSysAdminLogon", False)
    
    If lbAnyErrors Then
        LoadGlobalSettings = False
    Else
        LoadGlobalSettings = True
    End If
    
    Exit Function

ErrorHandler:

    'Display Error message
    DisplayErrorMessage ("PublicModule - LoadGlobalSettings")

    'Cleanup Error values
    CleanupErrorCodes
    
    lbAnyErrors = True
    
    Resume Next
    
End Function
Public Function SaveGlobalSettings() As Boolean
'This function saves all global settings to the registry.

Dim lbAnyErrors As Boolean

On Error GoTo ErrorHandler

    lbAnyErrors = False

    'Add any global settings you want to save here
    SaveSetting App.Title, "Settings", "UseSysAdminLogon", gbSettingUseSysAdminLogon
    
    If lbAnyErrors Then
        SaveGlobalSettings = False
    Else
        SaveGlobalSettings = True
    End If
    
Exit Function

ErrorHandler:

    'Display Error message
    DisplayErrorMessage ("PublicModule - SaveGlobalSettings")

    'Cleanup Error values
    CleanupErrorCodes
    
    lbAnyErrors = True
    
    Resume Next
    
End Function

