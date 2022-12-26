Attribute VB_Name = "Globals"
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
' Workfile:   Globals.bas

'----------------------------------------------------------------------
' All Rights Reserved.  Copyright (c) 1999 FileNet Corp.
'----------------------------------------------------------------------


Option Explicit

Public bCantQuitNow As Boolean
Public sSourceDir As String, sTargetDir As String
Public oHood As Neighborhood

Public oServicedLibraries As Collection
Public oFS As FileSystemObject

'constants
Public Const REND_ENG_ID As String = "RE_SAMPLE_ENGINE"
Public Const REND_ENG_NAME As String = "FileNET Sample Rendition Engine"
Public Const TEMPLATE_ID As String = "RE_SAMPLE_TEMP_"
Public Const INI_FILE As String = "Rend_Engine.ini"

'

Function GetRelativePath(sParentPath, sChildPath) As String
'Returns the relative path of the child document to the parent document
'Will only work if the child is in the same folder or below then the parent
    
    Dim sChildDir As String, sParentDir As String
    Dim sDifference As String
    Dim sChildFileName As String
    
    sParentPath = LCase(sParentPath)
    sChildPath = LCase(sChildPath)
    
    sChildDir = LCase(oFS.GetParentFolderName(sChildPath))
    
    'add the slash at the end
    sChildDir = LCase(IIf(Len(sChildDir) = 3, sChildDir, sChildDir + "\"))
    
    sParentDir = LCase(oFS.GetParentFolderName(sParentPath))
    
    'add slash at the end, unless its something like "c:\"
    sParentDir = IIf(Len(sParentDir) = 3, sParentDir, sParentDir + "\")
    
    'remove directory from child full path
    sChildFileName = LCase(oFS.GetFileName(sChildPath))
    
    'find difference between direcotries
    sDifference = Replace(sChildDir, sParentDir, "")
    If sDifference = "" Then
        sDifference = "\"
    Else
        sDifference = "\" + sDifference
    End If
    
    GetRelativePath = "." + sDifference + sChildFileName
    
End Function

Sub StatusBarText(sText As String)
'change status bar text
    
    frmMain.stbStatus.Panels("text").Text = sText
End Sub

Sub EmptyFolder(sFolderPath As String)
'Clear all the files from a directory

On Error GoTo ErrorHandler
    
    Dim oTempFile As File, oSubFolder As Scripting.Folder
        
    For Each oTempFile In oFS.GetFolder(sFolderPath).Files
        oTempFile.Delete True
    Next

    For Each oSubFolder In oFS.GetFolder(sFolderPath).SubFolders
        EmptyFolder (oSubFolder.Path)
        oSubFolder.Delete
    Next
    
Exit Sub
ErrorHandler:
    AddMessage "error while emptying folder " + sFolderPath + " " + Err.Description
    Resume Next
End Sub


Sub ClearTemporaryFolders()
'Clear both temporary directories

    EmptyFolder sTargetDir
    EmptyFolder sSourceDir
End Sub

Sub LoadLibrarySettings()
'Reads configuration file, and creates ServicedLib objects


    On Error GoTo ErrorHandler
    Dim nFileHandle As Integer, oTempServedLib As ServicedLib
    
    nFileHandle = FreeFile()
    
    Open App.Path + "\" + INI_FILE For Input As nFileHandle
    
    Set oServicedLibraries = New Collection
    
    Do While Not EOF(nFileHandle)
        Set oTempServedLib = New ServicedLib
        oTempServedLib.ReadSettings nFileHandle
        If oTempServedLib.IsWorkingOK Then
            oServicedLibraries.Add oTempServedLib
        End If
    Loop
    
    Close nFileHandle
Exit Sub
ErrorHandler:
    AddMessage "Configuration file is missing...  Please Add Libraries"
End Sub


Sub SaveLibrarySettings()
'Saves library configuration to .ini file

    On Error GoTo ErrorHandler
    
    Dim nFileHandle As Integer, oTempServedLib As ServicedLib
    nFileHandle = FreeFile()
    Open App.Path + "\" + INI_FILE For Output As nFileHandle
    
    For Each oTempServedLib In oServicedLibraries
        oTempServedLib.WriteSettings nFileHandle
    Next
    
    Close nFileHandle

Exit Sub
ErrorHandler:
    MsgBox Err.Description, , "Saving Library Settings"
End Sub

Public Sub AddMessage(Optional sMessage As String)
'Adds a message to the "Messages" text box
'if sMessage is blank, it does not put the time, just adds a blank line
   
    frmMain.txtMessages.Text = frmMain.txtMessages.Text + _
    IIf(sMessage <> "", _
        "[" + CStr(Now()) + "]   " + sMessage + vbCrLf, _
        vbCrLf)
    
    DoEvents
    
End Sub

Sub CreateDirectories()
'Create required sub folders, in the folder with the .EXE file
    On Error GoTo ErrorHandler:
    
    sSourceDir = App.Path + "\Source\"
    sTargetDir = App.Path + "\Published\"
    
    If Not oFS.FolderExists(sSourceDir) Then
        oFS.CreateFolder sSourceDir
    End If
    
    If Not oFS.FolderExists(sTargetDir) Then
        oFS.CreateFolder sTargetDir
    End If

Exit Sub
ErrorHandler:
    MsgBox "Error while creating a Sub Directory.  Make sure you have 'write' priveleges"
    End
End Sub

Sub CenterForm(frmForm As Form)
'centers the form on screen

    frmForm.Left = (Screen.Width - frmForm.Width) / 2
    frmForm.Top = (Screen.Height - frmForm.Height) / 2
End Sub

Sub ShowError(Optional strWhere As String, Optional bShowMessageBox As Boolean = False)
    
    AddMessage "Error '" + Err.Description + "' in '" + Err.Source + "'  (" + strWhere + ")"
    If bShowMessageBox Then
    MsgBox "Number:       " + CStr(Err.Number) + vbCrLf + _
           "Description:  " + Err.Description + vbCrLf + _
           "Source:        " + Err.Source, vbCritical, "ERROR!!! " + strWhere
    End If
End Sub

Sub DisplayStyleTemplate(oStyleTemp As StyleTemplate)
'Displays Style Template properties
    
    Dim frmDlg As New frmStyleTemplate
    Set frmDlg.oStyleTemplate = oStyleTemp
    frmDlg.Show vbModal
End Sub

Function FindAllFiles(sParentFile As String) As Collection
'returns a collection with strings with full path of all files in the directory and sub directory
'exect the sParentFile

    Dim oFldToLookIn As Scripting.Folder, oCollection As Collection, _
        i As Integer, nToRemove As Integer

    Set oCollection = New Collection
    
    Set oFldToLookIn = oFS.GetFolder(oFS.GetParentFolderName(sParentFile))
    
    'call the recursive function to put all file paths inside the folder in the collection
    GetAllFiles oCollection, oFldToLookIn
    
    
    'Find and remove sParentFile from the collection
    For i = 1 To oCollection.Count
        If UCase(oCollection(i)) = UCase(sParentFile) Then
            nToRemove = i
            Exit For
        End If
    Next
    
    
    If nToRemove <> 0 Then
        oCollection.Remove nToRemove
    End If
    
    
    Set FindAllFiles = oCollection
    
End Function

Sub GetAllFiles(oCollection As Collection, oDirecotry As Scripting.Folder)
    'add full paths of all the files in the folder and in subfolders
    
    Dim oFile As Scripting.File, oSubDirectory As Scripting.Folder
    
    For Each oFile In oDirecotry.Files
        oCollection.Add oFile.Path
    Next
    
    'recursive call
    For Each oSubDirectory In oDirecotry.SubFolders
        GetAllFiles oCollection, oSubDirectory
    Next
    
End Sub
