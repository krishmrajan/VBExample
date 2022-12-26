Attribute VB_Name = "modPreferences"
' This program is an example of how to use the new foundation objects
' for managing user preferences
'----------------------------------------------------------------------
'   Disclaimer: FileNET provides programming examples for illustration
'   only, without warranty either expressed or implied,  including, but
'   not limited to, the implied warranties of merchantability and/or
'   fitness for a particular purpose. This sample assumes that you are
'   familiar with the programming language being demonstrated and the
'   tools used to create and debug procedures.
'----------------------------------------------------------------------

'----------------------------------------------------------------------
' $Revision:   1.1  $
' $Date:   03 Nov 1999 14:48:54  $
' $Author:   nbader  $
' $Workfile:   modPreferences.bas  $
'----------------------------------------------------------------------
' All Rights Reserved.  Copyright (c) 1988,1999 FileNet Corp.
'----------------------------------------------------------------------

Option Explicit

Public Sub SetPreferenceValue(subSystemName As String, categoryName As String, preferenceName As String, newPreferenceValue As Variant)
       
    On Error GoTo ErrorHandler
    
    ' create a subsystem object and set its name and user type
    Dim oSubSystem As New IDMPreferences.SubSystem
    oSubSystem.Name = subSystemName
    oSubSystem.UserType = idmPoUserCurrent
    
    ' select the category of the subsystem
    Dim oCategory As IDMPreferences.Category
    Set oCategory = oSubSystem.GetCategory(categoryName)
    
    ' select the preference of the category
    Dim oPreference As IDMPreferences.Preference
    Set oPreference = oCategory.GetPreference(preferenceName)
    
    ' set the preference's new value and save it
    oPreference.Value.Value = newPreferenceValue
    oPreference.Save
    
    Exit Sub
    
ErrorHandler:
   MsgBox Err.Description, vbCritical, "Set Preference Value"
   
End Sub

Public Sub ResetPreferenceValue(subSystemName As String, categoryName As String, preferenceName As String)
          
    On Error GoTo ErrorHandler
    
    ' create a subsystem object and set its name and user type
    Dim oSubSystem As New IDMPreferences.SubSystem
    oSubSystem.Name = subSystemName
    oSubSystem.UserType = idmPoUserCurrent
    
    ' select the category of the subsystem
    Dim oCategory As IDMPreferences.Category
    Set oCategory = oSubSystem.GetCategory(categoryName)
    
    ' select the preference of the category
    Dim oPreference As IDMPreferences.Preference
    Set oPreference = oCategory.GetPreference(preferenceName)
    
    ' set the preference's default value and save it
    oPreference.Value.Value = oPreference.ValueDefault.Value
    oPreference.Save
    
    Exit Sub
    
ErrorHandler:
   MsgBox Err.Description, vbCritical, "Reset Preference Value"
   
End Sub

Public Function GetPreferenceValue(subSystemName As String, categoryName As String, preferenceName As String, _
        preferenceValue As Variant, preferenceValueStr As String) As Boolean
    
    On Error GoTo ErrorHandler
    
    ' create a subsystem object and set its name and user type
    Dim oSubSystem As New IDMPreferences.SubSystem
    oSubSystem.Name = subSystemName
    oSubSystem.UserType = idmPoUserCurrent
    
    ' select the category of the subsystem
    Dim oCategory As IDMPreferences.Category
    Set oCategory = oSubSystem.GetCategory(categoryName)
    
    ' select the preference of the category
    Dim oPreference As IDMPreferences.Preference
    Set oPreference = oCategory.GetPreference(preferenceName)
    
    ' get the currently selected option for the preference
    Dim oOption As IDMPreferences.Option
    Set oOption = oPreference.Value
    
    ' return the preference value as well as its textual description
    preferenceValue = oOption.Value
    preferenceValueStr = oOption.ValueLabel
    
    ' return flag that indicates we succeeded in getting the preference value
    GetPreferenceValue = True
    
    Exit Function
    
ErrorHandler:
    ' return flag that indicates there was an error (probably because the preference didn't exist)
    GetPreferenceValue = False
    
End Function

Public Sub ImportPreferences()

    On Error GoTo ErrorHandler
    
    ' create a preference manager to do the import (using current user)
    Dim oManager As New IDMPreferences.Manager
    oManager.UserType = idmPoUserCurrent
    
    ' import a file and check the results
    Dim result As Long
    Dim importedFileStr As String
    result = oManager.Import(idmPoIOWithUI, , importedFileStr)
    If result = 1 Then
        MsgBox "Imported " & importedFileStr & " successfully.", , "Import"
    End If

    Exit Sub
    
ErrorHandler:
    MsgBox Err.Description, vbCritical, "Import Error"

End Sub

Public Sub ExportPreferences()

    On Error GoTo ErrorHandler
    
    ' create a preference manager to do the export (using all subsystems of current user)
    Dim oManager As New IDMPreferences.Manager
    oManager.UserType = idmPoUserCurrent
    oManager.AddAll
    
    ' export the file and check the results
    Dim result As Long
    Dim exportedFileStr As String
    result = oManager.Export(idmPoIOWithUI, , , exportedFileStr)
    If result = 1 Then
        MsgBox "Exported " & exportedFileStr & " successfully.", , "Export"
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox Err.Description, vbCritical, "Export Error"

End Sub

