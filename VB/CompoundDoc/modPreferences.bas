Attribute VB_Name = "modPreferences"
Public Sub SetPreference(sSubSystemName As String, sCategorryName As String, sPreferenceName As String, sPreferenceValue As String)
   '--------------------------------------------------
   'Purpose: To set preference to a library
   'Input: SubSystemName,CategorryName,PreferenceValue
   'Output: None
   'Result: A preference is set
   '--------------------------------------------------
   Dim oPrefs As IDMPreferences.Preferences
   Dim oPref As New IDMPreferences.Preference
   Dim oSubSystem As New IDMPreferences.SubSystem
   Dim oCategorry As IDMPreferences.Category
   Dim oOption As New IDMPreferences.Option
   Dim oBehavior As IDMObjects.Behavior
    
   On Error GoTo ErrHandler
   
   Set oBehavior = oDocument.Compound.Behavior
   Set oPrefs = oBehavior.Preferences
   
   oSubSystem.Name = sSubSystemName
   oSubSystem.UserType = idmPoUserCurrent
   
   'check preferences exit or not if exit exit the routine, otherwise create preferences
   If GetPreferenceValue(sSubSystemName, sCategorryName, sPreferenceName) <> "NoValue" Then
      Exit Sub
   End If
   Set oCategorry = oSubSystem.CreateCategory(sCategorryName)
   Set oPref = oCategorry.CreatePreference(sPreferenceName)
   oOption.Value = sPreferenceValue
   oPref.Value = oOption
   oPrefs.Add oPref

   Exit Sub
ErrHandler:
   MsgBox Err.Number & "  " & Err.Description, vbCritical, "Set Preferences"
End Sub
Public Function GetPreferenceValue(sSubSystemName As String, sCategoryName As String, sPreferenceName As String) As Variant
    '--------------------------------------------------
    'Purpose: To get preference
    'Input: SubSystemName,CategorryName,PreferenceValue
    'Output: A preference value
    'Result: A preference value
    '--------------------------------------------------
    Dim oSubSystem As New IDMPreferences.SubSystem
    Dim oCategory As New IDMPreferences.Category
    Dim oCategories As New IDMPreferences.Categories
    Dim oPreference As IDMPreferences.Preference
    Dim oValue As IDMPreferences.Option
    Dim varPrefValue As Variant
    Dim CategoryName As String
    
    On Error GoTo ErrorHandler
    
    'Initializes the category object
    oSubSystem.Name = sSubSystemName
    oSubSystem.UserType = idmPoUserCurrent
    Set oCategory = oSubSystem.GetCategory(sCategoryName)
    
    'Gets the preference
    Set oPreference = oCategory.GetPreference(sPreferenceName)
    Set oValue = oPreference.Value
    varPrefValue = oValue.Value
    GetPreferenceValue = varPrefValue
    
    Exit Function
    
ErrorHandler:
    If Err.Number = -2147201493 Then
       GetPreferenceValue = "NoValue"
       Exit Function
    End If
    MsgBox Err.Number & " " & Err.Description, vbCritical, "Get Preference Value"
End Function


