VERSION 5.00
Begin VB.Form PreferencesForm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Simple Preferences Sample"
   ClientHeight    =   2655
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7110
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   7110
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Import 
      Caption         =   "Import"
      Height          =   375
      Left            =   1800
      TabIndex        =   7
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton Export 
      Caption         =   "Export"
      Height          =   375
      Left            =   360
      TabIndex        =   6
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton Exit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   5640
      TabIndex        =   4
      Top             =   2160
      Width           =   1335
   End
   Begin VB.TextBox CustomPrefTextBox 
      Height          =   525
      Left            =   3240
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   1320
      Width           =   3735
   End
   Begin VB.CommandButton StandardPrefReset 
      Caption         =   "Reset to Default"
      Height          =   375
      Left            =   5640
      TabIndex        =   2
      Top             =   600
      Width           =   1335
   End
   Begin VB.CheckBox CustomPrefCheckBox 
      Caption         =   "Display this message when starting up"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   3015
   End
   Begin VB.CheckBox StandardPrefCheckBox 
      Caption         =   "Keep a copy of the document on the hard disk after doing a check in"
      Height          =   435
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   5295
   End
   Begin VB.Label CacheLabel 
      Caption         =   "Your documents will be cached in xxxxxxxxxxxxxxxxxxx"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   6855
   End
End
Attribute VB_Name = "PreferencesForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
' $Revision:   1.2  $
' $Date:   30 Aug 2000 08:56:50  $
' $Author:   nbader  $
' $Workfile:   preferencesform.frm  $
'----------------------------------------------------------------------
' All Rights Reserved.  Copyright (c) 1988,1999 FileNet Corp.
'----------------------------------------------------------------------

Option Explicit

Private Sub Exit_Click()
    Unload Me
End Sub

Private Sub Form_Load()

    Dim preferenceValue As Variant
    Dim preferenceValueStr As String

    ' here is a sample of how to display a preference value for the user
    If (GetPreferenceValue("DirectoriesAndFiles", "LocalCaching", "CacheDir", preferenceValue, preferenceValueStr)) Then
        CacheLabel.Caption = "Your documents will be cached in " & preferenceValue
    Else
        MsgBox "Preference value for local cache directory is missing.", vbCritical, "Error loading form"
        CacheLabel.Caption = "Error!!! Missing Preference value for local cache directory."
    End If
    
    ' here is a sample of a standard preference which you can modify without having the user run the IDM Configure application
    If (GetPreferenceValue("Documents", "AddCheckinRetrieve", "KeepLocalCopy", preferenceValue, preferenceValueStr)) Then
        StandardPrefCheckBox.value = preferenceValue
    Else
        MsgBox "Preference value for keep local copy is missing.", vbCritical, "Error loading form"
        StandardPrefCheckBox.value = 0
    End If
    
    
    ' the rest of this procedure is a sample of how to add your own preferences to an application
    
    ' check if the custom preferences have been created
    Dim preference1Exists As Boolean
    Dim preference2Exists As Boolean
    preference1Exists = GetPreferenceValue("Documents", "PreferenceSample", "ShowMessage", preferenceValue, preferenceValueStr)
    preference2Exists = GetPreferenceValue("Documents", "PreferenceSample", "MessageText", preferenceValue, preferenceValueStr)
    If ((Not preference1Exists) Or (Not preference2Exists)) Then
        ' create the custom preferences
        Call CreateCustomPreferences
    End If
    
    ' get the custom preference for showing the message box
    Dim showMessage As Boolean
    If (GetPreferenceValue("Documents", "PreferenceSample", "ShowMessage", preferenceValue, preferenceValueStr)) Then
        CustomPrefCheckBox.value = preferenceValue
        showMessage = preferenceValue
    Else
        MsgBox "Preference value for showing message is missing.", vbCritical, "Error loading form"
        CustomPrefCheckBox.value = 0
        showMessage = False
    End If
    
    ' get the custom preference which contains the text for the message box
    If (GetPreferenceValue("Documents", "PreferenceSample", "MessageText", preferenceValue, preferenceValueStr)) Then
        CustomPrefTextBox.Text = preferenceValue
    Else
        MsgBox "Preference value for message text is missing.", vbCritical, "Error loading form"
        CustomPrefTextBox.Text = "Error!!! Missing Preference value for message text."
    End If
    
    ' modify the behavior of this sample using the custom preferences
    If (showMessage) Then
    Dim messageText As String
        messageText = CustomPrefTextBox.Text
        MsgBox messageText, , "Startup Message"
    End If
        
End Sub

Private Sub CustomPrefCheckBox_Click()
    ' save the preference when the user clicks on the checkbox
    Dim value As Integer
    value = CustomPrefCheckBox.value
    Call SetPreferenceValue("Documents", "PreferenceSample", "ShowMessage", value)
End Sub

Private Sub CustomPrefTextBox_Change()
    ' save the preference when the user modifies the text
    Dim message As String
    message = CustomPrefTextBox.Text
    Call SetPreferenceValue("Documents", "PreferenceSample", "MessageText", message)
End Sub

Private Sub StandardPrefCheckBox_Click()
    ' save the preference when the user clicks on the checkbox
    Dim value As Integer
    value = StandardPrefCheckBox.value
    Call SetPreferenceValue("Documents", "AddCheckinRetrieve", "KeepLocalCopy", value)
End Sub

Private Sub StandardPrefReset_Click()
    ' reset the preference to the default values
    Call ResetPreferenceValue("Documents", "AddCheckinRetrieve", "KeepLocalCopy")
    
    ' update the UI to match the new preference value
    Dim preferenceValue As Variant
    Dim preferenceValueStr As String
    If (GetPreferenceValue("Documents", "AddCheckinRetrieve", "KeepLocalCopy", preferenceValue, preferenceValueStr)) Then
        StandardPrefCheckBox.value = preferenceValue
    Else
        MsgBox "Preference value for keep local copy is missing.", vbCritical, "Error loading form"
        StandardPrefCheckBox.value = 0
    End If
End Sub

Private Sub CreateCustomPreferences()

    Dim oPreference As IDMPreferences.Preference
    Dim oOptions As IDMPreferences.Options
    Dim oOption As IDMPreferences.Option
    
    On Error GoTo ErrorHandler
    
    ' first we create a subsystem
    Dim oSubSystem As New IDMPreferences.SubSystem
    ' specify which subsystem to use, for application specific preferences, we suggest that
    ' you use "Documents" and the default (new users) type
    oSubSystem.Name = "Documents"
    oSubSystem.UserType = idmPoUserDefault
    
    ' now, set a category from the subsystem
    Dim oCategory As IDMPreferences.Category
    On Error GoTo ErrorGetCategory
    ' try to get the category .. if there is an error, the handler will create a new category
    Set oCategory = oSubSystem.GetCategory("PreferenceSample")
    On Error GoTo ErrorHandler
    ' set the properties for the category
    oCategory.Label = "Preferences Sample Code"
    oCategory.HelpString = "This is a category used by the preferences sample code."
    oCategory.Save

    ' now, set a preference from the category
    On Error GoTo ErrorGetPreference1
    ' try to get the preference .. if there is an error, the handler will create a new preference
    Set oPreference = oCategory.GetPreference("ShowMessage")
    On Error GoTo ErrorHandler
    ' set the properties for the preference .. this one is for a simple checkbox
    oPreference.Label = "Show message when preferences sample code starts"
    oPreference.UIType = idmPoUICheckbox
    
    ' remove all the options from the preference so we can add new ones
    Set oOptions = oPreference.Options
    oOptions.RemoveAll
    
    ' create an option for the default value
    Set oOption = CreateObject("IDMPreferences.Option")
    oOption.value = 1
    oOption.ValueLabel = "Yes"
    oOptions.Add oOption
    
    ' set the preference's current value and its default value
    oPreference.value = oOption
    oPreference.ValueDefault = oOption
    
    ' create an option for the other state of the checkbox
    Set oOption = CreateObject("IDMPreferences.Option")
    oOption.value = 0
    oOption.ValueLabel = "No"
    oOptions.Add oOption
    
    ' save the newly created preference
    oPreference.Save
    
    ' now, set another preference from the category
    On Error GoTo ErrorGetPreference2
    ' try to get the preference .. if there is an error, the handler will create a new preference
    Set oPreference = oCategory.GetPreference("MessageText")
    On Error GoTo ErrorHandler
    ' set the properties for the preference .. this one is the just plain text
    oPreference.Label = "The message that will be displayed when the preferences sample code starts"
    oPreference.UIType = idmPoUIEditbox
    
    ' the editbox and browse preferences do not have any options (text only)
    oPreference.Options.RemoveAll
    
    ' set the preference's current value and its default value
    oPreference.value.value = "Hey!"
    oPreference.ValueDefault.value = "Hey!"
    
    ' save the newly created preference
    oPreference.Save
    
    Exit Sub
    
ErrorHandler:
    ' general error handler
    MsgBox Err.Description, vbCritical, "Create Custom Preferences"
    Exit Sub
    
ErrorGetCategory:
    ' for errors with GetCategory
    Set oCategory = oSubSystem.CreateCategory("PreferenceSample")
    Resume Next

ErrorGetPreference1:
    ' for errors with GetPreference of "ShowMessage"
    Set oPreference = oCategory.CreatePreference("ShowMessage")
    Resume Next

ErrorGetPreference2:
    ' for errors with GetPreference of "MessageText"
    Set oPreference = oCategory.CreatePreference("MessageText")
    Resume Next
    
End Sub

Private Sub Export_Click()
    ' use the preference manager to export the current user's preferences to a file
    Call ExportPreferences
    ' note that the export file will contain the custom preferences
End Sub

Private Sub Import_Click()
    ' use the preference manager to import the current user's preferences from a file
    Call ImportPreferences
    ' note that the import will include the custom preferences, so let's refresh the form
    Call Form_Load
End Sub


