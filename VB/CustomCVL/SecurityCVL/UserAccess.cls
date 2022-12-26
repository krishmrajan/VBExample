VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UserAccess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'This ActiveX dll works for the Thick client only
'
'Requirements
'   On Server, using the Admin Tool:
'              1. Create a CVL Plugin named SecurityCVL.UserAccess For its data, add SecurityCVL.UserAccess
'                 SecurityCVL.UserAccess is this project's name and the class that implements IDMObjects.PropertyDescriptionPlugIn
'              2. Create a regular CVL JobTitles and add the data listed in Class_Initialize
'              3. Create a Doc Class as follows:
'
'                       Doc Class Name = SecurityPlugin (Whatever you want)
'
'                       PropertyName | CVL Name
'                       ------------ | ---------
'             (Optional)Title        |
'                       SVCP Str CVL |  JobTitles
'                       SVCP Str     |  <PI>SecurityPlug<\PI>
'
'Note: "SVCP Str CVL" and "SVCP Str" are just labels here.


Implements IDMObjects.PropertyDescriptionPlugIn
Option Explicit
Option Base 1

Dim errorDesc As String
Dim myVal As Integer
Dim foo As Integer

Dim JobTitles() As String

Private Sub Class_Initialize()
    '+++Initialize array of JobTitles()
    ReDim JobTitles(5) As String

    JobTitles(1) = "Director"
    JobTitles(2) = "Manager"
    JobTitles(3) = "Sr. Engineer"
    JobTitles(4) = "Engineer"
    JobTitles(5) = "Intern"

End Sub

Private Sub Class_Terminate()
    '
End Sub

Private Property Get PropertyDescriptionPlugin_DefaultValue(ByVal Doc As Object, ByVal PropDesc As Object) As Variant
    'Set default value to Engineer if the Property is SVCP Str
    'This will work only if HasDefaultValue = True
    If PropDesc = SVCPStr Then
        PropertyDescriptionPlugin_DefaultValue = "Engineer"
    End If
End Property

Private Property Get PropertyDescriptionPlugin_ErrorDescription(ByVal Doc As Object, ByVal PropDesc As Object) As String
    PropertyDescriptionPlugin_ErrorDescription = errorDesc
End Property

Private Property Get PropertyDescriptionPlugin_FetchCount(ByVal Doc As Object, ByVal PropDesc As Object) As Long
    'Make sure Doc object is valid and is a type of document
    PropertyDescriptionPlugin_FetchCount = foo
End Property

Private Function PropertyDescriptionPlugin_FormatValue(ByVal Doc As Object, ByVal PropDesc As Object, ByVal ValueIn As Variant) As String
    PropertyDescriptionPlugin_FormatValue = ValueIn
End Function

Private Property Get PropertyDescriptionPlugin_HasDefaultValue(ByVal Doc As Object, ByVal PropDesc As Object) As Boolean
    'In Conditional Compilation Arguments, USE_DEFAULT = 1
    'Set USE_DEFAULT to 0 if you don't want to set a default value
    If SVCPStr = "" Then SVCPStr = FindPropertyName(Doc, "SVCP Str")
    If SVCPStrCVL = "" Then SVCPStrCVL = FindPropertyName(Doc, "SVCP Str CVL")
    
    #If USE_DEFAULT Then
        If PropDesc = SVCPStr Then
            PropertyDescriptionPlugin_HasDefaultValue = True
        End If
    #Else
        PropertyDescriptionPlugin_HasDefaultValue = False
    #End If
End Property

Private Property Get PropertyDescriptionPlugin_Item(ByVal Doc As Object, ByVal PropDesc As Object, ByVal Index As Long) As String
    'Display the values for each corresponding JobTitle
    'Don't really need to display the access level but for demo purposes, just show it.
   
    Select Case myVal
     Case 1
        'Admin account, by default, has the Owner access rights
        '(When you logon as someone else other than Admin, the
        'maximum access level you can set for Admin is Owner)
        Doc.Permissions.Add Doc.Library.ActiveUser, idmDSAccessOwner
        PropertyDescriptionPlugin_Item = "Owner"
     Case 2
        Doc.Permissions.Add Doc.Library.ActiveUser, idmDSAccessOwner
        PropertyDescriptionPlugin_Item = "Owner"
     Case 3
        Doc.Permissions.Add Doc.Library.ActiveUser, idmDSAccessAuthor
        PropertyDescriptionPlugin_Item = "Author"
     Case 4
        Doc.Permissions.Add Doc.Library.ActiveUser, idmDSAccessViewer
        PropertyDescriptionPlugin_Item = "Viewer"
     Case 5
        Doc.Permissions.Add Doc.Library.ActiveUser, idmDSAccessNone
        PropertyDescriptionPlugin_Item = "None"
     Case Else
        'This will happen if we've not chosen a JobTitle  Return "" for now.
        PropertyDescriptionPlugin_Item = ""
    End Select
    
End Property

Private Function PropertyDescriptionPlugin_PagePrefetch(ByVal Doc As Object, ByVal PropDesc As Object, ByVal vKey As Variant, ByVal KeyType As IDMObjects.idmPropPlugKeyType, ByVal PageSize As Long) As Long

    On Error GoTo ErrHandler
    Dim Prop As Property
    Dim propIter As Property
    
    'Find which property is SVCPStrCVL
    For Each propIter In Doc.Properties
        If propIter.Name = SVCPStrCVL Then
            Set Prop = propIter
            Exit For
        End If
    Next
    
    PropertyDescriptionPlugin_PagePrefetch = 1
    foo = 1 'Only need to show one item for User Access (SVCP Str)
       
    Select Case Prop.Value
    'Set Number of items to be returned in the fetch
    Case "Director"       'Owner
        myVal = 1
        PropertyDescriptionPlugin_PagePrefetch = 5
    Case "Manager"        'Owner
        myVal = 2
        PropertyDescriptionPlugin_PagePrefetch = 4
    Case "Sr. Engineer"   'Author
        myVal = 3
        PropertyDescriptionPlugin_PagePrefetch = 3
    Case "Engineer"       'Viewer
        myVal = 4
        PropertyDescriptionPlugin_PagePrefetch = 2
    Case "Intern"         'None
        myVal = 5
        PropertyDescriptionPlugin_PagePrefetch = 1
    
    Case Else
        'PropDesc will be empty on initialize. Do nothing
        myVal = 0
    End Select
    
    If foo > PageSize Then
        foo = PageSize
    End If
    
    Exit Function

ErrHandler:
    MsgBox Err.Description, vbCritical + vbOKOnly, Err.Source
End Function

Public Function FindPropertyName(ByVal objDocument As IDMObjects.Document, PropLabel As String) As String
    'Purpose: Search the server and return the internal name of the Property
    'Note   : SVCP Str CVL might be labeled differently. So just pass the correct label

    On Error GoTo ErrHandler

    Dim i, intCount As Integer
    intCount = objDocument.Properties.Count
    
    'Go through the list of properties
    '    Find the name of the property that matches PropLabel
    For i = 1 To intCount
        If objDocument.Properties.Item(i).Label = PropLabel Then
            FindPropertyName = objDocument.Properties.Item(i).Name
            Exit For
        End If
    Next
    
    'Did not find property
    If FindPropertyName = "" Then
        errorDesc = "Missing property " & PropLabel & " in the '" & objDocument.ClassDescription.Label & "' Doc class"
        Err.Raise ERR_MISSING_PROPERTY, ERR_APP_SOURCE, errorDesc
        FindPropertyName = 0
    End If
    
    Exit Function
    
ErrHandler:
    MsgBox Err.Description, vbCritical + vbOKOnly, Err.Source
End Function

Private Sub PropertyDescriptionPlugIn_ValidateValue(ByVal Doc As Object, ByVal PropDesc As Object, ByVal Value As String, pbVal As Boolean)
    'Note: validation will not work properly if the user typed both values the first time
    'A value has to be selected (Not typed) from the list so that the dependent property can be set
    
    On Error GoTo ErrorHandler
    
    Dim x As Integer
    Dim blPass As Boolean
    Dim strJobTitle As String
    Dim tempValue As String
    
    blPass = False 'Initialization
    strJobTitle = LCase(Doc.Properties(2).Value)
    
    'Clear any previous error
    errorDesc = ""
    Err.Clear
    
    If PropDesc = SVCPStr Then
       'Verify that the Access Level set corresponds to the JobTitle
       'Need to verify in case user typed a value
        tempValue = LCase(Value)
        Select Case tempValue
            Case "owner"
                If strJobTitle = "director" Then blPass = True
                If strJobTitle = "manager" Then blPass = True
            
            Case "author"
                If strJobTitle = "sr. engineer" Then blPass = True
            
            Case "viewer"
                If strJobTitle = "engineer" Then blPass = True
            
            Case "none"
                If strJobTitle = "intern" Then blPass = True
            
            Case Else
                errorDesc = "Invalid Job Title! Please select from the SVCP Str list"
                GoTo ErrorHandler
        End Select
    End If
    
     'No match
    If blPass = False Then
        errorDesc = "Invalid User Access was set for " & strJobTitle
        GoTo ErrorHandler
    End If
    
    Exit Sub
    
ErrorHandler:
    pbVal = False
End Sub
