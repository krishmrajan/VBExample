VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Geography"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************************************************************************
' Requirements
'   On Server, using the Admin Tool:
'              1. Create a CVL Plugin named Whatever. For its data, add MultiDependents.Geography
'                 MultiDependents.Geography is this project's name and the class that implements IDMObjects.IFnPropertyDescriptionPlugin
'              2. Create a Doc Class as follows:
'
'                       Doc Class Name = MyDocClass
'
'                       PropertyName | CVL Name
'                       ------------ | ---------
'             (Optional)Title        |
'             (Country) SVCP Str     |  <PI>Whatever<\PI>
'             (State)   SVCP Str CVL |  <PI>Whatever<\PI>
'             (City)    SVCP Str CVL2|  <PI>Whatever<\PI>
'
'  Note: "SVCP Str *" are just some names here. They represent some
'        Single value properties that have a CVL assigned to them.
'
'  On Client:
'             1. Register the Dll using regsvr32.exe or use VB to recompile the project
'             2. CustomCVL.mdb must exist in the same folder where the dll is registered
'
'*******************************************************************************************************************************


Implements IDMObjects.PropertyDescriptionPlugIn
Option Explicit
Option Base 1

Dim ErrDesc As String
Dim foo As Integer
 
'Datasource Class Objects
Dim CountryCLS As Database
Dim StateCLS As Database
Dim CityCLS As Database

Private Sub Class_Terminate()
    'Terminate classes
    Set CountryCLS = Nothing
    Set StateCLS = Nothing
    Set CityCLS = Nothing
End Sub

Private Property Get PropertyDescriptionPlugin_DefaultValue(ByVal Doc As Object, ByVal PropDesc As Object) As Variant
    'Set default value to USA if the Property is SVCP STR (Country)
    'This will work only if HasDefaultValue = True
    'This will not work if the property already had a value
   
    If PropDesc = SVCPStr Then
        PropertyDescriptionPlugin_DefaultValue = ReadFile("DefaultValue")
        strCountry = PropertyDescriptionPlugin_DefaultValue 'Store country for validation later
    End If
End Property

Private Function ReadFile(section As String) As Variant
    'This function is used to read default values set in the Defaults.ini file
    'Returns a boolean for HasDefaultValue and returns a string for DefaultValue
    On Error GoTo ErrHandler
    Dim intFileNum As Integer
    Dim strLine As String
    Dim intLength As Integer
    Dim x As Integer
    
    intFileNum = FreeFile
    
    'Open Configuration file
    Open App.Path & "\Defaults.ini" For Input As intFileNum
    
    Do While Not EOF(intFileNum)
        Line Input #intFileNum, strLine
        intLength = Len(strLine)
        If intLength <> 0 And Left(strLine, 1) <> "'" Then  'Skip blanks or comments
            For x = 1 To Len(strLine)
                If Mid(strLine, x, 1) = "=" Then
                    If Mid(strLine, 1, x - 1) = section Then    'Find section
                        ReadFile = Mid(strLine, x + 1, intLength - x)   'Get Right Hand Side value
                    End If
                End If
            Next
            
        End If
    Loop
    Close intFileNum
    Exit Function
    
ErrHandler:
    If Err.Number = 53 Then 'Defaults.ini does not exist. Create one
        Open App.Path & "\Defaults.ini" For Output As intFileNum
        Print #intFileNum, "HasDefaultValue=true"
        Print #intFileNum, "DefaultValue=USA"
        Close #intFileNum
    End If
End Function

Private Property Get PropertyDescriptionPlugin_ErrorDescription(ByVal Doc As Object, ByVal PropDesc As Object) As String
    'Return the error description whenever there is a problem
    PropertyDescriptionPlugin_ErrorDescription = ErrDesc
End Property

Private Property Get PropertyDescriptionPlugin_FetchCount(ByVal Doc As Object, ByVal PropDesc As Object) As Long
    PropertyDescriptionPlugin_FetchCount = foo
End Property

Private Function PropertyDescriptionPlugin_FormatValue(ByVal Doc As Object, ByVal PropDesc As Object, ByVal ValueIn As Variant) As String
    'Format the Item that is returned to display
    'ValueIn is case sensitive
    PropertyDescriptionPlugin_FormatValue = ValueIn
End Function

Private Property Get PropertyDescriptionPlugin_HasDefaultValue(ByVal Doc As Object, ByVal PropDesc As Object) As Boolean
    'Purpose: Determine whether the property has a default value or not
    
    'This is the first method fired and the first time we can have access to the Doc object.
    'Initialize the property names here
    If Doc Is Nothing Then
        AssignPropertyNames False
    Else
        AssignPropertyNames True, Doc
    End If
   
    If PropDesc = SVCPStr Then
        PropertyDescriptionPlugin_HasDefaultValue = ReadFile("HasDefaultValue")
    End If
 
End Property

Private Sub AssignPropertyNames(Thick As Boolean, Optional oDoc As Object)
    If Thick Then
        'Use the document object to find the property name
        If SVCPStr = "" Then SVCPStr = FindPropertyName(oDoc, "SVCP Str")
        If SVCPStrCVL = "" Then SVCPStrCVL = FindPropertyName(oDoc, "SVCP Str CVL")
        If SVCPStrCVL2 = "" Then SVCPStrCVL2 = FindPropertyName(oDoc, "SVCP Str CVL 2")
    Else
        'Thin client - can't use Doc object
        'Hardcode the property names (You have to find out what they are on the server)
        If SVCPStr = "" Then SVCPStr = "idmDocCustom8"
        If SVCPStrCVL = "" Then SVCPStrCVL = "idmDocCustom9"
        If SVCPStrCVL2 = "" Then SVCPStrCVL2 = "idmDocCustom10"
    End If
End Sub

Private Property Get PropertyDescriptionPlugin_Item(ByVal Doc As Object, ByVal PropDesc As Object, ByVal Index As Long) As String
    'This property will be called for each index (number of records available)
    'Return each value from the array
    On Error Resume Next    'To ignore invalid indexes
    Select Case PropDesc
    Case SVCPStr
        PropertyDescriptionPlugin_Item = CountryArray(Index)
    Case SVCPStrCVL
        PropertyDescriptionPlugin_Item = StateArray(Index)
    Case SVCPStrCVL2
        PropertyDescriptionPlugin_Item = CityArray(Index)
    Case Else
        PropertyDescriptionPlugin_Item = "Nothing"
    End Select
End Property

Public Function SetorGetArrayValues(WhichArray As String, numRecords As Integer, ByVal vkey As String) As Variant
    'Purpose: Fill Arrays using data from database tables
    'Returns: An integer representing the index of the record in the database
    
    'The first time the user clicks on the drop down of the property value field, events will
    'fire and vKey will be null. So when this function is called the first time, the arrays
    'will be filled with all valid values. In the PagePrefetch method, when SetorGetArrayValues
    'is called, we don't need to go through the whole table and repopulate the array.
    'We just need to check which record vKey corresponds to the value in the database
    
    On Error Resume Next
    Dim x As Integer    'Counter

    Select Case WhichArray
        Case "Country"
            vkey = UCase(vkey)   'Set to uppercase in case user typed a value
            ReDim CountryArray(numRecords)
            For x = 1 To numRecords
                If CountryArray(x) = "" Then CountryArray(x) = CountryCLS.CorrespondingData
                If vkey = CountryArray(x) Then
                    SetorGetArrayValues = x
                    Exit Function
                Else
                    'The first time user clicks on the property, vKey will be null
                    'We need to return 1 so that PagePrefetch is 1 for this case
                    SetorGetArrayValues = 1
                End If
            Next
        
        Case "State"
            If numRecords = 0 Then
                StateArray(1) = "N/A"
            Else
                vkey = SetLetterCase(vkey)
                ReDim StateArray(numRecords)
                For x = 1 To numRecords
                    If StateArray(x) = "" Then StateArray(x) = StateCLS.CorrespondingData
                    If vkey = StateArray(x) Then
                       SetorGetArrayValues = x
                       Exit Function
                    Else
                        SetorGetArrayValues = 1
                    End If
                 Next
            End If
        
        Case "City"
            If numCities = 0 Then
                CityArray(1) = "N/A"
            Else
                vkey = SetLetterCase(vkey)
                ReDim CityArray(numRecords)
                For x = 1 To numRecords
                    If CityArray(x) = "" Then CityArray(x) = CityCLS.CorrespondingData
                    If vkey = CityArray(x) Then
                        SetorGetArrayValues = x
                        Exit Function
                    Else
                        SetorGetArrayValues = 1
                    End If
                Next
            End If
        
    End Select

End Function

Private Function PropertyDescriptionPlugin_PagePrefetch(ByVal Doc As Object, ByVal PropDesc As Object, ByVal vkey As Variant, ByVal KeyType As IDMObjects.idmPropPlugKeyType, ByVal PageSize As Long) As Long
    On Error Resume Next
     
    PropertyDescriptionPlugin_PagePrefetch = 1
    
    'FR 50886
    
    If Doc Is Nothing Then
        AssignPropertyNames False
    Else
        AssignPropertyNames True, Doc
    End If
    
    Select Case PropDesc
    'PropDesc tells which specific property is being used
    Case SVCPStr    'Country
        If vkey <> Null Or vkey <> "" Then
            strCountry = vkey   'set to current selected country
        End If
        
        strSQL = SetSQLStatement("Country")
        Set CountryCLS = New Database
        numCountries = CountryCLS.RecordCount
        
        If vkey <> Null Or vkey <> "" Then
            'Find the index of the record in the database
            PropertyDescriptionPlugin_PagePrefetch = SetorGetArrayValues("Country", numCountries, vkey)
        Else
            'Initialize array of countries
            PropertyDescriptionPlugin_PagePrefetch = SetorGetArrayValues("Country", numCountries, "")
        End If

        foo = numCountries
        
    Case SVCPStrCVL 'State
        If Doc Is Nothing Then            'Thin Client
            strSQL = SetSQLStatement("AllStates")
        Else                              'Thick Client
            'Set strCountry here as well because if user typed a value for country, then it would not be set above.
            'There's also currently a problem with the document object not knowing what the typed
            'value is until the document is committed
            strCountry = Doc.Properties(2).Value
            strSQL = SetSQLStatement("RelatedStates")
        End If
          
        Set StateCLS = New Database
        numStates = StateCLS.RecordCount
        
        If vkey <> Null Or vkey <> "" Then
            strState = vkey 'Store current State to be used for validation
            'Find the index of the record in the database
            PropertyDescriptionPlugin_PagePrefetch = SetorGetArrayValues("State", numStates, vkey)
        Else
            'Initialize array of States
            PropertyDescriptionPlugin_PagePrefetch = SetorGetArrayValues("State", numStates, "")
        End If

        foo = numStates
    
    Case SVCPStrCVL2    'City
        If Doc Is Nothing Then                'Thin Client
            strSQL = SetSQLStatement("AllCities")
        Else                                  'Thick Client
            strSQL = SetSQLStatement("RelatedCities")
        End If
        
        Set CityCLS = New Database
        numCities = CityCLS.RecordCount
        
        If vkey <> Null Or vkey <> "" Then
            strCity = vkey  'Store current City to be used for validation
            'Find the index of the record in the database
            PropertyDescriptionPlugin_PagePrefetch = SetorGetArrayValues("City", numCities, vkey)
        Else
            'Initialize array of Cities
            PropertyDescriptionPlugin_PagePrefetch = SetorGetArrayValues("City", numCities, "")
        End If
        
        foo = numCities
    
    Case Else
        'PropDesc will be empty on initialize. Do nothing
    End Select
    
    If foo > PageSize Then
        foo = PageSize
    End If
    Exit Function
End Function

Private Function FindPropertyName(ByVal objDocument As IDMObjects.Document, PropLabel As String) As String
'Purpose: Find the internal name of Property
'The property names are going to be different on each server
'Note: Those SVCPs might be labeled differently. So just pass the correct label

    On Error GoTo ErrHandler
    Dim i, intCount As Integer
    intCount = objDocument.Properties.Count 'Number of properties in document
    
    'Go through the list of properties
    'Find the name of the property that matches PropLabel
    For i = 1 To intCount
        If objDocument.Properties.Item(i).Label = PropLabel Then
            FindPropertyName = objDocument.Properties.Item(i).Name
            Exit For
        End If
    Next
    
    'Did not find property
    If FindPropertyName = "" Then
        ErrDesc = "Missing property " & PropLabel & " in the '" & objDocument.ClassDescription.Label & "' Doc class"
        Err.Raise ERR_MISSING_PROPERTY, ERR_APP_SOURCE, ErrDesc
        FindPropertyName = 0
    End If
    Exit Function
    
ErrHandler:
    MsgBox Err.Description, vbCritical + vbOKOnly, Err.Source
End Function


Private Function SetSQLStatement(SQLType As String) As String
    '+++ Purpose: Set the SQL statements to query the database
    
    Select Case SQLType
        Case "Country"
            'Get all countries
            SetSQLStatement = "SELECT Country.CountryID FROM Country " & _
                              "ORDER BY Country.CountryID"
        Case "AllStates"
            'Get all states
            SetSQLStatement = "SELECT State.StateName FROM State " & _
                              "ORDER BY State.StateName"
        Case "RelatedStates"
            'Get states related to selected country
            SetSQLStatement = "SELECT State.StateName FROM Country INNER JOIN State " & _
                              "ON Country.CountryID = State.CountryID " & _
                              "WHERE Country.CountryID LIKE '" & strCountry & _
                              "' ORDER BY State.StateName"
        Case "AllCities"
            'Get all cities
            SetSQLStatement = "SELECT City.CityName FROM City " & _
                              "ORDER BY City.CityName"
        Case "RelatedCities"
            'Get cities related to selected state
            SetSQLStatement = "SELECT City.CityName FROM State INNER JOIN City " & _
                              "ON State.StateID = City.StateID " & _
                              "WHERE State.StateName LIKE '" & strState & _
                              "' ORDER BY City.CityName"
   End Select
End Function

Private Function SetLetterCase(StrWord As String) As String
    'The Countries/States/Cities stored in the arrays are case sensitive
    'If the user types in a value, make sure that the case will correspond to the array value
        
    'For Countries, all the letters should be uppercase
    'For States & Cities, only the first letter of each word is upper case
    'This function will only take care of States & Cities
    
    Dim y As Integer            'Counter
    Dim intLength As Integer    'Length of word to be converted
    Dim strConverted As String  'Converted word to be returned by function
    
    intLength = Len(StrWord)
    
     For y = 1 To intLength
        If y = 1 Then
            'Set First letter to UpperCase
            strConverted = UCase(Left(StrWord, 1))
        Else
            If Mid(StrWord, y, 1) = " " Or Mid(StrWord, y, 1) = "-" Then
                'Set letter after " " or "-" to UpperCase
                strConverted = strConverted & Mid(StrWord, y, 1) & UCase(Mid(StrWord, y + 1, 1))
                y = y + 1 'Add 1 to skip next letter
            Else
                'Set letter to lower case
                strConverted = strConverted & LCase(Mid(StrWord, y, 1))
            End If
        End If
    Next
    
    SetLetterCase = strConverted
    
End Function


Private Sub PropertyDescriptionPlugIn_ValidateValue(ByVal Doc As Object, ByVal PropDesc As Object, ByVal Value As String, pbVal As Boolean)
    '+++ Note for Thin Client users!!!
    '+++ For this Scenario, we will not do validation for the dependent values since we do not have the Doc object
    
    On Error GoTo ErrHandler
    Dim x As Integer    'Counter
    Dim blPass As Boolean 'For Validations
    
    Dim numItems As Integer

    blPass = False
    ErrDesc = ""
    Err.Clear 'Clear any previous error that was raised
    
    Select Case PropDesc
    Case SVCPStr
        Value = UCase(Value)    'Set all letters to uppercase
        blPass = ValidateCountry(Value)
       
    Case SVCPStrCVL
        Value = SetLetterCase(Value)
        If Not Doc Is Nothing Then
            If Doc.Properties(2).Value <> Null Then
                strCountry = Doc.Properties(2).Value
            End If
        Else
            strCountry = "%"    'For Thin Clients, pass a wild card
        End If
        blPass = ValidateState(Value)
   
    Case SVCPStrCVL2
        Value = SetLetterCase(Value)
        If Not Doc Is Nothing Then
            If Doc.Properties(3).Value <> Null Then
                strCity = Doc.Properties(3).Value
            End If
        Else
            strCity = "%"   'For Thin Clients, pass a wild card
        End If
        blPass = ValidateCity(Value)
    End Select

    If blPass = False Then GoTo ErrHandler
    
    Exit Sub

ErrHandler:
    'This will cause the property description to raise an error
    pbVal = False
End Sub

Private Function ValidateCountry(country As String) As Boolean
    Dim x As Integer
    'Check for a valid Country
    For x = 1 To numCountries
        If CountryArray(x) = country Or country = "" Then
            ValidateCountry = True
            Exit Function
        End If
    Next
    'Did not find a match
    ValidateCountry = False
    ErrDesc = country & " is an invalid Country!" & vbCrLf & _
              "Please pick from the list or add it to the database."
End Function

Private Function ValidateState(state As String) As Boolean
    Dim x As Integer
    Dim numItems As Integer
    
    'First, check against current State array that typed/selected value is valid
    For x = 1 To numStates
        If StateArray(x) = state Or state = "" Then
            GoTo Continue
        End If
    Next
    ErrDesc = state & " is an invalid State!" & vbCrLf & _
              "Please pick from the list or add it to the database."
    GoTo Fail
    
Continue:
    'Second, check that this State is a valid dependent value for the selected Country
    strSQL = SetSQLStatement("RelatedStates")
    Dim ValidateStateCLS As Database
    Set ValidateStateCLS = New Database
    numItems = ValidateStateCLS.RecordCount
    'Check for a valid State
    For x = 1 To numItems
        If state = ValidateStateCLS.CorrespondingData Or state = "" Then
            ValidateState = True
            Exit Function
        End If
    Next
    Set ValidateStateCLS = Nothing
    ErrDesc = state & " is not a valid value for the selected country " & strCountry
        
Fail:
    'Did not find a match
    ValidateState = False
End Function

Private Function ValidateCity(city As String) As Boolean
    Dim x As Integer
    Dim numItems As Integer

    'First, check against current City array that typed/selected value is valid
    For x = 1 To numCities
        If CityArray(x) = city Or city = "" Then
            GoTo Continue
        End If
    Next
    
    ErrDesc = city & " is an invalid City!" & vbCrLf & _
              "Please pick from the list or add it to the database."
    GoTo Fail
    
Continue:
    'Second, check that this City is a valid dependent value for the selected State
    strSQL = SetSQLStatement("RelatedCities")
    Dim ValidateCityCLS As Database
    Set ValidateCityCLS = New Database
    numItems = ValidateCityCLS.RecordCount
    
    'Check for a valid State
    For x = 1 To numItems
        If city = ValidateCityCLS.CorrespondingData Or city = "" Then
            ValidateCity = True
            Exit Function
        End If
    Next
    Set ValidateCityCLS = Nothing
    ErrDesc = city & " is not a valid value for the selected state " & strState
     
Fail:
    'Did not find a match
    ValidateCity = False
End Function

