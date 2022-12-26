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
'This ActiveX dll works both for the Thick and Thin Clients
'Note that on Thin Client, the Doc document object will be null

'Requirements
'   On Server, using the Admin Tool:
'              1. Create a CVL Plugin named Whatever. For its data, add SimplePlugIn.Geography
'                 SimplePlugIn.Geography is this project's name and the class that implements IDMObjects.PropertyDescriptionPlugin
'              2. Create a Doc Class as follows:
'
'                       Doc Class Name = MyDocClass
'
'                       PropertyName | CVL Name
'                       ------------ | ---------
'             (Optional)Title        |
'                       SVCP Str CVL    |  <PI>Whatever<\PI>
'
'Note: "SVCP Str CVL" is just a label here. It represents a
'       Single value property that has a CVL assigned to them.
'This dll only needs to be registered on the web server to test this on a thin client
'Country.mdb must exist in the same location where this dll is registered


'Plugin Variables/Objects
Implements IDMObjects.PropertyDescriptionPlugIn
Option Explicit

'Conditional Compilation Constants
#Const USE_DEFAULT = True

Dim ErrDesc As String
Dim foo As Integer
 
'Datasource Class Object
Dim CountryCLS As Country


Private Sub Class_Terminate()
    Set CountryCLS = Nothing
End Sub

Private Property Get PropertyDescriptionPlugin_DefaultValue(ByVal Doc As Object, ByVal PropDesc As Object) As Variant
    'This will work only if HasDefaultValue = True
    PropertyDescriptionPlugin_DefaultValue = "United States"
End Property

Private Property Get PropertyDescriptionPlugin_ErrorDescription(ByVal Doc As Object, ByVal PropDesc As Object) As String
    PropertyDescriptionPlugin_ErrorDescription = ErrDesc
End Property

Private Property Get PropertyDescriptionPlugin_FetchCount(ByVal Doc As Object, ByVal PropDesc As Object) As Long
    PropertyDescriptionPlugin_FetchCount = foo
End Property

Private Function PropertyDescriptionPlugin_FormatValue(ByVal Doc As Object, ByVal PropDesc As Object, ByVal ValueIn As Variant) As String
   PropertyDescriptionPlugin_FormatValue = ValueIn
   'ValueIn is case sensitive - we won't be converting case in this sample
End Function

Private Property Get PropertyDescriptionPlugin_HasDefaultValue(ByVal Doc As Object, ByVal PropDesc As Object) As Boolean
    'In Conditional Compilation Arguments
    '   USE_DEFAULT = True
    '   USE_DEFAULT = False (Don't use a default)
    
    #If USE_DEFAULT Then
        PropertyDescriptionPlugin_HasDefaultValue = True
    #Else
        PropertyDescriptionPlugin_HasDefaultValue = False
    #End If
End Property

Private Property Get PropertyDescriptionPlugin_Item(ByVal Doc As Object, ByVal PropDesc As Object, ByVal Index As Long) As String
    On Error Resume Next    'To ignore invalid indexes
    PropertyDescriptionPlugin_Item = CountryArray(Index)
End Property

Private Function PropertyDescriptionPlugin_PagePrefetch(ByVal Doc As Object, ByVal PropDesc As Object, ByVal vKey As Variant, ByVal KeyType As IDMObjects.idmPropPlugKeyType, ByVal PageSize As Long) As Long

    On Error Resume Next
    Dim x As Integer

    PropertyDescriptionPlugin_PagePrefetch = 1
   

    Set CountryCLS = New Country
    numCountries = CountryCLS.RecordCount

    'Create array of countries
    ReDim CountryArray(1 To numCountries)

    For x = 1 To numCountries
        CountryArray(x) = CountryCLS.CorrespondingCountry

        If vKey = CountryArray(x) Then

            'Return the index of vkey

            PropertyDescriptionPlugin_PagePrefetch = x

         '   Exit For

        End If
    Next

   'foo is the size of the returned page
   If numCountries <= PageSize Then

    foo = PageSize    'No Paging

   Else

        If KeyType = idmPropPlugKeyStart Then 'Forward Display

            If vKey = Null Then

                foo = PageSize

            ElseIf PropertyDescriptionPlugin_PagePrefetch = 1 Then  ' startkey is null

                foo = PageSize

                PropertyDescriptionPlugin_PagePrefetch = 1

            ElseIf (PropertyDescriptionPlugin_PagePrefetch + PageSize) <= numCountries Then

                foo = PageSize

            Else

                foo = numCountries - PropertyDescriptionPlugin_PagePrefetch + 1

            End If

        Else ' backward display

             foo = PageSize

             If (PropertyDescriptionPlugin_PagePrefetch - PageSize + 1) >= 0 Then

              PropertyDescriptionPlugin_PagePrefetch = PropertyDescriptionPlugin_PagePrefetch - PageSize

             Else

                PropertyDescriptionPlugin_PagePrefetch = 1

                foo = PageSize - 1

             End If

        End If
    End If

   Exit Function

'Message boxes have potential to lock up the web server. Don't use them here.

End Function

 
Private Sub PropertyDescriptionPlugIn_ValidateValue(ByVal Doc As Object, ByVal PropDesc As Object, ByVal Value As String, pbVal As Boolean)
    On Error GoTo ErrHandler
    Dim x As Integer
    Dim blPass As Boolean
    Dim numItems As Integer
    
      
    blPass = False
    'Clear any previous error
    ErrDesc = ""
    Err.Clear
    
    If numCountries = 0 Then
        Set CountryCLS = New Country
        numCountries = CountryCLS.RecordCount
        
        'Create array of countries
        ReDim CountryArray(1 To numCountries)
        For x = 1 To numCountries
            CountryArray(x) = CountryCLS.CorrespondingCountry
        Next
    End If
        
    'Check for a valid Country
    For x = 1 To numCountries
        If CountryArray(x) = Value Or Value = "" Then   'Allow empty string
            blPass = True
            pbVal = True
            Exit For
        End If
    Next
    
    If blPass = False Then
        GoTo ErrHandler
    End If
    
    Exit Sub
    
ErrHandler:
   ErrDesc = "Invalid Country, please pick from the list"
   pbVal = False
   
End Sub
