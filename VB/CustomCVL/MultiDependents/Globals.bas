Attribute VB_Name = "Globals"
'Global Variables
Public strCountry As String
Public strState As String
Public strCity As String

Public Const ERR_APP_SOURCE = "Multi-Dependents Plug-In"
Public Const ERR_INVALID = 514      'For invalid values
Public Const ERR_MISSING_PROPERTY = 515     'Missing property on server

'Arrays
Public CountryArray() As String
Public StateArray() As String
Public CityArray() As String

Public strSQL As String

'To store the names of the properties
Public SVCPStr As String
Public SVCPStrCVL As String
Public SVCPStrCVL2  As String

'Record Counts
Public numCountries As Integer
Public numStates As Integer
Public numCities As Integer

