Attribute VB_Name = "Globals"
'Global Variables
Public strCountry As String

Public Const ERR_APP_SOURCE = "Simple CVL Plug-In"
Public Const ERR_INVALID = 514      'For invalid values
Public Const ERR_MISSING_PROPERTY = 515     'Missing property on server

'Putting these here for now. THere's a problem on thin clients where the
'variables would be cleared if declared under option explicit
Public CountryArray() As String 'Array to keep current list of countries
Public numCountries As Integer  'Record Counts
