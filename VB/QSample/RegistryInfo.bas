Attribute VB_Name = "RegistryInfo"
Option Explicit

Private Const HKEY_CLASSES_ROOT As Long = &H80000000
Private Const HKEY_CURRENT_USER As Long = &H80000001
Private Const HKEY_LOCAL_MACHINE As Long = &H80000002
Private Const HKEY_USERS As Long = &H80000003
Private Const REG_SZ As Long = &H1   'null terminated string
Private Const KEY_ALL_ACCESS As Long = &H3F
Private Const ERROR_SUCCESS = 0

Private Type SECURITY_ATTRIBUTES
  nLength  As Long
  lpSecurityDescriptor As Long
  bInheritHandle  As Long
End Type

Private Declare Function RegOpenKeyEx _
    Lib "advapi32.dll" Alias "RegOpenKeyExA" _
   (ByVal hKey As Long, _
    ByVal lpSubKey As String, _
    ByVal ulOptions As Long, _
    ByVal samDesired As Long, _
    phkResult As Long) As Long

Private Declare Function RegCloseKey Lib "advapi32.dll" _
   (ByVal hKey As Long) As Long
   
Private Declare Function RegQueryValueEx Lib "advapi32.dll" _
   Alias "RegQueryValueExA" _
  (ByVal hKey As Long, ByVal lpszValueName As String, _
   ByVal lpdwRes As Long, lpType As Long, _
   lpData As Any, nSize As Long) As Long

Public Function GetIDMInstallPath() As String
   Dim hKey As Long
   Dim dwKeyType As Long

   Dim sKeyName As String
   Dim sValue As String
   
   Dim dwDataType As Long
   Dim dwDataSize As Long
  
   '+++ 1. open key
   dwKeyType = HKEY_LOCAL_MACHINE
   sKeyName = "Software\FileNET\IDM\Install"
   sValue = "InstallPath"
   
   hKey = RegKeyOpen(HKEY_LOCAL_MACHINE, sKeyName)
   
   If hKey <> 0 Then
   
      Debug.Print "RegKeyOpen = "; hKey
      
      '+++ 2. Determine the size and type of data to be read.
      '+++    In this case it should be a string (REG_SZ) value.
      dwDataSize = RegGetStringSize(ByVal hKey, sValue, dwDataType)
      
      Debug.Print "RegGetStringSize = "; dwDataSize
      
      If dwDataSize > 0 Then
      
         '+++ 3. get the value for that key
         Dim sdataret As String
         
         sdataret = RegGetStringValue(hKey, sValue, dwDataSize)
         
        'if a value returned
         If sdataret > "" Then
            GetIDMInstallPath = sdataret
            Debug.Print "RegGetStringValue = "; sdataret
         End If
      End If
   End If
   
   Call RegCloseKey(hKey)
      
End Function

Private Function RegKeyOpen(dwKeyType As Long, sKeyPath As String) As Long

   Dim hKey As Long
   Dim dwOptions As Long
   Dim SA As SECURITY_ATTRIBUTES
   
   SA.nLength = Len(SA)
   SA.bInheritHandle = False

   dwOptions = 0&
   If RegOpenKeyEx(dwKeyType, _
                   sKeyPath, dwOptions, _
                   KEY_ALL_ACCESS, hKey) = ERROR_SUCCESS Then
   
      RegKeyOpen = hKey
   End If

End Function


Private Function RegGetStringSize(ByVal hKey As Long, _
                                  ByVal sValue As String, _
                                  dwDataType As Long) As Long
                                  
   Dim success As Long
   Dim dwDataSize As Long
   
   success = RegQueryValueEx(hKey, _
                             sValue, _
                             0&, _
                             dwDataType, _
                             ByVal 0&, _
                             dwDataSize)
         
   If success = ERROR_SUCCESS Then
      If dwDataType = REG_SZ Then
      
         RegGetStringSize = dwDataSize
         
      End If
   End If

End Function

Private Function RegGetStringValue(ByVal hKey As Long, _
                                   ByVal sValue As String, _
                                   dwDataSize As Long) As String

   Dim sdataret As String
   Dim dwDataRet As Long
   Dim success As Long
   Dim pos As Long
   
  'get the value of the passed key
   sdataret = Space$(dwDataSize)
   dwDataRet = Len(sdataret)
   
   success = RegQueryValueEx(hKey, sValue, _
                             ByVal 0&, dwDataSize, _
                             ByVal sdataret, dwDataRet)

   If success = ERROR_SUCCESS Then
      If dwDataRet > 0 Then
      
         pos = InStr(sdataret, Chr$(0))
         RegGetStringValue = Left$(sdataret, pos - 1)
         
      End If
   End If
   
End Function

