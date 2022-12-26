Attribute VB_Name = "WindowsVersion"
Option Explicit
'HOWTO: Determine Which 32-Bit Windows Version Is Being Used
'ID: Q189249
'--------------------------------------------------------------------------------
'The information in this article applies to:
'
'Microsoft Visual Basic Learning, Professional, and Enterprise Editions for Windows, versions 5.0, 6.0
'Microsoft Visual Basic Standard and Professional Editions, 32-bit only, for Windows, version 4.0
'Microsoft Visual Basic for Applications version 5.0
'
'--------------------------------------------------------------------------------
'
'SUMMARY
'An application may need to perform tasks differently depending on which
'operating system is running on the computer. This article shows, by example,
'how to differentiate between Windows 95, Windows 98, Window NT 3.51, and
'Windows NT 4.0.
'
'The Win32 GetVersionEx function returns information that a program can use
'to identify the operating system. Among those values are the major and
'minor revision numbers and a platform identifier. With the introduction of
'Windows 98, it now takes a more involved logical evaluation to determine
'which version of Windows is in use. The listing below provides the data
'needed to evaluate the OSVERSIONINFO structure populated by GetVersionEx:
'
'                   Win95     Win98     WinNT 3.51     WinNT 4.0
'                  ------------------------------------------------
'dwPlatFormID         1         1            2              2
'
'dwMajorVersion       4         4            3              4
'
'dwMinorVersion       0        10           51              0
'
'============================================================================

Private Const PROCESSOR_INTEL_386 = 386
Private Const PROCESSOR_INTEL_486 = 486
Private Const PROCESSOR_INTEL_PENTIUM = 586
Private Const PROCESSOR_MIPS_R4000 = 4000
Private Const PROCESSOR_ALPHA_21064 = 21064

Public gApplicationVersion As String
Public gBuildVersion As String
Public gFileVersion As String

'================================GetBuildInfo=====================================
Private Type VS_FIXEDFILEINFO
   dwSignature As Long
   dwStrucVersion As Long         'e.g. 0x00000042 = "0.42"
   dwFileVersionMS As Long        'e.g. 0x00030075 = "3.75"
   dwFileVersionLS As Long        'e.g. 0x00000031 = "0.31"
   dwProductVersionMS As Long     'e.g. 0x00030010 = "3.10"
   dwProductVersionLS As Long     'e.g. 0x00000031 = "0.31"
   dwFileFlagsMask As Long        '= 0x3F for version "0.42"
   dwFileFlags As Long            'e.g. VFF_DEBUG Or VFF_PRERELEASE
   dwFileOS As Long               'e.g. VOS_DOS_WINDOWS16
   dwFileType As Long             'e.g. VFT_DRIVER
   dwFileSubtype As Long          'e.g. VFT2_DRV_KEYBOARD
   dwFileDateMS As Long           'e.g. 0
   dwFileDateLS As Long           'e.g. 0
End Type

Private Declare Function GetFileVersionInfoSize Lib "Version.dll" _
   Alias "GetFileVersionInfoSizeA" _
  (ByVal lptstrFilename As String, _
   lpdwHandle As Long) As Long

Private Declare Function GetFileVersionInfo Lib "Version.dll" _
   Alias "GetFileVersionInfoA" _
  (ByVal lptstrFilename As String, _
   ByVal dwHandle As Long, _
   ByVal dwLen As Long, _
   lpData As Any) As Long
   
Private Declare Function VerQueryValue Lib "Version.dll" _
   Alias "VerQueryValueA" _
  (pBlock As Any, _
   ByVal lpSubBlock As String, _
   lplpBuffer As Any, nVerSize As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" _
   Alias "RtlMoveMemory" _
  (Destination As Any, _
   Source As Any, _
   ByVal Length As Long)
  
Private Declare Function lstrcpyA Lib "kernel32" _
  (ByVal RetVal As String, ByVal Ptr As Long) As Long
                        
Private Declare Function lstrlenA Lib "kernel32" _
  (ByVal Ptr As Any) As Long
  
'================================ End - GetBuildInfo ===============================

'================================ Begin - GetWinVersion ============================
Private Declare Function GetVersionExA Lib "kernel32" _
   (LpVersionInformation As OSVERSIONINFO) As Integer
Private Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As _
   MEMORYSTATUS)
Private Declare Sub GetSystemInfo Lib "kernel32" (lpSystemInfo As _
   SYSTEM_INFO)

Private Type OSVERSIONINFO
   dwOSVersionInfoSize As Long
   dwMajorVersion As Long
   dwMinorVersion As Long
   dwBuildNumber As Long
   dwPlatformId As Long
   szCSDVersion As String * 128
End Type

Type SYSTEM_INFO
      dwOemID As Long
      dwPageSize As Long
      lpMinimumApplicationAddress As Long
      lpMaximumApplicationAddress As Long
      dwActiveProcessorMask As Long
      dwNumberOrfProcessors As Long
      dwProcessorType As Long
      dwAllocationGranularity As Long
      dwReserved As Long
End Type
'================================ End - GetWinVersion ============================

Type MEMORYSTATUS
      dwLength As Long
      dwMemoryLoad As Long
      dwTotalPhys As Long
      dwAvailPhys As Long
      dwTotalPageFile As Long
      dwAvailPageFile As Long
      dwTotalVirtual As Long
      dwAvailVirtual As Long
End Type

Public Function GetWinVersion() As String
  Dim osinfo As OSVERSIONINFO
  Dim retvalue As Integer
  Dim sVersion As String
  Dim BuildNo As String
  Dim pos As Integer
  Dim Servicepack As String
  
  GetWinVersion = ""
  osinfo.dwOSVersionInfoSize = 148
  osinfo.szCSDVersion = Space$(128)
  retvalue = GetVersionExA(osinfo)

  With osinfo
    Select Case .dwPlatformId
       Case 1
          If .dwMinorVersion = 0 Then
             sVersion = "Windows 95"
          ElseIf .dwMinorVersion = 10 Then
             sVersion = "Windows 98"
          End If
       Case 2
          If .dwMajorVersion = 3 Then
             sVersion = "Windows NT 3.51"
          ElseIf .dwMajorVersion = 4 Then
             sVersion = "Windows NT 4.0"
          End If
       Case Else
          sVersion = "Unknown version"
    End Select
    
    BuildNo = (osinfo.dwBuildNumber And &HFFFF)
    
    'Any additional info. In Win9x, this can be
    '"any arbitrary string" provided by the
    'manufacturer. In NT, this is the service pack.
    
     pos = InStr(osinfo.szCSDVersion, Chr$(0))
     
     If pos Then
        Servicepack = Left$(osinfo.szCSDVersion, pos - 1)
     End If
     
     GetWinVersion = sVersion & " " & vbCrLf & _
                     "Build: " & BuildNo & " (" & _
                     Servicepack & ")"
  End With
 
End Function
 
Public Function CPUInfo() As String
 
  'Get CPU type and operating mode.
  Dim sysinfo As SYSTEM_INFO
  Dim msg As String
  
  GetSystemInfo sysinfo
  msg = msg & "CPU: "
  
  Select Case sysinfo.dwProcessorType
    Case PROCESSOR_INTEL_386
        msg = msg & "Intel 386" & vbCrLf
    Case PROCESSOR_INTEL_486
        msg = msg & "Intel 486" & vbCrLf
    Case PROCESSOR_INTEL_PENTIUM
        msg = msg & "Intel Pentium" & vbCrLf
    Case PROCESSOR_MIPS_R4000
        msg = msg & "MIPS R4000" & vbCrLf
    Case PROCESSOR_ALPHA_21064
        msg = msg & "DEC Alpha 21064" & vbCrLf
    Case Else
        msg = msg & "(unknown)" & vbCrLf
  End Select

  msg = msg & vbCrLf
' Get free memory.
  Dim memsts As MEMORYSTATUS
  Dim Memory As Long
  
  GlobalMemoryStatus memsts
  Memory = memsts.dwTotalPhys
  msg = msg & "Total Physical Memory: "
  msg = msg & Format$(Memory \ 1024, "###,###,###") & "K" _
            & vbCrLf
  
  Memory& = memsts.dwAvailPhys
  msg = msg & "Available Physical Memory: "
  msg = msg & Format$(Memory \ 1024, "###,###,###") & "K" _
            & vbCrLf
  
  Memory& = memsts.dwTotalVirtual
  msg = msg & "Total Virtual Memory: "
  msg = msg & Format$(Memory \ 1024, "###,###,###") & "K" _
            & vbCrLf
  
  Memory& = memsts.dwAvailVirtual
  msg = msg & "Available Virtual Memory: "
  msg = msg & Format$(Memory \ 1024, "###,###,###") & "K" _
            & vbCrLf & vbCrLf
  
  CPUInfo = msg
End Function

Public Function GetBuildInfo(sSourceFile As String) As String

  Dim FI As VS_FIXEDFILEINFO
  Dim sBuffer() As Byte
  Dim nBufferSize As Long
  Dim lpBuffer As Long
  Dim nVerSize As Long
  Dim nUnused As Long
  Dim tmpVer As String
  Dim sComments As String
  Dim sFileVersion As String
  Dim sProductVersion As String
   
  If sSourceFile > "" Then
    '+++ Set file that has the encryption level
    '+++ info and call to get required size
    nBufferSize = GetFileVersionInfoSize(sSourceFile, nUnused)
     
    ReDim sBuffer(nBufferSize)
  
    If nBufferSize > 0 Then
      
      'get the version info
      Call GetFileVersionInfo(sSourceFile, 0&, nBufferSize, sBuffer(0))
      Call VerQueryValue(sBuffer(0), "\", lpBuffer, nVerSize)
      Call CopyMemory(FI, ByVal lpBuffer, Len(FI))
   
      If VerQueryValue(sBuffer(0), "\VarFileInfo\Translation", lpBuffer, nVerSize) Then
        If nVerSize Then
          tmpVer = GetPointerToString(lpBuffer, nVerSize)
          tmpVer = Right("0" & Hex(Asc(Mid(tmpVer, 2, 1))), 2) & _
                   Right("0" & Hex(Asc(Mid(tmpVer, 1, 1))), 2) & _
                   Right("0" & Hex(Asc(Mid(tmpVer, 4, 1))), 2) & _
                   Right("0" & Hex(Asc(Mid(tmpVer, 3, 1))), 2)
                   
          sComments = "\StringFileInfo\" & tmpVer & "\Comments"
          sFileVersion = "\StringFileInfo\" & tmpVer & "\FileVersion"
          sProductVersion = "\StringFileInfo\" & tmpVer & "\ProductVersion"
          
          '+++ Get predefined "File Version" resources
         
          If VerQueryValue(sBuffer(0), sFileVersion, lpBuffer, nVerSize) Then
            If nVerSize Then
              '+++ get the file description
              gFileVersion = GetStrFromPtrA(lpBuffer)
            End If
          End If
          
          '+++ Get predefined "Product Version" resources
          If VerQueryValue(sBuffer(0), sProductVersion, lpBuffer, nVerSize) Then
            If nVerSize Then
              '+++ get the file description
              gApplicationVersion = GetStrFromPtrA(lpBuffer)
              GetBuildInfo = "Product Version: " & GetStrFromPtrA(lpBuffer)
            End If
          End If
               
          '+++ Get predefined "Comments" resources
          If VerQueryValue(sBuffer(0), sComments, lpBuffer, nVerSize) Then
            If nVerSize Then
              '+++ get the file description
              gBuildVersion = GetStrFromPtrA(lpBuffer)
              GetBuildInfo = GetBuildInfo & Chr(13) & _
                             "Build: " & GetStrFromPtrA(lpBuffer)
            End If  '+++ If nVerSize
          End If  '+++ If VerQueryValue
        End If  '+++ If nVerSize
      End If  '+++ If VerQueryValue
    End If  '+++ If nBufferSize
  End If  '+++ If sSourcefile

End Function

Private Function GetStrFromPtrA(ByVal lpszA As Long) As String
  GetStrFromPtrA = String$(lstrlenA(ByVal lpszA), 0)
  Call lstrcpyA(ByVal GetStrFromPtrA, ByVal lpszA)
End Function

Private Function GetPointerToString(lpString As Long, nBytes As Long) As String
  Dim Buffer As String
  
  If nBytes Then
    Buffer = Space$(nBytes)
    CopyMemory ByVal Buffer, ByVal lpString, nBytes
    GetPointerToString = Buffer
  End If
End Function

 
