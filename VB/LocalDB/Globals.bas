Attribute VB_Name = "Globals"
' This program is an example which uses the Local DB foundation objects

'----------------------------------------------------------------------
'   Disclaimer: FileNET provides programming examples for illustration
'   only, without warranty either expressed or implied,  including, but
'   not limited to, the implied warranties of merchantability and/or
'   fitness for a particular purpose. This sample assumes that you are
'   familiar with the programming language being demonstrated and the
'   tools used to create and debug procedures.
'----------------------------------------------------------------------

'----------------------------------------------------------------------
' $Revision:   2.0  $
' $Date:   7 November 1999 12:42:54  $
' $Author:   Vladimir Fridman  $
' $Workfile:   Globals.bas  $
'----------------------------------------------------------------------
' All Rights Reserved.  Copyright (c) 1988,1999 FileNet Corp.
'----------------------------------------------------------------------

Option Explicit
Public mbCancelAdd As Boolean


Type SHELLEXECUTEINFO
    cbSize        As Long
    fMask         As Long
    hwnd          As Long
    lpVerb        As String
    lpFile        As String
    lpParameters  As String
    lpDirectory   As String
    nShow         As Long
    hInstApp      As Long
    lpIDList      As Long     'Optional parameter
    lpClass       As String   'Optional parameter
    hkeyClass     As Long     'Optional parameter
    dwHotKey      As Long     'Optional parameter
    hIcon         As Long     'Optional parameter
    hProcess      As Long     'Optional parameter
End Type

Public Const SEE_MASK_INVOKEIDLIST = &HC
Public Const SEE_MASK_NOCLOSEPROCESS = &H40

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Declare Function ShellExecuteEX Lib "shell32.dll" Alias "ShellExecuteEx" _
  (SEI As SHELLEXECUTEINFO) As Long
  
Public Sub ShowFileProperties(sFilename As String, Optional sVerb As String = "properties")
  'open a file properties property page for

   Dim SEI As SHELLEXECUTEINFO

 
  'Fill in the SHELLEXECUTEINFO structure
   With SEI
      .cbSize = Len(SEI)
      .fMask = SEE_MASK_NOCLOSEPROCESS Or SEE_MASK_INVOKEIDLIST
      .hwnd = frmLocRecs.hwnd
      .lpVerb = sVerb
      .lpFile = sFilename
      .lpParameters = vbNullChar
      .lpDirectory = vbNullChar
      .nShow = 5
      .hInstApp = 0
      .lpIDList = 0
   End With
 
  'call the Windows API function to display the property sheet for the file or folder
   ShellExecuteEX SEI
  
End Sub

