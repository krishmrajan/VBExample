Attribute VB_Name = "modUtility"
' This program is an example of how to use the new foundation objects
' for compound documents

'----------------------------------------------------------------------
'   Disclaimer: FileNET provides programming examples for illustration
'   only, without warranty either expressed or implied,  including, but
'   not limited to, the implied warranties of merchantability and/or
'   fitness for a particular purpose. This sample assumes that you are
'   familiar with the programming language being demonstrated and the
'   tools used to create and debug procedures.
'----------------------------------------------------------------------

'----------------------------------------------------------------------
' $Revision:     $
' $Date:     $
' $Author:     $
' $Workfile:     $
'----------------------------------------------------------------------
' All Rights Reserved.  Copyright (c) 1988,1999 FileNet Corp.
'----------------------------------------------------------------------

Option Explicit

Sub main()
    MainForm.Show vbModeless
End Sub

Public Sub CenterForm(oForm As Form)
    oForm.Left = (Screen.Width - oForm.Width) / 2
    oForm.Top = (Screen.Height - oForm.Height) / 2
End Sub

Public Function GetTitle(vTitle As Variant) As Variant
    Dim iPos As Integer
    iPos = InStr(vTitle, ".")
    If iPos = 0 Then
       GetTitle = vTitle
    Else
       GetTitle = Left(vTitle, (iPos - 1))
    End If
End Function

Public Sub GetDirectory(sFilePath As String, sDirectory As String, sFileName As String)
   Dim lIndex As Long
   For lIndex = Len(sFilePath) To 1 Step -1
        If Mid$(sFilePath, lIndex, 1) = "\" Then
            sFileName = Right$(sFilePath, Len(sFilePath) - lIndex)
            sDirectory = Left$(sFilePath, lIndex - 1) ' Remove last Slash character
            Exit For
        End If
    Next lIndex

End Sub

