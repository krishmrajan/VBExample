VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CIDMMacroAPI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Enum eIDMReturnStatus
    CIDMOk = 0
    CIDMError = 1
    CIDMDriveSelection = 2
    CIDMSelection = 3
    CIDMCancel = 4
End Enum

Public Enum eIDMFlags
    CIDMCopy = 1
    CIDMReference = 2
    CIDMNoOpenFromDrive = 4
    CIDMTrackedFiles = 8
    CIDMNoSaveAsDialog = 16
    CIDMInsert = 32
    CIDMShortCut = 64
End Enum

Public Function IDMFileOpen(iApplType As Integer, strPathName As String, Optional oAppl As Variant, Optional efilter As Variant, Optional eFlag = 0, Optional vPathNames As Variant) As Integer

End Function

Public Function IDMFileCheckin(iApplType As Integer, strPathName As String, Optional oAppl As Object) As Long 'eIDMReturnStatus

End Function

Public Function IDMCancelCheckout(iApplType As Integer, strPathName As String, Optional oAppl As Object) As Long ' eIDMReturnStatus

End Function

Public Function IDMFileAdd(iApplType As Integer, strPathName As String, Optional oAppl As Variant) As Integer

End Function

Public Function IDMSave(iApplType As Integer, strPathName As String, Optional oAppl As Object) As Long 'eIDMReturnStatus

End Function

Public Function IDMShowProperties(iApplType As Integer, strPathName As String, Optional oAppl As Object) As Long 'eIDMReturnStatus

End Function

Public Function IDMInsertMezzProperties(iApplType As Integer, oAppl As Object) As Long 'eIDMReturnStatus

End Function

Public Function IDMUpdateMezzProperties(iApplType As Integer, oAppl As Object) As Long 'eIDMReturnStatus

End Function
Public Function IDMGetMezzPropertyValue(iApplType As Integer, oAppl As Object, strPropertyName As String, strPropertyValue As String) As Long 'eIDMReturnStatus

End Function

Public Function IDMShowPreferences(iApplType As Integer, oAppl As Variant) As Integer

End Function
Public Function IDMGetPreferenceValue(strPreference As String, iApplType As Integer, varPreferenceValue As Variant) As Integer

End Function
Public Function IDMStart(iApplType As Integer, oAppl As Object) As Long 'eIDMReturnStatus

End Function

Public Function IDMMenu(iApplType As Integer, oAppl As Object) As Long 'eIDMReturnStatus

End Function

Public Function IDMUpdateMenu(iApplType As Integer, oAppl As Object, Optional vDocClose As Variant) As Long 'eIDMReturnStatus

End Function

Public Function IDMSaveDialog(iApplType As Integer, strPathName As String, Optional oAppl As Object) As Long 'eIDMReturnStatus

End Function

Public Function IDMGetResource(strResourceName As String) As String

End Function

Public Function IDMDocStatus(strPathName As String) As Integer

End Function
Public Function IDMGetHelp(iApplType As Integer) As Integer

End Function

