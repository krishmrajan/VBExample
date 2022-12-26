Attribute VB_Name = "Module1"
Option Explicit
Public Enum CommitValues
    UnDecided = -1
    DontCommit = 0
    Commit = 1
End Enum
Public Type DocInfo
    CommitFlag As CommitValues
    FileName As String
End Type

Public Const InitArraySz = 50    ' Initial size of DocList, FinalList
Public ArraySz As Integer        ' Current array sizes
Public DocList() As DocInfo      ' Primary array of selected filenames for browsing
Public FolderList() As String    ' Optional folder for each document
Public FinalList() As IDMObjects.Document  ' Final array of FN DocObjects to be committed
Public TotalDocs As Integer      ' Number of docs in DocList
' Use form variables to insulate code from form name changes
Public MainForm As New FormMain
Public PropertyForm As New FormProperty
Public CommitForm As New FormProg
Public oErrManager As idmError.ErrorManager
Public HomeDirectory As String    ' Home directory for finding icons
Public CurrentLib As IDMObjects.Library   ' Current library in use

Public CurrentDocInx As Integer   ' Current offset into DocList
Public Online As Boolean          ' Allows offline debugging

Public Sub ShowError()
Dim oErrCollect As idmError.Errors
Dim oError As idmError.Error
Dim iCnt As Integer
Set oErrCollect = oErrManager.Errors
If oErrCollect.Count > 1 Then
    iCnt = 1
    For Each oError In oErrCollect
        MsgBox "Error " & iCnt & ": " & oError.Description & " : " & Hex(oError.Number)
        iCnt = iCnt + 1
    Next
Else
    If oErrCollect.Count = 1 Then
        oErrManager.ShowErrorDialog
    Else
        If Err.Number <> 0 Then
            MsgBox Err.Description & " : " & Err.Number
        End If
    End If
End If
End Sub


Public Function FormatDataType(ByVal TypeID As idmTypeID)
    Select Case TypeID
        Case idmTypeArray
            FormatDataType = "Array"
        Case idmTypeBoolean
            FormatDataType = "Boolean"
        Case idmTypeByte
            FormatDataType = "Byte"
        Case idmTypeCharacter
            FormatDataType = "Character"
        Case idmTypeCurrency
            FormatDataType = "Currency"
        Case idmTypeDate
            FormatDataType = "Date"
        Case idmTypeDouble
            FormatDataType = "Double"
        Case idmTypeEmpty
            FormatDataType = "Empty"
        Case idmTypeError
            FormatDataType = "Error"
        Case idmTypeGuid
            FormatDataType = "GUID"
        Case idmTypeLong
            FormatDataType = "Long"
        Case idmTypeNull
            FormatDataType = "NULL"
        Case idmTypeObject
            FormatDataType = "Object"
        Case idmTypeShort
            FormatDataType = "Short"
        Case idmTypeSingle
            FormatDataType = "Single"
        Case idmTypeString
            FormatDataType = "String"
        Case idmTypeUnknown
            FormatDataType = "Unknown"
        Case idmTypeUnsignedLong
            FormatDataType = "Unsigned Long"
        Case idmTypeUnsignedShort
            FormatDataType = "Unsigned Short"
        Case idmTypeVariant
            FormatDataType = "Variant"
    End Select
End Function


Sub Main()
HomeDirectory = App.Path
' These arrays will be resized after files are loaded...
ArraySz = InitArraySz
ReDim DocList(ArraySz)
ReDim FinalList(ArraySz)
ReDim FolderList(ArraySz)

Load MainForm
If TotalDocs > 0 Then
    MainForm.Show
Else
    Unload MainForm
End If
End Sub
