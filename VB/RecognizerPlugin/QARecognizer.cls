VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TextRecognizer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Option Base 1
Implements IDMObjects.RecognizerPlugIn
Private sFilePath As String
Private Const DeLimiter = "|"


'''''''''''''' Values used for negative testing by QA
Private Const INVALID_ABS_PATH = "Invalid Absolute Path"
Private Const INVALID_REL_PATH = "Invalid Relative Path"
Private Const EMPTY_REL_PATH = ""
Private Const EMPTY_ABS_PATH = ""

Private Enum AbsPathSetting
    AbsValidPath = 1
    AbsInvalidPath = 2
    AbsEmptyPath = 3
End Enum
Private Enum RelPathSetting
    RelValidPath = 1
    RelInvalidPath = 2
    RelEmptyPath = 3
End Enum
Private Sub RecognizerPlugIn_CanEditLink(ByVal Link As Object, CanEdit As Boolean)
    CanEdit = True
End Sub

Private Sub RecognizerPlugIn_ChangeLink(ByVal Link As Object, ByVal NewPath As String, ByVal bRelativePath As Boolean)
    
    ' In order to edit what could be a single path within a list of paths (parent has more than one child)
    ' we need to store the entire contents of the file line by line in an array. We will then search for
    ' the desired link within the string array, replace the old path with the new, then write the list back to the file
    ' over-writing the original list
    
    Dim PathLengthInBytes As Long
    Dim Path As String
    Dim OldAbsolutePathToChild As String
    Dim NewAbsolutePathToChild As String
    Dim Temp As String
    Dim Lines() As String
    Dim FileNum As Integer
    Dim fsFile As New Scripting.FileSystemObject
    Dim fsTextStream As Scripting.TextStream
 
    
    '''''''''''If Relative Path is passed then we need to construct a new aboslute path to the child
    If bRelativePath Then
        Dim ParentDirectory As String
        Dim ParentFilename As String
        
        ParsePath sFilePath, ParentFilename, ParentDirectory
        If Left(NewPath, 2) = ".\" Then  ' just tack it on to the end of the parent dir
            NewAbsolutePathToChild = ParentDirectory & Right(NewPath, Len(NewPath) - 2)
            
        ElseIf Left(NewPath, 3) = "..\" Then 'child is above parent - need to construct
            
            Dim TempNewPath
            
            TempNewPath = NewPath
            ''''''''''''''Chop up the parent directory then add the new path
            Do
                If StrComp("..\", Left(TempNewPath, 3), vbTextCompare) = 0 Then
                    TempNewPath = Right(TempNewPath, Len(TempNewPath) - 3)
                    Dim Index As Integer
                                        
                    For Index = Len(ParentDirectory) To 1 Step -1
                       If Mid$(ParentDirectory, Index, 1) = "\" Then
                             ParentDirectory = Left$(ParentDirectory, Index - 1)
                             Exit For
                       End If
                    Next Index
                Else
                    Exit Do
                End If
            Loop
            
            OldAbsolutePathToChild = Link.Properties("idmLinkAbsolutePath").Value
            NewAbsolutePathToChild = ParentDirectory & "\" & TempNewPath
            Link.Properties("idmLinkRelativePath").Value = NewPath
        End If
    Else
        OldAbsolutePathToChild = Link.Properties("idmLinkAbsolutePath").Value
        NewAbsolutePathToChild = NewPath
        Link.Properties("idmLinkAbsolutePath") = NewPath
    End If
    
    ''''''''''''Read Current List Of Aboslute Paths From Parent
    Set fsTextStream = fsFile.OpenTextFile(sFilePath, ForReading, False)
        
    ReDim Lines(1)
    
    Do While fsTextStream.AtEndOfStream <> True
        Lines(UBound(Lines)) = fsTextStream.ReadLine
        ReDim Preserve Lines(UBound(Lines) + 1)
    Loop
    
    ReDim Preserve Lines(UBound(Lines) - 1)
    fsTextStream.Close
    
    Dim i As Integer
    Dim bLinkPropSettingsExist As Boolean
    Dim sLinkProps As String
    
    For i = 1 To UBound(Lines)
        
        If InStr(1, Lines(i), "|", vbTextCompare) = 0 Then
                Temp = Lines(i)
                sLinkProps = ""
        Else
                Temp = Left(Lines(i), InStr(1, Lines(i), "|", vbTextCompare) - 1)
                sLinkProps = Right(Lines(i), Len(Lines(i)) - Len(Temp))
        End If
               
        If StrComp(Temp, OldAbsolutePathToChild, vbTextCompare) = 0 Then
              Temp = NewAbsolutePathToChild
              If sLinkProps <> "" Then
                 Lines(i) = Temp & sLinkProps
              Else
                Lines(i) = Temp
              End If
              
          Exit For
        End If
    Next i
    
    ' Write contents of array back to file
 
    Set fsTextStream = fsFile.OpenTextFile(sFilePath, ForWriting, False)
    
    For i = 1 To UBound(Lines)
        fsTextStream.WriteLine Lines(i)
    Next
    
    fsTextStream.Close
    Set fsTextStream = Nothing
    Set fsFile = Nothing
    
End Sub

Private Sub RecognizerPlugIn_CompoundBehaviorID(BehaviorID As String)
    
    BehaviorID = "IDMObjects.OLEBehavior"
    
End Sub

Private Sub RecognizerPlugIn_DetermineComponents(ByVal Recognizer As Object, ByVal Links As Object)
        
    Dim fsFile As New Scripting.FileSystemObject
    Dim fsTextStream As Scripting.TextStream
    Dim sCurrentLine As String
    Dim oLink As IDMObjects.Link
    Dim oLinks As IDMObjects.Links
    Dim AbsPathSelector As AbsPathSetting
    Dim RelPathSelector As RelPathSetting
    
    Dim sLinkAbsolutePath As String
    Dim sLinkRelativePath As String
    Dim sLinkStrength As String
    Dim sClassID As String
    Dim sChildUpdateMode As String
    Dim sLinkUsesRelative As String
    Dim sTemp As String
    
    Dim orecognizer As IDMObjects.CDRecognizer
    Set orecognizer = Recognizer
    
    Set oLinks = Links
       
    ' Open the parent to be parsed line-by-line for paths to children
    Set fsTextStream = fsFile.OpenTextFile(sFilePath, ForReading, False)
    
    
    ' Read each line within the parent and create links with paths retrieved from parent
    Do While fsTextStream.AtEndOfStream <> True
                          
        sCurrentLine = fsTextStream.ReadLine
        
        If Right(Left(sCurrentLine, 2), 1) = ":" Then    ' second character is a colon...we assume we have a path and not just random content
            
            Set oLink = New IDMObjects.Link
            Set oLink = orecognizer.CreateLink
            oLink.RecognizerID = "QA.TextRecognizer"
         
            Dim iCount As Integer
            Dim iLastCount As Integer
           
            For iCount = 1 To 6
                If InStr(1, sCurrentLine, "|", vbTextCompare) <> 0 Then
                   sTemp = Left(sCurrentLine, InStr(1, sCurrentLine, "|", vbTextCompare) - 1)
                Else
                    sTemp = sCurrentLine
                End If
                
                iLastCount = iCount
                
                Select Case iCount
                
                    Case 1
                        If sTemp = "" Then
                           sLinkAbsolutePath = sCurrentLine ' No delimiter so just take the line as the filepath and exit the select statement. Remaining parameters set to defaults in code below...
                           Exit For
                        Else
                           sLinkAbsolutePath = sTemp          ' Take current value and proceed
                        End If
                       
                    Case 2
                        If sTemp = "" Then
                           sClassID = sCurrentLine   ' Parameter string exhausted so use current value and exit the Select statement. Will set a default later in code below...
                           Exit For
                        Else
                            sClassID = sTemp            ' Take current value and proceed
                        End If
                    
                    Case 3
                        If sTemp = "" Then
                            sLinkUsesRelative = sCurrentLine   ' Parameter string exhausted so use current value and exit the Select statement. Will set a default later in code below...
                            Exit For
                        Else
                            sLinkUsesRelative = sTemp    ' Take current value and proceed
                        End If
                        
                    Case 4   ' Absolute Path
                        If sTemp = "" Then
                            AbsPathSelector = -1   ' Parameter string exhausted. Value of -1 will result in code below setting a default value,
                            Exit For                         ' No need to proceed with parameter string
                        Else
                           AbsPathSelector = CInt(sTemp)
                           
                        End If
                    
                    Case 5     ' Relative path
                        If sTemp = "" Then
                            RelPathSelector = -1   ' Parameter string is exhausted. Value of -1 will result in code below setting a default value
                            Exit For                         ' No need to proceed with parameter string
                        Else
                            RelPathSelector = CInt(sTemp)    ' Take current value and proceed
                          
                        End If
                    
                    Case 6
                        If sTemp = "" Then
                            sChildUpdateMode = sCurrentLine
                            Exit For
                        Else
                            sChildUpdateMode = sTemp      ' Take current value and exit
                            Exit For
                        End If
                    End Select
                    If sCurrentLine <> sTemp Then
                       sCurrentLine = Right(sCurrentLine, Len(sCurrentLine) - Len(sTemp) - 1)
                    Else
                        sCurrentLine = ""
                    End If
                    
            Next
           
           sLinkRelativePath = GetRelativePath(sFilePath, sLinkAbsolutePath)
           
           Select Case AbsPathSelector
                Case AbsValidPath
                    oLink.Properties("idmLinkAbsolutePath") = sLinkAbsolutePath
                Case AbsInvalidPath
                     oLink.Properties("idmLinkAbsolutePath") = INVALID_ABS_PATH
                Case AbsEmptyPath
                    oLink.Properties("idmLinkAbsolutePath") = EMPTY_ABS_PATH
                Case Else
                    oLink.Properties("idmLinkAbsolutePath") = sLinkAbsolutePath
            End Select
            
            Select Case RelPathSelector
                Case RelValidPath
                    oLink.Properties("idmLinkRelativePath") = sLinkRelativePath
                Case RelInvalidPath
                     oLink.Properties("idmLinkRelativePath") = INVALID_REL_PATH
                Case RelEmptyPath
                    oLink.Properties("idmLinkRelativePath") = EMPTY_REL_PATH
                Case Else
                    oLink.Properties("idmLinkRelativePath") = sLinkRelativePath
            End Select
            
            If sClassID = "" Then
                oLink.ClassID = idmStaticLink
            Else
                oLink.ClassID = CInt(sClassID)
            End If
            
            If sLinkUsesRelative = "" Then
                oLink.Properties("idmLinkUsesRelativePath") = idmFileLinkAbsolute
            Else
                oLink.Properties("idmLinkUsesRelativePath") = CInt(sLinkUsesRelative)
            End If
            
            If sChildUpdateMode = "" Then
                oLink.Properties("idmLinkChildUpdateMode") = idmUpdateAutomatic
            Else
                oLink.Properties("idmLinkChildUpdateMode") = CInt(sChildUpdateMode)
            End If
            
            ' Add this link to the collection
            oLinks.Add oLink
        End If
        
    Loop
    
    fsTextStream.Close
    Set fsFile = Nothing
    Set fsTextStream = Nothing
    Set oLinks = Nothing
    Set orecognizer = Nothing
    
    
End Sub

Private Sub RecognizerPlugIn_GetPlugInInfo(Name As String, Version As String, CompanyName As String)
    
    Name = "FileNET QA Custom Recognizer"
    Version = "1.0"
    CompanyName = "FileNET"
    
End Sub

Private Sub RecognizerPlugIn_IsSupportedFileType(IsSupported As Boolean)
    
    If Right(sFilePath, 4) = ".txt" Then
        IsSupported = True
    Else
        IsSupported = False
    End If
    
End Sub

Private Sub RecognizerPlugIn_SetFilePath(ByVal FilePath As String)
    sFilePath = FilePath
End Sub

Function GetRelativePath(RelativeTo As String, PathToCheck As String) As String
    
    Dim ParentDirPath As String
    Dim ParentFilename As String
    Dim ChildDirPath As String
    Dim ChildFileName As String
    Dim bContinue As Boolean
    Dim ParentBlock As String
    Dim ChildBlock As String
    Dim sRelativePath As String
    
    
    ' Seprate Filename from Dir Path
    ParsePath RelativeTo, ParentFilename, ParentDirPath
    ParsePath PathToCheck, ChildFileName, ChildDirPath
    
    ParentDirPath = ParentDirPath & "\"
    ChildDirPath = ChildDirPath & "\"
    
    If StrComp(Left(ParentDirPath, 2), Left(ChildDirPath, 2), vbTextCompare) <> 0 Then
        ' mismatched drives - no relative path can be created - return an empty string
        GetRelativePath = ""
        Exit Function
    Else ' Look for a relative path
        ParentDirPath = Right(ParentDirPath, Len(ParentDirPath) - 3)
        ChildDirPath = Right(ChildDirPath, Len(ChildDirPath) - 3)
        
        bContinue = True
        
        Do While bContinue
        
            ParentBlock = Left(ParentDirPath, InStr(1, ParentDirPath, "\", vbTextCompare))
            ChildBlock = Left(ChildDirPath, InStr(1, ChildDirPath, "\", vbTextCompare))
            
            If StrComp(ParentBlock, ChildBlock, vbTextCompare) = 0 And ParentBlock <> "" Then
                ' Continue and step to the next backslash
                ParentDirPath = Right(ParentDirPath, Len(ParentDirPath) - Len(ParentBlock))
                ChildDirPath = Right(ChildDirPath, Len(ChildDirPath) - Len(ChildBlock))
            
            ElseIf ParentBlock = "" And ChildBlock <> "" Then   ' If we run out of parent path and still have child path, then child is below parent
                bContinue = False
                sRelativePath = ".\" & ChildDirPath & ChildFileName
             
            ElseIf ParentBlock <> "" And ChildBlock = "" Then  ' Child is above parent in same branch
                Do While ParentBlock <> ""
                   sRelativePath = sRelativePath & "..\"
                   ParentDirPath = Right(ParentDirPath, Len(ParentDirPath) - Len(ParentBlock))
                   ParentBlock = Left(ParentDirPath, InStr(1, ParentDirPath, "\", vbTextCompare))
                Loop
                bContinue = False
                sRelativePath = sRelativePath & ChildFileName
                
            ElseIf StrComp(ParentBlock, ChildBlock, vbTextCompare) <> 0 And (ParentBlock <> "" And ChildBlock <> "") Then
                ' else child on different branch above
                 Do While ParentBlock <> ""
                   sRelativePath = sRelativePath & "..\"
                   ParentDirPath = Right(ParentDirPath, Len(ParentDirPath) - Len(ParentBlock))
                   ParentBlock = Left(ParentDirPath, InStr(1, ParentDirPath, "\", vbTextCompare))
                Loop
                bContinue = False
                sRelativePath = sRelativePath & ChildDirPath & ChildFileName
            ElseIf ParentBlock = "" And ChildBlock = "" Then
                bContinue = False
                sRelativePath = ".\" & ChildFileName
            End If
            
        Loop
        
    End If
    GetRelativePath = sRelativePath
   
End Function

Private Sub ParsePath(FilePath As String, ByRef sFileName As String, ByRef sDirPath As String)
    ' Sub-procedure to seperate filename from dir path.  Values extracted from argument, FilePath, and stored By Ref
    ' in sFileName and sDirPath
       
    Dim Index As Integer
    
    For Index = Len(FilePath) To 1 Step -1
        
        If Mid$(FilePath, Index, 1) = "\" Then
           sFileName = Right$(FilePath, Len(FilePath) - Index)
           sDirPath = Left$(FilePath, Index - 1)
           Exit For
        End If
    
    Next Index
             
End Sub
