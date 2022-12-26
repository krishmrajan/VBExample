Attribute VB_Name = "modDocuments"
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


Public Sub AddDocument(oFolder As IDMObjects.Folder, filePath As String, docClass As String, _
        showUserInterface As Boolean)
    
' a file specified by filePath will be added to the library of oFolder with docClass
' if the showUserInterface parameter is False, the document will be added without any user input

    On Error GoTo ErrHandler
    
    ' get the library we will be adding the document to
    Dim oLibrary As IDMObjects.Library
    Set oLibrary = oFolder.Library
    
    ' create a new document object and set its properties
    Dim oNewDocument As IDMObjects.Document
    Set oNewDocument = oLibrary.CreateObject(idmObjTypeDocument, docClass, oFolder)
    Dim directory As String
    Dim fileName As String
    Call GetDirectory(filePath, directory, fileName)
    oNewDocument.Properties("idmName") = GetTitle(fileName)
    
    ' set the file path for the recognizer .. this will cause the recognizer to examine the
    ' file and determine it's type and the number of links it contains (without actually building
    ' the links data structure
    Dim oRecognizer As New IDMObjects.CDRecognizer
    oRecognizer.filePath = filePath
    If (oRecognizer.Links.Count > 0) Then
        ' the document contains links, set up an action to deal with add operation
        
        ' the recognizer knows what type of document the user selected, so get the appropriate
        ' behavior name (CompoundBehaviorID) and create an object
        Dim oBehavior As IDMObjects.Behavior
        Set oBehavior = CreateObject(oRecognizer.CompoundBehaviorID)
        
        ' the command data is setup with defaults for adding a compound document to a library
        Dim oCmdData As New IDMObjects.CommandAddData
        oCmdData.Document = oNewDocument
        oCmdData.filePath = oRecognizer.filePath
        oCmdData.Links = oRecognizer.Links

        ' the command data is used to set up actions for each of the documents in the compound document
        Dim oAction As IDMObjects.action
        Set oAction = oBehavior.CreateRootAction(oCmdData)
    
        If (showUserInterface) Then
            ' display the action in the action grid control so the user can modify the action
            Call ShowAction(oAction)
        Else
            ' setup the action with predefined settings (which override the defaults) and execute it without
            ' any user interaction
            Call PerformAddAction(oAction)
        End If
        
    Else
        ' the document does not contain links, use the traditional add operation
        
        If (showUserInterface) Then
            oNewDocument.SaveNew filePath, idmDocSaveNewWithUIWizard
        Else
            oNewDocument.SaveNew filePath
        End If
        oFolder.File oNewDocument
    
    End If
    
    Exit Sub
ErrHandler:
    MsgBox Err.Description, vbCritical, "Add Document"
End Sub

Private Sub SetAddActionOptions(oAction As IDMObjects.action)

    On Error GoTo ErrHandler
    
    ' change the keepLocalCopy property of the selected command
    ' the default setting is to remove the document afer it has been added to the library; in this example, we
    ' want to show how to modify the commands of the action and the KeepLocalCopy property seemed like a
    ' useful way to do this.

    oAction.SelectedCommand.KeepLocalCopy = idmKeepOptionsKeep
    
    ' change the property for all the sub actions by calling this function recursively
    If (oAction.SubActions.Count > 0) Then
        Dim subActionIndex As Integer
        For subActionIndex = 1 To oAction.SubActions.Count
            Call SetAddActionOptions(oAction.SubActions.Item(subActionIndex))
        Next subActionIndex
    End If

    Exit Sub

ErrHandler:
    MsgBox Err.Description, vbCritical, "Set Action Options"
    
End Sub

Private Sub PerformAddAction(oAction As IDMObjects.action)

    On Error GoTo ErrHandler
    
    ' change some of the properties of the action object
    Call SetAddActionOptions(oAction)
    
    ' execute the action .. the compound document will be added to the library
    Dim success As Boolean
    success = oAction.Execute(idmAbortOnFailure, False)
    If (Not success) Then
        MsgBox "There were problems adding the compound document", vbCritical, "Setup Action"
    End If
    
    Exit Sub
   
ErrHandler:
    MsgBox Err.Description, vbCritical, "Setup Action"
   
End Sub

Public Sub CheckoutDocument(oDocument As IDMObjects.Document, showUserInterface As Boolean)
   
    On Error GoTo ErrHandler

    ' make sure we can checkout the document
    If (Not oDocument.GetState(idmDocCanCheckout)) Then
        MsgBox "Can't check out document", vbCritical, "Checkout Document"
        Exit Sub
    End If
    
    ' check if the document is compound or normal
    If (oDocument.GetState(idmDocHasChild)) Then
        ' the document is compound
        
        ' setup the command data with the defaults for checking out a document
        Dim oCmdData As New IDMObjects.CommandCheckoutData
        oCmdData.Document = oDocument
       
        ' use the document's behavior object
        Dim oBehavior As IDMObjects.Behavior
        Set oBehavior = oDocument.Compound.Behavior
       
        ' create the action
        Dim oAction As IDMObjects.action
        Set oAction = oBehavior.CreateRootAction(oCmdData)
       
        If (showUserInterface) Then
            ' display the action grid control in a dialog and let the user modify the data manually
            Call ShowAction(oAction)
            
        Else
            ' execute the action without any user interaction
            Dim success As Boolean
            success = oAction.Execute(idmAbortOnFailure, False)
            If (Not success) Then
                MsgBox "There were problems checking out the compound document", vbCritical, "Checkout Document"
            End If
       
       End If

    Else
        ' the document is normal
        
        Dim directory As String
        directory = GetPreferenceValue("DirectoriesAndFiles", "CheckoutsAndCopies", "CheckoutDir")
            
        Dim oVersion As IDMObjects.Version
        Set oVersion = oDocument.Version
        oVersion.Checkout directory, , , idmCheckoutOverwrite
    End If
    
    Exit Sub
   
ErrHandler:
    MsgBox Err.Description, vbCritical, "Checkout Document"
    
End Sub

Public Sub CheckinDocument(oDocument As IDMObjects.Document, showUserInterface As Boolean)
   
    On Error GoTo ErrHandler

    ' make sure we can check in the document
    If (Not oDocument.GetState(idmDocCanCheckin)) Then
        MsgBox "Can't check in document", vbCritical, "CheckinDocument"
        Exit Sub
    End If
    
    ' check if the document is compound or normal
    If (oDocument.GetState(idmDocHasChild)) Then
        ' the document is compound
       
        ' use the recognizer to determine the links in the checked out document
        Dim oRecognizer As New IDMObjects.CDRecognizer
        oRecognizer.filePath = oDocument.Version.CheckoutPath
 
        ' create the behavior using the recognizer
        Dim oBehavior As IDMObjects.Behavior
        Set oBehavior = CreateObject(oRecognizer.CompoundBehaviorID)
        
        ' setup the command data with the defaults for checking in a document
        Dim oCmdData As New IDMObjects.CommandCheckinData
        oCmdData.Document = oDocument
        oCmdData.filePath = oRecognizer.filePath
        oCmdData.Links = oRecognizer.Links
           
        ' create the action
        Dim oAction As IDMObjects.action
        Set oAction = oBehavior.CreateRootAction(oCmdData)
       
        If (showUserInterface) Then
            ' display the action grid control in a dialog and let the user modify the data manually
            Call ShowAction(oAction)
            
        Else
            ' execute the action without any user interaction
            Dim success As Boolean
            success = oAction.Execute(idmAbortOnFailure, False)
            If (Not success) Then
                MsgBox "There were problems checking in the compound document", vbCritical, "Checkin Document"
            End If
       
       End If
       
    Else
        ' the document is normal
        
        Dim directory As String
        directory = GetPreferenceValue("DirectoriesAndFiles", "CheckoutsAndCopies", "CheckoutDir")
            
        Dim oVersion As IDMObjects.Version
        Set oVersion = oDocument.Version
        If (showUserInterface) Then
            oVersion.Checkin directory, , idmCheckinWithUI + idmCheckinKeep
        Else
            oVersion.Checkin directory, , idmCheckinKeep
        End If
    End If
    
    Exit Sub
   
ErrHandler:
    MsgBox Err.Description, vbCritical, "Checkin Document"
    
End Sub

Public Sub OpenDocument(oDocument As IDMObjects.Document, showUserInterface As Boolean)
   
    On Error GoTo ErrHandler

    ' make sure we can open and view the document
    If (Not oDocument.GetState(idmDocCanView)) Then
        MsgBox "Can't view document", vbCritical, "OpenDocument"
        Exit Sub
    End If
    
    ' check if the document is compound or normal
    If (oDocument.GetState(idmDocHasChild)) Then
        ' the document is compound
        
        ' setup the command data with the defaults for launching a document (in native application)
        Dim oCmdData As New IDMObjects.CommandLaunchData
        oCmdData.Document = oDocument
        oCmdData.Option = idmCDLaunchNativeApplication
       
        ' use the document's behavior object
        Dim oBehavior As IDMObjects.Behavior
        Set oBehavior = oDocument.Compound.Behavior

        ' create the action
        Dim oAction As IDMObjects.action
        Set oAction = oBehavior.CreateRootAction(oCmdData)
       
        If (showUserInterface) Then
            ' display the action grid control in a dialog and let the user modify the data manually
            Call ShowAction(oAction)
            
        Else
            ' execute the action without any user interaction
            Dim success As Boolean
            success = oAction.Execute(idmAbortOnFailure, False)
            If (Not success) Then
                MsgBox "There were problems opening the compound document", vbCritical, "Open Document"
            End If
       
       End If

    Else
        ' the document is normal
        
        oDocument.Launch idmDocLaunchNativeApplication
    End If
    
    Exit Sub
   
ErrHandler:
    MsgBox Err.Description, vbCritical, "Open Document"
    
End Sub

Private Sub ShowAction(oAction As IDMObjects.action)
   
   On Error GoTo ErrHandler
   
    ' show a form with the action grid control .. this allows the user to modify the action before execution
    Dim oActionForm As New ActionForm
    ' we control the contents of the control by setting the root item
    oActionForm.ActionGrid.SetRootItem oAction, True
    oActionForm.Show vbModal
    
    Exit Sub
    
ErrHandler:
   MsgBox Err.Description, vbCritical, "Show Action"
   
End Sub
