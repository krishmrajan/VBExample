VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "idmwizard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public WithEvents oWizard As IDMObjects.DocumentWizard
Attribute oWizard.VB_VarHelpID = -1
Public iApplType As Integer
Public oAppl As Object
Public strInSubroutine As String
Public CallingOperation As AddCheckinEnum
Public WithEvents RootAction As IDMObjects.action
Attribute RootAction.VB_VarHelpID = -1
Dim ActiveName As String
Public SysType As idmSysTypeOptions
Public oDoc As IDMObjects.Document
Public enuSaveCheckinAction As AddCheckinEnum

Private Sub oWizard_RootActionAvailable(ByVal pRootAction As Object)
Dim oCmd As Object
    Set RootAction = oWizard.RootAction
    pRootAction.ProblemDescriptionFilter = idmActionFileBusy
    'for save and checkin we should reset command
    
    If enuSaveCheckinAction = idmSaveCheckin Then
        For Each oCmd In RootAction.Commands
            If oCmd.Name = LoadResString(STR_CHECKIN) Then
                oCmd.ShowTemporaryCheckin = False
                oCmd.ShowKeepLocal = False
            End If
        Next
    ElseIf enuSaveCheckinAction = idmSaveAdd Then
        For Each oCmd In RootAction.Commands
            If oCmd.Name = LoadResString(STR_ADD) Then
                oCmd.ShowKeepLocal = False
            End If
        Next
    End If
    If enuSaveCheckinAction = idmSaveCheckin Then
       Call ResetCommand(pRootAction, LoadResString(STR_CHECKIN_LINK))
    ElseIf enuSaveCheckinAction = idmSaveAdd Then
       Call ResetCommand2(pRootAction, LoadResString(STR_ADD_LINK))
    End If
End Sub
Private Sub oWizard_WizardFinish(ByVal pObject As Object)
   Call ConvertWord6_95ToCurrentWordVersion(oAppl, iApplType)
   Call gIdmEvent.AddFooter(0, LoadResString(STR_ADD)) '"Add"
   If SysType = idmSysTypeIS Then
        Call DocClose(oAppl, iApplType)
   End If
End Sub
Private Sub RootAction_BeforeExecuteFirst(ByVal pAction As Object)
   
   On Error Resume Next

   If Not (oAppl Is Nothing) Then
        DocMakeActive oAppl, iApplType, ActiveName
        
        ' to handle NNC problem for comp Doc
        Dim sFilePath1, sFilePath2 As String
        Dim bDocActive As Boolean
        bDocActive = False
        sFilePath1 = pAction.TargetObjectFilePath
        sFilePath2 = getFullName(oAppl, iApplType)
        bDocActive = CompareFilePath(sFilePath1, sFilePath2)
        
        If ((pAction.TargetObjectFilePath = getFullName(oAppl, iApplType)) Or (bDocActive = True)) Then
            Set gdoc = pAction.TargetObject
            Call UpdateMezzProperties(oAppl, iApplType, LoadResString(STR_UPDATE_PROP_BEFORE_CHECKIN))
            Set gdoc = Nothing
            If iApplType <> APPL_POWERPOINT Then
                 Call DocSave(oAppl, iApplType)  '###
                'a preference will be put in for the print on checkin
                 Call CheckPref_Print_Checkin(oAppl, iApplType)
            End If
            Call DocClose(oAppl, iApplType)
            Call IDMUpdateMenu(iApplType, oAppl)
        End If
        bDocActive = False
        
    End If
 
    pAction.ProblemDescriptionFilter = Empty
End Sub
Private Sub ResetCommand(oRootAction As IDMObjects.action, sCommand As String)
    Dim oSubAction As IDMObjects.action
    Dim oCmd As Object
    
    For Each oSubAction In oRootAction.SubActions
        If GetDocStatus(oSubAction.TargetObjectFilePath) = DocCheckedout Then
            For Each oCmd In oSubAction.Commands
                If oCmd.Name = sCommand Then ' "Checkin and Link" Then
                    oCmd.ShowTemporaryCheckin = False
                    oCmd.ShowKeepLocal = False
                    oSubAction.SelectedCommand = oCmd
                End If
            Next
        End If
        Call ResetCommand(oSubAction, sCommand)
    Next

End Sub

Private Sub ResetCommand2(oRootAction As IDMObjects.action, sCommand As String)
    Dim oSubAction As IDMObjects.action
    Dim oCmd As Object
    
    For Each oSubAction In oRootAction.SubActions
        For Each oCmd In oSubAction.Commands
            If oCmd.Name = sCommand Then ' "Add and Link" Then
                oCmd.ShowKeepLocal = False
            End If
        Next
        Call ResetCommand(oSubAction, sCommand)
    Next
End Sub

