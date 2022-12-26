VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmCommonDialogs 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   360
      Top             =   2400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      FontName        =   "Arial"
   End
End
Attribute VB_Name = "frmCommonDialogs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub SaveDialogSetup(iApplicationType As Integer)
    
    Dim sFilter As String
    Dim strTitle As String
    Dim strDefaultExt As String
    
    strTitle = LoadResString(DLG_SAVE_TO_LOCAL_DRIVE)               '"FileNET Save to local file"
    Select Case iApplicationType
        Case APPL_WORD:
            sFilter = LoadResString(FILTER_WORD_FILTER1)            '"Word Documents (*.doc)|*.doc"
            sFilter = sFilter & LoadResString(FILTER_WORD_FILTER2)  '"|Document Templates (*.dot)|*.dot"
            sFilter = sFilter & LoadResString(FILTER_WORD_FILTER3)  '"|Rich Text Format (*.rtf)|*.rtf"
            sFilter = sFilter & LoadResString(FILTER_WORD_FILTER4)  '"|Text Files (*.txt)|*.txt"
            ' sFilter = sFilter & "|All Files (*.*)|*.*"
            strDefaultExt = LoadResString(EXT_WORD) '"doc"
    
        Case APPL_EXCEL:
            sFilter = LoadResString(FILTER_EXCEL_FILTER1)           '"Microsoft Excel Files (*.xls)|*.xls"
            sFilter = sFilter & LoadResString(FILTER_EXCEL_FILTER2) '"|Formatted Text (*.prn)|*.prn"
            sFilter = sFilter & LoadResString(FILTER_EXCEL_FILTER3) '"|Text Files (*.txt)|*.txt"
            sFilter = sFilter & LoadResString(FILTER_EXCEL_FILTER4) '"|Text Files (*.csv)|*.csv"
            sFilter = sFilter & LoadResString(FILTER_EXCEL_FILTER5) '"|Microsoft Works 2.0 Files (*.wks)|*.wks"
            sFilter = sFilter & LoadResString(FILTER_EXCEL_FILTER6) '"|HTML Documents (*.htm)|*.htm"
            sFilter = sFilter & LoadResString(FILTER_EXCEL_FILTER7) '"|Microsoft Excel Addin (*.xla)"
            ' sFilter = sFilter & "|All Files (*.*)|*.*"
            
            strDefaultExt = LoadResString(EXT_EXCEL)                '"xls"
    
        Case APPL_POWERPOINT:
            sFilter = LoadResString(STR_PP_FILTER1)                 '"Presentation (*.ppt)|*.ppt"
            sFilter = sFilter & LoadResString(STR_PP_FILTER2)       '"|Outline/RTF (*.rtf)|*.rtf"
            sFilter = sFilter & LoadResString(STR_PP_FILTER3)       '"|Presentation Template (*.pot)|*.pot"
            sFilter = sFilter & LoadResString(STR_PP_FILTER4)       '"|Power Point Show (*.pps)|*.pps"
            sFilter = sFilter & LoadResString(STR_PP_FILTER5)       '"|Power Point 95 & 97 Presentation *.ppt|*.ppt"
            sFilter = sFilter & LoadResString(STR_PP_FILTER6)       '"|Power Point 95 (*.ppt)|*.ppt"
            sFilter = sFilter & LoadResString(STR_PP_FILTER7)       '"|Power Point 4.0 (*.ppt)|*.ppt"
            sFilter = sFilter & LoadResString(STR_PP_FILTER8)       '"|Power Point 3.0 (*.ppt)|*.ppt"
            sFilter = sFilter & LoadResString(STR_PP_FILTER9)       '"|Power Point Add-In (*.ppa)|*.ppa"
            strDefaultExt = LoadResString(EXT_PP)                   '"ppt"
    
        Case APPL_WORDPRO:
            sFilter = LoadResString(FILTER_WORDPPO_FILTER1)             '"Word Documents (*.doc)|*.doc"
            sFilter = sFilter & LoadResString(FILTER_WORDPPO_FILTER2)   '"|Document Templates (*.dot)|*.dot"
            sFilter = sFilter & LoadResString(FILTER_WORDPPO_FILTER3)   '"|Rich Text Format (*.rtf)|*.rtf"
            sFilter = sFilter & LoadResString(FILTER_WORDPPO_FILTER4)   '"|Text Files (*.txt)|*.txt"
            ' sFilter = sFilter & "|All Files (*.*)|*.*"
            strDefaultExt = LoadResString(EXT_WORD)                     '"doc"

    End Select
    With CommonDialog1
        .DialogTitle = strTitle
        .DefaultExt = strDefaultExt
        .Filter = sFilter
   '     .FilterIndex = 1
        .InitDir = DEFAULT_SAVE_PATH
        .CancelError = True
        .FLAGS = cdlOFNHideReadOnly
    End With
    
End Sub

Sub SaveLocalDialogSetup(strDefaultExt As String, sFilter As String)
    
    Dim strTitle As String
    'Dim strDefaultExt As String
    'Dim sFilter As String
    'sFilter = "Word Documents (*.doc)|*.doc"
            'sFilter = sFilter & "|Document Templates (*.dot)|*.dot"
            'sFilter = sFilter & "|Rich Text Format (*.rtf)|*.rtf"
            'sFilter = sFilter & "|Text Files (*.txt)|*.txt"
            ' sFilter = sFilter & "|All Files (*.*)|*.*"
            'strDefaultExt = "doc"
    With CommonDialog1
        .DialogTitle = strTitle
        .DefaultExt = strDefaultExt
        .InitDir = DEFAULT_SAVE_PATH
        .CancelError = True
        .FLAGS = cdlOFNHideReadOnly
        .Filter = sFilter
    End With
    
End Sub
Sub OpenDialogSetup(iApplType As String)

    Dim sFilter As String
    Dim sDefaultExt As String
    Select Case iApplType
           Case APPL_EXCEL
                sFilter = LoadResString(FILTER_EXCEL_FILTER1)           '"Microsoft Excel Files (*.xls)|*.xls"
                sFilter = sFilter & LoadResString(FILTER_EXCEL_FILTER2) '"|Formatted Text (*.prn)|*.prn"
                sFilter = sFilter & LoadResString(FILTER_EXCEL_FILTER3) '"|Text Files (*.txt)|*.txt"
                sFilter = sFilter & LoadResString(FILTER_EXCEL_FILTER4) '"|Text Files (*.csv)|*.csv"
                sFilter = sFilter & LoadResString(FILTER_EXCEL_FILTER5) '"|Microsoft Works 2.0 Files (*.wks)|*.wks"
                sFilter = sFilter & LoadResString(FILTER_EXCEL_FILTER6) '"|HTML Documents (*.htm)|*.htm"
                sFilter = sFilter & LoadResString(FILTER_EXCEL_FILTER7) '"|Microsoft Excel Addin (*.xla)"
                sFilter = sFilter & LoadResString(FILTER_ALL_FILE)      '"|All Files (*.*)|*.*"
                
                sDefaultExt = LoadResString(EXT_EXCEL)                  '"xls"
           
           Case APPL_POWERPOINT
                sFilter = LoadResString(STR_PP_FILTER1)                 '"Presentation (*.ppt)|*.ppt"
                sFilter = sFilter & LoadResString(STR_PP_FILTER2)       '"|Outline/RTF (*.rtf)|*.rtf"
                sFilter = sFilter & LoadResString(STR_PP_FILTER3)       '"|Presentation Template (*.pot)|*.pot"
                sFilter = sFilter & LoadResString(STR_PP_FILTER4)       '"|Power Point Show (*.pps)|*.pps"
                sFilter = sFilter & LoadResString(STR_PP_FILTER5)       '"|Power Point 95 & 97 Presentation *.ppt|*.ppt"
                sFilter = sFilter & LoadResString(STR_PP_FILTER6)       '"|Power Point 95 (*.ppt)|*.ppt"
                sFilter = sFilter & LoadResString(STR_PP_FILTER7)       '"|Power Point 4.0 (*.ppt)|*.ppt"
                sFilter = sFilter & LoadResString(STR_PP_FILTER8)       '"|Power Point 3.0 (*.ppt)|*.ppt"
                sFilter = sFilter & LoadResString(STR_PP_FILTER9)       '"|Power Point Add-In (*.ppa)|*.ppa"
                sFilter = sFilter & LoadResString(FILTER_ALL_FILE)      '"|All Files (*.*)|*.*"
                sDefaultExt = LoadResString(EXT_PP)                     '"ppt"
     End Select
    With CommonDialog1
        .DialogTitle = LoadResString(MSG_OPEN)
        .DefaultExt = sDefaultExt '"ppt"
        .Filter = sFilter
        .FilterIndex = 1
        .InitDir = DEFAULT_SAVE_PATH
        .CancelError = True
        .FLAGS = cdlOFNHideReadOnly
    End With
    
End Sub

Sub PrintDialogSetup(frm As CommonDialog)
    
    With frm
        .DialogTitle = LoadResString(DLG_PRINT)
    End With
    
End Sub

Private Sub Form_Load()

End Sub
