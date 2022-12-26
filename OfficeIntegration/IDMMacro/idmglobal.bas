Attribute VB_Name = "IDMGlobal"
Option Explicit
Global gbPropMgrStatus As Boolean
'Global golocalDB As IDMObjects.LocalDb
Global gsPathRegistryKey As String
Global golib As IDMObjects.Library
'for ms legal
Global gIdmEvent As idmEvent
Public gpWnd As Long
Public lParentHWND As Long
Global lpPrevWndProc As Long
Global gHW As Long
Public bRename As Boolean
Global gdoc As IDMObjects.Document
Global gbFileNETSave As Boolean
Public goAppl As Object
Public giApplType As Integer
Public gActionType As AddCheckinEnum
Public oAppObject As Object
Public bIsUpdateMenuToolbar As Boolean

'make menuitem and toolbar button as global
Public FnMenuAdd As CommandBarControl
Public FnMenuOpen As CommandBarControl
Public FnMenuFnCheckin As CommandBarControl
Public FnMenuCancelCheckout As CommandBarControl
Public FnMenuSave As CommandBarControl
Public FnMenuShowProperty As CommandBarControl
Public FnMenuUpdateProperty As CommandBarControl
Public FnMenuInsertProperty As CommandBarControl
Public FnMenuInsertFile As CommandBarControl
Public FnMenuPreferences As CommandBarControl

Public FnBtnAdd As CommandBarButton
Public FnBtnOpen As CommandBarButton
Public FnBtnCheckin As CommandBarButton
Public FnBtnCancelCheckout As CommandBarButton
Public FnBtnSave As CommandBarButton
Public FnBtnShowProperty As CommandBarButton
Public FnBtnUpdateProperty As CommandBarButton
Public FnBtnInsertProperty As CommandBarButton


'User Defined typedefs
Public Type tDocInfo
    eSystemType As idmSysTypeOptions
    sLibraryName As String
    vDocId As Variant
End Type

Enum ContainerTypeEnum
    NoContainer
    ContainerInt
    ContainerAtt
End Enum

Enum FilterEnum
    actionfilebusy
    setempty
End Enum

Enum DocStatusEnum
    docnew
    DocCheckedout
    DocCopied
End Enum

Enum AddCheckinEnum
    idmAdd
    idmCheckin
    idmSaveCheckin
    idmSaveAdd
End Enum

Enum AppIntOperationEnum
    IDMReplace
    IDMInsert
End Enum

Public Const APPL_WORD = 1
Public Const APPL_EXCEL = 2
Public Const APPL_POWERPOINT = 3
Public Const APPL_OUTLOOK = 4
Public Const APPL_WORDPRO = 5
Public Const DIRECTORIES_FILES = 7

'***see new global variable g_FN_CANCEL below ***
Public Const MB_FILE_SAVE = 3
Public Const MB_FILE_SAVEAS = 748
Public Const MB_FILE = 30002            'FILE MENU BAR
Public Const MB_INSERT = 30005          'INSERT MENU BAR
Public Const MB_TOOLS = 30007           'Tools Menu Bar
Public Const MB_HELP = 30010            'Help Menu Bar


Public Const CREATE_ERR_MGR               As String = "IDMError.ErrorManager"
Public Const CREATE_COMMON_DLG            As String = "IDMObjects.CommonDialogs"
Public Const CREATE_LIBRARY               As String = "IDMObjects.Library"

Global Const REG_SZ As Long = 1
Global Const REG_DWORD As Long = 4
'Global Const HKEY_CLASSES_ROOT = &H80000000
'Global Const HKEY_CURRENT_USER = &H80000001
'Global Const HKEY_LOCAL_MACHINE = &H80000002
'Global Const HKEY_USERS = &H80000003
Global Const ERROR_NONE = 0
'Global Const KEY_ALL_ACCESS = &H3F

Global DEFAULT_COPY_PATH As String
Global DEFAULT_CHECKOUT_PATH As String
Global g_FN_CANCEL(1 To 3) As String
Global goCmnDlg As IDMObjects.CommonDialogs

'this variable takes the value from TXT_CANCEL_CHECKOUT1/2 in initializeVars
Global gMNU_CANCEL_CHECKOUT As Integer
Public DEFAULT_SAVE_PATH As String

'Used by the frmPropertyManager to attach itself to the application window
Public Const SWP_NOMOVE = 2
Public Const SWP_NOSIZE = 1
Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Public Const HWND_TOPMOST = -1
Public Const HWND_TOP = 0
Public Const HWND_NOTOP = -2
Public Const HWND_BOTTOM = 1

'==========STRINGS========================
Public Const ADD_FN_DOC = 1001

Public Const READ_ONLY = 3001

Public Const DEFAULT = 4000
Public Const DEFAULT_PATH = 4001
Public Const DOC_PATH = 4002
Public Const DLG_ERR_EXCEL_DEL_PROP = 4003
Public Const DLG_ERR_EXCEL_GET_PROP = 4004
Public Const DLG_ERR_EXCEL_INSERT_PROP = 4005
Public Const DLG_ERR_EXCEL_UPDATE_PROP = 4006
Public Const DLG_ERR_FRMPROPMGR_CMDINSERT = 4007
Public Const DLG_ERR_FRMPROPMGR_CMDREPLACE_CLICK = 4008
Public Const DLG_ERR_FRMPROPMGR_CMDGOTO_CLICK = 4009
Public Const DLG_ERR_FRMPROPMGR_EXCEL_SUMMARY = 4010
Public Const DLG_ERR_FRMPROPMGR_REMOVE_ENTRY = 4011
Public Const DLG_ERR_FRMPROPMGR_WORD_SUMMARY = 4012
Public Const DLG_ERR_GET_DOC_STATUS = 4013
Public Const DLG_ERR_GET_OBJECT = 4014
Public Const DLG_ERR_GET_PROP_PROP = 4015
Public Const DLG_ERR_GOTO_PROP = 4016
Public Const DLG_ERR_INBOOKMARK = 4017
Public Const DLG_ERR_INSERT_PROP = 4018
Public Const DLG_ERR_REPLACE_PROP = 4019
Public Const DLG_ERR_SHOW_PROP = 4020
Public Const DLG_ERR_SHOW_PROP_MGR = 4021
Public Const DLG_ERR_SHOWPROP = 4022
Public Const DLG_ERR_WORD_DEL_PROP = 4023
Public Const DLG_ERR_WORD_GET_PROP = 4024
Public Const DLG_ERR_WORD_GOTO_PROP = 4025
Public Const DLG_ERR_WORD_INSERT_PROP = 4026
Public Const DLG_ERR_WORD_UPDATE_PROP = 4027
Public Const DLG_ERROR_TITLE = 4028
Public Const DLG_PRINT = 4029
Public Const DLG_SAVE_TO_LOCAL_DRIVE = 4030
Public Const DLG_WARNING = 4031
Public Const DLG_ERR_PAGE_ZERO = 4032
Public Const DLG_ERR_PAGE_FIRST = 4033
Public Const DLG_ERR_PAGE_LAST = 4034

Public Const EXT_WORD = 5001
Public Const EXT_EXCEL = 5002
Public Const EXT_PP = 5003

Public Const EXT_WORD0 = 5007
Public Const EXT_WORD1 = 5008
Public Const EXT_WORD2 = 5009
Public Const EXT_WORD3 = 5010
Public Const EXT_WORD4 = 5011
Public Const EXT_WORD5 = 5012
Public Const EXT_WORD6 = 5013
Public Const EXT_WORD7 = 5014
Public Const EXT_WORD8 = 5015
Public Const EXT_WORD9 = 5016
Public Const EXT_WORD10 = 5017
Public Const EXT_WORD11 = 5018
Public Const EXT_WORD12 = 5019
Public Const EXT_WORD13 = 5020
Public Const EXT_WORD14 = 5021
Public Const EXT_WORD15 = 5022
Public Const EXT_WORD16 = 5023
Public Const EXT_WORD17 = 5024

Public Const EXT_EXCEL1 = 5025
Public Const EXT_EXCEL2 = 5026
Public Const EXT_EXCEL3 = 5027
Public Const EXT_EXCEL4 = 5028
Public Const EXT_EXCEL5 = 5029
Public Const EXT_EXCEL6 = 5030
Public Const EXT_EXCEL7 = 5031
Public Const EXT_EXCEL8 = 5032
Public Const EXT_EXCEL9 = 5033
Public Const EXT_EXCEL10 = 5034
Public Const EXT_EXCEL11 = 5035
Public Const EXT_EXCEL12 = 5036
Public Const EXT_EXCEL13 = 5037
Public Const EXT_EXCEL14 = 5038
Public Const EXT_EXCEL15 = 5039
Public Const EXT_EXCEL16 = 5040
Public Const EXT_EXCEL17 = 5041
Public Const EXT_EXCEL18 = 5042
Public Const EXT_EXCEL19 = 5043

Public Const EXT_PP1 = 5044
Public Const EXT_PP2 = 5045
Public Const EXT_PP3 = 5046
Public Const EXT_PP4 = 5047
Public Const EXT_PP5 = 5048
Public Const EXT_PP6 = 5049

Public Const FILTER_ALL_FILE = 6001
Public Const FILTER_WORD_FILTER1 = 6002
Public Const FILTER_WORD_FILTER2 = 6003
Public Const FILTER_WORD_FILTER3 = 6004
Public Const FILTER_WORD_FILTER4 = 6005

Public Const FILTER_EXCEL_FILTER1 = 6006
Public Const FILTER_EXCEL_FILTER2 = 6007
Public Const FILTER_EXCEL_FILTER3 = 6008
Public Const FILTER_EXCEL_FILTER4 = 6009
Public Const FILTER_EXCEL_FILTER5 = 6010
Public Const FILTER_EXCEL_FILTER6 = 6011
Public Const FILTER_EXCEL_FILTER7 = 6012

Public Const FILTER_WORDPPO_FILTER1 = 6013
Public Const FILTER_WORDPPO_FILTER2 = 6014
Public Const FILTER_WORDPPO_FILTER3 = 6015
Public Const FILTER_WORDPPO_FILTER4 = 6016


Public Const IDM_GRID_HEADER = 9001
Public Const IDM_PROP_BODY = 9002
Public Const IDM_PROP_FOOTER_EVEN_PAGE = 9003
Public Const IDM_PROP_FOOTER_FIRST_PAGE = 9004
Public Const IDM_PROP_FOOTER_ODD_PAGE = 9005
Public Const IDM_PROP_HEADER_EVEN_PAGE = 9006
Public Const IDM_PROP_HEADER_FIRST_PAGE = 9007
Public Const IDM_PROP_HEADER_ODD_PAGE = 9008
Public Const IDM_QUESTION_MARK = 9009
Public Const IDM_SEPARATOR = 9010
Public Const IDM_SPACE = 9011
Public Const IDM_TAG = 9012
Public Const IDM_UPDATE_MENU = 9013
Public Const IDM_WORD_TEMPLATE_FILENAME = 9014
Public Const INSERT_FN_ATTACHMENT = 9015
Public Const INSERT_FN_FILE = 9016

Public Const LOCAL_CHECKOUT = 1201
Public Const LOCAL_COPIES = 1202

Public Const MNU_FN_ADD = 1301
Public Const MNU_FN_CHECKIN = 1302
Public Const MNU_FN_OPEN = 1303
Public Const MNU_FN_PREFERENCES = 1304
Public Const MNU_FN_PROPERTIES = 1305
Public Const MNU_FN_SAVE = 1306
Public Const MNU_FN_SHOW = 1307
Public Const MNU_INSERT_FILE = 1308
Public Const MNU_INSERT_MEZZ_PROP = 1309
Public Const MNU_SAVEAS = 1310
Public Const MNU_UPDATE_MEZZ_PROP = 1311

Public Const MSG_ADD = 1312
Public Const MSG_CANCEL_CHECKOUT = 1313
Public Const MSG_CANNOT_CHECKOUT = 1314
Public Const MSG_CANNOT_CREATE_CMNDIALOG = 1315
Public Const MSG_CANNOT_CREATE_DOC = 1316
Public Const MSG_CANNOT_CREATE_ERRMGR = 1317
Public Const MSG_CANNOT_CREATE_LIBRARY = 1318
Public Const MSG_CANNOT_GET_PROP_OBJECT = 1319
Public Const MSG_CANNOT_GET_PROPERTY = 1320
Public Const MSG_CANNOT_GET_VERSION = 1321
Public Const MSG_CHECK_PROP_MGR = 1322
Public Const MSG_CHECK_SDM_STRING = 1323
Public Const MSG_CHECKIN = 1324
Public Const MSG_CLOSE = 1325
Public Const MSG_CONFIRM_DELETE = 1326
Public Const MSG_CONFIRM_REPLACE = 1327
Public Const MSG_CONVERSION = 1328
Public Const MSG_COPY_TO_CLIPBOARD = 1329

Public Const MSG_DELETE_PROP = 1330
Public Const MSG_DELETE_MEZZ_PROP_EXCELL = 1331
Public Const MSG_DELETE_PROP_WORD = 1332

Public Const MSG_DO_YOU_WANT_TO_CONVERT_SDM_TO_IDM = 1333
Public Const MSG_DO_YOU_WANT_TO_PRINT_FILE = 1334
Public Const MSG_DO_YOU_WANT_TO_UPDATE_PROPERTIES = 1335

Public Const MSG_ERR_MANAGER_NOT_INITIALIZED = 1336
Public Const MSG_ERROR_IN_GET_BOOKMARKS_EXCELL = 1337
Public Const MSG_ERROR_IN_GET_BOOKMARKS_WORD = 1338
Public Const MSG_ERROR_IN_SHOW_PROPERTY_MGR = 1339
Public Const MSG_ERROR_IN_WORD_UPDATE_PROP = 1340
Public Const MSG_ERROR_WITHOUT_ERRMGR = 1341

Public Const MSG_FILE_ADD_CHECKEDOUT = 1342
Public Const MSG_FILE_EXISTS_OVERWRITE = 1343
Public Const MSG_FILE_NOT_CHECKOUT = 1344
Public Const MSG_FILE_NOT_EXIST = 1345
Public Const MSG_FILE_OPEN_OVERWRITE = 1346
Public Const MSG_FILE_OPENPPT_OVERWRITE = 1347
Public Const MSG_FILE_SAVE_CHANGES = 1348
Public Const MSG_FILE_SAVEAS_CHECKEDOUT = 1349
Public Const MSG_FILE_SAVEAS_OVERWRITE = 1350
Public Const MSG_FILECHECKIN = 1351
Public Const MSG_FILENET = 1352
Public Const MSG_FILENETSAVE = 1353

Public Const MSG_GET_ACTIVE_FILE = 1354
Public Const MSG_GET_DOC_PATH = 1355
Public Const MSG_GET_DOC_STATUS = 1356
Public Const MSG_GET_PROP_LABEL = 1357
Public Const MSG_GET_UNIQUE_NUMBER = 1358
Public Const MSG_GOTO_BOOKMARK = 1359

Public Const MSG_IN_CELL = 1360
Public Const MSG_IN_SECTION = 1361
Public Const MSG_INSERT_MEZZ_PROP = 1362
Public Const MSG_INSERT_MEZZ_PROP_EXCELL = 1363
Public Const MSG_INSERT_POINT_ALREADY_PROPERTY = 1364
Public Const MSG_INSERT_PROP = 1365
Public Const MSG_INSERT_PROPERTIES = 1366
Public Const MSG_INSERT_WORD_BOOKMARK = 1367
Public Const MSG_IS_DOC_CHECKEDOUT = 1368

Public Const MSG_KEEP_LOCAL_COPY = 1369
Public Const MSG_LOAD_IDM_STRING = 1370
Public Const MSG_MUST_ADD_FIRST = 1371
Public Const MSG_NO_DOC = 1372
Public Const MSG_NOT_A_CHECHEDOUT_DOC = 1373
Public Const MSG_NOT_A_CHECKEDOUT_FILE = 1374
Public Const MSG_NOT_CHECKED_OUT = 1375
Public Const MSG_NOT_CHECKOUT = 1376
Public Const MSG_ON_PAGE = 1377
Public Const MSG_OPEN = 1378
Public Const MSG_OPEN_DOC_NOT_SUPPORT = 1379
Public Const MSG_OPEN_METHOD_FAILED = 1380
Public Const MSG_PREFERENCES = 1381
Public Const MSG_PRINT = 1382

Public Const MSG_SAVE = 1383
Public Const MSG_SAVEAS = 1384
Public Const MSG_SHOW_PROP = 1385
Public Const MSG_SHOW_PROPERTY_MGR = 1386
Public Const MSG_SHOW_TOOLBAR = 1387
Public Const MSG_UPDATE_PROP = 1388
Public Const MSG_UPDATE_PROP_GRID = 1389
Public Const MSG_UPDATE_PROPERTIES = 1390
Public Const MSG_WITH = 1391
Public Const MSG_OPERATION_NOT_DEFINED = 1392
Public Const MNU_FN_HELP = 1393
Public Const MSG_POWERPOINT_BLOCK = 1394
Public Const MSG_RETURN = 1395
Public Const MSG_PRESENTATION = 1396
Public Const MSG_CONTAINSLINKS = 1397
Public Const MSG_CAN_NOT_INSERT_COMP_DOC = 1398
Public Const MSG_INSUFFICIENT_RIGHT = 1399
Public Const MSG_FILE_PATH_AND_NAME_TOO_LONG = 13100
Public Const MSG_FILE_NAME_CANNOT_CONTAIN_CHARACTER = 13101
Public Const MSG_DELETE_CHARACTER = 13102
Public Const MSG_CANOT_COPY_DOC = 13103
Public Const MSG_FILE_ADD_CHECKEDOUT1 = 13104
Public Const MSG_FILE_ADD_CHECKEDOUT2 = 13105
Public Const MSG_WOULD_YOU_LIKE_TO_CONTINE_ADD = 13106
Public Const MSG_DOCID = 13107
Public Const MSG_LIBRARY = 13108
Public Const MSG_FILENAME = 13109
Public Const MSG_THE_CHECKEDOUT_FILE_INFO = 13110
Public Const MSG_WOULD_YOU_LIKE_TO_CANCEL_CHECKOUT = 13111
Public Const MSG_CLIPBOARD = 13112

Public Const OPEN_FN_DOC = 1501
Public Const OPTIONS_DIALOG_ADD_TITLE = 1502
Public Const OPTIONS_DIALOG_CHECKIN_TITLE = 1503
Public Const OPTIONS_DIALOG_SHOW_PROPS = 1504
Public Const OPTIONS_GET_DOC_OBJECT = 1505

Public Const PERSONAL = 1601
Public Const PROMPT_DELETE_PROPERTY = 1602
Public Const PROMPT_REPLACE_PROPERTY = 1603
Public Const PRINTER_NOT_INSTALL = 1604

Public Const REG_KEY_DEFAULT_LOCATION = 1801
Public Const REG_KEY_EXCELL = 1802
Public Const REG_KEY_POWERPOINT = 1803
Public Const REG_KEY_SHELL_FOLDER = 1804
Public Const REG_KEY_WORD = 1805

Public Const SDM_N_TAG = 1901
Public Const SDM_TAG = 1902
Public Const STR_ADD_CHECKIN_RETRIEVE = 1903
Public Const STR_APP_INTEGRATION = 1904
Public Const STR_CAN_NOT_BE_CONVERT = 1905
Public Const STR_CHECKEDOUT = 1906
Public Const STR_COPIED = 1907
Public Const STR_DCDC = 1908
Public Const STR_DOCUMENT = 1909
Public Const STR_EXCEL_INTEGRATION = 1910
Public Const STR_OUTLOOK_INTEGRATION = 1911
Public Const STR_PP_FILTER1 = 1912
Public Const STR_PP_FILTER2 = 1913
Public Const STR_PP_FILTER3 = 1914
Public Const STR_PP_FILTER4 = 1915
Public Const STR_PP_FILTER5 = 1916
Public Const STR_PP_FILTER6 = 1917
Public Const STR_PP_FILTER7 = 1918
Public Const STR_PP_FILTER8 = 1919
Public Const STR_PP_FILTER9 = 1920
Public Const STR_PP_INTEGRATION = 1921
Public Const STR_SET_DOC_TITLE_TO_FILENAME = 1922
Public Const STR_SHOW_DOC_STATUS = 1923
Public Const STR_SHOW_TOOLBAR = 1924
Public Const STR_THE_BOOKMARK = 1925
Public Const STR_UPDATE_ENBEDDED_PROP = 1926
Public Const STR_UPDATE_PROP_AFTER_CHECKOUT = 1927
Public Const STR_UPDATE_PROP_BEFORE_CHECKIN = 1928
Public Const STR_UPDATE_PROPERTIES = 1929
Public Const STR_WORD_INTEGRATION = 1930
Public Const STR_PRT_SAVE_ADD = 1931
Public Const STR_PRT_SAVE_CHECKIN = 1932
Public Const STR_SDCDC = 1933
Public Const STR_DIRECTORIES_AND_FILES = 1934
Public Const STR_LOCAL_CACHING = 1935
Public Const STR_CACHE_DIRECTORY = 1936
Public Const STR_KEEP_LOCAL_COPY = 1937
Public Const STR_SHOW_KEEP_LOCAL_COPY = 1938
Public Const STR_WARNING = 1939
Public Const STR_HELP_FILE = 1940
Public Const STR_FN_HELP = 1941
Public Const STR_PATH_DOES_NOT_EXIST = 1942
Public Const STR_VERIFY_PATH = 1943
Public Const STR_CHECK_OR_COPY_TO = 1944
Public Const STR_FOLDER_REQ_WHEN_ADD = 1945
Public Const STR_CHECKIN_LINK = 1946
Public Const STR_EXTERNAL_DOC = 1947
Public Const STR_Mezzanine_Installed = 1948
Public Const STR_SubKey_IDM_Install = 1949
Public Const STR_CHECKIN = 1950
Public Const STR_ADD = 1951
Public Const STR_ADD_LINK = 1952
Public Const STR_FNI = 1953
Public Const STR_DIALOG = 1954
Public Const STR_MEZZ = 1955
Public Const STR_IDM = 1956
Public Const STR_COPYDIR = 1957
Public Const STR_CHECKOUTDIR = 1958
Public Const STR_PRINT_DOC_ON_ADD = 1959
Public Const STR_IN = 1960
Public Const STR_MODIFY_CONTROLS = 1961
Public Const STR_CHECKOUTS_AND_COPIES = 1962
Public Const STR_OTHER = 1963
Public Const STR_LOCALDB = 1964
Public Const STR_LOCALDB_KEY = 1965
Public Const STR_INDEX_OUT_RANGE = 1966
Public Const STR_PPT = 1967
Public Const STR_POWERPOINT = 1968
Public Const STR_BUTTON = 1969
Public Const STR_TOOLBAR_BUTTON_FACE = 1970
Public Const STR_TOOLBAR_BUTTON_MASK = 1971
Public Const STR_ALLFILES = 1972
Public Const STR_WORD = 1973
Public Const STR_WORD_APPLICATION = 1974
Public Const STR_CACHE = 1975
Public Const STR_MS_WORD = 1976
Public Const STR_REPLICA = 1977
Public Const STR_TOOLBAR_BUTTON_MASK_97 = 1978
Public Const STR_TOOLBAR_BUTTON_MASK_2000 = 1979
Public Const STR_TOOLBAR_BUTTON_FACE_97 = 1980
Public Const STR_TOOLBAR_BUTTON_FACE_2000 = 1981
Public Const STR_DO_YOU_WANT_TO_SAVE_THE_DOC = 1982
Public Const STR_P_LEFT = 1983
Public Const STR_P_RIGHT = 1984
Public Const STR_UPDATE_MENU_TOOLBAR = 1985
Public Const STR_CONVERT_WORD6_95_TO_CURRENT_WORD_VERSION = 1986
Public Const STR_REG_PREF_PATH = 1987
Public Const STR_VALUE = 1988
 

Public Const TXT_CANCEL_CHECKOUT1 = 2001
Public Const TXT_CANCEL_CHECKOUT2 = 2002
Public Const TXT_CANNOT_CONVERT_STRING = 2003
Public Const TXT_NO_THIS_VALUE = 2004
Public Const TXT_ALWAYS = 2005
Public Const TXT_NEVER = 2006
Public Const TXT_PROMPTUSER = 2007
Public Const TXT_INSERT = 2008
Public Const TXT_OPEN = 2009

Public Const GWL_WNDPROC = -4
Public Const WM_SETFOCUS = &H7
Public Const WM_ACTIVATEAPP = &H1C
Public Const WM_ACTIVATE = &H6
Public Const WM_SHOWWINDOW = &H18
Public Const WM_SIZE = &H5
Public Const WM_KILLFOCUS = &H8
'WM_SHOWWINDOW wParam codes
'Public Const SW_PARENTCLOSING = 1
'Public Const SW_OTHERMAXIMIZED = 2
'Public Const SW_PARENTOPENING = 3
'Public Const SW_OTHERRESTORED = 4

'===========API FUNCTIONS====================
Public Declare Function GetActiveWindow Lib "user32" () As Integer
Public Declare Function lstrcpy Lib "kernel32" Alias "lstrcatA" (ByVal lpString1$, ByVal lpString2&) As Long
Public Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal s&) As Long

Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias _
   "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, _
   ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As _
   Long) As Long
Public Declare Function RegQueryValueExNULL Lib "advapi32.dll" Alias _
   "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As _
   String, ByVal lpReserved As Long, lpType As Long, ByVal lpData _
   As Long, lpcbData As Long) As Long
Public Declare Function RegQueryValueExString Lib "advapi32.dll" Alias _
   "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As _
   String, ByVal lpReserved As Long, lpType As Long, ByVal lpData _
   As String, lpcbData As Long) As Long
Public Declare Function RegQueryValueExLong Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As _
   String, ByVal lpReserved As Long, lpType As Long, lpData As Long, lpcbData As Long) As Long

Public Const STANDARD_RIGHTS_ALL = &H1F0000
Public Const SYNCHRONIZE = &H100000
Public Const READ_CONTROL = &H20000
Public Const STANDARD_RIGHTS_READ = (READ_CONTROL)
Public Const STANDARD_RIGHTS_WRITE = (READ_CONTROL)
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_CONFIG = &H80000005
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_DYN_DATA = &H80000006
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_PERFORMANCE_DATA = &H80000004
Public Const HKEY_USERS = &H80000003
Public Const KEY_QUERY_VALUE = &H1
Public Const KEY_SET_VALUE = &H2
Public Const KEY_CREATE_LINK = &H20
Public Const KEY_CREATE_SUB_KEY = &H4
Public Const KEY_ENUMERATE_SUB_KEYS = &H8
Public Const KEY_NOTIFY = &H10
Public Const KEY_EVENT = &H1     '  Event contains key event record
Public Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))
Public Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
Public Const KEY_WRITE = ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))
Public Const KEY_EXECUTE = ((KEY_READ) And (Not SYNCHRONIZE))
Public Const ERROR_SUCCESS = 0&
Public Declare Function GetFullPathName Lib "kernel32" Alias "GetFullPathNameA" _
    (ByVal lpFileName As String, ByVal nBufferLength As Long, ByVal lpBuffer As String, ByVal lpFilePart As String) As Long
Public Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" _
    (ByVal LongName$, ByVal ShortName As String, ByVal Size1 As Long) As Integer

Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

''Declare Function GetFocus Lib "user32" () As Long
Declare Function SetFocus Lib "user32" (ByVal hWnd As Long) As Long

'Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Declare Function GetForegroundWindow Lib "user32" () As Long

Declare Function SetFocusAPI& Lib "user32" Alias "SetFocus" (ByVal hWnd As Long)

Declare Function GetFocus Lib "user32" () As Long

Declare Function CallWindowProc Lib "user32" Alias _
"CallWindowProcA" (ByVal lpPrevWndFunc As Long, _
   ByVal hWnd As Long, ByVal Msg As Long, _
   ByVal wParam As Long, ByVal lParam As Long) As Long

Declare Function SetWindowLong Lib "user32" Alias _
"SetWindowLongA" (ByVal hWnd As Long, _
ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

'Public Const GW_CHILD = 5

'Public Const GW_HWNDFIRST = 0

'Public Const GW_HWNDLAST = 1

'Public Const GW_HWNDNEXT = 2

'Public Const GW_HWNDPREV = 3

'Public Const GW_MAX = 5

'Public Const GW_OWNER = 4

'Public Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long

'Public Declare Function EnumChildWindows Lib "user32" (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long

'Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long

Public Const SW_HIDE = 0

Public Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long

'Public Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long

Public Const SW_MINIMIZE = 6



'Public Declare Function EnableWindow Lib "user32" (ByVal hWnd As Long, ByVal fEnable As Long) As Long

'Public Const WM_CANCELMODE = &H1F

Public Declare Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hWnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long

Public Const MB_OK = &H0&

Public Const MB_ICONEXCLAMATION = &H30&

Public Declare Function GetEnvironmentVariable Lib "kernel32" Alias "GetEnvironmentVariableA" (ByVal lpName As String, ByVal lpBuffer As String, ByVal nSize As Long) As Long

'these api for bitmaps
Public Declare Function CreateIC Lib "gdi32" Alias "CreateICA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, lpInitData As Long) As Long

Public Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long

Public Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long

'help file
Public Declare Function WinHelp Lib "user32" Alias "WinHelpA" (ByVal hWnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, ByVal dwData As Long) As Long

Public Const HELP_CONTEXT = &H1

' ===================================================================
'   Clipboard APIs
' ===================================================================
Public Declare Function OpenClipboard Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function CloseClipboard Lib "user32" () As Long
Public Declare Function RegisterClipboardFormat Lib "user32" Alias "RegisterClipboardFormatA" (ByVal lpString As String) As Long
Public Declare Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As Long
Public Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Public Declare Function EmptyClipboard Lib "user32" () As Long
Public Const CF_DIB = 8
Public Declare Function CountClipboardFormats Lib "user32" () As Long
' ===================================================================
'   Memory APIs (for clipboard transfers)
' ===================================================================
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal cbLength As Long)
Public Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Public Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Public Const GMEM_DDESHARE = &H2000
Public Const GMEM_MOVEABLE = &H2
Public Const CF_GDIOBJFIRST = &H300
Public Const CF_GDIOBJLAST = &H3FF
Public Const CF_PRIVATEFIRST = &H200
Public Const CF_PRIVATELAST = &H2FF

Public Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias _
    "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, _
    ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions _
    As Long, ByVal samDesired As Long, ByVal lpSecurityAttributes _
    As Long, phkResult As Long, lpdwDisposition As Long) As Long
    
Public Declare Function RegSetValueExString Lib "advapi32.dll" Alias _
    "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, _
    ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As _
    String, ByVal cbData As Long) As Long
    
Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" _
(ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, _
ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long

Public Const REG_CREATED_NEW_KEY = &H1           'New Registry Key created
Public Const REG_OPENED_EXISTING_KEY = &H2 'Existing Key opened

Public Declare Function timeGetTime Lib "winmm.dll" () As Long

Public Type TEXTMETRIC
    tmHeight As Long
    tmAscent As Long
    tmDescend As Long
    tmInternalLeading As Long
    tmExternalLeading As Long
    tmAveCharWidth As Long
    tmMaxCharWidth As Long
    tmWeight As Long
    tmOverhang As Long
    tmDigitizedAspectX As Long
    tmDigitizedAspectY As Long
    tmFirstChar As Byte
    tmLastChar As Byte
    tmDefaultChar As Byte
    tmBreakChar As Byte
    tmItalic As Byte
    tmUnderlined As Byte
    tmStruckOut As Byte
    tmPitchAndFamily As Byte
    tmCharSet As Byte
End Type

Public Type mzoMain
    SystemFont As StdFont
    bUseSystemFont As Boolean
End Type

Public g_mzoMain As mzoMain

'Contants used in setting form font.
Public Const m_iDEFAULT_GUI_FONT = 17
Public Const m_iLOG_PIXELS_Y = 90
Public Declare Function GetUserDefaultLCID Lib "kernel32" () As Long
Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long
Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Declare Function GetTextFace Lib "gdi32" Alias "GetTextFaceA" (ByVal hDC As Long, ByVal nCount As Long, ByVal lpFacename As String) As Long
Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Declare Function GetTextMetrics Lib "gdi32" Alias "GetTextMetricsA" (ByVal hDC As Long, lpMetrics As TEXTMETRIC) As Long
'Declare Function GetDeviceCaps Lib "gdi32" (ByVal lDC As Long, ByVal m_iLOG_PIXELS_Y As Long) As Long

Public Const FILE_ATTRIBUTE_NORMAL = &H80
Public Const FILE_ATTRIBUTE_READONLY = &H1
Public Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Public Declare Function SetFileAttributes Lib "kernel32" Alias "SetFileAttributesA" (ByVal lpFileName As String, ByVal dwFileAttributes As Long) As Long

Public Declare Function WNetGetConnection Lib "mpr.dll" Alias "WNetGetConnectionA" (ByVal lpszLocalName As String, ByVal lpszRemoteName As String, cbRemoteName As Long) As Long

Public Const No_ERROR As Long = 0
Public Const lBuffer_size As Long = 255


