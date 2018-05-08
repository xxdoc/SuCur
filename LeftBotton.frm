VERSION 5.00
Begin VB.Form LeftBotton 
   BorderStyle     =   0  'None
   ClientHeight    =   3195
   ClientLeft      =   120
   ClientTop       =   120
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Menu UserMenu 
      Caption         =   "SuperMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuControlPanel 
         Caption         =   "Windows控制面板(&C)"
      End
      Begin VB.Menu mnuCMD 
         Caption         =   "命令提示符(&M)"
      End
      Begin VB.Menu mnuCompmgmt 
         Caption         =   "計算機管理(&O)"
      End
      Begin VB.Menu mnuGpedit 
         Caption         =   "組策略編輯器(&G)"
      End
      Begin VB.Menu mnuMMC 
         Caption         =   "系統管理控制台(&A)"
      End
      Begin VB.Menu mnuWinVer 
         Caption         =   "Windows版本(&V)"
      End
      Begin VB.Menu mnuTaskmgr 
         Caption         =   "任務管理器(&T)"
      End
      Begin VB.Menu mnuSyskey 
         Caption         =   "用戶數據庫加密工具(&S)"
      End
      Begin VB.Menu mnuCharmap 
         Caption         =   "字符映射表(&H)"
      End
      Begin VB.Menu mnuSysedit 
         Caption         =   "系統配置文件編輯器(&E)"
      End
      Begin VB.Menu mnuDxDiag 
         Caption         =   "DirectX診斷工具(&D)"
      End
      Begin VB.Menu mnuCleanmgr 
         Caption         =   "Windows磁盤清理(&L)"
      End
      Begin VB.Menu mnuRegedit 
         Caption         =   "註冊表編輯器(&R)"
      End
      Begin VB.Menu mnuMsconfig 
         Caption         =   "Windows系統配置實用程序(&F)"
      End
      Begin VB.Menu mnuExplorer 
         Caption         =   "資源管理器(&X)"
      End
      Begin VB.Menu mnuDevMgmt 
         Caption         =   "設備管理器(&Q)"
      End
      Begin VB.Menu mnuEvent 
         Caption         =   "事件查看器(&N)"
      End
      Begin VB.Menu mnuDiskMgmt 
         Caption         =   "磁盤管理工具(&I)"
      End
      Begin VB.Menu mnuServices 
         Caption         =   "服務管理器(&U)"
      End
      Begin VB.Menu mnuTasks 
         Caption         =   "任務計劃管理器(&J)"
      End
      Begin VB.Menu mnuCert 
         Caption         =   "證書管理器(&Z)"
      End
      Begin VB.Menu mnuPerform 
         Caption         =   "性能監視器(&P)"
      End
      Begin VB.Menu mnuUsers 
         Caption         =   "本地用戶和組管理器(&Y)"
      End
      Begin VB.Menu mnuShare 
         Caption         =   "共享内容管理器(&F)"
      End
      Begin VB.Menu mnuCOM 
         Caption         =   "組件服務管理器(&K)"
      End
      Begin VB.Menu mnuSysInfo 
         Caption         =   "顯示系統信息(&S)"
      End
      Begin VB.Menu mnuDS 
         Caption         =   "定時關機工具(&G)"
      End
      Begin VB.Menu mnuAXA 
         Caption         =   "ActiveX控件/DLL助手(&A)"
      End
      Begin VB.Menu mnuFileInfo 
         Caption         =   "文件信息工具(&F)"
      End
      Begin VB.Menu mnuLargeFile 
         Caption         =   "大文件生成工具(&D)"
      End
      Begin VB.Menu mnuSysInfoV 
         Caption         =   "系統信息查看器(&S)"
      End
      Begin VB.Menu mnuAppwiz 
         Caption         =   "添加或刪除程序(&P)"
      End
      Begin VB.Menu mnuMblctr 
         Caption         =   "Windows顯示控制面板(&M)"
      End
      Begin VB.Menu mnuRc 
         Caption         =   "Windows遠程協助程序(&R)"
      End
      Begin VB.Menu mnuColor 
         Caption         =   "Windows顔色控制面板(&C)"
      End
      Begin VB.Menu mnuDia 
         Caption         =   "遠程桌面連接(&D)"
      End
      Begin VB.Menu mnuAdvUsrMgr 
         Caption         =   "高級用戶賬戶控制面板(&V)"
      End
      Begin VB.Menu mnuAdvSet 
         Caption         =   "高級系統設置(&T)"
      End
      Begin VB.Menu mnuMobile 
         Caption         =   "Windows移動設備中心(&I)"
      End
      Begin VB.Menu mnuRestore 
         Caption         =   "Windows係統還原工具(&E)"
      End
      Begin VB.Menu mnuSmartScreen 
         Caption         =   "SmartScreen篩選器設置(&S)"
      End
      Begin VB.Menu mnuSound 
         Caption         =   "音量合成器(&V)"
      End
      Begin VB.Menu mnuUAC 
         Caption         =   "用戶賬戶控制UAC設置(&U)"
      End
      Begin VB.Menu mnuPsr 
         Caption         =   "步驟記錄器(&P)"
      End
      Begin VB.Menu mnuRecdisk 
         Caption         =   "恢復驅動器創建工具(&O)"
      End
      Begin VB.Menu mnuMarrator 
         Caption         =   "講述人(&N)"
      End
      Begin VB.Menu mnuOSK 
         Caption         =   "虛擬屏幕鍵盤(&K)"
      End
      Begin VB.Menu mnuMaginfy 
         Caption         =   "屏幕放大鏡(&M)"
      End
      Begin VB.Menu mnuSched 
         Caption         =   "Windows内存診斷工具(&M)"
      End
      Begin VB.Menu mnuRekeywiz 
         Caption         =   "EFS證書管理器(&E)"
      End
      Begin VB.Menu mnuLockKeys 
         Caption         =   "鍵盤/按鍵鎖定工具(&L)"
      End
      Begin VB.Menu mnuKillGray 
         Caption         =   "灰色按鈕破解工具(&G)"
      End
      Begin VB.Menu mnuScreenOff 
         Caption         =   "關閉顯示器(&X)"
      End
      Begin VB.Menu mnuSCR 
         Caption         =   "啟動屏幕保護程序(&Q)"
      End
      Begin VB.Menu mnuTaskApp 
         Caption         =   "任務計劃工具(&R)"
      End
      Begin VB.Menu mnuVClick 
         Caption         =   "虛擬點擊器(&X)"
      End
      Begin VB.Menu mnuMClip 
         Caption         =   "超級剪貼板(&J)"
      End
      Begin VB.Menu b1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuShutdown 
         Caption         =   "關機(&W)"
      End
      Begin VB.Menu mnuReboot 
         Caption         =   "重啟(&B)"
      End
      Begin VB.Menu mnuLogoff 
         Caption         =   "註銷(&Z)"
      End
   End
End
Attribute VB_Name = "LeftBotton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1
Const REG_DWORD = 4
Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"
Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
Private Declare Function SetFileAttributes Lib "kernel32" Alias "SetFileAttributesA" (ByVal lpFileName As String, ByVal dwFileAttributes As Long) As Long
Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Private Const FILE_ATTRIBUTE_ARCHIVE = &H20
Private Const FILE_ATTRIBUTE_COMPRESSED = &H800
Private Const FILE_ATTRIBUTE_DIRECTORY = &H10
Private Const FILE_ATTRIBUTE_HIDDEN = &H2
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const FILE_ATTRIBUTE_READONLY = &H1
Private Const FILE_ATTRIBUTE_SYSTEM = &H4
Private Const FILE_ATTRIBUTE_TEMPORARY = &H100
Dim lpSize As Long
Dim bchk As Boolean
Dim lpFilePath As String
Const MAX_FILE_SIZE = 1.5 * (1024 ^ 3)
Private Const HWND_BOTTOM = 1
Private Const HWND_BROADCAST = &HFFFF&
Private Const HWND_DESKTOP = 0
Private Const HWND_NOTOPMOST = -2
Private Const WS_EX_TRANSPARENT = &H20&
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Private Declare Function ClipCursor Lib "user32" (lpRect As Any) As Long
Private Type RECTL
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Const WM_SYSCOMMAND = &H112&
Const SC_SCREENSAVE = &HF140&
Dim cRect As RECT
Const LCR_UNLOCK = 0
Dim dwMouseFlag As Integer
Const ME_LBCLICK = 1
Const ME_LBDBLCLICK = 2
Const ME_RBCLICK = 3
Private Const MOUSEEVENTF_LEFTDOWN = &H2
Private Const MOUSEEVENTF_ABSOLUTE = &H8000
Private Const MOUSEEVENTF_LEFTUP = &H4
Private Const MOUSEEVENTF_MIDDLEDOWN = &H20
Private Const MOUSEEVENTF_MIDDLEUP = &H40
Private Const MOUSEEVENTF_MOVE = &H1
Private Const MOUSEEVENTF_RIGHTDOWN = &H8
Private Const MOUSEEVENTF_RIGHTUP = &H10
Private Const MOUSETRAILS = 39
Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private Declare Function GetKeyboardState Lib "user32" (pbKeyState As Byte) As Long
Const SWP_NOACTIVATE = &H10
Private Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Dim lpszCaptionNew As String
Private Const SC_MINIMIZE = &HF020&
Private Const WS_MAXIMIZEBOX = &H10000
Dim HKStateCtrl As Integer
Dim HKStateFn As Integer
Private Const WS_MINIMIZEBOX = &H20000
Private Const WS_MAXIMIZE = &H1000000
Private Const WS_MINIMIZE = &H20000000
Private Const WS_ICONIC = WS_MINIMIZE
Const SC_ICON = SC_MINIMIZE
Const SC_TASKLIST = &HF130&
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Dim bCodeUse As Boolean
Private Const WS_CAPTION = &HC00000
Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Const PROCESS_ALL_ACCESS = &H1F0FFF
Private Type PROCESSENTRY32
dwSize As Long
cntUsage As Long
th32ProcessID As Long
th32DefaultHeapID As Long
th32ModuleID As Long
cntThreads As Long
th32ParentProcessID As Long
pcPriClassBase As Long
dwFlags As Long
szExeFile As String * 1024
End Type
Const SC_RESTORE = &HF120&
Const TH32CS_SNAPHEAPLIST = &H1
Const TH32CS_SNAPPROCESS = &H2
Const TH32CS_SNAPTHREAD = &H4
Const TH32CS_SNAPMODULE = &H8
Const TH32CS_SNAPALL = (TH32CS_SNAPHEAPLIST Or TH32CS_SNAPPROCESS Or TH32CS_SNAPTHREAD Or TH32CS_SNAPMODULE)
Const TH32CS_INHERIT = &H80000000
Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal dwFlags As Long, ByVal th32ProcessID As Long) As Long
Private Declare Function Process32First Lib "kernel32" (ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long
Private Declare Function Process32Next Lib "kernel32" (ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Dim lMeWinStyle As Long
Const SWP_SHOWWINDOW = &H40
Const SWP_HIDEWINDOW = &H80
Const SWP_NOOWNERZORDER = &H200
Const SWP_NOREPOSITION = SWP_NOOWNERZORDER
Private Const SC_MOVE = &HF010&
Private Const SC_SIZE = &HF000&
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Const WS_EX_APPWINDOW = &H40000
Private Type WINDOWINFORMATION
hWindow As Long
hWindowDC As Long
hThreadProcess As Long
hThreadProcessID As Long
lpszCaption As String
lpszClassName As String
lpszThreadProcessName As String * 1024
lpszThreadProcessPath As String
lpszExe As String
lpszPath As String
End Type
Private Type WINDOWPARAM
bEnabled As Boolean
bHide As Boolean
bTrans As Boolean
bClosable As Boolean
bSizable As Boolean
bMinisizable As Boolean
bTop As Boolean
lpTransValue As Integer
End Type
Dim lpWindow As WINDOWINFORMATION
Dim lpWindowParam() As WINDOWPARAM
Dim lpCur As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Const WM_GETTEXT = &HD
Private Const WM_GETTEXTLENGTH = &HE
Dim lpRtn As Long
Dim hWindow As Long
Dim lpLength As Long
Dim lpArray() As Byte
Dim lpArray2() As Byte
Dim lpBuff As String
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Const WS_EX_LAYERED = &H80000
Private Const GWL_EXSTYLE = (-20)
Private Const LWA_ALPHA = &H2
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Const LWA_COLORKEY = &H1
Private Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal fEnable As Long) As Long
Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Private Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Const MF_BYPOSITION = &H400&
Private Const MF_REMOVE = &H1000&
Private Const WS_SYSMENU = &H80000
Private Const GWL_STYLE = (-16)
Private Const MF_BYCOMMAND = &H0
Private Const SC_CLOSE = &HF060&
Private Declare Function SetMenu Lib "user32" (ByVal hwnd As Long, ByVal hMenu As Long) As Long
Private Const MF_INSERT = &H0&
Private Const SC_MAXIMIZE = &HF030&
Private Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
Private Type WINDOWINFOBOXDATA
lpszCaption As String
lpszClass As String
lpszThread As String
lpszHandle As String
lpszDC As String
End Type
Dim dwWinInfo As WINDOWINFOBOXDATA
Dim bError As Boolean
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Const HTCAPTION = 2
Private Const WM_NCLBUTTONDOWN = &HA1
Const WM_CLOSE = &H10
Private Const HWND_TOPMOST = -1
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOREDRAW = &H8
Private Const SWP_NOMOVE = &H2
Dim mov As Boolean
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
Const ANYSIZE_ARRAY = 1
Const EWX_LOGOFF = 0
Const EWX_SHUTDOWN = 1
Const EWX_REBOOT = 2
Const EWX_FORCE = 4
Private Type LUID
UsedPart As Long
IgnoredForNowHigh32BitPart As Long
End Type
Private Type TOKEN_PRIVILEGES
PrivilegeCount As Long
TheLuid As LUID
Attributes As Long
End Type
Private Declare Function OpenProcessToken Lib "advapi32" (ByVal _
ProcessHandle As Long, _
ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Private Declare Function LookupPrivilegeValue Lib "advapi32" _
Alias "LookupPrivilegeValueA" _
(ByVal lpSystemName As String, ByVal lpName As String, lpLuid _
As LUID) As Long
Private Declare Function AdjustTokenPrivileges Lib "advapi32" _
(ByVal TokenHandle As Long, _
ByVal DisableAllPrivileges As Long, NewState As TOKEN_PRIVILEGES _
, ByVal BufferLength As Long, _
PreviousState As TOKEN_PRIVILEGES, ReturnLength As Long) As Long
Private Type TestCounter
TimesLeft As Integer
ResetTime As Integer
End Type
Dim PassTest As TestCounter
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Type POINTAPI
X As Long
Y As Long
End Type
Private Const VK_ADD = &H6B
Private Const VK_ATTN = &HF6
Private Const VK_BACK = &H8
Private Const VK_CANCEL = &H3
Private Const VK_CAPITAL = &H14
Private Const VK_CLEAR = &HC
Private Const VK_CONTROL = &H11
Private Const VK_CRSEL = &HF7
Private Const VK_DECIMAL = &H6E
Private Const VK_DELETE = &H2E
Private Const VK_DIVIDE = &H6F
Private Const VK_DOWN = &H28
Private Const VK_END = &H23
Private Const VK_EREOF = &HF9
Private Const VK_ESCAPE = &H1B
Private Const VK_EXECUTE = &H2B
Private Const VK_EXSEL = &HF8
Private Const VK_F1 = &H70
Private Const VK_F10 = &H79
Private Const VK_F11 = &H7A
Private Const VK_F12 = &H7B
Private Const VK_F13 = &H7C
Private Const VK_F14 = &H7D
Private Const VK_F15 = &H7E
Private Const VK_F16 = &H7F
Private Const VK_F17 = &H80
Private Const VK_F18 = &H81
Private Const VK_F19 = &H82
Private Const VK_F2 = &H71
Private Const VK_F20 = &H83
Private Const VK_F21 = &H84
Private Const VK_F22 = &H85
Private Const VK_F23 = &H86
Private Const VK_F24 = &H87
Private Const VK_F3 = &H72
Private Const VK_F4 = &H73
Private Const VK_F5 = &H74
Private Const VK_F6 = &H75
Private Const VK_F7 = &H76
Private Const VK_F8 = &H77
Private Const VK_F9 = &H78
Private Const VK_HELP = &H2F
Private Const VK_HOME = &H24
Private Const VK_INSERT = &H2D
Private Const VK_LBUTTON = &H1
Private Const VK_LCONTROL = &HA2
Private Const VK_LEFT = &H25
Private Const VK_LMENU = &HA4
Private Const VK_LSHIFT = &HA0
Private Const VK_MBUTTON = &H4
Private Const VK_MENU = &H12
Private Const VK_MULTIPLY = &H6A
Private Const VK_NEXT = &H22
Private Const VK_NONAME = &HFC
Private Const VK_NUMLOCK = &H90
Private Const VK_NUMPAD0 = &H60
Private Const VK_NUMPAD1 = &H61
Private Const VK_NUMPAD2 = &H62
Private Const VK_NUMPAD3 = &H63
Private Const VK_NUMPAD4 = &H64
Private Const VK_NUMPAD5 = &H65
Private Const VK_NUMPAD6 = &H66
Private Const VK_NUMPAD7 = &H67
Private Const VK_NUMPAD8 = &H68
Private Const VK_NUMPAD9 = &H69
Private Const VK_OEM_CLEAR = &HFE
Private Const VK_PA1 = &HFD
Private Const VK_PAUSE = &H13
Private Const VK_PLAY = &HFA
Private Const VK_PRINT = &H2A
Private Const VK_PRIOR = &H21
Private Const VK_PROCESSKEY = &HE5
Private Const VK_RBUTTON = &H2
Private Const VK_RCONTROL = &HA3
Private Const VK_RETURN = &HD
Private Const VK_RIGHT = &H27
Private Const VK_RMENU = &HA5
Private Const VK_RSHIFT = &HA1
Private Const VK_SCROLL = &H91
Private Const VK_SELECT = &H29
Private Const VK_SEPARATOR = &H6C
Private Const VK_SHIFT = &H10
Private Const VK_SNAPSHOT = &H2C
Private Const VK_SPACE = &H20
Private Const VK_SUBTRACT = &H6D
Private Const VK_TAB = &H9
Private Const VK_UP = &H26
Private Const VK_ZOOM = &HFB
Private Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Dim lpX As Long
Dim lpY As Long
Private Type FILEINFO
lpPath As String
lpDateLastChanged As Date
lpAttribList As Integer
lpSize As Long
lpHeader As String * 25
lpType As String
lpAttrib As String
End Type
Dim lpFile As FILEINFO
Public act As Boolean
Dim regsvrvrt
Dim unregsvrvrt
Dim regflag As Boolean
Dim unregflag  As Boolean
Dim ream
Private Type BROWSEINFO
hOwner As Long
pidlRoot As Long
pszDisplayName As String
lpszTitle As String
ulFlags As Long
lpfn As Long
lParam As Long
iImage As Long
End Type
Private Const BIF_RETURNONLYFSDIRS = &H1
Private Const BIF_NEWDIALOGSTYLE = &H40
Private Const BIF_NONEWFOLDERBUTTON = &H200
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" _
(ByVal pidl As Long, _
ByVal pszPath As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" _
(lpBrowseInfo As BROWSEINFO) As Long
Private Declare Function CloseScreenFun Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Const SC_MONITORPOWER = &HF170&
Private Sub CloseScreenA(ByVal sWitch As Boolean)
If sWitch = True Then
CloseScreenFun GetForegroundWindow, WM_SYSCOMMAND, SC_MONITORPOWER, 1&
Else
CloseScreenFun GetForegroundWindow, WM_SYSCOMMAND, SC_MONITORPOWER, -1&
End If
End Sub
Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
On Error Resume Next
Dim i As Long
Dim rc As Long
Dim hKey As Long
Dim hDepth As Long
Dim KeyValType As Long
Dim tmpVal As String
Dim KeyValSize As Long
rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey)
If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError
tmpVal = String$(1024, 0)
KeyValSize = 1024
rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
 KeyValType, tmpVal, KeyValSize)
If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError
If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then
tmpVal = Left(tmpVal, KeyValSize - 1)
Else
tmpVal = Left(tmpVal, KeyValSize)
End If
Select Case KeyValType
Case REG_SZ
KeyVal = tmpVal
Case REG_DWORD
For i = Len(tmpVal) To 1 Step -1
KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))
Next
End Select
GetKeyValue = True
rc = RegCloseKey(hKey)
Exit Function
GetKeyError:
KeyVal = ""
GetKeyValue = False
rc = RegCloseKey(hKey)
End Function
Public Function GetFolderName(hwnd As Long, Text As String) As String
On Error Resume Next
Dim bi As BROWSEINFO
Dim pidl As Long
Dim path As String
With bi
.hOwner = hwnd
.pidlRoot = 0&
.lpszTitle = Text
.ulFlags = BIF_NONEWFOLDERBUTTON
End With
pidl = SHBrowseForFolder(bi)
path = Space$(512)
If SHGetPathFromIDList(ByVal pidl, ByVal path) Then
GetFolderName = Left(path, InStr(path, Chr(0)) - 1)
End If
End Function
Sub GetProcessName(ByVal processID As Long, szExeName As String, szPathName As String)
On Error Resume Next
Dim my As PROCESSENTRY32
Dim hProcessHandle As Long
Dim success As Long
Dim l As Long
l = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
If l Then
my.dwSize = 1060
If (Process32First(l, my)) Then
Do
If my.th32ProcessID = processID Then
CloseHandle l
szExeName = Left$(my.szExeFile, InStr(1, my.szExeFile, Chr$(0)) - 1)
For l = Len(szExeName) To 1 Step -1
If Mid$(szExeName, l, 1) = "\" Then
Exit For
End If
Next l
szPathName = Left$(szExeName, l)
Exit Sub
End If
Loop Until (Process32Next(l, my) < 1)
End If
CloseHandle l
End If
End Sub
Private Sub CreateFile(lpPath As String, lpSize As Long)
On Error Resume Next
End Sub
Private Sub DisableClose(hwnd As Long, Optional ByVal MDIChild As Boolean)
On Error Resume Next
Exit Sub
Dim hSysMenu As Long
Dim nCnt As Long
Dim cID As Long
hSysMenu = GetSystemMenu(hwnd, False)
If hSysMenu = 0 Then
Exit Sub
End If
nCnt = GetMenuItemCount(hSysMenu)
If MDIChild Then
cID = 3
Else
cID = 1
End If
If nCnt Then
RemoveMenu hSysMenu, nCnt - cID, MF_BYPOSITION Or MF_REMOVE
RemoveMenu hSysMenu, nCnt - cID - 1, MF_BYPOSITION Or MF_REMOVE
DrawMenuBar hwnd
End If
End Sub
Private Function GetPassword(hwnd As Long) As String
On Error Resume Next
lpLength = SendMessage(hwnd, WM_GETTEXTLENGTH, 0, 0)
If lpLength > 0 Then
ReDim lpArray(lpLength + 1) As Byte
ReDim lpArray2(lpLength - 1) As Byte
CopyMemory lpArray(0), lpLength, 2
SendMessage hwnd, WM_GETTEXT, lpLength + 1, lpArray(0)
CopyMemory lpArray2(0), lpArray(0), lpLength
GetPassword = StrConv(lpArray2, vbUnicode)
Else
GetPassword = ""
End If
End Function
Private Function GetWindowClassName(hwnd As Long) As String
On Error Resume Next
Dim lpszWindowClassName As String * 256
lpszWindowClassName = Space(256)
GetClassName hwnd, lpszWindowClassName, 256
lpszWindowClassName = Trim(lpszWindowClassName)
GetWindowClassName = lpszWindowClassName
End Function
Private Sub AdjustToken()
On Error Resume Next
Const TOKEN_ADJUST_PRIVILEGES = &H20
Const TOKEN_QUERY = &H8
Const SE_PRIVILEGE_ENABLED = &H2
Dim hdlProcessHandle As Long
Dim hdlTokenHandle As Long
Dim tmpLuid As LUID
Dim tkp As TOKEN_PRIVILEGES
Dim tkpNewButIgnored As TOKEN_PRIVILEGES
Dim lBufferNeeded As Long
hdlProcessHandle = GetCurrentProcess()
OpenProcessToken hdlProcessHandle, (TOKEN_ADJUST_PRIVILEGES Or _
TOKEN_QUERY), hdlTokenHandle
LookupPrivilegeValue "", "SeShutdownPrivilege", tmpLuid
tkp.PrivilegeCount = 1
tkp.TheLuid = tmpLuid
tkp.Attributes = SE_PRIVILEGE_ENABLED
AdjustTokenPrivileges hdlTokenHandle, False, _
tkp, Len(tkpNewButIgnored), tkpNewButIgnored, lBufferNeeded
End Sub
Private Function HexOpen(lpFilePath As String, bSafe As Boolean) As String
Dim strFileName As String
Dim arr() As Byte
strFileName = App.path & "\2.jpg"
Open lpFilePath For Binary As #1
ReDim arr(LOF(1))
Get #1, , arr()
Close #1
Dim t As String
Dim l As Integer
Dim te As String
Dim ASCII As String
l = 0
t = ""
te = ""
ASCII = ""
Dim i
For i = LBound(arr) To UBound(arr)
te = UCase(Hex$(arr(i)))
If arr(i) >= 32 And arr(i) <= 126 Then
ASCII = ASCII & Chr(arr(i))
Else
ASCII = ASCII & "."
End If
If Len(te) = 1 Then te = "0" & te
t = t & te & " "
l = l + 1
If l = 16 Then
t = t & " "
ASCII = ASCII & " "
End If
If l = 32 Then
't = t & " " & ASCII & vbCrLf
t = t
ASCII = ""
l = 0
End If
If bSafe = True Then
If Len(t) >= 72 Then
t = Left(t, 72)
Exit For
End If
End If
Next
HexOpen = t
End Function
Private Function OpenAsHexDocument(lpFile As String, lpHeadOnly As Boolean) As String
On Error Resume Next
Dim strFileName As String
Dim arr() As Byte
strFileName = lpFile
If 245 = 245 Then
Open strFileName For Binary As #1
ReDim arr(LOF(1))
Get #1, , arr()
Close #1
Dim t As String
Dim l As Integer
Dim te As String
Dim ASCII As String
l = 0
t = ""
te = ""
ASCII = ""
Dim i
For i = LBound(arr) To UBound(arr)
te = UCase(Hex$(arr(i)))
If arr(i) >= 32 And arr(i) <= 126 Then
ASCII = ASCII & Chr(arr(i))
Else
ASCII = ASCII & "."
End If
If Len(te) = 1 Then te = "0" & te
t = t & te & " "
If Len(t) >= 72 And lpHeadOnly = True Then
Exit For
End If
l = l + 1
If l = 16 Then
t = t & " "
ASCII = ASCII & " "
End If
If l = 32 Then
t = t
ASCII = ""
l = 0
End If
Next
End If
If lpHeadOnly = True Then
OpenAsHexDocument = Left(t, 72)
Else
OpenAsHexDocument = t
End If
End Function
Private Sub Form_Activate()
On Error Resume Next
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
Dim rtn As Long
If FormMain.Check1.Value = 1 Then
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
Else
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 25, LWA_ALPHA
End If
End Sub
Private Sub Form_Load()
On Error Resume Next
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
Dim rtn As Long
If FormMain.Check1.Value = 1 Then
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
Else
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 25, LWA_ALPHA
End If
End Sub
Private Sub Form_MouseMoveOld(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Dim lpCommand As String
lpCommand = Me.Tag
Dim ans As Integer
On Error Resume Next
Const TOKEN_ADJUST_PRIVILEGES = &H20
Const TOKEN_QUERY = &H8
Const SE_PRIVILEGE_ENABLED = &H2
Dim hdlProcessHandle As Long
Dim hdlTokenHandle As Long
Dim tmpLuid As LUID
Dim tkp As TOKEN_PRIVILEGES
Dim tkpNewButIgnored As TOKEN_PRIVILEGES
Dim lBufferNeeded As Long
Dim lpForce As Boolean
If FormMain.Check3.Value = 1 Then
lpForce = True
Else
lpForce = False
End If
Select Case lpCommand
Case "Shutdown"
If lpForce = True Then
ans = MsgBox("確定要關機嗎?請保存數據!", vbExclamation + vbYesNo + vbMsgBoxSetForeground, "Ask")
If ans = vbYes Then
hdlProcessHandle = GetCurrentProcess()
OpenProcessToken hdlProcessHandle, (TOKEN_ADJUST_PRIVILEGES Or _
TOKEN_QUERY), hdlTokenHandle
LookupPrivilegeValue "", "SeShutdownPrivilege", tmpLuid
tkp.PrivilegeCount = 1
tkp.TheLuid = tmpLuid
tkp.Attributes = SE_PRIVILEGE_ENABLED
AdjustTokenPrivileges hdlTokenHandle, False, _
tkp, Len(tkpNewButIgnored), tkpNewButIgnored, lBufferNeeded
AdjustToken
ExitWindowsEx EWX_SHUTDOWN Or EWX_FORCE, 0
Else
Exit Sub
End If
Else
ans = MsgBox("確定要關機嗎?請保存數據!", vbExclamation + vbYesNo + vbMsgBoxSetForeground, "Ask")
If ans = vbYes Then
hdlProcessHandle = GetCurrentProcess()
OpenProcessToken hdlProcessHandle, (TOKEN_ADJUST_PRIVILEGES Or _
TOKEN_QUERY), hdlTokenHandle
LookupPrivilegeValue "", "SeShutdownPrivilege", tmpLuid
tkp.PrivilegeCount = 1
tkp.TheLuid = tmpLuid
tkp.Attributes = SE_PRIVILEGE_ENABLED
AdjustTokenPrivileges hdlTokenHandle, False, _
tkp, Len(tkpNewButIgnored), tkpNewButIgnored, lBufferNeeded
AdjustToken
ExitWindowsEx EWX_SHUTDOWN, 0
Else
Exit Sub
End If
End If
Case "Reboot"
If lpForce = True Then
ans = MsgBox("確定要重啟嗎?請保存數據!", vbExclamation + vbYesNo + vbMsgBoxSetForeground, "Ask")
If ans = vbYes Then
hdlProcessHandle = GetCurrentProcess()
OpenProcessToken hdlProcessHandle, (TOKEN_ADJUST_PRIVILEGES Or _
TOKEN_QUERY), hdlTokenHandle
LookupPrivilegeValue "", "SeShutdownPrivilege", tmpLuid
tkp.PrivilegeCount = 1
tkp.TheLuid = tmpLuid
tkp.Attributes = SE_PRIVILEGE_ENABLED
AdjustTokenPrivileges hdlTokenHandle, False, _
tkp, Len(tkpNewButIgnored), tkpNewButIgnored, lBufferNeeded
AdjustToken
ExitWindowsEx EWX_REBOOT Or EWX_FORCE, 0
Else
Exit Sub
End If
Else
ans = MsgBox("確定要重啟嗎?請保存數據!", vbExclamation + vbYesNo + vbMsgBoxSetForeground, "Ask")
If ans = vbYes Then
hdlProcessHandle = GetCurrentProcess()
OpenProcessToken hdlProcessHandle, (TOKEN_ADJUST_PRIVILEGES Or _
TOKEN_QUERY), hdlTokenHandle
LookupPrivilegeValue "", "SeShutdownPrivilege", tmpLuid
tkp.PrivilegeCount = 1
tkp.TheLuid = tmpLuid
tkp.Attributes = SE_PRIVILEGE_ENABLED
AdjustTokenPrivileges hdlTokenHandle, False, _
tkp, Len(tkpNewButIgnored), tkpNewButIgnored, lBufferNeeded
AdjustToken
ExitWindowsEx EWX_REBOOT, 0
Else
Exit Sub
End If
End If
Case "Logoff"
If lpForce = True Then
ans = MsgBox("確定要註銷嗎?請保存數據!", vbExclamation + vbYesNo + vbMsgBoxSetForeground, "Ask")
If ans = vbYes Then
hdlProcessHandle = GetCurrentProcess()
OpenProcessToken hdlProcessHandle, (TOKEN_ADJUST_PRIVILEGES Or _
TOKEN_QUERY), hdlTokenHandle
LookupPrivilegeValue "", "SeShutdownPrivilege", tmpLuid
tkp.PrivilegeCount = 1
tkp.TheLuid = tmpLuid
tkp.Attributes = SE_PRIVILEGE_ENABLED
AdjustTokenPrivileges hdlTokenHandle, False, _
tkp, Len(tkpNewButIgnored), tkpNewButIgnored, lBufferNeeded
AdjustToken
ExitWindowsEx EWX_LOGOFF Or EWX_FORCE, 0
Else
Exit Sub
End If
Else
ans = MsgBox("確定要註銷嗎?請保存數據!", vbExclamation + vbYesNo + vbMsgBoxSetForeground, "Ask")
If ans = vbYes Then
hdlProcessHandle = GetCurrentProcess()
OpenProcessToken hdlProcessHandle, (TOKEN_ADJUST_PRIVILEGES Or _
TOKEN_QUERY), hdlTokenHandle
LookupPrivilegeValue "", "SeShutdownPrivilege", tmpLuid
tkp.PrivilegeCount = 1
tkp.TheLuid = tmpLuid
tkp.Attributes = SE_PRIVILEGE_ENABLED
AdjustTokenPrivileges hdlTokenHandle, False, _
tkp, Len(tkpNewButIgnored), tkpNewButIgnored, lBufferNeeded
AdjustToken
ExitWindowsEx EWX_LOGOFF, 0
Else
Exit Sub
End If
End If
Case "Menu"
PopupMenu Me.UserMenu
Case Else
Exit Sub
End Select
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Dim rtn As Long
Dim lRet As Long
'If Me.Tag = "ScreenProtection" Then
'lRet = SendMessage(Me.hwnd, WM_SYSCOMMAND, SC_SCREENSAVE, 0&)
'End If
If FormMain.Check1.Value = 1 Then
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
Else
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 25, LWA_ALPHA
End If
If Me.Tag = "None" Then
If FormMain.Check1.Value = 1 Then
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED Or WS_EX_TRANSPARENT
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
Else
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED Or WS_EX_TRANSPARENT
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 25, LWA_ALPHA
End If
Else
If FormMain.Check1.Value = 1 Then
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
Else
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 25, LWA_ALPHA
End If
End If
Dim lpCommand As String
lpCommand = Me.Tag
Dim ans As Integer
On Error Resume Next
Const TOKEN_ADJUST_PRIVILEGES = &H20
Const TOKEN_QUERY = &H8
Const SE_PRIVILEGE_ENABLED = &H2
Dim hdlProcessHandle As Long
Dim hdlTokenHandle As Long
Dim tmpLuid As LUID
Dim tkp As TOKEN_PRIVILEGES
Dim tkpNewButIgnored As TOKEN_PRIVILEGES
Dim lBufferNeeded As Long
Dim lpForce As Boolean
If FormMain.Check3.Value = 1 Then
lpForce = True
Else
lpForce = False
End If
Dim lpMsgProm As Boolean
If FormMain.Check2.Value = 1 Then
lpMsgProm = True
Else
lpMsgProm = False
End If
If lpMsgProm = True Then
Select Case lpCommand
Case "ScreenProtection"
lRet = SendMessage(Me.hwnd, WM_SYSCOMMAND, SC_SCREENSAVE, 0&)
Case "ScreenOff"
CloseScreenA True
Case "Shutdown"
If lpForce = True Then
ans = MsgBox("確定要關機嗎?請保存數據!", vbExclamation + vbYesNo + vbMsgBoxSetForeground, "Ask")
If ans = vbYes Then
hdlProcessHandle = GetCurrentProcess()
OpenProcessToken hdlProcessHandle, (TOKEN_ADJUST_PRIVILEGES Or _
TOKEN_QUERY), hdlTokenHandle
LookupPrivilegeValue "", "SeShutdownPrivilege", tmpLuid
tkp.PrivilegeCount = 1
tkp.TheLuid = tmpLuid
tkp.Attributes = SE_PRIVILEGE_ENABLED
AdjustTokenPrivileges hdlTokenHandle, False, _
tkp, Len(tkpNewButIgnored), tkpNewButIgnored, lBufferNeeded
AdjustToken
ExitWindowsEx EWX_SHUTDOWN Or EWX_FORCE, 0
Else
Exit Sub
End If
Else
ans = MsgBox("確定要關機嗎?請保存數據!", vbExclamation + vbYesNo + vbMsgBoxSetForeground, "Ask")
If ans = vbYes Then
hdlProcessHandle = GetCurrentProcess()
OpenProcessToken hdlProcessHandle, (TOKEN_ADJUST_PRIVILEGES Or _
TOKEN_QUERY), hdlTokenHandle
LookupPrivilegeValue "", "SeShutdownPrivilege", tmpLuid
tkp.PrivilegeCount = 1
tkp.TheLuid = tmpLuid
tkp.Attributes = SE_PRIVILEGE_ENABLED
AdjustTokenPrivileges hdlTokenHandle, False, _
tkp, Len(tkpNewButIgnored), tkpNewButIgnored, lBufferNeeded
AdjustToken
ExitWindowsEx EWX_SHUTDOWN, 0
Else
Exit Sub
End If
End If
Case "Reboot"
If lpForce = True Then
ans = MsgBox("確定要重啟嗎?請保存數據!", vbExclamation + vbYesNo + vbMsgBoxSetForeground, "Ask")
If ans = vbYes Then
hdlProcessHandle = GetCurrentProcess()
OpenProcessToken hdlProcessHandle, (TOKEN_ADJUST_PRIVILEGES Or _
TOKEN_QUERY), hdlTokenHandle
LookupPrivilegeValue "", "SeShutdownPrivilege", tmpLuid
tkp.PrivilegeCount = 1
tkp.TheLuid = tmpLuid
tkp.Attributes = SE_PRIVILEGE_ENABLED
AdjustTokenPrivileges hdlTokenHandle, False, _
tkp, Len(tkpNewButIgnored), tkpNewButIgnored, lBufferNeeded
AdjustToken
ExitWindowsEx EWX_REBOOT Or EWX_FORCE, 0
Else
Exit Sub
End If
Else
ans = MsgBox("確定要重啟嗎?請保存數據!", vbExclamation + vbYesNo + vbMsgBoxSetForeground, "Ask")
If ans = vbYes Then
hdlProcessHandle = GetCurrentProcess()
OpenProcessToken hdlProcessHandle, (TOKEN_ADJUST_PRIVILEGES Or _
TOKEN_QUERY), hdlTokenHandle
LookupPrivilegeValue "", "SeShutdownPrivilege", tmpLuid
tkp.PrivilegeCount = 1
tkp.TheLuid = tmpLuid
tkp.Attributes = SE_PRIVILEGE_ENABLED
AdjustTokenPrivileges hdlTokenHandle, False, _
tkp, Len(tkpNewButIgnored), tkpNewButIgnored, lBufferNeeded
AdjustToken
ExitWindowsEx EWX_REBOOT, 0
Else
Exit Sub
End If
End If
Case "Logoff"
If lpForce = True Then
ans = MsgBox("確定要註銷嗎?請保存數據!", vbExclamation + vbYesNo + vbMsgBoxSetForeground, "Ask")
If ans = vbYes Then
hdlProcessHandle = GetCurrentProcess()
OpenProcessToken hdlProcessHandle, (TOKEN_ADJUST_PRIVILEGES Or _
TOKEN_QUERY), hdlTokenHandle
LookupPrivilegeValue "", "SeShutdownPrivilege", tmpLuid
tkp.PrivilegeCount = 1
tkp.TheLuid = tmpLuid
tkp.Attributes = SE_PRIVILEGE_ENABLED
AdjustTokenPrivileges hdlTokenHandle, False, _
tkp, Len(tkpNewButIgnored), tkpNewButIgnored, lBufferNeeded
AdjustToken
ExitWindowsEx EWX_LOGOFF Or EWX_FORCE, 0
Else
Exit Sub
End If
Else
ans = MsgBox("確定要註銷嗎?請保存數據!", vbExclamation + vbYesNo + vbMsgBoxSetForeground, "Ask")
If ans = vbYes Then
hdlProcessHandle = GetCurrentProcess()
OpenProcessToken hdlProcessHandle, (TOKEN_ADJUST_PRIVILEGES Or _
TOKEN_QUERY), hdlTokenHandle
LookupPrivilegeValue "", "SeShutdownPrivilege", tmpLuid
tkp.PrivilegeCount = 1
tkp.TheLuid = tmpLuid
tkp.Attributes = SE_PRIVILEGE_ENABLED
AdjustTokenPrivileges hdlTokenHandle, False, _
tkp, Len(tkpNewButIgnored), tkpNewButIgnored, lBufferNeeded
AdjustToken
ExitWindowsEx EWX_LOGOFF, 0
Else
Exit Sub
End If
End If
Case "Menu"
PopupMenu Me.UserMenu
Case Else
Exit Sub
End Select
ElseIf lpMsgProm = False Then
Select Case lpCommand
Case "ScreenOff"
CloseScreenA True
Case "ScreenProtection"
lRet = SendMessage(Me.hwnd, WM_SYSCOMMAND, SC_SCREENSAVE, 0&)
Case "Shutdown"
If lpForce = True Then
ans = vbYes
If ans = vbYes Then
hdlProcessHandle = GetCurrentProcess()
OpenProcessToken hdlProcessHandle, (TOKEN_ADJUST_PRIVILEGES Or _
TOKEN_QUERY), hdlTokenHandle
LookupPrivilegeValue "", "SeShutdownPrivilege", tmpLuid
tkp.PrivilegeCount = 1
tkp.TheLuid = tmpLuid
tkp.Attributes = SE_PRIVILEGE_ENABLED
AdjustTokenPrivileges hdlTokenHandle, False, _
tkp, Len(tkpNewButIgnored), tkpNewButIgnored, lBufferNeeded
AdjustToken
ExitWindowsEx EWX_SHUTDOWN Or EWX_FORCE, 0
Else
Exit Sub
End If
Else
ans = vbYes
If ans = vbYes Then
hdlProcessHandle = GetCurrentProcess()
OpenProcessToken hdlProcessHandle, (TOKEN_ADJUST_PRIVILEGES Or _
TOKEN_QUERY), hdlTokenHandle
LookupPrivilegeValue "", "SeShutdownPrivilege", tmpLuid
tkp.PrivilegeCount = 1
tkp.TheLuid = tmpLuid
tkp.Attributes = SE_PRIVILEGE_ENABLED
AdjustTokenPrivileges hdlTokenHandle, False, _
tkp, Len(tkpNewButIgnored), tkpNewButIgnored, lBufferNeeded
AdjustToken
ExitWindowsEx EWX_SHUTDOWN, 0
Else
Exit Sub
End If
End If
Case "Reboot"
If lpForce = True Then
ans = vbYes
If ans = vbYes Then
hdlProcessHandle = GetCurrentProcess()
OpenProcessToken hdlProcessHandle, (TOKEN_ADJUST_PRIVILEGES Or _
TOKEN_QUERY), hdlTokenHandle
LookupPrivilegeValue "", "SeShutdownPrivilege", tmpLuid
tkp.PrivilegeCount = 1
tkp.TheLuid = tmpLuid
tkp.Attributes = SE_PRIVILEGE_ENABLED
AdjustTokenPrivileges hdlTokenHandle, False, _
tkp, Len(tkpNewButIgnored), tkpNewButIgnored, lBufferNeeded
AdjustToken
ExitWindowsEx EWX_REBOOT Or EWX_FORCE, 0
Else
Exit Sub
End If
Else
ans = vbYes
If ans = vbYes Then
hdlProcessHandle = GetCurrentProcess()
OpenProcessToken hdlProcessHandle, (TOKEN_ADJUST_PRIVILEGES Or _
TOKEN_QUERY), hdlTokenHandle
LookupPrivilegeValue "", "SeShutdownPrivilege", tmpLuid
tkp.PrivilegeCount = 1
tkp.TheLuid = tmpLuid
tkp.Attributes = SE_PRIVILEGE_ENABLED
AdjustTokenPrivileges hdlTokenHandle, False, _
tkp, Len(tkpNewButIgnored), tkpNewButIgnored, lBufferNeeded
AdjustToken
ExitWindowsEx EWX_REBOOT, 0
Else
Exit Sub
End If
End If
Case "Logoff"
If lpForce = True Then
ans = vbYes
If ans = vbYes Then
hdlProcessHandle = GetCurrentProcess()
OpenProcessToken hdlProcessHandle, (TOKEN_ADJUST_PRIVILEGES Or _
TOKEN_QUERY), hdlTokenHandle
LookupPrivilegeValue "", "SeShutdownPrivilege", tmpLuid
tkp.PrivilegeCount = 1
tkp.TheLuid = tmpLuid
tkp.Attributes = SE_PRIVILEGE_ENABLED
AdjustTokenPrivileges hdlTokenHandle, False, _
tkp, Len(tkpNewButIgnored), tkpNewButIgnored, lBufferNeeded
AdjustToken
ExitWindowsEx EWX_LOGOFF Or EWX_FORCE, 0
Else
Exit Sub
End If
Else
ans = vbYes
If ans = vbYes Then
hdlProcessHandle = GetCurrentProcess()
OpenProcessToken hdlProcessHandle, (TOKEN_ADJUST_PRIVILEGES Or _
TOKEN_QUERY), hdlTokenHandle
LookupPrivilegeValue "", "SeShutdownPrivilege", tmpLuid
tkp.PrivilegeCount = 1
tkp.TheLuid = tmpLuid
tkp.Attributes = SE_PRIVILEGE_ENABLED
AdjustTokenPrivileges hdlTokenHandle, False, _
tkp, Len(tkpNewButIgnored), tkpNewButIgnored, lBufferNeeded
AdjustToken
ExitWindowsEx EWX_LOGOFF, 0
Else
Exit Sub
End If
End If
Case "Menu"
PopupMenu Me.UserMenu
Case Else
Exit Sub
End Select
Else
Select Case lpCommand
Case "Shutdown"
If lpForce = True Then
ans = MsgBox("確定要關機嗎?請保存數據!", vbExclamation + vbYesNo + vbMsgBoxSetForeground, "Ask")
If ans = vbYes Then
hdlProcessHandle = GetCurrentProcess()
OpenProcessToken hdlProcessHandle, (TOKEN_ADJUST_PRIVILEGES Or _
TOKEN_QUERY), hdlTokenHandle
LookupPrivilegeValue "", "SeShutdownPrivilege", tmpLuid
tkp.PrivilegeCount = 1
tkp.TheLuid = tmpLuid
tkp.Attributes = SE_PRIVILEGE_ENABLED
AdjustTokenPrivileges hdlTokenHandle, False, _
tkp, Len(tkpNewButIgnored), tkpNewButIgnored, lBufferNeeded
AdjustToken
ExitWindowsEx EWX_SHUTDOWN Or EWX_FORCE, 0
Else
Exit Sub
End If
Else
ans = MsgBox("確定要關機嗎?請保存數據!", vbExclamation + vbYesNo + vbMsgBoxSetForeground, "Ask")
If ans = vbYes Then
hdlProcessHandle = GetCurrentProcess()
OpenProcessToken hdlProcessHandle, (TOKEN_ADJUST_PRIVILEGES Or _
TOKEN_QUERY), hdlTokenHandle
LookupPrivilegeValue "", "SeShutdownPrivilege", tmpLuid
tkp.PrivilegeCount = 1
tkp.TheLuid = tmpLuid
tkp.Attributes = SE_PRIVILEGE_ENABLED
AdjustTokenPrivileges hdlTokenHandle, False, _
tkp, Len(tkpNewButIgnored), tkpNewButIgnored, lBufferNeeded
AdjustToken
ExitWindowsEx EWX_SHUTDOWN, 0
Else
Exit Sub
End If
End If
Case "Reboot"
If lpForce = True Then
ans = MsgBox("確定要重啟嗎?請保存數據!", vbExclamation + vbYesNo + vbMsgBoxSetForeground, "Ask")
If ans = vbYes Then
hdlProcessHandle = GetCurrentProcess()
OpenProcessToken hdlProcessHandle, (TOKEN_ADJUST_PRIVILEGES Or _
TOKEN_QUERY), hdlTokenHandle
LookupPrivilegeValue "", "SeShutdownPrivilege", tmpLuid
tkp.PrivilegeCount = 1
tkp.TheLuid = tmpLuid
tkp.Attributes = SE_PRIVILEGE_ENABLED
AdjustTokenPrivileges hdlTokenHandle, False, _
tkp, Len(tkpNewButIgnored), tkpNewButIgnored, lBufferNeeded
AdjustToken
ExitWindowsEx EWX_REBOOT Or EWX_FORCE, 0
Else
Exit Sub
End If
Else
ans = MsgBox("確定要重啟嗎?請保存數據!", vbExclamation + vbYesNo + vbMsgBoxSetForeground, "Ask")
If ans = vbYes Then
hdlProcessHandle = GetCurrentProcess()
OpenProcessToken hdlProcessHandle, (TOKEN_ADJUST_PRIVILEGES Or _
TOKEN_QUERY), hdlTokenHandle
LookupPrivilegeValue "", "SeShutdownPrivilege", tmpLuid
tkp.PrivilegeCount = 1
tkp.TheLuid = tmpLuid
tkp.Attributes = SE_PRIVILEGE_ENABLED
AdjustTokenPrivileges hdlTokenHandle, False, _
tkp, Len(tkpNewButIgnored), tkpNewButIgnored, lBufferNeeded
AdjustToken
ExitWindowsEx EWX_REBOOT, 0
Else
Exit Sub
End If
End If
Case "Logoff"
If lpForce = True Then
ans = MsgBox("確定要註銷嗎?請保存數據!", vbExclamation + vbYesNo + vbMsgBoxSetForeground, "Ask")
If ans = vbYes Then
hdlProcessHandle = GetCurrentProcess()
OpenProcessToken hdlProcessHandle, (TOKEN_ADJUST_PRIVILEGES Or _
TOKEN_QUERY), hdlTokenHandle
LookupPrivilegeValue "", "SeShutdownPrivilege", tmpLuid
tkp.PrivilegeCount = 1
tkp.TheLuid = tmpLuid
tkp.Attributes = SE_PRIVILEGE_ENABLED
AdjustTokenPrivileges hdlTokenHandle, False, _
tkp, Len(tkpNewButIgnored), tkpNewButIgnored, lBufferNeeded
AdjustToken
ExitWindowsEx EWX_LOGOFF Or EWX_FORCE, 0
Else
Exit Sub
End If
Else
ans = MsgBox("確定要註銷嗎?請保存數據!", vbExclamation + vbYesNo + vbMsgBoxSetForeground, "Ask")
If ans = vbYes Then
hdlProcessHandle = GetCurrentProcess()
OpenProcessToken hdlProcessHandle, (TOKEN_ADJUST_PRIVILEGES Or _
TOKEN_QUERY), hdlTokenHandle
LookupPrivilegeValue "", "SeShutdownPrivilege", tmpLuid
tkp.PrivilegeCount = 1
tkp.TheLuid = tmpLuid
tkp.Attributes = SE_PRIVILEGE_ENABLED
AdjustTokenPrivileges hdlTokenHandle, False, _
tkp, Len(tkpNewButIgnored), tkpNewButIgnored, lBufferNeeded
AdjustToken
ExitWindowsEx EWX_LOGOFF, 0
Else
Exit Sub
End If
End If
Case "Menu"
PopupMenu Me.UserMenu
Case Else
Exit Sub
End Select
End If
End Sub
Private Sub mnuCert_Click()
On Error GoTo ep
Shell "mmc.exe -k certmgr.msc", vbNormalFocus
Exit Sub
ep:
MsgBox "发生系统错误:" & Chr(13) & Err.Description & Chr(13) & "您的系统版本可能不支持此功能或者您尚未安装这个功能.", vbCritical, "Error"
End Sub
Private Sub mnuCOM_Click()
On Error GoTo ep
Shell "mmc.exe -k comexp.msc", vbNormalFocus
Exit Sub
ep:
MsgBox "发生系统错误:" & Chr(13) & Err.Description & Chr(13) & "您的系统版本可能不支持此功能或者您尚未安装这个功能.", vbCritical, "Error"
End Sub
Private Sub mnuDevMgmt_Click()
On Error GoTo ep
Shell "mmc.exe -k devmgmt.msc", vbNormalFocus
Exit Sub
ep:
MsgBox "发生系统错误:" & Chr(13) & Err.Description & Chr(13) & "您的系统版本可能不支持此功能或者您尚未安装这个功能.", vbCritical, "Error"
End Sub
Private Sub mnuDiskMgmt_Click()
On Error GoTo ep
Shell "mmc.exe -k diskmgmt.msc", vbNormalFocus
Exit Sub
ep:
MsgBox "发生系统错误:" & Chr(13) & Err.Description & Chr(13) & "您的系统版本可能不支持此功能或者您尚未安装这个功能.", vbCritical, "Error"
End Sub
Private Sub mnuEnable_Click()
On Error Resume Next
On Error Resume Next
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
On Error Resume Next
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
On Error Resume Next
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Select Case Me.Tag
Case False
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Load LeftBotton
LeftBotton.Show
Load LeftTop
LeftTop.Show
Load RightBotton
RightBotton.Show
Load RightTop
RightTop.Show
With LeftBotton
.Left = 0
.Width = CInt(FormMain.Text1.Text) * Screen.TwipsPerPixelX
.Height = CInt(FormMain.Text2.Text) * Screen.TwipsPerPixelY
.Top = Screen.Height - .Height
.BackColor = RGB(0, 245, 245)
End With
With LeftTop
.Left = 0
.Width = CInt(FormMain.Text1.Text) * Screen.TwipsPerPixelX
.Height = CInt(FormMain.Text2.Text) * Screen.TwipsPerPixelY
.Top = 0
.BackColor = RGB(0, 245, 245)
End With
With RightTop
.Width = CInt(FormMain.Text1.Text) * Screen.TwipsPerPixelX
.Left = Screen.Width - .Width
.Height = CInt(FormMain.Text2.Text) * Screen.TwipsPerPixelY
.Top = 0
.BackColor = RGB(0, 245, 245)
End With
With RightBotton
.Width = CInt(FormMain.Text1.Text) * Screen.TwipsPerPixelX
.Height = CInt(FormMain.Text2.Text) * Screen.TwipsPerPixelY
.Top = Screen.Height - .Height
.Left = Screen.Width - .Width
.BackColor = RGB(0, 245, 245)
End With
On Error Resume Next
Dim rtn As Long
If FormMain.Check1.Value = 1 Then
rtn = GetWindowLong(LeftBotton.hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong LeftBotton.hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes LeftBotton.hwnd, 0, 192, LWA_ALPHA
Else
rtn = GetWindowLong(LeftBotton.hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong LeftBotton.hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes LeftBotton.hwnd, 0, 25, LWA_ALPHA
End If
If FormMain.Check1.Value = 1 Then
rtn = GetWindowLong(LeftTop.hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong LeftTop.hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes LeftTop.hwnd, 0, 192, LWA_ALPHA
Else
rtn = GetWindowLong(LeftTop.hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong LeftTop.hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes LeftTop.hwnd, 0, 25, LWA_ALPHA
End If
If FormMain.Check1.Value = 1 Then
rtn = GetWindowLong(RightTop.hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong RightTop.hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes RightTop.hwnd, 0, 192, LWA_ALPHA
Else
rtn = GetWindowLong(RightTop.hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong RightTop.hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes RightTop.hwnd, 0, 25, LWA_ALPHA
End If
If FormMain.Check1.Value = 1 Then
rtn = GetWindowLong(RightBotton.hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong RightBotton.hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes RightBotton.hwnd, 0, 192, LWA_ALPHA
Else
rtn = GetWindowLong(RightBotton.hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong RightBotton.hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes RightBotton.hwnd, 0, 25, LWA_ALPHA
End If
On Error Resume Next
Select Case FormMain.Combo1.ListIndex
Case 0
LeftTop.Tag = "Shutdown"
Case 1
LeftTop.Tag = "Reboot"
Case 2
LeftTop.Tag = "Logoff"
Case 3
LeftTop.Tag = "Menu"
Case 4
LeftTop.Tag = "None"
End Select
Select Case FormMain.Combo2.ListIndex
Case 0
LeftBotton.Tag = "Shutdown"
Case 1
LeftBotton.Tag = "Reboot"
Case 2
LeftBotton.Tag = "Logoff"
Case 3
LeftBotton.Tag = "Menu"
Case 4
LeftBotton.Tag = "None"
End Select
Select Case FormMain.Combo3.ListIndex
Case 0
RightTop.Tag = "Shutdown"
Case 1
RightTop.Tag = "Reboot"
Case 2
RightTop.Tag = "Logoff"
Case 3
RightTop.Tag = "Menu"
Case 4
RightTop.Tag = "None"
End Select
Select Case FormMain.Combo4.ListIndex
Case 0
RightBotton.Tag = "Shutdown"
Case 1
RightBotton.Tag = "Reboot"
Case 2
RightBotton.Tag = "Logoff"
Case 3
RightBotton.Tag = "Menu"
Case 4
RightBotton.Tag = "None"
End Select
Debug.Print LeftTop.Tag
Debug.Print LeftBotton.Tag
Debug.Print RightTop.Tag
Debug.Print RightBotton.Tag
Case True
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
On Error Resume Next
If FormMain.Check1.Value = 1 Then
rtn = GetWindowLong(LeftBotton.hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong LeftBotton.hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes LeftBotton.hwnd, 0, 192, LWA_ALPHA
Else
rtn = GetWindowLong(LeftBotton.hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong LeftBotton.hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes LeftBotton.hwnd, 0, 25, LWA_ALPHA
End If
If FormMain.Check1.Value = 1 Then
rtn = GetWindowLong(LeftTop.hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong LeftTop.hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes LeftTop.hwnd, 0, 192, LWA_ALPHA
Else
rtn = GetWindowLong(LeftTop.hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong LeftTop.hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes LeftTop.hwnd, 0, 25, LWA_ALPHA
End If
If FormMain.Check1.Value = 1 Then
rtn = GetWindowLong(RightTop.hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong RightTop.hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes RightTop.hwnd, 0, 192, LWA_ALPHA
Else
rtn = GetWindowLong(RightTop.hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong RightTop.hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes RightTop.hwnd, 0, 25, LWA_ALPHA
End If
If FormMain.Check1.Value = 1 Then
rtn = GetWindowLong(RightBotton.hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong RightBotton.hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes RightBotton.hwnd, 0, 192, LWA_ALPHA
Else
rtn = GetWindowLong(RightBotton.hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong RightBotton.hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes RightBotton.hwnd, 0, 25, LWA_ALPHA
End If
Unload LeftTop
Unload LeftBotton
Unload RightTop
Unload RightBotton
End Select
End Sub
Private Sub mnuEvent_Click()
On Error GoTo ep
Shell "mmc.exe -k eventvwr.msc", vbNormalFocus
Exit Sub
ep:
MsgBox "发生系统错误:" & Chr(13) & Err.Description & Chr(13) & "您的系统版本可能不支持此功能或者您尚未安装这个功能.", vbCritical, "Error"
End Sub
Private Sub mnuExit_Click()
On Error Resume Next
Unload LeftTop
Unload LeftBotton
Unload RightTop
Unload RightBotton
Unload Me
End
End Sub
Private Sub mnuMsconfig_Click()
On Error GoTo ep
Dim lpSysPath As String
lpSysPath = Environ("Windir")
If Right(lpSysPath, 1) = "\" Then
lpSysPath = lpSysPath & "System32"
Else
lpSysPath = lpSysPath & "\System32"
End If
lpSysPath = Trim(lpSysPath)
Shell lpSysPath & "\..\explorer.exe " & lpSysPath & "\msconfig.exe", vbNormalFocus
Exit Sub
ep:
MsgBox "发生系统错误:" & Chr(13) & Err.Description & Chr(13) & "您的系统版本可能不支持此功能或者您尚未安装这个功能.", vbCritical, "Error"
End Sub
Private Sub mnuPerform_Click()
On Error GoTo ep
Shell "mmc.exe -k perfmon.msc", vbNormalFocus
Exit Sub
ep:
MsgBox "发生系统错误:" & Chr(13) & Err.Description & Chr(13) & "您的系统版本可能不支持此功能或者您尚未安装这个功能.", vbCritical, "Error"
End Sub
Private Sub mnuServices_Click()
On Error GoTo ep
Shell "mmc.exe -k services.msc", vbNormalFocus
Exit Sub
ep:
MsgBox "发生系统错误:" & Chr(13) & Err.Description & Chr(13) & "您的系统版本可能不支持此功能或者您尚未安装这个功能.", vbCritical, "Error"
End Sub
Private Sub mnuShare_Click()
On Error GoTo ep
Shell "mmc.exe -k fsmgmt.msc", vbNormalFocus
Exit Sub
ep:
MsgBox "发生系统错误:" & Chr(13) & Err.Description & Chr(13) & "您的系统版本可能不支持此功能或者您尚未安装这个功能.", vbCritical, "Error"
End Sub
Private Sub mnuShow_Click()
On Error Resume Next
With Me
.Show
End With
With FormMain.cSysTray1
.InTray = False
.TrayTip = "Super Cursor - 雙擊還原主窗口,右擊顯示菜單"
End With
End Sub
Private Sub mnuTasks_Click()
On Error GoTo ep
Shell "mmc.exe -k taskschd.msc", vbNormalFocus
Exit Sub
ep:
MsgBox "发生系统错误:" & Chr(13) & Err.Description & Chr(13) & "您的系统版本可能不支持此功能或者您尚未安装这个功能.", vbCritical, "Error"
End Sub
Private Sub mnuTrayMenu_Click()
On Error Resume Next
End Sub
Private Sub mnuUsers_Click()
On Error GoTo ep
Shell "mmc.exe -k lusrmgr.msc", vbNormalFocus
Exit Sub
ep:
MsgBox "发生系统错误:" & Chr(13) & Err.Description & Chr(13) & "您的系统版本可能不支持此功能或者您尚未安装这个功能.", vbCritical, "Error"
End Sub
Private Sub mnuReboot_Click()
On Error Resume Next
Dim lpCommand As String
lpCommand = Me.Tag
Dim ans As Integer
On Error Resume Next
Const TOKEN_ADJUST_PRIVILEGES = &H20
Const TOKEN_QUERY = &H8
Const SE_PRIVILEGE_ENABLED = &H2
Dim hdlProcessHandle As Long
Dim hdlTokenHandle As Long
Dim tmpLuid As LUID
Dim tkp As TOKEN_PRIVILEGES
Dim tkpNewButIgnored As TOKEN_PRIVILEGES
Dim lBufferNeeded As Long
Dim lpForce As Boolean
If FormMain.Check3.Value = 1 Then
lpForce = True
Else
lpForce = False
End If
Dim lpMsgProm As Boolean
If FormMain.Check2.Value = 1 Then
lpMsgProm = True
Else
lpMsgProm = False
End If
If lpForce = True Then
ans = MsgBox("確定要重啟嗎?請保存數據!", vbExclamation + vbYesNo + vbMsgBoxSetForeground, "Ask")
If ans = vbYes Then
hdlProcessHandle = GetCurrentProcess()
OpenProcessToken hdlProcessHandle, (TOKEN_ADJUST_PRIVILEGES Or _
TOKEN_QUERY), hdlTokenHandle
LookupPrivilegeValue "", "SeShutdownPrivilege", tmpLuid
tkp.PrivilegeCount = 1
tkp.TheLuid = tmpLuid
tkp.Attributes = SE_PRIVILEGE_ENABLED
AdjustTokenPrivileges hdlTokenHandle, False, _
tkp, Len(tkpNewButIgnored), tkpNewButIgnored, lBufferNeeded
AdjustToken
ExitWindowsEx EWX_REBOOT Or EWX_FORCE, 0
Else
Exit Sub
End If
Else
ans = MsgBox("確定要重啟嗎?請保存數據!", vbExclamation + vbYesNo + vbMsgBoxSetForeground, "Ask")
If ans = vbYes Then
hdlProcessHandle = GetCurrentProcess()
OpenProcessToken hdlProcessHandle, (TOKEN_ADJUST_PRIVILEGES Or _
TOKEN_QUERY), hdlTokenHandle
LookupPrivilegeValue "", "SeShutdownPrivilege", tmpLuid
tkp.PrivilegeCount = 1
tkp.TheLuid = tmpLuid
tkp.Attributes = SE_PRIVILEGE_ENABLED
AdjustTokenPrivileges hdlTokenHandle, False, _
tkp, Len(tkpNewButIgnored), tkpNewButIgnored, lBufferNeeded
AdjustToken
ExitWindowsEx EWX_REBOOT, 0
Else
Exit Sub
End If
End If
End Sub
Private Sub mnuLogoff_Click()
On Error Resume Next
Dim lpCommand As String
lpCommand = Me.Tag
Dim ans As Integer
On Error Resume Next
Const TOKEN_ADJUST_PRIVILEGES = &H20
Const TOKEN_QUERY = &H8
Const SE_PRIVILEGE_ENABLED = &H2
Dim hdlProcessHandle As Long
Dim hdlTokenHandle As Long
Dim tmpLuid As LUID
Dim tkp As TOKEN_PRIVILEGES
Dim tkpNewButIgnored As TOKEN_PRIVILEGES
Dim lBufferNeeded As Long
Dim lpForce As Boolean
If FormMain.Check3.Value = 1 Then
lpForce = True
Else
lpForce = False
End If
Dim lpMsgProm As Boolean
If FormMain.Check2.Value = 1 Then
lpMsgProm = True
Else
lpMsgProm = False
End If
If lpForce = True Then
ans = MsgBox("確定要註銷嗎?請保存數據!", vbExclamation + vbYesNo + vbMsgBoxSetForeground, "Ask")
If ans = vbYes Then
hdlProcessHandle = GetCurrentProcess()
OpenProcessToken hdlProcessHandle, (TOKEN_ADJUST_PRIVILEGES Or _
TOKEN_QUERY), hdlTokenHandle
LookupPrivilegeValue "", "SeShutdownPrivilege", tmpLuid
tkp.PrivilegeCount = 1
tkp.TheLuid = tmpLuid
tkp.Attributes = SE_PRIVILEGE_ENABLED
AdjustTokenPrivileges hdlTokenHandle, False, _
tkp, Len(tkpNewButIgnored), tkpNewButIgnored, lBufferNeeded
AdjustToken
ExitWindowsEx EWX_LOGOFF Or EWX_FORCE, 0
Else
Exit Sub
End If
Else
ans = MsgBox("確定要註銷嗎?請保存數據!", vbExclamation + vbYesNo + vbMsgBoxSetForeground, "Ask")
If ans = vbYes Then
hdlProcessHandle = GetCurrentProcess()
OpenProcessToken hdlProcessHandle, (TOKEN_ADJUST_PRIVILEGES Or _
TOKEN_QUERY), hdlTokenHandle
LookupPrivilegeValue "", "SeShutdownPrivilege", tmpLuid
tkp.PrivilegeCount = 1
tkp.TheLuid = tmpLuid
tkp.Attributes = SE_PRIVILEGE_ENABLED
AdjustTokenPrivileges hdlTokenHandle, False, _
tkp, Len(tkpNewButIgnored), tkpNewButIgnored, lBufferNeeded
AdjustToken
ExitWindowsEx EWX_LOGOFF, 0
Else
Exit Sub
End If
End If
End Sub
Private Sub mnuShutdown_Click()
On Error Resume Next
Dim lpCommand As String
lpCommand = Me.Tag
Dim ans As Integer
On Error Resume Next
Const TOKEN_ADJUST_PRIVILEGES = &H20
Const TOKEN_QUERY = &H8
Const SE_PRIVILEGE_ENABLED = &H2
Dim hdlProcessHandle As Long
Dim hdlTokenHandle As Long
Dim tmpLuid As LUID
Dim tkp As TOKEN_PRIVILEGES
Dim tkpNewButIgnored As TOKEN_PRIVILEGES
Dim lBufferNeeded As Long
Dim lpForce As Boolean
If FormMain.Check3.Value = 1 Then
lpForce = True
Else
lpForce = False
End If
Dim lpMsgProm As Boolean
If FormMain.Check2.Value = 1 Then
lpMsgProm = True
Else
lpMsgProm = False
End If
If lpForce = True Then
ans = MsgBox("確定要關機嗎?請保存數據!", vbExclamation + vbYesNo + vbMsgBoxSetForeground, "Ask")
If ans = vbYes Then
hdlProcessHandle = GetCurrentProcess()
OpenProcessToken hdlProcessHandle, (TOKEN_ADJUST_PRIVILEGES Or _
TOKEN_QUERY), hdlTokenHandle
LookupPrivilegeValue "", "SeShutdownPrivilege", tmpLuid
tkp.PrivilegeCount = 1
tkp.TheLuid = tmpLuid
tkp.Attributes = SE_PRIVILEGE_ENABLED
AdjustTokenPrivileges hdlTokenHandle, False, _
tkp, Len(tkpNewButIgnored), tkpNewButIgnored, lBufferNeeded
AdjustToken
ExitWindowsEx EWX_SHUTDOWN Or EWX_FORCE, 0
Else
Exit Sub
End If
Else
ans = MsgBox("確定要關機嗎?請保存數據!", vbExclamation + vbYesNo + vbMsgBoxSetForeground, "Ask")
If ans = vbYes Then
hdlProcessHandle = GetCurrentProcess()
OpenProcessToken hdlProcessHandle, (TOKEN_ADJUST_PRIVILEGES Or _
TOKEN_QUERY), hdlTokenHandle
LookupPrivilegeValue "", "SeShutdownPrivilege", tmpLuid
tkp.PrivilegeCount = 1
tkp.TheLuid = tmpLuid
tkp.Attributes = SE_PRIVILEGE_ENABLED
AdjustTokenPrivileges hdlTokenHandle, False, _
tkp, Len(tkpNewButIgnored), tkpNewButIgnored, lBufferNeeded
AdjustToken
ExitWindowsEx EWX_SHUTDOWN, 0
Else
Exit Sub
End If
End If
End Sub
Private Sub mnuDS_Click()
On Error Resume Next
Form1.Show
End Sub
Private Sub mnuSysInfo_Click()
On Error GoTo ep
'Me.Hide
Shell "cmd.exe /c systeminfo.exe > c:\sys.nfo", vbHide
'Do
'MsgBox "請等待10秒...", vbInformation, "Info"
'Loop Until Dir("c:\sys.nfo") <> ""
Sleep 10000
'Me.Show
Shell "notepad.exe c:\sys.nfo", vbNormalFocus
Exit Sub
ep:
MsgBox "发生系统错误:" & Chr(13) & Err.Description & Chr(13) & "您的系统版本可能不支持此功能或者您尚未安装这个功能.", vbCritical, "Error"
End Sub
Private Sub mnuAXA_Click()
On Error Resume Next
frmMain.Show
End Sub
Private Sub mnuSysInfoV_Click()
On Error GoTo ep
Dim rc As Long
Dim SysInfoPath As String
If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then
If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
Else
GoTo ep
End If
Else
GoTo ep
End If
Call Shell(SysInfoPath, vbNormalFocus)
Exit Sub
ep:
MsgBox "发生系统错误:" & Chr(13) & Err.Description & Chr(13) & "您的系统版本可能不支持此功能或者您尚未安装这个功能.", vbCritical, "Error"
End Sub
Public Function GetKeyValueEx(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
On Error Resume Next
Dim i As Long
Dim rc As Long
Dim hKey As Long
Dim hDepth As Long
Dim KeyValType As Long
Dim tmpVal As String
Dim KeyValSize As Long
rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey)
If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError
tmpVal = String$(1024, 0)
KeyValSize = 1024
rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
 KeyValType, tmpVal, KeyValSize)
If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError
If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then
tmpVal = Left(tmpVal, KeyValSize - 1)
Else
tmpVal = Left(tmpVal, KeyValSize)
End If
Select Case KeyValType
Case REG_SZ
KeyVal = tmpVal
Case REG_DWORD
For i = Len(tmpVal) To 1 Step -1
KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))
Next
End Select
GetKeyValueEx = True
rc = RegCloseKey(hKey)
Exit Function
GetKeyError:
KeyVal = ""
GetKeyValueEx = False
rc = RegCloseKey(hKey)
End Function
Private Sub mnuFileInfo_Click()
On Error Resume Next
frmFI.Show
End Sub
Private Sub mnuAdvUsrMgr_Click()
On Error GoTo ep
Dim lpSysPath As String
lpSysPath = Environ("Windir")
If Right(lpSysPath, 1) = "\" Then
lpSysPath = lpSysPath & "System32"
Else
lpSysPath = lpSysPath & "\System32"
End If
Shell lpSysPath & "\Netplwiz.exe", vbNormalFocus
Exit Sub
ep:
MsgBox "发生系统错误:" & Chr(13) & Err.Description & Chr(13) & "您的系统版本可能不支持此功能或者您尚未安装这个功能.", vbCritical, "Error"
End Sub
Private Sub mnuAdvSet_Click()
On Error GoTo ep
Dim lpSysPath As String
lpSysPath = Environ("Windir")
If Right(lpSysPath, 1) = "\" Then
lpSysPath = lpSysPath & "System32"
Else
lpSysPath = lpSysPath & "\System32"
End If
Shell lpSysPath & "\SystemPropertiesAdvanced.exe", vbNormalFocus
Exit Sub
ep:
MsgBox "发生系统错误:" & Chr(13) & Err.Description & Chr(13) & "您的系统版本可能不支持此功能或者您尚未安装这个功能.", vbCritical, "Error"
End Sub
Private Sub mnuAppwiz_Click()
On Error GoTo ep
Dim lpSysPath As String
lpSysPath = Environ("Windir")
If Right(lpSysPath, 1) = "\" Then
lpSysPath = lpSysPath & "System32"
Else
lpSysPath = lpSysPath & "\System32"
End If
ShellExecute Me.hwnd, vbNullString, lpSysPath & "\appwiz.cpl", vbNullString, vbNullString, CLng(vbNormalFocus)
Exit Sub
ep:
MsgBox "发生系统错误:" & Chr(13) & Err.Description & Chr(13) & "您的系统版本可能不支持此功能或者您尚未安装这个功能.", vbCritical, "Error"
End Sub
Private Sub mnuColor_Click()
On Error GoTo ep
Dim lpSysPath As String
lpSysPath = Environ("Windir")
If Right(lpSysPath, 1) = "\" Then
lpSysPath = lpSysPath & "System32"
Else
lpSysPath = lpSysPath & "\System32"
End If
Shell lpSysPath & "\colorcpl.exe", vbNormalFocus
Exit Sub
ep:
MsgBox "发生系统错误:" & Chr(13) & Err.Description & Chr(13) & "您的系统版本可能不支持此功能或者您尚未安装这个功能.", vbCritical, "Error"
End Sub
Private Sub mnuDia_Click()
On Error GoTo ep
Dim lpSysPath As String
lpSysPath = Environ("Windir")
If Right(lpSysPath, 1) = "\" Then
lpSysPath = lpSysPath & "System32"
Else
lpSysPath = lpSysPath & "\System32"
End If
Shell lpSysPath & "\mstsc.exe", vbNormalFocus
Exit Sub
ep:
MsgBox "发生系统错误:" & Chr(13) & Err.Description & Chr(13) & "您的系统版本可能不支持此功能或者您尚未安装这个功能.", vbCritical, "Error"
End Sub
Private Sub mnuRc_Click()
On Error GoTo ep
Dim lpSysPath As String
lpSysPath = Environ("Windir")
If Right(lpSysPath, 1) = "\" Then
lpSysPath = lpSysPath & "System32"
Else
lpSysPath = lpSysPath & "\System32"
End If
Shell lpSysPath & "\msra.exe", vbNormalFocus
Exit Sub
ep:
MsgBox "发生系统错误:" & Chr(13) & Err.Description & Chr(13) & "您的系统版本可能不支持此功能或者您尚未安装这个功能.", vbCritical, "Error"
End Sub
Private Sub mnuMblCtr_Click()
On Error GoTo ep
Dim lpSysPath As String
lpSysPath = Environ("Windir")
If Right(lpSysPath, 1) = "\" Then
lpSysPath = lpSysPath & "System32"
Else
lpSysPath = lpSysPath & "\System32"
End If
Shell lpSysPath & "\DpiScaling.exe", vbNormalFocus
'ShellExecute Me.hwnd, vbNullString, lpSysPath & "\MdSched.exe", vbNullString, vbNullString, CLng(vbNormalFocus)
Exit Sub
ep:
MsgBox "发生系统错误:" & Chr(13) & Err.Description & Chr(13) & "您的系统版本可能不支持此功能或者您尚未安装这个功能.", vbCritical, "Error"
End Sub
Private Sub mnuOSK_Click()
On Error GoTo ep
Dim lpSysPath As String
lpSysPath = Environ("Windir")
If Right(lpSysPath, 1) = "\" Then
lpSysPath = lpSysPath & "System32"
Else
lpSysPath = lpSysPath & "\System32"
End If
lpSysPath = Trim(lpSysPath)
Shell lpSysPath & "\..\explorer.exe " & lpSysPath & "\osk.exe", vbNormalFocus
Exit Sub
ep:
MsgBox "发生系统错误:" & Chr(13) & Err.Description & Chr(13) & "您的系统版本可能不支持此功能或者您尚未安装这个功能.", vbCritical, "Error"
End Sub
Private Sub mnuMobile_Click()
On Error GoTo ep
Dim lpSysPath As String
lpSysPath = Environ("Windir")
If Right(lpSysPath, 1) = "\" Then
lpSysPath = lpSysPath & "System32"
Else
lpSysPath = lpSysPath & "\System32"
End If
lpSysPath = Trim(lpSysPath)
Shell lpSysPath & "\..\explorer.exe " & lpSysPath & "\mblctr.exe", vbNormalFocus
Exit Sub
ep:
MsgBox "发生系统错误:" & Chr(13) & Err.Description & Chr(13) & "您的系统版本可能不支持此功能或者您尚未安装这个功能.", vbCritical, "Error"
End Sub
Private Sub mnuRestore_Click()
On Error GoTo ep
Dim lpSysPath As String
lpSysPath = Environ("Windir")
If Right(lpSysPath, 1) = "\" Then
lpSysPath = lpSysPath & "System32"
Else
lpSysPath = lpSysPath & "\System32"
End If
lpSysPath = Trim(lpSysPath)
Shell lpSysPath & "\..\explorer.exe " & lpSysPath & "\rstrui.exe", vbNormalFocus
Exit Sub
ep:
MsgBox "发生系统错误:" & Chr(13) & Err.Description & Chr(13) & "您的系统版本可能不支持此功能或者您尚未安装这个功能.", vbCritical, "Error"
End Sub
Private Sub mnuSmartScreen_Click()
On Error GoTo ep
Dim lpSysPath As String
lpSysPath = Environ("Windir")
If Right(lpSysPath, 1) = "\" Then
lpSysPath = lpSysPath & "System32"
Else
lpSysPath = lpSysPath & "\System32"
End If
lpSysPath = Trim(lpSysPath)
Shell lpSysPath & "\..\explorer.exe " & lpSysPath & "\SmartScreenSettings.exe", vbNormalFocus
Exit Sub
ep:
MsgBox "发生系统错误:" & Chr(13) & Err.Description & Chr(13) & "您的系统版本可能不支持此功能或者您尚未安装这个功能.", vbCritical, "Error"
End Sub
Private Sub mnuSound_Cick()
On Error GoTo ep
Dim lpSysPath As String
lpSysPath = Environ("Windir")
If Right(lpSysPath, 1) = "\" Then
lpSysPath = lpSysPath & "System32"
Else
lpSysPath = lpSysPath & "\System32"
End If
lpSysPath = Trim(lpSysPath)
Shell lpSysPath & "\..\explorer.exe " & lpSysPath & "\SndVol.exe", vbNormalFocus
Exit Sub
ep:
MsgBox "发生系统错误:" & Chr(13) & Err.Description & Chr(13) & "您的系统版本可能不支持此功能或者您尚未安装这个功能.", vbCritical, "Error"
End Sub
Private Sub mnuUAC_Click()
On Error GoTo ep
Dim lpSysPath As String
lpSysPath = Environ("Windir")
If Right(lpSysPath, 1) = "\" Then
lpSysPath = lpSysPath & "System32"
Else
lpSysPath = lpSysPath & "\System32"
End If
lpSysPath = Trim(lpSysPath)
Shell lpSysPath & "\..\explorer.exe " & lpSysPath & "\UserAccountControlSettings.exe", vbNormalFocus
Exit Sub
ep:
MsgBox "发生系统错误:" & Chr(13) & Err.Description & Chr(13) & "您的系统版本可能不支持此功能或者您尚未安装这个功能.", vbCritical, "Error"
End Sub
Private Sub mnuPsr_Click()
On Error GoTo ep
Dim lpSysPath As String
lpSysPath = Environ("Windir")
If Right(lpSysPath, 1) = "\" Then
lpSysPath = lpSysPath & "System32"
Else
lpSysPath = lpSysPath & "\System32"
End If
lpSysPath = Trim(lpSysPath)
Shell lpSysPath & "\..\explorer.exe " & lpSysPath & "\psr.exe", vbNormalFocus
Exit Sub
ep:
MsgBox "发生系统错误:" & Chr(13) & Err.Description & Chr(13) & "您的系统版本可能不支持此功能或者您尚未安装这个功能.", vbCritical, "Error"
End Sub
Private Sub mnuRecdisk_Click()
On Error GoTo ep
Dim lpSysPath As String
lpSysPath = Environ("Windir")
If Right(lpSysPath, 1) = "\" Then
lpSysPath = lpSysPath & "System32"
Else
lpSysPath = lpSysPath & "\System32"
End If
lpSysPath = Trim(lpSysPath)
Shell lpSysPath & "\..\explorer.exe " & lpSysPath & "\RecoveryDrive.exe", vbNormalFocus
Exit Sub
ep:
MsgBox "发生系统错误:" & Chr(13) & Err.Description & Chr(13) & "您的系统版本可能不支持此功能或者您尚未安装这个功能.", vbCritical, "Error"
End Sub
Private Sub mnuMarrator_Click()
On Error GoTo ep
Dim lpSysPath As String
lpSysPath = Environ("Windir")
If Right(lpSysPath, 1) = "\" Then
lpSysPath = lpSysPath & "System32"
Else
lpSysPath = lpSysPath & "\System32"
End If
lpSysPath = Trim(lpSysPath)
Shell lpSysPath & "\..\explorer.exe " & lpSysPath & "\Narrator.exe", vbNormalFocus
Exit Sub
ep:
MsgBox "发生系统错误:" & Chr(13) & Err.Description & Chr(13) & "您的系统版本可能不支持此功能或者您尚未安装这个功能.", vbCritical, "Error"
End Sub
Private Sub mnuMaginfy_Click()
On Error GoTo ep
Dim lpSysPath As String
lpSysPath = Environ("Windir")
If Right(lpSysPath, 1) = "\" Then
lpSysPath = lpSysPath & "System32"
Else
lpSysPath = lpSysPath & "\System32"
End If
lpSysPath = Trim(lpSysPath)
Shell lpSysPath & "\..\explorer.exe " & lpSysPath & "\Magnify.exe", vbNormalFocus
Exit Sub
ep:
MsgBox "发生系统错误:" & Chr(13) & Err.Description & Chr(13) & "您的系统版本可能不支持此功能或者您尚未安装这个功能.", vbCritical, "Error"
End Sub
Private Sub mnuSched_Click()
On Error GoTo ep
Dim lpSysPath As String
lpSysPath = Environ("Windir")
If Right(lpSysPath, 1) = "\" Then
lpSysPath = lpSysPath & "System32"
Else
lpSysPath = lpSysPath & "\System32"
End If
lpSysPath = Trim(lpSysPath)
Shell lpSysPath & "\..\explorer.exe " & lpSysPath & "\MdSched.exe", vbNormalFocus
Exit Sub
ep:
MsgBox "发生系统错误:" & Chr(13) & Err.Description & Chr(13) & "您的系统版本可能不支持此功能或者您尚未安装这个功能.", vbCritical, "Error"
End Sub
Private Sub mnuRekeywiz_Click()
On Error GoTo ep
Dim lpSysPath As String
lpSysPath = Environ("Windir")
If Right(lpSysPath, 1) = "\" Then
lpSysPath = lpSysPath & "System32"
Else
lpSysPath = lpSysPath & "\System32"
End If
lpSysPath = Trim(lpSysPath)
Shell lpSysPath & "\..\explorer.exe " & lpSysPath & "\rekeywiz.exe", vbNormalFocus
Exit Sub
ep:
MsgBox "发生系统错误:" & Chr(13) & Err.Description & Chr(13) & "您的系统版本可能不支持此功能或者您尚未安装这个功能.", vbCritical, "Error"
End Sub
Private Sub mnuSound_Click()
On Error GoTo ep
Dim lpSysPath As String
lpSysPath = Environ("Windir")
If Right(lpSysPath, 1) = "\" Then
lpSysPath = lpSysPath & "System32"
Else
lpSysPath = lpSysPath & "\System32"
End If
lpSysPath = Trim(lpSysPath)
Shell lpSysPath & "\..\explorer.exe " & lpSysPath & "\SndVol.exe", vbNormalFocus
Exit Sub
ep:
MsgBox "发生系统错误:" & Chr(13) & Err.Description & Chr(13) & "您的系统版本可能不支持此功能或者您尚未安装这个功能.", vbCritical, "Error"
End Sub
Private Sub mnuLockKeys_Click()
On Error Resume Next
frmLock.Show
End Sub
Private Sub mnuKillGray_Click()
On Error Resume Next
frmKill.Show
End Sub
Private Sub Form_Click()
On Error Resume Next
Dim rtn As Long
Dim lRet As Long
'If Me.Tag = "ScreenProtection" Then
'lRet = SendMessage(Me.hwnd, WM_SYSCOMMAND, SC_SCREENSAVE, 0&)
'End If
If FormMain.Check1.Value = 1 Then
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
Else
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 25, LWA_ALPHA
End If
If Me.Tag = "None" Then
If FormMain.Check1.Value = 1 Then
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED Or WS_EX_TRANSPARENT
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
Else
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED Or WS_EX_TRANSPARENT
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 25, LWA_ALPHA
End If
Else
If FormMain.Check1.Value = 1 Then
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
Else
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 25, LWA_ALPHA
End If
End If
Dim lpCommand As String
lpCommand = Me.Tag
Dim ans As Integer
On Error Resume Next
Const TOKEN_ADJUST_PRIVILEGES = &H20
Const TOKEN_QUERY = &H8
Const SE_PRIVILEGE_ENABLED = &H2
Dim hdlProcessHandle As Long
Dim hdlTokenHandle As Long
Dim tmpLuid As LUID
Dim tkp As TOKEN_PRIVILEGES
Dim tkpNewButIgnored As TOKEN_PRIVILEGES
Dim lBufferNeeded As Long
Dim lpForce As Boolean
If FormMain.Check3.Value = 1 Then
lpForce = True
Else
lpForce = False
End If
Dim lpMsgProm As Boolean
If FormMain.Check2.Value = 1 Then
lpMsgProm = True
Else
lpMsgProm = False
End If
If lpMsgProm = True Then
Select Case lpCommand
Case "ScreenProtection"
lRet = SendMessage(Me.hwnd, WM_SYSCOMMAND, SC_SCREENSAVE, 0&)
Case "ScreenOff"
CloseScreenA True
Case "Shutdown"
If lpForce = True Then
ans = MsgBox("確定要關機嗎?請保存數據!", vbExclamation + vbYesNo + vbMsgBoxSetForeground, "Ask")
If ans = vbYes Then
hdlProcessHandle = GetCurrentProcess()
OpenProcessToken hdlProcessHandle, (TOKEN_ADJUST_PRIVILEGES Or _
TOKEN_QUERY), hdlTokenHandle
LookupPrivilegeValue "", "SeShutdownPrivilege", tmpLuid
tkp.PrivilegeCount = 1
tkp.TheLuid = tmpLuid
tkp.Attributes = SE_PRIVILEGE_ENABLED
AdjustTokenPrivileges hdlTokenHandle, False, _
tkp, Len(tkpNewButIgnored), tkpNewButIgnored, lBufferNeeded
AdjustToken
ExitWindowsEx EWX_SHUTDOWN Or EWX_FORCE, 0
Else
Exit Sub
End If
Else
ans = MsgBox("確定要關機嗎?請保存數據!", vbExclamation + vbYesNo + vbMsgBoxSetForeground, "Ask")
If ans = vbYes Then
hdlProcessHandle = GetCurrentProcess()
OpenProcessToken hdlProcessHandle, (TOKEN_ADJUST_PRIVILEGES Or _
TOKEN_QUERY), hdlTokenHandle
LookupPrivilegeValue "", "SeShutdownPrivilege", tmpLuid
tkp.PrivilegeCount = 1
tkp.TheLuid = tmpLuid
tkp.Attributes = SE_PRIVILEGE_ENABLED
AdjustTokenPrivileges hdlTokenHandle, False, _
tkp, Len(tkpNewButIgnored), tkpNewButIgnored, lBufferNeeded
AdjustToken
ExitWindowsEx EWX_SHUTDOWN, 0
Else
Exit Sub
End If
End If
Case "Reboot"
If lpForce = True Then
ans = MsgBox("確定要重啟嗎?請保存數據!", vbExclamation + vbYesNo + vbMsgBoxSetForeground, "Ask")
If ans = vbYes Then
hdlProcessHandle = GetCurrentProcess()
OpenProcessToken hdlProcessHandle, (TOKEN_ADJUST_PRIVILEGES Or _
TOKEN_QUERY), hdlTokenHandle
LookupPrivilegeValue "", "SeShutdownPrivilege", tmpLuid
tkp.PrivilegeCount = 1
tkp.TheLuid = tmpLuid
tkp.Attributes = SE_PRIVILEGE_ENABLED
AdjustTokenPrivileges hdlTokenHandle, False, _
tkp, Len(tkpNewButIgnored), tkpNewButIgnored, lBufferNeeded
AdjustToken
ExitWindowsEx EWX_REBOOT Or EWX_FORCE, 0
Else
Exit Sub
End If
Else
ans = MsgBox("確定要重啟嗎?請保存數據!", vbExclamation + vbYesNo + vbMsgBoxSetForeground, "Ask")
If ans = vbYes Then
hdlProcessHandle = GetCurrentProcess()
OpenProcessToken hdlProcessHandle, (TOKEN_ADJUST_PRIVILEGES Or _
TOKEN_QUERY), hdlTokenHandle
LookupPrivilegeValue "", "SeShutdownPrivilege", tmpLuid
tkp.PrivilegeCount = 1
tkp.TheLuid = tmpLuid
tkp.Attributes = SE_PRIVILEGE_ENABLED
AdjustTokenPrivileges hdlTokenHandle, False, _
tkp, Len(tkpNewButIgnored), tkpNewButIgnored, lBufferNeeded
AdjustToken
ExitWindowsEx EWX_REBOOT, 0
Else
Exit Sub
End If
End If
Case "Logoff"
If lpForce = True Then
ans = MsgBox("確定要註銷嗎?請保存數據!", vbExclamation + vbYesNo + vbMsgBoxSetForeground, "Ask")
If ans = vbYes Then
hdlProcessHandle = GetCurrentProcess()
OpenProcessToken hdlProcessHandle, (TOKEN_ADJUST_PRIVILEGES Or _
TOKEN_QUERY), hdlTokenHandle
LookupPrivilegeValue "", "SeShutdownPrivilege", tmpLuid
tkp.PrivilegeCount = 1
tkp.TheLuid = tmpLuid
tkp.Attributes = SE_PRIVILEGE_ENABLED
AdjustTokenPrivileges hdlTokenHandle, False, _
tkp, Len(tkpNewButIgnored), tkpNewButIgnored, lBufferNeeded
AdjustToken
ExitWindowsEx EWX_LOGOFF Or EWX_FORCE, 0
Else
Exit Sub
End If
Else
ans = MsgBox("確定要註銷嗎?請保存數據!", vbExclamation + vbYesNo + vbMsgBoxSetForeground, "Ask")
If ans = vbYes Then
hdlProcessHandle = GetCurrentProcess()
OpenProcessToken hdlProcessHandle, (TOKEN_ADJUST_PRIVILEGES Or _
TOKEN_QUERY), hdlTokenHandle
LookupPrivilegeValue "", "SeShutdownPrivilege", tmpLuid
tkp.PrivilegeCount = 1
tkp.TheLuid = tmpLuid
tkp.Attributes = SE_PRIVILEGE_ENABLED
AdjustTokenPrivileges hdlTokenHandle, False, _
tkp, Len(tkpNewButIgnored), tkpNewButIgnored, lBufferNeeded
AdjustToken
ExitWindowsEx EWX_LOGOFF, 0
Else
Exit Sub
End If
End If
Case "Menu"
PopupMenu Me.UserMenu
Case Else
Exit Sub
End Select
ElseIf lpMsgProm = False Then
Select Case lpCommand
Case "ScreenOff"
CloseScreenA True
Case "ScreenProtection"
lRet = SendMessage(Me.hwnd, WM_SYSCOMMAND, SC_SCREENSAVE, 0&)
Case "Shutdown"
If lpForce = True Then
ans = vbYes
If ans = vbYes Then
hdlProcessHandle = GetCurrentProcess()
OpenProcessToken hdlProcessHandle, (TOKEN_ADJUST_PRIVILEGES Or _
TOKEN_QUERY), hdlTokenHandle
LookupPrivilegeValue "", "SeShutdownPrivilege", tmpLuid
tkp.PrivilegeCount = 1
tkp.TheLuid = tmpLuid
tkp.Attributes = SE_PRIVILEGE_ENABLED
AdjustTokenPrivileges hdlTokenHandle, False, _
tkp, Len(tkpNewButIgnored), tkpNewButIgnored, lBufferNeeded
AdjustToken
ExitWindowsEx EWX_SHUTDOWN Or EWX_FORCE, 0
Else
Exit Sub
End If
Else
ans = vbYes
If ans = vbYes Then
hdlProcessHandle = GetCurrentProcess()
OpenProcessToken hdlProcessHandle, (TOKEN_ADJUST_PRIVILEGES Or _
TOKEN_QUERY), hdlTokenHandle
LookupPrivilegeValue "", "SeShutdownPrivilege", tmpLuid
tkp.PrivilegeCount = 1
tkp.TheLuid = tmpLuid
tkp.Attributes = SE_PRIVILEGE_ENABLED
AdjustTokenPrivileges hdlTokenHandle, False, _
tkp, Len(tkpNewButIgnored), tkpNewButIgnored, lBufferNeeded
AdjustToken
ExitWindowsEx EWX_SHUTDOWN, 0
Else
Exit Sub
End If
End If
Case "Reboot"
If lpForce = True Then
ans = vbYes
If ans = vbYes Then
hdlProcessHandle = GetCurrentProcess()
OpenProcessToken hdlProcessHandle, (TOKEN_ADJUST_PRIVILEGES Or _
TOKEN_QUERY), hdlTokenHandle
LookupPrivilegeValue "", "SeShutdownPrivilege", tmpLuid
tkp.PrivilegeCount = 1
tkp.TheLuid = tmpLuid
tkp.Attributes = SE_PRIVILEGE_ENABLED
AdjustTokenPrivileges hdlTokenHandle, False, _
tkp, Len(tkpNewButIgnored), tkpNewButIgnored, lBufferNeeded
AdjustToken
ExitWindowsEx EWX_REBOOT Or EWX_FORCE, 0
Else
Exit Sub
End If
Else
ans = vbYes
If ans = vbYes Then
hdlProcessHandle = GetCurrentProcess()
OpenProcessToken hdlProcessHandle, (TOKEN_ADJUST_PRIVILEGES Or _
TOKEN_QUERY), hdlTokenHandle
LookupPrivilegeValue "", "SeShutdownPrivilege", tmpLuid
tkp.PrivilegeCount = 1
tkp.TheLuid = tmpLuid
tkp.Attributes = SE_PRIVILEGE_ENABLED
AdjustTokenPrivileges hdlTokenHandle, False, _
tkp, Len(tkpNewButIgnored), tkpNewButIgnored, lBufferNeeded
AdjustToken
ExitWindowsEx EWX_REBOOT, 0
Else
Exit Sub
End If
End If
Case "Logoff"
If lpForce = True Then
ans = vbYes
If ans = vbYes Then
hdlProcessHandle = GetCurrentProcess()
OpenProcessToken hdlProcessHandle, (TOKEN_ADJUST_PRIVILEGES Or _
TOKEN_QUERY), hdlTokenHandle
LookupPrivilegeValue "", "SeShutdownPrivilege", tmpLuid
tkp.PrivilegeCount = 1
tkp.TheLuid = tmpLuid
tkp.Attributes = SE_PRIVILEGE_ENABLED
AdjustTokenPrivileges hdlTokenHandle, False, _
tkp, Len(tkpNewButIgnored), tkpNewButIgnored, lBufferNeeded
AdjustToken
ExitWindowsEx EWX_LOGOFF Or EWX_FORCE, 0
Else
Exit Sub
End If
Else
ans = vbYes
If ans = vbYes Then
hdlProcessHandle = GetCurrentProcess()
OpenProcessToken hdlProcessHandle, (TOKEN_ADJUST_PRIVILEGES Or _
TOKEN_QUERY), hdlTokenHandle
LookupPrivilegeValue "", "SeShutdownPrivilege", tmpLuid
tkp.PrivilegeCount = 1
tkp.TheLuid = tmpLuid
tkp.Attributes = SE_PRIVILEGE_ENABLED
AdjustTokenPrivileges hdlTokenHandle, False, _
tkp, Len(tkpNewButIgnored), tkpNewButIgnored, lBufferNeeded
AdjustToken
ExitWindowsEx EWX_LOGOFF, 0
Else
Exit Sub
End If
End If
Case "Menu"
PopupMenu Me.UserMenu
Case Else
Exit Sub
End Select
Else
Select Case lpCommand
Case "Shutdown"
If lpForce = True Then
ans = MsgBox("確定要關機嗎?請保存數據!", vbExclamation + vbYesNo + vbMsgBoxSetForeground, "Ask")
If ans = vbYes Then
hdlProcessHandle = GetCurrentProcess()
OpenProcessToken hdlProcessHandle, (TOKEN_ADJUST_PRIVILEGES Or _
TOKEN_QUERY), hdlTokenHandle
LookupPrivilegeValue "", "SeShutdownPrivilege", tmpLuid
tkp.PrivilegeCount = 1
tkp.TheLuid = tmpLuid
tkp.Attributes = SE_PRIVILEGE_ENABLED
AdjustTokenPrivileges hdlTokenHandle, False, _
tkp, Len(tkpNewButIgnored), tkpNewButIgnored, lBufferNeeded
AdjustToken
ExitWindowsEx EWX_SHUTDOWN Or EWX_FORCE, 0
Else
Exit Sub
End If
Else
ans = MsgBox("確定要關機嗎?請保存數據!", vbExclamation + vbYesNo + vbMsgBoxSetForeground, "Ask")
If ans = vbYes Then
hdlProcessHandle = GetCurrentProcess()
OpenProcessToken hdlProcessHandle, (TOKEN_ADJUST_PRIVILEGES Or _
TOKEN_QUERY), hdlTokenHandle
LookupPrivilegeValue "", "SeShutdownPrivilege", tmpLuid
tkp.PrivilegeCount = 1
tkp.TheLuid = tmpLuid
tkp.Attributes = SE_PRIVILEGE_ENABLED
AdjustTokenPrivileges hdlTokenHandle, False, _
tkp, Len(tkpNewButIgnored), tkpNewButIgnored, lBufferNeeded
AdjustToken
ExitWindowsEx EWX_SHUTDOWN, 0
Else
Exit Sub
End If
End If
Case "Reboot"
If lpForce = True Then
ans = MsgBox("確定要重啟嗎?請保存數據!", vbExclamation + vbYesNo + vbMsgBoxSetForeground, "Ask")
If ans = vbYes Then
hdlProcessHandle = GetCurrentProcess()
OpenProcessToken hdlProcessHandle, (TOKEN_ADJUST_PRIVILEGES Or _
TOKEN_QUERY), hdlTokenHandle
LookupPrivilegeValue "", "SeShutdownPrivilege", tmpLuid
tkp.PrivilegeCount = 1
tkp.TheLuid = tmpLuid
tkp.Attributes = SE_PRIVILEGE_ENABLED
AdjustTokenPrivileges hdlTokenHandle, False, _
tkp, Len(tkpNewButIgnored), tkpNewButIgnored, lBufferNeeded
AdjustToken
ExitWindowsEx EWX_REBOOT Or EWX_FORCE, 0
Else
Exit Sub
End If
Else
ans = MsgBox("確定要重啟嗎?請保存數據!", vbExclamation + vbYesNo + vbMsgBoxSetForeground, "Ask")
If ans = vbYes Then
hdlProcessHandle = GetCurrentProcess()
OpenProcessToken hdlProcessHandle, (TOKEN_ADJUST_PRIVILEGES Or _
TOKEN_QUERY), hdlTokenHandle
LookupPrivilegeValue "", "SeShutdownPrivilege", tmpLuid
tkp.PrivilegeCount = 1
tkp.TheLuid = tmpLuid
tkp.Attributes = SE_PRIVILEGE_ENABLED
AdjustTokenPrivileges hdlTokenHandle, False, _
tkp, Len(tkpNewButIgnored), tkpNewButIgnored, lBufferNeeded
AdjustToken
ExitWindowsEx EWX_REBOOT, 0
Else
Exit Sub
End If
End If
Case "Logoff"
If lpForce = True Then
ans = MsgBox("確定要註銷嗎?請保存數據!", vbExclamation + vbYesNo + vbMsgBoxSetForeground, "Ask")
If ans = vbYes Then
hdlProcessHandle = GetCurrentProcess()
OpenProcessToken hdlProcessHandle, (TOKEN_ADJUST_PRIVILEGES Or _
TOKEN_QUERY), hdlTokenHandle
LookupPrivilegeValue "", "SeShutdownPrivilege", tmpLuid
tkp.PrivilegeCount = 1
tkp.TheLuid = tmpLuid
tkp.Attributes = SE_PRIVILEGE_ENABLED
AdjustTokenPrivileges hdlTokenHandle, False, _
tkp, Len(tkpNewButIgnored), tkpNewButIgnored, lBufferNeeded
AdjustToken
ExitWindowsEx EWX_LOGOFF Or EWX_FORCE, 0
Else
Exit Sub
End If
Else
ans = MsgBox("確定要註銷嗎?請保存數據!", vbExclamation + vbYesNo + vbMsgBoxSetForeground, "Ask")
If ans = vbYes Then
hdlProcessHandle = GetCurrentProcess()
OpenProcessToken hdlProcessHandle, (TOKEN_ADJUST_PRIVILEGES Or _
TOKEN_QUERY), hdlTokenHandle
LookupPrivilegeValue "", "SeShutdownPrivilege", tmpLuid
tkp.PrivilegeCount = 1
tkp.TheLuid = tmpLuid
tkp.Attributes = SE_PRIVILEGE_ENABLED
AdjustTokenPrivileges hdlTokenHandle, False, _
tkp, Len(tkpNewButIgnored), tkpNewButIgnored, lBufferNeeded
AdjustToken
ExitWindowsEx EWX_LOGOFF, 0
Else
Exit Sub
End If
End If
Case "Menu"
PopupMenu Me.UserMenu
Case Else
Exit Sub
End Select
End If
End Sub
Private Sub Form_DblClick()
On Error Resume Next
Dim rtn As Long
Dim lRet As Long
'If Me.Tag = "ScreenProtection" Then
'lRet = SendMessage(Me.hwnd, WM_SYSCOMMAND, SC_SCREENSAVE, 0&)
'End If
If FormMain.Check1.Value = 1 Then
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
Else
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 25, LWA_ALPHA
End If
If Me.Tag = "None" Then
If FormMain.Check1.Value = 1 Then
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED Or WS_EX_TRANSPARENT
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
Else
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED Or WS_EX_TRANSPARENT
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 25, LWA_ALPHA
End If
Else
If FormMain.Check1.Value = 1 Then
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
Else
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 25, LWA_ALPHA
End If
End If
Dim lpCommand As String
lpCommand = Me.Tag
Dim ans As Integer
On Error Resume Next
Const TOKEN_ADJUST_PRIVILEGES = &H20
Const TOKEN_QUERY = &H8
Const SE_PRIVILEGE_ENABLED = &H2
Dim hdlProcessHandle As Long
Dim hdlTokenHandle As Long
Dim tmpLuid As LUID
Dim tkp As TOKEN_PRIVILEGES
Dim tkpNewButIgnored As TOKEN_PRIVILEGES
Dim lBufferNeeded As Long
Dim lpForce As Boolean
If FormMain.Check3.Value = 1 Then
lpForce = True
Else
lpForce = False
End If
Dim lpMsgProm As Boolean
If FormMain.Check2.Value = 1 Then
lpMsgProm = True
Else
lpMsgProm = False
End If
If lpMsgProm = True Then
Select Case lpCommand
Case "ScreenProtection"
lRet = SendMessage(Me.hwnd, WM_SYSCOMMAND, SC_SCREENSAVE, 0&)
Case "ScreenOff"
CloseScreenA True
Case "Shutdown"
If lpForce = True Then
ans = MsgBox("確定要關機嗎?請保存數據!", vbExclamation + vbYesNo + vbMsgBoxSetForeground, "Ask")
If ans = vbYes Then
hdlProcessHandle = GetCurrentProcess()
OpenProcessToken hdlProcessHandle, (TOKEN_ADJUST_PRIVILEGES Or _
TOKEN_QUERY), hdlTokenHandle
LookupPrivilegeValue "", "SeShutdownPrivilege", tmpLuid
tkp.PrivilegeCount = 1
tkp.TheLuid = tmpLuid
tkp.Attributes = SE_PRIVILEGE_ENABLED
AdjustTokenPrivileges hdlTokenHandle, False, _
tkp, Len(tkpNewButIgnored), tkpNewButIgnored, lBufferNeeded
AdjustToken
ExitWindowsEx EWX_SHUTDOWN Or EWX_FORCE, 0
Else
Exit Sub
End If
Else
ans = MsgBox("確定要關機嗎?請保存數據!", vbExclamation + vbYesNo + vbMsgBoxSetForeground, "Ask")
If ans = vbYes Then
hdlProcessHandle = GetCurrentProcess()
OpenProcessToken hdlProcessHandle, (TOKEN_ADJUST_PRIVILEGES Or _
TOKEN_QUERY), hdlTokenHandle
LookupPrivilegeValue "", "SeShutdownPrivilege", tmpLuid
tkp.PrivilegeCount = 1
tkp.TheLuid = tmpLuid
tkp.Attributes = SE_PRIVILEGE_ENABLED
AdjustTokenPrivileges hdlTokenHandle, False, _
tkp, Len(tkpNewButIgnored), tkpNewButIgnored, lBufferNeeded
AdjustToken
ExitWindowsEx EWX_SHUTDOWN, 0
Else
Exit Sub
End If
End If
Case "Reboot"
If lpForce = True Then
ans = MsgBox("確定要重啟嗎?請保存數據!", vbExclamation + vbYesNo + vbMsgBoxSetForeground, "Ask")
If ans = vbYes Then
hdlProcessHandle = GetCurrentProcess()
OpenProcessToken hdlProcessHandle, (TOKEN_ADJUST_PRIVILEGES Or _
TOKEN_QUERY), hdlTokenHandle
LookupPrivilegeValue "", "SeShutdownPrivilege", tmpLuid
tkp.PrivilegeCount = 1
tkp.TheLuid = tmpLuid
tkp.Attributes = SE_PRIVILEGE_ENABLED
AdjustTokenPrivileges hdlTokenHandle, False, _
tkp, Len(tkpNewButIgnored), tkpNewButIgnored, lBufferNeeded
AdjustToken
ExitWindowsEx EWX_REBOOT Or EWX_FORCE, 0
Else
Exit Sub
End If
Else
ans = MsgBox("確定要重啟嗎?請保存數據!", vbExclamation + vbYesNo + vbMsgBoxSetForeground, "Ask")
If ans = vbYes Then
hdlProcessHandle = GetCurrentProcess()
OpenProcessToken hdlProcessHandle, (TOKEN_ADJUST_PRIVILEGES Or _
TOKEN_QUERY), hdlTokenHandle
LookupPrivilegeValue "", "SeShutdownPrivilege", tmpLuid
tkp.PrivilegeCount = 1
tkp.TheLuid = tmpLuid
tkp.Attributes = SE_PRIVILEGE_ENABLED
AdjustTokenPrivileges hdlTokenHandle, False, _
tkp, Len(tkpNewButIgnored), tkpNewButIgnored, lBufferNeeded
AdjustToken
ExitWindowsEx EWX_REBOOT, 0
Else
Exit Sub
End If
End If
Case "Logoff"
If lpForce = True Then
ans = MsgBox("確定要註銷嗎?請保存數據!", vbExclamation + vbYesNo + vbMsgBoxSetForeground, "Ask")
If ans = vbYes Then
hdlProcessHandle = GetCurrentProcess()
OpenProcessToken hdlProcessHandle, (TOKEN_ADJUST_PRIVILEGES Or _
TOKEN_QUERY), hdlTokenHandle
LookupPrivilegeValue "", "SeShutdownPrivilege", tmpLuid
tkp.PrivilegeCount = 1
tkp.TheLuid = tmpLuid
tkp.Attributes = SE_PRIVILEGE_ENABLED
AdjustTokenPrivileges hdlTokenHandle, False, _
tkp, Len(tkpNewButIgnored), tkpNewButIgnored, lBufferNeeded
AdjustToken
ExitWindowsEx EWX_LOGOFF Or EWX_FORCE, 0
Else
Exit Sub
End If
Else
ans = MsgBox("確定要註銷嗎?請保存數據!", vbExclamation + vbYesNo + vbMsgBoxSetForeground, "Ask")
If ans = vbYes Then
hdlProcessHandle = GetCurrentProcess()
OpenProcessToken hdlProcessHandle, (TOKEN_ADJUST_PRIVILEGES Or _
TOKEN_QUERY), hdlTokenHandle
LookupPrivilegeValue "", "SeShutdownPrivilege", tmpLuid
tkp.PrivilegeCount = 1
tkp.TheLuid = tmpLuid
tkp.Attributes = SE_PRIVILEGE_ENABLED
AdjustTokenPrivileges hdlTokenHandle, False, _
tkp, Len(tkpNewButIgnored), tkpNewButIgnored, lBufferNeeded
AdjustToken
ExitWindowsEx EWX_LOGOFF, 0
Else
Exit Sub
End If
End If
Case "Menu"
PopupMenu Me.UserMenu
Case Else
Exit Sub
End Select
ElseIf lpMsgProm = False Then
Select Case lpCommand
Case "ScreenOff"
CloseScreenA True
Case "ScreenProtection"
lRet = SendMessage(Me.hwnd, WM_SYSCOMMAND, SC_SCREENSAVE, 0&)
Case "Shutdown"
If lpForce = True Then
ans = vbYes
If ans = vbYes Then
hdlProcessHandle = GetCurrentProcess()
OpenProcessToken hdlProcessHandle, (TOKEN_ADJUST_PRIVILEGES Or _
TOKEN_QUERY), hdlTokenHandle
LookupPrivilegeValue "", "SeShutdownPrivilege", tmpLuid
tkp.PrivilegeCount = 1
tkp.TheLuid = tmpLuid
tkp.Attributes = SE_PRIVILEGE_ENABLED
AdjustTokenPrivileges hdlTokenHandle, False, _
tkp, Len(tkpNewButIgnored), tkpNewButIgnored, lBufferNeeded
AdjustToken
ExitWindowsEx EWX_SHUTDOWN Or EWX_FORCE, 0
Else
Exit Sub
End If
Else
ans = vbYes
If ans = vbYes Then
hdlProcessHandle = GetCurrentProcess()
OpenProcessToken hdlProcessHandle, (TOKEN_ADJUST_PRIVILEGES Or _
TOKEN_QUERY), hdlTokenHandle
LookupPrivilegeValue "", "SeShutdownPrivilege", tmpLuid
tkp.PrivilegeCount = 1
tkp.TheLuid = tmpLuid
tkp.Attributes = SE_PRIVILEGE_ENABLED
AdjustTokenPrivileges hdlTokenHandle, False, _
tkp, Len(tkpNewButIgnored), tkpNewButIgnored, lBufferNeeded
AdjustToken
ExitWindowsEx EWX_SHUTDOWN, 0
Else
Exit Sub
End If
End If
Case "Reboot"
If lpForce = True Then
ans = vbYes
If ans = vbYes Then
hdlProcessHandle = GetCurrentProcess()
OpenProcessToken hdlProcessHandle, (TOKEN_ADJUST_PRIVILEGES Or _
TOKEN_QUERY), hdlTokenHandle
LookupPrivilegeValue "", "SeShutdownPrivilege", tmpLuid
tkp.PrivilegeCount = 1
tkp.TheLuid = tmpLuid
tkp.Attributes = SE_PRIVILEGE_ENABLED
AdjustTokenPrivileges hdlTokenHandle, False, _
tkp, Len(tkpNewButIgnored), tkpNewButIgnored, lBufferNeeded
AdjustToken
ExitWindowsEx EWX_REBOOT Or EWX_FORCE, 0
Else
Exit Sub
End If
Else
ans = vbYes
If ans = vbYes Then
hdlProcessHandle = GetCurrentProcess()
OpenProcessToken hdlProcessHandle, (TOKEN_ADJUST_PRIVILEGES Or _
TOKEN_QUERY), hdlTokenHandle
LookupPrivilegeValue "", "SeShutdownPrivilege", tmpLuid
tkp.PrivilegeCount = 1
tkp.TheLuid = tmpLuid
tkp.Attributes = SE_PRIVILEGE_ENABLED
AdjustTokenPrivileges hdlTokenHandle, False, _
tkp, Len(tkpNewButIgnored), tkpNewButIgnored, lBufferNeeded
AdjustToken
ExitWindowsEx EWX_REBOOT, 0
Else
Exit Sub
End If
End If
Case "Logoff"
If lpForce = True Then
ans = vbYes
If ans = vbYes Then
hdlProcessHandle = GetCurrentProcess()
OpenProcessToken hdlProcessHandle, (TOKEN_ADJUST_PRIVILEGES Or _
TOKEN_QUERY), hdlTokenHandle
LookupPrivilegeValue "", "SeShutdownPrivilege", tmpLuid
tkp.PrivilegeCount = 1
tkp.TheLuid = tmpLuid
tkp.Attributes = SE_PRIVILEGE_ENABLED
AdjustTokenPrivileges hdlTokenHandle, False, _
tkp, Len(tkpNewButIgnored), tkpNewButIgnored, lBufferNeeded
AdjustToken
ExitWindowsEx EWX_LOGOFF Or EWX_FORCE, 0
Else
Exit Sub
End If
Else
ans = vbYes
If ans = vbYes Then
hdlProcessHandle = GetCurrentProcess()
OpenProcessToken hdlProcessHandle, (TOKEN_ADJUST_PRIVILEGES Or _
TOKEN_QUERY), hdlTokenHandle
LookupPrivilegeValue "", "SeShutdownPrivilege", tmpLuid
tkp.PrivilegeCount = 1
tkp.TheLuid = tmpLuid
tkp.Attributes = SE_PRIVILEGE_ENABLED
AdjustTokenPrivileges hdlTokenHandle, False, _
tkp, Len(tkpNewButIgnored), tkpNewButIgnored, lBufferNeeded
AdjustToken
ExitWindowsEx EWX_LOGOFF, 0
Else
Exit Sub
End If
End If
Case "Menu"
PopupMenu Me.UserMenu
Case Else
Exit Sub
End Select
Else
Select Case lpCommand
Case "Shutdown"
If lpForce = True Then
ans = MsgBox("確定要關機嗎?請保存數據!", vbExclamation + vbYesNo + vbMsgBoxSetForeground, "Ask")
If ans = vbYes Then
hdlProcessHandle = GetCurrentProcess()
OpenProcessToken hdlProcessHandle, (TOKEN_ADJUST_PRIVILEGES Or _
TOKEN_QUERY), hdlTokenHandle
LookupPrivilegeValue "", "SeShutdownPrivilege", tmpLuid
tkp.PrivilegeCount = 1
tkp.TheLuid = tmpLuid
tkp.Attributes = SE_PRIVILEGE_ENABLED
AdjustTokenPrivileges hdlTokenHandle, False, _
tkp, Len(tkpNewButIgnored), tkpNewButIgnored, lBufferNeeded
AdjustToken
ExitWindowsEx EWX_SHUTDOWN Or EWX_FORCE, 0
Else
Exit Sub
End If
Else
ans = MsgBox("確定要關機嗎?請保存數據!", vbExclamation + vbYesNo + vbMsgBoxSetForeground, "Ask")
If ans = vbYes Then
hdlProcessHandle = GetCurrentProcess()
OpenProcessToken hdlProcessHandle, (TOKEN_ADJUST_PRIVILEGES Or _
TOKEN_QUERY), hdlTokenHandle
LookupPrivilegeValue "", "SeShutdownPrivilege", tmpLuid
tkp.PrivilegeCount = 1
tkp.TheLuid = tmpLuid
tkp.Attributes = SE_PRIVILEGE_ENABLED
AdjustTokenPrivileges hdlTokenHandle, False, _
tkp, Len(tkpNewButIgnored), tkpNewButIgnored, lBufferNeeded
AdjustToken
ExitWindowsEx EWX_SHUTDOWN, 0
Else
Exit Sub
End If
End If
Case "Reboot"
If lpForce = True Then
ans = MsgBox("確定要重啟嗎?請保存數據!", vbExclamation + vbYesNo + vbMsgBoxSetForeground, "Ask")
If ans = vbYes Then
hdlProcessHandle = GetCurrentProcess()
OpenProcessToken hdlProcessHandle, (TOKEN_ADJUST_PRIVILEGES Or _
TOKEN_QUERY), hdlTokenHandle
LookupPrivilegeValue "", "SeShutdownPrivilege", tmpLuid
tkp.PrivilegeCount = 1
tkp.TheLuid = tmpLuid
tkp.Attributes = SE_PRIVILEGE_ENABLED
AdjustTokenPrivileges hdlTokenHandle, False, _
tkp, Len(tkpNewButIgnored), tkpNewButIgnored, lBufferNeeded
AdjustToken
ExitWindowsEx EWX_REBOOT Or EWX_FORCE, 0
Else
Exit Sub
End If
Else
ans = MsgBox("確定要重啟嗎?請保存數據!", vbExclamation + vbYesNo + vbMsgBoxSetForeground, "Ask")
If ans = vbYes Then
hdlProcessHandle = GetCurrentProcess()
OpenProcessToken hdlProcessHandle, (TOKEN_ADJUST_PRIVILEGES Or _
TOKEN_QUERY), hdlTokenHandle
LookupPrivilegeValue "", "SeShutdownPrivilege", tmpLuid
tkp.PrivilegeCount = 1
tkp.TheLuid = tmpLuid
tkp.Attributes = SE_PRIVILEGE_ENABLED
AdjustTokenPrivileges hdlTokenHandle, False, _
tkp, Len(tkpNewButIgnored), tkpNewButIgnored, lBufferNeeded
AdjustToken
ExitWindowsEx EWX_REBOOT, 0
Else
Exit Sub
End If
End If
Case "Logoff"
If lpForce = True Then
ans = MsgBox("確定要註銷嗎?請保存數據!", vbExclamation + vbYesNo + vbMsgBoxSetForeground, "Ask")
If ans = vbYes Then
hdlProcessHandle = GetCurrentProcess()
OpenProcessToken hdlProcessHandle, (TOKEN_ADJUST_PRIVILEGES Or _
TOKEN_QUERY), hdlTokenHandle
LookupPrivilegeValue "", "SeShutdownPrivilege", tmpLuid
tkp.PrivilegeCount = 1
tkp.TheLuid = tmpLuid
tkp.Attributes = SE_PRIVILEGE_ENABLED
AdjustTokenPrivileges hdlTokenHandle, False, _
tkp, Len(tkpNewButIgnored), tkpNewButIgnored, lBufferNeeded
AdjustToken
ExitWindowsEx EWX_LOGOFF Or EWX_FORCE, 0
Else
Exit Sub
End If
Else
ans = MsgBox("確定要註銷嗎?請保存數據!", vbExclamation + vbYesNo + vbMsgBoxSetForeground, "Ask")
If ans = vbYes Then
hdlProcessHandle = GetCurrentProcess()
OpenProcessToken hdlProcessHandle, (TOKEN_ADJUST_PRIVILEGES Or _
TOKEN_QUERY), hdlTokenHandle
LookupPrivilegeValue "", "SeShutdownPrivilege", tmpLuid
tkp.PrivilegeCount = 1
tkp.TheLuid = tmpLuid
tkp.Attributes = SE_PRIVILEGE_ENABLED
AdjustTokenPrivileges hdlTokenHandle, False, _
tkp, Len(tkpNewButIgnored), tkpNewButIgnored, lBufferNeeded
AdjustToken
ExitWindowsEx EWX_LOGOFF, 0
Else
Exit Sub
End If
End If
Case "Menu"
PopupMenu Me.UserMenu
Case Else
Exit Sub
End Select
End If
End Sub
Private Sub mnuScreenOff_Click()
On Error Resume Next
CloseScreenA True
End Sub
Private Sub mnuSCR_Click()
On Error Resume Next
Dim lRet As Long
lRet = SendMessage(Form1.hwnd, WM_SYSCOMMAND, SC_SCREENSAVE, 0&)
End Sub
Private Sub mnuLargeFile_Click()
On Error Resume Next
frmLargefile.Show
End Sub
Private Sub mnuTaskApp_Click()
On Error Resume Next
frmTaskMain.Show
End Sub
Private Sub mnuVClick_Click()
On Error Resume Next
FormVC1.Show
End Sub
Private Sub mnuMClip_Click()
On Error Resume Next
frmMCM.Show
End Sub
