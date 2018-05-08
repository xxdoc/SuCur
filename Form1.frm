VERSION 5.00
Begin VB.Form FormMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Super Cursor - PC-DOS Workshop"
   ClientHeight    =   7770
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5535
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7770
   ScaleWidth      =   5535
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command4 
      Height          =   390
      Left            =   2880
      MaskColor       =   &H00C8D0D4&
      Picture         =   "Form1.frx":068A
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   7320
      UseMaskColor    =   -1  'True
      Width           =   255
   End
   Begin VB.CommandButton Command3 
      Caption         =   "預覽超級菜單(&V)"
      Height          =   390
      Left            =   360
      TabIndex        =   24
      Top             =   7320
      Width           =   2535
   End
   Begin VB.Timer Timer1 
      Interval        =   245
      Left            =   2160
      Top             =   3255
   End
   Begin 工程1.cSysTray cSysTray1 
      Left            =   2505
      Top             =   3240
      _extentx        =   900
      _extenty        =   900
      intray          =   0
      trayicon        =   "Form1.frx":0858
      traytip         =   "Super Cursor - 雙擊還原主窗口,右擊顯示菜單"
   End
   Begin VB.CommandButton Command2 
      Caption         =   "最小化(&M)"
      Height          =   375
      Left            =   3240
      TabIndex        =   20
      Top             =   7320
      Width           =   2250
   End
   Begin VB.Frame Frame1 
      Caption         =   "Super Cursor選項"
      Height          =   5835
      Left            =   375
      TabIndex        =   3
      Top             =   1410
      Width           =   5130
      Begin VB.Frame Frame3 
         Caption         =   "關機/重啟/註銷選項"
         Height          =   930
         Left            =   135
         TabIndex        =   21
         Top             =   4755
         Width           =   4875
         Begin VB.CheckBox Check2 
            Caption         =   "關機/重啟/註銷前提示(&P)"
            Height          =   270
            Left            =   120
            TabIndex        =   23
            Top             =   240
            Value           =   1  'Checked
            Width           =   2460
         End
         Begin VB.CheckBox Check3 
            Caption         =   "強制結束未響應進程(&F)"
            Height          =   270
            Left            =   120
            TabIndex        =   22
            Top             =   555
            Width           =   2250
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "屏幕"
         Height          =   4500
         Left            =   135
         TabIndex        =   4
         Top             =   210
         Width           =   4890
         Begin VB.PictureBox Picture4 
            BackColor       =   &H00FF0000&
            Height          =   180
            Left            =   3750
            ScaleHeight     =   120
            ScaleWidth      =   135
            TabIndex        =   29
            Top             =   2010
            Width           =   195
         End
         Begin VB.PictureBox Picture3 
            BackColor       =   &H00FF0000&
            Height          =   180
            Left            =   975
            ScaleHeight     =   120
            ScaleWidth      =   135
            TabIndex        =   28
            Top             =   2010
            Width           =   195
         End
         Begin VB.PictureBox Picture2 
            BackColor       =   &H00FF0000&
            Height          =   180
            Left            =   3750
            ScaleHeight     =   120
            ScaleWidth      =   135
            TabIndex        =   27
            Top             =   435
            Width           =   195
         End
         Begin VB.PictureBox Picture1 
            BackColor       =   &H00FF0000&
            Height          =   180
            Left            =   975
            ScaleHeight     =   120
            ScaleWidth      =   135
            TabIndex        =   26
            Top             =   435
            Width           =   195
         End
         Begin VB.CheckBox Check1 
            Caption         =   "增強感應區域可見性(&V)"
            Height          =   300
            Left            =   150
            TabIndex        =   19
            Top             =   4170
            Width           =   4650
         End
         Begin VB.TextBox Text2 
            Height          =   270
            Left            =   3390
            MaxLength       =   3
            TabIndex        =   17
            Text            =   "25"
            Top             =   3870
            Width           =   675
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Left            =   1920
            MaxLength       =   3
            TabIndex        =   15
            Text            =   "25"
            Top             =   3870
            Width           =   675
         End
         Begin VB.ComboBox Combo4 
            Height          =   300
            ItemData        =   "Form1.frx":0EF4
            Left            =   960
            List            =   "Form1.frx":0F0D
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   3540
            Width           =   3810
         End
         Begin VB.ComboBox Combo3 
            Height          =   300
            ItemData        =   "Form1.frx":0F6F
            Left            =   960
            List            =   "Form1.frx":0F88
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   3225
            Width           =   3810
         End
         Begin VB.ComboBox Combo2 
            Height          =   300
            ItemData        =   "Form1.frx":0FEA
            Left            =   960
            List            =   "Form1.frx":1003
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   2910
            Width           =   3810
         End
         Begin VB.ComboBox Combo1 
            Height          =   300
            ItemData        =   "Form1.frx":1065
            Left            =   960
            List            =   "Form1.frx":107E
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   2595
            Width           =   3810
         End
         Begin VB.CommandButton Command1 
            Height          =   2310
            Left            =   165
            MaskColor       =   &H00FFFFFF&
            Picture         =   "Form1.frx":10E0
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   210
            UseMaskColor    =   -1  'True
            Width           =   4605
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "像素"
            Height          =   180
            Left            =   4140
            TabIndex        =   18
            Top             =   3900
            Width           =   360
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "像素 X"
            Height          =   180
            Left            =   2670
            TabIndex        =   16
            Top             =   3900
            Width           =   540
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "感應區域大小:"
            Height          =   180
            Left            =   165
            TabIndex        =   14
            Top             =   3900
            Width           =   1170
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "右下角:"
            Height          =   180
            Left            =   165
            TabIndex        =   12
            Top             =   3600
            Width           =   630
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "右上角:"
            Height          =   180
            Left            =   165
            TabIndex        =   10
            Top             =   3285
            Width           =   630
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "左下角:"
            Height          =   180
            Left            =   165
            TabIndex        =   8
            Top             =   2970
            Width           =   630
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "左上角:"
            Height          =   180
            Left            =   165
            TabIndex        =   6
            Top             =   2655
            Width           =   630
         End
      End
   End
   Begin VB.OptionButton Option2 
      Caption         =   "啟用Super Cursor(&E)"
      Height          =   330
      Left            =   120
      TabIndex        =   2
      Top             =   1095
      Width           =   2205
   End
   Begin VB.OptionButton Option1 
      Caption         =   "禁用Super Cursor(&D)"
      Height          =   330
      Left            =   120
      TabIndex        =   1
      Top             =   780
      Value           =   -1  'True
      Width           =   2205
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"Form1.frx":2511E
      Height          =   600
      Left            =   765
      TabIndex        =   0
      Top             =   135
      Width           =   4725
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   30
      Picture         =   "Form1.frx":251B3
      Top             =   30
      Width           =   720
   End
   Begin VB.Menu mnuTrayMenu 
      Caption         =   "TrayMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuEnable 
         Caption         =   "啟用Super Cursor(&E)"
      End
      Begin VB.Menu b3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuShow 
         Caption         =   "顯示程序主窗口(&S)"
      End
      Begin VB.Menu b2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "退出(&E)"
      End
   End
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
Attribute VB_Name = "FormMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const SC_SCREENSAVE = &HF140&
Dim IsCodeUse As Boolean
Private Declare Function WinExec Lib "kernel32" (ByVal lpCmdLine As String, ByVal nCmdShow As Long) As Long
Private Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal MSG As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Const GWL_WNDPROC = (-4)
Private Const WM_SYSCOMMAND = &H112
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
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
'很多朋友都见到过能在托盘图标上出现气球提示的软件，不说软件，就是在“磁盘空间不足”时Windows给出的提示就属于气球提示，那么怎样在自己的程序中添加这样的气球提示呢？
   
'其实并不难，关键就在添加托盘图标时所使用的NOTIFYICONDATA结构，源代码如下：
Private Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
   
Private Type NOTIFYICONDATA
cbSize   As Long     '   结构大小(字节)
hwnd   As Long     '   处理消息的窗口的句柄
uID   As Long     '   唯一的标识符
uFlags   As Long     '   Flags
uCallbackMessage   As Long     '   处理消息的窗口接收的消息
hIcon   As Long     '   托盘图标句柄
szTip   As String * 128         '   Tooltip   提示文本
dwState   As Long     '   托盘图标状态
dwStateMask   As Long     '   状态掩码
szInfo   As String * 256         '   气球提示文本
uTimeoutOrVersion   As Long     '   气球提示消失时间或版本
'   uTimeout   -   气球提示消失时间(单位:ms,   10000   --   30000)
'   uVersion   -   版本(0   for   V4,   3   for   V5)
szInfoTitle   As String * 64         '   气球提示标题
dwInfoFlags   As Long     '   气球提示图标
End Type
   
'   dwState   to   NOTIFYICONDATA   structure
Private Const NIS_HIDDEN = &H1           '   隐藏图标
Private Const NIS_SHAREDICON = &H2           '   共享图标
   
'   dwInfoFlags   to   NOTIFIICONDATA   structure
Private Const NIIF_NONE = &H0           '   无图标
Private Const NIIF_INFO = &H1           '   "消息"图标
Private Const NIIF_WARNING = &H2           '   "警告"图标
Private Const NIIF_ERROR = &H3           '   "错误"图标
   
'   uFlags   to   NOTIFYICONDATA   structure
Private Const NIF_ICON       As Long = &H2
Private Const NIF_INFO       As Long = &H10
Private Const NIF_MESSAGE       As Long = &H1
Private Const NIF_STATE       As Long = &H8
Private Const NIF_TIP       As Long = &H4
   
'   dwMessage   to   Shell_NotifyIcon
Private Const NIM_ADD       As Long = &H0
Private Const NIM_DELETE       As Long = &H2
Private Const NIM_MODIFY       As Long = &H1
Private Const NIM_SETFOCUS       As Long = &H3
Private Const NIM_SETVERSION       As Long = &H4
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
Private Sub Check1_Click()
On Error Resume Next
SetWindowPos LeftTop.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
SetWindowPos LeftBotton.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
SetWindowPos RightTop.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
SetWindowPos RightBotton.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
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
End Sub
Private Sub Check1_GotFocus()
On Error Resume Next
Picture1.BackColor = RGB(0, 0, 245)
Picture2.BackColor = RGB(0, 0, 245)
Picture3.BackColor = RGB(0, 0, 245)
Picture4.BackColor = RGB(0, 0, 245)
End Sub
Private Sub Check2_Click()
On Error Resume Next
Picture1.BackColor = RGB(0, 0, 245)
Picture2.BackColor = RGB(0, 0, 245)
Picture3.BackColor = RGB(0, 0, 245)
Picture4.BackColor = RGB(0, 0, 245)
End Sub
Private Sub Check2_GotFocus()
On Error Resume Next
Picture1.BackColor = RGB(0, 0, 245)
Picture2.BackColor = RGB(0, 0, 245)
Picture3.BackColor = RGB(0, 0, 245)
Picture4.BackColor = RGB(0, 0, 245)
End Sub
Private Sub Check3_Click()
On Error Resume Next
If Check3.Value = 1 Then
Dim ans As Integer
ans = MsgBox("警告:不推薦強制結束進程,那可能導致數據丟失或不可知問題,繼續?", vbExclamation + vbYesNo, "Alert")
If ans = vbYes Then
Err.Clear
Exit Sub
Else
Check3.Value = 0
End If
End If
Err.Clear
End Sub
Private Sub Check3_GotFocus()
On Error Resume Next
Picture1.BackColor = RGB(0, 0, 245)
Picture2.BackColor = RGB(0, 0, 245)
Picture3.BackColor = RGB(0, 0, 245)
Picture4.BackColor = RGB(0, 0, 245)
End Sub
Private Sub Combo1_Change()
On Error Resume Next
Select Case Combo1.ListIndex
Case 0
LeftTop.Tag = "Shutdown"
Case 1
LeftTop.Tag = "Reboot"
Case 2
LeftTop.Tag = "Logoff"
Case 3
LeftTop.Tag = "ScreenOff"
Case 4
LeftTop.Tag = "ScreenProtection"
Case 5
LeftTop.Tag = "Menu"
Case 6
LeftTop.Tag = "None"
End Select
Select Case Combo2.ListIndex
Case 0
LeftBotton.Tag = "Shutdown"
Case 1
LeftBotton.Tag = "Reboot"
Case 2
LeftBotton.Tag = "Logoff"
Case 3
LeftBotton.Tag = "ScreenOff"
Case 4
LeftBotton.Tag = "ScreenProtection"
Case 5
LeftBotton.Tag = "Menu"
Case 6
LeftBotton.Tag = "None"
End Select
Select Case Combo3.ListIndex
Case 0
RightTop.Tag = "Shutdown"
Case 1
RightTop.Tag = "Reboot"
Case 2
RightTop.Tag = "Logoff"
Case 3
RightTop.Tag = "ScreenOff"
Case 4
RightTop.Tag = "ScreenProtection"
Case 5
RightTop.Tag = "Menu"
Case 6
RightTop.Tag = "None"
End Select
Select Case Combo4.ListIndex
Case 0
RightBotton.Tag = "Shutdown"
Case 1
RightBotton.Tag = "Reboot"
Case 2
RightBotton.Tag = "Logoff"
Case 3
RightBotton.Tag = "ScreenOff"
Case 4
RightBotton.Tag = "ScreenProtection"
Case 5
RightBotton.Tag = "Menu"
Case 6
RightBotton.Tag = "None"
End Select
Debug.Print LeftTop.Tag
Debug.Print LeftBotton.Tag
Debug.Print RightTop.Tag
Debug.Print RightBotton.Tag
End Sub
Private Sub Combo1_Click()
On Error Resume Next
Select Case Combo1.ListIndex
Case 0
LeftTop.Tag = "Shutdown"
Case 1
LeftTop.Tag = "Reboot"
Case 2
LeftTop.Tag = "Logoff"
Case 3
LeftTop.Tag = "ScreenOff"
Case 4
LeftTop.Tag = "ScreenProtection"
Case 5
LeftTop.Tag = "Menu"
Case 6
LeftTop.Tag = "None"
End Select
Select Case Combo2.ListIndex
Case 0
LeftBotton.Tag = "Shutdown"
Case 1
LeftBotton.Tag = "Reboot"
Case 2
LeftBotton.Tag = "Logoff"
Case 3
LeftBotton.Tag = "ScreenOff"
Case 4
LeftBotton.Tag = "ScreenProtection"
Case 5
LeftBotton.Tag = "Menu"
Case 6
LeftBotton.Tag = "None"
End Select
Select Case Combo3.ListIndex
Case 0
RightTop.Tag = "Shutdown"
Case 1
RightTop.Tag = "Reboot"
Case 2
RightTop.Tag = "Logoff"
Case 3
RightTop.Tag = "ScreenOff"
Case 4
RightTop.Tag = "ScreenProtection"
Case 5
RightTop.Tag = "Menu"
Case 6
RightTop.Tag = "None"
End Select
Select Case Combo4.ListIndex
Case 0
RightBotton.Tag = "Shutdown"
Case 1
RightBotton.Tag = "Reboot"
Case 2
RightBotton.Tag = "Logoff"
Case 3
RightBotton.Tag = "ScreenOff"
Case 4
RightBotton.Tag = "ScreenProtection"
Case 5
RightBotton.Tag = "Menu"
Case 6
RightBotton.Tag = "None"
End Select
Debug.Print LeftTop.Tag
Debug.Print LeftBotton.Tag
Debug.Print RightTop.Tag
Debug.Print RightBotton.Tag
End Sub
Private Sub Combo1_GotFocus()
On Error Resume Next
Picture1.BackColor = RGB(0, 245, 245)
Picture2.BackColor = RGB(0, 0, 245)
Picture3.BackColor = RGB(0, 0, 245)
Picture4.BackColor = RGB(0, 0, 245)
End Sub
Private Sub Combo1_LostFocus()
On Error Resume Next
Picture1.BackColor = RGB(0, 0, 245)
Picture2.BackColor = RGB(0, 0, 245)
Picture3.BackColor = RGB(0, 0, 245)
Picture4.BackColor = RGB(0, 0, 245)
End Sub
Private Sub Combo2_Change()
On Error Resume Next
Select Case Combo1.ListIndex
Case 0
LeftTop.Tag = "Shutdown"
Case 1
LeftTop.Tag = "Reboot"
Case 2
LeftTop.Tag = "Logoff"
Case 3
LeftTop.Tag = "ScreenOff"
Case 4
LeftTop.Tag = "ScreenProtection"
Case 5
LeftTop.Tag = "Menu"
Case 6
LeftTop.Tag = "None"
End Select
Select Case Combo2.ListIndex
Case 0
LeftBotton.Tag = "Shutdown"
Case 1
LeftBotton.Tag = "Reboot"
Case 2
LeftBotton.Tag = "Logoff"
Case 3
LeftBotton.Tag = "ScreenOff"
Case 4
LeftBotton.Tag = "ScreenProtection"
Case 5
LeftBotton.Tag = "Menu"
Case 6
LeftBotton.Tag = "None"
End Select
Select Case Combo3.ListIndex
Case 0
RightTop.Tag = "Shutdown"
Case 1
RightTop.Tag = "Reboot"
Case 2
RightTop.Tag = "Logoff"
Case 3
RightTop.Tag = "ScreenOff"
Case 4
RightTop.Tag = "ScreenProtection"
Case 5
RightTop.Tag = "Menu"
Case 6
RightTop.Tag = "None"
End Select
Select Case Combo4.ListIndex
Case 0
RightBotton.Tag = "Shutdown"
Case 1
RightBotton.Tag = "Reboot"
Case 2
RightBotton.Tag = "Logoff"
Case 3
RightBotton.Tag = "ScreenOff"
Case 4
RightBotton.Tag = "ScreenProtection"
Case 5
RightBotton.Tag = "Menu"
Case 6
RightBotton.Tag = "None"
End Select
Debug.Print LeftTop.Tag
Debug.Print LeftBotton.Tag
Debug.Print RightTop.Tag
Debug.Print RightBotton.Tag
End Sub
Private Sub Combo2_Click()
On Error Resume Next
Select Case Combo1.ListIndex
Case 0
LeftTop.Tag = "Shutdown"
Case 1
LeftTop.Tag = "Reboot"
Case 2
LeftTop.Tag = "Logoff"
Case 3
LeftTop.Tag = "ScreenOff"
Case 4
LeftTop.Tag = "ScreenProtection"
Case 5
LeftTop.Tag = "Menu"
Case 6
LeftTop.Tag = "None"
End Select
Select Case Combo2.ListIndex
Case 0
LeftBotton.Tag = "Shutdown"
Case 1
LeftBotton.Tag = "Reboot"
Case 2
LeftBotton.Tag = "Logoff"
Case 3
LeftBotton.Tag = "ScreenOff"
Case 4
LeftBotton.Tag = "ScreenProtection"
Case 5
LeftBotton.Tag = "Menu"
Case 6
LeftBotton.Tag = "None"
End Select
Select Case Combo3.ListIndex
Case 0
RightTop.Tag = "Shutdown"
Case 1
RightTop.Tag = "Reboot"
Case 2
RightTop.Tag = "Logoff"
Case 3
RightTop.Tag = "ScreenOff"
Case 4
RightTop.Tag = "ScreenProtection"
Case 5
RightTop.Tag = "Menu"
Case 6
RightTop.Tag = "None"
End Select
Select Case Combo4.ListIndex
Case 0
RightBotton.Tag = "Shutdown"
Case 1
RightBotton.Tag = "Reboot"
Case 2
RightBotton.Tag = "Logoff"
Case 3
RightBotton.Tag = "ScreenOff"
Case 4
RightBotton.Tag = "ScreenProtection"
Case 5
RightBotton.Tag = "Menu"
Case 6
RightBotton.Tag = "None"
End Select
Debug.Print LeftTop.Tag
Debug.Print LeftBotton.Tag
Debug.Print RightTop.Tag
Debug.Print RightBotton.Tag
End Sub
Private Sub Combo2_GotFocus()
On Error Resume Next
Picture1.BackColor = RGB(0, 0, 245)
Picture2.BackColor = RGB(0, 0, 245)
Picture3.BackColor = RGB(0, 245, 245)
Picture4.BackColor = RGB(0, 0, 245)
End Sub
Private Sub Combo2_LostFocus()
On Error Resume Next
Picture1.BackColor = RGB(0, 0, 245)
Picture2.BackColor = RGB(0, 0, 245)
Picture3.BackColor = RGB(0, 0, 245)
Picture4.BackColor = RGB(0, 0, 245)
End Sub
Private Sub Combo3_Change()
On Error Resume Next
Select Case Combo1.ListIndex
Case 0
LeftTop.Tag = "Shutdown"
Case 1
LeftTop.Tag = "Reboot"
Case 2
LeftTop.Tag = "Logoff"
Case 3
LeftTop.Tag = "ScreenOff"
Case 4
LeftTop.Tag = "ScreenProtection"
Case 5
LeftTop.Tag = "Menu"
Case 6
LeftTop.Tag = "None"
End Select
Select Case Combo2.ListIndex
Case 0
LeftBotton.Tag = "Shutdown"
Case 1
LeftBotton.Tag = "Reboot"
Case 2
LeftBotton.Tag = "Logoff"
Case 3
LeftBotton.Tag = "ScreenOff"
Case 4
LeftBotton.Tag = "ScreenProtection"
Case 5
LeftBotton.Tag = "Menu"
Case 6
LeftBotton.Tag = "None"
End Select
Select Case Combo3.ListIndex
Case 0
RightTop.Tag = "Shutdown"
Case 1
RightTop.Tag = "Reboot"
Case 2
RightTop.Tag = "Logoff"
Case 3
RightTop.Tag = "ScreenOff"
Case 4
RightTop.Tag = "ScreenProtection"
Case 5
RightTop.Tag = "Menu"
Case 6
RightTop.Tag = "None"
End Select
Select Case Combo4.ListIndex
Case 0
RightBotton.Tag = "Shutdown"
Case 1
RightBotton.Tag = "Reboot"
Case 2
RightBotton.Tag = "Logoff"
Case 3
RightBotton.Tag = "ScreenOff"
Case 4
RightBotton.Tag = "ScreenProtection"
Case 5
RightBotton.Tag = "Menu"
Case 6
RightBotton.Tag = "None"
End Select
Debug.Print LeftTop.Tag
Debug.Print LeftBotton.Tag
Debug.Print RightTop.Tag
Debug.Print RightBotton.Tag
End Sub
Private Sub Combo3_Click()
On Error Resume Next
Select Case Combo1.ListIndex
Case 0
LeftTop.Tag = "Shutdown"
Case 1
LeftTop.Tag = "Reboot"
Case 2
LeftTop.Tag = "Logoff"
Case 3
LeftTop.Tag = "ScreenOff"
Case 4
LeftTop.Tag = "ScreenProtection"
Case 5
LeftTop.Tag = "Menu"
Case 6
LeftTop.Tag = "None"
End Select
Select Case Combo2.ListIndex
Case 0
LeftBotton.Tag = "Shutdown"
Case 1
LeftBotton.Tag = "Reboot"
Case 2
LeftBotton.Tag = "Logoff"
Case 3
LeftBotton.Tag = "ScreenOff"
Case 4
LeftBotton.Tag = "ScreenProtection"
Case 5
LeftBotton.Tag = "Menu"
Case 6
LeftBotton.Tag = "None"
End Select
Select Case Combo3.ListIndex
Case 0
RightTop.Tag = "Shutdown"
Case 1
RightTop.Tag = "Reboot"
Case 2
RightTop.Tag = "Logoff"
Case 3
RightTop.Tag = "ScreenOff"
Case 4
RightTop.Tag = "ScreenProtection"
Case 5
RightTop.Tag = "Menu"
Case 6
RightTop.Tag = "None"
End Select
Select Case Combo4.ListIndex
Case 0
RightBotton.Tag = "Shutdown"
Case 1
RightBotton.Tag = "Reboot"
Case 2
RightBotton.Tag = "Logoff"
Case 3
RightBotton.Tag = "ScreenOff"
Case 4
RightBotton.Tag = "ScreenProtection"
Case 5
RightBotton.Tag = "Menu"
Case 6
RightBotton.Tag = "None"
End Select
Debug.Print LeftTop.Tag
Debug.Print LeftBotton.Tag
Debug.Print RightTop.Tag
Debug.Print RightBotton.Tag
End Sub
Private Sub Combo3_GotFocus()
On Error Resume Next
Picture1.BackColor = RGB(0, 0, 245)
Picture2.BackColor = RGB(0, 245, 245)
Picture3.BackColor = RGB(0, 0, 245)
Picture4.BackColor = RGB(0, 0, 245)
End Sub
Private Sub Combo3_LostFocus()
On Error Resume Next
Picture1.BackColor = RGB(0, 0, 245)
Picture2.BackColor = RGB(0, 0, 245)
Picture3.BackColor = RGB(0, 0, 245)
Picture4.BackColor = RGB(0, 0, 245)
End Sub
Private Sub Combo4_Change()
On Error Resume Next
Select Case Combo1.ListIndex
Case 0
LeftTop.Tag = "Shutdown"
Case 1
LeftTop.Tag = "Reboot"
Case 2
LeftTop.Tag = "Logoff"
Case 3
LeftTop.Tag = "ScreenOff"
Case 4
LeftTop.Tag = "ScreenProtection"
Case 5
LeftTop.Tag = "Menu"
Case 6
LeftTop.Tag = "None"
End Select
Select Case Combo2.ListIndex
Case 0
LeftBotton.Tag = "Shutdown"
Case 1
LeftBotton.Tag = "Reboot"
Case 2
LeftBotton.Tag = "Logoff"
Case 3
LeftBotton.Tag = "ScreenOff"
Case 4
LeftBotton.Tag = "ScreenProtection"
Case 5
LeftBotton.Tag = "Menu"
Case 6
LeftBotton.Tag = "None"
End Select
Select Case Combo3.ListIndex
Case 0
RightTop.Tag = "Shutdown"
Case 1
RightTop.Tag = "Reboot"
Case 2
RightTop.Tag = "Logoff"
Case 3
RightTop.Tag = "ScreenOff"
Case 4
RightTop.Tag = "ScreenProtection"
Case 5
RightTop.Tag = "Menu"
Case 6
RightTop.Tag = "None"
End Select
Select Case Combo4.ListIndex
Case 0
RightBotton.Tag = "Shutdown"
Case 1
RightBotton.Tag = "Reboot"
Case 2
RightBotton.Tag = "Logoff"
Case 3
RightBotton.Tag = "ScreenOff"
Case 4
RightBotton.Tag = "ScreenProtection"
Case 5
RightBotton.Tag = "Menu"
Case 6
RightBotton.Tag = "None"
End Select
Debug.Print LeftTop.Tag
Debug.Print LeftBotton.Tag
Debug.Print RightTop.Tag
Debug.Print RightBotton.Tag
End Sub
Private Sub Combo4_Click()
On Error Resume Next
Select Case Combo1.ListIndex
Case 0
LeftTop.Tag = "Shutdown"
Case 1
LeftTop.Tag = "Reboot"
Case 2
LeftTop.Tag = "Logoff"
Case 3
LeftTop.Tag = "ScreenOff"
Case 4
LeftTop.Tag = "ScreenProtection"
Case 5
LeftTop.Tag = "Menu"
Case 6
LeftTop.Tag = "None"
End Select
Select Case Combo2.ListIndex
Case 0
LeftBotton.Tag = "Shutdown"
Case 1
LeftBotton.Tag = "Reboot"
Case 2
LeftBotton.Tag = "Logoff"
Case 3
LeftBotton.Tag = "ScreenOff"
Case 4
LeftBotton.Tag = "ScreenProtection"
Case 5
LeftBotton.Tag = "Menu"
Case 6
LeftBotton.Tag = "None"
End Select
Select Case Combo3.ListIndex
Case 0
RightTop.Tag = "Shutdown"
Case 1
RightTop.Tag = "Reboot"
Case 2
RightTop.Tag = "Logoff"
Case 3
RightTop.Tag = "ScreenOff"
Case 4
RightTop.Tag = "ScreenProtection"
Case 5
RightTop.Tag = "Menu"
Case 6
RightTop.Tag = "None"
End Select
Select Case Combo4.ListIndex
Case 0
RightBotton.Tag = "Shutdown"
Case 1
RightBotton.Tag = "Reboot"
Case 2
RightBotton.Tag = "Logoff"
Case 3
RightBotton.Tag = "ScreenOff"
Case 4
RightBotton.Tag = "ScreenProtection"
Case 5
RightBotton.Tag = "Menu"
Case 6
RightBotton.Tag = "None"
End Select
Debug.Print LeftTop.Tag
Debug.Print LeftBotton.Tag
Debug.Print RightTop.Tag
Debug.Print RightBotton.Tag
End Sub
Private Sub Combo4_GotFocus()
On Error Resume Next
Picture1.BackColor = RGB(0, 0, 245)
Picture2.BackColor = RGB(0, 0, 245)
Picture3.BackColor = RGB(0, 0, 245)
Picture4.BackColor = RGB(0, 245, 245)
End Sub
Private Sub Combo4_LostFocus()
On Error Resume Next
Picture1.BackColor = RGB(0, 0, 245)
Picture2.BackColor = RGB(0, 0, 245)
Picture3.BackColor = RGB(0, 0, 245)
Picture4.BackColor = RGB(0, 0, 245)
End Sub
Private Sub Command2_Click()
On Error Resume Next
With Me
.Hide
End With
'With Me.cSysTray1
'.InTray = True
'.TrayTip = "Cursor - 雙擊還原主窗口,右擊顯示菜單"
'‘End With
SetWindowPos LeftTop.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
SetWindowPos LeftBotton.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
SetWindowPos RightTop.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
SetWindowPos RightBotton.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
Dim IconData     As NOTIFYICONDATA
Dim title     As String
title = "Super Cursor - 雙擊還原主窗口,右擊顯示菜單" & vbNullChar
With IconData
.cbSize = Len(IconData)
.hwnd = Me.hwnd
.uID = 0
.uFlags = NIF_TIP Or NIF_ICON Or NIF_MESSAGE Or NIF_INFO Or NIF_STATE
.uCallbackMessage = WM_NOTIFYICON
.szTip = title
.hIcon = Me.Icon.Handle
.dwState = 0
.dwStateMask = 0
.szInfo = "Super Cursor 已經最小化，點擊顯示窗口。雙擊還原主窗口,右擊顯示菜單" & vbNullChar
.szInfoTitle = title
.dwInfoFlags = NIIF_INFO
.uTimeoutOrVersion = 10000
End With
Shell_NotifyIcon NIM_ADD, IconData
preWndProc = SetWindowLong(Me.hwnd, GWL_WNDPROC, AddressOf WindowProc)
On Error Resume Next
With Me
.Hide
End With
'With Me.cSysTray1
'.InTray = True
'.TrayTip = "Cursor - 雙擊還原主窗口,右擊顯示菜單"
'End With
SetWindowPos LeftTop.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
SetWindowPos LeftBotton.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
SetWindowPos RightTop.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
SetWindowPos RightBotton.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End Sub
Private Sub Command3_Click()
On Error Resume Next
PopupMenu Me.UserMenu
End Sub
Private Sub Command4_Click()
On Error Resume Next
PopupMenu Me.UserMenu
End Sub
Private Sub cSysTray1_MouseDblClick(Button As Integer, Id As Long)
On Error Resume Next
FormMain.Show
On Error Resume Next
With FormMain
.Show
End With
'   删除托盘区图标
Dim IconData     As NOTIFYICONDATA
With IconData
.cbSize = Len(IconData)
.hwnd = FormMain.hwnd
.uID = 0
.uFlags = NIF_TIP Or NIF_ICON Or NIF_MESSAGE
.uCallbackMessage = WM_NOTIFYICON
.szTip = "Super Cursor - 雙擊顯示窗口，右擊顯示菜單"
.hIcon = FormMain.Icon.Handle
End With
Shell_NotifyIcon NIM_DELETE, IconData
SetWindowLong FormMain.hwnd, GWL_WNDPROC, preWndProc
'With FormMain.cSysTray1
'.InTray = False
'.TrayTip = "Super Cursor - 雙擊還原主窗口,右擊顯示菜單"
'End With
End Sub
Private Sub cSysTray1_MouseDown(Button As Integer, Id As Long)
On Error Resume Next
If Button = 2 Then
PopupMenu Me.mnuTrayMenu
End If
End Sub
Private Sub cSysTray1_MouseMove(Id As Long)
On Error Resume Next
Exit Sub
End Sub
Private Sub Form_Load()
On Error Resume Next
Dim hMenu As Long
hMenu = GetSystemMenu(hwnd, False)
AppendMenu hMenu, 0, 0, vbNullString
AppendMenu hMenu, 0, MENUITEM_1, "關於Super Cursor(&A)..."
PrevWndFunc = SetWindowLong(hwnd, GWL_WNDPROC, AddressOf SystemMenuCallback)
Command4.MaskColor = &HC8D0D4
Command4.UseMaskColor = True
Label2.Enabled = False
Label3.Enabled = False
Label4.Enabled = False
Label5.Enabled = False
Label6.Enabled = False
Label7.Enabled = False
Label8.Enabled = False
With Picture1
.Enabled = False
.BackColor = RGB(125, 125, 125)
End With
With Picture2
.Enabled = False
.BackColor = RGB(125, 125, 125)
End With
With Picture3
.Enabled = False
.BackColor = RGB(125, 125, 125)
End With
With Picture4
.Enabled = False
.BackColor = RGB(125, 125, 125)
End With
Picture1.Enabled = False
Picture2.Enabled = False
Picture3.Enabled = False
Picture4.Enabled = False
Check1.Enabled = False
Check2.Enabled = False
Check3.Enabled = False
Command1.Enabled = False
Text1.Enabled = False
Text2.Enabled = False
Combo1.Enabled = False
Combo2.Enabled = False
Combo3.Enabled = False
Combo4.Enabled = False
Option1.Value = True
Combo1.ListIndex = 0
Combo2.ListIndex = 1
Combo3.ListIndex = 2
Combo4.ListIndex = 5
Check2.Value = 1
Check3.Value = 0
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
.Width = CInt(Text1.Text) * Screen.TwipsPerPixelX
.Height = CInt(Text2.Text) * Screen.TwipsPerPixelY
.Top = Screen.Height - .Height
.BackColor = RGB(0, 245, 245)
End With
With LeftTop
.Left = 0
.Width = CInt(Text1.Text) * Screen.TwipsPerPixelX
.Height = CInt(Text2.Text) * Screen.TwipsPerPixelY
.Top = 0
.BackColor = RGB(0, 245, 245)
End With
With RightTop
.Width = CInt(Text1.Text) * Screen.TwipsPerPixelX
.Left = Screen.Width - .Width
.Height = CInt(Text2.Text) * Screen.TwipsPerPixelY
.Top = 0
.BackColor = RGB(0, 245, 245)
End With
With RightBotton
.Width = CInt(Text1.Text) * Screen.TwipsPerPixelX
.Height = CInt(Text2.Text) * Screen.TwipsPerPixelY
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
Unload LeftTop
Unload LeftBotton
Unload RightTop
Unload RightBotton
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
Unload LeftTop
Unload LeftBotton
Unload RightTop
Unload RightBotton
Unload Me
End Sub
Private Sub Form_Terminate()
On Error Resume Next
Unload LeftTop
Unload LeftBotton
Unload RightTop
Unload RightBotton
Unload Me
End
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Unload LeftTop
Unload LeftBotton
Unload RightTop
Unload RightBotton
Unload Me
End
End Sub
Private Sub mnuAXA_Click()
On Error Resume Next
frmMain.Show
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
Private Sub mnuDS_Click()
On Error Resume Next
Form1.Show
End Sub
Private Sub mnuEnable_Click()
On Error Resume Next
IsCodeUse = True
Select Case mnuEnable.Checked
Case False
mnuEnable.Checked = True
Me.Option2.Value = True
On Error Resume Next
Label2.Enabled = True
Label3.Enabled = True
Label4.Enabled = True
Label5.Enabled = True
Label6.Enabled = True
Label7.Enabled = True
Label8.Enabled = True
Check1.Enabled = True
Picture1.Enabled = True
Picture2.Enabled = True
Picture3.Enabled = True
Picture4.Enabled = True
Check2.Enabled = True
Check3.Enabled = True
Command1.Enabled = True
Text1.Enabled = True
Text2.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True
Combo4.Enabled = True
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
.Width = CInt(Text1.Text) * Screen.TwipsPerPixelX
.Height = CInt(Text2.Text) * Screen.TwipsPerPixelY
.Top = Screen.Height - .Height
.BackColor = RGB(0, 245, 245)
End With
With LeftTop
.Left = 0
.Width = CInt(Text1.Text) * Screen.TwipsPerPixelX
.Height = CInt(Text2.Text) * Screen.TwipsPerPixelY
.Top = 0
.BackColor = RGB(0, 245, 245)
End With
With RightTop
.Width = CInt(Text1.Text) * Screen.TwipsPerPixelX
.Left = Screen.Width - .Width
.Height = CInt(Text2.Text) * Screen.TwipsPerPixelY
.Top = 0
.BackColor = RGB(0, 245, 245)
End With
With RightBotton
.Width = CInt(Text1.Text) * Screen.TwipsPerPixelX
.Height = CInt(Text2.Text) * Screen.TwipsPerPixelY
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
Select Case Combo1.ListIndex
Case 0
LeftTop.Tag = "Shutdown"
Case 1
LeftTop.Tag = "Reboot"
Case 2
LeftTop.Tag = "Logoff"
Case 3
LeftTop.Tag = "ScreenOff"
Case 4
LeftTop.Tag = "ScreenProtection"
Case 5
LeftTop.Tag = "Menu"
Case 6
LeftTop.Tag = "None"
End Select
Select Case Combo2.ListIndex
Case 0
LeftBotton.Tag = "Shutdown"
Case 1
LeftBotton.Tag = "Reboot"
Case 2
LeftBotton.Tag = "Logoff"
Case 3
LeftBotton.Tag = "ScreenOff"
Case 4
LeftBotton.Tag = "ScreenProtection"
Case 5
LeftBotton.Tag = "Menu"
Case 6
LeftBotton.Tag = "None"
End Select
Select Case Combo3.ListIndex
Case 0
RightTop.Tag = "Shutdown"
Case 1
RightTop.Tag = "Reboot"
Case 2
RightTop.Tag = "Logoff"
Case 3
RightTop.Tag = "ScreenOff"
Case 4
RightTop.Tag = "ScreenProtection"
Case 5
RightTop.Tag = "Menu"
Case 6
RightTop.Tag = "None"
End Select
Select Case Combo4.ListIndex
Case 0
RightBotton.Tag = "Shutdown"
Case 1
RightBotton.Tag = "Reboot"
Case 2
RightBotton.Tag = "Logoff"
Case 3
RightBotton.Tag = "ScreenOff"
Case 4
RightBotton.Tag = "ScreenProtection"
Case 5
RightBotton.Tag = "Menu"
Case 6
RightBotton.Tag = "None"
End Select
Debug.Print LeftTop.Tag
Debug.Print LeftBotton.Tag
Debug.Print RightTop.Tag
Debug.Print RightBotton.Tag
Case True
mnuEnable.Checked = False
Option1.Value = True
On Error Resume Next
Label2.Enabled = False
Label3.Enabled = False
Label4.Enabled = False
Label5.Enabled = False
Label6.Enabled = False
Label7.Enabled = False
Label8.Enabled = False
Check1.Enabled = False
Picture1.Enabled = False
Picture2.Enabled = False
Picture3.Enabled = False
Picture4.Enabled = False
Check2.Enabled = False
Check3.Enabled = False
Command1.Enabled = False
Text1.Enabled = False
Text2.Enabled = False
Combo1.Enabled = False
Combo2.Enabled = False
Combo3.Enabled = False
Combo4.Enabled = False
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
Private Sub mnuFileInfo_Click()
On Error Resume Next
frmFI.Show
End Sub
Private Sub mnuKillGray_Click()
On Error Resume Next
frmKill.Show
End Sub
Private Sub mnuLargeFile_Click()
On Error Resume Next
frmLargefile.Show
End Sub
Private Sub mnuLockKeys_Click()
On Error Resume Next
frmLock.Show
End Sub
Private Sub mnuMClip_Click()
On Error Resume Next
frmMCM.Show
End Sub
Private Sub mnuPerform_Click()
On Error GoTo ep
Shell "mmc.exe -k perfmon.msc", vbNormalFocus
Exit Sub
ep:
MsgBox "发生系统错误:" & Chr(13) & Err.Description & Chr(13) & "您的系统版本可能不支持此功能或者您尚未安装这个功能.", vbCritical, "Error"
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
FormMain.Show
On Error Resume Next
With FormMain
.Show
End With
'   删除托盘区图标
Dim IconData     As NOTIFYICONDATA
With IconData
.cbSize = Len(IconData)
.hwnd = FormMain.hwnd
.uID = 0
.uFlags = NIF_TIP Or NIF_ICON Or NIF_MESSAGE
.uCallbackMessage = WM_NOTIFYICON
.szTip = "Super Cursor - 雙擊顯示窗口，右擊顯示菜單"
.hIcon = FormMain.Icon.Handle
End With
Shell_NotifyIcon NIM_DELETE, IconData
SetWindowLong FormMain.hwnd, GWL_WNDPROC, preWndProc
'With FormMain.cSysTray1
'.InTray = False
'.TrayTip = "Super Cursor - 雙擊還原主窗口,右擊顯示菜單"
'End With
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
Private Sub mnuTaskApp_Click()
On Error Resume Next
frmTaskMain.Show
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
Private Sub mnuVClick_Click()
On Error Resume Next
FormVC1.Show
End Sub
Private Sub Option1_Click()
On Error Resume Next
mnuEnable.Checked = False
Label2.Enabled = False
Label3.Enabled = False
Label4.Enabled = False
Label5.Enabled = False
Label6.Enabled = False
Label7.Enabled = False
Label8.Enabled = False
Check1.Enabled = False
Picture1.Enabled = False
Picture2.Enabled = False
Picture3.Enabled = False
Picture4.Enabled = False
Check2.Enabled = False
Check3.Enabled = False
Command1.Enabled = False
Text1.Enabled = False
Text2.Enabled = False
Combo1.Enabled = False
Combo2.Enabled = False
Combo3.Enabled = False
Combo4.Enabled = False
With Picture1
.Enabled = False
.BackColor = RGB(125, 125, 125)
End With
With Picture2
.Enabled = False
.BackColor = RGB(125, 125, 125)
End With
With Picture3
.Enabled = False
.BackColor = RGB(125, 125, 125)
End With
With Picture4
.Enabled = False
.BackColor = RGB(125, 125, 125)
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
Unload LeftTop
Unload LeftBotton
Unload RightTop
Unload RightBotton
End Sub
Private Sub Option2_Click()
On Error Resume Next
If IsCodeUse = True Then
IsCodeUse = False
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
End If
On Error Resume Next
Picture1.BackColor = RGB(0, 0, 245)
Picture2.BackColor = RGB(0, 0, 245)
Picture3.BackColor = RGB(0, 0, 245)
Picture4.BackColor = RGB(0, 0, 245)
mnuEnable.Checked = True
Label2.Enabled = True
Label3.Enabled = True
Label4.Enabled = True
Label5.Enabled = True
Label6.Enabled = True
Label7.Enabled = True
Label8.Enabled = True
Check1.Enabled = True
Picture1.Enabled = True
Picture2.Enabled = True
Picture3.Enabled = True
Picture4.Enabled = True
Check2.Enabled = True
Check3.Enabled = True
Command1.Enabled = True
Text1.Enabled = True
Text2.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True
Combo4.Enabled = True
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
.Width = CInt(Text1.Text) * Screen.TwipsPerPixelX
.Height = CInt(Text2.Text) * Screen.TwipsPerPixelY
.Top = Screen.Height - .Height
.BackColor = RGB(0, 245, 245)
End With
With LeftTop
.Left = 0
.Width = CInt(Text1.Text) * Screen.TwipsPerPixelX
.Height = CInt(Text2.Text) * Screen.TwipsPerPixelY
.Top = 0
.BackColor = RGB(0, 245, 245)
End With
With RightTop
.Width = CInt(Text1.Text) * Screen.TwipsPerPixelX
.Left = Screen.Width - .Width
.Height = CInt(Text2.Text) * Screen.TwipsPerPixelY
.Top = 0
.BackColor = RGB(0, 245, 245)
End With
With RightBotton
.Width = CInt(Text1.Text) * Screen.TwipsPerPixelX
.Height = CInt(Text2.Text) * Screen.TwipsPerPixelY
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
Select Case Combo1.ListIndex
Case 0
LeftTop.Tag = "Shutdown"
Case 1
LeftTop.Tag = "Reboot"
Case 2
LeftTop.Tag = "Logoff"
Case 3
LeftTop.Tag = "ScreenOff"
Case 4
LeftTop.Tag = "ScreenProtection"
Case 5
LeftTop.Tag = "Menu"
Case 6
LeftTop.Tag = "None"
End Select
Select Case Combo2.ListIndex
Case 0
LeftBotton.Tag = "Shutdown"
Case 1
LeftBotton.Tag = "Reboot"
Case 2
LeftBotton.Tag = "Logoff"
Case 3
LeftBotton.Tag = "ScreenOff"
Case 4
LeftBotton.Tag = "ScreenProtection"
Case 5
LeftBotton.Tag = "Menu"
Case 6
LeftBotton.Tag = "None"
End Select
Select Case Combo3.ListIndex
Case 0
RightTop.Tag = "Shutdown"
Case 1
RightTop.Tag = "Reboot"
Case 2
RightTop.Tag = "Logoff"
Case 3
RightTop.Tag = "ScreenOff"
Case 4
RightTop.Tag = "ScreenProtection"
Case 5
RightTop.Tag = "Menu"
Case 6
RightTop.Tag = "None"
End Select
Select Case Combo4.ListIndex
Case 0
RightBotton.Tag = "Shutdown"
Case 1
RightBotton.Tag = "Reboot"
Case 2
RightBotton.Tag = "Logoff"
Case 3
RightBotton.Tag = "ScreenOff"
Case 4
RightBotton.Tag = "ScreenProtection"
Case 5
RightBotton.Tag = "Menu"
Case 6
RightBotton.Tag = "None"
End Select
Debug.Print LeftTop.Tag
Debug.Print LeftBotton.Tag
Debug.Print RightTop.Tag
Debug.Print RightBotton.Tag
Me.SetFocus
SetWindowPos LeftTop.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
SetWindowPos LeftBotton.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
SetWindowPos RightTop.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
SetWindowPos RightBotton.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
Me.SetFocus
Me.Show
Exit Sub
Me.SetFocus
Me.Show
'   向托盘区添加图标
Dim IconData     As NOTIFYICONDATA
Dim title     As String
title = "Super Cursor - 雙擊還原主窗口,右擊顯示菜單" & vbNullChar
With IconData
.cbSize = Len(IconData)
.hwnd = Me.hwnd
.uID = 0
.uFlags = NIF_TIP Or NIF_ICON Or NIF_MESSAGE Or NIF_INFO Or NIF_STATE
.uCallbackMessage = WM_NOTIFYICON
.szTip = title
.hIcon = Me.Icon.Handle
.dwState = 0
.dwStateMask = 0
.szInfo = "Super Cursor 已經最小化，點擊顯示窗口。雙擊顯示窗口，右擊顯示菜單" & vbNullChar
.szInfoTitle = title
.dwInfoFlags = NIIF_INFO
.uTimeoutOrVersion = 10000
End With
Shell_NotifyIcon NIM_ADD, IconData
preWndProc = SetWindowLong(Me.hwnd, GWL_WNDPROC, AddressOf WindowProc)
On Error Resume Next
With Me
.Hide
End With
'With Me.cSysTray1
'.InTray = True
'.TrayTip = "Cursor - 雙擊還原主窗口,右擊顯示菜單"
'End With
SetWindowPos LeftTop.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
SetWindowPos LeftBotton.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
SetWindowPos RightTop.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
SetWindowPos RightBotton.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End Sub
Private Sub Picture1_Click()
On Error Resume Next
Picture1.BackColor = RGB(0, 245, 245)
Picture2.BackColor = RGB(0, 0, 245)
Picture3.BackColor = RGB(0, 0, 245)
Picture4.BackColor = RGB(0, 0, 245)
Combo1.SetFocus
End Sub
Private Sub Picture2_Click()
On Error Resume Next
Picture1.BackColor = RGB(0, 0, 245)
Picture2.BackColor = RGB(0, 245, 245)
Picture3.BackColor = RGB(0, 0, 245)
Picture4.BackColor = RGB(0, 0, 245)
Combo3.SetFocus
End Sub
Private Sub Picture3_Click()
On Error Resume Next
Picture1.BackColor = RGB(0, 0, 245)
Picture2.BackColor = RGB(0, 0, 245)
Picture3.BackColor = RGB(0, 245, 245)
Picture4.BackColor = RGB(0, 0, 245)
Combo2.SetFocus
End Sub
Private Sub Picture4_Click()
On Error Resume Next
Picture1.BackColor = RGB(0, 0, 245)
Picture2.BackColor = RGB(0, 0, 245)
Picture3.BackColor = RGB(0, 0, 245)
Picture4.BackColor = RGB(0, 245, 245)
Combo4.SetFocus
End Sub
Private Sub Text1_Change()
On Error Resume Next
If 1 = 2 Then
If Trim(Text1.Text) = "" Then
Text1.Text = 25
End If
End If
Select Case CInt(Left(Text1.Text, 1))
Case 1
Text1.MaxLength = 3
Case 2
Text1.MaxLength = 2
Case 3
Text1.MaxLength = 2
Case 4
Text1.MaxLength = 2
Case 5
Text1.MaxLength = 2
Case 6
Text1.MaxLength = 2
Case 8
Text1.MaxLength = 2
Case 7
Text1.MaxLength = 2
Case 9
Text1.MaxLength = 2
Case 0
Text1.MaxLength = 3
End Select
With LeftBotton
.Left = 0
.Width = CInt(Text1.Text) * Screen.TwipsPerPixelX
.Height = CInt(Text2.Text) * Screen.TwipsPerPixelY
.Top = Screen.Height - .Height
.BackColor = RGB(0, 245, 245)
End With
With LeftTop
.Left = 0
.Width = CInt(Text1.Text) * Screen.TwipsPerPixelX
.Height = CInt(Text2.Text) * Screen.TwipsPerPixelY
.Top = 0
.BackColor = RGB(0, 245, 245)
End With
With RightTop
.Width = CInt(Text1.Text) * Screen.TwipsPerPixelX
.Left = Screen.Width - .Width
.Height = CInt(Text2.Text) * Screen.TwipsPerPixelY
.Top = 0
.BackColor = RGB(0, 245, 245)
End With
With RightBotton
.Width = CInt(Text1.Text) * Screen.TwipsPerPixelX
.Height = CInt(Text2.Text) * Screen.TwipsPerPixelY
.Top = Screen.Height - .Height
.Left = Screen.Width - .Width
.BackColor = RGB(0, 245, 245)
End With
End Sub
Private Sub Text1_GotFocus()
On Error Resume Next
With Text1
.SelStart = 0
.SelLength = Len(.Text)
End With
On Error Resume Next
Picture1.BackColor = RGB(0, 0, 245)
Picture2.BackColor = RGB(0, 0, 245)
Picture3.BackColor = RGB(0, 0, 245)
Picture4.BackColor = RGB(0, 0, 245)
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii < 48 Or KeyAscii > 57 Then
If KeyAscii <> 8 Then
KeyAscii = 0
End If
End If
End Sub
Private Sub Text1_LostFocus()
On Error Resume Next
If Trim(Text1.Text) = "" Then
Text1.Text = 25
End If
If CInt(Text1.Text) = 0 Then
Text1.Text = 25
End If
End Sub
Private Sub Text2_Change()
On Error Resume Next
If 1 = 2 Then
If Trim(Text2.Text) = "" Then
Text2.Text = 25
End If
End If
Select Case CInt(Left(Text2.Text, 1))
Case 1
Text2.MaxLength = 3
Case 2
Text2.MaxLength = 2
Case 3
Text2.MaxLength = 2
Case 4
Text2.MaxLength = 2
Case 5
Text2.MaxLength = 2
Case 6
Text2.MaxLength = 2
Case 8
Text2.MaxLength = 2
Case 7
Text2.MaxLength = 2
Case 9
Text2.MaxLength = 2
Case 0
Text2.MaxLength = 3
End Select
With LeftBotton
.Left = 0
.Width = CInt(Text1.Text) * Screen.TwipsPerPixelX
.Height = CInt(Text2.Text) * Screen.TwipsPerPixelY
.Top = Screen.Height - .Height
.BackColor = RGB(0, 245, 245)
End With
With LeftTop
.Left = 0
.Width = CInt(Text1.Text) * Screen.TwipsPerPixelX
.Height = CInt(Text2.Text) * Screen.TwipsPerPixelY
.Top = 0
.BackColor = RGB(0, 245, 245)
End With
With RightTop
.Width = CInt(Text1.Text) * Screen.TwipsPerPixelX
.Left = Screen.Width - .Width
.Height = CInt(Text2.Text) * Screen.TwipsPerPixelY
.Top = 0
.BackColor = RGB(0, 245, 245)
End With
With RightBotton
.Width = CInt(Text1.Text) * Screen.TwipsPerPixelX
.Height = CInt(Text2.Text) * Screen.TwipsPerPixelY
.Top = Screen.Height - .Height
.Left = Screen.Width - .Width
.BackColor = RGB(0, 245, 245)
End With
End Sub
Private Sub Text2_GotFocus()
On Error Resume Next
With Text2
.SelStart = 0
.SelLength = Len(.Text)
End With
On Error Resume Next
Picture1.BackColor = RGB(0, 0, 245)
Picture2.BackColor = RGB(0, 0, 245)
Picture3.BackColor = RGB(0, 0, 245)
Picture4.BackColor = RGB(0, 0, 245)
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii < 48 Or KeyAscii > 57 Then
If KeyAscii <> 8 Then
KeyAscii = 0
End If
End If
End Sub
Private Sub Text2_LostFocus()
On Error Resume Next
If Trim(Text2.Text) = "" Then
Text2.Text = 25
End If
If CInt(Text2.Text) = 0 Then
Text2.Text = 25
End If
End Sub
Private Sub Timer1_Timer()
On Error Resume Next
Const SWP_NOACTIVATE = &H10
If Option2.Value = True Then
SetWindowPos LeftTop.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOACTIVATE
SetWindowPos LeftBotton.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOACTIVATE
SetWindowPos RightTop.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOACTIVATE
SetWindowPos RightBotton.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOACTIVATE
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
Select Case Combo1.ListIndex
Case 0
LeftTop.Tag = "Shutdown"
Case 1
LeftTop.Tag = "Reboot"
Case 2
LeftTop.Tag = "Logoff"
Case 3
LeftTop.Tag = "ScreenOff"
Case 4
LeftTop.Tag = "ScreenProtection"
Case 5
LeftTop.Tag = "Menu"
Case 6
LeftTop.Tag = "None"
End Select
Select Case Combo2.ListIndex
Case 0
LeftBotton.Tag = "Shutdown"
Case 1
LeftBotton.Tag = "Reboot"
Case 2
LeftBotton.Tag = "Logoff"
Case 3
LeftBotton.Tag = "ScreenOff"
Case 4
LeftBotton.Tag = "ScreenProtection"
Case 5
LeftBotton.Tag = "Menu"
Case 6
LeftBotton.Tag = "None"
End Select
Select Case Combo3.ListIndex
Case 0
RightTop.Tag = "Shutdown"
Case 1
RightTop.Tag = "Reboot"
Case 2
RightTop.Tag = "Logoff"
Case 3
RightTop.Tag = "ScreenOff"
Case 4
RightTop.Tag = "ScreenProtection"
Case 5
RightTop.Tag = "Menu"
Case 6
RightTop.Tag = "None"
End Select
Select Case Combo4.ListIndex
Case 0
RightBotton.Tag = "Shutdown"
Case 1
RightBotton.Tag = "Reboot"
Case 2
RightBotton.Tag = "Logoff"
Case 3
RightBotton.Tag = "ScreenOff"
Case 4
RightBotton.Tag = "ScreenProtection"
Case 5
RightBotton.Tag = "Menu"
Case 6
RightBotton.Tag = "None"
End Select
Debug.Print LeftTop.Tag
Debug.Print LeftBotton.Tag
Debug.Print RightTop.Tag
Debug.Print RightBotton.Tag
If 1 = 245 Then
Me.SetFocus
End If
Else
Exit Sub
End If
End Sub
Private Sub mnuCharmap_Click()
On Error GoTo ep
Shell "charmap.exe", vbNormalFocus
Exit Sub
ep:
MsgBox "发生系统错误:" & Chr(13) & Err.Description & Chr(13) & "您的系统版本可能不支持此功能或者您尚未安装这个功能.", vbCritical, "Error"
End Sub
Private Sub mnuCleanmgr_Click()
On Error GoTo ep
Shell "cleanmgr.exe", vbNormalFocus
Exit Sub
ep:
MsgBox "发生系统错误:" & Chr(13) & Err.Description & Chr(13) & "您的系统版本可能不支持此功能或者您尚未安装这个功能.", vbCritical, "Error"
End Sub
Private Sub mnuCMD_Click()
On Error GoTo ep
Shell "cmd.exe", vbNormalFocus
Exit Sub
ep:
MsgBox "发生系统错误:" & Chr(13) & Err.Description & Chr(13) & "您的系统版本可能不支持此功能或者您尚未安装这个功能.", vbCritical, "Error"
End Sub
Private Sub mnuCompmgmt_Click()
On Error GoTo ep
Shell "mmc.exe -k compmgmt.msc", vbNormalFocus
Exit Sub
ep:
MsgBox "发生系统错误:" & Chr(13) & Err.Description & Chr(13) & "您的系统版本可能不支持此功能或者您尚未安装这个功能.", vbCritical, "Error"
End Sub
Private Sub mnuControlPanel_Click()
On Error GoTo ep
Shell "control", vbNormalFocus
Exit Sub
ep:
MsgBox "发生系统错误:" & Chr(13) & Err.Description & Chr(13) & "您的系统版本可能不支持此功能或者您尚未安装这个功能.", vbCritical, "Error"
End Sub
Private Sub mnuDxDiag_Click()
On Error GoTo ep
Shell "dxdiag.exe", vbNormalFocus
Exit Sub
ep:
MsgBox "发生系统错误:" & Chr(13) & Err.Description & Chr(13) & "您的系统版本可能不支持此功能或者您尚未安装这个功能.", vbCritical, "Error"
End Sub
Private Sub mnuExplorer_Click()
On Error GoTo ep
Shell "explorer.exe", vbNormalFocus
Exit Sub
ep:
MsgBox "发生系统错误:" & Chr(13) & Err.Description & Chr(13) & "您的系统版本可能不支持此功能或者您尚未安装这个功能.", vbCritical, "Error"
End Sub
Private Sub mnuGpedit_Click()
On Error GoTo ep
Shell "mmc.exe -k gpedit.msc", vbNormalFocus
Exit Sub
ep:
MsgBox "发生系统错误:" & Chr(13) & Err.Description & Chr(13) & "您的系统版本可能不支持此功能或者您尚未安装这个功能.", vbCritical, "Error"
End Sub
Private Sub mnuMMC_Click()
On Error GoTo ep
Shell "mmc.exe", vbNormalFocus
Exit Sub
ep:
MsgBox "发生系统错误:" & Chr(13) & Err.Description & Chr(13) & "您的系统版本可能不支持此功能或者您尚未安装这个功能.", vbCritical, "Error"
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
Private Sub mnuRegedit_Click()
On Error GoTo ep
Shell "regedit.exe", vbNormalFocus
Exit Sub
ep:
MsgBox "发生系统错误:" & Chr(13) & Err.Description & Chr(13) & "您的系统版本可能不支持此功能或者您尚未安装这个功能.", vbCritical, "Error"
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
Private Sub mnuSysedit_Click()
On Error GoTo ep
Shell "sysedit.exe", vbNormalFocus
Exit Sub
ep:
MsgBox "发生系统错误:" & Chr(13) & Err.Description & Chr(13) & "您的系统版本可能不支持此功能或者您尚未安装这个功能.", vbCritical, "Error"
End Sub
Private Sub mnuSyskey_Click()
On Error GoTo ep
Shell "syskey.exe", vbNormalFocus
Exit Sub
ep:
MsgBox "发生系统错误:" & Chr(13) & Err.Description & Chr(13) & "您的系统版本可能不支持此功能或者您尚未安装这个功能.", vbCritical, "Error"
End Sub
Private Sub mnuTaskmgr_Click()
On Error GoTo ep
Shell "taskmgr.exe", vbNormalFocus
Exit Sub
ep:
MsgBox "发生系统错误:" & Chr(13) & Err.Description & Chr(13) & "您的系统版本可能不支持此功能或者您尚未安装这个功能.", vbCritical, "Error"
End Sub
Private Sub mnuWinVer_Click()
On Error GoTo ep
Shell "winver.exe", vbNormalFocus
Exit Sub
ep:
MsgBox "发生系统错误:" & Chr(13) & Err.Description & Chr(13) & "您的系统版本可能不支持此功能或者您尚未安装这个功能.", vbCritical, "Error"
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
