VERSION 5.00
Begin VB.Form frmTaskMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tasks - PC-DOS Workshop"
   ClientHeight    =   7770
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5070
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7770
   ScaleWidth      =   5070
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command2 
      Caption         =   "啟動(&S)"
      Height          =   375
      Left            =   3525
      TabIndex        =   37
      Top             =   7320
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "復位(&R)"
      Height          =   375
      Left            =   2490
      TabIndex        =   36
      Top             =   7320
      Width           =   975
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   1920
      Top             =   3450
   End
   Begin VB.CommandButton Command3 
      Caption         =   "新增更多計劃(&A)..."
      Height          =   375
      Left            =   1155
      TabIndex        =   33
      Top             =   9375
      Width           =   2310
   End
   Begin VB.Frame Frame1 
      Caption         =   "計劃與計時選項"
      Height          =   6300
      Left            =   90
      TabIndex        =   1
      Top             =   825
      Width           =   4905
      Begin VB.Frame Frame6 
         Caption         =   "運行程序選項"
         Height          =   1155
         Left            =   90
         TabIndex        =   27
         Top             =   5070
         Width           =   4710
         Begin VB.ComboBox Combo4 
            Height          =   300
            ItemData        =   "Form1.frx":068A
            Left            =   1695
            List            =   "Form1.frx":06A0
            Style           =   2  'Dropdown List
            TabIndex        =   35
            Top             =   780
            Width           =   2955
         End
         Begin VB.TextBox Text7 
            Height          =   285
            Left            =   1800
            TabIndex        =   32
            Top             =   480
            Width           =   2850
         End
         Begin VB.CommandButton Command1 
            Caption         =   "瀏覽(&B)..."
            Height          =   270
            Left            =   3450
            TabIndex        =   30
            Top             =   180
            Width           =   1200
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "程序窗口顯示方法"
            Height          =   180
            Left            =   150
            TabIndex        =   34
            Top             =   855
            Width           =   1440
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "程序執行參數(可選)"
            Height          =   180
            Left            =   135
            TabIndex        =   31
            Top             =   540
            Width           =   1620
         End
         Begin VB.Label Label11 
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   540
            TabIndex        =   29
            Top             =   195
            Width           =   2850
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "位置"
            Height          =   180
            Left            =   135
            TabIndex        =   28
            Top             =   240
            Width           =   360
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "任務"
         Height          =   3285
         Left            =   90
         TabIndex        =   14
         Top             =   1755
         Width           =   4710
         Begin VB.Frame Frame5 
            Caption         =   "用戶自定義消息選項"
            Height          =   1710
            Left            =   135
            TabIndex        =   20
            Top             =   1500
            Width           =   4455
            Begin VB.TextBox Text6 
               Height          =   840
               Left            =   510
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   26
               Top             =   795
               Width           =   2985
            End
            Begin VB.ComboBox Combo3 
               Height          =   300
               ItemData        =   "Form1.frx":06F1
               Left            =   510
               List            =   "Form1.frx":0704
               Style           =   2  'Dropdown List
               TabIndex        =   24
               Top             =   480
               Width           =   3840
            End
            Begin VB.TextBox Text5 
               Height          =   270
               Left            =   510
               TabIndex        =   22
               Top             =   195
               Width           =   3840
            End
            Begin VB.Image Image3 
               BorderStyle     =   1  'Fixed Single
               Height          =   840
               Left            =   3510
               Stretch         =   -1  'True
               Top             =   795
               Width           =   840
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "文本"
               Height          =   180
               Left            =   120
               TabIndex        =   25
               Top             =   825
               Width           =   360
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "圖標"
               Height          =   180
               Left            =   120
               TabIndex        =   23
               Top             =   555
               Width           =   360
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "標題"
               Height          =   180
               Left            =   120
               TabIndex        =   21
               Top             =   255
               Width           =   360
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "關機/重啟/註銷選項"
            Height          =   930
            Left            =   135
            TabIndex        =   17
            Top             =   525
            Width           =   4470
            Begin VB.CheckBox Check3 
               Caption         =   "強制結束未響應進程(&F)"
               Height          =   270
               Left            =   120
               TabIndex        =   19
               Top             =   555
               Width           =   2250
            End
            Begin VB.CheckBox Check2 
               Caption         =   "關機/重啟/註銷前提示(&P)"
               Height          =   270
               Left            =   120
               TabIndex        =   18
               Top             =   240
               Value           =   1  'Checked
               Width           =   2460
            End
         End
         Begin VB.ComboBox Combo2 
            Height          =   300
            ItemData        =   "Form1.frx":0730
            Left            =   1500
            List            =   "Form1.frx":0749
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   195
            Width           =   3120
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "定時器時間到時"
            Height          =   180
            Left            =   150
            TabIndex        =   15
            Top             =   270
            Width           =   1260
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "計時方式"
         Height          =   1500
         Left            =   90
         TabIndex        =   2
         Top             =   195
         Width           =   4725
         Begin VB.TextBox Text4 
            Height          =   300
            Left            =   1455
            MaxLength       =   2
            TabIndex        =   12
            Text            =   "24"
            Top             =   1110
            Width           =   270
         End
         Begin VB.TextBox Text3 
            Height          =   300
            Left            =   930
            MaxLength       =   2
            TabIndex        =   10
            Text            =   "25"
            Top             =   1110
            Width           =   270
         End
         Begin VB.TextBox Text2 
            Height          =   300
            Left            =   405
            MaxLength       =   2
            TabIndex        =   8
            Text            =   "24"
            Top             =   1110
            Width           =   270
         End
         Begin VB.OptionButton Option2 
            Caption         =   "與系統時間比對(&O)"
            Height          =   240
            Left            =   120
            TabIndex        =   7
            Top             =   840
            Width           =   4395
         End
         Begin VB.ComboBox Combo1 
            Height          =   300
            ItemData        =   "Form1.frx":07B3
            Left            =   825
            List            =   "Form1.frx":07C0
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   510
            Width           =   555
         End
         Begin VB.TextBox Text1 
            Height          =   300
            Left            =   405
            MaxLength       =   3
            TabIndex        =   4
            Text            =   "245"
            Top             =   510
            Width           =   390
         End
         Begin VB.OptionButton Option1 
            Caption         =   "倒計時(&C)"
            Height          =   240
            Left            =   120
            TabIndex        =   3
            Top             =   225
            Value           =   -1  'True
            Width           =   4395
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "秒時執行計劃的任務"
            Height          =   330
            Left            =   1770
            TabIndex        =   13
            Top             =   1170
            Width           =   2910
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "分"
            Height          =   330
            Left            =   1245
            TabIndex        =   11
            Top             =   1170
            Width           =   180
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "時"
            Height          =   330
            Left            =   720
            TabIndex        =   9
            Top             =   1170
            Width           =   180
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "之后執行計劃的任務"
            Height          =   255
            Left            =   1500
            TabIndex        =   6
            Top             =   555
            Width           =   3120
         End
      End
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   -60
      X2              =   24470
      Y1              =   7215
      Y2              =   7215
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   0
      X1              =   -60
      X2              =   24470
      Y1              =   7200
      Y2              =   7200
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   0
      Left            =   195
      Picture         =   "Form1.frx":07D0
      Top             =   -585
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   1
      Left            =   570
      Picture         =   "Form1.frx":0C12
      Top             =   -720
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   2
      Left            =   0
      Picture         =   "Form1.frx":1054
      Top             =   -750
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   3
      Left            =   870
      Picture         =   "Form1.frx":1496
      Top             =   -600
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "這個應用程序可以幫助您計劃一天的工作"
      Height          =   810
      Left            =   975
      TabIndex        =   0
      Top             =   105
      Width           =   4215
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   90
      Picture         =   "Form1.frx":18D8
      Top             =   60
      Width           =   720
   End
End
Attribute VB_Name = "frmTaskMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Type USER_DIALOG_CONFIG
lpTitle As String
lpIcon As Integer
lpMessage As String
End Type
Private Type USER_APP_RUN
lpAppPath As String
lpAppParam As String
lpRunMode As Integer
End Type
Private Type APP_TASK_PARAM
lpTimerType As Integer
lpDelay As Long
lpRunHour As Integer
lpRunMinute As Integer
lpRunSecond As Integer
lpCurrentHour As Integer
lpCurrentMinute As Integer
lpCurrentSecond As Integer
lpTaskEnum As Integer
lpTaskFriendlyDisplayName As String
lpRunning As Boolean
End Type
Dim lpDialogCfg As USER_DIALOG_CONFIG
Dim lpAppCfg As USER_APP_RUN
Dim lpTaskCfg As APP_TASK_PARAM
Const SC_SCREENSAVE = &HF140&
Dim IsCodeUse As Boolean
Private Declare Function WinExec Lib "kernel32" (ByVal lpCmdLine As String, ByVal nCmdShow As Long) As Long
Private Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
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
Private Sub Combo1_Change()
On Error Resume Next
Select Case Combo1.ListIndex
Case 0
With Text1
.MaxLength = 3
End With
If Len(Text1.Text) > 3 Then
Text1.Text = Left(Text1.Text, 3)
End If
Case 1
With Text1
.MaxLength = 3
End With
If Len(Text1.Text) > 3 Then
Text1.Text = Left(Text1.Text, 3)
End If
If CLng(Text1.Text) > 500 Then Text1.Text = 500
Case 2
With Text1
.MaxLength = 1
End With
If Len(Text1.Text) > 1 Then
Text1.Text = Left(Text1.Text, 1)
End If
End Select
End Sub
Private Sub Combo1_Click()
On Error Resume Next
Select Case Combo1.ListIndex
Case 0
With Text1
.MaxLength = 3
End With
If Len(Text1.Text) > 3 Then
Text1.Text = Left(Text1.Text, 3)
End If
Case 1
With Text1
.MaxLength = 3
End With
If Len(Text1.Text) > 3 Then
Text1.Text = Left(Text1.Text, 3)
End If
If CLng(Text1.Text) > 500 Then Text1.Text = 500
Case 2
With Text1
.MaxLength = 1
End With
If Len(Text1.Text) > 1 Then
Text1.Text = Left(Text1.Text, 1)
End If
End Select
End Sub
Private Sub Combo2_Change()
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
End Sub
Private Sub Combo2_Click()
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
End Sub
Private Sub Combo3_Change()
On Error Resume Next
Select Case Combo3.ListIndex
Case 0
With Image3
.Stretch = True
.BorderStyle = 1
.Picture = LoadPicture("")
End With
Case 1
With Image3
.Stretch = True
.BorderStyle = 1
.Picture = Me.Image2(0).Picture
End With
Case 2
With Image3
.Stretch = True
.BorderStyle = 1
.Picture = Me.Image2(1).Picture
End With
Case 3
With Image3
.Stretch = True
.BorderStyle = 1
.Picture = Me.Image2(2).Picture
End With
Case 4
With Image3
.Stretch = True
.BorderStyle = 1
.Picture = Me.Image2(3).Picture
End With
End Select
End Sub
Private Sub Combo3_Click()
On Error Resume Next
Select Case Combo3.ListIndex
Case 0
With Image3
.Stretch = True
.BorderStyle = 1
.Picture = LoadPicture("")
End With
Case 1
With Image3
.Stretch = True
.BorderStyle = 1
.Picture = Me.Image2(0).Picture
End With
Case 2
With Image3
.Stretch = True
.BorderStyle = 1
.Picture = Me.Image2(1).Picture
End With
Case 3
With Image3
.Stretch = True
.BorderStyle = 1
.Picture = Me.Image2(2).Picture
End With
Case 4
With Image3
.Stretch = True
.BorderStyle = 1
.Picture = Me.Image2(3).Picture
End With
End Select
End Sub
Private Sub Command1_Click()
On Error Resume Next
frmTaskOpenFile.Show
End Sub
Private Sub Command2_Click()
On Error Resume Next
If Option1.Value = True Then
lpTaskCfg.lpTimerType = 1
End If
If Option2.Value = True Then
lpTaskCfg.lpTimerType = 2
End If
Text1.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
Text5.Enabled = False
Text6.Enabled = False
Text7.Enabled = False
Command1.Enabled = False
Command2.Enabled = False
Command3.Enabled = False
Command4.Enabled = False
Check2.Enabled = False
Check3.Enabled = False
Combo1.Enabled = False
Combo2.Enabled = False
Combo3.Enabled = False
Option1.Enabled = False
Option2.Enabled = False
With Me.Combo4
.Enabled = False
End With
Select Case lpTaskCfg.lpTimerType
Case 1
Option1.Value = True
Case 2
Option2.Value = True
End Select
Text1.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
Text5.Enabled = False
Text6.Enabled = False
Text7.Enabled = False
Text1.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
Text5.Enabled = False
Text6.Enabled = False
Text7.Enabled = False
Combo1.Enabled = False
Combo2.Enabled = False
Combo3.Enabled = False
With Me.Combo4
.Enabled = False
End With
If lpTaskCfg.lpTimerType = 1 Then
If CInt(Text1.Text) > 500 And Combo1.ListIndex = 1 Then
MsgBox "輸入的時間超出許可范圍", vbCritical, "Error"
lpTaskCfg.lpRunning = False
If Option1.Value = True Then
lpTaskCfg.lpTimerType = 1
End If
If Option2.Value = True Then
lpTaskCfg.lpTimerType = 2
End If
lpTaskCfg.lpRunning = False
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Command1.Enabled = True
Command2.Enabled = True
Command3.Enabled = True
Command4.Enabled = True
Check2.Enabled = True
Check3.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True
Option1.Enabled = True
Option2.Enabled = True
Select Case lpTaskCfg.lpTimerType
Case 1
Option1.Value = True
Case 2
Option2.Value = True
End Select
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True
Combo4.Enabled = True
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
Exit Sub
End If
End If
If lpTaskCfg.lpTimerType = 2 Then
If CInt(Text2.Text) >= 24 Then
MsgBox "輸入的時間超出系統許可范圍", vbCritical, "Error"
lpTaskCfg.lpRunning = False
If Option1.Value = True Then
lpTaskCfg.lpTimerType = 1
End If
If Option2.Value = True Then
lpTaskCfg.lpTimerType = 2
End If
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Command1.Enabled = True
Command2.Enabled = True
Command3.Enabled = True
Command4.Enabled = True
Check2.Enabled = True
Check3.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True
Option1.Enabled = True
Option2.Enabled = True
Select Case lpTaskCfg.lpTimerType
Case 1
Option1.Value = True
Case 2
Option2.Value = True
End Select
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True
Combo4.Enabled = True
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
End With
With Me.Combo4
.Enabled = False
End With
End Select
End With
Exit Sub
End If
If CInt(Text3.Text) >= 60 Then
MsgBox "輸入的時間超出系統許可范圍", vbCritical, "Error"
lpTaskCfg.lpRunning = False
If Option1.Value = True Then
lpTaskCfg.lpTimerType = 1
End If
If Option2.Value = True Then
lpTaskCfg.lpTimerType = 2
End If
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Command1.Enabled = True
Command2.Enabled = True
Command3.Enabled = True
Command4.Enabled = True
Check2.Enabled = True
Check3.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True
Option1.Enabled = True
Option2.Enabled = True
Select Case lpTaskCfg.lpTimerType
Case 1
Option1.Value = True
Case 2
Option2.Value = True
End Select
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True
Combo4.Enabled = True
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
Exit Sub
End If
If CInt(Text4.Text) >= 60 Then
MsgBox "輸入的時間超出系統許可范圍", vbCritical, "Error"
lpTaskCfg.lpRunning = False
If Option1.Value = True Then
lpTaskCfg.lpTimerType = 1
End If
If Option2.Value = True Then
lpTaskCfg.lpTimerType = 2
End If
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Command1.Enabled = True
Command2.Enabled = True
Command3.Enabled = True
Command4.Enabled = True
Check2.Enabled = True
Check3.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True
Option1.Enabled = True
Option2.Enabled = True
Select Case lpTaskCfg.lpTimerType
Case 1
Option1.Value = True
Case 2
Option2.Value = True
End Select
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True
Combo4.Enabled = True
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
Exit Sub
End If
End If
With lpTaskCfg
.lpRunning = True
If .lpTimerType = 1 Then
Select Case Combo1.ListIndex
Case 0
.lpDelay = CLng(Text1.Text)
Case 1
.lpDelay = CLng(CLng(Text1.Text) * 60)
Case 2
.lpDelay = CLng(CLng(Text1.Text) * 3600)
End Select
End If
If .lpTimerType = 2 Then
.lpRunHour = CInt(Text2.Text)
.lpRunMinute = CInt(Text3.Text)
.lpRunSecond = CInt(Text4.Text)
End If
End With
With Timer1
.Enabled = True
.Interval = 1000
End With
If lpTaskCfg.lpTimerType = 1 Then
frmTaskDelay.Label1.Caption = lpTaskCfg.lpDelay
frmTaskDelay.Label3.Caption = "秒后" & Combo2.List(Combo2.ListIndex)
frmTaskDelay.Show
End If
If lpTaskCfg.lpTimerType = 2 Then
frmTaskTime.Label3.Caption = CInt(Text2.Text) & ":" & CInt(Text3.Text) & ":" & CInt(Text4.Text)
frmTaskTime.Label1.Caption = Hour(Now) & ":" & Minute(Now) & ":" & Second(Now)
frmTaskTime.Label5.Caption = Combo2.List(Combo2.ListIndex)
frmTaskTime.Show
End If
End Sub
Private Sub Command3_Click()
'On Error Resume Next
Dim lpThisPath As String
Dim lpThisName As String
Dim lpThisDir As String
lpThisDir = App.path
lpThisName = App.EXEName
If Right(lpThisDir, 1) = "\" Then
lpThisPath = lpThisDir & lpThisName & ".exe"
Else
lpThisPath = lpThisDir & "\" & lpThisName & ".exe"
End If
If Dir(lpThisPath, vbNormal Or vbReadOnly Or vbHidden Or vbSystem Or vbArchive) = "" Then
If Right(lpThisDir, 1) = "\" Then
lpThisPath = lpThisDir & lpThisName & ".com"
Else
lpThisPath = lpThisDir & "\" & lpThisName & ".com"
End If
End If
If Dir(lpThisPath, vbNormal Or vbReadOnly Or vbHidden Or vbSystem Or vbArchive) = "" Then
If Right(lpThisDir, 1) = "\" Then
lpThisPath = lpThisDir & lpThisName & ".bat"
Else
lpThisPath = lpThisDir & "\" & lpThisName & ".bat"
End If
End If
If Dir(lpThisPath, vbNormal Or vbReadOnly Or vbHidden Or vbSystem Or vbArchive) = "" Then
If Right(lpThisDir, 1) = "\" Then
lpThisPath = lpThisDir & lpThisName & ".cmd"
Else
lpThisPath = lpThisDir & "\" & lpThisName & ".cmd"
End If
End If
If Dir(lpThisPath, vbNormal Or vbReadOnly Or vbHidden Or vbSystem Or vbArchive) = "" Then
If Right(lpThisDir, 1) = "\" Then
lpThisPath = lpThisDir & lpThisName & ".scr"
Else
lpThisPath = lpThisDir & "\" & lpThisName & ".scr"
End If
End If
If Dir(lpThisPath, vbNormal Or vbReadOnly Or vbHidden Or vbSystem Or vbArchive) = "" Then
If Right(lpThisDir, 1) = "\" Then
lpThisPath = lpThisDir & lpThisName & ".pif"
Else
lpThisPath = lpThisDir & "\" & lpThisName & ".pif"
End If
End If
Shell lpThisPath, vbNormalFocus
End Sub
Private Sub Command4_Click()
On Error Resume Next
Dim ResetAns As Integer
ResetAns = MsgBox("確定要復位嗎?", vbQuestion + vbYesNo, "Ask")
If ResetAns = vbYes Then
Label11.Caption = ""
With Me.Combo4
.Enabled = True
.ListIndex = 0
End With
With Timer1
.Enabled = False
.Interval = 1000
End With
With Me.Option1
.Enabled = True
.Value = True
End With
With Me.Option2
.Enabled = True
.Value = False
End With
With Me.Text1
.Enabled = True
.Text = 25
.MaxLength = 3
End With
With Me.Combo1
.Enabled = True
.ListIndex = 0
End With
With Me.Text2
.MaxLength = 2
.Text = Hour(Now) + 1
.Enabled = False
End With
With Me.Text3
.Enabled = False
.Text = Minute(Now)
.MaxLength = 2
End With
With Text4
.MaxLength = 2
.Text = Second(Now)
.Enabled = False
End With
With Me.Combo2
.ListIndex = 3
.Enabled = True
End With
With Me.Text5
.Enabled = False
.Text = ""
End With
With Me.Combo3
.Enabled = False
.ListIndex = 0
End With
With Me.Text6
.Enabled = False
.Text = ""
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
.Text = ""
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
Else
Exit Sub
End If
End Sub
Private Sub TaskApp_ResetRequired()
On Error Resume Next
Label11.Caption = ""
With Me.Combo4
.Enabled = True
.ListIndex = 0
End With
With Timer1
.Enabled = False
.Interval = 1000
End With
With Me.Option1
.Enabled = True
.Value = True
End With
With Me.Option2
.Enabled = True
.Value = False
End With
With Me.Text1
.Enabled = True
.Text = 25
.MaxLength = 3
End With
With Me.Combo1
.Enabled = True
.ListIndex = 0
End With
With Me.Text2
.MaxLength = 2
.Text = Hour(Now) + 1
.Enabled = False
End With
With Me.Text3
.Enabled = False
.Text = Minute(Now)
.MaxLength = 2
End With
With Text4
.MaxLength = 2
.Text = Second(Now)
.Enabled = False
End With
With Me.Combo2
.ListIndex = 3
.Enabled = True
End With
With Me.Text5
.Enabled = False
.Text = ""
End With
With Me.Combo3
.Enabled = False
.ListIndex = 0
End With
With Me.Text6
.Enabled = False
.Text = ""
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
.Text = ""
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
End Sub
Private Sub Form_Load()
On Error Resume Next
If 25 = 245 Then
Dim hMenu As Long
hMenu = GetSystemMenu(hwnd, False)
AppendMenu hMenu, 0, 0, vbNullString
AppendMenu hMenu, 0, MENUITEM_1, "關於Tasks(&A)..."
PrevWndFunc = SetWindowLong(hwnd, GWL_WNDPROC, AddressOf SystemMenuCallback)
End If
With Command3
.Left = -245 * 25
.Top = -245 * 25
.Visible = False
End With
With Me.Combo4
.Enabled = True
.ListIndex = 0
End With
With Timer1
.Enabled = False
.Interval = 1000
End With
With Me.Option1
.Enabled = True
.Value = True
End With
With Me.Option2
.Enabled = True
.Value = False
End With
With Me.Text1
.Enabled = True
.Text = 25
.MaxLength = 3
End With
With Me.Combo1
.Enabled = True
.ListIndex = 0
End With
With Me.Text2
.MaxLength = 2
.Text = Hour(Now) + 1
.Enabled = False
End With
With Me.Text3
.Enabled = False
.Text = Minute(Now)
.MaxLength = 2
End With
With Text4
.MaxLength = 2
.Text = Second(Now)
.Enabled = False
End With
With Me.Combo2
.ListIndex = 3
.Enabled = True
End With
With Me.Text5
.Enabled = False
.Text = ""
End With
With Me.Combo3
.Enabled = False
.ListIndex = 0
End With
With Me.Text6
.Enabled = False
.Text = ""
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
If lpTaskCfg.lpRunning = True Then
Dim ans As Integer
ans = MsgBox("當前有正在運行的任務,確定退出嗎?", vbQuestion + vbYesNo, "Ask")
If ans = vbYes Then
Unload Me
Unload frmTaskOpenFile
End
Else
Cancel = 245
End If
Else
Unload Me
Unload frmTaskOpenFile
End
End If
End Sub
Private Sub Option1_Click()
On Error Resume Next
With Me.Text1
.Enabled = True
'.Text = 25
.MaxLength = 3
End With
With Me.Combo1
.Enabled = True
If 1 = 245 Then
.ListIndex = 0
End If
End With
With Me.Text2
.MaxLength = 2
.Enabled = False
End With
With Me.Text3
.Enabled = False
.MaxLength = 2
End With
With Text4
.MaxLength = 2
.Enabled = False
End With
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
End With
With Me.Combo4
.Enabled = False
End With
End Select
End With
Text1.SetFocus
End Sub
Private Sub Option2_Click()
On Error Resume Next
With Me.Text1
.Enabled = False
'.Text = 25
.MaxLength = 3
End With
With Me.Combo1
.Enabled = False
If 1 = 245 Then
.ListIndex = 0
End If
End With
With Me.Text2
.MaxLength = 2
.Enabled = True
End With
With Me.Text3
.Enabled = True
.MaxLength = 2
End With
With Text4
.MaxLength = 2
.Enabled = True
End With
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
End With
With Me.Combo4
.Enabled = False
End With
End Select
End With
Text2.SetFocus
End Sub
Private Sub Text1_Change()
On Error Resume Next
End Sub
Private Sub Text1_GotFocus()
On Error Resume Next
Text1.SelStart = 0
Text1.SelLength = Len(Text1.Text)
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii <> vbKeyBack Then
If KeyAscii <> vbKey0 Then
If KeyAscii <> vbKey1 Then
If KeyAscii <> vbKey2 Then
If KeyAscii <> vbKey3 Then
If KeyAscii <> vbKey4 Then
If KeyAscii <> vbKey5 Then
If KeyAscii <> vbKey6 Then
If KeyAscii <> vbKey7 Then
If KeyAscii <> vbKey8 Then
If KeyAscii <> vbKey9 Then
KeyAscii = 0
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End Sub
Private Sub Text2_GotFocus()
On Error Resume Next
Text2.SelStart = 0
Text2.SelLength = Len(Text2.Text)
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii <> vbKeyBack Then
If KeyAscii <> vbKey0 Then
If KeyAscii <> vbKey1 Then
If KeyAscii <> vbKey2 Then
If KeyAscii <> vbKey3 Then
If KeyAscii <> vbKey4 Then
If KeyAscii <> vbKey5 Then
If KeyAscii <> vbKey6 Then
If KeyAscii <> vbKey7 Then
If KeyAscii <> vbKey8 Then
If KeyAscii <> vbKey9 Then
KeyAscii = 0
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End Sub
Private Sub Text3_GotFocus()
On Error Resume Next
Text3.SelStart = 0
Text3.SelLength = Len(Text3.Text)
End Sub
Private Sub Text3_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii <> vbKeyBack Then
If KeyAscii <> vbKey0 Then
If KeyAscii <> vbKey1 Then
If KeyAscii <> vbKey2 Then
If KeyAscii <> vbKey3 Then
If KeyAscii <> vbKey4 Then
If KeyAscii <> vbKey5 Then
If KeyAscii <> vbKey6 Then
If KeyAscii <> vbKey7 Then
If KeyAscii <> vbKey8 Then
If KeyAscii <> vbKey9 Then
KeyAscii = 0
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End Sub
Private Sub Text4_GotFocus()
On Error Resume Next
Text4.SelStart = 0
Text4.SelLength = Len(Text4.Text)
End Sub
Private Sub Text4_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii <> vbKeyBack Then
If KeyAscii <> vbKey0 Then
If KeyAscii <> vbKey1 Then
If KeyAscii <> vbKey2 Then
If KeyAscii <> vbKey3 Then
If KeyAscii <> vbKey4 Then
If KeyAscii <> vbKey5 Then
If KeyAscii <> vbKey6 Then
If KeyAscii <> vbKey7 Then
If KeyAscii <> vbKey8 Then
If KeyAscii <> vbKey9 Then
KeyAscii = 0
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End Sub
Private Sub Timer1_Timer()
On Error Resume Next
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
Dim ans As Integer
Dim ResetAns As Integer
Select Case lpTaskCfg.lpTimerType
Case 1
lpTaskCfg.lpDelay = lpTaskCfg.lpDelay - 1
If lpTaskCfg.lpDelay <= 0 Then lpTaskCfg.lpDelay = 0
frmTaskDelay.Label1.Caption = CStr(lpTaskCfg.lpDelay)
If lpTaskCfg.lpDelay <= 0 Then
lpTaskCfg.lpDelay = 0
Timer1.Enabled = False
Select Case lpTaskCfg.lpTaskFriendlyDisplayName
Case "Shutdown"
If Check2.Value = 1 Then
ans = MsgBox("確定要關機嗎?請注意保存數據!", vbExclamation + vbYesNo, "Ask")
If ans = vbYes Then
If Check3.Value = 1 Then
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
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Command1.Enabled = True
Command2.Enabled = True
Command3.Enabled = True
Check2.Enabled = True
Check3.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True
Option1.Enabled = True
Option2.Enabled = True
Unload frmTaskDelay
lpTaskCfg.lpRunning = False
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
End With
End Select
End With
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
End With
With Me.Combo4
.Enabled = False
End With
End Select
End With
Command4.Enabled = True
ResetAns = MsgBox("任務運行完畢,要復位任務設定嗎?", vbQuestion + vbYesNo, "Ask")
If ResetAns = vbYes Then
TaskApp_ResetRequired
On Error Resume Next
Label11.Caption = ""
With Me.Combo4
.Enabled = True
.ListIndex = 0
End With
With Timer1
.Enabled = False
.Interval = 1000
End With
With Me.Option1
.Enabled = True
.Value = True
End With
With Me.Option2
.Enabled = True
.Value = False
End With
With Me.Text1
.Enabled = True
.Text = 25
.MaxLength = 3
End With
With Me.Combo1
.Enabled = True
.ListIndex = 0
End With
With Me.Text2
.MaxLength = 2
.Text = Hour(Now) + 1
.Enabled = False
End With
With Me.Text3
.Enabled = False
.Text = Minute(Now)
.MaxLength = 2
End With
With Text4
.MaxLength = 2
.Text = Second(Now)
.Enabled = False
End With
With Me.Combo2
.ListIndex = 3
.Enabled = True
End With
With Me.Text5
.Enabled = False
.Text = ""
End With
With Me.Combo3
.Enabled = False
.ListIndex = 0
End With
With Me.Text6
.Enabled = False
.Text = ""
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
.Text = ""
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
TaskApp_ResetRequired
On Error Resume Next
Label11.Caption = ""
With Me.Combo4
.Enabled = True
.ListIndex = 0
End With
With Timer1
.Enabled = False
.Interval = 1000
End With
With Me.Option1
.Enabled = True
.Value = True
End With
With Me.Option2
.Enabled = True
.Value = False
End With
With Me.Text1
.Enabled = True
.Text = 25
.MaxLength = 3
End With
With Me.Combo1
.Enabled = True
.ListIndex = 0
End With
With Me.Text2
.MaxLength = 2
.Text = Hour(Now) + 1
.Enabled = False
End With
With Me.Text3
.Enabled = False
.Text = Minute(Now)
.MaxLength = 2
End With
With Text4
.MaxLength = 2
.Text = Second(Now)
.Enabled = False
End With
With Me.Combo2
.ListIndex = 3
.Enabled = True
End With
With Me.Text5
.Enabled = False
.Text = ""
End With
With Me.Combo3
.Enabled = False
.ListIndex = 0
End With
With Me.Text6
.Enabled = False
.Text = ""
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
.Text = ""
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
With Me.Combo4
.ListIndex = 0
End With
Label11.Caption = ""
With Timer1
.Enabled = False
.Interval = 1000
End With
With Me.Option1
.Enabled = True
.Value = True
End With
With Me.Option2
.Enabled = True
.Value = False
End With
With Me.Text1
.Enabled = True
.Text = 25
.MaxLength = 3
End With
With Me.Combo1
.Enabled = True
.ListIndex = 0
End With
With Me.Text2
.MaxLength = 2
.Text = Hour(Now) + 1
.Enabled = False
End With
With Me.Text3
.Enabled = False
.Text = Minute(Now)
.MaxLength = 2
End With
With Text4
.MaxLength = 2
.Text = Second(Now)
.Enabled = False
End With
With Me.Combo2
.ListIndex = 3
.Enabled = True
End With
With Me.Text5
.Enabled = False
.Text = ""
End With
With Me.Combo3
.Enabled = False
.ListIndex = 0
End With
With Me.Text6
.Enabled = False
.Text = ""
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
End With
With Me.Combo4
.Enabled = False
End With
End Select
End With
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
End With
End Select
End With
Else
Exit Sub
End If
Exit Sub
Select Case lpTaskCfg.lpTimerType
Case 1
Option1.Value = True
Case 2
Option2.Value = True
End Select
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True
Unload frmTaskDelay
lpTaskCfg.lpRunning = False
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
End With
With Me.Combo4
.Enabled = False
End With
End Select
End With
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
Exit Sub
Else
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
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Command1.Enabled = True
Command2.Enabled = True
Command3.Enabled = True
Check2.Enabled = True
Check3.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True
Option1.Enabled = True
Option2.Enabled = True
Select Case lpTaskCfg.lpTimerType
Case 1
Option1.Value = True
Case 2
Option2.Value = True
End Select
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True
Unload frmTaskDelay
lpTaskCfg.lpRunning = False
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
Command4.Enabled = True
ResetAns = MsgBox("任務運行完畢,要復位任務設定嗎?", vbQuestion + vbYesNo, "Ask")
If ResetAns = vbYes Then
TaskApp_ResetRequired
On Error Resume Next
Label11.Caption = ""
With Me.Combo4
.Enabled = True
.ListIndex = 0
End With
With Timer1
.Enabled = False
.Interval = 1000
End With
With Me.Option1
.Enabled = True
.Value = True
End With
With Me.Option2
.Enabled = True
.Value = False
End With
With Me.Text1
.Enabled = True
.Text = 25
.MaxLength = 3
End With
With Me.Combo1
.Enabled = True
.ListIndex = 0
End With
With Me.Text2
.MaxLength = 2
.Text = Hour(Now) + 1
.Enabled = False
End With
With Me.Text3
.Enabled = False
.Text = Minute(Now)
.MaxLength = 2
End With
With Text4
.MaxLength = 2
.Text = Second(Now)
.Enabled = False
End With
With Me.Combo2
.ListIndex = 3
.Enabled = True
End With
With Me.Text5
.Enabled = False
.Text = ""
End With
With Me.Combo3
.Enabled = False
.ListIndex = 0
End With
With Me.Text6
.Enabled = False
.Text = ""
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
.Text = ""
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
With Me.Combo4
.ListIndex = 0
End With
Label11.Caption = ""
With Timer1
.Enabled = False
.Interval = 1000
End With
With Me.Option1
.Enabled = True
.Value = True
End With
With Me.Option2
.Enabled = True
.Value = False
End With
With Me.Text1
.Enabled = True
.Text = 25
.MaxLength = 3
End With
With Me.Combo1
.Enabled = True
.ListIndex = 0
End With
With Me.Text2
.MaxLength = 2
.Text = Hour(Now) + 1
.Enabled = False
End With
With Me.Text3
.Enabled = False
.Text = Minute(Now)
.MaxLength = 2
End With
With Text4
.MaxLength = 2
.Text = Second(Now)
.Enabled = False
End With
With Me.Combo2
.ListIndex = 3
.Enabled = True
End With
With Me.Text5
.Enabled = False
.Text = ""
End With
With Me.Combo3
.Enabled = False
.ListIndex = 0
End With
With Me.Text6
.Enabled = False
.Text = ""
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
Else
Exit Sub
End If
Exit Sub
Select Case lpTaskCfg.lpTimerType
Case 1
Option1.Value = True
Case 2
Option2.Value = True
End Select
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True
Unload frmTaskDelay
lpTaskCfg.lpRunning = False
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
Exit Sub
End If
Else
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Command1.Enabled = True
Command2.Enabled = True
Command3.Enabled = True
Check2.Enabled = True
Check3.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True
Option1.Enabled = True
Option2.Enabled = True
Select Case lpTaskCfg.lpTimerType
Case 1
Option1.Value = True
Case 2
Option2.Value = True
End Select
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True
Unload frmTaskDelay
lpTaskCfg.lpRunning = False
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
Command4.Enabled = True
ResetAns = MsgBox("任務運行完畢,要復位任務設定嗎?", vbQuestion + vbYesNo, "Ask")
If ResetAns = vbYes Then
TaskApp_ResetRequired
On Error Resume Next
Label11.Caption = ""
With Me.Combo4
.Enabled = True
.ListIndex = 0
End With
With Timer1
.Enabled = False
.Interval = 1000
End With
With Me.Option1
.Enabled = True
.Value = True
End With
With Me.Option2
.Enabled = True
.Value = False
End With
With Me.Text1
.Enabled = True
.Text = 25
.MaxLength = 3
End With
With Me.Combo1
.Enabled = True
.ListIndex = 0
End With
With Me.Text2
.MaxLength = 2
.Text = Hour(Now) + 1
.Enabled = False
End With
With Me.Text3
.Enabled = False
.Text = Minute(Now)
.MaxLength = 2
End With
With Text4
.MaxLength = 2
.Text = Second(Now)
.Enabled = False
End With
With Me.Combo2
.ListIndex = 3
.Enabled = True
End With
With Me.Text5
.Enabled = False
.Text = ""
End With
With Me.Combo3
.Enabled = False
.ListIndex = 0
End With
With Me.Text6
.Enabled = False
.Text = ""
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
.Text = ""
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
With Me.Combo4
.ListIndex = 0
End With
Label11.Caption = ""
With Timer1
.Enabled = False
.Interval = 1000
End With
With Me.Option1
.Enabled = True
.Value = True
End With
With Me.Option2
.Enabled = True
.Value = False
End With
With Me.Text1
.Enabled = True
.Text = 25
.MaxLength = 3
End With
With Me.Combo1
.Enabled = True
.ListIndex = 0
End With
With Me.Text2
.MaxLength = 2
.Text = Hour(Now) + 1
.Enabled = False
End With
With Me.Text3
.Enabled = False
.Text = Minute(Now)
.MaxLength = 2
End With
With Text4
.MaxLength = 2
.Text = Second(Now)
.Enabled = False
End With
With Me.Combo2
.ListIndex = 3
.Enabled = True
End With
With Me.Text5
.Enabled = False
.Text = ""
End With
With Me.Combo3
.Enabled = False
.ListIndex = 0
End With
With Me.Text6
.Enabled = False
.Text = ""
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
Else
Exit Sub
End If
Exit Sub
Select Case lpTaskCfg.lpTimerType
Case 1
Option1.Value = True
Case 2
Option2.Value = True
End Select
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True
Unload frmTaskDelay
lpTaskCfg.lpRunning = False
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
Exit Sub
Exit Sub
End If
Else
If Check3.Value = 1 Then
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
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Command1.Enabled = True
Command2.Enabled = True
Command3.Enabled = True
Check2.Enabled = True
Check3.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True
Option1.Enabled = True
Option2.Enabled = True
Select Case lpTaskCfg.lpTimerType
Case 1
Option1.Value = True
Case 2
Option2.Value = True
End Select
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True
Unload frmTaskDelay
lpTaskCfg.lpRunning = False
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
Command4.Enabled = True
ResetAns = MsgBox("任務運行完畢,要復位任務設定嗎?", vbQuestion + vbYesNo, "Ask")
If ResetAns = vbYes Then
TaskApp_ResetRequired
On Error Resume Next
Label11.Caption = ""
With Me.Combo4
.Enabled = True
.ListIndex = 0
End With
With Timer1
.Enabled = False
.Interval = 1000
End With
With Me.Option1
.Enabled = True
.Value = True
End With
With Me.Option2
.Enabled = True
.Value = False
End With
With Me.Text1
.Enabled = True
.Text = 25
.MaxLength = 3
End With
With Me.Combo1
.Enabled = True
.ListIndex = 0
End With
With Me.Text2
.MaxLength = 2
.Text = Hour(Now) + 1
.Enabled = False
End With
With Me.Text3
.Enabled = False
.Text = Minute(Now)
.MaxLength = 2
End With
With Text4
.MaxLength = 2
.Text = Second(Now)
.Enabled = False
End With
With Me.Combo2
.ListIndex = 3
.Enabled = True
End With
With Me.Text5
.Enabled = False
.Text = ""
End With
With Me.Combo3
.Enabled = False
.ListIndex = 0
End With
With Me.Text6
.Enabled = False
.Text = ""
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
.Text = ""
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
With Me.Combo4
.ListIndex = 0
End With
Label11.Caption = ""
With Timer1
.Enabled = False
.Interval = 1000
End With
With Me.Option1
.Enabled = True
.Value = True
End With
With Me.Option2
.Enabled = True
.Value = False
End With
With Me.Text1
.Enabled = True
.Text = 25
.MaxLength = 3
End With
With Me.Combo1
.Enabled = True
.ListIndex = 0
End With
With Me.Text2
.MaxLength = 2
.Text = Hour(Now) + 1
.Enabled = False
End With
With Me.Text3
.Enabled = False
.Text = Minute(Now)
.MaxLength = 2
End With
With Text4
.MaxLength = 2
.Text = Second(Now)
.Enabled = False
End With
With Me.Combo2
.ListIndex = 3
.Enabled = True
End With
With Me.Text5
.Enabled = False
.Text = ""
End With
With Me.Combo3
.Enabled = False
.ListIndex = 0
End With
With Me.Text6
.Enabled = False
.Text = ""
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
Else
Exit Sub
End If
Exit Sub
Select Case lpTaskCfg.lpTimerType
Case 1
Option1.Value = True
Case 2
Option2.Value = True
End Select
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True
Unload frmTaskDelay
lpTaskCfg.lpRunning = False
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
Exit Sub
End If
End If
Case "Reboot"
If Check2.Value = 1 Then
ans = MsgBox("確定要重啟嗎?請注意保存數據!", vbExclamation + vbYesNo, "Ask")
If ans = vbYes Then
If Check3.Value = 1 Then
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
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Command1.Enabled = True
Command2.Enabled = True
Command3.Enabled = True
Check2.Enabled = True
Check3.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True
Option1.Enabled = True
Option2.Enabled = True
Select Case lpTaskCfg.lpTimerType
Case 1
Option1.Value = True
Case 2
Option2.Value = True
End Select
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True
Unload frmTaskDelay
lpTaskCfg.lpRunning = False
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
Command4.Enabled = True
ResetAns = MsgBox("任務運行完畢,要復位任務設定嗎?", vbQuestion + vbYesNo, "Ask")
If ResetAns = vbYes Then
TaskApp_ResetRequired
On Error Resume Next
Label11.Caption = ""
With Me.Combo4
.Enabled = True
.ListIndex = 0
End With
With Timer1
.Enabled = False
.Interval = 1000
End With
With Me.Option1
.Enabled = True
.Value = True
End With
With Me.Option2
.Enabled = True
.Value = False
End With
With Me.Text1
.Enabled = True
.Text = 25
.MaxLength = 3
End With
With Me.Combo1
.Enabled = True
.ListIndex = 0
End With
With Me.Text2
.MaxLength = 2
.Text = Hour(Now) + 1
.Enabled = False
End With
With Me.Text3
.Enabled = False
.Text = Minute(Now)
.MaxLength = 2
End With
With Text4
.MaxLength = 2
.Text = Second(Now)
.Enabled = False
End With
With Me.Combo2
.ListIndex = 3
.Enabled = True
End With
With Me.Text5
.Enabled = False
.Text = ""
End With
With Me.Combo3
.Enabled = False
.ListIndex = 0
End With
With Me.Text6
.Enabled = False
.Text = ""
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
.Text = ""
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
With Me.Combo4
.ListIndex = 0
End With
Label11.Caption = ""
With Timer1
.Enabled = False
.Interval = 1000
End With
With Me.Option1
.Enabled = True
.Value = True
End With
With Me.Option2
.Enabled = True
.Value = False
End With
With Me.Text1
.Enabled = True
.Text = 25
.MaxLength = 3
End With
With Me.Combo1
.Enabled = True
.ListIndex = 0
End With
With Me.Text2
.MaxLength = 2
.Text = Hour(Now) + 1
.Enabled = False
End With
With Me.Text3
.Enabled = False
.Text = Minute(Now)
.MaxLength = 2
End With
With Text4
.MaxLength = 2
.Text = Second(Now)
.Enabled = False
End With
With Me.Combo2
.ListIndex = 3
.Enabled = True
End With
With Me.Text5
.Enabled = False
.Text = ""
End With
With Me.Combo3
.Enabled = False
.ListIndex = 0
End With
With Me.Text6
.Enabled = False
.Text = ""
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
Else
Exit Sub
End If
Exit Sub
Select Case lpTaskCfg.lpTimerType
Case 1
Option1.Value = True
Case 2
Option2.Value = True
End Select
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True
Unload frmTaskDelay
lpTaskCfg.lpRunning = False
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
Exit Sub
Else
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
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Command1.Enabled = True
Command2.Enabled = True
Command3.Enabled = True
Check2.Enabled = True
Check3.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True
Option1.Enabled = True
Option2.Enabled = True
Select Case lpTaskCfg.lpTimerType
Case 1
Option1.Value = True
Case 2
Option2.Value = True
End Select
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True
Unload frmTaskDelay
lpTaskCfg.lpRunning = False
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
Command4.Enabled = True
ResetAns = MsgBox("任務運行完畢,要復位任務設定嗎?", vbQuestion + vbYesNo, "Ask")
If ResetAns = vbYes Then
TaskApp_ResetRequired
On Error Resume Next
Label11.Caption = ""
With Me.Combo4
.Enabled = True
.ListIndex = 0
End With
With Timer1
.Enabled = False
.Interval = 1000
End With
With Me.Option1
.Enabled = True
.Value = True
End With
With Me.Option2
.Enabled = True
.Value = False
End With
With Me.Text1
.Enabled = True
.Text = 25
.MaxLength = 3
End With
With Me.Combo1
.Enabled = True
.ListIndex = 0
End With
With Me.Text2
.MaxLength = 2
.Text = Hour(Now) + 1
.Enabled = False
End With
With Me.Text3
.Enabled = False
.Text = Minute(Now)
.MaxLength = 2
End With
With Text4
.MaxLength = 2
.Text = Second(Now)
.Enabled = False
End With
With Me.Combo2
.ListIndex = 3
.Enabled = True
End With
With Me.Text5
.Enabled = False
.Text = ""
End With
With Me.Combo3
.Enabled = False
.ListIndex = 0
End With
With Me.Text6
.Enabled = False
.Text = ""
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
.Text = ""
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
With Me.Combo4
.ListIndex = 0
End With
Label11.Caption = ""
With Timer1
.Enabled = False
.Interval = 1000
End With
With Me.Option1
.Enabled = True
.Value = True
End With
With Me.Option2
.Enabled = True
.Value = False
End With
With Me.Text1
.Enabled = True
.Text = 25
.MaxLength = 3
End With
With Me.Combo1
.Enabled = True
.ListIndex = 0
End With
With Me.Text2
.MaxLength = 2
.Text = Hour(Now) + 1
.Enabled = False
End With
With Me.Text3
.Enabled = False
.Text = Minute(Now)
.MaxLength = 2
End With
With Text4
.MaxLength = 2
.Text = Second(Now)
.Enabled = False
End With
With Me.Combo2
.ListIndex = 3
.Enabled = True
End With
With Me.Text5
.Enabled = False
.Text = ""
End With
With Me.Combo3
.Enabled = False
.ListIndex = 0
End With
With Me.Text6
.Enabled = False
.Text = ""
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
Else
Exit Sub
End If
Exit Sub
Select Case lpTaskCfg.lpTimerType
Case 1
Option1.Value = True
Case 2
Option2.Value = True
End Select
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True
Unload frmTaskDelay
lpTaskCfg.lpRunning = False
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
Exit Sub
End If
Else
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Command1.Enabled = True
Command2.Enabled = True
Command3.Enabled = True
Check2.Enabled = True
Check3.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True
Option1.Enabled = True
Option2.Enabled = True
Select Case lpTaskCfg.lpTimerType
Case 1
Option1.Value = True
Case 2
Option2.Value = True
End Select
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True
Unload frmTaskDelay
lpTaskCfg.lpRunning = False
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
Command4.Enabled = True
ResetAns = MsgBox("任務運行完畢,要復位任務設定嗎?", vbQuestion + vbYesNo, "Ask")
If ResetAns = vbYes Then
TaskApp_ResetRequired
On Error Resume Next
Label11.Caption = ""
With Me.Combo4
.Enabled = True
.ListIndex = 0
End With
With Timer1
.Enabled = False
.Interval = 1000
End With
With Me.Option1
.Enabled = True
.Value = True
End With
With Me.Option2
.Enabled = True
.Value = False
End With
With Me.Text1
.Enabled = True
.Text = 25
.MaxLength = 3
End With
With Me.Combo1
.Enabled = True
.ListIndex = 0
End With
With Me.Text2
.MaxLength = 2
.Text = Hour(Now) + 1
.Enabled = False
End With
With Me.Text3
.Enabled = False
.Text = Minute(Now)
.MaxLength = 2
End With
With Text4
.MaxLength = 2
.Text = Second(Now)
.Enabled = False
End With
With Me.Combo2
.ListIndex = 3
.Enabled = True
End With
With Me.Text5
.Enabled = False
.Text = ""
End With
With Me.Combo3
.Enabled = False
.ListIndex = 0
End With
With Me.Text6
.Enabled = False
.Text = ""
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
.Text = ""
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
With Me.Combo4
.ListIndex = 0
End With
Label11.Caption = ""
With Timer1
.Enabled = False
.Interval = 1000
End With
With Me.Option1
.Enabled = True
.Value = True
End With
With Me.Option2
.Enabled = True
.Value = False
End With
With Me.Text1
.Enabled = True
.Text = 25
.MaxLength = 3
End With
With Me.Combo1
.Enabled = True
.ListIndex = 0
End With
With Me.Text2
.MaxLength = 2
.Text = Hour(Now) + 1
.Enabled = False
End With
With Me.Text3
.Enabled = False
.Text = Minute(Now)
.MaxLength = 2
End With
With Text4
.MaxLength = 2
.Text = Second(Now)
.Enabled = False
End With
With Me.Combo2
.ListIndex = 3
.Enabled = True
End With
With Me.Text5
.Enabled = False
.Text = ""
End With
With Me.Combo3
.Enabled = False
.ListIndex = 0
End With
With Me.Text6
.Enabled = False
.Text = ""
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
Else
Exit Sub
End If
Exit Sub
Select Case lpTaskCfg.lpTimerType
Case 1
Option1.Value = True
Case 2
Option2.Value = True
End Select
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True
Unload frmTaskDelay
lpTaskCfg.lpRunning = False
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
Exit Sub
Exit Sub
End If
Else
If Check3.Value = 1 Then
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
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Command1.Enabled = True
Command2.Enabled = True
Command3.Enabled = True
Check2.Enabled = True
Check3.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True
Option1.Enabled = True
Option2.Enabled = True
Select Case lpTaskCfg.lpTimerType
Case 1
Option1.Value = True
Case 2
Option2.Value = True
End Select
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True
Unload frmTaskDelay
lpTaskCfg.lpRunning = False
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
Command4.Enabled = True
ResetAns = MsgBox("任務運行完畢,要復位任務設定嗎?", vbQuestion + vbYesNo, "Ask")
If ResetAns = vbYes Then
TaskApp_ResetRequired
On Error Resume Next
Label11.Caption = ""
With Me.Combo4
.Enabled = True
.ListIndex = 0
End With
With Timer1
.Enabled = False
.Interval = 1000
End With
With Me.Option1
.Enabled = True
.Value = True
End With
With Me.Option2
.Enabled = True
.Value = False
End With
With Me.Text1
.Enabled = True
.Text = 25
.MaxLength = 3
End With
With Me.Combo1
.Enabled = True
.ListIndex = 0
End With
With Me.Text2
.MaxLength = 2
.Text = Hour(Now) + 1
.Enabled = False
End With
With Me.Text3
.Enabled = False
.Text = Minute(Now)
.MaxLength = 2
End With
With Text4
.MaxLength = 2
.Text = Second(Now)
.Enabled = False
End With
With Me.Combo2
.ListIndex = 3
.Enabled = True
End With
With Me.Text5
.Enabled = False
.Text = ""
End With
With Me.Combo3
.Enabled = False
.ListIndex = 0
End With
With Me.Text6
.Enabled = False
.Text = ""
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
.Text = ""
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
With Me.Combo4
.ListIndex = 0
End With
Label11.Caption = ""
With Timer1
.Enabled = False
.Interval = 1000
End With
With Me.Option1
.Enabled = True
.Value = True
End With
With Me.Option2
.Enabled = True
.Value = False
End With
With Me.Text1
.Enabled = True
.Text = 25
.MaxLength = 3
End With
With Me.Combo1
.Enabled = True
.ListIndex = 0
End With
With Me.Text2
.MaxLength = 2
.Text = Hour(Now) + 1
.Enabled = False
End With
With Me.Text3
.Enabled = False
.Text = Minute(Now)
.MaxLength = 2
End With
With Text4
.MaxLength = 2
.Text = Second(Now)
.Enabled = False
End With
With Me.Combo2
.ListIndex = 3
.Enabled = True
End With
With Me.Text5
.Enabled = False
.Text = ""
End With
With Me.Combo3
.Enabled = False
.ListIndex = 0
End With
With Me.Text6
.Enabled = False
.Text = ""
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
Else
Exit Sub
End If
Exit Sub
Select Case lpTaskCfg.lpTimerType
Case 1
Option1.Value = True
Case 2
Option2.Value = True
End Select
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True
Unload frmTaskDelay
lpTaskCfg.lpRunning = False
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
Exit Sub
Else
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
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Command1.Enabled = True
Command2.Enabled = True
Command3.Enabled = True
Check2.Enabled = True
Check3.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True
Option1.Enabled = True
Option2.Enabled = True
Select Case lpTaskCfg.lpTimerType
Case 1
Option1.Value = True
Case 2
Option2.Value = True
End Select
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True
Unload frmTaskDelay
lpTaskCfg.lpRunning = False
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
End With
With Me.Combo4
.Enabled = False
End With
End Select
End With
Command4.Enabled = True
ResetAns = MsgBox("任務運行完畢,要復位任務設定嗎?", vbQuestion + vbYesNo, "Ask")
If ResetAns = vbYes Then
TaskApp_ResetRequired
On Error Resume Next
Label11.Caption = ""
With Me.Combo4
.Enabled = True
.ListIndex = 0
End With
With Timer1
.Enabled = False
.Interval = 1000
End With
With Me.Option1
.Enabled = True
.Value = True
End With
With Me.Option2
.Enabled = True
.Value = False
End With
With Me.Text1
.Enabled = True
.Text = 25
.MaxLength = 3
End With
With Me.Combo1
.Enabled = True
.ListIndex = 0
End With
With Me.Text2
.MaxLength = 2
.Text = Hour(Now) + 1
.Enabled = False
End With
With Me.Text3
.Enabled = False
.Text = Minute(Now)
.MaxLength = 2
End With
With Text4
.MaxLength = 2
.Text = Second(Now)
.Enabled = False
End With
With Me.Combo2
.ListIndex = 3
.Enabled = True
End With
With Me.Text5
.Enabled = False
.Text = ""
End With
With Me.Combo3
.Enabled = False
.ListIndex = 0
End With
With Me.Text6
.Enabled = False
.Text = ""
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
.Text = ""
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
With Me.Combo4
.ListIndex = 0
End With
Label11.Caption = ""
With Timer1
.Enabled = False
.Interval = 1000
End With
With Me.Option1
.Enabled = True
.Value = True
End With
With Me.Option2
.Enabled = True
.Value = False
End With
With Me.Text1
.Enabled = True
.Text = 25
.MaxLength = 3
End With
With Me.Combo1
.Enabled = True
.ListIndex = 0
End With
With Me.Text2
.MaxLength = 2
.Text = Hour(Now) + 1
.Enabled = False
End With
With Me.Text3
.Enabled = False
.Text = Minute(Now)
.MaxLength = 2
End With
With Text4
.MaxLength = 2
.Text = Second(Now)
.Enabled = False
End With
With Me.Combo2
.ListIndex = 3
.Enabled = True
End With
With Me.Text5
.Enabled = False
.Text = ""
End With
With Me.Combo3
.Enabled = False
.ListIndex = 0
End With
With Me.Text6
.Enabled = False
.Text = ""
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
Else
Exit Sub
End If
Exit Sub
Select Case lpTaskCfg.lpTimerType
Case 1
Option1.Value = True
Case 2
Option2.Value = True
End Select
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True
Unload frmTaskDelay
lpTaskCfg.lpRunning = False
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
Exit Sub
End If
End If
Case "LogOff"
If Check2.Value = 1 Then
ans = MsgBox("確定要注銷嗎?請注意保存數據!", vbExclamation + vbYesNo, "Ask")
If ans = vbYes Then
If Check3.Value = 1 Then
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
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Command1.Enabled = True
Command2.Enabled = True
Command3.Enabled = True
Check2.Enabled = True
Check3.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True
Option1.Enabled = True
Option2.Enabled = True
Select Case lpTaskCfg.lpTimerType
Case 1
Option1.Value = True
Case 2
Option2.Value = True
End Select
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True
Unload frmTaskDelay
lpTaskCfg.lpRunning = False
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
Command4.Enabled = True
ResetAns = MsgBox("任務運行完畢,要復位任務設定嗎?", vbQuestion + vbYesNo, "Ask")
If ResetAns = vbYes Then
TaskApp_ResetRequired
On Error Resume Next
Label11.Caption = ""
With Me.Combo4
.Enabled = True
.ListIndex = 0
End With
With Timer1
.Enabled = False
.Interval = 1000
End With
With Me.Option1
.Enabled = True
.Value = True
End With
With Me.Option2
.Enabled = True
.Value = False
End With
With Me.Text1
.Enabled = True
.Text = 25
.MaxLength = 3
End With
With Me.Combo1
.Enabled = True
.ListIndex = 0
End With
With Me.Text2
.MaxLength = 2
.Text = Hour(Now) + 1
.Enabled = False
End With
With Me.Text3
.Enabled = False
.Text = Minute(Now)
.MaxLength = 2
End With
With Text4
.MaxLength = 2
.Text = Second(Now)
.Enabled = False
End With
With Me.Combo2
.ListIndex = 3
.Enabled = True
End With
With Me.Text5
.Enabled = False
.Text = ""
End With
With Me.Combo3
.Enabled = False
.ListIndex = 0
End With
With Me.Text6
.Enabled = False
.Text = ""
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
.Text = ""
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
With Me.Combo4
.ListIndex = 0
End With
Label11.Caption = ""
With Timer1
.Enabled = False
.Interval = 1000
End With
With Me.Option1
.Enabled = True
.Value = True
End With
With Me.Option2
.Enabled = True
.Value = False
End With
With Me.Text1
.Enabled = True
.Text = 25
.MaxLength = 3
End With
With Me.Combo1
.Enabled = True
.ListIndex = 0
End With
With Me.Text2
.MaxLength = 2
.Text = Hour(Now) + 1
.Enabled = False
End With
With Me.Text3
.Enabled = False
.Text = Minute(Now)
.MaxLength = 2
End With
With Text4
.MaxLength = 2
.Text = Second(Now)
.Enabled = False
End With
With Me.Combo2
.ListIndex = 3
.Enabled = True
End With
With Me.Text5
.Enabled = False
.Text = ""
End With
With Me.Combo3
.Enabled = False
.ListIndex = 0
End With
With Me.Text6
.Enabled = False
.Text = ""
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
Else
Exit Sub
End If
Exit Sub
Select Case lpTaskCfg.lpTimerType
Case 1
Option1.Value = True
Case 2
Option2.Value = True
End Select
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True
Unload frmTaskDelay
lpTaskCfg.lpRunning = False
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
Exit Sub
Else
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
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Command1.Enabled = True
Command2.Enabled = True
Command3.Enabled = True
Check2.Enabled = True
Check3.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True
Option1.Enabled = True
Option2.Enabled = True
Select Case lpTaskCfg.lpTimerType
Case 1
Option1.Value = True
Case 2
Option2.Value = True
End Select
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True
Unload frmTaskDelay
lpTaskCfg.lpRunning = False
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
Command4.Enabled = True
ResetAns = MsgBox("任務運行完畢,要復位任務設定嗎?", vbQuestion + vbYesNo, "Ask")
If ResetAns = vbYes Then
TaskApp_ResetRequired
On Error Resume Next
Label11.Caption = ""
With Me.Combo4
.Enabled = True
.ListIndex = 0
End With
With Timer1
.Enabled = False
.Interval = 1000
End With
With Me.Option1
.Enabled = True
.Value = True
End With
With Me.Option2
.Enabled = True
.Value = False
End With
With Me.Text1
.Enabled = True
.Text = 25
.MaxLength = 3
End With
With Me.Combo1
.Enabled = True
.ListIndex = 0
End With
With Me.Text2
.MaxLength = 2
.Text = Hour(Now) + 1
.Enabled = False
End With
With Me.Text3
.Enabled = False
.Text = Minute(Now)
.MaxLength = 2
End With
With Text4
.MaxLength = 2
.Text = Second(Now)
.Enabled = False
End With
With Me.Combo2
.ListIndex = 3
.Enabled = True
End With
With Me.Text5
.Enabled = False
.Text = ""
End With
With Me.Combo3
.Enabled = False
.ListIndex = 0
End With
With Me.Text6
.Enabled = False
.Text = ""
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
.Text = ""
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
With Me.Combo4
.ListIndex = 0
End With
Label11.Caption = ""
With Timer1
.Enabled = False
.Interval = 1000
End With
With Me.Option1
.Enabled = True
.Value = True
End With
With Me.Option2
.Enabled = True
.Value = False
End With
With Me.Text1
.Enabled = True
.Text = 25
.MaxLength = 3
End With
With Me.Combo1
.Enabled = True
.ListIndex = 0
End With
With Me.Text2
.MaxLength = 2
.Text = Hour(Now) + 1
.Enabled = False
End With
With Me.Text3
.Enabled = False
.Text = Minute(Now)
.MaxLength = 2
End With
With Text4
.MaxLength = 2
.Text = Second(Now)
.Enabled = False
End With
With Me.Combo2
.ListIndex = 3
.Enabled = True
End With
With Me.Text5
.Enabled = False
.Text = ""
End With
With Me.Combo3
.Enabled = False
.ListIndex = 0
End With
With Me.Text6
.Enabled = False
.Text = ""
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
End With
With Me.Combo4
.Enabled = False
End With
End Select
End With
Else
Exit Sub
End If
Exit Sub
Select Case lpTaskCfg.lpTimerType
Case 1
Option1.Value = True
Case 2
Option2.Value = True
End Select
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True
Unload frmTaskDelay
lpTaskCfg.lpRunning = False
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
Exit Sub
End If
Else
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Command1.Enabled = True
Command2.Enabled = True
Command3.Enabled = True
Check2.Enabled = True
Check3.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True
Option1.Enabled = True
Option2.Enabled = True
Select Case lpTaskCfg.lpTimerType
Case 1
Option1.Value = True
Case 2
Option2.Value = True
End Select
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True
Unload frmTaskDelay
lpTaskCfg.lpRunning = False
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
Command4.Enabled = True
ResetAns = MsgBox("任務運行完畢,要復位任務設定嗎?", vbQuestion + vbYesNo, "Ask")
If ResetAns = vbYes Then
TaskApp_ResetRequired
On Error Resume Next
Label11.Caption = ""
With Me.Combo4
.Enabled = True
.ListIndex = 0
End With
With Timer1
.Enabled = False
.Interval = 1000
End With
With Me.Option1
.Enabled = True
.Value = True
End With
With Me.Option2
.Enabled = True
.Value = False
End With
With Me.Text1
.Enabled = True
.Text = 25
.MaxLength = 3
End With
With Me.Combo1
.Enabled = True
.ListIndex = 0
End With
With Me.Text2
.MaxLength = 2
.Text = Hour(Now) + 1
.Enabled = False
End With
With Me.Text3
.Enabled = False
.Text = Minute(Now)
.MaxLength = 2
End With
With Text4
.MaxLength = 2
.Text = Second(Now)
.Enabled = False
End With
With Me.Combo2
.ListIndex = 3
.Enabled = True
End With
With Me.Text5
.Enabled = False
.Text = ""
End With
With Me.Combo3
.Enabled = False
.ListIndex = 0
End With
With Me.Text6
.Enabled = False
.Text = ""
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
.Text = ""
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
With Me.Combo4
.ListIndex = 0
End With
Label11.Caption = ""
With Timer1
.Enabled = False
.Interval = 1000
End With
With Me.Option1
.Enabled = True
.Value = True
End With
With Me.Option2
.Enabled = True
.Value = False
End With
With Me.Text1
.Enabled = True
.Text = 25
.MaxLength = 3
End With
With Me.Combo1
.Enabled = True
.ListIndex = 0
End With
With Me.Text2
.MaxLength = 2
.Text = Hour(Now) + 1
.Enabled = False
End With
With Me.Text3
.Enabled = False
.Text = Minute(Now)
.MaxLength = 2
End With
With Text4
.MaxLength = 2
.Text = Second(Now)
.Enabled = False
End With
With Me.Combo2
.ListIndex = 3
.Enabled = True
End With
With Me.Text5
.Enabled = False
.Text = ""
End With
With Me.Combo3
.Enabled = False
.ListIndex = 0
End With
With Me.Text6
.Enabled = False
.Text = ""
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
Else
Exit Sub
End If
Exit Sub
Select Case lpTaskCfg.lpTimerType
Case 1
Option1.Value = True
Case 2
Option2.Value = True
End Select
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True
Unload frmTaskDelay
lpTaskCfg.lpRunning = False
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
Exit Sub
Exit Sub
End If
Else
If Check3.Value = 1 Then
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
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Command1.Enabled = True
Command2.Enabled = True
Command3.Enabled = True
Check2.Enabled = True
Check3.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True
Option1.Enabled = True
Option2.Enabled = True
Select Case lpTaskCfg.lpTimerType
Case 1
Option1.Value = True
Case 2
Option2.Value = True
End Select
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True
Unload frmTaskDelay
lpTaskCfg.lpRunning = False
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
Command4.Enabled = True
ResetAns = MsgBox("任務運行完畢,要復位任務設定嗎?", vbQuestion + vbYesNo, "Ask")
If ResetAns = vbYes Then
TaskApp_ResetRequired
On Error Resume Next
Label11.Caption = ""
With Me.Combo4
.Enabled = True
.ListIndex = 0
End With
With Timer1
.Enabled = False
.Interval = 1000
End With
With Me.Option1
.Enabled = True
.Value = True
End With
With Me.Option2
.Enabled = True
.Value = False
End With
With Me.Text1
.Enabled = True
.Text = 25
.MaxLength = 3
End With
With Me.Combo1
.Enabled = True
.ListIndex = 0
End With
With Me.Text2
.MaxLength = 2
.Text = Hour(Now) + 1
.Enabled = False
End With
With Me.Text3
.Enabled = False
.Text = Minute(Now)
.MaxLength = 2
End With
With Text4
.MaxLength = 2
.Text = Second(Now)
.Enabled = False
End With
With Me.Combo2
.ListIndex = 3
.Enabled = True
End With
With Me.Text5
.Enabled = False
.Text = ""
End With
With Me.Combo3
.Enabled = False
.ListIndex = 0
End With
With Me.Text6
.Enabled = False
.Text = ""
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
.Text = ""
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
With Me.Combo4
.ListIndex = 0
End With
Label11.Caption = ""
With Timer1
.Enabled = False
.Interval = 1000
End With
With Me.Option1
.Enabled = True
.Value = True
End With
With Me.Option2
.Enabled = True
.Value = False
End With
With Me.Text1
.Enabled = True
.Text = 25
.MaxLength = 3
End With
With Me.Combo1
.Enabled = True
.ListIndex = 0
End With
With Me.Text2
.MaxLength = 2
.Text = Hour(Now) + 1
.Enabled = False
End With
With Me.Text3
.Enabled = False
.Text = Minute(Now)
.MaxLength = 2
End With
With Text4
.MaxLength = 2
.Text = Second(Now)
.Enabled = False
End With
With Me.Combo2
.ListIndex = 3
.Enabled = True
End With
With Me.Text5
.Enabled = False
.Text = ""
End With
With Me.Combo3
.Enabled = False
.ListIndex = 0
End With
With Me.Text6
.Enabled = False
.Text = ""
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
Else
Exit Sub
End If
Exit Sub
Select Case lpTaskCfg.lpTimerType
Case 1
Option1.Value = True
Case 2
Option2.Value = True
End Select
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True
Unload frmTaskDelay
lpTaskCfg.lpRunning = False
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
Exit Sub
Else
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
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Command1.Enabled = True
Command2.Enabled = True
Command3.Enabled = True
Check2.Enabled = True
Check3.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True
Option1.Enabled = True
Option2.Enabled = True
Select Case lpTaskCfg.lpTimerType
Case 1
Option1.Value = True
Case 2
Option2.Value = True
End Select
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True
Unload frmTaskDelay
lpTaskCfg.lpRunning = False
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
Command4.Enabled = True
ResetAns = MsgBox("任務運行完畢,要復位任務設定嗎?", vbQuestion + vbYesNo, "Ask")
If ResetAns = vbYes Then
TaskApp_ResetRequired
On Error Resume Next
Label11.Caption = ""
With Me.Combo4
.Enabled = True
.ListIndex = 0
End With
With Timer1
.Enabled = False
.Interval = 1000
End With
With Me.Option1
.Enabled = True
.Value = True
End With
With Me.Option2
.Enabled = True
.Value = False
End With
With Me.Text1
.Enabled = True
.Text = 25
.MaxLength = 3
End With
With Me.Combo1
.Enabled = True
.ListIndex = 0
End With
With Me.Text2
.MaxLength = 2
.Text = Hour(Now) + 1
.Enabled = False
End With
With Me.Text3
.Enabled = False
.Text = Minute(Now)
.MaxLength = 2
End With
With Text4
.MaxLength = 2
.Text = Second(Now)
.Enabled = False
End With
With Me.Combo2
.ListIndex = 3
.Enabled = True
End With
With Me.Text5
.Enabled = False
.Text = ""
End With
With Me.Combo3
.Enabled = False
.ListIndex = 0
End With
With Me.Text6
.Enabled = False
.Text = ""
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
.Text = ""
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
With Me.Combo4
.ListIndex = 0
End With
Label11.Caption = ""
With Timer1
.Enabled = False
.Interval = 1000
End With
With Me.Option1
.Enabled = True
.Value = True
End With
With Me.Option2
.Enabled = True
.Value = False
End With
With Me.Text1
.Enabled = True
.Text = 25
.MaxLength = 3
End With
With Me.Combo1
.Enabled = True
.ListIndex = 0
End With
With Me.Text2
.MaxLength = 2
.Text = Hour(Now) + 1
.Enabled = False
End With
With Me.Text3
.Enabled = False
.Text = Minute(Now)
.MaxLength = 2
End With
With Text4
.MaxLength = 2
.Text = Second(Now)
.Enabled = False
End With
With Me.Combo2
.ListIndex = 3
.Enabled = True
End With
With Me.Text5
.Enabled = False
.Text = ""
End With
With Me.Combo3
.Enabled = False
.ListIndex = 0
End With
With Me.Text6
.Enabled = False
.Text = ""
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
Else
Exit Sub
End If
Exit Sub
Select Case lpTaskCfg.lpTimerType
Case 1
Option1.Value = True
Case 2
Option2.Value = True
End Select
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True
Unload frmTaskDelay
lpTaskCfg.lpRunning = False
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
Exit Sub
End If
End If
Case "ScreenOff"
CloseScreenA True
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Command1.Enabled = True
Command2.Enabled = True
Command3.Enabled = True
Check2.Enabled = True
Check3.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True
Option1.Enabled = True
Option2.Enabled = True
Select Case lpTaskCfg.lpTimerType
Case 1
Option1.Value = True
Case 2
Option2.Value = True
End Select
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True
Unload frmTaskDelay
lpTaskCfg.lpRunning = False
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
Command4.Enabled = True
ResetAns = MsgBox("任務運行完畢,要復位任務設定嗎?", vbQuestion + vbYesNo, "Ask")
If ResetAns = vbYes Then
TaskApp_ResetRequired
On Error Resume Next
Label11.Caption = ""
With Me.Combo4
.Enabled = True
.ListIndex = 0
End With
With Timer1
.Enabled = False
.Interval = 1000
End With
With Me.Option1
.Enabled = True
.Value = True
End With
With Me.Option2
.Enabled = True
.Value = False
End With
With Me.Text1
.Enabled = True
.Text = 25
.MaxLength = 3
End With
With Me.Combo1
.Enabled = True
.ListIndex = 0
End With
With Me.Text2
.MaxLength = 2
.Text = Hour(Now) + 1
.Enabled = False
End With
With Me.Text3
.Enabled = False
.Text = Minute(Now)
.MaxLength = 2
End With
With Text4
.MaxLength = 2
.Text = Second(Now)
.Enabled = False
End With
With Me.Combo2
.ListIndex = 3
.Enabled = True
End With
With Me.Text5
.Enabled = False
.Text = ""
End With
With Me.Combo3
.Enabled = False
.ListIndex = 0
End With
With Me.Text6
.Enabled = False
.Text = ""
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
.Text = ""
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
With Me.Combo4
.ListIndex = 0
End With
Label11.Caption = ""
With Timer1
.Enabled = False
.Interval = 1000
End With
With Me.Option1
.Enabled = True
.Value = True
End With
With Me.Option2
.Enabled = True
.Value = False
End With
With Me.Text1
.Enabled = True
.Text = 25
.MaxLength = 3
End With
With Me.Combo1
.Enabled = True
.ListIndex = 0
End With
With Me.Text2
.MaxLength = 2
.Text = Hour(Now) + 1
.Enabled = False
End With
With Me.Text3
.Enabled = False
.Text = Minute(Now)
.MaxLength = 2
End With
With Text4
.MaxLength = 2
.Text = Second(Now)
.Enabled = False
End With
With Me.Combo2
.ListIndex = 3
.Enabled = True
End With
With Me.Text5
.Enabled = False
.Text = ""
End With
With Me.Combo3
.Enabled = False
.ListIndex = 0
End With
With Me.Text6
.Enabled = False
.Text = ""
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
Else
Exit Sub
End If
Exit Sub
Select Case lpTaskCfg.lpTimerType
Case 1
Option1.Value = True
Case 2
Option2.Value = True
End Select
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True
Unload frmTaskDelay
lpTaskCfg.lpRunning = False
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
Exit Sub
Case "ScreenSaver"
On Error Resume Next
Dim lRet As Long
lRet = SendMessage(frmTaskMain.hwnd, WM_SYSCOMMAND, SC_SCREENSAVE, 0&)
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Command1.Enabled = True
Command2.Enabled = True
Command3.Enabled = True
Check2.Enabled = True
Check3.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True
Option1.Enabled = True
Option2.Enabled = True
Select Case lpTaskCfg.lpTimerType
Case 1
Option1.Value = True
Case 2
Option2.Value = True
End Select
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True
Unload frmTaskDelay
lpTaskCfg.lpRunning = False
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
Command4.Enabled = True
ResetAns = MsgBox("任務運行完畢,要復位任務設定嗎?", vbQuestion + vbYesNo, "Ask")
If ResetAns = vbYes Then
TaskApp_ResetRequired
On Error Resume Next
Label11.Caption = ""
With Me.Combo4
.Enabled = True
.ListIndex = 0
End With
With Timer1
.Enabled = False
.Interval = 1000
End With
With Me.Option1
.Enabled = True
.Value = True
End With
With Me.Option2
.Enabled = True
.Value = False
End With
With Me.Text1
.Enabled = True
.Text = 25
.MaxLength = 3
End With
With Me.Combo1
.Enabled = True
.ListIndex = 0
End With
With Me.Text2
.MaxLength = 2
.Text = Hour(Now) + 1
.Enabled = False
End With
With Me.Text3
.Enabled = False
.Text = Minute(Now)
.MaxLength = 2
End With
With Text4
.MaxLength = 2
.Text = Second(Now)
.Enabled = False
End With
With Me.Combo2
.ListIndex = 3
.Enabled = True
End With
With Me.Text5
.Enabled = False
.Text = ""
End With
With Me.Combo3
.Enabled = False
.ListIndex = 0
End With
With Me.Text6
.Enabled = False
.Text = ""
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
.Text = ""
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
With Me.Combo4
.ListIndex = 0
End With
Label11.Caption = ""
With Timer1
.Enabled = False
.Interval = 1000
End With
With Me.Option1
.Enabled = True
.Value = True
End With
With Me.Option2
.Enabled = True
.Value = False
End With
With Me.Text1
.Enabled = True
.Text = 25
.MaxLength = 3
End With
With Me.Combo1
.Enabled = True
.ListIndex = 0
End With
With Me.Text2
.MaxLength = 2
.Text = Hour(Now) + 1
.Enabled = False
End With
With Me.Text3
.Enabled = False
.Text = Minute(Now)
.MaxLength = 2
End With
With Text4
.MaxLength = 2
.Text = Second(Now)
.Enabled = False
End With
With Me.Combo2
.ListIndex = 3
.Enabled = True
End With
With Me.Text5
.Enabled = False
.Text = ""
End With
With Me.Combo3
.Enabled = False
.ListIndex = 0
End With
With Me.Text6
.Enabled = False
.Text = ""
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
Else
Exit Sub
End If
Exit Sub
Select Case lpTaskCfg.lpTimerType
Case 1
Option1.Value = True
Case 2
Option2.Value = True
End Select
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True
Unload frmTaskDelay
lpTaskCfg.lpRunning = False
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
Exit Sub
Case "Message"
With lpDialogCfg
.lpIcon = Combo3.ListIndex
.lpTitle = Trim(Text5.Text)
.lpMessage = Trim(Text6.Text)
End With
If lpDialogCfg.lpIcon = 0 Then MsgBox lpDialogCfg.lpMessage, vbOKOnly, lpDialogCfg.lpTitle
If lpDialogCfg.lpIcon = 1 Then MsgBox lpDialogCfg.lpMessage, vbOKOnly + vbCritical, lpDialogCfg.lpTitle
If lpDialogCfg.lpIcon = 2 Then MsgBox lpDialogCfg.lpMessage, vbOKOnly + vbExclamation, lpDialogCfg.lpTitle
If lpDialogCfg.lpIcon = 3 Then MsgBox lpDialogCfg.lpMessage, vbOKOnly + vbInformation, lpDialogCfg.lpTitle
If lpDialogCfg.lpIcon = 4 Then MsgBox lpDialogCfg.lpMessage, vbOKOnly + vbQuestion, lpDialogCfg.lpTitle
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Command1.Enabled = True
Command2.Enabled = True
Command3.Enabled = True
Check2.Enabled = True
Check3.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True
Option1.Enabled = True
Option2.Enabled = True
Select Case lpTaskCfg.lpTimerType
Case 1
Option1.Value = True
Case 2
Option2.Value = True
End Select
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True
Unload frmTaskDelay
lpTaskCfg.lpRunning = False
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
Command4.Enabled = True
ResetAns = MsgBox("任務運行完畢,要復位任務設定嗎?", vbQuestion + vbYesNo, "Ask")
If ResetAns = vbYes Then
TaskApp_ResetRequired
On Error Resume Next
Label11.Caption = ""
With Me.Combo4
.Enabled = True
.ListIndex = 0
End With
With Timer1
.Enabled = False
.Interval = 1000
End With
With Me.Option1
.Enabled = True
.Value = True
End With
With Me.Option2
.Enabled = True
.Value = False
End With
With Me.Text1
.Enabled = True
.Text = 25
.MaxLength = 3
End With
With Me.Combo1
.Enabled = True
.ListIndex = 0
End With
With Me.Text2
.MaxLength = 2
.Text = Hour(Now) + 1
.Enabled = False
End With
With Me.Text3
.Enabled = False
.Text = Minute(Now)
.MaxLength = 2
End With
With Text4
.MaxLength = 2
.Text = Second(Now)
.Enabled = False
End With
With Me.Combo2
.ListIndex = 3
.Enabled = True
End With
With Me.Text5
.Enabled = False
.Text = ""
End With
With Me.Combo3
.Enabled = False
.ListIndex = 0
End With
With Me.Text6
.Enabled = False
.Text = ""
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
.Text = ""
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
With Me.Combo4
.ListIndex = 0
End With
Label11.Caption = ""
With Timer1
.Enabled = False
.Interval = 1000
End With
With Me.Option1
.Enabled = True
.Value = True
End With
With Me.Option2
.Enabled = True
.Value = False
End With
With Me.Text1
.Enabled = True
.Text = 25
.MaxLength = 3
End With
With Me.Combo1
.Enabled = True
.ListIndex = 0
End With
With Me.Text2
.MaxLength = 2
.Text = Hour(Now) + 1
.Enabled = False
End With
With Me.Text3
.Enabled = False
.Text = Minute(Now)
.MaxLength = 2
End With
With Text4
.MaxLength = 2
.Text = Second(Now)
.Enabled = False
End With
With Me.Combo2
.ListIndex = 3
.Enabled = True
End With
With Me.Text5
.Enabled = False
.Text = ""
End With
With Me.Combo3
.Enabled = False
.ListIndex = 0
End With
With Me.Text6
.Enabled = False
.Text = ""
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
Else
Exit Sub
End If
Exit Sub
Select Case lpTaskCfg.lpTimerType
Case 1
Option1.Value = True
Case 2
Option2.Value = True
End Select
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True
Unload frmTaskDelay
lpTaskCfg.lpRunning = False
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
Exit Sub
Case "RunApp"
With lpAppCfg
.lpAppParam = Trim(Text7.Text)
.lpAppPath = Trim(Label11.Caption)
End With
With lpAppCfg
Select Case Combo4.ListIndex
Case 0
.lpRunMode = vbNormalFocus
Case 1
.lpRunMode = vbNormalNoFocus
Case 2
.lpRunMode = vbMinimizedFocus
Case 3
.lpRunMode = vbMinimizedNoFocus
Case 4
.lpRunMode = vbMaximizedFocus
Case 5
.lpRunMode = vbHide
End Select
End With
ShellExecute hwnd, "Open", lpAppCfg.lpAppPath, lpAppCfg.lpAppParam, vbNullString, lpAppCfg.lpRunMode
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Command1.Enabled = True
Command2.Enabled = True
Command3.Enabled = True
Check2.Enabled = True
Check3.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True
Option1.Enabled = True
Option2.Enabled = True
Select Case lpTaskCfg.lpTimerType
Case 1
Option1.Value = True
Case 2
Option2.Value = True
End Select
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True
Unload frmTaskDelay
lpTaskCfg.lpRunning = False
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
Command4.Enabled = True
ResetAns = MsgBox("任務運行完畢,要復位任務設定嗎?", vbQuestion + vbYesNo, "Ask")
If ResetAns = vbYes Then
TaskApp_ResetRequired
On Error Resume Next
Label11.Caption = ""
With Me.Combo4
.Enabled = True
.ListIndex = 0
End With
With Timer1
.Enabled = False
.Interval = 1000
End With
With Me.Option1
.Enabled = True
.Value = True
End With
With Me.Option2
.Enabled = True
.Value = False
End With
With Me.Text1
.Enabled = True
.Text = 25
.MaxLength = 3
End With
With Me.Combo1
.Enabled = True
.ListIndex = 0
End With
With Me.Text2
.MaxLength = 2
.Text = Hour(Now) + 1
.Enabled = False
End With
With Me.Text3
.Enabled = False
.Text = Minute(Now)
.MaxLength = 2
End With
With Text4
.MaxLength = 2
.Text = Second(Now)
.Enabled = False
End With
With Me.Combo2
.ListIndex = 3
.Enabled = True
End With
With Me.Text5
.Enabled = False
.Text = ""
End With
With Me.Combo3
.Enabled = False
.ListIndex = 0
End With
With Me.Text6
.Enabled = False
.Text = ""
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
.Text = ""
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
With Me.Combo4
.ListIndex = 0
End With
Label11.Caption = ""
With Timer1
.Enabled = False
.Interval = 1000
End With
With Me.Option1
.Enabled = True
.Value = True
End With
With Me.Option2
.Enabled = True
.Value = False
End With
With Me.Text1
.Enabled = True
.Text = 25
.MaxLength = 3
End With
With Me.Combo1
.Enabled = True
.ListIndex = 0
End With
With Me.Text2
.MaxLength = 2
.Text = Hour(Now) + 1
.Enabled = False
End With
With Me.Text3
.Enabled = False
.Text = Minute(Now)
.MaxLength = 2
End With
With Text4
.MaxLength = 2
.Text = Second(Now)
.Enabled = False
End With
With Me.Combo2
.ListIndex = 3
.Enabled = True
End With
With Me.Text5
.Enabled = False
.Text = ""
End With
With Me.Combo3
.Enabled = False
.ListIndex = 0
End With
With Me.Text6
.Enabled = False
.Text = ""
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
Else
Exit Sub
End If
Exit Sub
Select Case lpTaskCfg.lpTimerType
Case 1
Option1.Value = True
Case 2
Option2.Value = True
End Select
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True
Unload frmTaskDelay
lpTaskCfg.lpRunning = False
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
Exit Sub
End Select
End If
Case 2
With lpTaskCfg
.lpCurrentHour = Hour(Now)
.lpCurrentMinute = Minute(Now)
.lpCurrentSecond = Second(Now)
frmTaskTime.Label1.Caption = CStr(.lpCurrentHour) & ":" & CStr(.lpCurrentMinute) & ":" & CStr(.lpCurrentSecond)
End With
If (lpTaskCfg.lpCurrentHour = lpTaskCfg.lpRunHour) And (lpTaskCfg.lpCurrentMinute = lpTaskCfg.lpRunMinute) And (lpTaskCfg.lpCurrentSecond = lpTaskCfg.lpRunSecond) Then
With lpTaskCfg
.lpCurrentHour = .lpRunHour
.lpCurrentMinute = .lpRunMinute
.lpCurrentSecond = .lpRunSecond
frmTaskTime.Label1.Caption = CStr(.lpCurrentHour) & ":" & CStr(.lpCurrentMinute) & ":" & CStr(.lpCurrentSecond)
End With
Select Case lpTaskCfg.lpTaskFriendlyDisplayName
Case "Shutdown"
If Check2.Value = 1 Then
ans = MsgBox("確定要關機嗎?請注意保存數據!", vbExclamation + vbYesNo, "Ask")
If ans = vbYes Then
If Check3.Value = 1 Then
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
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Command1.Enabled = True
Command2.Enabled = True
Command3.Enabled = True
Check2.Enabled = True
Check3.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True
Option1.Enabled = True
Option2.Enabled = True
Select Case lpTaskCfg.lpTimerType
Case 1
Option1.Value = True
Case 2
Option2.Value = True
End Select
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True
Unload frmTaskTime
lpTaskCfg.lpRunning = False
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
Command4.Enabled = True
ResetAns = MsgBox("任務運行完畢,要復位任務設定嗎?", vbQuestion + vbYesNo, "Ask")
If ResetAns = vbYes Then
TaskApp_ResetRequired
On Error Resume Next
Label11.Caption = ""
With Me.Combo4
.Enabled = True
.ListIndex = 0
End With
With Timer1
.Enabled = False
.Interval = 1000
End With
With Me.Option1
.Enabled = True
.Value = True
End With
With Me.Option2
.Enabled = True
.Value = False
End With
With Me.Text1
.Enabled = True
.Text = 25
.MaxLength = 3
End With
With Me.Combo1
.Enabled = True
.ListIndex = 0
End With
With Me.Text2
.MaxLength = 2
.Text = Hour(Now) + 1
.Enabled = False
End With
With Me.Text3
.Enabled = False
.Text = Minute(Now)
.MaxLength = 2
End With
With Text4
.MaxLength = 2
.Text = Second(Now)
.Enabled = False
End With
With Me.Combo2
.ListIndex = 3
.Enabled = True
End With
With Me.Text5
.Enabled = False
.Text = ""
End With
With Me.Combo3
.Enabled = False
.ListIndex = 0
End With
With Me.Text6
.Enabled = False
.Text = ""
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
.Text = ""
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
With Me.Combo4
.ListIndex = 0
End With
Label11.Caption = ""
With Timer1
.Enabled = False
.Interval = 1000
End With
With Me.Option1
.Enabled = True
.Value = True
End With
With Me.Option2
.Enabled = True
.Value = False
End With
With Me.Text1
.Enabled = True
.Text = 25
.MaxLength = 3
End With
With Me.Combo1
.Enabled = True
.ListIndex = 0
End With
With Me.Text2
.MaxLength = 2
.Text = Hour(Now) + 1
.Enabled = False
End With
With Me.Text3
.Enabled = False
.Text = Minute(Now)
.MaxLength = 2
End With
With Text4
.MaxLength = 2
.Text = Second(Now)
.Enabled = False
End With
With Me.Combo2
.ListIndex = 3
.Enabled = True
End With
With Me.Text5
.Enabled = False
.Text = ""
End With
With Me.Combo3
.Enabled = False
.ListIndex = 0
End With
With Me.Text6
.Enabled = False
.Text = ""
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
Else
Exit Sub
End If
Exit Sub
Select Case lpTaskCfg.lpTimerType
Case 1
Option1.Value = True
Case 2
Option2.Value = True
End Select
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True
Unload frmTaskDelay
lpTaskCfg.lpRunning = False
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
Exit Sub
Else
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
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Command1.Enabled = True
Command2.Enabled = True
Command3.Enabled = True
Check2.Enabled = True
Check3.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True
Option1.Enabled = True
Option2.Enabled = True
Select Case lpTaskCfg.lpTimerType
Case 1
Option1.Value = True
Case 2
Option2.Value = True
End Select
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True
Unload frmTaskTime
lpTaskCfg.lpRunning = False
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
Command4.Enabled = True
ResetAns = MsgBox("任務運行完畢,要復位任務設定嗎?", vbQuestion + vbYesNo, "Ask")
If ResetAns = vbYes Then
TaskApp_ResetRequired
On Error Resume Next
Label11.Caption = ""
With Me.Combo4
.Enabled = True
.ListIndex = 0
End With
With Timer1
.Enabled = False
.Interval = 1000
End With
With Me.Option1
.Enabled = True
.Value = True
End With
With Me.Option2
.Enabled = True
.Value = False
End With
With Me.Text1
.Enabled = True
.Text = 25
.MaxLength = 3
End With
With Me.Combo1
.Enabled = True
.ListIndex = 0
End With
With Me.Text2
.MaxLength = 2
.Text = Hour(Now) + 1
.Enabled = False
End With
With Me.Text3
.Enabled = False
.Text = Minute(Now)
.MaxLength = 2
End With
With Text4
.MaxLength = 2
.Text = Second(Now)
.Enabled = False
End With
With Me.Combo2
.ListIndex = 3
.Enabled = True
End With
With Me.Text5
.Enabled = False
.Text = ""
End With
With Me.Combo3
.Enabled = False
.ListIndex = 0
End With
With Me.Text6
.Enabled = False
.Text = ""
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
.Text = ""
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
With Me.Combo4
.ListIndex = 0
End With
Label11.Caption = ""
With Timer1
.Enabled = False
.Interval = 1000
End With
With Me.Option1
.Enabled = True
.Value = True
End With
With Me.Option2
.Enabled = True
.Value = False
End With
With Me.Text1
.Enabled = True
.Text = 25
.MaxLength = 3
End With
With Me.Combo1
.Enabled = True
.ListIndex = 0
End With
With Me.Text2
.MaxLength = 2
.Text = Hour(Now) + 1
.Enabled = False
End With
With Me.Text3
.Enabled = False
.Text = Minute(Now)
.MaxLength = 2
End With
With Text4
.MaxLength = 2
.Text = Second(Now)
.Enabled = False
End With
With Me.Combo2
.ListIndex = 3
.Enabled = True
End With
With Me.Text5
.Enabled = False
.Text = ""
End With
With Me.Combo3
.Enabled = False
.ListIndex = 0
End With
With Me.Text6
.Enabled = False
.Text = ""
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
Else
Exit Sub
End If
Exit Sub
Select Case lpTaskCfg.lpTimerType
Case 1
Option1.Value = True
Case 2
Option2.Value = True
End Select
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True
Unload frmTaskDelay
lpTaskCfg.lpRunning = False
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
End With
With Me.Combo4
.Enabled = False
End With
End Select
End With
Exit Sub
End If
Else
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Command1.Enabled = True
Command2.Enabled = True
Command3.Enabled = True
Check2.Enabled = True
Check3.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True
Option1.Enabled = True
Option2.Enabled = True
Select Case lpTaskCfg.lpTimerType
Case 1
Option1.Value = True
Case 2
Option2.Value = True
End Select
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True
Unload frmTaskTime
lpTaskCfg.lpRunning = False
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
Command4.Enabled = True
ResetAns = MsgBox("任務運行完畢,要復位任務設定嗎?", vbQuestion + vbYesNo, "Ask")
If ResetAns = vbYes Then
TaskApp_ResetRequired
On Error Resume Next
Label11.Caption = ""
With Me.Combo4
.Enabled = True
.ListIndex = 0
End With
With Timer1
.Enabled = False
.Interval = 1000
End With
With Me.Option1
.Enabled = True
.Value = True
End With
With Me.Option2
.Enabled = True
.Value = False
End With
With Me.Text1
.Enabled = True
.Text = 25
.MaxLength = 3
End With
With Me.Combo1
.Enabled = True
.ListIndex = 0
End With
With Me.Text2
.MaxLength = 2
.Text = Hour(Now) + 1
.Enabled = False
End With
With Me.Text3
.Enabled = False
.Text = Minute(Now)
.MaxLength = 2
End With
With Text4
.MaxLength = 2
.Text = Second(Now)
.Enabled = False
End With
With Me.Combo2
.ListIndex = 3
.Enabled = True
End With
With Me.Text5
.Enabled = False
.Text = ""
End With
With Me.Combo3
.Enabled = False
.ListIndex = 0
End With
With Me.Text6
.Enabled = False
.Text = ""
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
.Text = ""
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
With Me.Combo4
.ListIndex = 0
End With
Label11.Caption = ""
With Timer1
.Enabled = False
.Interval = 1000
End With
With Me.Option1
.Enabled = True
.Value = True
End With
With Me.Option2
.Enabled = True
.Value = False
End With
With Me.Text1
.Enabled = True
.Text = 25
.MaxLength = 3
End With
With Me.Combo1
.Enabled = True
.ListIndex = 0
End With
With Me.Text2
.MaxLength = 2
.Text = Hour(Now) + 1
.Enabled = False
End With
With Me.Text3
.Enabled = False
.Text = Minute(Now)
.MaxLength = 2
End With
With Text4
.MaxLength = 2
.Text = Second(Now)
.Enabled = False
End With
With Me.Combo2
.ListIndex = 3
.Enabled = True
End With
With Me.Text5
.Enabled = False
.Text = ""
End With
With Me.Combo3
.Enabled = False
.ListIndex = 0
End With
With Me.Text6
.Enabled = False
.Text = ""
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
Else
Exit Sub
End If
Exit Sub
Select Case lpTaskCfg.lpTimerType
Case 1
Option1.Value = True
Case 2
Option2.Value = True
End Select
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True
Unload frmTaskDelay
lpTaskCfg.lpRunning = False
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
Exit Sub
Exit Sub
End If
Else
If Check3.Value = 1 Then
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
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Command1.Enabled = True
Command2.Enabled = True
Command3.Enabled = True
Check2.Enabled = True
Check3.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True
Option1.Enabled = True
Option2.Enabled = True
Select Case lpTaskCfg.lpTimerType
Case 1
Option1.Value = True
Case 2
Option2.Value = True
End Select
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True
Unload frmTaskTime
lpTaskCfg.lpRunning = False
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
End With
With Me.Combo4
.Enabled = False
End With
End Select
End With
Command4.Enabled = True
ResetAns = MsgBox("任務運行完畢,要復位任務設定嗎?", vbQuestion + vbYesNo, "Ask")
If ResetAns = vbYes Then
TaskApp_ResetRequired
On Error Resume Next
Label11.Caption = ""
With Me.Combo4
.Enabled = True
.ListIndex = 0
End With
With Timer1
.Enabled = False
.Interval = 1000
End With
With Me.Option1
.Enabled = True
.Value = True
End With
With Me.Option2
.Enabled = True
.Value = False
End With
With Me.Text1
.Enabled = True
.Text = 25
.MaxLength = 3
End With
With Me.Combo1
.Enabled = True
.ListIndex = 0
End With
With Me.Text2
.MaxLength = 2
.Text = Hour(Now) + 1
.Enabled = False
End With
With Me.Text3
.Enabled = False
.Text = Minute(Now)
.MaxLength = 2
End With
With Text4
.MaxLength = 2
.Text = Second(Now)
.Enabled = False
End With
With Me.Combo2
.ListIndex = 3
.Enabled = True
End With
With Me.Text5
.Enabled = False
.Text = ""
End With
With Me.Combo3
.Enabled = False
.ListIndex = 0
End With
With Me.Text6
.Enabled = False
.Text = ""
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
.Text = ""
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
With Me.Combo4
.ListIndex = 0
End With
Label11.Caption = ""
With Timer1
.Enabled = False
.Interval = 1000
End With
With Me.Option1
.Enabled = True
.Value = True
End With
With Me.Option2
.Enabled = True
.Value = False
End With
With Me.Text1
.Enabled = True
.Text = 25
.MaxLength = 3
End With
With Me.Combo1
.Enabled = True
.ListIndex = 0
End With
With Me.Text2
.MaxLength = 2
.Text = Hour(Now) + 1
.Enabled = False
End With
With Me.Text3
.Enabled = False
.Text = Minute(Now)
.MaxLength = 2
End With
With Text4
.MaxLength = 2
.Text = Second(Now)
.Enabled = False
End With
With Me.Combo2
.ListIndex = 3
.Enabled = True
End With
With Me.Text5
.Enabled = False
.Text = ""
End With
With Me.Combo3
.Enabled = False
.ListIndex = 0
End With
With Me.Text6
.Enabled = False
.Text = ""
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
Else
Exit Sub
End If
Exit Sub
Select Case lpTaskCfg.lpTimerType
Case 1
Option1.Value = True
Case 2
Option2.Value = True
End Select
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True
Unload frmTaskDelay
lpTaskCfg.lpRunning = False
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
Exit Sub
End If
End If
Case "Reboot"
If Check2.Value = 1 Then
ans = MsgBox("確定要重啟嗎?請注意保存數據!", vbExclamation + vbYesNo, "Ask")
If ans = vbYes Then
If Check3.Value = 1 Then
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
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Command1.Enabled = True
Command2.Enabled = True
Command3.Enabled = True
Check2.Enabled = True
Check3.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True
Option1.Enabled = True
Option2.Enabled = True
Select Case lpTaskCfg.lpTimerType
Case 1
Option1.Value = True
Case 2
Option2.Value = True
End Select
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True
Unload frmTaskTime
lpTaskCfg.lpRunning = False
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
Command4.Enabled = True
ResetAns = MsgBox("任務運行完畢,要復位任務設定嗎?", vbQuestion + vbYesNo, "Ask")
If ResetAns = vbYes Then
TaskApp_ResetRequired
On Error Resume Next
Label11.Caption = ""
With Me.Combo4
.Enabled = True
.ListIndex = 0
End With
With Timer1
.Enabled = False
.Interval = 1000
End With
With Me.Option1
.Enabled = True
.Value = True
End With
With Me.Option2
.Enabled = True
.Value = False
End With
With Me.Text1
.Enabled = True
.Text = 25
.MaxLength = 3
End With
With Me.Combo1
.Enabled = True
.ListIndex = 0
End With
With Me.Text2
.MaxLength = 2
.Text = Hour(Now) + 1
.Enabled = False
End With
With Me.Text3
.Enabled = False
.Text = Minute(Now)
.MaxLength = 2
End With
With Text4
.MaxLength = 2
.Text = Second(Now)
.Enabled = False
End With
With Me.Combo2
.ListIndex = 3
.Enabled = True
End With
With Me.Text5
.Enabled = False
.Text = ""
End With
With Me.Combo3
.Enabled = False
.ListIndex = 0
End With
With Me.Text6
.Enabled = False
.Text = ""
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
.Text = ""
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
With Me.Combo4
.ListIndex = 0
End With
Label11.Caption = ""
With Timer1
.Enabled = False
.Interval = 1000
End With
With Me.Option1
.Enabled = True
.Value = True
End With
With Me.Option2
.Enabled = True
.Value = False
End With
With Me.Text1
.Enabled = True
.Text = 25
.MaxLength = 3
End With
With Me.Combo1
.Enabled = True
.ListIndex = 0
End With
With Me.Text2
.MaxLength = 2
.Text = Hour(Now) + 1
.Enabled = False
End With
With Me.Text3
.Enabled = False
.Text = Minute(Now)
.MaxLength = 2
End With
With Text4
.MaxLength = 2
.Text = Second(Now)
.Enabled = False
End With
With Me.Combo2
.ListIndex = 3
.Enabled = True
End With
With Me.Text5
.Enabled = False
.Text = ""
End With
With Me.Combo3
.Enabled = False
.ListIndex = 0
End With
With Me.Text6
.Enabled = False
.Text = ""
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
Else
Exit Sub
End If
Exit Sub
Select Case lpTaskCfg.lpTimerType
Case 1
Option1.Value = True
Case 2
Option2.Value = True
End Select
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True
Unload frmTaskDelay
lpTaskCfg.lpRunning = False
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
Exit Sub
Else
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
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Command1.Enabled = True
Command2.Enabled = True
Command3.Enabled = True
Check2.Enabled = True
Check3.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True
Option1.Enabled = True
Option2.Enabled = True
Select Case lpTaskCfg.lpTimerType
Case 1
Option1.Value = True
Case 2
Option2.Value = True
End Select
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True
Unload frmTaskTime
lpTaskCfg.lpRunning = False
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
End With
With Me.Combo4
.Enabled = False
End With
End Select
End With
Command4.Enabled = True
ResetAns = MsgBox("任務運行完畢,要復位任務設定嗎?", vbQuestion + vbYesNo, "Ask")
If ResetAns = vbYes Then
TaskApp_ResetRequired
On Error Resume Next
Label11.Caption = ""
With Me.Combo4
.Enabled = True
.ListIndex = 0
End With
With Timer1
.Enabled = False
.Interval = 1000
End With
With Me.Option1
.Enabled = True
.Value = True
End With
With Me.Option2
.Enabled = True
.Value = False
End With
With Me.Text1
.Enabled = True
.Text = 25
.MaxLength = 3
End With
With Me.Combo1
.Enabled = True
.ListIndex = 0
End With
With Me.Text2
.MaxLength = 2
.Text = Hour(Now) + 1
.Enabled = False
End With
With Me.Text3
.Enabled = False
.Text = Minute(Now)
.MaxLength = 2
End With
With Text4
.MaxLength = 2
.Text = Second(Now)
.Enabled = False
End With
With Me.Combo2
.ListIndex = 3
.Enabled = True
End With
With Me.Text5
.Enabled = False
.Text = ""
End With
With Me.Combo3
.Enabled = False
.ListIndex = 0
End With
With Me.Text6
.Enabled = False
.Text = ""
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
.Text = ""
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
With Me.Combo4
.ListIndex = 0
End With
Label11.Caption = ""
With Timer1
.Enabled = False
.Interval = 1000
End With
With Me.Option1
.Enabled = True
.Value = True
End With
With Me.Option2
.Enabled = True
.Value = False
End With
With Me.Text1
.Enabled = True
.Text = 25
.MaxLength = 3
End With
With Me.Combo1
.Enabled = True
.ListIndex = 0
End With
With Me.Text2
.MaxLength = 2
.Text = Hour(Now) + 1
.Enabled = False
End With
With Me.Text3
.Enabled = False
.Text = Minute(Now)
.MaxLength = 2
End With
With Text4
.MaxLength = 2
.Text = Second(Now)
.Enabled = False
End With
With Me.Combo2
.ListIndex = 3
.Enabled = True
End With
With Me.Text5
.Enabled = False
.Text = ""
End With
With Me.Combo3
.Enabled = False
.ListIndex = 0
End With
With Me.Text6
.Enabled = False
.Text = ""
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
Else
Exit Sub
End If
Exit Sub
Select Case lpTaskCfg.lpTimerType
Case 1
Option1.Value = True
Case 2
Option2.Value = True
End Select
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True
Unload frmTaskDelay
lpTaskCfg.lpRunning = False
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
Exit Sub
End If
Else
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Command1.Enabled = True
Command2.Enabled = True
Command3.Enabled = True
Check2.Enabled = True
Check3.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True
Option1.Enabled = True
Option2.Enabled = True
Select Case lpTaskCfg.lpTimerType
Case 1
Option1.Value = True
Case 2
Option2.Value = True
End Select
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True
Unload frmTaskTime
lpTaskCfg.lpRunning = False
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
Command4.Enabled = True
ResetAns = MsgBox("任務運行完畢,要復位任務設定嗎?", vbQuestion + vbYesNo, "Ask")
If ResetAns = vbYes Then
TaskApp_ResetRequired
On Error Resume Next
Label11.Caption = ""
With Me.Combo4
.Enabled = True
.ListIndex = 0
End With
With Timer1
.Enabled = False
.Interval = 1000
End With
With Me.Option1
.Enabled = True
.Value = True
End With
With Me.Option2
.Enabled = True
.Value = False
End With
With Me.Text1
.Enabled = True
.Text = 25
.MaxLength = 3
End With
With Me.Combo1
.Enabled = True
.ListIndex = 0
End With
With Me.Text2
.MaxLength = 2
.Text = Hour(Now) + 1
.Enabled = False
End With
With Me.Text3
.Enabled = False
.Text = Minute(Now)
.MaxLength = 2
End With
With Text4
.MaxLength = 2
.Text = Second(Now)
.Enabled = False
End With
With Me.Combo2
.ListIndex = 3
.Enabled = True
End With
With Me.Text5
.Enabled = False
.Text = ""
End With
With Me.Combo3
.Enabled = False
.ListIndex = 0
End With
With Me.Text6
.Enabled = False
.Text = ""
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
.Text = ""
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
With Me.Combo4
.ListIndex = 0
End With
Label11.Caption = ""
With Timer1
.Enabled = False
.Interval = 1000
End With
With Me.Option1
.Enabled = True
.Value = True
End With
With Me.Option2
.Enabled = True
.Value = False
End With
With Me.Text1
.Enabled = True
.Text = 25
.MaxLength = 3
End With
With Me.Combo1
.Enabled = True
.ListIndex = 0
End With
With Me.Text2
.MaxLength = 2
.Text = Hour(Now) + 1
.Enabled = False
End With
With Me.Text3
.Enabled = False
.Text = Minute(Now)
.MaxLength = 2
End With
With Text4
.MaxLength = 2
.Text = Second(Now)
.Enabled = False
End With
With Me.Combo2
.ListIndex = 3
.Enabled = True
End With
With Me.Text5
.Enabled = False
.Text = ""
End With
With Me.Combo3
.Enabled = False
.ListIndex = 0
End With
With Me.Text6
.Enabled = False
.Text = ""
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
Else
Exit Sub
End If
Exit Sub
Select Case lpTaskCfg.lpTimerType
Case 1
Option1.Value = True
Case 2
Option2.Value = True
End Select
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True
Unload frmTaskDelay
lpTaskCfg.lpRunning = False
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
Exit Sub
Exit Sub
End If
Else
If Check3.Value = 1 Then
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
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Command1.Enabled = True
Command2.Enabled = True
Command3.Enabled = True
Check2.Enabled = True
Check3.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True
Option1.Enabled = True
Option2.Enabled = True
Select Case lpTaskCfg.lpTimerType
Case 1
Option1.Value = True
Case 2
Option2.Value = True
End Select
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True
Unload frmTaskTime
lpTaskCfg.lpRunning = False
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
Command4.Enabled = True
ResetAns = MsgBox("任務運行完畢,要復位任務設定嗎?", vbQuestion + vbYesNo, "Ask")
If ResetAns = vbYes Then
TaskApp_ResetRequired
On Error Resume Next
Label11.Caption = ""
With Me.Combo4
.Enabled = True
.ListIndex = 0
End With
With Timer1
.Enabled = False
.Interval = 1000
End With
With Me.Option1
.Enabled = True
.Value = True
End With
With Me.Option2
.Enabled = True
.Value = False
End With
With Me.Text1
.Enabled = True
.Text = 25
.MaxLength = 3
End With
With Me.Combo1
.Enabled = True
.ListIndex = 0
End With
With Me.Text2
.MaxLength = 2
.Text = Hour(Now) + 1
.Enabled = False
End With
With Me.Text3
.Enabled = False
.Text = Minute(Now)
.MaxLength = 2
End With
With Text4
.MaxLength = 2
.Text = Second(Now)
.Enabled = False
End With
With Me.Combo2
.ListIndex = 3
.Enabled = True
End With
With Me.Text5
.Enabled = False
.Text = ""
End With
With Me.Combo3
.Enabled = False
.ListIndex = 0
End With
With Me.Text6
.Enabled = False
.Text = ""
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
.Text = ""
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
With Me.Combo4
.ListIndex = 0
End With
Label11.Caption = ""
With Timer1
.Enabled = False
.Interval = 1000
End With
With Me.Option1
.Enabled = True
.Value = True
End With
With Me.Option2
.Enabled = True
.Value = False
End With
With Me.Text1
.Enabled = True
.Text = 25
.MaxLength = 3
End With
With Me.Combo1
.Enabled = True
.ListIndex = 0
End With
With Me.Text2
.MaxLength = 2
.Text = Hour(Now) + 1
.Enabled = False
End With
With Me.Text3
.Enabled = False
.Text = Minute(Now)
.MaxLength = 2
End With
With Text4
.MaxLength = 2
.Text = Second(Now)
.Enabled = False
End With
With Me.Combo2
.ListIndex = 3
.Enabled = True
End With
With Me.Text5
.Enabled = False
.Text = ""
End With
With Me.Combo3
.Enabled = False
.ListIndex = 0
End With
With Me.Text6
.Enabled = False
.Text = ""
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
Else
Exit Sub
End If
Exit Sub
Select Case lpTaskCfg.lpTimerType
Case 1
Option1.Value = True
Case 2
Option2.Value = True
End Select
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True
Unload frmTaskDelay
lpTaskCfg.lpRunning = False
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
Exit Sub
Else
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
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Command1.Enabled = True
Command2.Enabled = True
Command3.Enabled = True
Check2.Enabled = True
Check3.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True
Option1.Enabled = True
Option2.Enabled = True
Select Case lpTaskCfg.lpTimerType
Case 1
Option1.Value = True
Case 2
Option2.Value = True
End Select
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True
Unload frmTaskTime
lpTaskCfg.lpRunning = False
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
Command4.Enabled = True
ResetAns = MsgBox("任務運行完畢,要復位任務設定嗎?", vbQuestion + vbYesNo, "Ask")
If ResetAns = vbYes Then
TaskApp_ResetRequired
On Error Resume Next
Label11.Caption = ""
With Me.Combo4
.Enabled = True
.ListIndex = 0
End With
With Timer1
.Enabled = False
.Interval = 1000
End With
With Me.Option1
.Enabled = True
.Value = True
End With
With Me.Option2
.Enabled = True
.Value = False
End With
With Me.Text1
.Enabled = True
.Text = 25
.MaxLength = 3
End With
With Me.Combo1
.Enabled = True
.ListIndex = 0
End With
With Me.Text2
.MaxLength = 2
.Text = Hour(Now) + 1
.Enabled = False
End With
With Me.Text3
.Enabled = False
.Text = Minute(Now)
.MaxLength = 2
End With
With Text4
.MaxLength = 2
.Text = Second(Now)
.Enabled = False
End With
With Me.Combo2
.ListIndex = 3
.Enabled = True
End With
With Me.Text5
.Enabled = False
.Text = ""
End With
With Me.Combo3
.Enabled = False
.ListIndex = 0
End With
With Me.Text6
.Enabled = False
.Text = ""
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
.Text = ""
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
With Me.Combo4
.ListIndex = 0
End With
Label11.Caption = ""
With Timer1
.Enabled = False
.Interval = 1000
End With
With Me.Option1
.Enabled = True
.Value = True
End With
With Me.Option2
.Enabled = True
.Value = False
End With
With Me.Text1
.Enabled = True
.Text = 25
.MaxLength = 3
End With
With Me.Combo1
.Enabled = True
.ListIndex = 0
End With
With Me.Text2
.MaxLength = 2
.Text = Hour(Now) + 1
.Enabled = False
End With
With Me.Text3
.Enabled = False
.Text = Minute(Now)
.MaxLength = 2
End With
With Text4
.MaxLength = 2
.Text = Second(Now)
.Enabled = False
End With
With Me.Combo2
.ListIndex = 3
.Enabled = True
End With
With Me.Text5
.Enabled = False
.Text = ""
End With
With Me.Combo3
.Enabled = False
.ListIndex = 0
End With
With Me.Text6
.Enabled = False
.Text = ""
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
Else
Exit Sub
End If
Exit Sub
Select Case lpTaskCfg.lpTimerType
Case 1
Option1.Value = True
Case 2
Option2.Value = True
End Select
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True
Unload frmTaskDelay
lpTaskCfg.lpRunning = False
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
Exit Sub
End If
End If
Case "LogOff"
If Check2.Value = 1 Then
ans = MsgBox("確定要注銷嗎?請注意保存數據!", vbExclamation + vbYesNo, "Ask")
If ans = vbYes Then
If Check3.Value = 1 Then
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
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Command1.Enabled = True
Command2.Enabled = True
Command3.Enabled = True
Check2.Enabled = True
Check3.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True
Option1.Enabled = True
Option2.Enabled = True
Select Case lpTaskCfg.lpTimerType
Case 1
Option1.Value = True
Case 2
Option2.Value = True
End Select
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True
Unload frmTaskTime
lpTaskCfg.lpRunning = False
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
Command4.Enabled = True
ResetAns = MsgBox("任務運行完畢,要復位任務設定嗎?", vbQuestion + vbYesNo, "Ask")
If ResetAns = vbYes Then
TaskApp_ResetRequired
On Error Resume Next
Label11.Caption = ""
With Me.Combo4
.Enabled = True
.ListIndex = 0
End With
With Timer1
.Enabled = False
.Interval = 1000
End With
With Me.Option1
.Enabled = True
.Value = True
End With
With Me.Option2
.Enabled = True
.Value = False
End With
With Me.Text1
.Enabled = True
.Text = 25
.MaxLength = 3
End With
With Me.Combo1
.Enabled = True
.ListIndex = 0
End With
With Me.Text2
.MaxLength = 2
.Text = Hour(Now) + 1
.Enabled = False
End With
With Me.Text3
.Enabled = False
.Text = Minute(Now)
.MaxLength = 2
End With
With Text4
.MaxLength = 2
.Text = Second(Now)
.Enabled = False
End With
With Me.Combo2
.ListIndex = 3
.Enabled = True
End With
With Me.Text5
.Enabled = False
.Text = ""
End With
With Me.Combo3
.Enabled = False
.ListIndex = 0
End With
With Me.Text6
.Enabled = False
.Text = ""
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
.Text = ""
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
With Me.Combo4
.ListIndex = 0
End With
Label11.Caption = ""
With Timer1
.Enabled = False
.Interval = 1000
End With
With Me.Option1
.Enabled = True
.Value = True
End With
With Me.Option2
.Enabled = True
.Value = False
End With
With Me.Text1
.Enabled = True
.Text = 25
.MaxLength = 3
End With
With Me.Combo1
.Enabled = True
.ListIndex = 0
End With
With Me.Text2
.MaxLength = 2
.Text = Hour(Now) + 1
.Enabled = False
End With
With Me.Text3
.Enabled = False
.Text = Minute(Now)
.MaxLength = 2
End With
With Text4
.MaxLength = 2
.Text = Second(Now)
.Enabled = False
End With
With Me.Combo2
.ListIndex = 3
.Enabled = True
End With
With Me.Text5
.Enabled = False
.Text = ""
End With
With Me.Combo3
.Enabled = False
.ListIndex = 0
End With
With Me.Text6
.Enabled = False
.Text = ""
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
Else
Exit Sub
End If
Exit Sub
Select Case lpTaskCfg.lpTimerType
Case 1
Option1.Value = True
Case 2
Option2.Value = True
End Select
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True
Unload frmTaskDelay
lpTaskCfg.lpRunning = False
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
Exit Sub
Else
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
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Command1.Enabled = True
Command2.Enabled = True
Command3.Enabled = True
Check2.Enabled = True
Check3.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True
Option1.Enabled = True
Option2.Enabled = True
Select Case lpTaskCfg.lpTimerType
Case 1
Option1.Value = True
Case 2
Option2.Value = True
End Select
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True
Unload frmTaskTime
lpTaskCfg.lpRunning = False
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
Command4.Enabled = True
ResetAns = MsgBox("任務運行完畢,要復位任務設定嗎?", vbQuestion + vbYesNo, "Ask")
If ResetAns = vbYes Then
TaskApp_ResetRequired
On Error Resume Next
Label11.Caption = ""
With Me.Combo4
.Enabled = True
.ListIndex = 0
End With
With Timer1
.Enabled = False
.Interval = 1000
End With
With Me.Option1
.Enabled = True
.Value = True
End With
With Me.Option2
.Enabled = True
.Value = False
End With
With Me.Text1
.Enabled = True
.Text = 25
.MaxLength = 3
End With
With Me.Combo1
.Enabled = True
.ListIndex = 0
End With
With Me.Text2
.MaxLength = 2
.Text = Hour(Now) + 1
.Enabled = False
End With
With Me.Text3
.Enabled = False
.Text = Minute(Now)
.MaxLength = 2
End With
With Text4
.MaxLength = 2
.Text = Second(Now)
.Enabled = False
End With
With Me.Combo2
.ListIndex = 3
.Enabled = True
End With
With Me.Text5
.Enabled = False
.Text = ""
End With
With Me.Combo3
.Enabled = False
.ListIndex = 0
End With
With Me.Text6
.Enabled = False
.Text = ""
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
.Text = ""
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
With Me.Combo4
.ListIndex = 0
End With
Label11.Caption = ""
With Timer1
.Enabled = False
.Interval = 1000
End With
With Me.Option1
.Enabled = True
.Value = True
End With
With Me.Option2
.Enabled = True
.Value = False
End With
With Me.Text1
.Enabled = True
.Text = 25
.MaxLength = 3
End With
With Me.Combo1
.Enabled = True
.ListIndex = 0
End With
With Me.Text2
.MaxLength = 2
.Text = Hour(Now) + 1
.Enabled = False
End With
With Me.Text3
.Enabled = False
.Text = Minute(Now)
.MaxLength = 2
End With
With Text4
.MaxLength = 2
.Text = Second(Now)
.Enabled = False
End With
With Me.Combo2
.ListIndex = 3
.Enabled = True
End With
With Me.Text5
.Enabled = False
.Text = ""
End With
With Me.Combo3
.Enabled = False
.ListIndex = 0
End With
With Me.Text6
.Enabled = False
.Text = ""
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
Else
Exit Sub
End If
Exit Sub
Select Case lpTaskCfg.lpTimerType
Case 1
Option1.Value = True
Case 2
Option2.Value = True
End Select
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True
Unload frmTaskDelay
lpTaskCfg.lpRunning = False
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
Exit Sub
End If
Else
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Command1.Enabled = True
Command2.Enabled = True
Command3.Enabled = True
Check2.Enabled = True
Check3.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True
Option1.Enabled = True
Option2.Enabled = True
Select Case lpTaskCfg.lpTimerType
Case 1
Option1.Value = True
Case 2
Option2.Value = True
End Select
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True
Unload frmTaskTime
lpTaskCfg.lpRunning = False
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
Command4.Enabled = True
ResetAns = MsgBox("任務運行完畢,要復位任務設定嗎?", vbQuestion + vbYesNo, "Ask")
If ResetAns = vbYes Then
TaskApp_ResetRequired
On Error Resume Next
Label11.Caption = ""
With Me.Combo4
.Enabled = True
.ListIndex = 0
End With
With Timer1
.Enabled = False
.Interval = 1000
End With
With Me.Option1
.Enabled = True
.Value = True
End With
With Me.Option2
.Enabled = True
.Value = False
End With
With Me.Text1
.Enabled = True
.Text = 25
.MaxLength = 3
End With
With Me.Combo1
.Enabled = True
.ListIndex = 0
End With
With Me.Text2
.MaxLength = 2
.Text = Hour(Now) + 1
.Enabled = False
End With
With Me.Text3
.Enabled = False
.Text = Minute(Now)
.MaxLength = 2
End With
With Text4
.MaxLength = 2
.Text = Second(Now)
.Enabled = False
End With
With Me.Combo2
.ListIndex = 3
.Enabled = True
End With
With Me.Text5
.Enabled = False
.Text = ""
End With
With Me.Combo3
.Enabled = False
.ListIndex = 0
End With
With Me.Text6
.Enabled = False
.Text = ""
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
.Text = ""
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
With Me.Combo4
.ListIndex = 0
End With
Label11.Caption = ""
With Timer1
.Enabled = False
.Interval = 1000
End With
With Me.Option1
.Enabled = True
.Value = True
End With
With Me.Option2
.Enabled = True
.Value = False
End With
With Me.Text1
.Enabled = True
.Text = 25
.MaxLength = 3
End With
With Me.Combo1
.Enabled = True
.ListIndex = 0
End With
With Me.Text2
.MaxLength = 2
.Text = Hour(Now) + 1
.Enabled = False
End With
With Me.Text3
.Enabled = False
.Text = Minute(Now)
.MaxLength = 2
End With
With Text4
.MaxLength = 2
.Text = Second(Now)
.Enabled = False
End With
With Me.Combo2
.ListIndex = 3
.Enabled = True
End With
With Me.Text5
.Enabled = False
.Text = ""
End With
With Me.Combo3
.Enabled = False
.ListIndex = 0
End With
With Me.Text6
.Enabled = False
.Text = ""
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
Else
Exit Sub
End If
Exit Sub
Select Case lpTaskCfg.lpTimerType
Case 1
Option1.Value = True
Case 2
Option2.Value = True
End Select
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True
Unload frmTaskDelay
lpTaskCfg.lpRunning = False
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
Exit Sub
Exit Sub
End If
Else
If Check3.Value = 1 Then
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
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Command1.Enabled = True
Command2.Enabled = True
Command3.Enabled = True
Check2.Enabled = True
Check3.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True
Option1.Enabled = True
Option2.Enabled = True
Select Case lpTaskCfg.lpTimerType
Case 1
Option1.Value = True
Case 2
Option2.Value = True
End Select
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True
Unload frmTaskTime
lpTaskCfg.lpRunning = False
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
End With
With Me.Combo4
.Enabled = False
End With
End Select
End With
Command4.Enabled = True
ResetAns = MsgBox("任務運行完畢,要復位任務設定嗎?", vbQuestion + vbYesNo, "Ask")
If ResetAns = vbYes Then
TaskApp_ResetRequired
On Error Resume Next
Label11.Caption = ""
With Me.Combo4
.Enabled = True
.ListIndex = 0
End With
With Timer1
.Enabled = False
.Interval = 1000
End With
With Me.Option1
.Enabled = True
.Value = True
End With
With Me.Option2
.Enabled = True
.Value = False
End With
With Me.Text1
.Enabled = True
.Text = 25
.MaxLength = 3
End With
With Me.Combo1
.Enabled = True
.ListIndex = 0
End With
With Me.Text2
.MaxLength = 2
.Text = Hour(Now) + 1
.Enabled = False
End With
With Me.Text3
.Enabled = False
.Text = Minute(Now)
.MaxLength = 2
End With
With Text4
.MaxLength = 2
.Text = Second(Now)
.Enabled = False
End With
With Me.Combo2
.ListIndex = 3
.Enabled = True
End With
With Me.Text5
.Enabled = False
.Text = ""
End With
With Me.Combo3
.Enabled = False
.ListIndex = 0
End With
With Me.Text6
.Enabled = False
.Text = ""
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
.Text = ""
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
With Me.Combo4
.ListIndex = 0
End With
Label11.Caption = ""
With Timer1
.Enabled = False
.Interval = 1000
End With
With Me.Option1
.Enabled = True
.Value = True
End With
With Me.Option2
.Enabled = True
.Value = False
End With
With Me.Text1
.Enabled = True
.Text = 25
.MaxLength = 3
End With
With Me.Combo1
.Enabled = True
.ListIndex = 0
End With
With Me.Text2
.MaxLength = 2
.Text = Hour(Now) + 1
.Enabled = False
End With
With Me.Text3
.Enabled = False
.Text = Minute(Now)
.MaxLength = 2
End With
With Text4
.MaxLength = 2
.Text = Second(Now)
.Enabled = False
End With
With Me.Combo2
.ListIndex = 3
.Enabled = True
End With
With Me.Text5
.Enabled = False
.Text = ""
End With
With Me.Combo3
.Enabled = False
.ListIndex = 0
End With
With Me.Text6
.Enabled = False
.Text = ""
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
Else
Exit Sub
End If
Exit Sub
Select Case lpTaskCfg.lpTimerType
Case 1
Option1.Value = True
Case 2
Option2.Value = True
End Select
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True
Unload frmTaskDelay
lpTaskCfg.lpRunning = False
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
Exit Sub
Else
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
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Command1.Enabled = True
Command2.Enabled = True
Command3.Enabled = True
Check2.Enabled = True
Check3.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True
Option1.Enabled = True
Option2.Enabled = True
Select Case lpTaskCfg.lpTimerType
Case 1
Option1.Value = True
Case 2
Option2.Value = True
End Select
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True
Unload frmTaskTime
lpTaskCfg.lpRunning = False
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
Command4.Enabled = True
ResetAns = MsgBox("任務運行完畢,要復位任務設定嗎?", vbQuestion + vbYesNo, "Ask")
If ResetAns = vbYes Then
TaskApp_ResetRequired
On Error Resume Next
Label11.Caption = ""
With Me.Combo4
.Enabled = True
.ListIndex = 0
End With
With Timer1
.Enabled = False
.Interval = 1000
End With
With Me.Option1
.Enabled = True
.Value = True
End With
With Me.Option2
.Enabled = True
.Value = False
End With
With Me.Text1
.Enabled = True
.Text = 25
.MaxLength = 3
End With
With Me.Combo1
.Enabled = True
.ListIndex = 0
End With
With Me.Text2
.MaxLength = 2
.Text = Hour(Now) + 1
.Enabled = False
End With
With Me.Text3
.Enabled = False
.Text = Minute(Now)
.MaxLength = 2
End With
With Text4
.MaxLength = 2
.Text = Second(Now)
.Enabled = False
End With
With Me.Combo2
.ListIndex = 3
.Enabled = True
End With
With Me.Text5
.Enabled = False
.Text = ""
End With
With Me.Combo3
.Enabled = False
.ListIndex = 0
End With
With Me.Text6
.Enabled = False
.Text = ""
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
.Text = ""
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
With Me.Combo4
.ListIndex = 0
End With
Label11.Caption = ""
With Timer1
.Enabled = False
.Interval = 1000
End With
With Me.Option1
.Enabled = True
.Value = True
End With
With Me.Option2
.Enabled = True
.Value = False
End With
With Me.Text1
.Enabled = True
.Text = 25
.MaxLength = 3
End With
With Me.Combo1
.Enabled = True
.ListIndex = 0
End With
With Me.Text2
.MaxLength = 2
.Text = Hour(Now) + 1
.Enabled = False
End With
With Me.Text3
.Enabled = False
.Text = Minute(Now)
.MaxLength = 2
End With
With Text4
.MaxLength = 2
.Text = Second(Now)
.Enabled = False
End With
With Me.Combo2
.ListIndex = 3
.Enabled = True
End With
With Me.Text5
.Enabled = False
.Text = ""
End With
With Me.Combo3
.Enabled = False
.ListIndex = 0
End With
With Me.Text6
.Enabled = False
.Text = ""
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
Else
Exit Sub
End If
Exit Sub
Select Case lpTaskCfg.lpTimerType
Case 1
Option1.Value = True
Case 2
Option2.Value = True
End Select
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True
Unload frmTaskDelay
lpTaskCfg.lpRunning = False
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
Exit Sub
End If
End If
Case "ScreenOff"
CloseScreenA True
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Command1.Enabled = True
Command2.Enabled = True
Command3.Enabled = True
Check2.Enabled = True
Check3.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True
Option1.Enabled = True
Option2.Enabled = True
Select Case lpTaskCfg.lpTimerType
Case 1
Option1.Value = True
Case 2
Option2.Value = True
End Select
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True
Unload frmTaskTime
lpTaskCfg.lpRunning = False
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
Command4.Enabled = True
ResetAns = MsgBox("任務運行完畢,要復位任務設定嗎?", vbQuestion + vbYesNo, "Ask")
If ResetAns = vbYes Then
TaskApp_ResetRequired
On Error Resume Next
Label11.Caption = ""
With Me.Combo4
.Enabled = True
.ListIndex = 0
End With
With Timer1
.Enabled = False
.Interval = 1000
End With
With Me.Option1
.Enabled = True
.Value = True
End With
With Me.Option2
.Enabled = True
.Value = False
End With
With Me.Text1
.Enabled = True
.Text = 25
.MaxLength = 3
End With
With Me.Combo1
.Enabled = True
.ListIndex = 0
End With
With Me.Text2
.MaxLength = 2
.Text = Hour(Now) + 1
.Enabled = False
End With
With Me.Text3
.Enabled = False
.Text = Minute(Now)
.MaxLength = 2
End With
With Text4
.MaxLength = 2
.Text = Second(Now)
.Enabled = False
End With
With Me.Combo2
.ListIndex = 3
.Enabled = True
End With
With Me.Text5
.Enabled = False
.Text = ""
End With
With Me.Combo3
.Enabled = False
.ListIndex = 0
End With
With Me.Text6
.Enabled = False
.Text = ""
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
.Text = ""
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
With Me.Combo4
.ListIndex = 0
End With
Label11.Caption = ""
With Timer1
.Enabled = False
.Interval = 1000
End With
With Me.Option1
.Enabled = True
.Value = True
End With
With Me.Option2
.Enabled = True
.Value = False
End With
With Me.Text1
.Enabled = True
.Text = 25
.MaxLength = 3
End With
With Me.Combo1
.Enabled = True
.ListIndex = 0
End With
With Me.Text2
.MaxLength = 2
.Text = Hour(Now) + 1
.Enabled = False
End With
With Me.Text3
.Enabled = False
.Text = Minute(Now)
.MaxLength = 2
End With
With Text4
.MaxLength = 2
.Text = Second(Now)
.Enabled = False
End With
With Me.Combo2
.ListIndex = 3
.Enabled = True
End With
With Me.Text5
.Enabled = False
.Text = ""
End With
With Me.Combo3
.Enabled = False
.ListIndex = 0
End With
With Me.Text6
.Enabled = False
.Text = ""
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
Else
Exit Sub
End If
Exit Sub
Select Case lpTaskCfg.lpTimerType
Case 1
Option1.Value = True
Case 2
Option2.Value = True
End Select
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True
Unload frmTaskDelay
lpTaskCfg.lpRunning = False
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
Exit Sub
Case "ScreenSaver"
On Error Resume Next
lRet = SendMessage(frmTaskMain.hwnd, WM_SYSCOMMAND, SC_SCREENSAVE, 0&)
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Command1.Enabled = True
Command2.Enabled = True
Command3.Enabled = True
Check2.Enabled = True
Check3.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True
Option1.Enabled = True
Option2.Enabled = True
Select Case lpTaskCfg.lpTimerType
Case 1
Option1.Value = True
Case 2
Option2.Value = True
End Select
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True
Unload frmTaskTime
lpTaskCfg.lpRunning = False
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
Command4.Enabled = True
ResetAns = MsgBox("任務運行完畢,要復位任務設定嗎?", vbQuestion + vbYesNo, "Ask")
If ResetAns = vbYes Then
TaskApp_ResetRequired
On Error Resume Next
Label11.Caption = ""
With Me.Combo4
.Enabled = True
.ListIndex = 0
End With
With Timer1
.Enabled = False
.Interval = 1000
End With
With Me.Option1
.Enabled = True
.Value = True
End With
With Me.Option2
.Enabled = True
.Value = False
End With
With Me.Text1
.Enabled = True
.Text = 25
.MaxLength = 3
End With
With Me.Combo1
.Enabled = True
.ListIndex = 0
End With
With Me.Text2
.MaxLength = 2
.Text = Hour(Now) + 1
.Enabled = False
End With
With Me.Text3
.Enabled = False
.Text = Minute(Now)
.MaxLength = 2
End With
With Text4
.MaxLength = 2
.Text = Second(Now)
.Enabled = False
End With
With Me.Combo2
.ListIndex = 3
.Enabled = True
End With
With Me.Text5
.Enabled = False
.Text = ""
End With
With Me.Combo3
.Enabled = False
.ListIndex = 0
End With
With Me.Text6
.Enabled = False
.Text = ""
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
.Text = ""
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
With Me.Combo4
.ListIndex = 0
End With
Label11.Caption = ""
With Timer1
.Enabled = False
.Interval = 1000
End With
With Me.Option1
.Enabled = True
.Value = True
End With
With Me.Option2
.Enabled = True
.Value = False
End With
With Me.Text1
.Enabled = True
.Text = 25
.MaxLength = 3
End With
With Me.Combo1
.Enabled = True
.ListIndex = 0
End With
With Me.Text2
.MaxLength = 2
.Text = Hour(Now) + 1
.Enabled = False
End With
With Me.Text3
.Enabled = False
.Text = Minute(Now)
.MaxLength = 2
End With
With Text4
.MaxLength = 2
.Text = Second(Now)
.Enabled = False
End With
With Me.Combo2
.ListIndex = 3
.Enabled = True
End With
With Me.Text5
.Enabled = False
.Text = ""
End With
With Me.Combo3
.Enabled = False
.ListIndex = 0
End With
With Me.Text6
.Enabled = False
.Text = ""
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
Else
Exit Sub
End If
Exit Sub
Select Case lpTaskCfg.lpTimerType
Case 1
Option1.Value = True
Case 2
Option2.Value = True
End Select
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True
Unload frmTaskDelay
lpTaskCfg.lpRunning = False
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
Exit Sub
Case "Message"
With lpDialogCfg
.lpIcon = Combo3.ListIndex
.lpTitle = Trim(Text5.Text)
.lpMessage = Trim(Text6.Text)
End With
If lpDialogCfg.lpIcon = 0 Then MsgBox lpDialogCfg.lpMessage, vbOKOnly, lpDialogCfg.lpTitle
If lpDialogCfg.lpIcon = 1 Then MsgBox lpDialogCfg.lpMessage, vbOKOnly + vbCritical, lpDialogCfg.lpTitle
If lpDialogCfg.lpIcon = 2 Then MsgBox lpDialogCfg.lpMessage, vbOKOnly + vbExclamation, lpDialogCfg.lpTitle
If lpDialogCfg.lpIcon = 3 Then MsgBox lpDialogCfg.lpMessage, vbOKOnly + vbInformation, lpDialogCfg.lpTitle
If lpDialogCfg.lpIcon = 4 Then MsgBox lpDialogCfg.lpMessage, vbOKOnly + vbQuestion, lpDialogCfg.lpTitle
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Command1.Enabled = True
Command2.Enabled = True
Command3.Enabled = True
Check2.Enabled = True
Check3.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True
Option1.Enabled = True
Option2.Enabled = True
Select Case lpTaskCfg.lpTimerType
Case 1
Option1.Value = True
Case 2
Option2.Value = True
End Select
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True
Unload frmTaskTime
lpTaskCfg.lpRunning = False
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
Command4.Enabled = True
ResetAns = MsgBox("任務運行完畢,要復位任務設定嗎?", vbQuestion + vbYesNo, "Ask")
If ResetAns = vbYes Then
TaskApp_ResetRequired
On Error Resume Next
Label11.Caption = ""
With Me.Combo4
.Enabled = True
.ListIndex = 0
End With
With Timer1
.Enabled = False
.Interval = 1000
End With
With Me.Option1
.Enabled = True
.Value = True
End With
With Me.Option2
.Enabled = True
.Value = False
End With
With Me.Text1
.Enabled = True
.Text = 25
.MaxLength = 3
End With
With Me.Combo1
.Enabled = True
.ListIndex = 0
End With
With Me.Text2
.MaxLength = 2
.Text = Hour(Now) + 1
.Enabled = False
End With
With Me.Text3
.Enabled = False
.Text = Minute(Now)
.MaxLength = 2
End With
With Text4
.MaxLength = 2
.Text = Second(Now)
.Enabled = False
End With
With Me.Combo2
.ListIndex = 3
.Enabled = True
End With
With Me.Text5
.Enabled = False
.Text = ""
End With
With Me.Combo3
.Enabled = False
.ListIndex = 0
End With
With Me.Text6
.Enabled = False
.Text = ""
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
.Text = ""
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
With Me.Combo4
.ListIndex = 0
End With
Label11.Caption = ""
With Timer1
.Enabled = False
.Interval = 1000
End With
With Me.Option1
.Enabled = True
.Value = True
End With
With Me.Option2
.Enabled = True
.Value = False
End With
With Me.Text1
.Enabled = True
.Text = 25
.MaxLength = 3
End With
With Me.Combo1
.Enabled = True
.ListIndex = 0
End With
With Me.Text2
.MaxLength = 2
.Text = Hour(Now) + 1
.Enabled = False
End With
With Me.Text3
.Enabled = False
.Text = Minute(Now)
.MaxLength = 2
End With
With Text4
.MaxLength = 2
.Text = Second(Now)
.Enabled = False
End With
With Me.Combo2
.ListIndex = 3
.Enabled = True
End With
With Me.Text5
.Enabled = False
.Text = ""
End With
With Me.Combo3
.Enabled = False
.ListIndex = 0
End With
With Me.Text6
.Enabled = False
.Text = ""
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
Else
Exit Sub
End If
Exit Sub
Select Case lpTaskCfg.lpTimerType
Case 1
Option1.Value = True
Case 2
Option2.Value = True
End Select
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True
Unload frmTaskDelay
lpTaskCfg.lpRunning = False
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
Exit Sub
Case "RunApp"
With lpAppCfg
.lpAppParam = Trim(Text7.Text)
.lpAppPath = Trim(Label11.Caption)
End With
With lpAppCfg
Select Case Combo4.ListIndex
Case 0
.lpRunMode = vbNormalFocus
Case 1
.lpRunMode = vbNormalNoFocus
Case 2
.lpRunMode = vbMinimizedFocus
Case 3
.lpRunMode = vbMinimizedNoFocus
Case 4
.lpRunMode = vbMaximizedFocus
Case 5
.lpRunMode = vbHide
End Select
End With
ShellExecute hwnd, "Open", lpAppCfg.lpAppPath, lpAppCfg.lpAppParam, vbNullString, lpAppCfg.lpRunMode
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Command1.Enabled = True
Command2.Enabled = True
Command3.Enabled = True
Check2.Enabled = True
Check3.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True
Option1.Enabled = True
Option2.Enabled = True
Select Case lpTaskCfg.lpTimerType
Case 1
Option1.Value = True
Case 2
Option2.Value = True
End Select
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True
Unload frmTaskTime
lpTaskCfg.lpRunning = False
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
Command4.Enabled = True
ResetAns = MsgBox("任務運行完畢,要復位任務設定嗎?", vbQuestion + vbYesNo, "Ask")
If ResetAns = vbYes Then
TaskApp_ResetRequired
On Error Resume Next
Label11.Caption = ""
With Me.Combo4
.Enabled = True
.ListIndex = 0
End With
With Timer1
.Enabled = False
.Interval = 1000
End With
With Me.Option1
.Enabled = True
.Value = True
End With
With Me.Option2
.Enabled = True
.Value = False
End With
With Me.Text1
.Enabled = True
.Text = 25
.MaxLength = 3
End With
With Me.Combo1
.Enabled = True
.ListIndex = 0
End With
With Me.Text2
.MaxLength = 2
.Text = Hour(Now) + 1
.Enabled = False
End With
With Me.Text3
.Enabled = False
.Text = Minute(Now)
.MaxLength = 2
End With
With Text4
.MaxLength = 2
.Text = Second(Now)
.Enabled = False
End With
With Me.Combo2
.ListIndex = 3
.Enabled = True
End With
With Me.Text5
.Enabled = False
.Text = ""
End With
With Me.Combo3
.Enabled = False
.ListIndex = 0
End With
With Me.Text6
.Enabled = False
.Text = ""
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
.Text = ""
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
.Text = ""
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
With Me.Combo4
.ListIndex = 0
End With
Label11.Caption = ""
With Timer1
.Enabled = False
.Interval = 1000
End With
With Me.Option1
.Enabled = True
.Value = True
End With
With Me.Option2
.Enabled = True
.Value = False
End With
With Me.Text1
.Enabled = True
.Text = 25
.MaxLength = 3
End With
With Me.Combo1
.Enabled = True
.ListIndex = 0
End With
With Me.Text2
.MaxLength = 2
.Text = Hour(Now) + 1
.Enabled = False
End With
With Me.Text3
.Enabled = False
.Text = Minute(Now)
.MaxLength = 2
End With
With Text4
.MaxLength = 2
.Text = Second(Now)
.Enabled = False
End With
With Me.Combo2
.ListIndex = 3
.Enabled = True
End With
With Me.Text5
.Enabled = False
.Text = ""
End With
With Me.Combo3
.Enabled = False
.ListIndex = 0
End With
With Me.Text6
.Enabled = False
.Text = ""
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
Else
Exit Sub
End If
Exit Sub
Select Case lpTaskCfg.lpTimerType
Case 1
Option1.Value = True
Case 2
Option2.Value = True
End Select
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True
Unload frmTaskDelay
lpTaskCfg.lpRunning = False
On Error Resume Next
lpTaskCfg.lpTaskEnum = Combo2.ListIndex
With lpTaskCfg
Select Case .lpTaskEnum
Case 0
.lpTaskFriendlyDisplayName = "Shutdown"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 1
.lpTaskFriendlyDisplayName = "Reboot"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 2
.lpTaskFriendlyDisplayName = "LogOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 3
.lpTaskFriendlyDisplayName = "ScreenOff"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 4
.lpTaskFriendlyDisplayName = "ScreenSaver"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 5
.lpTaskFriendlyDisplayName = "Message"
With Me.Text5
.Enabled = True
End With
With Me.Combo3
.Enabled = True
End With
With Me.Text6
.Enabled = True
End With
With Me.Command1
.Enabled = False
End With
With Me.Text7
.Enabled = False
End With
With Me.Combo4
.Enabled = False
End With
Case 6
.lpTaskFriendlyDisplayName = "RunApp"
With Me.Text5
.Enabled = False
End With
With Me.Combo3
.Enabled = False
End With
With Me.Text6
.Enabled = False
End With
With Me.Command1
.Enabled = True
End With
With Me.Text7
.Enabled = True
End With
With Me.Combo4
.Enabled = True
End With
End Select
End With
Exit Sub
End Select
End If
End Select
End Sub


