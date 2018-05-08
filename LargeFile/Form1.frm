VERSION 5.00
Begin VB.Form frmLargefile 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Create Large File - PC-DOS Workshop"
   ClientHeight    =   5535
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9870
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   9870
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "製作(&M)"
      Height          =   375
      Left            =   7695
      TabIndex        =   17
      Top             =   5130
      Width           =   2145
   End
   Begin VB.Frame Frame3 
      Caption         =   "文件的內容"
      Height          =   5010
      Left            =   4695
      TabIndex        =   12
      Top             =   75
      Width           =   5145
      Begin VB.TextBox Text3 
         Height          =   3720
         Left            =   405
         MaxLength       =   2048
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   16
         Text            =   "Form1.frx":030A
         Top             =   1185
         Width           =   4665
      End
      Begin VB.OptionButton Option3 
         Caption         =   "用戶自定義內容(&U)"
         Height          =   375
         Left            =   150
         TabIndex        =   15
         Top             =   825
         Value           =   -1  'True
         Width           =   4875
      End
      Begin VB.OptionButton Option2 
         Caption         =   "隨機產生的內容(慢)(&R)"
         Height          =   375
         Left            =   150
         TabIndex        =   14
         Top             =   510
         Width           =   4875
      End
      Begin VB.OptionButton Option1 
         Caption         =   "空格(&S)"
         Height          =   375
         Left            =   150
         TabIndex        =   13
         Top             =   195
         Width           =   4875
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "文件大小"
      Height          =   1365
      Left            =   30
      TabIndex        =   6
      Top             =   3720
      Width           =   4620
      Begin VB.ComboBox Combo1 
         Height          =   300
         ItemData        =   "Form1.frx":0331
         Left            =   3240
         List            =   "Form1.frx":0341
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   195
         Width           =   1305
      End
      Begin VB.TextBox Text2 
         Height          =   315
         Left            =   1500
         MaxLength       =   10
         TabIndex        =   8
         Text            =   "1.5"
         Top             =   195
         Width           =   1695
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   120
         TabIndex        =   11
         Top             =   840
         Width           =   4395
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "換算文件大小(字節)"
         Height          =   180
         Left            =   120
         TabIndex        =   10
         Top             =   570
         Width           =   1620
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "大小(0-100GB):"
         Height          =   180
         Left            =   120
         TabIndex        =   7
         Top             =   270
         Width           =   1260
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "文件保存選項"
      Height          =   3000
      Left            =   30
      TabIndex        =   1
      Top             =   690
      Width           =   4620
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   105
         MaxLength       =   255
         TabIndex        =   5
         Text            =   "LargeFile.LGF"
         Top             =   2595
         Width           =   4410
      End
      Begin VB.DirListBox Dir1 
         Height          =   1770
         Left            =   90
         TabIndex        =   3
         Top             =   570
         Width           =   4410
      End
      Begin VB.DriveListBox Drive1 
         Height          =   300
         Left            =   90
         TabIndex        =   2
         Top             =   225
         Width           =   4410
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "文件名:"
         Height          =   180
         Left            =   90
         TabIndex        =   4
         Top             =   2385
         Width           =   630
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "本工具可以幫助您創建一個制定大小的文件--不管您是要填充磁盤空間或者替換文件,都是很有用的."
      Height          =   540
      Left            =   630
      TabIndex        =   0
      Top             =   165
      Width           =   4020
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   0
      Picture         =   "Form1.frx":0354
      Top             =   105
      Width           =   480
   End
End
Attribute VB_Name = "frmLargefile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lpSize As Currency
Dim bchk As Boolean
Dim lpFilePath As String
Const MAX_FILE_SIZE = 100 * (1024 ^ 3)
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
Private Sub Combo1_Change()
On Error Resume Next
On Error Resume Next
Dim lpLength As Long
Select Case Combo1.ListIndex
Case 0
With Me.Text2
.MaxLength = 10
.Text = CStr(1.5 * 1024 * 1024 * 1024)
End With
With Me.Label5
.Alignment = 2
.BackStyle = 0
.BorderStyle = 1
.Caption = CStr(Text2.Text)
.Enabled = True
End With
lpSize = CLng(Label5.Caption)
Case 1
With Me.Text2
.MaxLength = Len(CStr(1.5 * 1024 * 1024))
.Text = CStr(1.5 * 1024 * 1024)
End With
With Me.Label5
.Alignment = 2
.BackStyle = 0
.BorderStyle = 1
.Caption = CStr(CLng(Text2.Text) * 1024)
.Enabled = True
End With
lpSize = CLng(Label5.Caption)
Case 2
With Me.Text2
.MaxLength = Len(CStr(1.5 * 1024))
.Text = CStr(1.5 * 1024)
End With
With Me.Label5
.Alignment = 2
.BackStyle = 0
.BorderStyle = 1
.Caption = CStr(CLng(Text2.Text) * 1024 * 1024)
.Enabled = True
End With
lpSize = CLng(Label5.Caption)
Case 3
With Me.Text2
.MaxLength = Len(CStr(1.5))
.Text = CStr(1.5)
End With
With Me.Label5
.Alignment = 2
.BackStyle = 0
.BorderStyle = 1
.Caption = CStr(CLng(Text2.Text) * 1024 * 1024 * 1024)
.Enabled = True
End With
lpSize = CLng(Label5.Caption)
End Select
End Sub
Private Sub Combo1_Click()
On Error Resume Next
Dim lpLength As Long
Select Case Combo1.ListIndex
Case 0
With Me.Text2
.MaxLength = 10
.Text = CStr(1.5 * 1024 * 1024 * 1024)
End With
With Me.Label5
.Alignment = 2
.BackStyle = 0
.BorderStyle = 1
.Caption = CStr(Text2.Text)
.Enabled = True
End With
lpSize = CLng(Label5.Caption)
Case 1
With Me.Text2
.MaxLength = Len(CStr(1.5 * 1024 * 1024))
.Text = CStr(1.5 * 1024 * 1024)
End With
With Me.Label5
.Alignment = 2
.BackStyle = 0
.BorderStyle = 1
.Caption = CStr(CLng(Text2.Text) * 1024)
.Enabled = True
End With
lpSize = CLng(Label5.Caption)
Case 2
With Me.Text2
.MaxLength = Len(CStr(1.5 * 1024))
.Text = CStr(1.5 * 1024)
End With
With Me.Label5
.Alignment = 2
.BackStyle = 0
.BorderStyle = 1
.Caption = CStr(CLng(Text2.Text) * 1024 * 1024)
.Enabled = True
End With
lpSize = CLng(Label5.Caption)
Case 3
With Me.Text2
.MaxLength = Len(CStr(1.5))
.Text = CStr(1.5)
End With
With Me.Label5
.Alignment = 2
.BackStyle = 0
.BorderStyle = 1
.Caption = CStr(CLng(Text2.Text) * 1024 * 1024 * 1024)
.Enabled = True
End With
lpSize = CLng(Label5.Caption)
End Select
End Sub
Private Sub Command1_Click()
'On Error GoTo ep
Drive1.Enabled = False
Dir1.Enabled = False
Text1.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
Combo1.Enabled = False
Label1.Enabled = False
Label2.Enabled = False
Label3.Enabled = False
Label4.Enabled = False
Label5.Enabled = False
Command1.Enabled = False
Option1.Enabled = False
Option2.Enabled = False
Me.Enabled = False
Option3.Enabled = False
Frame1.Enabled = False
Frame2.Enabled = False
Frame3.Enabled = False
If CDbl(Label5.Caption) > MAX_FILE_SIZE Then
MsgBox "請輸入有效的文件大小,範圍0-100GB(0Bytes - " & MAX_FILE_SIZE & "Bytes)", vbCritical, "Error"
Drive1.Enabled = True
Dir1.Enabled = True
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Combo1.Enabled = True
Label1.Enabled = True
Label2.Enabled = True
Label3.Enabled = True
Me.Enabled = True
Label4.Enabled = True
Label5.Enabled = True
Command1.Enabled = True
Option1.Enabled = True
Option2.Enabled = True
Option3.Enabled = True
Frame1.Enabled = True
Frame2.Enabled = True
Frame3.Enabled = True
Exit Sub
End If
If CDbl(Label5.Caption) < 0 Then
MsgBox "請輸入有效的文件大小,範圍0-100GB(0Bytes - " & MAX_FILE_SIZE & "Bytes)", vbCritical, "Error"
Drive1.Enabled = True
Dir1.Enabled = True
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Combo1.Enabled = True
Label1.Enabled = True
Label2.Enabled = True
Label3.Enabled = True
Label4.Enabled = True
Me.Enabled = True
Label5.Enabled = True
Command1.Enabled = True
Option1.Enabled = True
Option2.Enabled = True
Option3.Enabled = True
Frame1.Enabled = True
Frame2.Enabled = True
Frame3.Enabled = True
Exit Sub
End If
lpSize = CCur(Label5.Caption)
If lpSize = 0 Then
If Right(Dir1.path, 1) = "\" Then
lpFilePath = Dir1.path & Trim(Text1.Text)
Else
lpFilePath = Dir1.path & "\" & Trim(Text1.Text)
End If
Dim ans As Integer
If Dir(lpFilePath) <> "" Then
ans = MsgBox("文件 " & lpFilePath & " 已經存在,是否替換?", vbYesNo + vbExclamation, "Ask")
Select Case ans
Case vbYes
Kill lpFilePath
Open lpFilePath For Output As #1
Close
MsgBox "文件 " & lpFilePath & " 被成功創建", vbInformation, "Info"
Drive1.Enabled = True
Dir1.Enabled = True
Text1.Enabled = True
Text2.Enabled = True
Me.Enabled = True
Text3.Enabled = True
Combo1.Enabled = True
Label1.Enabled = True
Label2.Enabled = True
Label3.Enabled = True
Label4.Enabled = True
Label5.Enabled = True
Command1.Enabled = True
Option1.Enabled = True
Option2.Enabled = True
Option3.Enabled = True
Frame1.Enabled = True
Frame2.Enabled = True
Frame3.Enabled = True
Exit Sub
Case vbNo
MsgBox "操作已被用戶取消", vbInformation, "Info"
Drive1.Enabled = True
Dir1.Enabled = True
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Combo1.Enabled = True
Label1.Enabled = True
Label2.Enabled = True
Label3.Enabled = True
Label4.Enabled = True
Label5.Enabled = True
Command1.Enabled = True
Option1.Enabled = True
Option2.Enabled = True
Option3.Enabled = True
Frame1.Enabled = True
Frame2.Enabled = True
Me.Enabled = True
Frame3.Enabled = True
End Select
Else
Open lpFilePath For Output As #1
Close
MsgBox "文件 " & lpFilePath & " 被成功創建", vbInformation, "Info"
Drive1.Enabled = True
Dir1.Enabled = True
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Combo1.Enabled = True
Label1.Enabled = True
Label2.Enabled = True
Label3.Enabled = True
Label4.Enabled = True
Label5.Enabled = True
Command1.Enabled = True
Option1.Enabled = True
Option2.Enabled = True
Option3.Enabled = True
Frame1.Enabled = True
Frame2.Enabled = True
Frame3.Enabled = True
Me.Enabled = True
Exit Sub
End If
Else
'If File Size Is Larger Then 0.00 Bytes
'Init Progress Bar
With frmProgress.Shape1
.Left = 0
.Top = 0
.Width = 0
.Height = frmProgress.Picture1.Height
.Visible = True
.BackColor = RGB(0, 0, 255)
.BorderColor = RGB(0, 0, 255)
End With
With frmProgress.Picture1
.BackColor = RGB(255, 255, 255)
End With
'Start Creating File
If Right(Dir1.path, 1) = "\" Then
lpFilePath = Dir1.path & Trim(Text1.Text)
Else
lpFilePath = Dir1.path & "\" & Trim(Text1.Text)
End If
Dim lpFileSizeCurrent As Long
If Dir(lpFilePath) <> "" Then
ans = MsgBox("文件 " & lpFilePath & " 已經存在,是否替換?", vbYesNo + vbExclamation, "Ask")
Select Case ans
Case vbYes
Kill lpFilePath
Open lpFilePath For Output As #1
Close
frmProgress.Show
Do
Open lpFilePath For Append As #1
If Option1.Value = True Then
Print #1, "                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                "
ElseIf Option2.Value = True Then
Dim lpChr As Integer
Randomize
lpChr = Int(Rnd * 245 + 1)
Print #1, Chr(lpChr) & Chr(lpChr) & Chr(lpChr) & Chr(lpChr) & Chr(lpChr) & Chr(lpChr) & Chr(lpChr) & Chr(lpChr) & Chr(lpChr) & Chr(lpChr) & Chr(lpChr) & Chr(lpChr) & Chr(lpChr) & Chr(lpChr) & Chr(lpChr) & Chr(lpChr) & Chr(lpChr)
ElseIf Option3.Value = True Then
Print #1, Text3.Text & Text3.Text & Text3.Text & Text3.Text & Text3.Text & Text3.Text & Text3.Text & Text3.Text & Text3.Text & Text3.Text & Text3.Text & Text3.Text & Text3.Text & Text3.Text & Text3.Text & Text3.Text & Text3.Text & Text3.Text & Text3.Text
End If
Close
If frmProgress.Tag = "Stop" Then
GoTo es
End If
lpFileSizeCurrent = FileLen(lpFilePath)
Sleep 50
frmProgress.Shape1.Width = frmProgress.Picture1.Width * (lpFileSizeCurrent / lpSize)
frmProgress.Caption = "Working On File... " & CStr(lpFileSizeCurrent / lpSize) & "% Finished"
frmProgress.Caption = "Working On File... " & Round((lpFileSizeCurrent / lpSize) * 100) & "% Finished"
DoEvents
Loop Until lpFileSizeCurrent >= lpSize
MsgBox "文件 " & lpFilePath & " 被成功創建", vbInformation, "Info"
Unload frmProgress
Drive1.Enabled = True
Dir1.Enabled = True
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Combo1.Enabled = True
Label1.Enabled = True
Label2.Enabled = True
Label3.Enabled = True
Label4.Enabled = True
Label5.Enabled = True
Command1.Enabled = True
Option1.Enabled = True
Option2.Enabled = True
Option3.Enabled = True
Frame1.Enabled = True
Frame2.Enabled = True
Frame3.Enabled = True
Me.Enabled = True
Exit Sub
Case vbNo
MsgBox "操作已被用戶取消", vbInformation, "Info"
Drive1.Enabled = True
Dir1.Enabled = True
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Combo1.Enabled = True
Label1.Enabled = True
Label2.Enabled = True
Label3.Enabled = True
Label4.Enabled = True
Label5.Enabled = True
Command1.Enabled = True
Option1.Enabled = True
Option2.Enabled = True
Option3.Enabled = True
Frame1.Enabled = True
Frame2.Enabled = True
Frame3.Enabled = True
Me.Enabled = True
End Select
Else
frmProgress.Show
Open lpFilePath For Output As #1
Close
Do
Open lpFilePath For Append As #1
If Option1.Value = True Then
Print #1, "                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                "
ElseIf Option2.Value = True Then
Randomize
lpChr = Int(Rnd * 245 + 1)
Print #1, Chr(lpChr) & Chr(lpChr) & Chr(lpChr) & Chr(lpChr) & Chr(lpChr) & Chr(lpChr) & Chr(lpChr) & Chr(lpChr) & Chr(lpChr) & Chr(lpChr) & Chr(lpChr) & Chr(lpChr) & Chr(lpChr) & Chr(lpChr) & Chr(lpChr) & Chr(lpChr) & Chr(lpChr)
ElseIf Option3.Value = True Then
Print #1, Text3.Text & Text3.Text & Text3.Text & Text3.Text & Text3.Text & Text3.Text & Text3.Text & Text3.Text & Text3.Text & Text3.Text & Text3.Text & Text3.Text & Text3.Text & Text3.Text & Text3.Text & Text3.Text & Text3.Text & Text3.Text & Text3.Text
End If
Close
lpFileSizeCurrent = FileLen(lpFilePath)
If frmProgress.Tag = "Stop" Then
GoTo es
End If
Sleep 50
frmProgress.Shape1.Width = frmProgress.Picture1.Width * (lpFileSizeCurrent / lpSize)
frmProgress.Caption = "Working On File... " & CStr(lpFileSizeCurrent / lpSize) & "% Finished"
frmProgress.Caption = "Working On File... " & Left(CStr(lpFileSizeCurrent / lpSize), 25) & "% Finished"
frmProgress.Caption = "Working On File... " & Round((lpFileSizeCurrent / lpSize) * 100) & "% Finished"
DoEvents
Loop Until lpFileSizeCurrent >= lpSize
MsgBox "文件 " & lpFilePath & " 被成功創建", vbInformation, "Info"
Unload frmProgress
Drive1.Enabled = True
Dir1.Enabled = True
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Combo1.Enabled = True
Label1.Enabled = True
Label2.Enabled = True
Label3.Enabled = True
Label4.Enabled = True
Label5.Enabled = True
Command1.Enabled = True
Option1.Enabled = True
Option2.Enabled = True
Option3.Enabled = True
Frame1.Enabled = True
Frame2.Enabled = True
Frame3.Enabled = True
Me.Enabled = True
Exit Sub
End If
End If
Exit Sub
ep:
MsgBox "文件 " & lpFilePath & " 創建失敗,錯誤:" & Err.Description, vbCritical, "Error"
Close
Unload frmProgress
Drive1.Enabled = True
Dir1.Enabled = True
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Combo1.Enabled = True
Label1.Enabled = True
Label2.Enabled = True
Label3.Enabled = True
Label4.Enabled = True
Label5.Enabled = True
Command1.Enabled = True
Option1.Enabled = True
Option2.Enabled = True
Option3.Enabled = True
Frame1.Enabled = True
Frame2.Enabled = True
Frame3.Enabled = True
Me.Enabled = True
Exit Sub
es:
Unload frmProgress
MsgBox "操作已被用戶取消", vbInformation, "Info"
Drive1.Enabled = True
Dir1.Enabled = True
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Combo1.Enabled = True
Label1.Enabled = True
Label2.Enabled = True
Label3.Enabled = True
Label4.Enabled = True
Label5.Enabled = True
Command1.Enabled = True
Option1.Enabled = True
Option2.Enabled = True
Option3.Enabled = True
Frame1.Enabled = True
Frame2.Enabled = True
Frame3.Enabled = True
Me.Enabled = True
Close
End Sub
Private Sub Drive1_Change()
On Error GoTo ep
With Me.Dir1
.path = Drive1.Drive
End With
Exit Sub
ep:
Dim ans As Integer
ans = MsgBox(Err.Description, vbCritical + vbRetryCancel, "Error")
Select Case ans
Case vbRetry
DriveChange
Case Else
With Me.Drive1
.Drive = "C:"
End With
End Select
End Sub
Private Sub DriveChange()
On Error GoTo ep
With Me.Dir1
.path = Drive1.Drive
End With
Exit Sub
ep:
Dim ans As Integer
ans = MsgBox(Err.Description, vbCritical + vbRetryCancel, "Error")
Select Case ans
Case vbRetry
Drive1_Change
Case Else
With Me.Drive1
.Drive = "C:"
End With
End Select
End Sub
Private Sub Form_Load()
On Error Resume Next
Me.Combo1.ListIndex = 1
End Sub
Private Sub Option1_Click()
On Error Resume Next
With Me.Text3
.Enabled = False
.MaxLength = 2048
If 1 = 245 Then
.SetFocus
End If
End With
End Sub
Private Sub Option2_Click()
On Error Resume Next
With Me.Text3
.Enabled = False
.MaxLength = 2048
If 1 = 245 Then
.SetFocus
End If
End With
End Sub
Private Sub Option3_Click()
On Error Resume Next
With Me.Text3
.Enabled = True
.MaxLength = 2048
.SetFocus
.SelStart = 0
.SelLength = Len(.Text)
End With
End Sub
Private Sub Text2_Change()
On Error Resume Next
Select Case Combo1.ListIndex
Case 0
With Label5
.Alignment = 2
.BorderStyle = 1
.BackStyle = 0
.Caption = CStr(Text2.Text)
lpSize = CLng(.Caption)
End With
Case 1
With Label5
.Alignment = 2
.BorderStyle = 1
.BackStyle = 0
.Caption = CStr(Text2.Text * 1024)
lpSize = CLng(.Caption)
End With
Case 2
With Label5
.Alignment = 2
.BorderStyle = 1
.BackStyle = 0
.Caption = CStr(Text2.Text * 1024 * 1024)
lpSize = CLng(.Caption)
End With
Case 3
With Label5
.Alignment = 2
.BorderStyle = 1
.BackStyle = 0
.Caption = CStr(Text2.Text * 1024 * 1024 * 1024)
lpSize = CLng(.Caption)
End With
End Select
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
On Error Resume Next
Select Case Chr(KeyAscii)
Case "1"
KeyAscii = KeyAscii
Case "2"
KeyAscii = KeyAscii
Case "3"
KeyAscii = KeyAscii
Case "4"
KeyAscii = KeyAscii
Case "5"
KeyAscii = KeyAscii
Case "6"
KeyAscii = KeyAscii
Case "7"
KeyAscii = KeyAscii
Case "8"
KeyAscii = KeyAscii
Case "9"
KeyAscii = KeyAscii
Case "0"
KeyAscii = KeyAscii
Case "."
KeyAscii = KeyAscii
Case Else
If KeyAscii = vbKeyBack Then
KeyAscii = KeyAscii
Exit Sub
End If
KeyAscii = 0
End Select
End Sub
