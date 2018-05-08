VERSION 5.00
Begin VB.Form frmKill 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gray Killer - PC-DOS Workshop"
   ClientHeight    =   780
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   3930
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   780
   ScaleWidth      =   3930
   Begin 工程1.cSysTray cSysTray1 
      Left            =   1710
      Top             =   525
      _ExtentX        =   900
      _ExtentY        =   900
      InTray          =   0   'False
      TrayIcon        =   "Form1.frx":0442
      TrayTip         =   "Gray Killer - 双击还原窗口,右键单击显示菜单"
   End
   Begin VB.Timer Timer1 
      Left            =   3705
      Top             =   2655
   End
   Begin VB.Frame Frame1 
      Caption         =   "当前菜单/窗口句柄(Handle)"
      Height          =   735
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   3870
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   135
         TabIndex        =   1
         Top             =   210
         Width           =   3600
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "文件(&F)"
      Begin VB.Menu mnuMini 
         Caption         =   "最小化到系统托盘(&M)"
      End
      Begin VB.Menu b1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMonitor 
         Caption         =   "控件定点监视(&C)"
         Begin VB.Menu mnuEnable 
            Caption         =   "启用(&N)"
            Shortcut        =   {F6}
         End
         Begin VB.Menu mnuDisable 
            Caption         =   "禁用(&D)"
            Shortcut        =   {F7}
         End
      End
      Begin VB.Menu b4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "退出(&E)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "编辑(&E)"
      Begin VB.Menu mnuCopyHandle 
         Caption         =   "复制窗口句柄(&C)"
         Shortcut        =   ^C
      End
      Begin VB.Menu b7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRefresh 
         Caption         =   "刷新(&R)"
         Shortcut        =   {F5}
      End
      Begin VB.Menu b11 
         Caption         =   "-"
      End
      Begin VB.Menu mnuShowInfo 
         Caption         =   "显示当前窗口信息(&S)..."
         Shortcut        =   ^I
      End
   End
   Begin VB.Menu MnuView 
      Caption         =   "查看(&V)"
      Begin VB.Menu mnuTop 
         Caption         =   "总是在最前面(&A)"
         Checked         =   -1  'True
      End
      Begin VB.Menu b2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTrans 
         Caption         =   "本窗口透明(&T)"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "工具(&T)"
      Begin VB.Menu mnuViewInfo 
         Caption         =   "查看活动窗口信息(&V)..."
      End
      Begin VB.Menu b9 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDOpt 
         Caption         =   "选项"
         Enabled         =   0   'False
      End
      Begin VB.Menu b6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEGC 
         Caption         =   "灰色控件可用(&E)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuEGM 
         Caption         =   "灰色菜单可用(&M)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuSHC 
         Caption         =   "显示隐藏控件(&S)"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "帮助(&H)"
      Begin VB.Menu mnuShowHelp 
         Caption         =   "显示应用程序帮助(&H)..."
      End
      Begin VB.Menu b8 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "关于 Gray Killer(&A)..."
      End
   End
   Begin VB.Menu mnuTray 
      Caption         =   "Tray"
      Visible         =   0   'False
      Begin VB.Menu mnuRestore 
         Caption         =   "还原程序窗口(&R)"
      End
      Begin VB.Menu mnuInfoTray 
         Caption         =   "显示活动窗口信息(&S)"
      End
      Begin VB.Menu b5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMonT 
         Caption         =   "控件监视程序(&C)"
         Begin VB.Menu mnuET 
            Caption         =   "启用(&N)"
         End
         Begin VB.Menu mnuDT 
            Caption         =   "禁用(&D)"
         End
      End
      Begin VB.Menu b3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEnd 
         Caption         =   "退出(&E)"
      End
   End
End
Attribute VB_Name = "frmKill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim var_90 As Long
Dim var_88 As Long
Dim var_A8 As Long
Private Declare Function IsMenu Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function EnableMenuItem Lib "user32" (ByVal hMenu As Long, ByVal wIDEnableItem As Long, ByVal wEnable As Long) As Long
Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function EnumChildWindows Lib "user32" (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Private Declare Function MenuItemFromPoint Lib "user32" (ByVal hwnd As Long, ByVal hMenu As Long, ByRef ptScreen As POINTAPI) As Long
Private Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Dim lpszCaptionNew As String
Private Type POINTAPI
x As Long
y As Long
End Type
Private Const SC_MINIMIZE = &HF020&
Private Const WS_MAXIMIZEBOX = &H10000
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
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
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
Private Sub b1_Click()
On Error Resume Next
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Enabled = False
.Checked = True
End With
With Me.mnuET
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Enabled = True
.Checked = False
End With
With Me.mnuDT
.Enabled = True
.Checked = False
End With
Case False
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Enabled = True
.Checked = False
End With
With Me.mnuET
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Enabled = False
.Checked = True
End With
With Me.mnuDT
.Enabled = False
.Checked = True
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
Case True
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
End Select
End Sub
Private Sub b2_Click()
On Error Resume Next
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Enabled = False
.Checked = True
End With
With Me.mnuET
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Enabled = True
.Checked = False
End With
With Me.mnuDT
.Enabled = True
.Checked = False
End With
Case False
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Enabled = True
.Checked = False
End With
With Me.mnuET
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Enabled = False
.Checked = True
End With
With Me.mnuDT
.Enabled = False
.Checked = True
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
Case True
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
End Select
End Sub
Private Sub b3_Click()
On Error Resume Next
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Enabled = False
.Checked = True
End With
With Me.mnuET
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Enabled = True
.Checked = False
End With
With Me.mnuDT
.Enabled = True
.Checked = False
End With
Case False
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Enabled = True
.Checked = False
End With
With Me.mnuET
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Enabled = False
.Checked = True
End With
With Me.mnuDT
.Enabled = False
.Checked = True
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
Case True
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
End Select
End Sub
Private Sub b4_Click()
On Error Resume Next
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Enabled = False
.Checked = True
End With
With Me.mnuET
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Enabled = True
.Checked = False
End With
With Me.mnuDT
.Enabled = True
.Checked = False
End With
Case False
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Enabled = True
.Checked = False
End With
With Me.mnuET
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Enabled = False
.Checked = True
End With
With Me.mnuDT
.Enabled = False
.Checked = True
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
Case True
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
End Select
End Sub
Private Sub b5_Click()
On Error Resume Next
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Enabled = False
.Checked = True
End With
With Me.mnuET
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Enabled = True
.Checked = False
End With
With Me.mnuDT
.Enabled = True
.Checked = False
End With
Case False
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Enabled = True
.Checked = False
End With
With Me.mnuET
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Enabled = False
.Checked = True
End With
With Me.mnuDT
.Enabled = False
.Checked = True
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
Case True
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
End Select
End Sub

Private Sub b6_Click()
On Error Resume Next
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Enabled = False
.Checked = True
End With
With Me.mnuET
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Enabled = True
.Checked = False
End With
With Me.mnuDT
.Enabled = True
.Checked = False
End With
Case False
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Enabled = True
.Checked = False
End With
With Me.mnuET
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Enabled = False
.Checked = True
End With
With Me.mnuDT
.Enabled = False
.Checked = True
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
Case True
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
End Select
End Sub
Private Sub b7_Click()
On Error Resume Next
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Enabled = False
.Checked = True
End With
With Me.mnuET
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Enabled = True
.Checked = False
End With
With Me.mnuDT
.Enabled = True
.Checked = False
End With
Case False
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Enabled = True
.Checked = False
End With
With Me.mnuET
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Enabled = False
.Checked = True
End With
With Me.mnuDT
.Enabled = False
.Checked = True
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
Case True
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
End Select
End Sub
Private Sub b8_Click()
On Error Resume Next
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Enabled = False
.Checked = True
End With
With Me.mnuET
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Enabled = True
.Checked = False
End With
With Me.mnuDT
.Enabled = True
.Checked = False
End With
Case False
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Enabled = True
.Checked = False
End With
With Me.mnuET
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Enabled = False
.Checked = True
End With
With Me.mnuDT
.Enabled = False
.Checked = True
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
Case True
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
End Select
End Sub
Private Sub b9_Click()
On Error Resume Next
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Enabled = False
.Checked = True
End With
With Me.mnuET
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Enabled = True
.Checked = False
End With
With Me.mnuDT
.Enabled = True
.Checked = False
End With
Case False
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Enabled = True
.Checked = False
End With
With Me.mnuET
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Enabled = False
.Checked = True
End With
With Me.mnuDT
.Enabled = False
.Checked = True
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
Case True
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
End Select
End Sub
Private Sub Blank_Click()
On Error Resume Next
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Enabled = False
.Checked = True
End With
With Me.mnuET
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Enabled = True
.Checked = False
End With
With Me.mnuDT
.Enabled = True
.Checked = False
End With
Case False
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Enabled = True
.Checked = False
End With
With Me.mnuET
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Enabled = False
.Checked = True
End With
With Me.mnuDT
.Enabled = False
.Checked = True
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
Case True
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
End Select
End Sub
Private Sub cSysTray1_MouseDblClick(Button As Integer, Id As Long)
On Error Resume Next
If Button = 1 Then
With frmKill
.WindowState = 0
.Show
.Enabled = True
End With
With Me.cSysTray1
.InTray = False
.TrayTip = "Gray Killer - 双击还原窗口,右键单击显示菜单"
End With
On Error Resume Next
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Enabled = False
.Checked = True
End With
With Me.mnuET
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Enabled = True
.Checked = False
End With
With Me.mnuDT
.Enabled = True
.Checked = False
End With
Case False
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Enabled = True
.Checked = False
End With
With Me.mnuET
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Enabled = False
.Checked = True
End With
With Me.mnuDT
.Enabled = False
.Checked = True
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
Case True
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
End Select
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Enabled = False
.Checked = True
End With
With Me.mnuET
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Enabled = True
.Checked = False
End With
With Me.mnuDT
.Enabled = True
.Checked = False
End With
Case False
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Enabled = True
.Checked = False
End With
With Me.mnuET
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Enabled = False
.Checked = True
End With
With Me.mnuDT
.Enabled = False
.Checked = True
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
Case True
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
End Select
Exit Sub
Else
On Error Resume Next
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Enabled = False
.Checked = True
End With
With Me.mnuET
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Enabled = True
.Checked = False
End With
With Me.mnuDT
.Enabled = True
.Checked = False
End With
Case False
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Enabled = True
.Checked = False
End With
With Me.mnuET
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Enabled = False
.Checked = True
End With
With Me.mnuDT
.Enabled = False
.Checked = True
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
Case True
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
End Select
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Enabled = False
.Checked = True
End With
With Me.mnuET
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Enabled = True
.Checked = False
End With
With Me.mnuDT
.Enabled = True
.Checked = False
End With
Case False
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Enabled = True
.Checked = False
End With
With Me.mnuET
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Enabled = False
.Checked = True
End With
With Me.mnuDT
.Enabled = False
.Checked = True
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
Case True
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
End Select
Exit Sub
End If
End Sub
Private Sub cSysTray1_MouseDown(Button As Integer, Id As Long)
On Error Resume Next
If Button = 2 Then
PopupMenu Me.mnuTray
Else
Exit Sub
End If
End Sub
Private Sub Form_Activate()
On Error Resume Next
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Enabled = False
.Checked = True
End With
With Me.mnuET
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Enabled = True
.Checked = False
End With
With Me.mnuDT
.Enabled = True
.Checked = False
End With
Case False
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Enabled = True
.Checked = False
End With
With Me.mnuET
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Enabled = False
.Checked = True
End With
With Me.mnuDT
.Enabled = False
.Checked = True
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
Case True
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
End Select
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Enabled = False
.Checked = True
End With
With Me.mnuET
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Enabled = True
.Checked = False
End With
With Me.mnuDT
.Enabled = True
.Checked = False
End With
Case False
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Enabled = True
.Checked = False
End With
With Me.mnuET
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Enabled = False
.Checked = True
End With
With Me.mnuDT
.Enabled = False
.Checked = True
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
Case True
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
End Select
End Sub
Private Sub Form_Click()
On Error Resume Next
On Error Resume Next
Exit Sub
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Enabled = False
.Checked = True
End With
With Me.mnuET
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Enabled = True
.Checked = False
End With
With Me.mnuDT
.Enabled = True
.Checked = False
End With
Case False
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Enabled = True
.Checked = False
End With
With Me.mnuET
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Enabled = False
.Checked = True
End With
With Me.mnuDT
.Enabled = False
.Checked = True
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
Case True
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
End Select
End Sub
Private Sub Form_DblClick()
On Error Resume Next
On Error Resume Next
Exit Sub
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Enabled = False
.Checked = True
End With
With Me.mnuET
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Enabled = True
.Checked = False
End With
With Me.mnuDT
.Enabled = True
.Checked = False
End With
Case False
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Enabled = True
.Checked = False
End With
With Me.mnuET
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Enabled = False
.Checked = True
End With
With Me.mnuDT
.Enabled = False
.Checked = True
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
Case True
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
End Select
End Sub
Private Sub Form_Deactivate()
On Error Resume Next
On Error Resume Next
Exit Sub
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Enabled = False
.Checked = True
End With
With Me.mnuET
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Enabled = True
.Checked = False
End With
With Me.mnuDT
.Enabled = True
.Checked = False
End With
Case False
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Enabled = True
.Checked = False
End With
With Me.mnuET
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Enabled = False
.Checked = True
End With
With Me.mnuDT
.Enabled = False
.Checked = True
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
Case True
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
End Select
End Sub
Private Sub Form_DragDrop(Source As Control, x As Single, y As Single)
On Error Resume Next
On Error Resume Next
Exit Sub
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Enabled = False
.Checked = True
End With
With Me.mnuET
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Enabled = True
.Checked = False
End With
With Me.mnuDT
.Enabled = True
.Checked = False
End With
Case False
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Enabled = True
.Checked = False
End With
With Me.mnuET
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Enabled = False
.Checked = True
End With
With Me.mnuDT
.Enabled = False
.Checked = True
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
Case True
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
End Select
End Sub
Private Sub Form_DragOver(Source As Control, x As Single, y As Single, State As Integer)
On Error Resume Next
On Error Resume Next
Exit Sub
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Enabled = False
.Checked = True
End With
With Me.mnuET
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Enabled = True
.Checked = False
End With
With Me.mnuDT
.Enabled = True
.Checked = False
End With
Case False
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Enabled = True
.Checked = False
End With
With Me.mnuET
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Enabled = False
.Checked = True
End With
With Me.mnuDT
.Enabled = False
.Checked = True
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
Case True
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
End Select
End Sub
Private Sub Form_GotFocus()
On Error Resume Next
On Error Resume Next
Exit Sub
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Enabled = False
.Checked = True
End With
With Me.mnuET
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Enabled = True
.Checked = False
End With
With Me.mnuDT
.Enabled = True
.Checked = False
End With
Case False
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Enabled = True
.Checked = False
End With
With Me.mnuET
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Enabled = False
.Checked = True
End With
With Me.mnuDT
.Enabled = False
.Checked = True
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
Case True
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
End Select
End Sub
Private Sub Form_Initialize()
On Error Resume Next
Dim lpParam As String
lpParam = Command$
Select Case lpParam
Case "-b"
With Me
.Hide
.WindowState = 1
End With
With Me.cSysTray1
.TrayTip = "Gray Killer - 双击还原窗口,右键单击显示菜单"
.InTray = True
End With
Case "-B"
With Me
.Hide
.WindowState = 1
End With
With Me.cSysTray1
.TrayTip = "Gray Killer - 双击还原窗口,右键单击显示菜单"
.InTray = True
End With
Case "/b"
With Me
.Hide
.WindowState = 1
End With
With Me.cSysTray1
.TrayTip = "Gray Killer - 双击还原窗口,右键单击显示菜单"
.InTray = True
End With
Case "/B"
With Me
.Hide
.WindowState = 1
End With
With Me.cSysTray1
.TrayTip = "Gray Killer - 双击还原窗口,右键单击显示菜单"
.InTray = True
End With
Case "-?"
With Timer1
.Interval = 1000
.Enabled = False
End With
With Me
.Visible = False
.Hide
End With
frmParamK.Show
Case "/?"
With Timer1
.Interval = 1000
.Enabled = False
End With
With Me
.Visible = False
.Hide
End With
frmParamK.Show
End Select
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
On Error Resume Next
Exit Sub
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Enabled = False
.Checked = True
End With
With Me.mnuET
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Enabled = True
.Checked = False
End With
With Me.mnuDT
.Enabled = True
.Checked = False
End With
Case False
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Enabled = True
.Checked = False
End With
With Me.mnuET
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Enabled = False
.Checked = True
End With
With Me.mnuDT
.Enabled = False
.Checked = True
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
Case True
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
End Select
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
On Error Resume Next
On Error Resume Next
Exit Sub
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Enabled = False
.Checked = True
End With
With Me.mnuET
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Enabled = True
.Checked = False
End With
With Me.mnuDT
.Enabled = True
.Checked = False
End With
Case False
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Enabled = True
.Checked = False
End With
With Me.mnuET
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Enabled = False
.Checked = True
End With
With Me.mnuDT
.Enabled = False
.Checked = True
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
Case True
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
End Select
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
On Error Resume Next
Exit Sub
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Enabled = False
.Checked = True
End With
With Me.mnuET
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Enabled = True
.Checked = False
End With
With Me.mnuDT
.Enabled = True
.Checked = False
End With
Case False
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Enabled = True
.Checked = False
End With
With Me.mnuET
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Enabled = False
.Checked = True
End With
With Me.mnuDT
.Enabled = False
.Checked = True
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
Case True
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
End Select
End Sub
Private Sub Form_LinkClose()
On Error Resume Next
On Error Resume Next
Exit Sub
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Enabled = False
.Checked = True
End With
With Me.mnuET
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Enabled = True
.Checked = False
End With
With Me.mnuDT
.Enabled = True
.Checked = False
End With
Case False
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Enabled = True
.Checked = False
End With
With Me.mnuET
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Enabled = False
.Checked = True
End With
With Me.mnuDT
.Enabled = False
.Checked = True
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
Case True
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
End Select
End Sub
Private Sub Form_LinkError(LinkErr As Integer)
On Error Resume Next
On Error Resume Next
Exit Sub
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Enabled = False
.Checked = True
End With
With Me.mnuET
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Enabled = True
.Checked = False
End With
With Me.mnuDT
.Enabled = True
.Checked = False
End With
Case False
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Enabled = True
.Checked = False
End With
With Me.mnuET
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Enabled = False
.Checked = True
End With
With Me.mnuDT
.Enabled = False
.Checked = True
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
Case True
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
End Select
End Sub
Private Sub Form_LinkExecute(CmdStr As String, Cancel As Integer)
On Error Resume Next
On Error Resume Next
Exit Sub
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Enabled = False
.Checked = True
End With
With Me.mnuET
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Enabled = True
.Checked = False
End With
With Me.mnuDT
.Enabled = True
.Checked = False
End With
Case False
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Enabled = True
.Checked = False
End With
With Me.mnuET
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Enabled = False
.Checked = True
End With
With Me.mnuDT
.Enabled = False
.Checked = True
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
Case True
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
End Select
End Sub
Private Sub Form_LinkOpen(Cancel As Integer)
On Error Resume Next
On Error Resume Next
Exit Sub
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Enabled = False
.Checked = True
End With
With Me.mnuET
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Enabled = True
.Checked = False
End With
With Me.mnuDT
.Enabled = True
.Checked = False
End With
Case False
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Enabled = True
.Checked = False
End With
With Me.mnuET
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Enabled = False
.Checked = True
End With
With Me.mnuDT
.Enabled = False
.Checked = True
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
Case True
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
End Select
End Sub
Private Sub Form_Load()
On Error Resume Next
With Me
.Left = 0
.Top = 0
.Enabled = True
End With
With Me.Timer1
.Interval = 1000
.Enabled = True
End With
With Me.mnuTop
.Checked = True
.Enabled = True
End With
With Me.mnuTrans
.Checked = False
.Enabled = True
End With
With Me.mnuEnable
.Enabled = True
.Checked = True
End With
With Me.mnuEGC
.Checked = True
End With
With Me.mnuEGM
.Checked = True
End With
With Me.mnuSHC
.Checked = False
End With
On Error Resume Next
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Enabled = False
.Checked = True
End With
With Me.mnuET
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Enabled = True
.Checked = False
End With
With Me.mnuDT
.Enabled = True
.Checked = False
End With
Case False
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Enabled = True
.Checked = False
End With
With Me.mnuET
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Enabled = False
.Checked = True
End With
With Me.mnuDT
.Enabled = False
.Checked = True
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
Case True
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
End Select
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Enabled = False
.Checked = True
End With
With Me.mnuET
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Enabled = True
.Checked = False
End With
With Me.mnuDT
.Enabled = True
.Checked = False
End With
Case False
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Enabled = True
.Checked = False
End With
With Me.mnuET
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Enabled = False
.Checked = True
End With
With Me.mnuDT
.Enabled = False
.Checked = True
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
Case True
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
End Select
End Sub
Private Sub Form_LostFocus()
On Error Resume Next
On Error Resume Next
Exit Sub
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Enabled = False
.Checked = True
End With
With Me.mnuET
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Enabled = True
.Checked = False
End With
With Me.mnuDT
.Enabled = True
.Checked = False
End With
Case False
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Enabled = True
.Checked = False
End With
With Me.mnuET
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Enabled = False
.Checked = True
End With
With Me.mnuDT
.Enabled = False
.Checked = True
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
Case True
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
End Select
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
On Error Resume Next
Exit Sub
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Enabled = False
.Checked = True
End With
With Me.mnuET
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Enabled = True
.Checked = False
End With
With Me.mnuDT
.Enabled = True
.Checked = False
End With
Case False
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Enabled = True
.Checked = False
End With
With Me.mnuET
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Enabled = False
.Checked = True
End With
With Me.mnuDT
.Enabled = False
.Checked = True
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
Case True
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
End Select
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
On Error Resume Next
Exit Sub
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Enabled = False
.Checked = True
End With
With Me.mnuET
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Enabled = True
.Checked = False
End With
With Me.mnuDT
.Enabled = True
.Checked = False
End With
Case False
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Enabled = True
.Checked = False
End With
With Me.mnuET
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Enabled = False
.Checked = True
End With
With Me.mnuDT
.Enabled = False
.Checked = True
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
Case True
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
End Select
End Sub
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
On Error Resume Next
Exit Sub
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Enabled = False
.Checked = True
End With
With Me.mnuET
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Enabled = True
.Checked = False
End With
With Me.mnuDT
.Enabled = True
.Checked = False
End With
Case False
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Enabled = True
.Checked = False
End With
With Me.mnuET
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Enabled = False
.Checked = True
End With
With Me.mnuDT
.Enabled = False
.Checked = True
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
Case True
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
End Select
End Sub
Private Sub Form_OLECompleteDrag(Effect As Long)
On Error Resume Next
On Error Resume Next
Exit Sub
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Enabled = False
.Checked = True
End With
With Me.mnuET
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Enabled = True
.Checked = False
End With
With Me.mnuDT
.Enabled = True
.Checked = False
End With
Case False
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Enabled = True
.Checked = False
End With
With Me.mnuET
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Enabled = False
.Checked = True
End With
With Me.mnuDT
.Enabled = False
.Checked = True
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
Case True
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
End Select
End Sub
Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
On Error Resume Next
Exit Sub
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Enabled = False
.Checked = True
End With
With Me.mnuET
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Enabled = True
.Checked = False
End With
With Me.mnuDT
.Enabled = True
.Checked = False
End With
Case False
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Enabled = True
.Checked = False
End With
With Me.mnuET
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Enabled = False
.Checked = True
End With
With Me.mnuDT
.Enabled = False
.Checked = True
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
Case True
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
End Select
End Sub
Private Sub Form_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
On Error Resume Next
On Error Resume Next
Exit Sub
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Enabled = False
.Checked = True
End With
With Me.mnuET
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Enabled = True
.Checked = False
End With
With Me.mnuDT
.Enabled = True
.Checked = False
End With
Case False
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Enabled = True
.Checked = False
End With
With Me.mnuET
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Enabled = False
.Checked = True
End With
With Me.mnuDT
.Enabled = False
.Checked = True
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
Case True
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
End Select
End Sub
Private Sub Form_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
On Error Resume Next
On Error Resume Next
Exit Sub
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Enabled = False
.Checked = True
End With
With Me.mnuET
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Enabled = True
.Checked = False
End With
With Me.mnuDT
.Enabled = True
.Checked = False
End With
Case False
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Enabled = True
.Checked = False
End With
With Me.mnuET
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Enabled = False
.Checked = True
End With
With Me.mnuDT
.Enabled = False
.Checked = True
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
Case True
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
End Select
End Sub
Private Sub Form_OLESetData(Data As DataObject, DataFormat As Integer)
On Error Resume Next
On Error Resume Next
Exit Sub
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Enabled = False
.Checked = True
End With
With Me.mnuET
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Enabled = True
.Checked = False
End With
With Me.mnuDT
.Enabled = True
.Checked = False
End With
Case False
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Enabled = True
.Checked = False
End With
With Me.mnuET
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Enabled = False
.Checked = True
End With
With Me.mnuDT
.Enabled = False
.Checked = True
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
Case True
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
End Select
End Sub
Private Sub Form_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
On Error Resume Next
On Error Resume Next
Exit Sub
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Enabled = False
.Checked = True
End With
With Me.mnuET
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Enabled = True
.Checked = False
End With
With Me.mnuDT
.Enabled = True
.Checked = False
End With
Case False
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Enabled = True
.Checked = False
End With
With Me.mnuET
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Enabled = False
.Checked = True
End With
With Me.mnuDT
.Enabled = False
.Checked = True
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
Case True
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
End Select
End Sub
Private Sub Form_Paint()
On Error Resume Next
On Error Resume Next
Exit Sub
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Enabled = False
.Checked = True
End With
With Me.mnuET
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Enabled = True
.Checked = False
End With
With Me.mnuDT
.Enabled = True
.Checked = False
End With
Case False
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Enabled = True
.Checked = False
End With
With Me.mnuET
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Enabled = False
.Checked = True
End With
With Me.mnuDT
.Enabled = False
.Checked = True
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
Case True
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
End Select
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
With Me.cSysTray1
.TrayTip = "Gray Killer - 双击还原窗口,右键单击显示菜单"
.InTray = False
End With
End
End Sub
Private Sub Form_Resize()
On Error Resume Next
If Me.WindowState = 1 Then
With Me
.Hide
End With
With Me.cSysTray1
.TrayTip = "Gray Killer - 双击还原窗口,右键单击显示菜单"
.InTray = True
End With
Exit Sub
Else
With Me.cSysTray1
.TrayTip = "Gray Killer - 双击还原窗口,右键单击显示菜单"
.InTray = False
End With
Exit Sub
End If
End Sub
Private Sub Form_Terminate()
On Error Resume Next
End
On Error Resume Next
Exit Sub
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Enabled = False
.Checked = True
End With
With Me.mnuET
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Enabled = True
.Checked = False
End With
With Me.mnuDT
.Enabled = True
.Checked = False
End With
Case False
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Enabled = True
.Checked = False
End With
With Me.mnuET
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Enabled = False
.Checked = True
End With
With Me.mnuDT
.Enabled = False
.Checked = True
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
Case True
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
End Select
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
With Me.cSysTray1
.TrayTip = "Gray Killer - 双击还原窗口,右键单击显示菜单"
.InTray = False
End With
End
End Sub
Private Sub Frame1_Click()
On Error Resume Next
On Error Resume Next
Exit Sub
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Enabled = False
.Checked = True
End With
With Me.mnuET
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Enabled = True
.Checked = False
End With
With Me.mnuDT
.Enabled = True
.Checked = False
End With
Case False
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Enabled = True
.Checked = False
End With
With Me.mnuET
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Enabled = False
.Checked = True
End With
With Me.mnuDT
.Enabled = False
.Checked = True
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
Case True
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
End Select
End Sub
Private Sub Frame1_DblClick()
On Error Resume Next
On Error Resume Next
Exit Sub
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Enabled = False
.Checked = True
End With
With Me.mnuET
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Enabled = True
.Checked = False
End With
With Me.mnuDT
.Enabled = True
.Checked = False
End With
Case False
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Enabled = True
.Checked = False
End With
With Me.mnuET
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Enabled = False
.Checked = True
End With
With Me.mnuDT
.Enabled = False
.Checked = True
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
Case True
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
End Select
End Sub
Private Sub Frame1_DragDrop(Source As Control, x As Single, y As Single)
On Error Resume Next
On Error Resume Next
Exit Sub
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Enabled = False
.Checked = True
End With
With Me.mnuET
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Enabled = True
.Checked = False
End With
With Me.mnuDT
.Enabled = True
.Checked = False
End With
Case False
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Enabled = True
.Checked = False
End With
With Me.mnuET
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Enabled = False
.Checked = True
End With
With Me.mnuDT
.Enabled = False
.Checked = True
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
Case True
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
End Select
End Sub
Private Sub Frame1_DragOver(Source As Control, x As Single, y As Single, State As Integer)
On Error Resume Next
On Error Resume Next
Exit Sub
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Enabled = False
.Checked = True
End With
With Me.mnuET
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Enabled = True
.Checked = False
End With
With Me.mnuDT
.Enabled = True
.Checked = False
End With
Case False
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Enabled = True
.Checked = False
End With
With Me.mnuET
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Enabled = False
.Checked = True
End With
With Me.mnuDT
.Enabled = False
.Checked = True
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
Case True
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
End Select
End Sub
Private Sub Frame1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
On Error Resume Next
Exit Sub
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Enabled = False
.Checked = True
End With
With Me.mnuET
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Enabled = True
.Checked = False
End With
With Me.mnuDT
.Enabled = True
.Checked = False
End With
Case False
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Enabled = True
.Checked = False
End With
With Me.mnuET
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Enabled = False
.Checked = True
End With
With Me.mnuDT
.Enabled = False
.Checked = True
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
Case True
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
End Select
End Sub
Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
On Error Resume Next
Exit Sub
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Enabled = False
.Checked = True
End With
With Me.mnuET
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Enabled = True
.Checked = False
End With
With Me.mnuDT
.Enabled = True
.Checked = False
End With
Case False
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Enabled = True
.Checked = False
End With
With Me.mnuET
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Enabled = False
.Checked = True
End With
With Me.mnuDT
.Enabled = False
.Checked = True
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
Case True
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
End Select
End Sub
Private Sub Frame1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
On Error Resume Next
Exit Sub
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Enabled = False
.Checked = True
End With
With Me.mnuET
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Enabled = True
.Checked = False
End With
With Me.mnuDT
.Enabled = True
.Checked = False
End With
Case False
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Enabled = True
.Checked = False
End With
With Me.mnuET
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Enabled = False
.Checked = True
End With
With Me.mnuDT
.Enabled = False
.Checked = True
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
Case True
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
End Select
End Sub
Private Sub Frame1_OLECompleteDrag(Effect As Long)
On Error Resume Next
On Error Resume Next
Exit Sub
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Enabled = False
.Checked = True
End With
With Me.mnuET
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Enabled = True
.Checked = False
End With
With Me.mnuDT
.Enabled = True
.Checked = False
End With
Case False
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Enabled = True
.Checked = False
End With
With Me.mnuET
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Enabled = False
.Checked = True
End With
With Me.mnuDT
.Enabled = False
.Checked = True
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
Case True
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
End Select
End Sub
Private Sub Frame1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
On Error Resume Next
Exit Sub
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Enabled = False
.Checked = True
End With
With Me.mnuET
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Enabled = True
.Checked = False
End With
With Me.mnuDT
.Enabled = True
.Checked = False
End With
Case False
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Enabled = True
.Checked = False
End With
With Me.mnuET
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Enabled = False
.Checked = True
End With
With Me.mnuDT
.Enabled = False
.Checked = True
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
Case True
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
End Select
End Sub
Private Sub Frame1_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
On Error Resume Next
On Error Resume Next
Exit Sub
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Enabled = False
.Checked = True
End With
With Me.mnuET
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Enabled = True
.Checked = False
End With
With Me.mnuDT
.Enabled = True
.Checked = False
End With
Case False
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Enabled = True
.Checked = False
End With
With Me.mnuET
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Enabled = False
.Checked = True
End With
With Me.mnuDT
.Enabled = False
.Checked = True
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
Case True
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
End Select
End Sub
Private Sub Frame1_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
On Error Resume Next
On Error Resume Next
Exit Sub
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Enabled = False
.Checked = True
End With
With Me.mnuET
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Enabled = True
.Checked = False
End With
With Me.mnuDT
.Enabled = True
.Checked = False
End With
Case False
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Enabled = True
.Checked = False
End With
With Me.mnuET
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Enabled = False
.Checked = True
End With
With Me.mnuDT
.Enabled = False
.Checked = True
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
Case True
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
End Select
End Sub
Private Sub Frame1_OLESetData(Data As DataObject, DataFormat As Integer)
On Error Resume Next
On Error Resume Next
Exit Sub
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Enabled = False
.Checked = True
End With
With Me.mnuET
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Enabled = True
.Checked = False
End With
With Me.mnuDT
.Enabled = True
.Checked = False
End With
Case False
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Enabled = True
.Checked = False
End With
With Me.mnuET
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Enabled = False
.Checked = True
End With
With Me.mnuDT
.Enabled = False
.Checked = True
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
Case True
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
End Select
End Sub
Private Sub Frame1_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
On Error Resume Next
On Error Resume Next
Exit Sub
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Enabled = False
.Checked = True
End With
With Me.mnuET
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Enabled = True
.Checked = False
End With
With Me.mnuDT
.Enabled = True
.Checked = False
End With
Case False
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Enabled = True
.Checked = False
End With
With Me.mnuET
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Enabled = False
.Checked = True
End With
With Me.mnuDT
.Enabled = False
.Checked = True
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
Case True
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
End Select
End Sub
Private Sub Label1_Change()
On Error Resume Next
Exit Sub
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Enabled = False
.Checked = True
End With
With Me.mnuET
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Enabled = True
.Checked = False
End With
With Me.mnuDT
.Enabled = True
.Checked = False
End With
Case False
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Enabled = True
.Checked = False
End With
With Me.mnuET
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Enabled = False
.Checked = True
End With
With Me.mnuDT
.Enabled = False
.Checked = True
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
Case True
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
End Select
End Sub
Private Sub Label1_Click()
On Error Resume Next
On Error Resume Next
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Dim rtn As Long
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
With Me
.Height = 1455
.Width = 4020
End With
With Me.Timer1
.Enabled = False
.Interval = 1000
End With
If Label1.Caption = "" Then
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Enabled = False
.Checked = True
End With
With Me.mnuET
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Enabled = True
.Checked = False
End With
With Me.mnuDT
.Enabled = True
.Checked = False
End With
Case False
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Enabled = True
.Checked = False
End With
With Me.mnuET
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Enabled = False
.Checked = True
End With
With Me.mnuDT
.Enabled = False
.Checked = True
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
Case True
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
End Select
Exit Sub
End If
If Label1.Caption = "0" Then
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Enabled = False
.Checked = True
End With
With Me.mnuET
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Enabled = True
.Checked = False
End With
With Me.mnuDT
.Enabled = True
.Checked = False
End With
Case False
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Enabled = True
.Checked = False
End With
With Me.mnuET
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Enabled = False
.Checked = True
End With
With Me.mnuDT
.Enabled = False
.Checked = True
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
Case True
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
End Select
Exit Sub
End If
With lpWindow
.hWindow = Label1.Caption
.hWindowDC = GetWindowDC(.hWindow)
.lpszCaption = GetPassword(.hWindow)
.hThreadProcessID = GetWindowThreadProcessId(.hWindow, 0)
.hThreadProcessID = OpenProcess(PROCESS_ALL_ACCESS, False, .hThreadProcessID)
GetModuleFileName Null, .lpszThreadProcessName, 1024
GetProcessName .hThreadProcessID, .lpszThreadProcessName, .lpszThreadProcessPath
.hThreadProcessID = OpenProcess(PROCESS_ALL_ACCESS, False, .hThreadProcessID)
GetModuleFileName .hThreadProcessID, .lpszThreadProcessName, 256
.lpszClassName = GetWindowClassName(.hWindow)
Dim lpszClsName As String * 256
GetClassName .hWindow, lpszClsName, 256
.lpszClassName = Trim(lpszClsName)
End With
With lpWindow
.hThreadProcessID = GetWindowThreadProcessId(Me.Label1.Caption, 0)
.hThreadProcess = OpenProcess(PROCESS_ALL_ACCESS, False, .hThreadProcessID)
GetModuleFileName .hThreadProcess, .lpszThreadProcessName, 1024
GetProcessName .hThreadProcessID, .lpszThreadProcessPath, .lpszThreadProcessPath
End With
With lpWindow
GetWindowThreadProcessId Me.Label1.Caption, .hThreadProcessID
Debug.Print .hThreadProcessID
.hThreadProcess = OpenProcess(PROCESS_ALL_ACCESS, False, .hThreadProcessID)
GetModuleFileName .hThreadProcess, .lpszThreadProcessName, 1024
GetProcessName .hThreadProcessID, .lpszExe, .lpszPath
End With
With lpWindow
GetWindowThreadProcessId Me.Label1.Caption, .hThreadProcessID
Debug.Print .hThreadProcessID
.hThreadProcess = OpenProcess(PROCESS_ALL_ACCESS, False, .hThreadProcessID)
GetModuleFileName .hThreadProcess, .lpszThreadProcessName, 1024
GetProcessName .hThreadProcessID, .lpszExe, .lpszPath
End With
With lpWindow
.lpszClassName = Left$(.lpszClassName, InStr(1, .lpszClassName, Chr$(0)) - 1)
End With
With lpWindow
Dim lpStr As String
lpStr = "当前窗口信息:" & vbCrLf & vbCrLf & "句柄:" & vbCrLf & .hWindow & vbCrLf & "设备上下文句柄:" & vbCrLf & .hWindowDC & vbCrLf & "标题:" & vbCrLf & .lpszCaption & vbCrLf & "类名:" & vbCrLf & .lpszClassName
lpStr = lpStr & vbCrLf & "隶属进程ID:" & vbCrLf & .hThreadProcessID & vbCrLf & "隶属进程可执行文件名:" & vbCrLf & .lpszExe
MsgBox lpStr, vbInformation, "Info"
End With
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Enabled = False
.Checked = True
End With
With Me.mnuET
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Enabled = True
.Checked = False
End With
With Me.mnuDT
.Enabled = True
.Checked = False
End With
Case False
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Enabled = True
.Checked = False
End With
With Me.mnuET
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Enabled = False
.Checked = True
End With
With Me.mnuDT
.Enabled = False
.Checked = True
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
Case True
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
End Select
End Sub
Private Sub Label1_DblClick()
On Error Resume Next
Exit Sub
On Error Resume Next
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Enabled = False
.Checked = True
End With
With Me.mnuET
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Enabled = True
.Checked = False
End With
With Me.mnuDT
.Enabled = True
.Checked = False
End With
Case False
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Enabled = True
.Checked = False
End With
With Me.mnuET
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Enabled = False
.Checked = True
End With
With Me.mnuDT
.Enabled = False
.Checked = True
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
Case True
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
End Select
End Sub
Private Sub Label1_DragDrop(Source As Control, x As Single, y As Single)
On Error Resume Next
On Error Resume Next
Exit Sub
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Enabled = False
.Checked = True
End With
With Me.mnuET
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Enabled = True
.Checked = False
End With
With Me.mnuDT
.Enabled = True
.Checked = False
End With
Case False
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Enabled = True
.Checked = False
End With
With Me.mnuET
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Enabled = False
.Checked = True
End With
With Me.mnuDT
.Enabled = False
.Checked = True
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
Case True
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
End Select
End Sub
Private Sub Label1_DragOver(Source As Control, x As Single, y As Single, State As Integer)
On Error Resume Next
On Error Resume Next
Exit Sub
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Enabled = False
.Checked = True
End With
With Me.mnuET
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Enabled = True
.Checked = False
End With
With Me.mnuDT
.Enabled = True
.Checked = False
End With
Case False
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Enabled = True
.Checked = False
End With
With Me.mnuET
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Enabled = False
.Checked = True
End With
With Me.mnuDT
.Enabled = False
.Checked = True
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
Case True
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
End Select
End Sub
Private Sub Label1_LinkClose()
On Error Resume Next
On Error Resume Next
Exit Sub
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Enabled = False
.Checked = True
End With
With Me.mnuET
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Enabled = True
.Checked = False
End With
With Me.mnuDT
.Enabled = True
.Checked = False
End With
Case False
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Enabled = True
.Checked = False
End With
With Me.mnuET
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Enabled = False
.Checked = True
End With
With Me.mnuDT
.Enabled = False
.Checked = True
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
Case True
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
End Select
End Sub
Private Sub Label1_LinkError(LinkErr As Integer)
On Error Resume Next
On Error Resume Next
Exit Sub
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Enabled = False
.Checked = True
End With
With Me.mnuET
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Enabled = True
.Checked = False
End With
With Me.mnuDT
.Enabled = True
.Checked = False
End With
Case False
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Enabled = True
.Checked = False
End With
With Me.mnuET
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Enabled = False
.Checked = True
End With
With Me.mnuDT
.Enabled = False
.Checked = True
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
Case True
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
End Select
End Sub
Private Sub Label1_LinkNotify()
On Error Resume Next
On Error Resume Next
Exit Sub
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Enabled = False
.Checked = True
End With
With Me.mnuET
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Enabled = True
.Checked = False
End With
With Me.mnuDT
.Enabled = True
.Checked = False
End With
Case False
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Enabled = True
.Checked = False
End With
With Me.mnuET
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Enabled = False
.Checked = True
End With
With Me.mnuDT
.Enabled = False
.Checked = True
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
Case True
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
End Select
End Sub
Private Sub Label1_LinkOpen(Cancel As Integer)
On Error Resume Next
On Error Resume Next
Exit Sub
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Enabled = False
.Checked = True
End With
With Me.mnuET
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Enabled = True
.Checked = False
End With
With Me.mnuDT
.Enabled = True
.Checked = False
End With
Case False
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Enabled = True
.Checked = False
End With
With Me.mnuET
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Enabled = False
.Checked = True
End With
With Me.mnuDT
.Enabled = False
.Checked = True
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
Case True
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
End Select
End Sub
Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
On Error Resume Next
Exit Sub
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Enabled = False
.Checked = True
End With
With Me.mnuET
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Enabled = True
.Checked = False
End With
With Me.mnuDT
.Enabled = True
.Checked = False
End With
Case False
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Enabled = True
.Checked = False
End With
With Me.mnuET
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Enabled = False
.Checked = True
End With
With Me.mnuDT
.Enabled = False
.Checked = True
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
Case True
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
End Select
End Sub
Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
On Error Resume Next
Exit Sub
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Enabled = False
.Checked = True
End With
With Me.mnuET
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Enabled = True
.Checked = False
End With
With Me.mnuDT
.Enabled = True
.Checked = False
End With
Case False
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Enabled = True
.Checked = False
End With
With Me.mnuET
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Enabled = False
.Checked = True
End With
With Me.mnuDT
.Enabled = False
.Checked = True
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
Case True
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
End Select
End Sub
Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
On Error Resume Next
Exit Sub
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Enabled = False
.Checked = True
End With
With Me.mnuET
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Enabled = True
.Checked = False
End With
With Me.mnuDT
.Enabled = True
.Checked = False
End With
Case False
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Enabled = True
.Checked = False
End With
With Me.mnuET
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Enabled = False
.Checked = True
End With
With Me.mnuDT
.Enabled = False
.Checked = True
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
Case True
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
End Select
End Sub
Private Sub Label1_OLECompleteDrag(Effect As Long)
On Error Resume Next
On Error Resume Next
Exit Sub
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Enabled = False
.Checked = True
End With
With Me.mnuET
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Enabled = True
.Checked = False
End With
With Me.mnuDT
.Enabled = True
.Checked = False
End With
Case False
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Enabled = True
.Checked = False
End With
With Me.mnuET
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Enabled = False
.Checked = True
End With
With Me.mnuDT
.Enabled = False
.Checked = True
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
Case True
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
End Select
End Sub
Private Sub Label1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
On Error Resume Next
Exit Sub
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Enabled = False
.Checked = True
End With
With Me.mnuET
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Enabled = True
.Checked = False
End With
With Me.mnuDT
.Enabled = True
.Checked = False
End With
Case False
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Enabled = True
.Checked = False
End With
With Me.mnuET
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Enabled = False
.Checked = True
End With
With Me.mnuDT
.Enabled = False
.Checked = True
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
Case True
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
End Select
End Sub
Private Sub Label1_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
On Error Resume Next
On Error Resume Next
Exit Sub
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Enabled = False
.Checked = True
End With
With Me.mnuET
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Enabled = True
.Checked = False
End With
With Me.mnuDT
.Enabled = True
.Checked = False
End With
Case False
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Enabled = True
.Checked = False
End With
With Me.mnuET
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Enabled = False
.Checked = True
End With
With Me.mnuDT
.Enabled = False
.Checked = True
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
Case True
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
End Select
End Sub
Private Sub Label1_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
On Error Resume Next
On Error Resume Next
Exit Sub
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Enabled = False
.Checked = True
End With
With Me.mnuET
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Enabled = True
.Checked = False
End With
With Me.mnuDT
.Enabled = True
.Checked = False
End With
Case False
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Enabled = True
.Checked = False
End With
With Me.mnuET
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Enabled = False
.Checked = True
End With
With Me.mnuDT
.Enabled = False
.Checked = True
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
Case True
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
End Select
End Sub
Private Sub Label1_OLESetData(Data As DataObject, DataFormat As Integer)
On Error Resume Next
On Error Resume Next
Exit Sub
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Enabled = False
.Checked = True
End With
With Me.mnuET
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Enabled = True
.Checked = False
End With
With Me.mnuDT
.Enabled = True
.Checked = False
End With
Case False
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Enabled = True
.Checked = False
End With
With Me.mnuET
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Enabled = False
.Checked = True
End With
With Me.mnuDT
.Enabled = False
.Checked = True
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
Case True
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
End Select
End Sub
Private Sub Label1_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
On Error Resume Next
On Error Resume Next
Exit Sub
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Enabled = False
.Checked = True
End With
With Me.mnuET
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Enabled = True
.Checked = False
End With
With Me.mnuDT
.Enabled = True
.Checked = False
End With
Case False
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Enabled = True
.Checked = False
End With
With Me.mnuET
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Enabled = False
.Checked = True
End With
With Me.mnuDT
.Enabled = False
.Checked = True
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
Case True
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
End Select
End Sub
Private Sub mnuAbout_Click()
On Error Resume Next
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
AK.Show 1
End Sub
Private Sub mnuCopyHandle_Click()
On Error GoTo ep
With Me.Timer1
.Enabled = False
.Interval = 1000
End With
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Dim rtn As Long
SetWindowPos Me.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
If Label1.Caption = "0" Then
MsgBox "没有选择前景窗口!", vbCritical, "Error"
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Enabled = False
.Checked = True
End With
With Me.mnuET
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Enabled = True
.Checked = False
End With
With Me.mnuDT
.Enabled = True
.Checked = False
End With
Case False
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Enabled = True
.Checked = False
End With
With Me.mnuET
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Enabled = False
.Checked = True
End With
With Me.mnuDT
.Enabled = False
.Checked = True
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
Case True
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
End Select
Exit Sub
End If
If Label1.Caption = "" Then
MsgBox "没有选择前景窗口!", vbCritical, "Error"
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Enabled = False
.Checked = True
End With
With Me.mnuET
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Enabled = True
.Checked = False
End With
With Me.mnuDT
.Enabled = True
.Checked = False
End With
Case False
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Enabled = True
.Checked = False
End With
With Me.mnuET
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Enabled = False
.Checked = True
End With
With Me.mnuDT
.Enabled = False
.Checked = True
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
Case True
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
End Select
Exit Sub
End If
Clipboard.SetText Me.Label1.Caption
MsgBox "复制句柄成功!", vbExclamation, "Info"
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Enabled = False
.Checked = True
End With
With Me.mnuET
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Enabled = True
.Checked = False
End With
With Me.mnuDT
.Enabled = True
.Checked = False
End With
Case False
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Enabled = True
.Checked = False
End With
With Me.mnuET
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Enabled = False
.Checked = True
End With
With Me.mnuDT
.Enabled = False
.Checked = True
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
Case True
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
End Select
Exit Sub
ep:
With Timer1
.Enabled = False
.Interval = 1000
End With
SetWindowPos Me.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
MsgBox Err.Description, vbCritical, "Error"
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Enabled = False
.Checked = True
End With
With Me.mnuET
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Enabled = True
.Checked = False
End With
With Me.mnuDT
.Enabled = True
.Checked = False
End With
Case False
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Enabled = True
.Checked = False
End With
With Me.mnuET
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Enabled = False
.Checked = True
End With
With Me.mnuDT
.Enabled = False
.Checked = True
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
Case True
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
End Select
End Sub
Private Sub mnuDisable_Click()
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Enabled = True
.Checked = False
End With
With Me.mnuET
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Enabled = False
.Checked = True
End With
With Me.mnuDT
.Enabled = False
.Checked = True
End With
On Error Resume Next
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Dim rtn As Long
Select Case mnuTop.Checked
Case False
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
Case True
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
End Select
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Enabled = False
.Checked = True
End With
With Me.mnuET
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Enabled = True
.Checked = False
End With
With Me.mnuDT
.Enabled = True
.Checked = False
End With
Case False
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Enabled = True
.Checked = False
End With
With Me.mnuET
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Enabled = False
.Checked = True
End With
With Me.mnuDT
.Enabled = False
.Checked = True
End With
End Select
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Enabled = False
.Checked = True
End With
With Me.mnuET
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Enabled = True
.Checked = False
End With
With Me.mnuDT
.Enabled = True
.Checked = False
End With
Case False
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Enabled = True
.Checked = False
End With
With Me.mnuET
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Enabled = False
.Checked = True
End With
With Me.mnuDT
.Enabled = False
.Checked = True
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
Case True
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
End Select
End Sub
Private Sub mnuDOpt_Click()
On Error Resume Next
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Enabled = False
.Checked = True
End With
With Me.mnuET
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Enabled = True
.Checked = False
End With
With Me.mnuDT
.Enabled = True
.Checked = False
End With
Case False
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Enabled = True
.Checked = False
End With
With Me.mnuET
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Enabled = False
.Checked = True
End With
With Me.mnuDT
.Enabled = False
.Checked = True
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
Case True
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
End Select
End Sub
Private Sub mnuDT_Click()
On Error Resume Next
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Enabled = True
.Checked = False
End With
With Me.mnuET
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Enabled = False
.Checked = True
End With
With Me.mnuDT
.Enabled = False
.Checked = True
End With
On Error Resume Next
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Dim rtn As Long
Select Case mnuTop.Checked
Case False
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
Case True
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
End Select
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Enabled = False
.Checked = True
End With
With Me.mnuET
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Enabled = True
.Checked = False
End With
With Me.mnuDT
.Enabled = True
.Checked = False
End With
Case False
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Enabled = True
.Checked = False
End With
With Me.mnuET
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Enabled = False
.Checked = True
End With
With Me.mnuDT
.Enabled = False
.Checked = True
End With
End Select
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Enabled = False
.Checked = True
End With
With Me.mnuET
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Enabled = True
.Checked = False
End With
With Me.mnuDT
.Enabled = True
.Checked = False
End With
Case False
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Enabled = True
.Checked = False
End With
With Me.mnuET
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Enabled = False
.Checked = True
End With
With Me.mnuDT
.Enabled = False
.Checked = True
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
Case True
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
End Select
End Sub
Private Sub mnuEGC_Click()
On Error Resume Next
With mnuEGC
Select Case .Checked
Case True
.Checked = False
Case False
.Checked = True
End Select
End With
End Sub
Private Sub mnuEGM_Click()
On Error Resume Next
With mnuEGM
Select Case .Checked
Case True
.Checked = False
Case False
.Checked = True
End Select
End With
End Sub
Private Sub mnuEnable_Click()
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Enabled = False
.Checked = True
End With
With Me.mnuET
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Enabled = True
.Checked = False
End With
With Me.mnuDT
.Enabled = True
.Checked = False
End With
On Error Resume Next
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Select Case mnuTop.Checked
Case False
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
Case True
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Enabled = False
.Checked = True
End With
With Me.mnuET
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Enabled = True
.Checked = False
End With
With Me.mnuDT
.Enabled = True
.Checked = False
End With
Case False
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Enabled = True
.Checked = False
End With
With Me.mnuET
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Enabled = False
.Checked = True
End With
With Me.mnuDT
.Enabled = False
.Checked = True
End With
End Select
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Enabled = False
.Checked = True
End With
With Me.mnuET
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Enabled = True
.Checked = False
End With
With Me.mnuDT
.Enabled = True
.Checked = False
End With
Case False
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Enabled = True
.Checked = False
End With
With Me.mnuET
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Enabled = False
.Checked = True
End With
With Me.mnuDT
.Enabled = False
.Checked = True
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
Case True
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
End Select
End Sub
Private Sub mnuEnd_Click()
On Error Resume Next
With Me.cSysTray1
.InTray = False
.TrayTip = "Gray Killer - 双击还原窗口,右键单击显示菜单"
End With
Unload Me
Unload AK
End
Unload Me
Unload AK
End
End Sub
Private Sub mnuET_Click()
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Enabled = False
.Checked = True
End With
With Me.mnuET
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Enabled = True
.Checked = False
End With
With Me.mnuDT
.Enabled = True
.Checked = False
End With
On Error Resume Next
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Enabled = False
.Checked = True
End With
With Me.mnuET
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Enabled = True
.Checked = False
End With
With Me.mnuDT
.Enabled = True
.Checked = False
End With
Case False
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Enabled = True
.Checked = False
End With
With Me.mnuET
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Enabled = False
.Checked = True
End With
With Me.mnuDT
.Enabled = False
.Checked = True
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
Case True
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
End Select
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Enabled = False
.Checked = True
End With
With Me.mnuET
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Enabled = True
.Checked = False
End With
With Me.mnuDT
.Enabled = True
.Checked = False
End With
Case False
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Enabled = True
.Checked = False
End With
With Me.mnuET
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Enabled = False
.Checked = True
End With
With Me.mnuDT
.Enabled = False
.Checked = True
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
Case True
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
End Select
End Sub
Private Sub mnuExit_Click()
On Error Resume Next
With Me.cSysTray1
.InTray = False
.TrayTip = "Gray Killer - 双击还原窗口,右键单击显示菜单"
End With
Unload Me
Unload AK
End
End Sub
Private Sub SetWindowConfig()
On Error Resume Next
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Enabled = False
.Checked = True
End With
With Me.mnuET
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Enabled = True
.Checked = False
End With
With Me.mnuDT
.Enabled = True
.Checked = False
End With
Case False
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Enabled = True
.Checked = False
End With
With Me.mnuET
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Enabled = False
.Checked = True
End With
With Me.mnuDT
.Enabled = False
.Checked = True
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
Case True
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
End Select
End Sub
Private Sub mnuInfoTray_Click()
On Error Resume Next
On Error Resume Next
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Dim rtn As Long
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
With Me
.Height = 1455
.Width = 4020
End With
With Me.Timer1
.Enabled = False
.Interval = 1000
End With
If Label1.Caption = "" Then
MsgBox "没有有效窗口", vbCritical, "Error"
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Enabled = False
.Checked = True
End With
With Me.mnuET
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Enabled = True
.Checked = False
End With
With Me.mnuDT
.Enabled = True
.Checked = False
End With
Case False
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Enabled = True
.Checked = False
End With
With Me.mnuET
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Enabled = False
.Checked = True
End With
With Me.mnuDT
.Enabled = False
.Checked = True
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
Case True
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
End Select
Exit Sub
End If
If Label1.Caption = "0" Then
MsgBox "没有有效窗口", vbCritical, "Error"
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Enabled = False
.Checked = True
End With
With Me.mnuET
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Enabled = True
.Checked = False
End With
With Me.mnuDT
.Enabled = True
.Checked = False
End With
Case False
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Enabled = True
.Checked = False
End With
With Me.mnuET
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Enabled = False
.Checked = True
End With
With Me.mnuDT
.Enabled = False
.Checked = True
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
Case True
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
End Select
Exit Sub
End If
With lpWindow
.hWindow = Label1.Caption
.hWindowDC = GetWindowDC(.hWindow)
.lpszCaption = GetPassword(.hWindow)
.hThreadProcessID = GetWindowThreadProcessId(.hWindow, 0)
.hThreadProcessID = OpenProcess(PROCESS_ALL_ACCESS, False, .hThreadProcessID)
GetModuleFileName Null, .lpszThreadProcessName, 1024
GetProcessName .hThreadProcessID, .lpszThreadProcessName, .lpszThreadProcessPath
.hThreadProcessID = OpenProcess(PROCESS_ALL_ACCESS, False, .hThreadProcessID)
GetModuleFileName .hThreadProcessID, .lpszThreadProcessName, 256
.lpszClassName = GetWindowClassName(.hWindow)
Dim lpszClsName As String * 256
GetClassName .hWindow, lpszClsName, 256
.lpszClassName = Trim(lpszClsName)
End With
With lpWindow
.hThreadProcessID = GetWindowThreadProcessId(Me.Label1.Caption, 0)
.hThreadProcess = OpenProcess(PROCESS_ALL_ACCESS, False, .hThreadProcessID)
GetModuleFileName .hThreadProcess, .lpszThreadProcessName, 1024
GetProcessName .hThreadProcessID, .lpszThreadProcessPath, .lpszThreadProcessPath
End With
With lpWindow
GetWindowThreadProcessId Me.Label1.Caption, .hThreadProcessID
Debug.Print .hThreadProcessID
.hThreadProcess = OpenProcess(PROCESS_ALL_ACCESS, False, .hThreadProcessID)
GetModuleFileName .hThreadProcess, .lpszThreadProcessName, 1024
GetProcessName .hThreadProcessID, .lpszExe, .lpszPath
End With
With lpWindow
GetWindowThreadProcessId Me.Label1.Caption, .hThreadProcessID
Debug.Print .hThreadProcessID
.hThreadProcess = OpenProcess(PROCESS_ALL_ACCESS, False, .hThreadProcessID)
GetModuleFileName .hThreadProcess, .lpszThreadProcessName, 1024
GetProcessName .hThreadProcessID, .lpszExe, .lpszPath
End With
With lpWindow
.lpszClassName = Left$(.lpszClassName, InStr(1, .lpszClassName, Chr$(0)) - 1)
End With
With lpWindow
Dim lpStr As String
lpStr = "当前窗口信息:" & vbCrLf & vbCrLf & "句柄:" & vbCrLf & .hWindow & vbCrLf & "设备上下文句柄:" & vbCrLf & .hWindowDC & vbCrLf & "标题:" & vbCrLf & .lpszCaption & vbCrLf & "类名:" & vbCrLf & .lpszClassName
lpStr = lpStr & vbCrLf & "隶属进程ID:" & vbCrLf & .hThreadProcessID & vbCrLf & "隶属进程可执行文件名:" & vbCrLf & .lpszExe
MsgBox lpStr, vbInformation, "Info"
End With
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Enabled = False
.Checked = True
End With
With Me.mnuET
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Enabled = True
.Checked = False
End With
With Me.mnuDT
.Enabled = True
.Checked = False
End With
Case False
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Enabled = True
.Checked = False
End With
With Me.mnuET
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Enabled = False
.Checked = True
End With
With Me.mnuDT
.Enabled = False
.Checked = True
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
Case True
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
End Select
End Sub
Private Sub mnuMini_Click()
On Error Resume Next
With Me
.Hide
.WindowState = 1
End With
With Me.cSysTray1
.TrayTip = "Gray Killer - 双击还原窗口,右键单击显示菜单"
.InTray = True
End With
End Sub
Private Sub mnuRefresh_Click()
On Error Resume Next
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Dim rtn As Long
With Me.Label1
.BackColor = RGB(255, 255, 255)
.BackStyle = 1
.BorderStyle = 1
.Alignment = 2
.Caption = ""
End With
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Enabled = False
.Checked = True
End With
With Me.mnuET
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Enabled = True
.Checked = False
End With
With Me.mnuDT
.Enabled = True
.Checked = False
End With
Case False
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Enabled = True
.Checked = False
End With
With Me.mnuET
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Enabled = False
.Checked = True
End With
With Me.mnuDT
.Enabled = False
.Checked = True
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
Case True
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
End Select
End Sub
Private Sub mnuRestore_Click()
On Error Resume Next
With frmKill
.WindowState = 0
.Show
.Enabled = True
End With
With Me.cSysTray1
.InTray = False
.TrayTip = "Gray Killer - 双击还原窗口,右键单击显示菜单"
End With
On Error Resume Next
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Enabled = False
.Checked = True
End With
With Me.mnuET
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Enabled = True
.Checked = False
End With
With Me.mnuDT
.Enabled = True
.Checked = False
End With
Case False
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Enabled = True
.Checked = False
End With
With Me.mnuET
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Enabled = False
.Checked = True
End With
With Me.mnuDT
.Enabled = False
.Checked = True
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
Case True
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
End Select
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Enabled = False
.Checked = True
End With
With Me.mnuET
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Enabled = True
.Checked = False
End With
With Me.mnuDT
.Enabled = True
.Checked = False
End With
Case False
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Enabled = True
.Checked = False
End With
With Me.mnuET
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Enabled = False
.Checked = True
End With
With Me.mnuDT
.Enabled = False
.Checked = True
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
Case True
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
End Select
End Sub
Private Sub mnuSHC_Click()
On Error Resume Next
With mnuSHC
Select Case .Checked
Case True
.Checked = False
Case False
.Checked = True
End Select
End With
End Sub
Private Sub mnuShowHelp_Click()
On Error Resume Next
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
frmHelp.Show 1
End Sub
Private Sub mnuShowInfo_Click()
On Error Resume Next
On Error Resume Next
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Dim rtn As Long
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
With Me
.Height = 1455
.Width = 4020
End With
With Me.Timer1
.Enabled = False
.Interval = 1000
End With
If Label1.Caption = "" Then
MsgBox "没有有效窗口", vbCritical, "Error"
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Enabled = False
.Checked = True
End With
With Me.mnuET
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Enabled = True
.Checked = False
End With
With Me.mnuDT
.Enabled = True
.Checked = False
End With
Case False
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Enabled = True
.Checked = False
End With
With Me.mnuET
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Enabled = False
.Checked = True
End With
With Me.mnuDT
.Enabled = False
.Checked = True
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
Case True
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
End Select
Exit Sub
End If
If Label1.Caption = "0" Then
MsgBox "没有有效窗口", vbCritical, "Error"
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Enabled = False
.Checked = True
End With
With Me.mnuET
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Enabled = True
.Checked = False
End With
With Me.mnuDT
.Enabled = True
.Checked = False
End With
Case False
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Enabled = True
.Checked = False
End With
With Me.mnuET
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Enabled = False
.Checked = True
End With
With Me.mnuDT
.Enabled = False
.Checked = True
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
Case True
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
End Select
Exit Sub
End If
With lpWindow
.hWindow = Label1.Caption
.hWindowDC = GetWindowDC(.hWindow)
.lpszCaption = GetPassword(.hWindow)
.hThreadProcessID = GetWindowThreadProcessId(.hWindow, 0)
.hThreadProcessID = OpenProcess(PROCESS_ALL_ACCESS, False, .hThreadProcessID)
GetModuleFileName Null, .lpszThreadProcessName, 1024
GetProcessName .hThreadProcessID, .lpszThreadProcessName, .lpszThreadProcessPath
.hThreadProcessID = OpenProcess(PROCESS_ALL_ACCESS, False, .hThreadProcessID)
GetModuleFileName .hThreadProcessID, .lpszThreadProcessName, 256
.lpszClassName = GetWindowClassName(.hWindow)
Dim lpszClsName As String * 256
GetClassName .hWindow, lpszClsName, 256
.lpszClassName = Trim(lpszClsName)
End With
With lpWindow
.hThreadProcessID = GetWindowThreadProcessId(Me.Label1.Caption, 0)
.hThreadProcess = OpenProcess(PROCESS_ALL_ACCESS, False, .hThreadProcessID)
GetModuleFileName .hThreadProcess, .lpszThreadProcessName, 1024
GetProcessName .hThreadProcessID, .lpszThreadProcessPath, .lpszThreadProcessPath
End With
With lpWindow
GetWindowThreadProcessId Me.Label1.Caption, .hThreadProcessID
Debug.Print .hThreadProcessID
.hThreadProcess = OpenProcess(PROCESS_ALL_ACCESS, False, .hThreadProcessID)
GetModuleFileName .hThreadProcess, .lpszThreadProcessName, 1024
GetProcessName .hThreadProcessID, .lpszExe, .lpszPath
End With
With lpWindow
GetWindowThreadProcessId Me.Label1.Caption, .hThreadProcessID
Debug.Print .hThreadProcessID
.hThreadProcess = OpenProcess(PROCESS_ALL_ACCESS, False, .hThreadProcessID)
GetModuleFileName .hThreadProcess, .lpszThreadProcessName, 1024
GetProcessName .hThreadProcessID, .lpszExe, .lpszPath
End With
With lpWindow
.lpszClassName = Left$(.lpszClassName, InStr(1, .lpszClassName, Chr$(0)) - 1)
End With
With lpWindow
Dim lpStr As String
lpStr = "当前窗口信息:" & vbCrLf & vbCrLf & "句柄:" & vbCrLf & .hWindow & vbCrLf & "设备上下文句柄:" & vbCrLf & .hWindowDC & vbCrLf & "标题:" & vbCrLf & .lpszCaption & vbCrLf & "类名:" & vbCrLf & .lpszClassName
lpStr = lpStr & vbCrLf & "隶属进程ID:" & vbCrLf & .hThreadProcessID & vbCrLf & "隶属进程可执行文件名:" & vbCrLf & .lpszExe
MsgBox lpStr, vbInformation, "Info"
End With
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Enabled = False
.Checked = True
End With
With Me.mnuET
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Enabled = True
.Checked = False
End With
With Me.mnuDT
.Enabled = True
.Checked = False
End With
Case False
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Enabled = True
.Checked = False
End With
With Me.mnuET
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Enabled = False
.Checked = True
End With
With Me.mnuDT
.Enabled = False
.Checked = True
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
Case True
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
End Select
End Sub
Private Sub mnuTop_Click()
On Error Resume Next
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Select Case mnuTop.Checked
Case False
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
mnuTop.Checked = True
Case True
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
mnuTop.Checked = False
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
On Error Resume Next
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Enabled = False
.Checked = True
End With
With Me.mnuET
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Enabled = True
.Checked = False
End With
With Me.mnuDT
.Enabled = True
.Checked = False
End With
Case False
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Enabled = True
.Checked = False
End With
With Me.mnuET
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Enabled = False
.Checked = True
End With
With Me.mnuDT
.Enabled = False
.Checked = True
End With
End Select
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Enabled = False
.Checked = True
End With
With Me.mnuET
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Enabled = True
.Checked = False
End With
With Me.mnuDT
.Enabled = True
.Checked = False
End With
Case False
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Enabled = True
.Checked = False
End With
With Me.mnuET
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Enabled = False
.Checked = True
End With
With Me.mnuDT
.Enabled = False
.Checked = True
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
Case True
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
End Select
End Sub
Private Sub mnuTrans_Click()
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
mnuTrans.Checked = True
Exit Sub
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
mnuTrans.Checked = False
Exit Sub
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos frmKill.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
Case True
SetWindowPos frmKill.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
End Select
On Error Resume Next
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Enabled = False
.Checked = True
End With
With Me.mnuET
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Enabled = True
.Checked = False
End With
With Me.mnuDT
.Enabled = True
.Checked = False
End With
Case False
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Enabled = True
.Checked = False
End With
With Me.mnuET
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Enabled = False
.Checked = True
End With
With Me.mnuDT
.Enabled = False
.Checked = True
End With
End Select
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Enabled = False
.Checked = True
End With
With Me.mnuET
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Enabled = True
.Checked = False
End With
With Me.mnuDT
.Enabled = True
.Checked = False
End With
Case False
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Enabled = True
.Checked = False
End With
With Me.mnuET
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Enabled = False
.Checked = True
End With
With Me.mnuDT
.Enabled = False
.Checked = True
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
Case True
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
End Select
End Sub
Private Sub mnuViewInfo_Click()
On Error Resume Next
On Error Resume Next
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Dim rtn As Long
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
With Me
.Height = 1455
.Width = 4020
End With
With Me.Timer1
.Enabled = False
.Interval = 1000
End With
If Label1.Caption = "" Then
MsgBox "没有有效窗口", vbCritical, "Error"
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Enabled = False
.Checked = True
End With
With Me.mnuET
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Enabled = True
.Checked = False
End With
With Me.mnuDT
.Enabled = True
.Checked = False
End With
Case False
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Enabled = True
.Checked = False
End With
With Me.mnuET
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Enabled = False
.Checked = True
End With
With Me.mnuDT
.Enabled = False
.Checked = True
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
Case True
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
End Select
Exit Sub
End If
If Label1.Caption = "0" Then
MsgBox "没有有效窗口", vbCritical, "Error"
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Enabled = False
.Checked = True
End With
With Me.mnuET
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Enabled = True
.Checked = False
End With
With Me.mnuDT
.Enabled = True
.Checked = False
End With
Case False
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Enabled = True
.Checked = False
End With
With Me.mnuET
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Enabled = False
.Checked = True
End With
With Me.mnuDT
.Enabled = False
.Checked = True
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
Case True
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
End Select
Exit Sub
End If
With lpWindow
.hWindow = Label1.Caption
.hWindowDC = GetWindowDC(.hWindow)
.lpszCaption = GetPassword(.hWindow)
.hThreadProcessID = GetWindowThreadProcessId(.hWindow, 0)
.hThreadProcessID = OpenProcess(PROCESS_ALL_ACCESS, False, .hThreadProcessID)
GetModuleFileName Null, .lpszThreadProcessName, 1024
GetProcessName .hThreadProcessID, .lpszThreadProcessName, .lpszThreadProcessPath
.hThreadProcessID = OpenProcess(PROCESS_ALL_ACCESS, False, .hThreadProcessID)
GetModuleFileName .hThreadProcessID, .lpszThreadProcessName, 256
.lpszClassName = GetWindowClassName(.hWindow)
Dim lpszClsName As String * 256
GetClassName .hWindow, lpszClsName, 256
.lpszClassName = Trim(lpszClsName)
End With
With lpWindow
.hThreadProcessID = GetWindowThreadProcessId(Me.Label1.Caption, 0)
.hThreadProcess = OpenProcess(PROCESS_ALL_ACCESS, False, .hThreadProcessID)
GetModuleFileName .hThreadProcess, .lpszThreadProcessName, 1024
GetProcessName .hThreadProcessID, .lpszThreadProcessPath, .lpszThreadProcessPath
End With
With lpWindow
GetWindowThreadProcessId Me.Label1.Caption, .hThreadProcessID
Debug.Print .hThreadProcessID
.hThreadProcess = OpenProcess(PROCESS_ALL_ACCESS, False, .hThreadProcessID)
GetModuleFileName .hThreadProcess, .lpszThreadProcessName, 1024
GetProcessName .hThreadProcessID, .lpszExe, .lpszPath
End With
With lpWindow
GetWindowThreadProcessId Me.Label1.Caption, .hThreadProcessID
Debug.Print .hThreadProcessID
.hThreadProcess = OpenProcess(PROCESS_ALL_ACCESS, False, .hThreadProcessID)
GetModuleFileName .hThreadProcess, .lpszThreadProcessName, 1024
GetProcessName .hThreadProcessID, .lpszExe, .lpszPath
End With
With lpWindow
.lpszClassName = Left$(.lpszClassName, InStr(1, .lpszClassName, Chr$(0)) - 1)
End With
With lpWindow
Dim lpStr As String
lpStr = "当前窗口信息:" & vbCrLf & vbCrLf & "句柄:" & vbCrLf & .hWindow & vbCrLf & "设备上下文句柄:" & vbCrLf & .hWindowDC & vbCrLf & "标题:" & vbCrLf & .lpszCaption & vbCrLf & "类名:" & vbCrLf & .lpszClassName
lpStr = lpStr & vbCrLf & "隶属进程ID:" & vbCrLf & .hThreadProcessID & vbCrLf & "隶属进程可执行文件名:" & vbCrLf & .lpszExe
MsgBox lpStr, vbInformation, "Info"
End With
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Enabled = False
.Checked = True
End With
With Me.mnuET
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Enabled = True
.Checked = False
End With
With Me.mnuDT
.Enabled = True
.Checked = False
End With
Case False
On Error Resume Next
With Me.Timer1
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Enabled = True
.Checked = False
End With
With Me.mnuET
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Enabled = False
.Checked = True
End With
With Me.mnuDT
.Enabled = False
.Checked = True
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
Case True
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1455
.Width = 4020
End With
End Select
End Sub
Private Sub Timer1_Timer()
On Error Resume Next
Dim lpPoint As POINTAPI
GetCursorPos lpPoint
Dim hMenuPos As Long
With lpPoint
End With
Dim hWinTemp As Long
hWinTemp = GetForegroundWindow()
If hWinTemp = Me.hwnd Then
DoEvents
Exit Sub
End If
If hWinTemp = frmHelp.hwnd Then
DoEvents
Exit Sub
End If
If hWinTemp = AK.hwnd Then
DoEvents
Exit Sub
End If
If hWinTemp = AK.hwnd Then
DoEvents
Exit Sub
End If
Dim hWin As Long
hWin = GetForegroundWindow()
With Me.Label1
.BorderStyle = 1
.BackStyle = 1
.BackColor = RGB(255, 255, 255)
.Alignment = 2
.Caption = hWin
End With
If Me.mnuEGC.Checked = True Then
EnumChildWindows hWin, AddressOf EnableDisabledControls, 0
End If
If Me.mnuSHC.Checked = True Then
EnumChildWindows hWin, AddressOf ShowHiddenControls, 0
End If
If Me.mnuEGM.Checked = True Then
Dim hWindow As Long
hWindow = GetForegroundWindow()
Dim hMenu As Long
Dim hSubMenu As Long
Dim nCount As Long
Const EMI_ENABLE = True
Const EMI_DISABLE = False
Const MF_ENABLED = &H0&
Const MF_DISABLED = &H2&
Dim i As Long
Dim nPos As Long
hMenu = GetMenu(hWindow)
nPos = GetMenuItemCount(hMenu)
Debug.Print hMenu
hSubMenu = GetSubMenu(hMenu, 1)
If 1 = 245 Then
Debug.Print hSubMenu
End If
If hSubMenu <> 0 Then
nCount = GetMenuItemCount(hSubMenu)
Debug.Print nCount
If 1 = 245 Then
Dim lpMenuItemID As Long
End If
Dim l As Long
For l = 1 To nPos
hSubMenu = GetSubMenu(hMenu, l - 1)
For i = 1 To nCount
lpMenuItemID = GetMenuItemID(GetSubMenu(hMenu, 0), i - 1)
If 1 = 2 Then
Debug.Print lpMenuItemID
End If
EnableMenuItem hMenu, lpMenuItemID, MF_ENABLED Or MF_BYPOSITION Or MF_BYCOMMAND
Next
Next
End If
End If
If Me.mnuEGM.Checked = True Then
Dim MenuHandle As Long
Dim SubHandle As Long
Dim Id As Long
Dim MainCount As Long
Dim SubCount
MenuHandle = GetMenu(hWin)
MainCount = GetMenuItemCount(MenuHandle)
Dim n1 As Long
Dim n2 As Long
For n1 = 1 To 100
SubHandle = GetSubMenu(MenuHandle, n1 - 1)
SubCount = GetMenuItemCount(SubHandle)
For n2 = 1 To 100
Id = GetMenuItemID(GetSubMenu(MenuHandle, n1 - 1), n2 - 1)
EnableMenuItem SubHandle, n2 - 1, MF_ENABLED Or MF_BYPOSITION
EnableMenuItem MenuHandle, n2 - 1, MF_ENABLED Or MF_BYPOSITION
EnableMenuItem SubHandle, Id, MF_ENABLED
EnableMenuItem MenuHandle, Id, MF_ENABLED
Next
EnableMenuItem SubHandle, Id, MF_ENABLED
EnableMenuItem MenuHandle, Id, MF_ENABLED
EnableMenuItem MenuHandle, n1 - 1, MF_ENABLED Or MF_BYPOSITION
EnableMenuItem SubHandle, n1 - 1, MF_ENABLED Or MF_BYPOSITION
Next
End If
If Me.mnuEGM.Checked = True Then
Dim var_90 As Long
Dim var_8C As Long
Dim hWinActivated As Long
Dim hWinMainMenu As Long
Dim hWinSubMenu As Long
Dim nMainMenuCount As Long
Dim nSubMenuCount As Long
Dim nLoopFirst As Long
Dim nLoopSecond As Long
hWinActivated = GetForegroundWindow()
If hWinActivated = Me.hwnd Then
DoEvents
Exit Sub
End If
If hWinActivated = frmHelp.hwnd Then
DoEvents
Exit Sub
End If
If hWinActivated = AK.hwnd Then
DoEvents
Exit Sub
End If
hWinMainMenu = GetMenu(hWinActivated)
nMainMenuCount = GetMenuItemCount(hWinMainMenu)
Debug.Print nMainMenuCount
For nLoopFirst = 0 To nMainMenuCount - 1
hWinSubMenu = GetSubMenu(hWinMainMenu, nLoopFirst)
nSubMenuCount = GetMenuItemCount(hWinSubMenu)
EnableMenuItem hWinSubMenu, nLoopFirst, MF_ENABLED Or MF_BYPOSITION
EnableMenuItem hWinMainMenu, nLoopFirst, MF_ENABLED Or MF_BYPOSITION
For nLoopSecond = 0 To nSubMenuCount - 1
EnableMenuItem hWinSubMenu, nLoopSecond, MF_ENABLED Or MF_BYPOSITION
EnableMenuItem hWinMainMenu, nLoopSecond, MF_ENABLED Or MF_BYPOSITION
Next
EnableMenuItem hWinSubMenu, nLoopFirst, MF_ENABLED Or MF_BYPOSITION
EnableMenuItem hWinMainMenu, nLoopFirst, MF_ENABLED Or MF_BYPOSITION
Next
End If
If Me.mnuEGM.Checked = True Then
Dim h1 As Long
Dim h2 As Long
h1 = GetForegroundWindow()
h2 = GetMenu(h1)
EnableMenu (h2)
End If
If Me.mnuEGM.Checked = True Then
Dim lpCurPos As POINTAPI
GetCursorPos lpCurPos
Dim hMenuMain As Long
Dim hMenuPoint As Long
Dim lpID As Long
hMenuMain = GetMenu(hWin)
hMenuPoint = MenuItemFromPoint(hWin, hMenuMain, lpCurPos)
EnableMenuItem hMenuPoint, 0, MF_ENABLED Or MF_BYPOSITION
End If
DoEvents
End Sub
Private Sub EnableMenu(hMenu As Long)
Dim var_90 As Long
Dim var_8C As Long
Dim var_A8 As Long
Dim var_88 As Long
Dim var_98 As Long
On Error Resume Next
var_8C = GetMenuItemCount(hMenu)
If (var_8C > 0) Then
For var_98 = 0 To (var_8C - 1): var_88 = var_98
EnableMenuItem hMenu, var_88, &H400
var_90 = GetSubMenu(hMenu, var_88)
var_A8 = CVar(var_90)
If CBool(var_A8 <> 0) Then
EnableMenuItem CLng(var_A8), var_90, 0
End If
Next var_98
End If
End Sub
