VERSION 5.00
Begin VB.Form AK 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About"
   ClientHeight    =   3765
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6540
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3765
   ScaleWidth      =   6540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "确定"
      Default         =   -1  'True
      Height          =   315
      Left            =   5325
      TabIndex        =   5
      Top             =   3405
      Width           =   1170
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   1140
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   4
      Text            =   "frmAbout.frx":0000
      Top             =   2730
      Width           =   5340
   End
   Begin VB.Image Image3 
      Height          =   75
      Index           =   1
      Left            =   0
      Picture         =   "frmAbout.frx":0076
      Stretch         =   -1  'True
      Top             =   1050
      Width           =   6600
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   1095
      Left            =   6360
      Top             =   0
      Width           =   270
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   1
      Left            =   270
      Picture         =   "frmAbout.frx":04E2
      Top             =   2085
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "基于Microsoft(R) Visual Studio(R) 6.00 制作"
      Height          =   180
      Index           =   3
      Left            =   1140
      TabIndex        =   3
      Top             =   2370
      Width           =   3870
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PC-DOS Workshop开发"
      Height          =   180
      Index           =   2
      Left            =   1140
      TabIndex        =   2
      Top             =   1995
      Width           =   1710
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "版本1.00"
      Height          =   180
      Index           =   1
      Left            =   1140
      TabIndex        =   1
      Top             =   1635
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   180
      Index           =   0
      Left            =   1140
      TabIndex        =   0
      Top             =   1275
      Width           =   90
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   0
      Left            =   270
      Picture         =   "frmAbout.frx":0DAC
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   1125
      Left            =   0
      Picture         =   "frmAbout.frx":11EE
      Stretch         =   -1  'True
      Top             =   -30
      Width           =   6480
   End
End
Attribute VB_Name = "AK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Dim lpszCaptionNew As String
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
Private Sub Command1_Click()
On Error Resume Next
Unload Me
If 1 = 2 Then
On Error Resume Next
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Dim rtn As Long
Select Case frmKill.mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong frmKill.hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes frmKill.hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong frmKill.hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes frmKill.hwnd, 0, 192, LWA_ALPHA
End Select
Select Case frmKill.mnuEnable.Checked
Case True
On Error Resume Next
With frmKill.Timer1
.Interval = 1000
.Enabled = True
End With
With frmKill.mnuEnable
.Enabled = False
.Checked = True
End With
With frmKill.mnuET
.Checked = True
.Enabled = False
End With
With frmKill.mnuDisable
.Enabled = True
.Checked = False
End With
With frmKill.mnuDT
.Enabled = True
.Checked = False
End With
Case False
On Error Resume Next
With frmKill.Timer1
.Interval = 1000
.Enabled = False
End With
With frmKill.mnuEnable
.Enabled = True
.Checked = False
End With
With frmKill.mnuET
.Checked = False
.Enabled = True
End With
With frmKill.mnuDisable
.Enabled = False
.Checked = True
End With
With frmKill.mnuDT
.Enabled = False
.Checked = True
End With
End Select
Select Case frmKill.mnuTop.Checked
Case False
SetWindowPos frmKill.hwnd, HWND_TOPMOST, 0, 0, frmKill.Width, frmKill.Height, SWP_NOMOVE Or SWP_NOSIZE
With frmKill
.Height = 1860
.Width = 4020
End With
Case True
SetWindowPos frmKill.hwnd, HWND_NOTOPMOST, 0, 0, frmKill.Width, frmKill.Height, SWP_NOMOVE Or SWP_NOSIZE
With frmKill
.Height = 1860
.Width = 4020
End With
End Select
End If
frmKill.Show
frmKill.SetFocus
If 1 = 2 Then
With frmKill
.WindowState = 1
.WindowState = 0
End With
End If
Select Case frmKill.mnuEnable.Checked
Case True
On Error Resume Next
With frmKill.Timer1
.Interval = 1000
.Enabled = True
End With
With frmKill.mnuEnable
.Enabled = False
.Checked = True
End With
With frmKill.mnuET
.Checked = True
.Enabled = False
End With
With frmKill.mnuDisable
.Enabled = True
.Checked = False
End With
With frmKill.mnuDT
.Enabled = True
.Checked = False
End With
Case False
On Error Resume Next
With frmKill.Timer1
.Interval = 1000
.Enabled = False
End With
With frmKill.mnuEnable
.Enabled = True
.Checked = False
End With
With frmKill.mnuET
.Checked = False
.Enabled = True
End With
With frmKill.mnuDisable
.Enabled = False
.Checked = True
End With
With frmKill.mnuDT
.Enabled = False
.Checked = True
End With
End Select
End Sub
Private Sub Form_Activate()
On Error Resume Next
Me.Command1.SetFocus
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End Sub
Private Sub Form_Load()
On Error Resume Next
With Me.Image2(0)
.Stretch = True
End With
With Label1(0)
.AutoSize = True
.BackStyle = 0
.BorderStyle = 0
.Caption = App.Title
End With
With Label1(1)
.AutoSize = True
.BackStyle = 0
.BorderStyle = 0
.Caption = "版本" & App.Major & "." & App.Minor & App.Revision
End With
With Me.Text1
.Locked = True
End With
With Me
.Left = Screen.Width / 2 - .Width / 2
.Top = Screen.Height / 2 - .Height / 2
.Icon = LoadPicture()
End With
With Me.Command1
.Cancel = True
.Default = True
End With
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
Select Case frmKill.mnuEnable.Checked
Case True
On Error Resume Next
With frmKill.Timer1
.Interval = 1000
.Enabled = True
End With
With frmKill.mnuEnable
.Enabled = False
.Checked = True
End With
With frmKill.mnuET
.Checked = True
.Enabled = False
End With
With frmKill.mnuDisable
.Enabled = True
.Checked = False
End With
With frmKill.mnuDT
.Enabled = True
.Checked = False
End With
Case False
On Error Resume Next
With frmKill.Timer1
.Interval = 1000
.Enabled = False
End With
With frmKill.mnuEnable
.Enabled = True
.Checked = False
End With
With frmKill.mnuET
.Checked = False
.Enabled = True
End With
With frmKill.mnuDisable
.Enabled = False
.Checked = True
End With
With frmKill.mnuDT
.Enabled = False
.Checked = True
End With
End Select
Exit Sub
If 1 = 245 Then
On Error Resume Next
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Dim rtn As Long
Select Case frmKill.mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong frmKill.hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes frmKill.hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong frmKill.hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes frmKill.hwnd, 0, 192, LWA_ALPHA
End Select
Select Case frmKill.mnuEnable.Checked
Case True
On Error Resume Next
With frmKill.Timer1
.Interval = 1000
.Enabled = True
End With
With frmKill.mnuEnable
.Enabled = False
.Checked = True
End With
With frmKill.mnuET
.Checked = True
.Enabled = False
End With
With frmKill.mnuDisable
.Enabled = True
.Checked = False
End With
With frmKill.mnuDT
.Enabled = True
.Checked = False
End With
Case False
On Error Resume Next
With frmKill.Timer1
.Interval = 1000
.Enabled = False
End With
With frmKill.mnuEnable
.Enabled = True
.Checked = False
End With
With frmKill.mnuET
.Checked = False
.Enabled = True
End With
With frmKill.mnuDisable
.Enabled = False
.Checked = True
End With
With frmKill.mnuDT
.Enabled = False
.Checked = True
End With
End Select
Select Case frmKill.mnuTop.Checked
Case False
SetWindowPos frmKill.hwnd, HWND_TOPMOST, 0, 0, frmKill.Width, frmKill.Height, SWP_NOMOVE Or SWP_NOSIZE
With frmKill
.Height = 1860
.Width = 4020
End With
Case True
SetWindowPos frmKill.hwnd, HWND_NOTOPMOST, 0, 0, frmKill.Width, frmKill.Height, SWP_NOMOVE Or SWP_NOSIZE
With frmKill
.Height = 1860
.Width = 4020
End With
End Select
With frmKill
.WindowState = 1
.WindowState = 0
End With
End If
Select Case frmKill.mnuEnable.Checked
Case True
On Error Resume Next
With frmKill.Timer1
.Interval = 1000
.Enabled = True
End With
With frmKill.mnuEnable
.Enabled = False
.Checked = True
End With
With frmKill.mnuET
.Checked = True
.Enabled = False
End With
With frmKill.mnuDisable
.Enabled = True
.Checked = False
End With
With frmKill.mnuDT
.Enabled = True
.Checked = False
End With
Case False
On Error Resume Next
With frmKill.Timer1
.Interval = 1000
.Enabled = False
End With
With frmKill.mnuEnable
.Enabled = True
.Checked = False
End With
With frmKill.mnuET
.Checked = False
.Enabled = True
End With
With frmKill.mnuDisable
.Enabled = False
.Checked = True
End With
With frmKill.mnuDT
.Enabled = False
.Checked = True
End With
End Select
End Sub
