VERSION 5.00
Begin VB.Form lgfLogOff 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Logoff"
   ClientHeight    =   4785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame1 
      Caption         =   "信息"
      Height          =   3390
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   4455
      Begin VB.TextBox Text1 
         Height          =   2310
         Left            =   150
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   210
         Width           =   4155
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "秒"
         Height          =   180
         Left            =   4110
         TabIndex        =   7
         Top             =   2820
         Width           =   180
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "30"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Left            =   1395
         TabIndex        =   6
         Top             =   2580
         Width           =   2415
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "距离注销还有"
         Height          =   180
         Left            =   180
         TabIndex        =   5
         Top             =   2820
         Width           =   1080
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "立刻注销(&L)"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   4320
      Width           =   1350
   End
   Begin VB.CommandButton Command2 
      Caption         =   "暂停(&P)"
      Default         =   -1  'True
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   4320
      Width           =   1350
   End
   Begin VB.CommandButton Command3 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   375
      Left            =   3240
      TabIndex        =   0
      Top             =   4320
      Width           =   1350
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1740
      Top             =   2220
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   3
      Left            =   3930
      Picture         =   "Form2.frx":0000
      Top             =   375
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   2
      Left            =   3060
      Picture         =   "Form2.frx":0442
      Top             =   225
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   1
      Left            =   3630
      Picture         =   "Form2.frx":0884
      Top             =   255
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   0
      Left            =   3255
      Picture         =   "Form2.frx":0CC6
      Top             =   390
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   210
      Picture         =   "Form2.frx":1108
      Top             =   135
      Width           =   480
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "    系统即将注销,请保存正在进行的工作和活动应用程序的数据,所有尚未保存的数据都可能丢失."
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   960
      TabIndex        =   8
      Top             =   120
      Width           =   3480
   End
End
Attribute VB_Name = "lgfLogOff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Const WS_EX_LAYERED = &H80000
Private Const GWL_EXSTYLE = (-20)
Private Const LWA_ALPHA = &H2
Private Const LWA_COLORKEY = &H1
Dim timeleft As Integer
Private Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
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
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
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
Sub AdjustTokenPrivilegesForNT()
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
Private Sub Command1_Click()
On Error Resume Next
Const HWND_NOTOPMOST = -2
SetWindowPos Me.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 5160
.Width = 4770
End With
Dim ans As Integer
ans = MsgBox("确定立刻注销吗?", vbQuestion + vbYesNo, "Ask")
If ans = vbYes Then
SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 5160
.Width = 4770
End With
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
If Form1.Check2.Value = 0 Then
SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 5160
.Width = 4770
End With
ExitWindowsEx EWX_LOGOFF, &HFFFF
Else
SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 5160
.Width = 4770
End With
ExitWindowsEx EWX_LOGOFF Or EWX_FORCE, 0
End If
Else
SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 5160
.Width = 4770
End With
Exit Sub
End If
End Sub
Private Sub Command2_Click()
On Error Resume Next
If Command2.Caption = "暂停(&P)" Then
Timer1.Enabled = False
Command2.Caption = "继续(&O)"
Else
Timer1.Enabled = True
Command2.Caption = "暂停(&P)"
End If
End Sub
Private Sub Command3_Click()
On Error Resume Next
Const HWND_NOTOPMOST = -2
SetWindowPos Me.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 5160
.Width = 4770
End With
Dim ans As Integer
ans = MsgBox("确定要取消吗?", vbQuestion + vbYesNo, "Ask")
If ans = vbYes Then
SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 5160
.Width = 4770
End With
Form1.Show
Unload Me
Else
SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 5160
.Width = 4770
End With
Exit Sub
End If
End Sub
Private Sub Form_Load()
On Error Resume Next
Dim rtn     As Long
Select Case Form1.Check3.Value
Case 1
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, Form1.HScroll1.Value, LWA_ALPHA
Case 0
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
End Select
SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Image1.Picture = Image2(Form1.Combo1.ListIndex).Picture
.Height = 5160
.Width = 4770
End With
Me.KeyPreview = True
Me.Icon = LoadPicture("")
Me.Command3.Cancel = True
Me.Command2.Default = True
Text1.Text = Form1.Text2.Text
Text1.Locked = True
timeleft = Val(Form1.Text1.Text)
Label3.Caption = Val(Form1.Text1.Text)
Timer1.Interval = 1000
Timer1.Enabled = True
End Sub
Private Sub Timer1_Timer()
On Error Resume Next
If timeleft <= 0 Then
Label3.Caption = "处理中..."
Exit Sub
End If
timeleft = timeleft - 1
Label3.Caption = timeleft
If timeleft = 0 Then
Label3.Caption = "处理中..."
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
If Form1.Check2.Value = 0 Then
ExitWindowsEx EWX_LOGOFF, &HFFFF
Else
ExitWindowsEx EWX_LOGOFF Or EWX_FORCE, 0
End If
End If
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 And Shift = vbAltMask Then
On Error Resume Next
Dim ans As Integer
ans = MsgBox("确定要取消吗?", vbQuestion + vbYesNo, "Ask")
If ans = vbYes Then
Form1.Show
Unload Me
Else
Exit Sub
End If
End If
End Sub

