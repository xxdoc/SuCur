VERSION 5.00
Begin VB.Form delReboot 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reboot"
   ClientHeight    =   4080
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4080
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer DelayExecuter 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1740
      Top             =   1800
   End
   Begin VB.Timer TimeGetter 
      Interval        =   1000
      Left            =   540
      Top             =   645
   End
   Begin VB.Frame Frame1 
      Caption         =   "信息"
      Height          =   2520
      Left            =   90
      TabIndex        =   3
      Top             =   840
      Width           =   4500
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "当前系统时间"
         Height          =   180
         Left            =   195
         TabIndex        =   9
         Top             =   405
         Width           =   1080
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "66:66"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   1515
         TabIndex        =   8
         Top             =   165
         Width           =   2880
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "设定重启时间"
         Height          =   180
         Left            =   195
         TabIndex        =   7
         Top             =   1155
         Width           =   1080
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "66:66"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   1515
         TabIndex        =   6
         Top             =   915
         Width           =   2880
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "操作执行延迟"
         Height          =   180
         Left            =   195
         TabIndex        =   5
         Top             =   1905
         Width           =   1080
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "启用"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   1515
         TabIndex        =   4
         Top             =   1665
         Width           =   2880
      End
   End
   Begin VB.CommandButton Command3 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   525
      Left            =   2655
      TabIndex        =   1
      Top             =   3480
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "立刻重启(&R)"
      Height          =   525
      Left            =   120
      TabIndex        =   0
      Top             =   3480
      Width           =   1935
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   3
      Left            =   870
      Picture         =   "delReboot.frx":0000
      Top             =   150
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   2
      Left            =   0
      Picture         =   "delReboot.frx":0442
      Top             =   0
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   1
      Left            =   570
      Picture         =   "delReboot.frx":0884
      Top             =   30
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   0
      Left            =   195
      Picture         =   "delReboot.frx":0CC6
      Top             =   165
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "    系统即将重启,请保存正在进行的工作和活动应用程序的数据,所有尚未保存的数据都可能丢失."
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
      TabIndex        =   2
      Top             =   120
      Width           =   3480
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   210
      Picture         =   "delReboot.frx":1108
      Top             =   135
      Width           =   480
   End
End
Attribute VB_Name = "delReboot"
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
Private Type TimeData
HourValue As Integer
MinuteValue As Integer
TimeUserSet As String
TimeSystem As String
ExecuteDelay As Integer
End Type
Dim TimeVar As TimeData
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
.Height = 4455
.Width = 4770
End With
Dim ans As Integer
ans = MsgBox("确定立刻重启吗?", vbQuestion + vbYesNo, "Ask")
If ans = vbYes Then
SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 4455
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
If Form2.Check2.Value = 0 Then
SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 4455
.Width = 4770
End With
ExitWindowsEx EWX_REBOOT, &HFFFF
Else
SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 4455
.Width = 4770
End With
ExitWindowsEx EWX_REBOOT Or EWX_FORCE, 0
End If
Else
SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 4455
.Width = 4770
End With
Exit Sub
End If
End Sub
Private Sub Command3_Click()
On Error Resume Next
Const HWND_NOTOPMOST = -2
SetWindowPos Me.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 4455
.Width = 4770
End With
Dim ans As Integer
ans = MsgBox("确定要取消吗?", vbQuestion + vbYesNo, "Ask")
If ans = vbYes Then
SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 4455
.Width = 4770
End With
Form2.Show
Unload Me
Else
SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 4455
.Width = 4770
End With
Exit Sub
End If
End Sub
Private Sub DelayExecuter_Timer()
On Error Resume Next
With TimeVar
.ExecuteDelay = .ExecuteDelay - 1
Label7.Caption = .ExecuteDelay
End With
If TimeVar.ExecuteDelay = 0 Then
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
If Form2.Check2.Value = 0 Then
ExitWindowsEx EWX_REBOOT, &HFFFF
Else
ExitWindowsEx EWX_REBOOT Or EWX_FORCE, 0
End If
End If
End Sub
Private Sub Form_Load()
On Error Resume Next
Dim rtn     As Long
Select Case Form2.Check3.Value
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
.Height = 4455
.Width = 4770
End With
With TimeVar
.HourValue = Hour(Now)
.MinuteValue = Minute(Now)
.TimeSystem = (.HourValue) & ":" & (.MinuteValue)
Label3.Caption = .TimeSystem
.ExecuteDelay = 30
.HourValue = (Form2.Text1.Text)
.MinuteValue = Str(Form2.Text2.Text)
.TimeUserSet = .HourValue & ":" & .MinuteValue
Label5.Caption = .TimeUserSet
End With
With Me
.Command1.Default = True
.Command3.Cancel = True
.Image1.Picture = Image2(Form2.Combo1.ListIndex).Picture
.KeyPreview = True
.TimeGetter.Interval = 1000
.TimeGetter.Enabled = True
.DelayExecuter.Interval = 1000
.DelayExecuter.Enabled = False
End With
Select Case Form2.Check1.Value
Case 0
Me.Label7.Caption = "禁用"
Case 1
Me.Label7.Caption = "启用"
End Select
End Sub
Private Sub TimeGetter_Timer()
With TimeVar
.HourValue = Hour(Now)
.MinuteValue = Minute(Now)
.TimeSystem = (.HourValue) & ":" & (.MinuteValue)
Label3.Caption = .TimeSystem
End With
If Label3.Caption = Label5.Caption Then
Select Case Form2.Check1.Value
Case 0
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
If Form2.Check2.Value = 0 Then
ExitWindowsEx EWX_REBOOT, &HFFFF
Else
ExitWindowsEx EWX_REBOOT Or EWX_FORCE, 0
End If
Case 1
Label5.Caption = "延迟执行中"
Label6.Caption = "延迟剩余时间"
TimeVar.ExecuteDelay = 30
Label7.Caption = TimeVar.ExecuteDelay
Me.DelayExecuter.Enabled = True
Me.TimeGetter.Enabled = False
End Select
End If
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 And Shift = vbAltMask Then
On Error Resume Next
Dim ans As Integer
ans = MsgBox("确定要取消吗?", vbQuestion + vbYesNo, "Ask")
If ans = vbYes Then
Form2.Show
Unload Me
Else
Exit Sub
End If
End If
End Sub
