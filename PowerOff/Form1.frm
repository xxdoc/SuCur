VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C0C000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Power Cotrol & Timer - PC-DOS Workshop"
   ClientHeight    =   4950
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6495
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4950
   ScaleWidth      =   6495
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C5791D&
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   0
      ScaleHeight     =   195
      ScaleWidth      =   6495
      TabIndex        =   20
      Top             =   0
      Width           =   6495
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "工具(H)"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   2
         Left            =   1500
         TabIndex        =   23
         Top             =   0
         Width           =   630
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "操作(O)"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   1
         Left            =   780
         TabIndex        =   22
         Top             =   0
         Width           =   630
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "程序(P)"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   0
         Left            =   60
         TabIndex        =   21
         Top             =   0
         Width           =   630
      End
   End
   Begin VB.CommandButton Command7 
      Caption         =   "按时间执行电源操作(&T)"
      Height          =   435
      Left            =   2925
      TabIndex        =   14
      Top             =   4425
      Width           =   2085
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00C5791D&
      Caption         =   "强制关闭没有响应的进程(&F)"
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   4455
      Width           =   2655
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C000&
      Caption         =   "选项"
      Height          =   2235
      Left            =   135
      TabIndex        =   5
      Top             =   1980
      Width           =   6255
      Begin VB.HScrollBar HScroll1 
         Height          =   255
         Left            =   2760
         Max             =   255
         Min             =   100
         TabIndex        =   18
         Top             =   1335
         Value           =   199
         Width           =   2655
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00C0C000&
         Caption         =   "启用提示窗口半透明功能(&E)"
         Height          =   255
         Left            =   165
         TabIndex        =   17
         Top             =   1335
         Width           =   2610
      End
      Begin VB.ComboBox Combo1 
         Height          =   300
         ItemData        =   "Form1.frx":030A
         Left            =   1440
         List            =   "Form1.frx":031A
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   1740
         Width           =   3990
      End
      Begin VB.TextBox Text2 
         Height          =   765
         Left            =   2205
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   12
         Text            =   "Form1.frx":0342
         Top             =   525
         Width           =   3915
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0C000&
         Caption         =   "不要倒计时而是直接执行操作(&D)"
         Height          =   330
         Left            =   3270
         TabIndex        =   10
         Top             =   195
         Width           =   2910
      End
      Begin VB.CommandButton Command6 
         Caption         =   "+"
         Height          =   285
         Left            =   2985
         TabIndex        =   9
         Top             =   210
         Width           =   240
      End
      Begin VB.CommandButton Command5 
         Caption         =   "-"
         Height          =   285
         Left            =   2190
         TabIndex        =   8
         Top             =   210
         Width           =   240
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   270
         Left            =   2430
         MaxLength       =   2
         TabIndex        =   7
         Text            =   "30"
         Top             =   210
         Width           =   555
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "199"
         Height          =   225
         Left            =   5445
         TabIndex        =   19
         Top             =   1350
         Width           =   690
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "提示窗口图标"
         Height          =   180
         Left            =   120
         TabIndex        =   15
         Top             =   1785
         Width           =   1080
      End
      Begin VB.Image Image4 
         BorderStyle     =   1  'Fixed Single
         Height          =   540
         Left            =   5580
         Picture         =   "Form1.frx":035F
         Top             =   1635
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "提示文本(可选):"
         Height          =   180
         Left            =   165
         TabIndex        =   11
         Top             =   525
         Width           =   1350
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "倒计时阀值(1秒-99秒)"
         Height          =   180
         Left            =   150
         TabIndex        =   6
         Top             =   270
         Width           =   1800
      End
   End
   Begin VB.CommandButton Command4 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   435
      Left            =   5100
      TabIndex        =   4
      Top             =   4425
      Width           =   1290
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C000&
      Caption         =   "想要计算机做什么?"
      ForeColor       =   &H00000000&
      Height          =   1740
      Left            =   1320
      TabIndex        =   0
      Top             =   240
      Width           =   5070
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FFFF00&
         Caption         =   "注销(&L)"
         Height          =   1200
         Left            =   3495
         MaskColor       =   &H00000000&
         Picture         =   "Form1.frx":07A1
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   345
         UseMaskColor    =   -1  'True
         Width           =   1365
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H0000FF00&
         Caption         =   "重新启动(&R)"
         Height          =   1200
         Left            =   1845
         MaskColor       =   &H00000000&
         Picture         =   "Form1.frx":2BE3
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   345
         UseMaskColor    =   -1  'True
         Width           =   1365
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H000000FF&
         Caption         =   "关机(&S)"
         Height          =   1200
         Left            =   210
         MaskColor       =   &H00000000&
         Picture         =   "Form1.frx":4725
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   345
         UseMaskColor    =   -1  'True
         Width           =   1365
      End
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   3
      Left            =   870
      Picture         =   "Form1.frx":6B67
      Top             =   1260
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   2
      Left            =   0
      Picture         =   "Form1.frx":6FA9
      Top             =   1110
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   1
      Left            =   570
      Picture         =   "Form1.frx":73EB
      Top             =   1140
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   0
      Left            =   195
      Picture         =   "Form1.frx":782D
      Top             =   1275
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image3 
      Height          =   75
      Index           =   0
      Left            =   -15
      Picture         =   "Form1.frx":7C6F
      Stretch         =   -1  'True
      Top             =   4260
      Width           =   6600
   End
   Begin VB.Image Image3 
      Height          =   480
      Index           =   1
      Left            =   540
      Picture         =   "Form1.frx":80DB
      Top             =   630
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   225
      Picture         =   "Form1.frx":83E5
      Top             =   300
      Width           =   720
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00C5791D&
      FillStyle       =   0  'Solid
      Height          =   705
      Left            =   -45
      Top             =   4335
      Width           =   7995
   End
   Begin VB.Menu mnuP 
      Caption         =   "程序(&P)"
      Visible         =   0   'False
      Begin VB.Menu mnuExit 
         Caption         =   "退出(&E)"
      End
   End
   Begin VB.Menu mnuOP 
      Caption         =   "操作(&O)"
      Visible         =   0   'False
      Begin VB.Menu mnushut 
         Caption         =   "关机(&S)..."
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuREBO 
         Caption         =   "重启(&R)..."
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuLOGO 
         Caption         =   "注销(&L)..."
         Shortcut        =   ^L
      End
      Begin VB.Menu b1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuED 
         Caption         =   "立即执行操作(&E)"
         Begin VB.Menu mnuS 
            Caption         =   "立刻关闭系统(&H)"
         End
         Begin VB.Menu mnuR 
            Caption         =   "立刻重新启动(&B)"
         End
         Begin VB.Menu mnuL 
            Caption         =   "立刻注销用户(&O)"
         End
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "工具(&T)"
      Visible         =   0   'False
      Begin VB.Menu mnuTT 
         Caption         =   "定时电源操作程序(&D)..."
         Shortcut        =   ^D
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim flag As Boolean
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
Private Sub Check1_Click()
On Error Resume Next
On Error Resume Next
Dim aryCur As Integer
For aryCur = 0 To Me.Label5.UBound
With Me.Label5(aryCur)
.AutoSize = True
.BackStyle = 1
.BackColor = Me.Picture1.BackColor
End With
Next
Select Case Check1.Value
Case 0
Text1.Enabled = True
Command5.Enabled = True
Command6.Enabled = True
Text2.Enabled = True
Me.Check3.Enabled = True
Me.HScroll1.Enabled = True
With Me.Label4
.Enabled = True
.Caption = Me.HScroll1.Value
End With
Me.Label3.Enabled = True
Me.Combo1.Enabled = True
Me.Image4.Enabled = True
Label1.Enabled = True
Label2.Enabled = True
If Me.Check3.Value = 0 Then
With Label4
.Enabled = False
.Caption = "Disable"
End With
Me.HScroll1.Enabled = False
Else
Me.HScroll1.Enabled = True
With Me.Label4
.Caption = Me.HScroll1.Value
.Enabled = True
End With
End If
Case 1
Text1.Enabled = False
Text2.Enabled = False
Command5.Enabled = False
Command6.Enabled = False
Me.Check3.Enabled = False
Me.HScroll1.Enabled = False
With Me.Label4
.Enabled = False
.Caption = "Disable"
End With
Me.Label3.Enabled = False
Me.Combo1.Enabled = False
Me.Image4.Enabled = False
Label1.Enabled = False
Label2.Enabled = False
End Select
Select Case Check3.Value
Case 0
Me.HScroll1.Enabled = False
If Check3.Enabled = False Then
With Me.Label4
.Caption = "Disable"
.Enabled = False
End With
Me.HScroll1.Enabled = False
Else
Me.HScroll1.Enabled = True
With Me.Label4
.Caption = Me.HScroll1.Value
.Enabled = True
End With
End If
Case 1
Me.HScroll1.Enabled = True
With Me.Label4
.Caption = Me.HScroll1.Value
.Enabled = True
End With
If Check3.Enabled = False Then
With Me.Label4
.Caption = "Disable"
.Enabled = False
End With
Me.HScroll1.Enabled = False
ElseIf Check3.Enabled = True Then
Select Case Me.Check3.Value
Case 1
Me.HScroll1.Enabled = True
With Me.Label4
.Caption = Me.HScroll1.Value
.Enabled = True
End With
Case 0
Me.HScroll1.Enabled = False
With Me.Label4
.Caption = "Disable"
.Enabled = False
End With
End Select
End If
End Select
Err.Clear
Err.Clear
If Me.Check3.Value = 0 Then
With Me.Label4
.Caption = "Disable"
.Enabled = False
End With
Me.HScroll1.Enabled = False
End If
End Sub
Private Sub Check1_GotFocus()
On Error Resume Next
Dim aryCur As Integer
For aryCur = 0 To Me.Label5.UBound
With Me.Label5(aryCur)
.AutoSize = True
.BackStyle = 1
.BackColor = Me.Picture1.BackColor
End With
Next
On Error Resume Next
flag = False
Err.Clear
End Sub
Private Sub Check2_Click()
On Error Resume Next
Dim aryCur As Integer
For aryCur = 0 To Me.Label5.UBound
With Me.Label5(aryCur)
.AutoSize = True
.BackStyle = 1
.BackColor = Me.Picture1.BackColor
End With
Next
On Error Resume Next
If Check2.Value = 1 Then
Dim ans As Integer
ans = MsgBox("警告:不推荐强制结束进程,因为这样可能导致不可知的问题,继续?", vbExclamation + vbYesNo, "Alert")
If ans = vbYes Then
Err.Clear
Exit Sub
Else
Check2.Value = 0
End If
End If
Err.Clear
End Sub
Private Sub Check2_GotFocus()
On Error Resume Next
Dim aryCur As Integer
For aryCur = 0 To Me.Label5.UBound
With Me.Label5(aryCur)
.AutoSize = True
.BackStyle = 1
.BackColor = Me.Picture1.BackColor
End With
Next
On Error Resume Next
flag = False
Err.Clear
End Sub
Private Sub Check3_Click()
On Error Resume Next
Dim aryCur As Integer
For aryCur = 0 To Me.Label5.UBound
With Me.Label5(aryCur)
.AutoSize = True
.BackStyle = 1
.BackColor = Me.Picture1.BackColor
End With
Next
On Error Resume Next
Select Case Check3.Value
Case 0
Me.HScroll1.Enabled = False
With Me.Label4
.Caption = "Disable"
.Enabled = False
End With
Case 1
Me.HScroll1.Enabled = True
With Me.Label4
.Caption = Me.HScroll1.Value
.Enabled = True
End With
End Select
Err.Clear
End Sub
Private Sub Check3_GotFocus()
On Error Resume Next
Dim aryCur As Integer
For aryCur = 0 To Me.Label5.UBound
With Me.Label5(aryCur)
.AutoSize = True
.BackStyle = 1
.BackColor = Me.Picture1.BackColor
End With
Next
On Error Resume Next
flag = False
Err.Clear
End Sub
Private Sub Combo1_Change()
On Error Resume Next
Dim aryCur As Integer
For aryCur = 0 To Me.Label5.UBound
With Me.Label5(aryCur)
.AutoSize = True
.BackStyle = 1
.BackColor = Me.Picture1.BackColor
End With
Next
On Error Resume Next
Me.Image4.Picture = Me.Image2(Combo1.ListIndex).Picture
Err.Clear
End Sub
Private Sub Combo1_Click()
On Error Resume Next
Dim aryCur As Integer
For aryCur = 0 To Me.Label5.UBound
With Me.Label5(aryCur)
.AutoSize = True
.BackStyle = 1
.BackColor = Me.Picture1.BackColor
End With
Next
On Error Resume Next
Me.Image4.Picture = Me.Image2(Combo1.ListIndex).Picture
Err.Clear
End Sub
Private Sub Combo1_GotFocus()
On Error Resume Next
Dim aryCur As Integer
For aryCur = 0 To Me.Label5.UBound
With Me.Label5(aryCur)
.AutoSize = True
.BackStyle = 1
.BackColor = Me.Picture1.BackColor
End With
Next
On Error Resume Next
flag = False
Err.Clear
End Sub
Private Sub Command1_Click()
On Error Resume Next
On Error Resume Next
Dim aryCur As Integer
For aryCur = 0 To Me.Label5.UBound
With Me.Label5(aryCur)
.AutoSize = True
.BackStyle = 1
.BackColor = Me.Picture1.BackColor
End With
Next
If Check1.Value = 0 Then
Me.Hide
shdShutdown.Show
Else
Dim ans As Integer
ans = MsgBox("确定执行这个操作吗?请注意保持文件好数据!", vbExclamation + vbYesNo, "Ask")
If ans = vbYes Then
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
If Check2.Value = 0 Then
ExitWindowsEx EWX_SHUTDOWN, &HFFFF
Else
ExitWindowsEx EWX_SHUTDOWN Or EWX_FORCE, 0
End If
Else
Exit Sub
End If
End If
Err.Clear
End Sub
Private Sub Command1_GotFocus()
On Error Resume Next
flag = False
Err.Clear
End Sub
Private Sub Command2_Click()
On Error Resume Next
On Error Resume Next
Dim aryCur As Integer
For aryCur = 0 To Me.Label5.UBound
With Me.Label5(aryCur)
.AutoSize = True
.BackStyle = 1
.BackColor = Me.Picture1.BackColor
End With
Next
If Check1.Value = 0 Then
Me.Hide
rbtReboot.Show
Else
Dim ans As Integer
ans = MsgBox("确定执行这个操作吗?请注意保持文件好数据!", vbExclamation + vbYesNo, "Ask")
If ans = vbYes Then
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
If Check2.Value = 0 Then
ExitWindowsEx EWX_REBOOT, &HFFFF
Else
ExitWindowsEx EWX_REBOOT Or EWX_FORCE, 0
End If
Else
Exit Sub
End If
End If
Err.Clear
End Sub
Private Sub Command2_GotFocus()
On Error Resume Next
flag = False
Err.Clear
End Sub
Private Sub Command3_Click()
On Error Resume Next
On Error Resume Next
Dim aryCur As Integer
For aryCur = 0 To Me.Label5.UBound
With Me.Label5(aryCur)
.AutoSize = True
.BackStyle = 1
.BackColor = Me.Picture1.BackColor
End With
Next
If Check1.Value = 0 Then
Me.Hide
lgfLogOff.Show
Else
Dim ans As Integer
ans = MsgBox("确定执行这个操作吗?请注意保持文件好数据!", vbExclamation + vbYesNo, "Ask")
If ans = vbYes Then
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
If Check2.Value = 0 Then
ExitWindowsEx EWX_LOGOFF, &HFFFF
Else
ExitWindowsEx EWX_LOGOFF Or EWX_FORCE, 0
End If
Else
Err.Clear
Exit Sub
End If
End If
Err.Clear
End Sub
Private Sub Command3_GotFocus()
On Error Resume Next
flag = False
Err.Clear
End Sub
Private Sub Command4_Click()
On Error Resume Next
Err.Clear
Unload Form1
Unload Form2
Unload delLogoff
Unload delReboot
Unload delShutdown
Unload lgfLogOff
Unload rbtReboot
Unload shdShutdown
End Sub
Private Sub Command4_GotFocus()
On Error Resume Next
flag = False
Err.Clear
End Sub
Private Sub Command5_Click()
On Error Resume Next
If Val(Text1.Text) <= 1 Then
Text1.Text = 1
Err.Clear
Exit Sub
End If
Text1.Text = Val(Text1.Text) - 1
Err.Clear
End Sub
Private Sub Command5_GotFocus()
flag = False
Err.Clear
End Sub
Private Sub Command6_Click()
On Error Resume Next
On Error Resume Next
Dim aryCur As Integer
For aryCur = 0 To Me.Label5.UBound
With Me.Label5(aryCur)
.AutoSize = True
.BackStyle = 1
.BackColor = Me.Picture1.BackColor
End With
Next
If Val(Text1.Text) >= 99 Then
Text1.Text = 99
Exit Sub
End If
Text1.Text = Val(Text1.Text) + 1
Err.Clear
End Sub
Private Sub Command6_GotFocus()
On Error Resume Next
On Error Resume Next
Dim aryCur As Integer
For aryCur = 0 To Me.Label5.UBound
With Me.Label5(aryCur)
.AutoSize = True
.BackStyle = 1
.BackColor = Me.Picture1.BackColor
End With
Next
flag = False
Err.Clear
End Sub
Private Sub Command7_Click()
On Error Resume Next
On Error Resume Next
Dim aryCur As Integer
For aryCur = 0 To Me.Label5.UBound
With Me.Label5(aryCur)
.AutoSize = True
.BackStyle = 1
.BackColor = Me.Picture1.BackColor
End With
Next
Form1.Hide
Form2.Show
Err.Clear
End Sub
Private Sub Form_Click()
On Error Resume Next
Dim aryCur As Integer
For aryCur = 0 To Me.Label5.UBound
With Me.Label5(aryCur)
.AutoSize = True
.BackStyle = 1
.BackColor = Me.Picture1.BackColor
End With
Next
End Sub
Private Sub Form_GotFocus()
On Error Resume Next
On Error Resume Next
Dim aryCur As Integer
For aryCur = 0 To Me.Label5.UBound
With Me.Label5(aryCur)
.AutoSize = True
.BackStyle = 1
.BackColor = Me.Picture1.BackColor
End With
Next
flag = False
Err.Clear
End Sub
Private Sub Form_Initialize()
On Error Resume Next
Dim lpCommand As String
Dim ans As Integer
Const TOKEN_ADJUST_PRIVILEGES = &H20
Const TOKEN_QUERY = &H8
Const SE_PRIVILEGE_ENABLED = &H2
Dim hdlProcessHandle As Long
Dim hdlTokenHandle As Long
Dim tmpLuid As LUID
Dim tkp As TOKEN_PRIVILEGES
Dim tkpNewButIgnored As TOKEN_PRIVILEGES
Dim lBufferNeeded As Long
lpCommand = Command$()
Select Case lpCommand
Case "-s"
ans = MsgBox("确定执行这个操作吗?请注意保持文件好数据!", vbExclamation + vbYesNo, "Ask")
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
If Check2.Value = 0 Then
ExitWindowsEx EWX_SHUTDOWN, &HFFFF
Else
ExitWindowsEx EWX_SHUTDOWN Or EWX_FORCE, 0
End If
Else
Exit Sub
End If
Case "-S"
ans = MsgBox("确定执行这个操作吗?请注意保持文件好数据!", vbExclamation + vbYesNo, "Ask")
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
If Check2.Value = 0 Then
ExitWindowsEx EWX_SHUTDOWN, &HFFFF
Else
ExitWindowsEx EWX_SHUTDOWN Or EWX_FORCE, 0
End If
Else
Exit Sub
End If
Case "/s"
ans = MsgBox("确定执行这个操作吗?请注意保持文件好数据!", vbExclamation + vbYesNo, "Ask")
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
If Check2.Value = 0 Then
ExitWindowsEx EWX_SHUTDOWN, &HFFFF
Else
ExitWindowsEx EWX_SHUTDOWN Or EWX_FORCE, 0
End If
Else
Exit Sub
End If
Case "/S"
ans = MsgBox("确定执行这个操作吗?请注意保持文件好数据!", vbExclamation + vbYesNo, "Ask")
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
If Check2.Value = 0 Then
ExitWindowsEx EWX_SHUTDOWN, &HFFFF
Else
ExitWindowsEx EWX_SHUTDOWN Or EWX_FORCE, 0
End If
Else
Exit Sub
End If
Case "-r"
ans = MsgBox("确定执行这个操作吗?请注意保持文件好数据!", vbExclamation + vbYesNo, "Ask")
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
If Check2.Value = 0 Then
ExitWindowsEx EWX_REBOOT, &HFFFF
Else
ExitWindowsEx EWX_REBOOT Or EWX_FORCE, 0
End If
Else
Exit Sub
End If
Case "-R"
ans = MsgBox("确定执行这个操作吗?请注意保持文件好数据!", vbExclamation + vbYesNo, "Ask")
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
If Check2.Value = 0 Then
ExitWindowsEx EWX_REBOOT, &HFFFF
Else
ExitWindowsEx EWX_REBOOT Or EWX_FORCE, 0
End If
Else
Exit Sub
End If
Case "/r"
ans = MsgBox("确定执行这个操作吗?请注意保持文件好数据!", vbExclamation + vbYesNo, "Ask")
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
If Check2.Value = 0 Then
ExitWindowsEx EWX_REBOOT, &HFFFF
Else
ExitWindowsEx EWX_REBOOT Or EWX_FORCE, 0
End If
Else
Exit Sub
End If
Case "/R"
ans = MsgBox("确定执行这个操作吗?请注意保持文件好数据!", vbExclamation + vbYesNo, "Ask")
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
If Check2.Value = 0 Then
ExitWindowsEx EWX_REBOOT, &HFFFF
Else
ExitWindowsEx EWX_REBOOT Or EWX_FORCE, 0
End If
Else
Exit Sub
End If
Case "-l"
ans = MsgBox("确定执行这个操作吗?请注意保持文件好数据!", vbExclamation + vbYesNo, "Ask")
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
If Check2.Value = 0 Then
ExitWindowsEx EWX_LOGOFF, &HFFFF
Else
ExitWindowsEx EWX_LOGOFF Or EWX_FORCE, 0
End If
Else
Exit Sub
End If
Case "-L"
ans = MsgBox("确定执行这个操作吗?请注意保持文件好数据!", vbExclamation + vbYesNo, "Ask")
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
If Check2.Value = 0 Then
ExitWindowsEx EWX_LOGOFF, &HFFFF
Else
ExitWindowsEx EWX_LOGOFF Or EWX_FORCE, 0
End If
Else
Exit Sub
End If
Case "/l"
ans = MsgBox("确定执行这个操作吗?请注意保持文件好数据!", vbExclamation + vbYesNo, "Ask")
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
If Check2.Value = 0 Then
ExitWindowsEx EWX_LOGOFF, &HFFFF
Else
ExitWindowsEx EWX_LOGOFF Or EWX_FORCE, 0
End If
Else
Exit Sub
End If
Case "/L"
ans = MsgBox("确定执行这个操作吗?请注意保持文件好数据!", vbExclamation + vbYesNo, "Ask")
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
If Check2.Value = 0 Then
ExitWindowsEx EWX_LOGOFF, &HFFFF
Else
ExitWindowsEx EWX_LOGOFF Or EWX_FORCE, 0
End If
Else
Exit Sub
End If
Case "-ts"
Me.Hide
cmdSD.Show
Case "-TS"
Me.Hide
cmdSD.Show
Case "-Ts"
Me.Hide
cmdSD.Show
Case "-tS"
Me.Hide
cmdSD.Show
Case "/ts"
Me.Hide
cmdSD.Show
Case "/TS"
Me.Hide
cmdSD.Show
Case "/Ts"
Me.Hide
cmdSD.Show
Case "/tS"
Me.Hide
cmdSD.Show
Case "-tr"
Me.Hide
cmdRB.Show
Case "-TR"
Me.Hide
cmdRB.Show
Case "-Ts"
Me.Hide
cmdRB.Show
Case "-tR"
Me.Hide
cmdRB.Show
Case "/tr"
Me.Hide
cmdRB.Show
Case "/TR"
Me.Hide
cmdRB.Show
Case "/Tr"
Me.Hide
cmdRB.Show
Case "/tR"
Me.Hide
cmdRB.Show
Case "-tl"
Me.Hide
cmdLF.Show
Case "-TL"
Me.Hide
cmdLF.Show
Case "-Tl"
Me.Hide
cmdLF.Show
Case "-tL"
Me.Hide
cmdLF.Show
Case "/tl"
Me.Hide
cmdLF.Show
Case "/TL"
Me.Hide
cmdLF.Show
Case "/Tl"
Me.Hide
cmdRB.Show
Case "/tL"
Me.Hide
cmdRB.Show
Case "/?"
With Me
.Visible = False
End With
frmParam.Show
Case "-?"
With Me
.Visible = False
End With
frmParam.Show
End Select
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
On Error Resume Next
If KeyCode = vbKeyF4 And Shift = vbAltMask Then
On Error Resume Next
Err.Clear
Unload Form1
Unload Form2
Unload delLogoff
Unload delReboot
Unload delShutdown
Unload lgfLogOff
Unload rbtReboot
Unload shdShutdown
End If
Dim aryCur As Integer
For aryCur = 0 To Me.Label5.UBound
With Me.Label5(aryCur)
.AutoSize = True
.BackStyle = 1
.BackColor = Me.Picture1.BackColor
End With
Next
If Shift = vbCtrlMask Then
Select Case KeyCode
Case vbKeyS
On Error Resume Next
If Check1.Value = 0 Then
Me.Hide
shdShutdown.Show
Else
Dim ans As Integer
ans = MsgBox("确定执行这个操作吗?请注意保持文件好数据!", vbExclamation + vbYesNo, "Ask")
If ans = vbYes Then
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
If Check2.Value = 0 Then
ExitWindowsEx EWX_SHUTDOWN, &HFFFF
Else
ExitWindowsEx EWX_SHUTDOWN Or EWX_FORCE, 0
End If
Else
Exit Sub
End If
End If
Err.Clear
Exit Sub
Case vbKeyR
On Error Resume Next
If Check1.Value = 0 Then
Me.Hide
rbtReboot.Show
Else
ans = MsgBox("确定执行这个操作吗?请注意保持文件好数据!", vbExclamation + vbYesNo, "Ask")
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
If Check2.Value = 0 Then
ExitWindowsEx EWX_REBOOT, &HFFFF
Else
ExitWindowsEx EWX_REBOOT Or EWX_FORCE, 0
End If
Else
Exit Sub
End If
End If
Err.Clear
Exit Sub
Case vbKeyL
On Error Resume Next
If Check1.Value = 0 Then
Me.Hide
lgfLogOff.Show
Else
ans = MsgBox("确定执行这个操作吗?请注意保持文件好数据!", vbExclamation + vbYesNo, "Ask")
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
If Check2.Value = 0 Then
ExitWindowsEx EWX_LOGOFF, &HFFFF
Else
ExitWindowsEx EWX_LOGOFF Or EWX_FORCE, 0
End If
Else
Err.Clear
Exit Sub
End If
End If
Err.Clear
Exit Sub
Case vbKeyD
On Error Resume Next
Form1.Hide
Form2.Show
Err.Clear
End Select
Exit Sub
End If
If Shift = vbAltMask Then
Dim X As Single
Dim Y As Single
Dim z As Integer
Select Case KeyCode
Case vbKeyP
On Error Resume Next
X = Label5(0).Left
Y = Label5(0).Top + Label5(0).Height
For z = 0 To 2
If z <> 0 Then
Label5(z).BackColor = Picture1.BackColor
End If
Next
PopupMenu Me.mnuP, , X, Y
Exit Sub
Case vbKeyO
On Error Resume Next
X = Label5(1).Left
Y = Label5(1).Top + Label5(1).Height
For z = 0 To 2
If z <> 1 Then
Label5(z).BackColor = Picture1.BackColor
End If
Next
PopupMenu Me.mnuOP, , X, Y
Exit Sub
Case vbKeyH
On Error Resume Next
X = Label5(2).Left
Y = Label5(2).Top + Label5(2).Height
For z = 0 To 2
If z <> 2 Then
Label5(z).BackColor = Picture1.BackColor
End If
Next
PopupMenu Me.mnuTools, , X, Y
Exit Sub
Case vbKeyS
On Error Resume Next
If Check1.Value = 0 Then
Me.Hide
shdShutdown.Show
Else
ans = MsgBox("确定执行这个操作吗?请注意保持文件好数据!", vbExclamation + vbYesNo, "Ask")
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
If Check2.Value = 0 Then
ExitWindowsEx EWX_SHUTDOWN, &HFFFF
Else
ExitWindowsEx EWX_SHUTDOWN Or EWX_FORCE, 0
End If
Else
Exit Sub
End If
End If
Err.Clear
Exit Sub
Case vbKeyR
On Error Resume Next
If Check1.Value = 0 Then
Me.Hide
rbtReboot.Show
Else
ans = MsgBox("确定执行这个操作吗?请注意保持文件好数据!", vbExclamation + vbYesNo, "Ask")
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
If Check2.Value = 0 Then
ExitWindowsEx EWX_REBOOT, &HFFFF
Else
ExitWindowsEx EWX_REBOOT Or EWX_FORCE, 0
End If
Else
Exit Sub
End If
End If
Err.Clear
Exit Sub
Case vbKeyL
On Error Resume Next
If Check1.Value = 0 Then
Me.Hide
rbtReboot.Show
Else
ans = MsgBox("确定执行这个操作吗?请注意保持文件好数据!", vbExclamation + vbYesNo, "Ask")
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
If Check2.Value = 0 Then
ExitWindowsEx EWX_LOGOFF, &HFFFF
Else
ExitWindowsEx EWX_LOGOFF Or EWX_FORCE, 0
End If
Else
Exit Sub
End If
End If
Err.Clear
Exit Sub
Case vbKeyD
On Error Resume Next
Select Case Check1.Value
Case 0
Text1.Enabled = True
Command5.Enabled = True
Command6.Enabled = True
Text2.Enabled = True
Me.Check3.Enabled = True
Me.HScroll1.Enabled = True
With Me.Label4
.Enabled = True
.Caption = Me.HScroll1.Value
End With
Me.Label3.Enabled = True
Me.Combo1.Enabled = True
Me.Image4.Enabled = True
Label1.Enabled = True
Label2.Enabled = True
If Me.Check3.Value = 0 Then
With Label4
.Enabled = False
.Caption = "Disable"
End With
Me.HScroll1.Enabled = False
Else
Me.HScroll1.Enabled = True
With Me.Label4
.Caption = Me.HScroll1.Value
.Enabled = True
End With
End If
Case 1
Text1.Enabled = False
Text2.Enabled = False
Command5.Enabled = False
Command6.Enabled = False
Me.Check3.Enabled = False
Me.HScroll1.Enabled = False
With Me.Label4
.Enabled = False
.Caption = "Disable"
End With
Me.Label3.Enabled = False
Me.Combo1.Enabled = False
Me.Image4.Enabled = False
Label1.Enabled = False
Label2.Enabled = False
End Select
Select Case Check3.Value
Case 0
Me.HScroll1.Enabled = False
If Check3.Enabled = False Then
With Me.Label4
.Caption = "Disable"
.Enabled = False
End With
Me.HScroll1.Enabled = False
Else
Me.HScroll1.Enabled = True
With Me.Label4
.Caption = Me.HScroll1.Value
.Enabled = True
End With
End If
Case 1
Me.HScroll1.Enabled = True
With Me.Label4
.Caption = Me.HScroll1.Value
.Enabled = True
End With
If Check3.Enabled = False Then
With Me.Label4
.Caption = "Disable"
.Enabled = False
End With
Me.HScroll1.Enabled = False
ElseIf Check3.Enabled = True Then
Select Case Me.Check3.Value
Case 1
Me.HScroll1.Enabled = True
With Me.Label4
.Caption = Me.HScroll1.Value
.Enabled = True
End With
Case 0
Me.HScroll1.Enabled = False
With Me.Label4
.Caption = "Disable"
.Enabled = False
End With
End Select
End If
End Select
Err.Clear
Err.Clear
If Me.Check3.Value = 0 Then
With Me.Label4
.Caption = "Disable"
.Enabled = False
End With
Me.HScroll1.Enabled = False
End If
Exit Sub
Case vbKeyE
On Error Resume Next
Select Case Check3.Value
Case 0
Me.HScroll1.Enabled = False
With Me.Label4
.Caption = "Disable"
.Enabled = False
End With
Case 1
Me.HScroll1.Enabled = True
With Me.Label4
.Caption = Me.HScroll1.Value
.Enabled = True
End With
End Select
Err.Clear
Case vbKeyF
On Error Resume Next
If Check2.Value = 1 Then
ans = MsgBox("警告:不推荐强制结束进程,因为这样可能导致不可知的问题,继续?", vbExclamation + vbYesNo, "Alert")
If ans = vbYes Then
Err.Clear
Exit Sub
Else
Check2.Value = 0
End If
End If
Err.Clear
Exit Sub
Case vbKeyC
On Error Resume Next
Err.Clear
Unload Form1
Unload Form2
Unload delLogoff
Unload delReboot
Unload delShutdown
Unload lgfLogOff
Unload rbtReboot
Unload shdShutdown
Exit Sub
Case vbKeyT
On Error Resume Next
Form1.Hide
Form2.Show
Err.Clear
Exit Sub
Case Else
KeyCode = 0
Exit Sub
End Select
End If
If KeyCode = vbKeyF4 And Shift = vbAltMask Then
On Error Resume Next
Err.Clear
Unload Form1
Unload Form2
Unload delLogoff
Unload delReboot
Unload delShutdown
Unload lgfLogOff
Unload rbtReboot
Unload shdShutdown
End If
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
On Error Resume Next
On Error Resume Next
Dim aryCur As Integer
For aryCur = 0 To Me.Label5.UBound
With Me.Label5(aryCur)
.AutoSize = True
.BackStyle = 1
.BackColor = Me.Picture1.BackColor
End With
Next
If flag = True Then
Exit Sub
End If
Const TOKEN_ADJUST_PRIVILEGES = &H20
Const TOKEN_QUERY = &H8
Const SE_PRIVILEGE_ENABLED = &H2
Dim hdlProcessHandle As Long
Dim hdlTokenHandle As Long
Dim tmpLuid As LUID
Dim tkp As TOKEN_PRIVILEGES
Dim tkpNewButIgnored As TOKEN_PRIVILEGES
Dim lBufferNeeded As Long
Dim ans As Integer
Select Case KeyAscii
Case 115
If Check1.Value = 0 Then
Me.Hide
shdShutdown.Show
Else
ans = MsgBox("确定执行这个操作吗?请注意保持文件好数据!", vbExclamation + vbYesNo, "Ask")
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
If Check2.Value = 0 Then
ExitWindowsEx EWX_SHUTDOWN, &HFFFF
Else
ExitWindowsEx EWX_SHUTDOWN Or EWX_FORCE, 0
End If
Else
Err.Clear
Exit Sub
End If
End If
Case 114
If Check1.Value = 0 Then
Me.Hide
rbtReboot.Show
Else
ans = MsgBox("确定执行这个操作吗?请注意保持文件好数据!", vbExclamation + vbYesNo, "Ask")
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
If Check2.Value = 0 Then
ExitWindowsEx EWX_REBOOT, &HFFFF
Else
ExitWindowsEx EWX_REBOOT Or EWX_FORCE, 0
End If
Else
Err.Clear
Exit Sub
End If
End If
Case 108
If Check1.Value = 0 Then
Me.Hide
lgfLogOff.Show
Else
ans = MsgBox("确定执行这个操作吗?请注意保持文件好数据!", vbExclamation + vbYesNo, "Ask")
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
If Check2.Value = 0 Then
ExitWindowsEx EWX_LOGOFF, &HFFFF
Else
ExitWindowsEx EWX_LOGOFF Or EWX_FORCE, 0
End If
Else
Exit Sub
End If
End If
Case Else
KeyAscii = 0
Exit Sub
End Select
End Sub
Private Sub Form_Load()
On Error Resume Next
If App.PrevInstance = False Then
On Error Resume Next
Dim aryCur As Integer
For aryCur = 0 To Me.Label5.UBound
With Me.Label5(aryCur)
.AutoSize = True
.BackStyle = 1
.BackColor = Me.Picture1.BackColor
End With
Next
With Me.HScroll1
.Max = 255
.Min = 100
.LargeChange = 10
.SmallChange = 1
.Value = 199
End With
Me.Check3.Value = 1
Select Case Check3.Value
Case 0
Me.HScroll1.Enabled = False
Me.Label4.Enabled = False
Case 1
Me.HScroll1.Enabled = True
Me.Label4.Enabled = True
End Select
With Me.Picture1
.Left = 0
.Top = 0
.Width = Me.Width
.Visible = True
End With
With Me
.Combo1.ListIndex = 2
.Image4.Picture = .Image2(2).Picture
.Left = Screen.Width / 2 - Me.Width / 2
.Top = Screen.Height / 2 - Me.Height / 2
.KeyPreview = True
.Check1.Value = 0
.Check2.Value = 0
.Command1.ToolTipText = "结束所有打开的应用程序并结束Windows会话"
.Command2.ToolTipText = "临时性结束Windows会话并重新启动计算机"
.Command3.ToolTipText = "关闭所有打开的程序并断开当前用户连接,但是不会结束Windows会话"
End With
Me.Check3.Value = 0
Select Case Check3.Value
Case 0
Me.HScroll1.Enabled = False
With Me.Label4
.Caption = "Disable"
.Enabled = False
End With
Case 1
Me.HScroll1.Enabled = True
With Me.Label4
.Caption = Me.HScroll1.Value
.Enabled = True
End With
For aryCur = 0 To Me.Label5.UBound
With Me.Label5(aryCur)
.AutoSize = True
.BackStyle = 1
.BackColor = Me.Picture1.BackColor
End With
Next
End Select
Else
MsgBox "Windows检测到应用程序已经有一个活动的实例,为保证系统稳定性,应用程序将退出", vbCritical, "Error"
End
End If
Err.Clear
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Dim aryCur As Integer
For aryCur = 0 To Me.Label5.UBound
With Me.Label5(aryCur)
.AutoSize = True
.BackStyle = 1
.BackColor = Me.Picture1.BackColor
End With
Next
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Err.Clear
Unload Form1
Unload Form2
Unload delLogoff
Unload delReboot
Unload delShutdown
Unload lgfLogOff
Unload rbtReboot
Unload shdShutdown
End Sub
Private Sub HScroll1_Change()
On Error Resume Next
Label4.Caption = Me.HScroll1.Value
On Error Resume Next
Dim aryCur As Integer
For aryCur = 0 To Me.Label5.UBound
With Me.Label5(aryCur)
.AutoSize = True
.BackStyle = 1
.BackColor = Me.Picture1.BackColor
End With
Next
End Sub
Private Sub HScroll1_GotFocus()
On Error Resume Next
flag = False
On Error Resume Next
Dim aryCur As Integer
For aryCur = 0 To Me.Label5.UBound
With Me.Label5(aryCur)
.AutoSize = True
.BackStyle = 1
.BackColor = Me.Picture1.BackColor
End With
Next
End Sub
Private Sub Label4_Click()
On Error Resume Next
On Error Resume Next
Dim aryCur As Integer
For aryCur = 0 To Me.Label5.UBound
With Me.Label5(aryCur)
.AutoSize = True
.BackStyle = 1
.BackColor = Me.Picture1.BackColor
End With
Next
Dim alp As Integer
Dim oldalp As Integer
Dim rtn As Long
oldalp = Label4.Caption
alp = Val(InputBox$("请输入透明度" & vbCrLf & "范围:100-255", "Alpha", 199))
If Val(alp) = 0 Then
Me.HScroll1.Value = oldalp
Label4.Caption = Me.HScroll1.Value
Exit Sub
End If
If 100 <= alp And alp <= 255 Then
Me.HScroll1.Value = alp
Label4.Caption = Me.HScroll1.Value
Else
MsgBox "无效透明度数值", vbCritical, "Error"
End If
End Sub
Private Sub Label5_Click(Index As Integer)
On Error Resume Next
Dim X As Single
Dim Y As Single
X = Label5(Index).Left
Y = Label5(Index).Top + Label5(Index).Height
Dim z As Integer
For z = 0 To 2
If z <> Index Then
Label5(z).BackColor = Picture1.BackColor
End If
Next
Select Case Index
Case 0
PopupMenu Me.mnuP, , X, Y
Case 1
PopupMenu Me.mnuOP, , X, Y
Case 2
PopupMenu Me.mnuTools, , X, Y
End Select
End Sub
Private Sub Label5_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Label5(Index).BackColor = &H6C2F35
End Sub
Private Sub mnuExit_Click()
On Error Resume Next
Err.Clear
Unload Form1
Unload Form2
Unload delLogoff
Unload delReboot
Unload delShutdown
Unload lgfLogOff
Unload rbtReboot
Unload shdShutdown
End Sub
Private Sub mnuL_Click()
On Error Resume Next
On Error Resume Next
Dim aryCur As Integer
For aryCur = 0 To Me.Label5.UBound
With Me.Label5(aryCur)
.AutoSize = True
.BackStyle = 1
.BackColor = Me.Picture1.BackColor
End With
Next
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
If Check2.Value = 0 Then
ExitWindowsEx EWX_LOGOFF, &HFFFF
Else
ExitWindowsEx EWX_LOGOFF Or EWX_FORCE, 0
End If
End Sub
Private Sub mnuLOGO_Click()
On Error Resume Next
On Error Resume Next
Dim aryCur As Integer
For aryCur = 0 To Me.Label5.UBound
With Me.Label5(aryCur)
.AutoSize = True
.BackStyle = 1
.BackColor = Me.Picture1.BackColor
End With
Next
If Check1.Value = 0 Then
Me.Hide
lgfLogOff.Show
Else
Dim ans As Integer
ans = MsgBox("确定执行这个操作吗?请注意保持文件好数据!", vbExclamation + vbYesNo, "Ask")
If ans = vbYes Then
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
If Check2.Value = 0 Then
ExitWindowsEx EWX_LOGOFF, &HFFFF
Else
ExitWindowsEx EWX_LOGOFF Or EWX_FORCE, 0
End If
Else
Err.Clear
Exit Sub
End If
End If
Err.Clear
End Sub
Private Sub mnuR_Click()
On Error Resume Next
Dim aryCur As Integer
For aryCur = 0 To Me.Label5.UBound
With Me.Label5(aryCur)
.AutoSize = True
.BackStyle = 1
.BackColor = Me.Picture1.BackColor
End With
Next
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
If Check2.Value = 0 Then
ExitWindowsEx EWX_REBOOT, &HFFFF
Else
ExitWindowsEx EWX_REBOOT Or EWX_FORCE, 0
End If
End Sub
Private Sub mnuREBO_Click()
On Error Resume Next
Dim aryCur As Integer
For aryCur = 0 To Me.Label5.UBound
With Me.Label5(aryCur)
.AutoSize = True
.BackStyle = 1
.BackColor = Me.Picture1.BackColor
End With
Next
On Error Resume Next
If Check1.Value = 0 Then
Me.Hide
rbtReboot.Show
Else
Dim ans As Integer
ans = MsgBox("确定执行这个操作吗?请注意保持文件好数据!", vbExclamation + vbYesNo, "Ask")
If ans = vbYes Then
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
If Check2.Value = 0 Then
ExitWindowsEx EWX_REBOOT, &HFFFF
Else
ExitWindowsEx EWX_REBOOT Or EWX_FORCE, 0
End If
Else
Exit Sub
End If
End If
Err.Clear
End Sub
Private Sub mnuS_Click()
On Error Resume Next
Dim aryCur As Integer
For aryCur = 0 To Me.Label5.UBound
With Me.Label5(aryCur)
.AutoSize = True
.BackStyle = 1
.BackColor = Me.Picture1.BackColor
End With
Next
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
If Check2.Value = 0 Then
ExitWindowsEx EWX_SHUTDOWN, &HFFFF
Else
ExitWindowsEx EWX_SHUTDOWN Or EWX_FORCE, 0
End If
End Sub
Private Sub mnushut_Click()
On Error Resume Next
Dim aryCur As Integer
For aryCur = 0 To Me.Label5.UBound
With Me.Label5(aryCur)
.AutoSize = True
.BackStyle = 1
.BackColor = Me.Picture1.BackColor
End With
Next
On Error Resume Next
If Check1.Value = 0 Then
Me.Hide
shdShutdown.Show
Else
Dim ans As Integer
ans = MsgBox("确定执行这个操作吗?请注意保持文件好数据!", vbExclamation + vbYesNo, "Ask")
If ans = vbYes Then
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
If Check2.Value = 0 Then
ExitWindowsEx EWX_SHUTDOWN, &HFFFF
Else
ExitWindowsEx EWX_SHUTDOWN Or EWX_FORCE, 0
End If
Else
Exit Sub
End If
End If
Err.Clear
End Sub
Private Sub mnuTT_Click()
On Error Resume Next
Dim aryCur As Integer
For aryCur = 0 To Me.Label5.UBound
With Me.Label5(aryCur)
.AutoSize = True
.BackStyle = 1
.BackColor = Me.Picture1.BackColor
End With
Next
On Error Resume Next
Form1.Hide
Form2.Show
Err.Clear
End Sub
Private Sub Picture1_Click()
On Error Resume Next
Dim aryCur As Integer
For aryCur = 0 To Me.Label5.UBound
With Me.Label5(aryCur)
.AutoSize = True
.BackStyle = 1
.BackColor = Me.Picture1.BackColor
End With
Next
End Sub
Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Dim aryCur As Integer
For aryCur = 0 To Me.Label5.UBound
With Me.Label5(aryCur)
.AutoSize = True
.BackStyle = 1
.BackColor = Me.Picture1.BackColor
End With
Next
End Sub
Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
On Error Resume Next
Dim aryCur As Integer
For aryCur = 0 To Me.Label5.UBound
With Me.Label5(aryCur)
.AutoSize = True
.BackStyle = 1
.BackColor = Me.Picture1.BackColor
End With
Next
If KeyCode = vbKeyUp Then
If Val(Text1.Text) >= 99 Then
Text1.Text = 99
Exit Sub
End If
Text1.Text = Val(Text1.Text) + 1
Err.Clear
ElseIf KeyCode = vbKeyRight Then
If Val(Text1.Text) >= 99 Then
Text1.Text = 99
Exit Sub
End If
Text1.Text = Val(Text1.Text) + 1
Err.Clear
ElseIf KeyCode = vbKeyRight Then
If Val(Text1.Text) <= 1 Then
Text1.Text = 1
Err.Clear
Exit Sub
End If
Text1.Text = Val(Text1.Text) - 1
Err.Clear
ElseIf KeyCode = vbKeyDown Then
If Val(Text1.Text) <= 1 Then
Text1.Text = 1
Err.Clear
Exit Sub
End If
Text1.Text = Val(Text1.Text) - 1
Err.Clear
Else
Exit Sub
End If
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
On Error Resume Next
Dim aryCur As Integer
For aryCur = 0 To Me.Label5.UBound
With Me.Label5(aryCur)
.AutoSize = True
.BackStyle = 1
.BackColor = Me.Picture1.BackColor
End With
Next
On Error Resume Next
If KeyAscii < 48 Or KeyAscii > 57 Then
If KeyAscii <> 8 Then
KeyAscii = 0
End If
End If
End Sub
Private Sub Text1_LostFocus()
On Error Resume Next
Dim aryCur As Integer
For aryCur = 0 To Me.Label5.UBound
With Me.Label5(aryCur)
.AutoSize = True
.BackStyle = 1
.BackColor = Me.Picture1.BackColor
End With
Next
On Error Resume Next
If Trim(Val(Text1.Text)) < 1 Or Trim(Val(Text1.Text)) > 99 Then
Text1.Text = 30
End If
End Sub
Private Sub Text2_GotFocus()
On Error Resume Next
Dim aryCur As Integer
For aryCur = 0 To Me.Label5.UBound
With Me.Label5(aryCur)
.AutoSize = True
.BackStyle = 1
.BackColor = Me.Picture1.BackColor
End With
Next
On Error Resume Next
flag = True
End Sub
Private Sub Text2_LostFocus()
On Error Resume Next
Dim aryCur As Integer
For aryCur = 0 To Me.Label5.UBound
With Me.Label5(aryCur)
.AutoSize = True
.BackStyle = 1
.BackColor = Me.Picture1.BackColor
End With
Next
On Error Resume Next
flag = False
End Sub
Private Sub Text1_GotFocus()
On Error Resume Next
Dim aryCur As Integer
For aryCur = 0 To Me.Label5.UBound
With Me.Label5(aryCur)
.AutoSize = True
.BackStyle = 1
.BackColor = Me.Picture1.BackColor
End With
Next
On Error Resume Next
flag = True
End Sub
Private Sub Text1f_LostFocus()
On Error Resume Next
Dim aryCur As Integer
For aryCur = 0 To Me.Label5.UBound
With Me.Label5(aryCur)
.AutoSize = True
.BackStyle = 1
.BackColor = Me.Picture1.BackColor
End With
Next
On Error Resume Next
flag = False
End Sub
