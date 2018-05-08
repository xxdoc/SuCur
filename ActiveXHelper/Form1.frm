VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "控件小助手 - PC_DOS Workshop"
   ClientHeight    =   6525
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   14430
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6525
   ScaleWidth      =   14430
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   6315
      Left            =   7245
      ScaleHeight     =   6315
      ScaleWidth      =   7140
      TabIndex        =   32
      Top             =   210
      Width           =   7140
      Begin SHDocVwCtl.WebBrowser WebBrowser1 
         Height          =   6315
         Left            =   0
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   0
         Width           =   7125
         ExtentX         =   12568
         ExtentY         =   11139
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   0
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   "http:///"
      End
   End
   Begin VB.CommandButton Command10 
      Caption         =   "跳转到目录(&J)"
      Height          =   375
      Left            =   3600
      TabIndex        =   17
      Top             =   5640
      Width           =   1830
   End
   Begin VB.Timer RegSvr 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   4395
      Top             =   60
   End
   Begin VB.Timer UnRegSvr 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   3045
      Top             =   -75
   End
   Begin VB.CommandButton Command9 
      Caption         =   "退出(ESC)(&X)"
      Height          =   375
      Left            =   5520
      TabIndex        =   16
      Top             =   5640
      Width           =   1575
   End
   Begin VB.CommandButton Command8 
      Caption         =   "选项(&P)"
      Height          =   375
      Left            =   2250
      TabIndex        =   15
      Top             =   5640
      Width           =   1245
   End
   Begin VB.CommandButton Command7 
      Caption         =   "打开Windows功能(&N)"
      Height          =   375
      Left            =   105
      TabIndex        =   14
      Top             =   5640
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      Caption         =   "浏览/注册/反注册控件"
      Height          =   5580
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   7125
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FFFFFF&
         Height          =   3900
         Left            =   60
         ScaleHeight     =   3840
         ScaleWidth      =   6960
         TabIndex        =   19
         Top             =   600
         Visible         =   0   'False
         Width           =   7020
         Begin VB.CommandButton Command13 
            Caption         =   "r"
            BeginProperty Font 
               Name            =   "Marlett"
               Size            =   9
               Charset         =   2
               Weight          =   500
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   6615
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   45
            Width           =   300
         End
         Begin VB.CommandButton Command12 
            Cancel          =   -1  'True
            Caption         =   "取消(&C)"
            Height          =   420
            Left            =   5370
            TabIndex        =   22
            Top             =   3285
            Width           =   1515
         End
         Begin VB.CommandButton Command11 
            Caption         =   "暂停(&P)"
            Height          =   420
            Left            =   3660
            TabIndex        =   21
            Top             =   3285
            Width           =   1515
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "项"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   4335
            TabIndex        =   30
            Top             =   1845
            Width           =   210
         End
         Begin VB.Label Label12 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "666"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   1620
            TabIndex        =   29
            Top             =   1845
            Width           =   2490
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "秒"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   4335
            TabIndex        =   28
            Top             =   1545
            Width           =   210
         End
         Begin VB.Label Label10 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "666"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   1620
            TabIndex        =   27
            Top             =   1545
            Width           =   2490
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "估计剩余项目:"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   75
            TabIndex        =   26
            Top             =   1845
            Width           =   1365
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "估计剩余时间:"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   75
            TabIndex        =   25
            Top             =   1545
            Width           =   1365
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ActiveX Controls Install/Uninstall"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   405
            Left            =   75
            TabIndex        =   24
            Top             =   585
            Width           =   5715
         End
         Begin VB.Shape Shape1 
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00C00000&
            FillColor       =   &H00C00000&
            FillStyle       =   0  'Solid
            Height          =   375
            Left            =   -60
            Top             =   0
            Width           =   7035
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "正在执行操作,请稍候..."
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   75
            TabIndex        =   20
            Top             =   1230
            Width           =   4575
         End
         Begin VB.Shape Shape2 
            BorderColor     =   &H00C0C0C0&
            FillColor       =   &H00C0C0C0&
            FillStyle       =   0  'Solid
            Height          =   1005
            Left            =   0
            Top             =   3090
            Width           =   7110
         End
      End
      Begin VB.CommandButton Command6 
         Caption         =   "刷新(&E)"
         Height          =   420
         Left            =   5160
         TabIndex        =   13
         Top             =   5040
         Width           =   1830
      End
      Begin VB.CommandButton Command5 
         Caption         =   "循环反注册列表中的控件(&O)"
         Height          =   420
         Left            =   2640
         TabIndex        =   12
         Top             =   5040
         Width           =   2415
      End
      Begin VB.CommandButton Command4 
         Caption         =   "反注册(卸载)选定的控件(&U)"
         Height          =   420
         Left            =   120
         TabIndex        =   11
         Top             =   5040
         Width           =   2415
      End
      Begin VB.CommandButton Command3 
         Caption         =   "删除选中的控件(&D)"
         Height          =   420
         Left            =   4920
         TabIndex        =   9
         Top             =   4560
         Width           =   2085
      End
      Begin VB.CommandButton Command2 
         Caption         =   "循环注册列表中的控件(&L)"
         Height          =   420
         Left            =   2280
         TabIndex        =   8
         Top             =   4560
         Width           =   2535
      End
      Begin VB.CommandButton Command1 
         Caption         =   "注册选定的控件(&R)"
         Height          =   420
         Left            =   120
         TabIndex        =   7
         Top             =   4560
         Width           =   2055
      End
      Begin VB.FileListBox File1 
         Height          =   3690
         Hidden          =   -1  'True
         Left            =   3660
         Pattern         =   "*.DLL;*.ocx;*.cpl"
         System          =   -1  'True
         TabIndex        =   6
         Top             =   825
         Width           =   3375
      End
      Begin VB.DirListBox Dir1 
         Height          =   3660
         Left            =   120
         TabIndex        =   4
         Top             =   840
         Width           =   3480
      End
      Begin VB.DriveListBox Drive1 
         Height          =   300
         Left            =   720
         TabIndex        =   2
         Top             =   240
         Width           =   6285
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "文件(*.ocx;*.dll;*.cpl):"
         Height          =   180
         Left            =   3675
         TabIndex        =   5
         Top             =   600
         Width           =   2160
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "目录(文件夹):"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   630
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "磁盘:"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "目录友好查看"
      Height          =   180
      Left            =   7215
      TabIndex        =   31
      Top             =   15
      Width           =   1080
   End
   Begin VB.Label Label5 
      Height          =   495
      Left            =   3000
      TabIndex        =   18
      Top             =   3120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "C:\"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   0
      TabIndex        =   10
      Top             =   6120
      Width           =   7215
   End
   Begin VB.Menu file 
      Caption         =   "文件(&F)"
      Begin VB.Menu reg 
         Caption         =   "导出循环注册当前目录控件批处理文件(&E)"
         Shortcut        =   ^E
         Visible         =   0   'False
      End
      Begin VB.Menu unreg 
         Caption         =   "导出循环反注册当前目录控件批处理文件(&X)"
         Shortcut        =   ^X
         Visible         =   0   'False
      End
      Begin VB.Menu mnuB3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuReg 
         Caption         =   "循环注册控件(&L)"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuUnReg 
         Caption         =   "循环反注册控件(&O)"
         Shortcut        =   ^U
      End
      Begin VB.Menu mnuB2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBrowse 
         Caption         =   "浏览当前目录(&B)"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuJump 
         Caption         =   "跳转(&J)..."
         Shortcut        =   ^J
      End
      Begin VB.Menu mnuB 
         Caption         =   "-"
      End
      Begin VB.Menu mnuToolLink 
         Caption         =   "Windows工具箱(&W)..."
      End
      Begin VB.Menu b2 
         Caption         =   "-"
      End
      Begin VB.Menu exit 
         Caption         =   "退出(&T)"
      End
   End
   Begin VB.Menu hid 
      Caption         =   "hid"
      Visible         =   0   'False
      Begin VB.Menu regsv 
         Caption         =   "注册选中的控件(&G)"
      End
      Begin VB.Menu unin 
         Caption         =   "反注册选中的控件(&I)"
      End
      Begin VB.Menu killfl 
         Caption         =   "删除选中的控件(&K)"
      End
      Begin VB.Menu Explorer 
         Caption         =   "用资源管理器打开选中的控件所在目录(&S)"
      End
      Begin VB.Menu ren 
         Caption         =   "重命名选中的文件(&M)"
      End
   End
   Begin VB.Menu perform 
      Caption         =   "性能(&P)"
      Begin VB.Menu normalspeed 
         Caption         =   "标准速度(&N)"
         Checked         =   -1  'True
      End
      Begin VB.Menu highspeed 
         Caption         =   "高速模式(&H)"
      End
      Begin VB.Menu anti 
         Caption         =   "启用防停止响应功能(&E)"
      End
      Begin VB.Menu b1 
         Caption         =   "-"
      End
      Begin VB.Menu COST 
         Caption         =   "自定义(&C)"
      End
   End
   Begin VB.Menu hlp 
      Caption         =   "帮助(&L)"
      Begin VB.Menu regsvrhlp 
         Caption         =   "Regsvr32.exe指令说明(&R)"
      End
      Begin VB.Menu help 
         Caption         =   "帮助(&H)"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
Private Sub anti_Click()
On Error GoTo ep
Unload Form10
Dim ans As Integer
ans = MsgBox("启用防停止响应模式可以一定程度上避免程序的假死/停止响应等问题,但是会大大减缓循环注册的时间,仅建议在配置较差的计算机上启用此模式,继续吗?", vbExclamation + vbYesNo, "Ask")
If ans = vbYes Then
Me.normalspeed.Checked = False
Me.highspeed.Checked = False
Me.anti.Checked = True
COST.Checked = False
Me.RegSvr.Interval = 100
Me.UnRegSvr.Interval = 100
Form3.slider1.Value = 100
Me.Show
Else
MsgBox "系统的性能模式将返回标准模式", vbInformation, "Info"
On Error Resume Next
Me.normalspeed.Checked = True
Me.highspeed.Checked = False
Me.anti.Checked = False
COST.Checked = False
Me.RegSvr.Interval = 25
Me.UnRegSvr.Interval = 25
Form3.slider1.Value = 25
End If
If 1 = 2 Then
With Form10
.Top = frmMain.WebBrowser1.Top + frmMain.Top + 651
.Left = frmMain.WebBrowser1.Left + frmMain.Left
.Move frmMain.WebBrowser1.Left + frmMain.Left, frmMain.WebBrowser1.Top + frmMain.Top + 651
.Height = frmMain.WebBrowser1.Height
.Width = frmMain.WebBrowser1.Width - 200
.Enabled = False
.Visible = True
.Show
End With
End If
Me.Show
Exit Sub
ep:
MsgBox "发生了错误:" & Chr(13) & Err.Description, vbCritical, "Error"
Me.normalspeed.Checked = True
Me.highspeed.Checked = False
Me.anti.Checked = False
COST.Checked = False
Me.RegSvr.Interval = 25
Me.UnRegSvr.Interval = 25
Form3.slider1.Value = 25
Me.Show
If 1 = 2 Then
With Form10
.Top = frmMain.WebBrowser1.Top + frmMain.Top + 651
.Left = frmMain.WebBrowser1.Left + frmMain.Left
.Move frmMain.WebBrowser1.Left + frmMain.Left, frmMain.WebBrowser1.Top + frmMain.Top + 651
.Height = frmMain.WebBrowser1.Height
.Width = frmMain.WebBrowser1.Width - 200
.Enabled = False
.Visible = True
.Show
End With
End If
End Sub
Private Sub Command1_Click()
On Error GoTo ep
Unload Form10
If File1.ListIndex >= 0 Then
Shell ("regsvr32 " & Chr(34) & Label4.Caption & Chr(34))
Else
MsgBox "您尚未选择文件!", vbCritical, "Error"
End If
If 1 = 2 Then
With Form10
.Top = frmMain.WebBrowser1.Top + frmMain.Top + 651
.Left = frmMain.WebBrowser1.Left + frmMain.Left
.Move frmMain.WebBrowser1.Left + frmMain.Left, frmMain.WebBrowser1.Top + frmMain.Top + 651
.Height = frmMain.WebBrowser1.Height
.Width = frmMain.WebBrowser1.Width - 200
.Enabled = False
.Visible = True
.Show
End With
End If
Exit Sub
ep:
MsgBox "发生了错误:" & Chr(13) & Err.Description, vbCritical, "Error"
End Sub
Private Sub Command10_Click()
On Error GoTo ep
Unload Form10
Dir1.path = GetFolderName(Me.hwnd, "请选择一个文件夹")
If 1 = 2 Then
With Form10
.Top = frmMain.WebBrowser1.Top + frmMain.Top + 651
.Left = frmMain.WebBrowser1.Left + frmMain.Left
.Move frmMain.WebBrowser1.Left + frmMain.Left, frmMain.WebBrowser1.Top + frmMain.Top + 651
.Height = frmMain.WebBrowser1.Height
.Width = frmMain.WebBrowser1.Width - 200
.Enabled = False
.Visible = True
.Show
End With
End If
Exit Sub
ep:
MsgBox "发生应用程序错误:" & Chr(13) & "错误号:" & Err.Number & Chr(13) & "错误描述:" & Chr(13) & "    " & Err.Description & Chr(13) & Chr(13) & Chr(13) & "请检查输入的目录是否有效,且您是否有权访问.", vbCritical, "Error"
End Sub
Private Sub Command11_Click()
On Error Resume Next
Select Case Left(Command11.Caption, 1)
Case "暂"
Me.RegSvr.Enabled = False
Me.UnRegSvr.Enabled = False
Command11.Caption = "继续(&O)"
Label6.Caption = "已经被暂停..."
Case "继"
If regflag = True Then
Me.RegSvr.Enabled = True
Me.UnRegSvr.Enabled = False
Command11.Caption = "暂停(&S)"
Label6.Caption = "正在执行操作,请稍候..."
End If
If unregflag = True Then
Me.UnRegSvr.Enabled = True
Me.RegSvr.Enabled = False
Command11.Caption = "暂停(&S)"
Label6.Caption = "正在执行操作,请稍候..."
End If
End Select
End Sub
Private Sub Command12_Click()
On Error Resume Next
Dim ans As Integer
Unload Form10
ans = MsgBox("确定要停止吗?", vbQuestion + vbYesNo, "Ask")
If ans = vbYes Then
ream = File1.ListCount
Me.WebBrowser1.Navigate Dir1.path
Label10.Caption = Int(ream * Me.RegSvr.Interval / 1000) + 1
Label12.Caption = ream
Label4.Enabled = True
regflag = False
unregflag = False
Me.Picture1.Visible = False
RegSvr.Enabled = False
UnRegSvr.Enabled = False
Command1.Enabled = True
Command3.Enabled = True
Command4.Enabled = True
Command5.Enabled = True
Command6.Enabled = True
Command7.Enabled = True
Command8.Enabled = True
Command9.Enabled = True
Command2.Enabled = True
Command10.Enabled = True
Me.perform.Enabled = True
Dir1.Enabled = True
File1.Enabled = True
Drive1.Enabled = True
file.Enabled = True
hlp.Enabled = True
regsvrvrt = 0
unregsvrvrt = 0
Else
Me.WebBrowser1.Navigate "About:Processing..."
If 1 = 2 Then
With Form10
.Top = frmMain.WebBrowser1.Top + frmMain.Top + 651
.Left = frmMain.WebBrowser1.Left + frmMain.Left
.Move frmMain.WebBrowser1.Left + frmMain.Left, frmMain.WebBrowser1.Top + frmMain.Top + 651
.Height = frmMain.WebBrowser1.Height
.Width = frmMain.WebBrowser1.Width - 200
.Enabled = False
.Visible = True
.Show
End With
End If
Exit Sub
End If
If 1 = 2 Then
With Form10
.Top = frmMain.WebBrowser1.Top + frmMain.Top + 651
.Left = frmMain.WebBrowser1.Left + frmMain.Left
.Move frmMain.WebBrowser1.Left + frmMain.Left, frmMain.WebBrowser1.Top + frmMain.Top + 651
.Height = frmMain.WebBrowser1.Height
.Width = frmMain.WebBrowser1.Width - 200
.Enabled = False
.Visible = True
.Show
End With
End If
End Sub
Private Sub Command13_Click()
On Error Resume Next
Dim ans As Integer
Unload Form10
ans = MsgBox("确定要停止吗?", vbQuestion + vbYesNo, "Ask")
ream = File1.ListCount
Me.WebBrowser1.Navigate Dir1.path
Label10.Caption = Int(ream * Me.RegSvr.Interval / 1000) + 1
Label12.Caption = ream
If ans = vbYes Then
Label4.Enabled = True
regflag = False
unregflag = False
Me.Picture1.Visible = False
RegSvr.Enabled = False
UnRegSvr.Enabled = False
Command1.Enabled = True
Command3.Enabled = True
Command4.Enabled = True
Command5.Enabled = True
Command6.Enabled = True
Command7.Enabled = True
Command8.Enabled = True
Command9.Enabled = True
Command2.Enabled = True
Command10.Enabled = True
Me.perform.Enabled = True
Dir1.Enabled = True
File1.Enabled = True
Drive1.Enabled = True
file.Enabled = True
hlp.Enabled = True
regsvrvrt = 0
unregsvrvrt = 0
Else
Me.WebBrowser1.Navigate "About:Processing..."
If 1 = 2 Then
With Form10
.Top = frmMain.WebBrowser1.Top + frmMain.Top + 651
.Left = frmMain.WebBrowser1.Left + frmMain.Left
.Move frmMain.WebBrowser1.Left + frmMain.Left, frmMain.WebBrowser1.Top + frmMain.Top + 651
.Height = frmMain.WebBrowser1.Height
.Width = frmMain.WebBrowser1.Width - 200
.Enabled = False
.Visible = True
.Show
End With
End If
Exit Sub
End If
If 1 = 2 Then
With Form10
.Top = frmMain.WebBrowser1.Top + frmMain.Top + 651
.Left = frmMain.WebBrowser1.Left + frmMain.Left
.Move frmMain.WebBrowser1.Left + frmMain.Left, frmMain.WebBrowser1.Top + frmMain.Top + 651
.Height = frmMain.WebBrowser1.Height
.Width = frmMain.WebBrowser1.Width - 200
.Enabled = False
.Visible = True
.Show
End With
End If
End Sub
Private Sub COST_Click()
On Error Resume Next
frmMain.Visible = False
Form9.Show
Form9.slider1.SetFocus
End Sub
Private Sub File1_KeyPress(KeyAscii As Integer)
On Error GoTo ep
Unload Form10
If File1.ListCount > 0 Then
If KeyAscii = vbKeyReturn Then
If File1.ListIndex >= 0 Then
Shell ("regsvr32 " & Chr(34) & Label4.Caption & Chr(34))
If 1 = 2 Then
With Form10
.Top = frmMain.WebBrowser1.Top + frmMain.Top + 651
.Left = frmMain.WebBrowser1.Left + frmMain.Left
.Move frmMain.WebBrowser1.Left + frmMain.Left, frmMain.WebBrowser1.Top + frmMain.Top + 651
.Height = frmMain.WebBrowser1.Height
.Width = frmMain.WebBrowser1.Width - 200
.Enabled = False
.Visible = True
.Show
End With
End If
Else
Exit Sub
End If
Else
Exit Sub
End If
Else
Exit Sub
End If
If 1 = 2 Then
With Form10
.Top = frmMain.WebBrowser1.Top + frmMain.Top + 651
.Left = frmMain.WebBrowser1.Left + frmMain.Left
.Move frmMain.WebBrowser1.Left + frmMain.Left, frmMain.WebBrowser1.Top + frmMain.Top + 651
.Height = frmMain.WebBrowser1.Height
.Width = frmMain.WebBrowser1.Width - 200
.Enabled = False
.Visible = True
.Show
End With
End If
Exit Sub
ep:
MsgBox "发生了错误:" & Chr(13) & Err.Description, vbCritical, "Error"
End Sub
Private Sub Form_Activate()
On Error Resume Next
Me.SetFocus
With Me.Picture2
.Enabled = False
End With
If 1 = 2 Then
With Form10
.Top = frmMain.WebBrowser1.Top + frmMain.Top + 651
.Left = frmMain.WebBrowser1.Left + frmMain.Left
.Move frmMain.WebBrowser1.Left + frmMain.Left, frmMain.WebBrowser1.Top + frmMain.Top + 651
.Height = frmMain.WebBrowser1.Height
.Width = frmMain.WebBrowser1.Width - 200
.Enabled = False
.Visible = True
.Show
End With
End If
Me.SetFocus
act = False
End Sub
Private Sub Form_Deactivate()
On Error Resume Next
act = True
End Sub
Private Sub Form_GotFocus()
On Error Resume Next
Exit Sub
If 1 = 2 Then
With Form10
.Top = frmMain.WebBrowser1.Top + frmMain.Top + 651
.Left = frmMain.WebBrowser1.Left + frmMain.Left
.Move frmMain.WebBrowser1.Left + frmMain.Left, frmMain.WebBrowser1.Top + frmMain.Top + 651
.Height = frmMain.WebBrowser1.Height
.Width = frmMain.WebBrowser1.Width - 200
.Enabled = False
.Visible = True
.Show
End With
End If
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = vbKeyEscape Then
Unload Me
End If
End Sub
Private Sub Form_LostFocus()
On Error Resume Next
act = True
End Sub
Private Sub Form_Paint()
On Error Resume Next
If 1 = 2 Then
If frmMain.Visible = True Then Exit Sub
With Form10
.Top = frmMain.WebBrowser1.Top + frmMain.Top + 651
.Left = frmMain.WebBrowser1.Left + frmMain.Left
.Move frmMain.WebBrowser1.Left + frmMain.Left, frmMain.WebBrowser1.Top + frmMain.Top + 651
.Height = frmMain.WebBrowser1.Height
.Width = frmMain.WebBrowser1.Width - 200
.Enabled = False
.Visible = True
.Show
End With
End If
End Sub
Private Sub help_Click()
On Error Resume Next
MsgBox "关于 [控件小助手(Windows Control Helper)]" & Chr(13) & Chr(13) & Chr(13) & "版本:1.0.0" & Chr(13) & Chr(13) & "这个应用程序可以帮助您对Windows中的控件进行注册/反注册/删除/循环注册/循环反注册等操作,而且可以实现对部分Windows自带功能的调用.在文件列表中双击控件文件可以直接注册控件,而右击可以显示高级菜单" & Chr(13) & Chr(13) & Chr(13) & Chr(13) & Chr(13) & "PC-DOS Workshop 出品" & Chr(13) & "版权没有,翻版不究", vbInformation, "Help"
End Sub
Private Sub Command2_Click()
On Error GoTo ep
Unload Form10
If File1.ListCount <= 0 Then
MsgBox "没有可以注册的控件!", vbCritical, "Error"
Label4.Enabled = True
Label6.Caption = "正在执行操作,请稍候..."
ream = File1.ListCount
Label10.Caption = Int(ream * Me.RegSvr.Interval / 1000) + 1
Label12.Caption = ream
Me.WebBrowser1.Navigate Dir1.path
Command11.Caption = "暂停(&S)"
regflag = False
unregflag = False
Me.Picture1.Visible = False
RegSvr.Enabled = False
UnRegSvr.Enabled = False
Command1.Enabled = True
Command3.Enabled = True
Command4.Enabled = True
Command5.Enabled = True
Command6.Enabled = True
Command7.Enabled = True
Command8.Enabled = True
Command9.Enabled = True
Command2.Enabled = True
Command10.Enabled = True
Me.perform.Enabled = True
Dir1.Enabled = True
File1.Enabled = True
Drive1.Enabled = True
file.Enabled = True
hlp.Enabled = True
If 1 = 2 Then
With Form10
.Top = frmMain.WebBrowser1.Top + frmMain.Top + 651
.Left = frmMain.WebBrowser1.Left + frmMain.Left
.Move frmMain.WebBrowser1.Left + frmMain.Left, frmMain.WebBrowser1.Top + frmMain.Top + 651
.Height = frmMain.WebBrowser1.Height
.Width = frmMain.WebBrowser1.Width - 200
.Enabled = False
.Visible = True
.Show
End With
End If
Exit Sub
End If
Label4.Enabled = False
Command11.Caption = "暂停(&S)"
Label6.Caption = "正在执行操作,请稍候..."
Me.WebBrowser1.Navigate "About:Processing..."
ream = File1.ListCount
Label10.Caption = Int(ream * Me.RegSvr.Interval / 1000) + 1
Label12.Caption = ream
regflag = True
unregflag = False
Me.Picture1.Visible = True
Command2.Enabled = False
Command1.Enabled = False
Command3.Enabled = False
Command4.Enabled = False
Command5.Enabled = False
Command6.Enabled = False
Command7.Enabled = False
Command8.Enabled = False
Command9.Enabled = False
Command10.Enabled = False
hlp.Enabled = False
Dir1.Enabled = False
File1.Enabled = False
Drive1.Enabled = False
file.Enabled = False
Me.perform.Enabled = False
RegSvr.Enabled = True
If 1 = 2 Then
With Form10
.Top = frmMain.WebBrowser1.Top + frmMain.Top + 651
.Left = frmMain.WebBrowser1.Left + frmMain.Left
.Move frmMain.WebBrowser1.Left + frmMain.Left, frmMain.WebBrowser1.Top + frmMain.Top + 651
.Height = frmMain.WebBrowser1.Height
.Width = frmMain.WebBrowser1.Width - 200
.Enabled = False
.Visible = True
.Show
End With
With Form10
.Top = frmMain.WebBrowser1.Top + frmMain.Top + 651
.Left = frmMain.WebBrowser1.Left + frmMain.Left
.Move frmMain.WebBrowser1.Left + frmMain.Left, frmMain.WebBrowser1.Top + frmMain.Top + 651
.Height = frmMain.WebBrowser1.Height
.Width = frmMain.WebBrowser1.Width - 200
.Enabled = False
.Visible = True
.Show
End With
End If
Exit Sub
ep:
MsgBox "发生了错误:" & Chr(13) & Err.Description, vbCritical, "Error"
Me.WebBrowser1.Navigate Dir1.path
ream = File1.ListCount
Label10.Caption = Int(ream * Me.RegSvr.Interval / 1000) + 1
Label12.Caption = ream
Label4.Enabled = True
Label6.Caption = "正在执行操作,请稍候..."
Command11.Caption = "暂停(&S)"
regflag = False
unregflag = False
Me.Picture1.Visible = False
RegSvr.Enabled = False
UnRegSvr.Enabled = False
Command1.Enabled = True
Command3.Enabled = True
Command4.Enabled = True
Command5.Enabled = True
Command6.Enabled = True
Command7.Enabled = True
Command8.Enabled = True
Command9.Enabled = True
Command2.Enabled = True
Command10.Enabled = True
Me.perform.Enabled = True
Dir1.Enabled = True
File1.Enabled = True
Drive1.Enabled = True
file.Enabled = True
hlp.Enabled = True
If 1 = 2 Then
With Form10
.Top = frmMain.WebBrowser1.Top + frmMain.Top + 651
.Left = frmMain.WebBrowser1.Left + frmMain.Left
.Move frmMain.WebBrowser1.Left + frmMain.Left, frmMain.WebBrowser1.Top + frmMain.Top + 651
.Height = frmMain.WebBrowser1.Height
.Width = frmMain.WebBrowser1.Width - 200
.Enabled = False
.Visible = True
.Show
End With
End If
End Sub
Private Sub Command3_Click()
On Error GoTo ep
Unload Form10
Dim a As Integer
If File1.ListIndex >= 0 Then
a = MsgBox("确定删除这个文件吗?删除后不可恢复!", vbExclamation + vbYesNo, "Question")
If a = vbYes Then
Kill (Label4.Caption)
File1.Refresh
Call Dir1_Change
End If
Else
MsgBox "您尚未选择文件!", vbCritical, "Error"
End If
If 1 = 2 Then
With Form10
.Top = frmMain.WebBrowser1.Top + frmMain.Top + 651
.Left = frmMain.WebBrowser1.Left + frmMain.Left
.Move frmMain.WebBrowser1.Left + frmMain.Left, frmMain.WebBrowser1.Top + frmMain.Top + 651
.Height = frmMain.WebBrowser1.Height
.Width = frmMain.WebBrowser1.Width - 200
.Enabled = False
.Visible = True
.Show
End With
End If
Exit Sub
ep:
MsgBox "发生了错误:" & Chr(13) & Err.Description, vbCritical, "Error"
File1.Refresh
If 1 = 2 Then
With Form10
.Top = frmMain.WebBrowser1.Top + frmMain.Top + 651
.Left = frmMain.WebBrowser1.Left + frmMain.Left
.Move frmMain.WebBrowser1.Left + frmMain.Left, frmMain.WebBrowser1.Top + frmMain.Top + 651
.Height = frmMain.WebBrowser1.Height
.Width = frmMain.WebBrowser1.Width - 200
.Enabled = False
.Visible = True
.Show
End With
End If
End Sub
Private Sub Command4_Click()
On Error GoTo ep
Unload Form10
If File1.ListIndex >= 0 Then
Shell ("regsvr32 /u " & Chr(34) & Label4.Caption & Chr(34))
Else
MsgBox "您尚未选择文件!", vbCritical, "Error"
End If
If 1 = 2 Then
With Form10
.Top = frmMain.WebBrowser1.Top + frmMain.Top + 651
.Left = frmMain.WebBrowser1.Left + frmMain.Left
.Move frmMain.WebBrowser1.Left + frmMain.Left, frmMain.WebBrowser1.Top + frmMain.Top + 651
.Height = frmMain.WebBrowser1.Height
.Width = frmMain.WebBrowser1.Width - 200
.Enabled = False
.Visible = True
.Show
End With
End If
Exit Sub
ep:
MsgBox "发生了错误:" & Chr(13) & Err.Description, vbCritical, "Error"
End Sub
Private Sub Command5_Click()
On Error GoTo ep
Unload Form10
If File1.ListCount <= 0 Then
MsgBox "没有可以反注册的控件!", vbCritical, "Error"
ream = File1.ListCount
Me.WebBrowser1.Navigate Dir1.path
Label10.Caption = Int(ream * Me.RegSvr.Interval / 1000) + 1
Label12.Caption = ream
Label4.Enabled = True
Label6.Caption = "正在执行操作,请稍候..."
Command11.Caption = "暂停(&S)"
regflag = False
unregflag = False
Me.Picture1.Visible = False
RegSvr.Enabled = False
UnRegSvr.Enabled = False
Command1.Enabled = True
Command3.Enabled = True
Command4.Enabled = True
Command5.Enabled = True
Command6.Enabled = True
Command7.Enabled = True
Command8.Enabled = True
Command9.Enabled = True
Command2.Enabled = True
Command10.Enabled = True
Me.perform.Enabled = True
Dir1.Enabled = True
File1.Enabled = True
Drive1.Enabled = True
file.Enabled = True
hlp.Enabled = True
If 1 = 2 Then
With Form10
.Top = frmMain.WebBrowser1.Top + frmMain.Top + 651
.Left = frmMain.WebBrowser1.Left + frmMain.Left
.Move frmMain.WebBrowser1.Left + frmMain.Left, frmMain.WebBrowser1.Top + frmMain.Top + 651
.Height = frmMain.WebBrowser1.Height
.Width = frmMain.WebBrowser1.Width - 200
.Enabled = False
.Visible = True
.Show
End With
End If
Exit Sub
End If
Label4.Enabled = False
Label6.Caption = "正在执行操作,请稍候..."
ream = File1.ListCount
Me.WebBrowser1.Navigate "About:Processing..."
Label10.Caption = Int(ream * Me.RegSvr.Interval / 1000) + 1
Label12.Caption = ream
Command11.Caption = "暂停(&S)"
regflag = False
unregflag = True
Me.Picture1.Visible = True
Me.perform.Enabled = False
Command2.Enabled = False
Command1.Enabled = False
Command3.Enabled = False
Command5.Enabled = False
Command6.Enabled = False
Command4.Enabled = False
Command7.Enabled = False
Command8.Enabled = False
Command9.Enabled = False
Command10.Enabled = False
Dir1.Enabled = False
File1.Enabled = False
Drive1.Enabled = False
file.Enabled = False
hlp.Enabled = False
If 1 = 2 Then
With Form10
.Top = frmMain.WebBrowser1.Top + frmMain.Top + 651
.Left = frmMain.WebBrowser1.Left + frmMain.Left
.Move frmMain.WebBrowser1.Left + frmMain.Left, frmMain.WebBrowser1.Top + frmMain.Top + 651
.Height = frmMain.WebBrowser1.Height
.Width = frmMain.WebBrowser1.Width - 200
.Enabled = False
.Visible = True
.Show
End With
End If
UnRegSvr.Enabled = True
Exit Sub
ep:
MsgBox "发生了错误:" & Chr(13) & Err.Description, vbCritical, "Error"
Me.WebBrowser1.Navigate Dir1.path
ream = File1.ListCount
Label10.Caption = Int(ream * Me.RegSvr.Interval / 1000) + 1
Label12.Caption = ream
Label4.Enabled = True
Label6.Caption = "正在执行操作,请稍候..."
Command11.Caption = "暂停(&S)"
regflag = False
unregflag = False
Me.Picture1.Visible = False
RegSvr.Enabled = False
UnRegSvr.Enabled = False
Command1.Enabled = True
Command3.Enabled = True
Command4.Enabled = True
Command5.Enabled = True
Command6.Enabled = True
Command7.Enabled = True
Command8.Enabled = True
Command9.Enabled = True
Command2.Enabled = True
Command10.Enabled = True
Me.perform.Enabled = True
Dir1.Enabled = True
File1.Enabled = True
Drive1.Enabled = True
file.Enabled = True
If 1 = 2 Then
With Form10
.Top = frmMain.WebBrowser1.Top + frmMain.Top + 651
.Left = frmMain.WebBrowser1.Left + frmMain.Left
.Move frmMain.WebBrowser1.Left + frmMain.Left, frmMain.WebBrowser1.Top + frmMain.Top + 651
.Height = frmMain.WebBrowser1.Height
.Width = frmMain.WebBrowser1.Width - 200
.Enabled = False
.Visible = True
.Show
End With
End If
End Sub
Private Sub Command6_Click()
On Error Resume Next
File1.Refresh
End Sub
Private Sub Command7_Click()
On Error Resume Next
frmMain.Visible = False
Me.Hide
frmTool.Show
End Sub
Private Sub Command8_Click()
On Error Resume Next
frmMain.Visible = False
Form3.Show
End Sub
Private Sub Command9_Click()
On Error Resume Next
End
End Sub
Private Sub Dir1_Change()
On Error GoTo ep
File1.path = Dir1.path
If Right(Dir1.path, 1) <> "\" Then
Label4.Caption = Dir1.path & "\"
ream = File1.ListCount
Me.WebBrowser1.Navigate Dir1.path
Label10.Caption = Int(ream * Me.RegSvr.Interval / 1000) + 1
Label12.Caption = ream
Else
Label4.Caption = Dir1.path
ream = File1.ListCount
Me.WebBrowser1.Navigate Dir1.path
Label10.Caption = Int(ream * Me.RegSvr.Interval / 1000) + 1
Label12.Caption = ream
End If
Exit Sub
ep:
MsgBox "发生了错误:" & Chr(13) & Err.Description, vbCritical, "Error"
End Sub
Private Sub Drive1_Change()
On Error GoTo ep
Dir1.path = Drive1.Drive
Me.WebBrowser1.Navigate Dir1.path
Exit Sub
ep:
MsgBox "发生了错误:" & Chr(13) & Err.Description, vbCritical, "Error"
Drive1.Drive = "C:"
End Sub
Private Sub exit_Click()
On Error Resume Next
End
End Sub
Private Sub Explorer_Click()
On Error GoTo ep
Unload Form10
Shell "explorer " & Dir1.path, vbNormalFocus
Exit Sub
ep:
MsgBox "发生了错误:" & Chr(13) & Err.Description, vbCritical, "Error"
End Sub
Private Sub File1_Click()
On Error GoTo ep
If Right(Dir1.path, 1) <> "\" Then
Label4.Caption = Dir1.path & "\" & File1.FileName
Else
Label4.Caption = Dir1.path & File1.FileName
End If
Exit Sub
ep:
MsgBox "发生了错误:" & Chr(13) & Err.Description, vbCritical, "Error"
End Sub
Private Sub File1_DblClick()
On Error GoTo ep
Unload Form10
If File1.ListCount > 0 Then
If File1.ListIndex >= 0 Then
Shell ("regsvr32 " & Chr(34) & Label4.Caption & Chr(34))
Else
MsgBox "您尚未选择文件!", vbCritical, "Error"
End If
Else
If 1 = 2 Then
With Form10
.Top = frmMain.WebBrowser1.Top + frmMain.Top + 651
.Left = frmMain.WebBrowser1.Left + frmMain.Left
.Move frmMain.WebBrowser1.Left + frmMain.Left, frmMain.WebBrowser1.Top + frmMain.Top + 651
.Height = frmMain.WebBrowser1.Height
.Width = frmMain.WebBrowser1.Width - 200
.Enabled = False
.Visible = True
.Show
End With
End If
Exit Sub
End If
If 1 = 2 Then
With Form10
.Top = frmMain.WebBrowser1.Top + frmMain.Top + 651
.Left = frmMain.WebBrowser1.Left + frmMain.Left
.Move frmMain.WebBrowser1.Left + frmMain.Left, frmMain.WebBrowser1.Top + frmMain.Top + 651
.Height = frmMain.WebBrowser1.Height
.Width = frmMain.WebBrowser1.Width - 200
.Enabled = False
.Visible = True
.Show
End With
End If
Exit Sub
ep:
MsgBox "发生了错误:" & Chr(13) & Err.Description, vbCritical, "Error"
End Sub
Private Sub File1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If File1.ListIndex >= 0 Then
If Button = 2 Then PopupMenu hid
Else
Exit Sub
End If
End Sub
Private Sub Form_Load()
On Error GoTo ep
If App.PrevInstance = False Then
With Picture2
.Enabled = False
End With
If 1 = 2 Then
With Form10
.Top = frmMain.WebBrowser1.Top + frmMain.Top + 651
.Left = frmMain.WebBrowser1.Left + frmMain.Left
.Move frmMain.WebBrowser1.Left + frmMain.Left, frmMain.WebBrowser1.Top + frmMain.Top + 651
.Height = frmMain.WebBrowser1.Height - 200
.Width = frmMain.WebBrowser1.Width
.Enabled = False
.Show
End With
End If
Me.KeyPreview = True
File1.path = Dir1.path
If Right(File1.path, 1) <> "\" Then
ream = File1.ListCount
Me.WebBrowser1.Navigate Dir1.path
Label4.Caption = File1.path & "\"
Else
ream = File1.ListCount
Me.WebBrowser1.Navigate Dir1.path
Label4.Caption = File1.path
End If
Me.normalspeed.Checked = True
Me.highspeed.Checked = False
Me.anti.Checked = False
Me.RegSvr.Interval = 25
Me.UnRegSvr.Interval = 25
Else
MsgBox "本程序不允许同时执行2个及以上的实例,程序即将退出...", vbCritical, "Info"
End
End If
Exit Sub
ep:
MsgBox "发生了错误:" & Chr(13) & Err.Description, vbCritical, "Error"
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
End Sub
Private Sub highspeed_Click()
On Error GoTo ep
Unload Form10
Dim ans As Integer
ans = MsgBox("启用高速模式可以加快循环注册控件的速度,但是不建议在配置较差的计算机上启用此模式,因为这可能引起程序停止响应,继续吗?", vbExclamation + vbYesNo, "Ask")
If ans = vbYes Then
Me.normalspeed.Checked = False
Me.highspeed.Checked = True
Me.anti.Checked = False
COST.Checked = False
Me.RegSvr.Interval = 10
Me.UnRegSvr.Interval = 10
Form3.slider1.Value = 10
Me.Show
Else
MsgBox "系统的性能模式将返回标准模式", vbInformation, "Info"
On Error Resume Next
Me.normalspeed.Checked = True
Me.highspeed.Checked = False
Me.anti.Checked = False
COST.Checked = False
Me.RegSvr.Interval = 25
Me.UnRegSvr.Interval = 25
Form3.slider1.Value = 25
Me.Show
End If
If 1 = 2 Then
With Form10
.Top = frmMain.WebBrowser1.Top + frmMain.Top + 651
.Left = frmMain.WebBrowser1.Left + frmMain.Left
.Move frmMain.WebBrowser1.Left + frmMain.Left, frmMain.WebBrowser1.Top + frmMain.Top + 651
.Height = frmMain.WebBrowser1.Height
.Width = frmMain.WebBrowser1.Width - 200
.Enabled = False
.Visible = True
.Show
End With
End If
Exit Sub
ep:
MsgBox "发生了错误:" & Chr(13) & Err.Description, vbCritical, "Error"
Me.normalspeed.Checked = True
Me.highspeed.Checked = False
Me.anti.Checked = False
COST.Checked = False
Me.RegSvr.Interval = 25
Me.UnRegSvr.Interval = 25
Form3.slider1.Value = 25
Me.Show
If 1 = 2 Then
With Form10
.Top = frmMain.WebBrowser1.Top + frmMain.Top + 651
.Left = frmMain.WebBrowser1.Left + frmMain.Left
.Move frmMain.WebBrowser1.Left + frmMain.Left, frmMain.WebBrowser1.Top + frmMain.Top + 651
.Height = frmMain.WebBrowser1.Height
.Width = frmMain.WebBrowser1.Width - 200
.Enabled = False
.Visible = True
.Show
End With
End If
End Sub
Private Sub killfl_Click()
On Error GoTo ep
Unload Form10
Dim a As Integer
If File1.ListIndex >= 0 Then
a = MsgBox("确定删除这个文件吗?删除后不可恢复!", vbExclamation + vbYesNo, "Question")
If a = vbYes Then Kill (Label4.Caption)
File1.Refresh
Call Dir1_Change
Else
MsgBox "您尚未选择文件!", vbCritical, "Error"
End If
If 1 = 2 Then
With Form10
.Top = frmMain.WebBrowser1.Top + frmMain.Top + 651
.Left = frmMain.WebBrowser1.Left + frmMain.Left
.Move frmMain.WebBrowser1.Left + frmMain.Left, frmMain.WebBrowser1.Top + frmMain.Top + 651
.Height = frmMain.WebBrowser1.Height
.Width = frmMain.WebBrowser1.Width - 200
.Enabled = False
.Visible = True
.Show
End With
End If
Exit Sub
ep:
MsgBox "发生了错误:" & Chr(13) & Err.Description, vbCritical, "Error"
File1.Refresh
If 1 = 2 Then
With Form10
.Top = frmMain.WebBrowser1.Top + frmMain.Top + 651
.Left = frmMain.WebBrowser1.Left + frmMain.Left
.Move frmMain.WebBrowser1.Left + frmMain.Left, frmMain.WebBrowser1.Top + frmMain.Top + 651
.Height = frmMain.WebBrowser1.Height
.Width = frmMain.WebBrowser1.Width - 200
.Enabled = False
.Visible = True
.Show
End With
End If
End Sub
Private Sub Label4_Click()
On Error Resume Next
MsgBox "您当前的位置(目录+文件(如果存在)):" & Chr(13) & Label4.Caption, vbInformation, "Info"
End Sub
Private Sub Label4_DblClick()
On Error GoTo ep
Unload Form10
Dir1.path = GetFolderName(Me.hwnd, "请选择一个文件夹")
If 1 = 2 Then
With Form10
.Top = frmMain.WebBrowser1.Top + frmMain.Top + 651
.Left = frmMain.WebBrowser1.Left + frmMain.Left
.Move frmMain.WebBrowser1.Left + frmMain.Left, frmMain.WebBrowser1.Top + frmMain.Top + 651
.Height = frmMain.WebBrowser1.Height
.Width = frmMain.WebBrowser1.Width - 200
.Enabled = False
.Visible = True
.Show
End With
End If
Exit Sub
ep:
MsgBox "发生应用程序错误:" & Chr(13) & "错误号:" & Err.Number & Chr(13) & "错误描述:" & Chr(13) & "    " & Err.Description & Chr(13) & Chr(13) & Chr(13) & "请检查输入的目录是否有效,且您是否有权访问.", vbCritical, "Error"
If 1 = 2 Then
With Form10
.Top = frmMain.WebBrowser1.Top + frmMain.Top + 651
.Left = frmMain.WebBrowser1.Left + frmMain.Left
.Move frmMain.WebBrowser1.Left + frmMain.Left, frmMain.WebBrowser1.Top + frmMain.Top + 651
.Height = frmMain.WebBrowser1.Height
.Width = frmMain.WebBrowser1.Width - 200
.Enabled = False
.Visible = True
.Show
End With
End If
End Sub
Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Label4.ToolTipText = Label4.Caption
End Sub
Private Sub mnuBrowse_Click()
On Error Resume Next
Shell "Explorer.exe" & " " & Me.Dir1.path, vbNormalFocus
End Sub
Private Sub mnuJump_Click()
On Error GoTo ep
Unload Form10
Dir1.path = GetFolderName(Me.hwnd, "请选择一个文件夹")
If 1 = 2 Then
With Form10
.Top = frmMain.WebBrowser1.Top + frmMain.Top + 651
.Left = frmMain.WebBrowser1.Left + frmMain.Left
.Move frmMain.WebBrowser1.Left + frmMain.Left, frmMain.WebBrowser1.Top + frmMain.Top + 651
.Height = frmMain.WebBrowser1.Height
.Width = frmMain.WebBrowser1.Width - 200
.Enabled = False
.Visible = True
.Show
End With
End If
Exit Sub
ep:
MsgBox "发生应用程序错误:" & Chr(13) & "错误号:" & Err.Number & Chr(13) & "错误描述:" & Chr(13) & "    " & Err.Description & Chr(13) & Chr(13) & Chr(13) & "请检查输入的目录是否有效,且您是否有权访问.", vbCritical, "Error"
End Sub
Private Sub mnuReg_Click()
On Error GoTo ep
Unload Form10
If File1.ListCount <= 0 Then
MsgBox "没有可以注册的控件!", vbCritical, "Error"
Label4.Enabled = True
Label6.Caption = "正在执行操作,请稍候..."
ream = File1.ListCount
Label10.Caption = Int(ream * Me.RegSvr.Interval / 1000) + 1
Label12.Caption = ream
Me.WebBrowser1.Navigate Dir1.path
Command11.Caption = "暂停(&S)"
regflag = False
unregflag = False
Me.Picture1.Visible = False
RegSvr.Enabled = False
UnRegSvr.Enabled = False
Command1.Enabled = True
Command3.Enabled = True
Command4.Enabled = True
Command5.Enabled = True
Command6.Enabled = True
Command7.Enabled = True
Command8.Enabled = True
Command9.Enabled = True
Command2.Enabled = True
Command10.Enabled = True
Me.perform.Enabled = True
Dir1.Enabled = True
File1.Enabled = True
Drive1.Enabled = True
file.Enabled = True
hlp.Enabled = True
If 1 = 2 Then
With Form10
.Top = frmMain.WebBrowser1.Top + frmMain.Top + 651
.Left = frmMain.WebBrowser1.Left + frmMain.Left
.Move frmMain.WebBrowser1.Left + frmMain.Left, frmMain.WebBrowser1.Top + frmMain.Top + 651
.Height = frmMain.WebBrowser1.Height
.Width = frmMain.WebBrowser1.Width - 200
.Enabled = False
.Visible = True
.Show
End With
End If
Exit Sub
End If
Label4.Enabled = False
Command11.Caption = "暂停(&S)"
Label6.Caption = "正在执行操作,请稍候..."
Me.WebBrowser1.Navigate "About:Processing..."
ream = File1.ListCount
Label10.Caption = Int(ream * Me.RegSvr.Interval / 1000) + 1
Label12.Caption = ream
regflag = True
unregflag = False
Me.Picture1.Visible = True
Command2.Enabled = False
Command1.Enabled = False
Command3.Enabled = False
Command4.Enabled = False
Command5.Enabled = False
Command6.Enabled = False
Command7.Enabled = False
Command8.Enabled = False
Command9.Enabled = False
Command10.Enabled = False
hlp.Enabled = False
Dir1.Enabled = False
File1.Enabled = False
Drive1.Enabled = False
file.Enabled = False
Me.perform.Enabled = False
RegSvr.Enabled = True
If 1 = 2 Then
With Form10
.Top = frmMain.WebBrowser1.Top + frmMain.Top + 651
.Left = frmMain.WebBrowser1.Left + frmMain.Left
.Move frmMain.WebBrowser1.Left + frmMain.Left, frmMain.WebBrowser1.Top + frmMain.Top + 651
.Height = frmMain.WebBrowser1.Height
.Width = frmMain.WebBrowser1.Width - 200
.Enabled = False
.Visible = True
.Show
End With
With Form10
.Top = frmMain.WebBrowser1.Top + frmMain.Top + 651
.Left = frmMain.WebBrowser1.Left + frmMain.Left
.Move frmMain.WebBrowser1.Left + frmMain.Left, frmMain.WebBrowser1.Top + frmMain.Top + 651
.Height = frmMain.WebBrowser1.Height
.Width = frmMain.WebBrowser1.Width - 200
.Enabled = False
.Visible = True
.Show
End With
End If
Exit Sub
ep:
MsgBox "发生了错误:" & Chr(13) & Err.Description, vbCritical, "Error"
Me.WebBrowser1.Navigate Dir1.path
ream = File1.ListCount
Label10.Caption = Int(ream * Me.RegSvr.Interval / 1000) + 1
Label12.Caption = ream
Label4.Enabled = True
Label6.Caption = "正在执行操作,请稍候..."
Command11.Caption = "暂停(&S)"
regflag = False
unregflag = False
Me.Picture1.Visible = False
RegSvr.Enabled = False
UnRegSvr.Enabled = False
Command1.Enabled = True
Command3.Enabled = True
Command4.Enabled = True
Command5.Enabled = True
Command6.Enabled = True
Command7.Enabled = True
Command8.Enabled = True
Command9.Enabled = True
Command2.Enabled = True
Command10.Enabled = True
Me.perform.Enabled = True
Dir1.Enabled = True
File1.Enabled = True
Drive1.Enabled = True
file.Enabled = True
hlp.Enabled = True
If 1 = 2 Then
With Form10
.Top = frmMain.WebBrowser1.Top + frmMain.Top + 651
.Left = frmMain.WebBrowser1.Left + frmMain.Left
.Move frmMain.WebBrowser1.Left + frmMain.Left, frmMain.WebBrowser1.Top + frmMain.Top + 651
.Height = frmMain.WebBrowser1.Height
.Width = frmMain.WebBrowser1.Width - 200
.Enabled = False
.Visible = True
.Show
End With
End If
End Sub
Private Sub mnuToolLink_Click()
On Error Resume Next
frmMain.Visible = False
Me.Hide
frmTool.Show
End Sub
Private Sub mnuUnReg_Click()
On Error GoTo ep
Unload Form10
If File1.ListCount <= 0 Then
MsgBox "没有可以反注册的控件!", vbCritical, "Error"
ream = File1.ListCount
Me.WebBrowser1.Navigate Dir1.path
Label10.Caption = Int(ream * Me.RegSvr.Interval / 1000) + 1
Label12.Caption = ream
Label4.Enabled = True
Label6.Caption = "正在执行操作,请稍候..."
Command11.Caption = "暂停(&S)"
regflag = False
unregflag = False
Me.Picture1.Visible = False
RegSvr.Enabled = False
UnRegSvr.Enabled = False
Command1.Enabled = True
Command3.Enabled = True
Command4.Enabled = True
Command5.Enabled = True
Command6.Enabled = True
Command7.Enabled = True
Command8.Enabled = True
Command9.Enabled = True
Command2.Enabled = True
Command10.Enabled = True
Me.perform.Enabled = True
Dir1.Enabled = True
File1.Enabled = True
Drive1.Enabled = True
file.Enabled = True
hlp.Enabled = True
If 1 = 2 Then
With Form10
.Top = frmMain.WebBrowser1.Top + frmMain.Top + 651
.Left = frmMain.WebBrowser1.Left + frmMain.Left
.Move frmMain.WebBrowser1.Left + frmMain.Left, frmMain.WebBrowser1.Top + frmMain.Top + 651
.Height = frmMain.WebBrowser1.Height
.Width = frmMain.WebBrowser1.Width - 200
.Enabled = False
.Visible = True
.Show
End With
End If
Exit Sub
End If
Label4.Enabled = False
Label6.Caption = "正在执行操作,请稍候..."
ream = File1.ListCount
Me.WebBrowser1.Navigate "About:Processing..."
Label10.Caption = Int(ream * Me.RegSvr.Interval / 1000) + 1
Label12.Caption = ream
Command11.Caption = "暂停(&S)"
regflag = False
unregflag = True
Me.Picture1.Visible = True
Me.perform.Enabled = False
Command2.Enabled = False
Command1.Enabled = False
Command3.Enabled = False
Command5.Enabled = False
Command6.Enabled = False
Command4.Enabled = False
Command7.Enabled = False
Command8.Enabled = False
Command9.Enabled = False
Command10.Enabled = False
Dir1.Enabled = False
File1.Enabled = False
Drive1.Enabled = False
file.Enabled = False
hlp.Enabled = False
If 1 = 2 Then
With Form10
.Top = frmMain.WebBrowser1.Top + frmMain.Top + 651
.Left = frmMain.WebBrowser1.Left + frmMain.Left
.Move frmMain.WebBrowser1.Left + frmMain.Left, frmMain.WebBrowser1.Top + frmMain.Top + 651
.Height = frmMain.WebBrowser1.Height
.Width = frmMain.WebBrowser1.Width - 200
.Enabled = False
.Visible = True
.Show
End With
End If
UnRegSvr.Enabled = True
Exit Sub
ep:
MsgBox "发生了错误:" & Chr(13) & Err.Description, vbCritical, "Error"
Me.WebBrowser1.Navigate Dir1.path
ream = File1.ListCount
Label10.Caption = Int(ream * Me.RegSvr.Interval / 1000) + 1
Label12.Caption = ream
Label4.Enabled = True
Label6.Caption = "正在执行操作,请稍候..."
Command11.Caption = "暂停(&S)"
regflag = False
unregflag = False
Me.Picture1.Visible = False
RegSvr.Enabled = False
UnRegSvr.Enabled = False
Command1.Enabled = True
Command3.Enabled = True
Command4.Enabled = True
Command5.Enabled = True
Command6.Enabled = True
Command7.Enabled = True
Command8.Enabled = True
Command9.Enabled = True
Command2.Enabled = True
Command10.Enabled = True
Me.perform.Enabled = True
Dir1.Enabled = True
File1.Enabled = True
Drive1.Enabled = True
file.Enabled = True
If 1 = 2 Then
With Form10
.Top = frmMain.WebBrowser1.Top + frmMain.Top + 651
.Left = frmMain.WebBrowser1.Left + frmMain.Left
.Move frmMain.WebBrowser1.Left + frmMain.Left, frmMain.WebBrowser1.Top + frmMain.Top + 651
.Height = frmMain.WebBrowser1.Height
.Width = frmMain.WebBrowser1.Width - 200
.Enabled = False
.Visible = True
.Show
End With
End If
End Sub
Private Sub normalspeed_Click()
On Error GoTo ep
Unload Form10
Me.normalspeed.Checked = True
Me.highspeed.Checked = False
Me.anti.Checked = False
Me.COST.Checked = False
Me.RegSvr.Interval = 25
Me.UnRegSvr.Interval = 25
Form3.slider1.Value = 25
Me.Show
If 1 = 2 Then
With Form10
.Top = frmMain.WebBrowser1.Top + frmMain.Top + 651
.Left = frmMain.WebBrowser1.Left + frmMain.Left
.Move frmMain.WebBrowser1.Left + frmMain.Left, frmMain.WebBrowser1.Top + frmMain.Top + 651
.Height = frmMain.WebBrowser1.Height
.Width = frmMain.WebBrowser1.Width - 200
.Enabled = False
.Visible = True
.Show
End With
End If
Exit Sub
ep:
MsgBox "发生了错误:" & Chr(13) & Err.Description, vbCritical, "Error"
Me.normalspeed.Checked = True
Me.highspeed.Checked = False
Me.anti.Checked = False
COST.Checked = False
Me.RegSvr.Interval = 25
Me.UnRegSvr.Interval = 25
Form3.slider1.Value = 25
Me.Show
If 1 = 2 Then
With Form10
.Top = frmMain.WebBrowser1.Top + frmMain.Top + 651
.Left = frmMain.WebBrowser1.Left + frmMain.Left
.Move frmMain.WebBrowser1.Left + frmMain.Left, frmMain.WebBrowser1.Top + frmMain.Top + 651
.Height = frmMain.WebBrowser1.Height
.Width = frmMain.WebBrowser1.Width - 200
.Enabled = False
.Visible = True
.Show
End With
End If
End Sub
Private Sub reg_Click()
On Error GoTo Error
Unload Form10
Dim filewrite
Dim ans As Integer
If File1.ListCount = 0 Then
MsgBox "没有可以注册的控件,结束操作.", vbCritical, "Error"
If 1 = 2 Then
With Form10
.Top = frmMain.WebBrowser1.Top + frmMain.Top + 651
.Left = frmMain.WebBrowser1.Left + frmMain.Left
.Move frmMain.WebBrowser1.Left + frmMain.Left, frmMain.WebBrowser1.Top + frmMain.Top + 651
.Height = frmMain.WebBrowser1.Height
.Width = frmMain.WebBrowser1.Width - 200
.Enabled = False
.Visible = True
.Show
End With
End If
Exit Sub
End If
Exit Sub
Error:
If Err.Number = 32755 Then Exit Sub
MsgBox "发生错误:" & Chr(13) & Err.Description, vbCritical, "Error"
If 1 = 2 Then
With Form10
.Top = frmMain.WebBrowser1.Top + frmMain.Top + 651
.Left = frmMain.WebBrowser1.Left + frmMain.Left
.Move frmMain.WebBrowser1.Left + frmMain.Left, frmMain.WebBrowser1.Top + frmMain.Top + 651
.Height = frmMain.WebBrowser1.Height
.Width = frmMain.WebBrowser1.Width - 200
.Enabled = False
.Visible = True
.Show
End With
End If
End Sub
Private Sub regsv_Click()
On Error GoTo ep
Unload Form10
If File1.ListIndex >= 0 Then
Shell ("regsvr32 " & Chr(34) & Label4.Caption & Chr(34)), vbNormalFocus
Else
MsgBox "您尚未选择文件!", vbCritical, "Error"
End If
If 1 = 2 Then
With Form10
.Top = frmMain.WebBrowser1.Top + frmMain.Top + 651
.Left = frmMain.WebBrowser1.Left + frmMain.Left
.Move frmMain.WebBrowser1.Left + frmMain.Left, frmMain.WebBrowser1.Top + frmMain.Top + 651
.Height = frmMain.WebBrowser1.Height
.Width = frmMain.WebBrowser1.Width - 200
.Enabled = False
.Visible = True
.Show
End With
End If
Exit Sub
ep:
MsgBox "发生了错误:" & Chr(13) & Err.Description, vbCritical, "Error"
If 1 = 2 Then
With Form10
.Top = frmMain.WebBrowser1.Top + frmMain.Top + 651
.Left = frmMain.WebBrowser1.Left + frmMain.Left
.Move frmMain.WebBrowser1.Left + frmMain.Left, frmMain.WebBrowser1.Top + frmMain.Top + 651
.Height = frmMain.WebBrowser1.Height
.Width = frmMain.WebBrowser1.Width - 200
.Enabled = False
.Visible = True
.Show
End With
End If
End Sub
Private Sub RegSvr_Timer()
On Error GoTo ep
DoEvents
regflag = True
unregflag = False
Me.Picture1.Visible = True
Shell ("regsvr32.exe /s " & Chr(34) & Dir1.path & "\" & File1.List(regsvrvrt) & Chr(34))
ream = ream - 1
Label10.Caption = Int(ream * Me.RegSvr.Interval / 1000) + 1
Label12.Caption = ream
regsvrvrt = regsvrvrt + 1
If regsvrvrt = File1.ListCount Then
Me.WebBrowser1.Navigate Dir1.path
ream = File1.ListCount
Label10.Caption = Int(ream * Me.RegSvr.Interval / 1000) + 1
Label12.Caption = ream
Label4.Enabled = True
Label6.Caption = "正在执行操作,请稍候..."
Command11.Caption = "暂停(&S)"
regflag = False
unregflag = False
Me.Picture1.Visible = False
RegSvr.Enabled = False
UnRegSvr.Enabled = False
Command1.Enabled = True
Command3.Enabled = True
Command4.Enabled = True
Command5.Enabled = True
Command6.Enabled = True
Command7.Enabled = True
Command8.Enabled = True
Command9.Enabled = True
Command2.Enabled = True
Command10.Enabled = True
Me.perform.Enabled = True
Dir1.Enabled = True
File1.Enabled = True
Drive1.Enabled = True
file.Enabled = True
hlp.Enabled = True
MsgBox "注册控件成功!如果仍然无法正常调用,请检查:" & Chr(13) & "1:控件完整性(是否有效)" & Chr(13) & "2:是否支持当前系统" & Chr(13) & "3:是否有权访问" & Chr(13) & "4:是否正被其它程序调用", vbExclamation, "Info"
regsvrvrt = 0
If 1 = 2 Then
With Form10
.Top = frmMain.WebBrowser1.Top + frmMain.Top + 651
.Left = frmMain.WebBrowser1.Left + frmMain.Left
.Move frmMain.WebBrowser1.Left + frmMain.Left, frmMain.WebBrowser1.Top + frmMain.Top + 651
.Height = frmMain.WebBrowser1.Height
.Width = frmMain.WebBrowser1.Width - 200
.Enabled = False
.Visible = True
.Show
End With
End If
End If
Exit Sub
ep:
MsgBox "发生了错误:" & Chr(13) & Err.Description, vbCritical, "Error"
Me.WebBrowser1.Navigate Dir1.path
ream = File1.ListCount
Label10.Caption = Int(ream * Me.RegSvr.Interval / 1000) + 1
Label12.Caption = ream
Label4.Enabled = True
Label6.Caption = "正在执行操作,请稍候..."
Command11.Caption = "暂停(&S)"
regflag = False
unregflag = False
Me.Picture1.Visible = False
RegSvr.Enabled = False
UnRegSvr.Enabled = False
regsvrvrt = 0
Unregsvrrt = 0
Me.perform.Enabled = True
Command1.Enabled = True
Command3.Enabled = True
Command4.Enabled = True
Command5.Enabled = True
Command6.Enabled = True
Command7.Enabled = True
Command8.Enabled = True
Command9.Enabled = True
Command2.Enabled = True
Command10.Enabled = True
Dir1.Enabled = True
File1.Enabled = True
Drive1.Enabled = True
hlp.Enabled = True
End Sub
Private Sub regsvrhlp_Click()
On Error GoTo ep
Load Form8
Form8.Show 1
Exit Sub
ep:
MsgBox "发生系统错误:" & Chr(13) & Err.Description & Chr(13) & "您的系统版本可能不支持此功能或者您尚未安装这个功能.", vbCritical, "Error"
End Sub
Private Sub ren_Click()
On Error GoTo ep
Unload Form10
Dim newname As String
newname = InputBox("请输入要将选中的文件改名为的新名称.注意:改名包括扩展名" & Chr(13) & "比如,要将选中的文件改名为'Windows32.dll',应在文本框中输入Windows32.dll", "Rename")
If Trim(newname) = "" Then Exit Sub
If Right(File1.path, 1) = "\" Then
Name Label4.Caption As File1.path & Trim(newname)
Else
Name Label4.Caption As File1.path & "\" & Trim(newname)
End If
MsgBox "重命名文件成功!", vbInformation, "Info"
File1.Refresh
If 1 = 2 Then
With Form10
.Top = frmMain.WebBrowser1.Top + frmMain.Top + 651
.Left = frmMain.WebBrowser1.Left + frmMain.Left
.Move frmMain.WebBrowser1.Left + frmMain.Left, frmMain.WebBrowser1.Top + frmMain.Top + 651
.Height = frmMain.WebBrowser1.Height
.Width = frmMain.WebBrowser1.Width - 200
.Enabled = False
.Visible = True
.Show
End With
End If
Exit Sub
ep:
MsgBox "发生了错误:" & Chr(13) & Err.Description, vbCritical, "Error"
If 1 = 2 Then
With Form10
.Top = frmMain.WebBrowser1.Top + frmMain.Top + 651
.Left = frmMain.WebBrowser1.Left + frmMain.Left
.Move frmMain.WebBrowser1.Left + frmMain.Left, frmMain.WebBrowser1.Top + frmMain.Top + 651
.Height = frmMain.WebBrowser1.Height
.Width = frmMain.WebBrowser1.Width - 200
.Enabled = False
.Visible = True
.Show
End With
End If
End Sub
Private Sub unin_Click()
On Error GoTo ep
Unload Form10
If File1.ListIndex >= 0 Then
Shell ("regsvr32 /u " & Chr(34) & Label4.Caption & Chr(34))
Else
MsgBox "您尚未选择文件!", vbCritical, "Error"
End If
If 1 = 2 Then
With Form10
.Top = frmMain.WebBrowser1.Top + frmMain.Top + 651
.Left = frmMain.WebBrowser1.Left + frmMain.Left
.Move frmMain.WebBrowser1.Left + frmMain.Left, frmMain.WebBrowser1.Top + frmMain.Top + 651
.Height = frmMain.WebBrowser1.Height
.Width = frmMain.WebBrowser1.Width - 200
.Enabled = False
.Visible = True
.Show
End With
End If
Exit Sub
ep:
MsgBox "发生了错误:" & Chr(13) & Err.Description, vbCritical, "Error"
If 1 = 2 Then
With Form10
.Top = frmMain.WebBrowser1.Top + frmMain.Top + 651
.Left = frmMain.WebBrowser1.Left + frmMain.Left
.Move frmMain.WebBrowser1.Left + frmMain.Left, frmMain.WebBrowser1.Top + frmMain.Top + 651
.Height = frmMain.WebBrowser1.Height
.Width = frmMain.WebBrowser1.Width - 200
.Enabled = False
.Visible = True
.Show
End With
End If
End Sub
Private Sub unreg_Click()
On Error GoTo Error
Unload Form10
Dim filewrite
Dim ans As Integer
If File1.ListCount = 0 Then
MsgBox "没有可以反注册的控件,结束操作.", vbCritical, "Error"
If 1 = 2 Then
With Form10
.Top = frmMain.WebBrowser1.Top + frmMain.Top + 651
.Left = frmMain.WebBrowser1.Left + frmMain.Left
.Move frmMain.WebBrowser1.Left + frmMain.Left, frmMain.WebBrowser1.Top + frmMain.Top + 651
.Height = frmMain.WebBrowser1.Height
.Width = frmMain.WebBrowser1.Width - 200
.Enabled = False
.Visible = True
.Show
End With
End If
Exit Sub
End If
Exit Sub
Error:
If Err.Number = 32755 Then Exit Sub
MsgBox "发生错误:" & Chr(13) & Err.Description, vbCritical, "Error"
If 1 = 2 Then
With Form10
.Top = frmMain.WebBrowser1.Top + frmMain.Top + 651
.Left = frmMain.WebBrowser1.Left + frmMain.Left
.Move frmMain.WebBrowser1.Left + frmMain.Left, frmMain.WebBrowser1.Top + frmMain.Top + 651
.Height = frmMain.WebBrowser1.Height
.Width = frmMain.WebBrowser1.Width - 200
.Enabled = False
.Visible = True
.Show
End With
End If
End Sub
Private Sub UnRegSvr_Timer()
On Error GoTo ep
DoEvents
regflag = False
unregflag = True
Me.Picture1.Visible = True
Shell ("regsvr32.exe /s /u " & Chr(34) & Dir1.path & "\" & File1.List(unregsvrvrt) & Chr(34))
ream = ream - 1
Label10.Caption = Int(ream * Me.RegSvr.Interval / 1000) + 1
Label12.Caption = ream
unregsvrvrt = unregsvrvrt + 1
If unregsvrvrt = File1.ListCount Then
Label4.Enabled = True
Label6.Caption = "正在执行操作,请稍候..."
Me.WebBrowser1.Navigate Dir1.path
ream = File1.ListCount
Label10.Caption = Int(ream * Me.RegSvr.Interval / 1000) + 1
Label12.Caption = ream
Command11.Caption = "暂停(&S)"
regflag = False
unregflag = False
Me.Picture1.Visible = False
UnRegSvr.Enabled = False
RegSvr.Enabled = False
Command2.Enabled = True
Command1.Enabled = True
Command3.Enabled = True
Command4.Enabled = True
Command5.Enabled = True
Command6.Enabled = True
Command7.Enabled = True
Command8.Enabled = True
Command9.Enabled = True
Command10.Enabled = True
Me.perform.Enabled = True
Dir1.Enabled = True
File1.Enabled = True
Drive1.Enabled = True
file.Enabled = True
hlp.Enabled = True
MsgBox "反注册控件成功!", vbExclamation, "Info"
unregsvrvrt = 0
End If
Exit Sub
ep:
MsgBox "发生了错误:" & Chr(13) & Err.Description, vbCritical, "Error"
Me.WebBrowser1.Navigate Dir1.path
ream = File1.ListCount
Label10.Caption = Int(ream * Me.RegSvr.Interval / 1000) + 1
Label12.Caption = ream
Label4.Enabled = True
Label6.Caption = "正在执行操作,请稍候..."
Command11.Caption = "暂停(&S)"
regflag = False
unregflag = False
Me.Picture1.Visible = False
UnRegSvr.Enabled = False
RegSvr.Enabled = False
unregsvrvrt = 0
Command1.Enabled = True
Command3.Enabled = True
Command4.Enabled = True
Command5.Enabled = True
Command6.Enabled = True
Command7.Enabled = True
Command8.Enabled = True
Me.perform.Enabled = True
Command9.Enabled = True
Command2.Enabled = True
Command10.Enabled = True
Dir1.Enabled = True
File1.Enabled = True
Drive1.Enabled = True
hlp.Enabled = True
End Sub
Private Sub WebBrowser1_GotFocus()
On Error Resume Next
frmMain.SetFocus
End Sub
