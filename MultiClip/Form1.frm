VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Multi Clip Board 多次剪贴板 Version 3.0.0 - PC_DOS Workshop"
   ClientHeight    =   7815
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11370
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   9
      Charset         =   134
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7815
   ScaleWidth      =   11370
   Begin VB.CheckBox Check1 
      Caption         =   "保持在其它窗口前端(&K)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   9000
      TabIndex        =   31
      Top             =   7050
      Width           =   2310
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   330
      Left            =   1305
      TabIndex        =   29
      Top             =   7440
      Width           =   9045
   End
   Begin VB.CommandButton Command16 
      Caption         =   "帮助(&H)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5805
      TabIndex        =   27
      Top             =   7020
      Width           =   1620
   End
   Begin VB.CommandButton Command15 
      Caption         =   "最小化到托盘(&V)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4095
      TabIndex        =   26
      Top             =   7020
      Width           =   1635
   End
   Begin VB.CommandButton Command14 
      Cancel          =   -1  'True
      Caption         =   "退出(&T)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7530
      TabIndex        =   25
      Top             =   7020
      Width           =   1365
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   5880
      Top             =   3000
   End
   Begin VB.CommandButton Command11 
      Caption         =   "图像多次剪贴板(&P)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2130
      TabIndex        =   19
      Top             =   7020
      Width           =   1890
   End
   Begin VB.CommandButton Command10 
      Caption         =   "小窗口模式(&M)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   105
      TabIndex        =   18
      Top             =   7020
      Width           =   1950
   End
   Begin VB.Frame Frame2 
      Caption         =   "剪切板管理"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   0
      TabIndex        =   11
      Top             =   6000
      Width           =   11310
      Begin VB.Frame Frame3 
         Caption         =   "自动清空剪切板(每10秒)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3360
         TabIndex        =   13
         Top             =   120
         Width           =   7140
         Begin VB.CommandButton Command13 
            Caption         =   "关于此选项(&U)"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4860
            TabIndex        =   24
            Top             =   240
            Width           =   2055
         End
         Begin VB.OptionButton Option2 
            Caption         =   "禁用(&D)"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3060
            TabIndex        =   15
            Top             =   255
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.OptionButton Option1 
            Caption         =   "启用(&E)"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   360
            TabIndex        =   14
            Top             =   255
            Width           =   1455
         End
      End
      Begin VB.CommandButton Command7 
         Caption         =   "清空剪切板(&L)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   12
         Top             =   240
         Width           =   2775
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   10635
         Picture         =   "Form1.frx":0442
         Stretch         =   -1  'True
         Top             =   270
         Width           =   480
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "文字多次剪切/复制/粘贴"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5895
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11310
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FFFFFF&
         Height          =   1200
         Left            =   3615
         ScaleHeight     =   1140
         ScaleWidth      =   5970
         TabIndex        =   32
         Top             =   3330
         Width           =   6030
         Begin VB.Shape Shape1 
            BackColor       =   &H00800000&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00800000&
            Height          =   315
            Left            =   -15
            Top             =   0
            Width           =   6000
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "正在导入文件,请稍候..."
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   26.25
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   0
            TabIndex        =   33
            Top             =   420
            Width           =   6120
         End
      End
      Begin 工程1.cSysTray mni 
         Left            =   2910
         Top             =   3210
         _ExtentX        =   900
         _ExtentY        =   900
         InTray          =   0   'False
         TrayIcon        =   "Form1.frx":0884
         TrayTip         =   "Multi Cilpboard-双击还原窗口,右击显示菜单"
      End
      Begin VB.Frame Frame4 
         Caption         =   "监视剪切板"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   9195
         TabIndex        =   20
         Top             =   4305
         Width           =   1935
         Begin VB.CommandButton Command12 
            Caption         =   "关于此选项(&B)"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   23
            Top             =   960
            Width           =   1695
         End
         Begin VB.OptionButton Option4 
            Caption         =   "禁用(&S)"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   480
            TabIndex        =   22
            Top             =   600
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton Option3 
            Caption         =   "启用(&N)"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   480
            TabIndex        =   21
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.CommandButton Command9 
         Caption         =   "修改列表中选定的项目(&D)"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   1800
         Width           =   4695
      End
      Begin VB.CommandButton Command8 
         Caption         =   "导出列表中的全部项目为文本文件(&X)..."
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   3720
         Width           =   4695
      End
      Begin VB.Timer Timer2 
         Enabled         =   0   'False
         Interval        =   10000
         Left            =   7320
         Top             =   1680
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   6480
         Top             =   960
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         CancelError     =   -1  'True
      End
      Begin VB.CommandButton Command6 
         Caption         =   "从TXT文本文件导入(&I)..."
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   3240
         Width           =   4695
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Top             =   4440
         Width           =   8970
      End
      Begin VB.CommandButton Command5 
         Caption         =   "从剪切板中获取文字项目(&G)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   2760
         Width           =   4695
      End
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3840
         ItemData        =   "Form1.frx":0B9E
         Left            =   4920
         List            =   "Form1.frx":0BA0
         TabIndex        =   6
         Top             =   330
         Width           =   6225
      End
      Begin VB.CommandButton Command4 
         Caption         =   "移除选定的项目(&R)"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   1320
         Width           =   4695
      End
      Begin VB.CommandButton Command3 
         Caption         =   "向列表中添加项目(&A)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   2265
         Width           =   4695
      End
      Begin VB.CommandButton Command2 
         Caption         =   "清空列表中的项目(&C)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   4695
      End
      Begin VB.CommandButton Command1 
         Caption         =   "复制所选文字到剪贴板(&O)"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   4695
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "当前选定的项目:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   4200
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "可用的文字项目:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4920
         TabIndex        =   1
         Top             =   120
         Width           =   1815
      End
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "230"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   10440
      TabIndex        =   30
      Top             =   7440
      Width           =   855
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "窗口透明度:"
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
      Left            =   135
      TabIndex        =   28
      Top             =   7485
      Width           =   1155
   End
   Begin VB.Menu ctl 
      Caption         =   "control"
      Visible         =   0   'False
      Begin VB.Menu show 
         Caption         =   "显示主窗口(&S)"
      End
      Begin VB.Menu auto 
         Caption         =   "自动监视剪切板(&A)"
         Begin VB.Menu en 
            Caption         =   "启用(&E)"
         End
         Begin VB.Menu di 
            Caption         =   "禁用(&D)"
         End
      End
      Begin VB.Menu clr 
         Caption         =   "自动清空剪切板(&U)"
         Begin VB.Menu eee 
            Caption         =   "启用(&N)"
         End
         Begin VB.Menu ddd 
            Caption         =   "禁用(&I)"
         End
      End
      Begin VB.Menu CBC 
         Caption         =   "清空剪切板(&C)"
      End
      Begin VB.Menu pic 
         Caption         =   "图像多次剪切板(&P)"
      End
      Begin VB.Menu mmni 
         Caption         =   "退出托盘并打开小窗口(&X)"
      End
      Begin VB.Menu exit 
         Caption         =   "退出(&E)"
      End
   End
   Begin VB.Menu listctl 
      Caption         =   "Listctl"
      Visible         =   0   'False
      Begin VB.Menu mnucopy 
         Caption         =   "复制选定项(&C)"
      End
      Begin VB.Menu mnuremove 
         Caption         =   "移除选定项(&R)"
      End
      Begin VB.Menu l1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuclear 
         Caption         =   "清除列表(&L)"
      End
      Begin VB.Menu mnuexport 
         Caption         =   "导出列表(&E)"
      End
      Begin VB.Menu l2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuadd 
         Caption         =   "从剪切板添加(&A)"
      End
      Begin VB.Menu mnuinport 
         Caption         =   "从TXT文件导入(&I)"
      End
      Begin VB.Menu mnumadd 
         Caption         =   "手工输入项目添加(&M)"
      End
      Begin VB.Menu l3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuedit 
         Caption         =   "修改选定项(&D)"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim adddata As String
Dim loopback As Variant
Dim d As Integer
Dim autodata As String
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Const WS_EX_LAYERED = &H80000
Private Const GWL_EXSTYLE = (-20)
Private Const LWA_ALPHA = &H2
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Const LWA_COLORKEY = &H1
Private Sub CBC_Click()
Clipboard.Clear
End Sub
Private Sub Check1_Click()
On Error Resume Next
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Select Case Check1.Value
Case 1
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
Me.Height = 8190
Me.Width = 11460
Case 0
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
Me.Height = 8190
Me.Width = 11460
End Select
End Sub
Private Sub Command1_Click()
On Error GoTo ep
If List1.ListIndex > -1 Then
Clipboard.Clear
Clipboard.SetText List1.List(List1.ListIndex)
MsgBox "已经将选定的文字复制到剪贴板!", vbExclamation, "Copied"
Else
MsgBox "您尚未选择任何项目!", vbExclamation, "Error"
End If
Exit Sub
ep:
MsgBox "发生了错误" & Chr(13) & "错误号:" & Err.Number & Chr(13) & "错误描述符:" & Err.Description, vbCritical, "Error"
End Sub
Private Sub Command10_Click()
Me.Hide
frmMCM.show
End Sub
Private Sub Command11_Click()
Me.Hide
Form3.show
End Sub
Private Sub Command12_Click()
MsgBox "当选项为'已启用'时，程序每间隔10秒钟将从剪切板中获取文字数据并添加到列表中." & Chr(10) & Chr(10) & "注意：此选项只在完整窗口有效,并且迷你窗口不会继承这个选项的设定(迷你窗口强制启用该功能).", vbInformation, "Info"
End Sub
Private Sub Command13_Click()
MsgBox "当选项为'已启用'时，程序每间隔10秒钟清空剪切板的内容,以防别有用心的人获得您不慎保留在剪切板中的数据." & Chr(10) & Chr(10) & "注意：此选项只在完整窗口/迷你窗口中有效,并且迷你窗口会继承这个选项的设定." & Chr(10) & Chr(10) & "例如:将本选项启用并进入迷你窗口,那么迷你窗口仍然具有自动清空剪切板的功能,而且您必须返回到默认窗口才可以修改这项设置;但是假如您在未启用本选项的情况下进入迷你窗口,那么自动清空将不可用,除非您返回默认窗口修改.", vbInformation, "Info"
End Sub
Private Sub Command14_Click()
Dim a As Integer
a = MsgBox("确认退出吗？所以数据都会丢失!", vbExclamation + vbYesNo, "Alert")
If a = vbYes Then
End
Else
Cancel = 666
End If
End Sub
Private Sub Command15_Click()
Me.Hide
mni.InTray = True
End Sub
Private Sub Command16_Click()
MsgBox "这个程序可以帮助您获得更加高级 Windows剪切板,包括:" & Chr(13) & "多次剪切的操作" & Chr(13) & "从文件导入/导出数据" & Chr(13) & "图像多次剪切操作" & Chr(13) & "自动清除剪切板" & Chr(13) & "剪切板内容的导入与导出" & Chr(13) & "自动监视剪切板" & Chr(13) & Chr(13) & "如果您觉得大窗口太大,可以选择[小窗口模式]打开小窗口" & Chr(13) & Chr(13) & "注意:完整窗口的文字列表框和迷你窗口的文字列表框是不会自动同步的" & Chr(13) & Chr(13) & "提示:在可复制文本列表中选择项目直接按Ctrl+C可直接复制选定的文本" & Chr(13) & Chr(13) & Chr(13) & "PC-DOS Workshop 出品,版权没有,翻版不究", vbInformation, "Help"
End Sub
Private Sub Command2_Click()
On Error GoTo ep
d = MsgBox("确定要清除全部项目吗?", vbQuestion + vbYesNo, "Clear?")
If d = vbYes Then
List1.Clear
MsgBox "清空列表已经完成!", vbExclamation, "Cleared"
Command1.Enabled = False
Command4.Enabled = False
Command9.Enabled = False
Text1.Text = ""
End If
Exit Sub
ep:
MsgBox "发生了错误" & Chr(13) & "错误号:" & Err.Number & Chr(13) & "错误描述符:" & Err.Description, vbCritical, "Error"
End Sub
Private Sub Command3_Click()
On Error GoTo ep
adddata = InputBox("请输入要在剪切列表中添加的项目!", "Add Data")
If adddata <> "" Then
List1.AddItem adddata
MsgBox "数据已经添加!" & Chr(13) & "添加了:" & adddata, vbExclamation, "Added"
Else
MsgBox "对不起,不允许添加空值!", vbExclamation, "Error"
End If
Exit Sub
ep:
MsgBox "发生了错误" & Chr(13) & "错误号:" & Err.Number & Chr(13) & "错误描述符:" & Err.Description, vbCritical, "Error"
End Sub
Private Sub Command4_Click()
On Error GoTo ep
If List1.ListIndex > -1 Then
d = MsgBox("确定移除所选定的项目吗?", vbQuestion + vbYesNo, "Remove?")
If d = vbYes Then
List1.RemoveItem (List1.ListIndex)
MsgBox "清除已经完成!", vbExclamation, "Cleared"
Command1.Enabled = False
Command4.Enabled = False
Command9.Enabled = False
Text1.Text = ""
Else
MsgBox "您尚未选择任何项目!", vbExclamation, "Error"
End If
End If
Exit Sub
ep:
MsgBox "发生了错误" & Chr(13) & "错误号:" & Err.Number & Chr(13) & "错误描述符:" & Err.Description, vbCritical, "Error"
End Sub
Private Sub Command5_Click()
On Error GoTo ep
adddata = Clipboard.GetText
If adddata <> "" Then
List1.AddItem adddata
MsgBox "已经从剪贴板中向列表添加了项目!" & Chr(13) & "添加了:" & adddata, vbExclamation, "Added"
Else
MsgBox "对不起,当前剪贴板为空!", vbExclamation, "Error"
End If
Exit Sub
ep:
MsgBox "发生了错误" & Chr(13) & "错误号:" & Err.Number & Chr(13) & "错误描述符:" & Err.Description, vbCritical, "Error"
End Sub
Private Sub Command6_Click()
On Error GoTo ep
With CommonDialog1
.DialogTitle = "请选择要导入的TXT文本文档"
.Filter = "TXT文本(*.TXT)|*.TXT"
.ShowOpen
End With
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
Me.Height = 8190
Me.Width = 11460
If CommonDialog1.FileName <> "" Then
Open CommonDialog1.FileName For Input As 1
Me.Check1.Enabled = False
Me.Picture1.Enabled = False
Me.Command1.Enabled = False
Me.Command10.Enabled = False
Me.Command11.Enabled = False
Me.Command12.Enabled = False
Me.Command13.Enabled = False
Me.Command14.Enabled = False
Me.Command15.Enabled = False
Me.Command16.Enabled = False
Me.Command2.Enabled = False
Me.Command3.Enabled = False
Me.Command4.Enabled = False
Me.Command5.Enabled = False
Me.Command6.Enabled = False
Me.Command7.Enabled = False
Me.Command8.Enabled = False
Me.Command9.Enabled = False
Me.Option1.Enabled = False
Me.Option2.Enabled = False
Me.Option3.Enabled = False
Me.Option4.Enabled = False
Me.List1.Enabled = False
Me.Text1.Enabled = False
Me.HScroll1.Enabled = False
Me.Label1.Enabled = False
Me.Label12.Enabled = False
Me.Label3.Enabled = False
Me.Label4.Enabled = False
With Me.Picture1
.Left = Me.Width / 2 - .Width / 2
.Top = Me.Height / 2 - .Height / 2
.Visible = True
End With
With Label4
.Visible = True
.Top = Me.Shape1.Height
.Left = 0
.Height = Picture1.Height - Shape1.Height
.Width = Picture1.Width
.Caption = "正在导入文件,请稍候..."
End With
With Me.Shape1
.Visible = True
.BackColor = RGB(0, 0, 255)
.BorderColor = RGB(0, 0, 255)
End With
Sleep 100
Form4.show 1
Select Case Me.Tag
Case "Line"
Do While Not EOF(1)
Input #1, adddata
List1.AddItem adddata
Me.Check1.Enabled = False
Me.Picture1.Enabled = False
Me.Command1.Enabled = False
Me.Command10.Enabled = False
Me.Command11.Enabled = False
Me.Command12.Enabled = False
Me.Command13.Enabled = False
Me.Command14.Enabled = False
Me.Command15.Enabled = False
Me.Command16.Enabled = False
Me.Command2.Enabled = False
Me.Command3.Enabled = False
Me.Command4.Enabled = False
Me.Command5.Enabled = False
Me.Command6.Enabled = False
Me.Command7.Enabled = False
Me.Command8.Enabled = False
Me.Command9.Enabled = False
Me.Option1.Enabled = False
Me.Option2.Enabled = False
Me.Option3.Enabled = False
Me.Option4.Enabled = False
Me.List1.Enabled = False
Me.Text1.Enabled = False
Me.HScroll1.Enabled = False
Me.Label1.Enabled = False
Me.Label12.Enabled = False
Me.Label3.Enabled = False
Me.Label4.Enabled = False
With Me.Picture1
.Left = Me.Width / 2 - .Width / 2
.Top = Me.Height / 2 - .Height / 2
.Visible = True
End With
With Label4
.Visible = True
.Top = Me.Shape1.Height
.Left = 0
.Height = Picture1.Height - Shape1.Height
.Width = Picture1.Width
.Caption = "正在导入文件,请稍候..."
End With
With Me.Shape1
.Visible = True
.BackColor = RGB(0, 0, 255)
.BorderColor = RGB(0, 0, 255)
End With
DoEvents
Loop
Case "Whole"
Close
Open CommonDialog1.FileName For Binary As #2
adddata = StrConv(InputB(LOF(2), #2), vbUnicode)
List1.AddItem adddata
Me.Check1.Enabled = False
Me.Picture1.Enabled = False
Me.Command1.Enabled = False
Me.Command10.Enabled = False
Me.Command11.Enabled = False
Me.Command12.Enabled = False
Me.Command13.Enabled = False
Me.Command14.Enabled = False
Me.Command15.Enabled = False
Me.Command16.Enabled = False
Me.Command2.Enabled = False
Me.Command3.Enabled = False
Me.Command4.Enabled = False
Me.Command5.Enabled = False
Me.Command6.Enabled = False
Me.Command7.Enabled = False
Me.Command8.Enabled = False
Me.Command9.Enabled = False
Me.Option1.Enabled = False
Me.Option2.Enabled = False
Me.Option3.Enabled = False
Me.Option4.Enabled = False
Me.List1.Enabled = False
Me.Text1.Enabled = False
Me.HScroll1.Enabled = False
Me.Label1.Enabled = False
Me.Label12.Enabled = False
Me.Label3.Enabled = False
Me.Label4.Enabled = False
With Me.Picture1
.Left = Me.Width / 2 - .Width / 2
.Top = Me.Height / 2 - .Height / 2
.Visible = True
End With
With Label4
.Visible = True
.Top = Me.Shape1.Height
.Left = 0
.Height = Picture1.Height - Shape1.Height
.Width = Picture1.Width
.Caption = "正在导入文件,请稍候..."
End With
With Me.Shape1
.Visible = True
.BackColor = RGB(0, 0, 255)
.BorderColor = RGB(0, 0, 255)
End With
DoEvents
End Select
Close
Me.Picture1.Enabled = True
Me.Command1.Enabled = True
Me.Command10.Enabled = True
Me.Command11.Enabled = True
Me.Command12.Enabled = True
Me.Command13.Enabled = True
Me.Command14.Enabled = True
Me.Command15.Enabled = True
Me.Command16.Enabled = True
Me.Command2.Enabled = True
Me.Command3.Enabled = True
Me.Command4.Enabled = True
Me.Command5.Enabled = True
Me.Command6.Enabled = True
Me.Command7.Enabled = True
Me.Command8.Enabled = True
Me.Command9.Enabled = True
Me.Option1.Enabled = True
Me.Option2.Enabled = True
Me.Option3.Enabled = True
Me.Option4.Enabled = True
Me.Check1.Enabled = True
Me.List1.Enabled = True
Me.Text1.Enabled = True
Me.HScroll1.Enabled = True
Me.Label1.Enabled = True
Me.Label12.Enabled = True
Me.Label3.Enabled = True
Me.Label4.Enabled = True
With Me.Picture1
.Left = Me.Width / 2 - .Width / 2
.Top = Me.Height / 2 - .Height / 2
.Visible = False
End With
With Label4
.Visible = False
.Top = Me.Shape1.Height
.Left = 0
.Height = Picture1.Height - Shape1.Height
.Width = Picture1.Width
.Caption = "正在导入文件,请稍候..."
End With
With Me.Shape1
.Visible = False
.BackColor = RGB(0, 0, 255)
.BorderColor = RGB(0, 0, 255)
End With
Select Case Check1.Value
Case 1
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
Me.Height = 8190
Me.Width = 11460
Case 0
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
Me.Height = 8190
Me.Width = 11460
End Select
Refresh
Else
MsgBox "尚未选择文件!", vbCritical, "Error"
Select Case Check1.Value
Case 1
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
Me.Height = 8190
Me.Width = 11460
Case 0
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
Me.Height = 8190
Me.Width = 11460
End Select
Refresh
Exit Sub
End If
Exit Sub
ep:
If Err.Description <> "选定“取消”。" Then
MsgBox "发生了错误,可能是因为操作不当导致的" & Chr(13) & "错误:" & Err.Description, vbCritical, "Error"
Me.Picture1.Enabled = True
Me.Command1.Enabled = True
Me.Command10.Enabled = True
Me.Command11.Enabled = True
Me.Command12.Enabled = True
Me.Command13.Enabled = True
Me.Command14.Enabled = True
Me.Command15.Enabled = True
Me.Command16.Enabled = True
Me.Command2.Enabled = True
Me.Command3.Enabled = True
Me.Command4.Enabled = True
Me.Command5.Enabled = True
Me.Command6.Enabled = True
Me.Command7.Enabled = True
Me.Command8.Enabled = True
Me.Command9.Enabled = True
Me.Option1.Enabled = True
Me.Option2.Enabled = True
Me.Option3.Enabled = True
Me.Option4.Enabled = True
Me.Check1.Enabled = True
Me.List1.Enabled = True
Me.Text1.Enabled = True
Me.HScroll1.Enabled = True
Me.Label1.Enabled = True
Me.Label12.Enabled = True
Me.Label3.Enabled = True
Me.Label4.Enabled = True
With Me.Picture1
.Left = Me.Width / 2 - .Width / 2
.Top = Me.Height / 2 - .Height / 2
.Visible = False
End With
With Label4
.Visible = False
.Top = Me.Shape1.Height
.Left = 0
.Height = Picture1.Height - Shape1.Height
.Width = Picture1.Width
.Caption = "正在导入文件,请稍候..."
End With
With Me.Shape1
.Visible = False
.BackColor = RGB(0, 0, 255)
.BorderColor = RGB(0, 0, 255)
End With
Else
Me.Picture1.Enabled = True
Me.Command1.Enabled = True
Me.Command10.Enabled = True
Me.Command11.Enabled = True
Me.Command12.Enabled = True
Me.Command13.Enabled = True
Me.Command14.Enabled = True
Me.Command15.Enabled = True
Me.Command16.Enabled = True
Me.Command2.Enabled = True
Me.Command3.Enabled = True
Me.Command4.Enabled = True
Me.Command5.Enabled = True
Me.Command6.Enabled = True
Me.Command7.Enabled = True
Me.Command8.Enabled = True
Me.Command9.Enabled = True
Me.Option1.Enabled = True
Me.Option2.Enabled = True
Me.Option3.Enabled = True
Me.Option4.Enabled = True
Me.Check1.Enabled = True
Me.List1.Enabled = True
Me.Text1.Enabled = True
Me.HScroll1.Enabled = True
Me.Label1.Enabled = True
Me.Label12.Enabled = True
Me.Label3.Enabled = True
Me.Label4.Enabled = True
With Me.Picture1
.Left = Me.Width / 2 - .Width / 2
.Top = Me.Height / 2 - .Height / 2
.Visible = False
End With
With Label4
.Visible = False
.Top = Me.Shape1.Height
.Left = 0
.Height = Picture1.Height - Shape1.Height
.Width = Picture1.Width
.Caption = "正在导入文件,请稍候..."
End With
With Me.Shape1
.Visible = False
.BackColor = RGB(0, 0, 255)
.BorderColor = RGB(0, 0, 255)
End With
Exit Sub
End If
End Sub
Private Sub Command7_Click()
On Error GoTo ep
Clipboard.Clear
MsgBox "清空已经完成!", vbExclamation, "Cleared"
Exit Sub
ep:
MsgBox "发生了错误" & Chr(13) & "错误号:" & Err.Number & Chr(13) & "错误描述符:" & Err.Description, vbCritical, "Error"
End Sub
Private Sub Command8_Click()
On Error GoTo ep
Dim ans As Integer
With CommonDialog1
.DialogTitle = "请指定导出文件的选项"
.Filter = "TXT文本(*.TXT)|*.TXT"
.ShowSave
End With
If CommonDialog1.FileName <> "" Then
If Dir(Me.CommonDialog1.FileName) = "" Then
Open CommonDialog1.FileName For Output As #1
For loopback = 0 To List1.ListCount
Print #1, List1.List(loopback)
Next loopback
MsgBox "导出已经完成!", vbExclamation, "Outputted"
Close #1
Else
ans = MsgBox("目标文件已经存在,是否替换?", vbExclamation + vbYesNo, "Ask")
If ans = vbYes Then
Open CommonDialog1.FileName For Output As #1
For loopback = 0 To List1.ListCount
Print #1, List1.List(loopback)
Next loopback
MsgBox "导出已经完成!", vbExclamation, "Outputted"
Close #1
Else
Exit Sub
End If
End If
Else
MsgBox "对不起,文件名不允许为空!", vbCritical, "Error"
End If
Exit Sub
ep:
If Err.Description <> "选定“取消”。" Then
MsgBox "发生了错误,可能是因为操作不当导致的" & Chr(13) & "错误:" & Err.Description, vbCritical, "Error"
Else
Exit Sub
End If
End Sub
Private Sub Command9_Click()
On Error GoTo ep
If List1.ListIndex > -1 Then
adddata = InputBox("请输入要修改的项目的值!", "Edit Data")
If adddata <> "" Then
List1.List(List1.ListIndex) = adddata
MsgBox "修改已经完成!", vbExclamation, "Changed"
Text1.Text = List1.List(List1.ListIndex)
Else
MsgBox "对不起,不允许输入空值!", vbExclamation, "Error"
Text1.Text = List1.List(List1.ListIndex)
End If
Else
MsgBox "您尚未选择项目!", vbExclamation, "Error"
End If
Exit Sub
ep:
MsgBox "发生了错误" & Chr(13) & "错误号:" & Err.Number & Chr(13) & "错误描述符:" & Err.Description, vbCritical, "Error"
End Sub
Private Sub ddd_Click()
Timer2.Enabled = False
ddd.Enabled = False
eee.Enabled = True
Option2.Value = True
End Sub
Private Sub di_Click()
Timer1.Enabled = False
di.Enabled = False
en.Enabled = True
Option4.Value = True
End Sub
Private Sub eee_Click()
Timer2.Enabled = True
ddd.Enabled = True
eee.Enabled = False
Option1.Value = True
End Sub
Private Sub en_Click()
Timer1.Enabled = True
en.Enabled = False
di.Enabled = True
Option3.Value = True
End Sub
Private Sub exit_Click()
Unload Me
Unload Form1
Unload frmMCM
Unload Form3
Unload Form5
End
End Sub
Private Sub Form_Activate()
On Error Resume Next
List1.SetFocus
End Sub
Private Sub Form_Load()
On Error Resume Next
If App.PrevInstance = False Then
Dim rtn     As Long
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 230, LWA_ALPHA
Me.HScroll1.Value = 230
Label12.Caption = Me.HScroll1.Value
Me.Picture1.Enabled = True
Me.Command1.Enabled = True
Me.Command10.Enabled = True
Me.Command11.Enabled = True
Me.Command12.Enabled = True
Me.Command13.Enabled = True
Me.Command14.Enabled = True
Me.Command15.Enabled = True
Me.Command16.Enabled = True
Me.Command2.Enabled = True
Me.Command3.Enabled = True
Me.Command4.Enabled = True
Me.Command5.Enabled = True
Me.Command6.Enabled = True
Me.Command7.Enabled = True
Me.Command8.Enabled = True
Me.Command9.Enabled = True
Me.Option1.Enabled = True
Me.Option2.Enabled = True
Me.Option3.Enabled = True
Me.Option4.Enabled = True
Me.List1.Enabled = True
Me.Text1.Enabled = True
Me.HScroll1.Enabled = True
Me.Label1.Enabled = True
Me.Label12.Enabled = True
Me.Label3.Enabled = True
Me.Label4.Enabled = True
Me.Check1.Enabled = True
With Me.Picture1
.Left = Me.Width / 2 - .Width / 2
.Top = Me.Height / 2 - .Height / 2
.Visible = False
End With
With Label4
.Visible = False
.Top = Me.Shape1.Height
.Left = 0
.Height = Picture1.Height - Shape1.Height
.Width = Picture1.Width
.Caption = "正在导入文件,请稍候..."
End With
With Me.Shape1
.Visible = False
.BackColor = RGB(0, 0, 255)
.BorderColor = RGB(0, 0, 255)
End With
With Me.HScroll1
.Max = 255
.Min = 155
.LargeChange = 5
.SmallChange = 1
.Enabled = True
.Value = 230
End With
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
Me.Height = 8190
Me.Width = 11460
Me.Check1.Value = 1
ddd.Enabled = False
di.Enabled = False
Else
MsgBox "您已经打开本程序的一个实例,但是,因为剪切板程序的特殊性,系统不允许您同时打开程序的2个及以上的实例,为了系统稳定性,请点击[确定]立即退出.", vbCritical, "Error"
End
End If
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
If List1.ListCount = 0 Then
End
End If
Dim a As Integer
a = MsgBox("确认退出吗？所以数据都会丢失!", vbExclamation + vbYesNo, "Alert")
If a = vbYes Then
Unload Me
Else
Cancel = 666
End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
End
End Sub
Private Sub HScroll1_Change()
On Error Resume Next
Dim rtn     As Long
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, Me.HScroll1.Value, LWA_ALPHA
Label12.Caption = Me.HScroll1.Value
End Sub
Private Sub Label12_Click()
On Error Resume Next
Dim alp As Integer
Dim oldalp As Integer
Dim rtn As Long
oldalp = Label12.Caption
alp = Val(InputBox$("请输入透明度" & vbCrLf & "范围:155-255", "Alpha", 230))
If Val(alp) = 0 Then
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, oldalp, LWA_ALPHA
Me.HScroll1.Value = oldalp
Label12.Caption = Me.HScroll1.Value
Exit Sub
End If
If 155 <= alp And alp <= 255 Then
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, alp, LWA_ALPHA
Me.HScroll1.Value = alp
Label12.Caption = Me.HScroll1.Value
Else
MsgBox "无效透明度数值", vbCritical, "Error"
End If
End Sub
Private Sub List1_Click()
On Error GoTo ep
If List1.ListIndex > -1 Then
Command1.Enabled = True
Command4.Enabled = True
Command9.Enabled = True
Text1.Text = List1.List(List1.ListIndex)
Else
Command1.Enabled = False
Command4.Enabled = False
Command9.Enabled = False
End If
Exit Sub
ep:
MsgBox "发生了错误" & Chr(13) & "错误号:" & Err.Number & Chr(13) & "错误描述符:" & Err.Description, vbCritical, "Error"
End Sub
Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ep
If List1.ListIndex >= 0 And KeyCode = vbKeyC And Shift = 2 Then Clipboard.SetText List1.List(List1.ListIndex)
If KeyCode = vbKeyV And Shift = 2 Then
adddata = Clipboard.GetText
If adddata <> "" Then
List1.AddItem adddata
MsgBox "已经从剪贴板中向列表添加了项目!" & Chr(13) & "添加了:" & adddata, vbExclamation, "Added"
Else
MsgBox "对不起,当前剪贴板为空!", vbExclamation, "Error"
End If
End If
Exit Sub
ep:
MsgBox "发生了错误" & Chr(13) & "错误号:" & Err.Number & Chr(13) & "错误描述符:" & Err.Description, vbCritical, "Error"
End Sub
Private Sub List1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If List1.ListIndex >= 0 And Button = 2 Then
PopupMenu Me.listctl
Else
Exit Sub
End If
End Sub
Private Sub mmni_Click()
mni.InTray = False
Form1.Hide
frmMCM.show
End Sub
Private Sub mni_MouseDblClick(Button As Integer, Id As Long)
Form1.show
mni.InTray = False
End Sub
Private Sub mni_MouseDown(Button As Integer, Id As Long)
If Button = 2 Then PopupMenu ctl
End Sub
Private Sub mnuadd_Click()
On Error GoTo ep
adddata = Clipboard.GetText
If adddata <> "" Then
List1.AddItem adddata
MsgBox "已经从剪贴板中向列表添加了项目!" & Chr(13) & "添加了:" & adddata, vbExclamation, "Added"
Else
MsgBox "对不起,当前剪贴板为空!", vbExclamation, "Error"
End If
Exit Sub
ep:
MsgBox "发生了错误" & Chr(13) & "错误号:" & Err.Number & Chr(13) & "错误描述符:" & Err.Description, vbCritical, "Error"
End Sub
Private Sub mnuclear_Click()
On Error GoTo ep
d = MsgBox("确定要清除全部项目吗?", vbQuestion + vbYesNo, "Clear?")
If d = vbYes Then
List1.Clear
MsgBox "清空列表已经完成!", vbExclamation, "Cleared"
Command1.Enabled = False
Command4.Enabled = False
Command9.Enabled = False
Text1.Text = ""
End If
Exit Sub
ep:
MsgBox "发生了错误" & Chr(13) & "错误号:" & Err.Number & Chr(13) & "错误描述符:" & Err.Description, vbCritical, "Error"
End Sub
Private Sub mnucopy_Click()
On Error GoTo ep
If List1.ListIndex > -1 Then
Clipboard.Clear
Clipboard.SetText List1.List(List1.ListIndex)
MsgBox "已经将选定的文字复制到剪贴板!", vbExclamation, "Copied"
Else
MsgBox "您尚未选择任何项目!", vbExclamation, "Error"
End If
Exit Sub
ep:
MsgBox "发生了错误" & Chr(13) & "错误号:" & Err.Number & Chr(13) & "错误描述符:" & Err.Description, vbCritical, "Error"
End Sub
Private Sub mnuedit_Click()
On Error GoTo ep
If List1.ListIndex > -1 Then
adddata = InputBox("请输入要修改的项目的值!", "Edit Data")
If adddata <> "" Then
List1.List(List1.ListIndex) = adddata
MsgBox "修改已经完成!", vbExclamation, "Changed"
Text1.Text = List1.List(List1.ListIndex)
Else
MsgBox "对不起,不允许输入空值!", vbExclamation, "Error"
Text1.Text = List1.List(List1.ListIndex)
End If
Else
MsgBox "您尚未选择项目!", vbExclamation, "Error"
End If
Exit Sub
ep:
MsgBox "发生了错误" & Chr(13) & "错误号:" & Err.Number & Chr(13) & "错误描述符:" & Err.Description, vbCritical, "Error"
End Sub
Private Sub mnuexport_Click()
On Error GoTo ep
Dim ans As Integer
With CommonDialog1
.DialogTitle = "请指定导出文件的选项"
.Filter = "TXT文本(*.TXT)|*.TXT"
.ShowSave
End With
If CommonDialog1.FileName <> "" Then
If Dir(Me.CommonDialog1.FileName) = "" Then
Open CommonDialog1.FileName For Output As #1
For loopback = 0 To List1.ListCount
Print #1, List1.List(loopback)
Next loopback
MsgBox "导出已经完成!", vbExclamation, "Outputted"
Close #1
Else
ans = MsgBox("目标文件已经存在,是否替换?", vbExclamation + vbYesNo, "Ask")
If ans = vbYes Then
Open CommonDialog1.FileName For Output As #1
For loopback = 0 To List1.ListCount
Print #1, List1.List(loopback)
Next loopback
MsgBox "导出已经完成!", vbExclamation, "Outputted"
Close #1
Else
Exit Sub
End If
End If
Else
MsgBox "对不起,文件名不允许为空!", vbCritical, "Error"
End If
Exit Sub
ep:
If Err.Description <> "选定“取消”。" Then
MsgBox "发生了错误,可能是因为操作不当导致的" & Chr(13) & "错误:" & Err.Description, vbCritical, "Error"
Else
Exit Sub
End If
End Sub
Private Sub mnuinport_Click()
On Error GoTo ep
With CommonDialog1
.DialogTitle = "请选择要导入的TXT文本文档"
.Filter = "TXT文本(*.TXT)|*.TXT"
.ShowOpen
End With
If CommonDialog1.FileName <> "" Then
Open CommonDialog1.FileName For Input As 1
Me.Check1.Enabled = False
Me.Picture1.Enabled = False
Me.Command1.Enabled = False
Me.Command10.Enabled = False
Me.Command11.Enabled = False
Me.Command12.Enabled = False
Me.Command13.Enabled = False
Me.Command14.Enabled = False
Me.Command15.Enabled = False
Me.Command16.Enabled = False
Me.Command2.Enabled = False
Me.Command3.Enabled = False
Me.Command4.Enabled = False
Me.Command5.Enabled = False
Me.Command6.Enabled = False
Me.Command7.Enabled = False
Me.Command8.Enabled = False
Me.Command9.Enabled = False
Me.Option1.Enabled = False
Me.Option2.Enabled = False
Me.Option3.Enabled = False
Me.Option4.Enabled = False
Me.List1.Enabled = False
Me.Text1.Enabled = False
Me.HScroll1.Enabled = False
Me.Label1.Enabled = False
Me.Label12.Enabled = False
Me.Label3.Enabled = False
Me.Label4.Enabled = False
With Me.Picture1
.Left = Me.Width / 2 - .Width / 2
.Top = Me.Height / 2 - .Height / 2
.Visible = True
End With
With Label4
.Visible = True
.Top = Me.Shape1.Height
.Left = 0
.Height = Picture1.Height - Shape1.Height
.Width = Picture1.Width
.Caption = "正在导入文件,请稍候..."
End With
With Me.Shape1
.Visible = True
.BackColor = RGB(0, 0, 255)
.BorderColor = RGB(0, 0, 255)
End With
Sleep 100
Do While Not EOF(1)
Input #1, adddata
List1.AddItem adddata
Me.Check1.Enabled = False
Me.Picture1.Enabled = False
Me.Command1.Enabled = False
Me.Command10.Enabled = False
Me.Command11.Enabled = False
Me.Command12.Enabled = False
Me.Command13.Enabled = False
Me.Command14.Enabled = False
Me.Command15.Enabled = False
Me.Command16.Enabled = False
Me.Command2.Enabled = False
Me.Command3.Enabled = False
Me.Command4.Enabled = False
Me.Command5.Enabled = False
Me.Command6.Enabled = False
Me.Command7.Enabled = False
Me.Command8.Enabled = False
Me.Command9.Enabled = False
Me.Option1.Enabled = False
Me.Option2.Enabled = False
Me.Option3.Enabled = False
Me.Option4.Enabled = False
Me.List1.Enabled = False
Me.Text1.Enabled = False
Me.HScroll1.Enabled = False
Me.Label1.Enabled = False
Me.Label12.Enabled = False
Me.Label3.Enabled = False
Me.Label4.Enabled = False
With Me.Picture1
.Left = Me.Width / 2 - .Width / 2
.Top = Me.Height / 2 - .Height / 2
.Visible = True
End With
With Label4
.Visible = True
.Top = Me.Shape1.Height
.Left = 0
.Height = Picture1.Height - Shape1.Height
.Width = Picture1.Width
.Caption = "正在导入文件,请稍候..."
End With
With Me.Shape1
.Visible = True
.BackColor = RGB(0, 0, 255)
.BorderColor = RGB(0, 0, 255)
End With
DoEvents
Loop
Close 1
Me.Picture1.Enabled = True
Me.Command1.Enabled = True
Me.Command10.Enabled = True
Me.Command11.Enabled = True
Me.Command12.Enabled = True
Me.Command13.Enabled = True
Me.Command14.Enabled = True
Me.Command15.Enabled = True
Me.Command16.Enabled = True
Me.Command2.Enabled = True
Me.Command3.Enabled = True
Me.Command4.Enabled = True
Me.Command5.Enabled = True
Me.Command6.Enabled = True
Me.Command7.Enabled = True
Me.Command8.Enabled = True
Me.Command9.Enabled = True
Me.Option1.Enabled = True
Me.Option2.Enabled = True
Me.Option3.Enabled = True
Me.Option4.Enabled = True
Me.Check1.Enabled = True
Me.List1.Enabled = True
Me.Text1.Enabled = True
Me.HScroll1.Enabled = True
Me.Label1.Enabled = True
Me.Label12.Enabled = True
Me.Label3.Enabled = True
Me.Label4.Enabled = True
With Me.Picture1
.Left = Me.Width / 2 - .Width / 2
.Top = Me.Height / 2 - .Height / 2
.Visible = False
End With
With Label4
.Visible = False
.Top = Me.Shape1.Height
.Left = 0
.Height = Picture1.Height - Shape1.Height
.Width = Picture1.Width
.Caption = "正在导入文件,请稍候..."
End With
With Me.Shape1
.Visible = False
.BackColor = RGB(0, 0, 255)
.BorderColor = RGB(0, 0, 255)
End With
Else
MsgBox "尚未选择文件!", vbCritical, "Error"
End If
Exit Sub
ep:
If Err.Description <> "选定“取消”。" Then
MsgBox "发生了错误,可能是因为操作不当导致的" & Chr(13) & "错误:" & Err.Description, vbCritical, "Error"
Me.Picture1.Enabled = True
Me.Command1.Enabled = True
Me.Command10.Enabled = True
Me.Command11.Enabled = True
Me.Command12.Enabled = True
Me.Command13.Enabled = True
Me.Command14.Enabled = True
Me.Command15.Enabled = True
Me.Command16.Enabled = True
Me.Command2.Enabled = True
Me.Command3.Enabled = True
Me.Command4.Enabled = True
Me.Command5.Enabled = True
Me.Command6.Enabled = True
Me.Command7.Enabled = True
Me.Command8.Enabled = True
Me.Command9.Enabled = True
Me.Option1.Enabled = True
Me.Option2.Enabled = True
Me.Option3.Enabled = True
Me.Option4.Enabled = True
Me.Check1.Enabled = True
Me.List1.Enabled = True
Me.Text1.Enabled = True
Me.HScroll1.Enabled = True
Me.Label1.Enabled = True
Me.Label12.Enabled = True
Me.Label3.Enabled = True
Me.Label4.Enabled = True
With Me.Picture1
.Left = Me.Width / 2 - .Width / 2
.Top = Me.Height / 2 - .Height / 2
.Visible = False
End With
With Label4
.Visible = False
.Top = Me.Shape1.Height
.Left = 0
.Height = Picture1.Height - Shape1.Height
.Width = Picture1.Width
.Caption = "正在导入文件,请稍候..."
End With
With Me.Shape1
.Visible = False
.BackColor = RGB(0, 0, 255)
.BorderColor = RGB(0, 0, 255)
End With
Else
Me.Picture1.Enabled = True
Me.Command1.Enabled = True
Me.Command10.Enabled = True
Me.Command11.Enabled = True
Me.Command12.Enabled = True
Me.Command13.Enabled = True
Me.Command14.Enabled = True
Me.Command15.Enabled = True
Me.Command16.Enabled = True
Me.Command2.Enabled = True
Me.Command3.Enabled = True
Me.Command4.Enabled = True
Me.Command5.Enabled = True
Me.Command6.Enabled = True
Me.Command7.Enabled = True
Me.Command8.Enabled = True
Me.Command9.Enabled = True
Me.Option1.Enabled = True
Me.Option2.Enabled = True
Me.Option3.Enabled = True
Me.Option4.Enabled = True
Me.Check1.Enabled = True
Me.List1.Enabled = True
Me.Text1.Enabled = True
Me.HScroll1.Enabled = True
Me.Label1.Enabled = True
Me.Label12.Enabled = True
Me.Label3.Enabled = True
Me.Label4.Enabled = True
With Me.Picture1
.Left = Me.Width / 2 - .Width / 2
.Top = Me.Height / 2 - .Height / 2
.Visible = False
End With
With Label4
.Visible = False
.Top = Me.Shape1.Height
.Left = 0
.Height = Picture1.Height - Shape1.Height
.Width = Picture1.Width
.Caption = "正在导入文件,请稍候..."
End With
With Me.Shape1
.Visible = False
.BackColor = RGB(0, 0, 255)
.BorderColor = RGB(0, 0, 255)
End With
Exit Sub
End If
End Sub
Private Sub mnumadd_Click()
On Error GoTo ep
adddata = InputBox("请输入要在剪切列表中添加的项目!", "Add Data")
If adddata <> "" Then
List1.AddItem adddata
MsgBox "数据已经添加!" & Chr(13) & "添加了:" & adddata, vbExclamation, "Added"
Else
MsgBox "对不起,不允许添加空值!", vbExclamation, "Error"
End If
Exit Sub
ep:
MsgBox "发生了错误" & Chr(13) & "错误号:" & Err.Number & Chr(13) & "错误描述符:" & Err.Description, vbCritical, "Error"
End Sub
Private Sub mnuremove_Click()
On Error GoTo ep
If List1.ListIndex > -1 Then
d = MsgBox("确定移除所选定的项目吗?", vbQuestion + vbYesNo, "Remove?")
If d = vbYes Then
List1.RemoveItem (List1.ListIndex)
MsgBox "清除已经完成!", vbExclamation, "Cleared"
Command1.Enabled = False
Command4.Enabled = False
Command9.Enabled = False
Text1.Text = ""
Else
MsgBox "您尚未选择任何项目!", vbExclamation, "Error"
End If
End If
Exit Sub
ep:
MsgBox "发生了错误" & Chr(13) & "错误号:" & Err.Number & Chr(13) & "错误描述符:" & Err.Description, vbCritical, "Error"
End Sub
Private Sub Option1_Click()
On Error GoTo ep
Timer2.Enabled = True
ddd.Enabled = True
eee.Enabled = False
Exit Sub
ep:
MsgBox "发生了错误" & Chr(13) & "错误号:" & Err.Number & Chr(13) & "错误描述符:" & Err.Description, vbCritical, "Error"
End Sub
Private Sub Option2_Click()
On Error GoTo ep
Timer2.Enabled = False
ddd.Enabled = False
eee.Enabled = True
Exit Sub
ep:
MsgBox "发生了错误" & Chr(13) & "错误号:" & Err.Number & Chr(13) & "错误描述符:" & Err.Description, vbCritical, "Error"
End Sub
Private Sub Option3_Click()
Timer1.Enabled = True
di.Enabled = True
en.Enabled = False
End Sub
Private Sub Option4_Click()
Timer1.Enabled = False
di.Enabled = False
en.Enabled = True
End Sub
Private Sub pic_Click()
mni.InTray = False
Form1.Hide
Form3.show
End Sub
Private Sub show_Click()
mni.InTray = False
Form1.show
End Sub
Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If List1.ListIndex >= 0 And KeyCode = vbKeyC And Shift = 2 Then Clipboard.SetText List1.List(List1.ListIndex)
End Sub
Private Sub Timer1_Timer()
On Error GoTo ep
autodata = Clipboard.GetText
If autodata = "" Then Exit Sub
If List1.ListCount > 0 Then
Dim addvar As Integer
Dim cunt As Integer
For cunt = 0 To List1.ListCount - 1
If autodata <> List1.List(cunt) Then
addvar = addvar + 1
If addvar = List1.ListCount Then
List1.AddItem autodata
autovar = 0
End If
Else
Exit Sub
End If
Next
Else
List1.AddItem autodata
End If
Exit Sub
ep:
MsgBox "发生了错误" & Chr(13) & "错误号:" & Err.Number & Chr(13) & "错误描述符:" & Err.Description, vbCritical, "Error"
End Sub
Private Sub Timer2_Timer()
On Error GoTo ep
Clipboard.Clear
Exit Sub
ep:
MsgBox "发生了错误" & Chr(13) & "错误号:" & Err.Number & Chr(13) & "错误描述符:" & Err.Description, vbCritical, "Error"
End Sub
