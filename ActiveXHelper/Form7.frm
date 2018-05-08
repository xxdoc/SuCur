VERSION 5.00
Begin VB.Form Form7 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "选择一个磁盘来扫描并修复错误"
   ClientHeight    =   1350
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5400
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1350
   ScaleWidth      =   5400
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "退出(&C)"
      Height          =   390
      Left            =   3420
      TabIndex        =   3
      Top             =   570
      Width           =   1755
   End
   Begin VB.CommandButton Command1 
      Caption         =   "开始扫描(&S)"
      Height          =   390
      Left            =   270
      TabIndex        =   2
      Top             =   555
      Width           =   1755
   End
   Begin VB.DriveListBox Drive1 
      Height          =   300
      Left            =   795
      TabIndex        =   1
      Top             =   120
      Width           =   4350
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   45
      TabIndex        =   4
      Top             =   1065
      Width           =   5295
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "磁盘:"
      Height          =   255
      Left            =   255
      TabIndex        =   0
      Top             =   180
      Width           =   900
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command1_Click()
On Error GoTo ep
Dim ans As Integer
ans = MsgBox("确定扫描文件系统吗?", vbQuestion + vbYesNo, "Ask")
If ans = vbYes Then
Command2.Enabled = False
Command1.Enabled = False
Drive1.Enabled = False
Shell "cmd.exe /k " & Chr(34) & "chkdsk " & Left(Drive1.List(Drive1.ListIndex), 2) & " /x /f", vbNormalFocus
Command1.Enabled = True
Command2.Enabled = True
Drive1.Enabled = True
End If
Exit Sub
ep:
MsgBox "发生系统错误:" & Chr(13) & Err.Description & Chr(13) & "您的系统版本可能不支持此功能或者您尚未安装这个功能.", vbCritical, "Error"
Command1.Enabled = True
Command2.Enabled = True
Drive1.Enabled = True
End Sub
Private Sub Command2_Click()
On Error Resume Next
Unload Me
End Sub
Private Sub Drive1_Change()
On Error GoTo ep
Label2.Caption = "当前驱动器:      " & Left(Drive1.Drive, 2)
Exit Sub
ep:
MsgBox "发生错误:" & Chr(13) & Err.Description & Chr(13) & "请检查设备是否可以正常使用", vbCritical, "Error"
Drive1.Drive = "C:"
End Sub
Private Sub Drive1_KeyPress(KeyAscii As Integer)
On Error GoTo ep
If KeyAscii = vbKeyReturn Then
Dim ans As Integer
ans = MsgBox("确定扫描文件系统吗?", vbQuestion + vbYesNo, "Ask")
If ans = vbYes Then
Command2.Enabled = False
Command1.Enabled = False
Drive1.Enabled = False
Shell "cmd.exe /k " & Chr(34) & "chkdsk " & Left(Drive1.List(Drive1.ListIndex), 2) & " /x /f", vbNormalFocus
Command1.Enabled = True
Command2.Enabled = True
Drive1.Enabled = True
End If
End If
Exit Sub
ep:
MsgBox "发生系统错误:" & Chr(13) & Err.Description & Chr(13) & "您的系统版本可能不支持此功能或者您尚未安装这个功能.", vbCritical, "Error"
Command1.Enabled = True
Command2.Enabled = True
Drive1.Enabled = True
End Sub
Private Sub Form_Load()
On Error Resume Next
frmMain.Visible = False
Label2.Caption = "当前驱动器:      " & Left(Drive1.Drive, 2)
Me.Command1.Default = True
Me.Command2.Cancel = True
End Sub
