VERSION 5.00
Begin VB.Form Form6 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "选择一个磁盘格式化"
   ClientHeight    =   2055
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5400
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   5400
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   1320
      TabIndex        =   10
      Text            =   "Disk"
      Top             =   1035
      Width           =   3930
   End
   Begin VB.CheckBox Check2 
      Caption         =   "强制卸除卷(仅卷被占用时才推荐使用)(&F)"
      Height          =   300
      Left            =   270
      TabIndex        =   8
      Top             =   780
      Value           =   1  'Checked
      Width           =   4965
   End
   Begin VB.CheckBox Check1 
      Caption         =   "使用快速格式化(&U)"
      Height          =   300
      Left            =   3000
      TabIndex        =   7
      Top             =   465
      Value           =   1  'Checked
      Width           =   2190
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      ItemData        =   "Form6.frx":0000
      Left            =   1215
      List            =   "Form6.frx":0010
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   480
      Width           =   1710
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "退出(&C)"
      Height          =   390
      Left            =   3420
      TabIndex        =   3
      Top             =   1365
      Width           =   1755
   End
   Begin VB.CommandButton Command1 
      Caption         =   "开始格式化(&S)"
      Height          =   390
      Left            =   270
      TabIndex        =   2
      Top             =   1380
      Width           =   1755
   End
   Begin VB.DriveListBox Drive1 
      Height          =   300
      Left            =   795
      TabIndex        =   1
      Top             =   120
      Width           =   4350
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "卷标:"
      Height          =   225
      Left            =   300
      TabIndex        =   9
      Top             =   1095
      Width           =   840
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "文件系统:"
      Height          =   180
      Left            =   240
      TabIndex        =   5
      Top             =   480
      Width           =   810
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   30
      TabIndex        =   4
      Top             =   1785
      Width           =   5340
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
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Combo1_KeyPress(KeyAscii As Integer)
On Error GoTo ep
If KeyAscii = vbKeyReturn Then
Dim ans As Integer
ans = MsgBox("确认格式化该卷吗?所有数据都会丢失!", vbExclamation + vbYesNo, "Alert")
If ans = vbYes Then
Drive1.Enabled = False
Combo1.Enabled = False
Check1.Enabled = False
Check2.Enabled = False
Command1.Enabled = False
Command2.Enabled = False
Text1.Enabled = False
Select Case Check1.Value
Case 0
If Check2.Value = 0 Then
Shell "cmd.exe /k " & Chr(34) & "format " & Left(Drive1.List(Drive1.ListIndex), 2) & " /fs:" & Combo1.Text & " /v:" & Text1.Text, vbNormalFocus
Else
Shell "cmd.exe /k " & Chr(34) & "format " & Left(Drive1.List(Drive1.ListIndex), 2) & " /fs:" & Combo1.Text & " /x" & " /v:" & Text1.Text, vbNormalFocus
End If
Case 1
If Check2.Value = 0 Then
Shell "cmd.exe /k " & Chr(34) & "format " & Left(Drive1.List(Drive1.ListIndex), 2) & " /fs:" & Combo1.Text & " /q" & " /v:" & Text1.Text, vbNormalFocus
Else
Shell "cmd.exe /k " & Chr(34) & "format " & Left(Drive1.List(Drive1.ListIndex), 2) & " /fs:" & Combo1.Text & " /x /q" & " /v:" & Text1.Text, vbNormalFocus
End If
End Select
Drive1.Enabled = True
Combo1.Enabled = True
Check1.Enabled = True
Check2.Enabled = True
Command1.Enabled = True
Command2.Enabled = True
Text1.Enabled = True
End If
End If
Exit Sub
ep:
MsgBox "发生系统错误:" & Chr(13) & Err.Description & Chr(13) & "您的系统版本可能不支持此功能或者您尚未安装这个功能.", vbCritical, "Error"
Drive1.Enabled = True
Combo1.Enabled = True
Check1.Enabled = True
Check2.Enabled = True
Command1.Enabled = True
Command2.Enabled = True
Text1.Enabled = True
End Sub
Private Sub Command1_Click()
On Error GoTo ep
Dim ans As Integer
ans = MsgBox("确认格式化该卷吗?所有数据都会丢失!", vbExclamation + vbYesNo, "Alert")
If ans = vbYes Then
Drive1.Enabled = False
Combo1.Enabled = False
Check1.Enabled = False
Check2.Enabled = False
Command1.Enabled = False
Command2.Enabled = False
Text1.Enabled = False
Select Case Check1.Value
Case 0
If Check2.Value = 0 Then
Shell "cmd.exe /k " & Chr(34) & "format " & Left(Drive1.List(Drive1.ListIndex), 2) & " /fs:" & Combo1.Text & " /v:" & Text1.Text & " /y", vbNormalFocus
Else
Shell "cmd.exe /k " & Chr(34) & "format " & Left(Drive1.List(Drive1.ListIndex), 2) & " /fs:" & Combo1.Text & " /x /y" & " /v:" & Text1.Text, vbNormalFocus
End If
Case 1
If Check2.Value = 0 Then
Shell "cmd.exe /k " & Chr(34) & "format " & Left(Drive1.List(Drive1.ListIndex), 2) & " /fs:" & Combo1.Text & " /q /y" & " /v:" & Text1.Text, vbNormalFocus
Else
Shell "cmd.exe /k " & Chr(34) & "format " & Left(Drive1.List(Drive1.ListIndex), 2) & " /fs:" & Combo1.Text & " /x /q /y" & " /v:" & Text1.Text, vbNormalFocus
End If
End Select
Drive1.Enabled = True
Combo1.Enabled = True
Check1.Enabled = True
Check2.Enabled = True
Command1.Enabled = True
Command2.Enabled = True
Text1.Enabled = True
End If
Exit Sub
ep:
MsgBox "发生系统错误:" & Chr(13) & Err.Description & Chr(13) & "您的系统版本可能不支持此功能或者您尚未安装这个功能.", vbCritical, "Error"
Drive1.Enabled = True
Combo1.Enabled = True
Check1.Enabled = True
Check2.Enabled = True
Command1.Enabled = True
Command2.Enabled = True
Text1.Enabled = True
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
Dim ans As Integer
ans = MsgBox("确认格式化该卷吗?所有数据都会丢失!", vbExclamation + vbYesNo, "Alert")
If ans = vbYes Then
Drive1.Enabled = False
Combo1.Enabled = False
Check1.Enabled = False
Check2.Enabled = False
Command1.Enabled = False
Command2.Enabled = False
Text1.Enabled = False
Select Case Check1.Value
Case 0
If Check2.Value = 0 Then
Shell "cmd.exe /k " & Chr(34) & "format " & Left(Drive1.List(Drive1.ListIndex), 2) & " /fs:" & Combo1.Text & " /v:" & Text1.Text & " /y", vbNormalFocus
Else
Shell "cmd.exe /k " & Chr(34) & "format " & Left(Drive1.List(Drive1.ListIndex), 2) & " /fs:" & Combo1.Text & " /x /y" & " /v:" & Text1.Text, vbNormalFocus
End If
Case 1
If Check2.Value = 0 Then
Shell "cmd.exe /k " & Chr(34) & "format " & Left(Drive1.List(Drive1.ListIndex), 2) & " /fs:" & Combo1.Text & " /q /y" & " /v:" & Text1.Text, vbNormalFocus
Else
Shell "cmd.exe /k " & Chr(34) & "format " & Left(Drive1.List(Drive1.ListIndex), 2) & " /fs:" & Combo1.Text & " /x /q /y" & " /v:" & Text1.Text, vbNormalFocus
End If
End Select
Drive1.Enabled = True
Combo1.Enabled = True
Check1.Enabled = True
Check2.Enabled = True
Command1.Enabled = True
Command2.Enabled = True
Text1.Enabled = True
End If
Exit Sub
ep:
MsgBox "发生系统错误:" & Chr(13) & Err.Description & Chr(13) & "您的系统版本可能不支持此功能或者您尚未安装这个功能.", vbCritical, "Error"
Drive1.Enabled = True
Combo1.Enabled = True
Check1.Enabled = True
Check2.Enabled = True
Command1.Enabled = True
Command2.Enabled = True
Text1.Enabled = True
End Sub
Private Sub Form_Load()
On Error Resume Next
frmMain.Visible = False
Me.Command1.Default = True
Me.Command2.Cancel = True
Label2.Caption = "当前驱动器:      " & Left(Drive1.Drive, 2)
Combo1.ListIndex = 0
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
On Error GoTo ep
Dim ans As Integer
ans = MsgBox("确认格式化该卷吗?所有数据都会丢失!", vbExclamation + vbYesNo, "Alert")
If ans = vbYes Then
Drive1.Enabled = False
Combo1.Enabled = False
Check1.Enabled = False
Check2.Enabled = False
Command1.Enabled = False
Command2.Enabled = False
Text1.Enabled = False
Select Case Check1.Value
Case 0
If Check2.Value = 0 Then
Shell "cmd.exe /k " & Chr(34) & "format " & Left(Drive1.List(Drive1.ListIndex), 2) & " /fs:" & Combo1.Text & " /v:" & Text1.Text & " /y", vbNormalFocus
Else
Shell "cmd.exe /k " & Chr(34) & "format " & Left(Drive1.List(Drive1.ListIndex), 2) & " /fs:" & Combo1.Text & " /x /y" & " /v:" & Text1.Text, vbNormalFocus
End If
Case 1
If Check2.Value = 0 Then
Shell "cmd.exe /k " & Chr(34) & "format " & Left(Drive1.List(Drive1.ListIndex), 2) & " /fs:" & Combo1.Text & " /q /y" & " /v:" & Text1.Text, vbNormalFocus
Else
Shell "cmd.exe /k " & Chr(34) & "format " & Left(Drive1.List(Drive1.ListIndex), 2) & " /fs:" & Combo1.Text & " /x /q /y" & " /v:" & Text1.Text, vbNormalFocus
End If
End Select
Drive1.Enabled = True
Combo1.Enabled = True
Check1.Enabled = True
Check2.Enabled = True
Command1.Enabled = True
Command2.Enabled = True
Text1.Enabled = True
End If
Exit Sub
ep:
MsgBox "发生系统错误:" & Chr(13) & Err.Description & Chr(13) & "您的系统版本可能不支持此功能或者您尚未安装这个功能.", vbCritical, "Error"
Drive1.Enabled = True
Combo1.Enabled = True
Check1.Enabled = True
Check2.Enabled = True
Command1.Enabled = True
Command2.Enabled = True
Text1.Enabled = True
End Sub
