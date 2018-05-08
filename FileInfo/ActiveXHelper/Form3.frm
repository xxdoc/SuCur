VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "应用程序选项"
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8100
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   8100
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Apply 
      Caption         =   "应用(&A)"
      Enabled         =   0   'False
      Height          =   420
      Left            =   6825
      TabIndex        =   17
      Top             =   3105
      Width           =   1230
   End
   Begin VB.Frame Frame1 
      Caption         =   "文件浏览器选项"
      Height          =   1935
      Left            =   75
      TabIndex        =   8
      Top             =   60
      Width           =   7935
      Begin VB.ComboBox Combo1 
         Height          =   300
         ItemData        =   "Form3.frx":0000
         Left            =   120
         List            =   "Form3.frx":0019
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   1500
         Width           =   7710
      End
      Begin VB.CheckBox Check1 
         Caption         =   "系统(&S)"
         Height          =   255
         Index           =   4
         Left            =   4800
         TabIndex        =   13
         Top             =   720
         Width           =   1005
      End
      Begin VB.CheckBox Check1 
         Caption         =   "隐藏(&H)"
         Height          =   255
         Index           =   3
         Left            =   6480
         TabIndex        =   12
         Top             =   720
         Width           =   1005
      End
      Begin VB.CheckBox Check1 
         Caption         =   "存档(&A)"
         Height          =   255
         Index           =   2
         Left            =   3120
         TabIndex        =   11
         Top             =   720
         Width           =   1005
      End
      Begin VB.CheckBox Check1 
         Caption         =   "标准(&N)"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   1005
      End
      Begin VB.CheckBox Check1 
         Caption         =   "只读(&R)"
         Height          =   255
         Index           =   1
         Left            =   1590
         TabIndex        =   9
         Top             =   720
         Width           =   1005
      End
      Begin VB.Label Label2 
         Caption         =   "显示以下类型的文件:"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   1200
         Width           =   3015
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "显示具有以下属性的文件:"
         Height          =   180
         Left            =   120
         TabIndex        =   15
         Top             =   405
         Width           =   2070
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "处理性能"
      Height          =   1005
      Left            =   75
      TabIndex        =   2
      Top             =   2025
      Width           =   7935
      Begin VB.HScrollBar slider1 
         Height          =   255
         LargeChange     =   10
         Left            =   945
         Max             =   100
         Min             =   10
         SmallChange     =   5
         TabIndex        =   18
         Top             =   570
         Value           =   10
         Width           =   5910
      End
      Begin VB.CommandButton Command3 
         Caption         =   "输入(&T)"
         Height          =   300
         Left            =   6960
         TabIndex        =   3
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "性能:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "最高速度"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   630
         Width           =   735
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "最大保护"
         Height          =   255
         Index           =   1
         Left            =   7080
         TabIndex        =   5
         Top             =   645
         Width           =   735
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   960
         TabIndex        =   4
         Top             =   240
         Width           =   5895
      End
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   420
      Left            =   5505
      TabIndex        =   1
      Top             =   3105
      Width           =   1230
   End
   Begin VB.CommandButton Command1 
      Caption         =   "保存(&V)"
      Default         =   -1  'True
      Height          =   420
      Left            =   4185
      TabIndex        =   0
      Top             =   3105
      Width           =   1230
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click(Index As Integer)
On Error Resume Next
Apply.Enabled = True
End Sub
Private Sub Check1_KeyPress(Index As Integer, KeyAscii As Integer)
On Error Resume Next
If KeyAscii = vbKeyReturn Then
On Error Resume Next
If Check1(0).Value = 0 And Check1(1).Value = 0 And Check1(2).Value = 0 And Check1(3).Value = 0 And Check1(4).Value = 0 Then
MsgBox "没有勾选可显示的文件属性任何项目!", vbCritical, "Error"
Exit Sub
End If
Form1.RegSvr.Interval = Me.Slider1.Value
Form1.UnRegSvr.Interval = Me.Slider1.Value
Form1.normalspeed.Checked = False
Form1.anti.Checked = False
Form1.highspeed.Checked = False
Form1.COST.Checked = True
Dim ans As Integer
If Slider1.Value <= 15 And Form1.highspeed.Checked = False Then
ans = MsgBox("您选择的值可能太小,可能导致程序假死,仅建议在配置较好的计算机上使用,继续吗?" & vbCrLf & "点击[是]继续" & vbCrLf & "点击[否]启用默认设置(每25毫秒执行一次)", vbQuestion + vbYesNo, "Ask")
If ans = vbYes Then
Form1.RegSvr.Interval = Me.Slider1.Value
Form1.UnRegSvr.Interval = Me.Slider1.Value
Form1.normalspeed.Checked = False
Form1.anti.Checked = False
Form1.highspeed.Checked = False
Form1.COST.Checked = True
Else
Form1.RegSvr.Interval = 25
Form1.UnRegSvr.Interval = 25
Me.Slider1.Value = 25
Form1.normalspeed.Checked = True
Form1.anti.Checked = False
Form1.highspeed.Checked = False
Form1.COST.Checked = False
End If
End If
If Slider1.Value >= 75 And Form1.anti.Checked = False Then
ans = MsgBox("您选择的值可能太大,可能导致程序执行缓慢,仅建议在配置较差的计算机上使用,继续吗?" & vbCrLf & "点击[是]继续" & vbCrLf & "点击[否]启用默认设置(每25毫秒执行一次)", vbQuestion + vbYesNo, "Ask")
If ans = vbYes Then
Form1.RegSvr.Interval = Me.Slider1.Value
Form1.UnRegSvr.Interval = Me.Slider1.Value
Form1.normalspeed.Checked = False
Form1.anti.Checked = False
Form1.highspeed.Checked = False
Form1.COST.Checked = True
Else
Form1.RegSvr.Interval = 25
Form1.UnRegSvr.Interval = 25
Me.Slider1.Value = 25
Form1.normalspeed.Checked = True
Form1.anti.Checked = False
Form1.highspeed.Checked = False
Form1.COST.Checked = False
End If
End If
Select Case Check1(0).Value
Case 0
Form1.File1.Normal = False
Case 1
Form1.File1.Normal = True
End Select
Select Case Check1(1).Value
Case 0
Form1.File1.ReadOnly = False
Case 1
Form1.File1.ReadOnly = True
End Select
Select Case Check1(2).Value
Case 0
Form1.File1.Archive = False
Case 1
Form1.File1.Archive = True
End Select
Select Case Check1(3).Value
Case 0
Form1.File1.Hidden = False
Case 1
Form1.File1.Hidden = True
End Select
Select Case Check1(4).Value
Case 0
Form1.File1.System = False
Case 1
Form1.File1.System = True
End Select
Form1.File1.Pattern = Combo1.List(Combo1.ListIndex)
Form1.File1.Refresh
Form1.normalspeed.Checked = False
Form1.highspeed.Checked = False
Form1.anti.Checked = False
Form1.COST.Checked = False
If Slider1.Value = 25 Then
Form1.normalspeed.Checked = True
Form1.highspeed.Checked = False
Form1.anti.Checked = False
Form1.COST.Checked = False
Unload Me
Exit Sub
End If
If Slider1.Value = 10 Then
Form1.normalspeed.Checked = False
Form1.highspeed.Checked = True
Form1.anti.Checked = False
Form1.COST.Checked = False
Unload Me
Exit Sub
End If
If Slider1.Value = 100 Then
Form1.normalspeed.Checked = False
Form1.highspeed.Checked = False
Form1.anti.Checked = True
Form1.COST.Checked = False
Unload Me
Exit Sub
End If
Form1.normalspeed.Checked = False
Form1.highspeed.Checked = False
Form1.anti.Checked = False
Form1.COST.Checked = True
Unload Me
End If
End Sub
Private Sub Command1_Click()
On Error Resume Next
If Check1(0).Value = 0 And Check1(1).Value = 0 And Check1(2).Value = 0 And Check1(3).Value = 0 And Check1(4).Value = 0 Then
MsgBox "没有勾选可显示的文件属性任何项目!", vbCritical, "Error"
Exit Sub
End If
Form1.RegSvr.Interval = Me.Slider1.Value
Form1.UnRegSvr.Interval = Me.Slider1.Value
Form1.normalspeed.Checked = False
Form1.anti.Checked = False
Form1.highspeed.Checked = False
Form1.COST.Checked = True
Dim ans As Integer
If Slider1.Value <= 15 And Form1.highspeed.Checked = False Then
ans = MsgBox("您选择的值可能太小,可能导致程序假死,仅建议在配置较好的计算机上使用,继续吗?" & vbCrLf & "点击[是]继续" & vbCrLf & "点击[否]启用默认设置(每25毫秒执行一次)", vbQuestion + vbYesNo, "Ask")
If ans = vbYes Then
Form1.RegSvr.Interval = Me.Slider1.Value
Form1.UnRegSvr.Interval = Me.Slider1.Value
Form1.normalspeed.Checked = False
Form1.anti.Checked = False
Form1.highspeed.Checked = False
Form1.COST.Checked = True
Form1.Show
Else
Form1.RegSvr.Interval = 25
Form1.UnRegSvr.Interval = 25
Me.Slider1.Value = 25
Form1.normalspeed.Checked = True
Form1.anti.Checked = False
Form1.highspeed.Checked = False
Form1.COST.Checked = False
Form1.Show
End If
End If
If Slider1.Value >= 75 And Form1.anti.Checked = False Then
ans = MsgBox("您选择的值可能太大,可能导致程序执行缓慢,仅建议在配置较差的计算机上使用,继续吗?" & vbCrLf & "点击[是]继续" & vbCrLf & "点击[否]启用默认设置(每25毫秒执行一次)", vbQuestion + vbYesNo, "Ask")
If ans = vbYes Then
Form1.RegSvr.Interval = Me.Slider1.Value
Form1.UnRegSvr.Interval = Me.Slider1.Value
Form1.normalspeed.Checked = False
Form1.anti.Checked = False
Form1.highspeed.Checked = False
Form1.COST.Checked = True
Form1.Show
Else
Form1.RegSvr.Interval = 25
Form1.UnRegSvr.Interval = 25
Me.Slider1.Value = 25
Form1.normalspeed.Checked = True
Form1.anti.Checked = False
Form1.highspeed.Checked = False
Form1.COST.Checked = False
Form1.Show
End If
End If
Select Case Check1(0).Value
Case 0
Form1.File1.Normal = False
Case 1
Form1.File1.Normal = True
End Select
Select Case Check1(1).Value
Case 0
Form1.File1.ReadOnly = False
Case 1
Form1.File1.ReadOnly = True
End Select
Select Case Check1(2).Value
Case 0
Form1.File1.Archive = False
Case 1
Form1.File1.Archive = True
End Select
Select Case Check1(3).Value
Case 0
Form1.File1.Hidden = False
Case 1
Form1.File1.Hidden = True
End Select
Select Case Check1(4).Value
Case 0
Form1.File1.System = False
Case 1
Form1.File1.System = True
End Select
Form1.File1.Pattern = Combo1.List(Combo1.ListIndex)
Form1.File1.Refresh
Form1.normalspeed.Checked = False
Form1.highspeed.Checked = False
Form1.anti.Checked = False
Form1.COST.Checked = False
If Slider1.Value = 25 Then
Form1.normalspeed.Checked = True
Form1.highspeed.Checked = False
Form1.anti.Checked = False
Form1.COST.Checked = False
Unload Me
Form1.Show
Exit Sub
End If
If Slider1.Value = 10 Then
Form1.normalspeed.Checked = False
Form1.highspeed.Checked = True
Form1.anti.Checked = False
Form1.COST.Checked = False
Unload Me
Form1.Show
Exit Sub
End If
If Slider1.Value = 100 Then
Form1.normalspeed.Checked = False
Form1.highspeed.Checked = False
Form1.anti.Checked = True
Form1.COST.Checked = False
Unload Me
Form1.Show
Exit Sub
End If
Form1.normalspeed.Checked = False
Form1.highspeed.Checked = False
Form1.anti.Checked = False
Form1.COST.Checked = True
Unload Me
Form1.Show
End Sub
Private Sub Command2_Click()
Form1.Visible = True
Unload Me
Form1.Show
End Sub
Private Sub Command3_Click()
On Error Resume Next
Dim ms As Integer
ms = InputBox("请输入要使循环注册/反注册控件时间隔的时间,单位毫秒,范围5-100", "Timer Value", 25)
If 5 <= ms And ms <= 100 Then
Me.Slider1.Value = ms
Dim pct As Single
pct = Slider1.Value / 100 * 100
Label5.Caption = "加速" & 100 - pct & "%                " & "循环注册/反注册控件时每隔" & Me.Slider1.Value & "毫秒执行一次"
Apply.Enabled = True
Else
MsgBox "输入错误,请检查输入!", vbCritical, "Error"
End If
End Sub
Private Sub Form_Load()
On Error Resume Next
Form1.Visible = False
Me.Command2.Cancel = True
Me.Command1.Default = True
Me.Slider1.Value = Form1.RegSvr.Interval
Dim pct As Single
pct = Slider1.Value / 100 * 100
Label5.Caption = "加速" & 100 - pct & "%                " & "循环注册/反注册控件时每隔" & Me.Slider1.Value & "毫秒执行一次"
If Form1.File1.Archive = True Then Check1(2).Value = 1
If Form1.File1.Normal = True Then Check1(0).Value = 1
If Form1.File1.ReadOnly = True Then Check1(1).Value = 1
If Form1.File1.Hidden = True Then Check1(3).Value = 1
If Form1.File1.System = True Then Check1(4).Value = 1
Combo1.ListIndex = 6
Apply.Enabled = False
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Form1.Show
End Sub
Private Sub Slider1_Change()
On Error Resume Next
Dim pct As Single
pct = Slider1.Value / 100 * 100
Label5.Caption = "加速" & 100 - pct & "%                " & "循环注册/反注册控件时每隔" & Me.Slider1.Value & "毫秒执行一次"
Apply.Enabled = True
End Sub
Private Sub Apply_Click()
On Error Resume Next
Form1.RegSvr.Interval = Me.Slider1.Value
Form1.UnRegSvr.Interval = Me.Slider1.Value
Form1.normalspeed.Checked = False
Form1.anti.Checked = False
Form1.highspeed.Checked = False
Form1.COST.Checked = True
Dim ans As Integer
If Slider1.Value <= 15 And Form1.highspeed.Checked = False Then
ans = MsgBox("您选择的值可能太小,可能导致程序假死,仅建议在配置较好的计算机上使用,继续吗?" & vbCrLf & "点击[是]继续" & vbCrLf & "点击[否]启用默认设置(每25毫秒执行一次)", vbQuestion + vbYesNo, "Ask")
If ans = vbYes Then
Form1.RegSvr.Interval = Me.Slider1.Value
Form1.UnRegSvr.Interval = Me.Slider1.Value
Form1.normalspeed.Checked = False
Form1.anti.Checked = False
Form1.highspeed.Checked = False
Form1.COST.Checked = True
Else
Form1.RegSvr.Interval = 25
Form1.UnRegSvr.Interval = 25
Me.Slider1.Value = 25
Form1.normalspeed.Checked = True
Form1.anti.Checked = False
Form1.highspeed.Checked = False
Form1.COST.Checked = False
End If
End If
If Slider1.Value >= 75 And Form1.anti.Checked = False Then
ans = MsgBox("您选择的值可能太大,可能导致程序执行缓慢,仅建议在配置较差的计算机上使用,继续吗?" & vbCrLf & "点击[是]继续" & vbCrLf & "点击[否]启用默认设置(每25毫秒执行一次)", vbQuestion + vbYesNo, "Ask")
If ans = vbYes Then
Form1.RegSvr.Interval = Me.Slider1.Value
Form1.UnRegSvr.Interval = Me.Slider1.Value
Form1.normalspeed.Checked = False
Form1.anti.Checked = False
Form1.highspeed.Checked = False
Form1.COST.Checked = True
Else
Form1.RegSvr.Interval = 25
Form1.UnRegSvr.Interval = 25
Me.Slider1.Value = 25
Form1.normalspeed.Checked = True
Form1.anti.Checked = False
Form1.highspeed.Checked = False
Form1.COST.Checked = False
End If
End If
Form1.normalspeed.Checked = False
Form1.highspeed.Checked = False
Form1.anti.Checked = False
Form1.COST.Checked = False
If Slider1.Value = 25 Then
Form1.normalspeed.Checked = True
Form1.highspeed.Checked = False
Form1.anti.Checked = False
Form1.COST.Checked = False
Apply.Enabled = False
Exit Sub
End If
If Slider1.Value = 10 Then
Form1.normalspeed.Checked = False
Form1.highspeed.Checked = True
Form1.anti.Checked = False
Form1.COST.Checked = False
Apply.Enabled = False
Exit Sub
End If
If Slider1.Value = 100 Then
Form1.normalspeed.Checked = False
Form1.highspeed.Checked = False
Form1.anti.Checked = True
Form1.COST.Checked = False
Apply.Enabled = False
Exit Sub
End If
Form1.normalspeed.Checked = False
Form1.highspeed.Checked = False
Form1.anti.Checked = False
Form1.COST.Checked = True
Apply.Enabled = False
End Sub
