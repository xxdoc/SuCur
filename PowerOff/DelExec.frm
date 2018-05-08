VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00C0C000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Delay Power Control"
   ClientHeight    =   4290
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6255
   Icon            =   "DelExec.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   6255
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C000&
      Caption         =   "选项"
      Height          =   1755
      Left            =   30
      TabIndex        =   6
      Top             =   1785
      Width           =   6210
      Begin VB.HScrollBar HScroll1 
         Height          =   255
         Left            =   2715
         Max             =   255
         Min             =   100
         TabIndex        =   20
         Top             =   870
         Value           =   199
         Width           =   2655
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00C0C000&
         Caption         =   "启用提示窗口半透明功能(&E)"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   870
         Width           =   2610
      End
      Begin VB.ComboBox Combo1 
         Height          =   300
         ItemData        =   "DelExec.frx":030A
         Left            =   1455
         List            =   "DelExec.frx":031A
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   1290
         Width           =   3990
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0C000&
         Caption         =   "在定时器时到后执行30秒延迟时间(推荐)(&E)"
         Height          =   375
         Left            =   135
         TabIndex        =   16
         Top             =   510
         Value           =   1  'Checked
         Width           =   5895
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         Height          =   270
         Left            =   4050
         MaxLength       =   2
         TabIndex        =   14
         Text            =   "30"
         Top             =   240
         Width           =   555
      End
      Begin VB.CommandButton Command8 
         Caption         =   "-"
         Height          =   285
         Left            =   3810
         TabIndex        =   13
         Top             =   240
         Width           =   240
      End
      Begin VB.CommandButton Command7 
         Caption         =   "+"
         Height          =   285
         Left            =   4605
         TabIndex        =   12
         Top             =   240
         Width           =   240
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   270
         Left            =   2535
         MaxLength       =   2
         TabIndex        =   9
         Text            =   "30"
         Top             =   240
         Width           =   555
      End
      Begin VB.CommandButton Command5 
         Caption         =   "-"
         Height          =   285
         Left            =   2295
         TabIndex        =   8
         Top             =   240
         Width           =   240
      End
      Begin VB.CommandButton Command6 
         Caption         =   "+"
         Height          =   285
         Left            =   3090
         TabIndex        =   7
         Top             =   240
         Width           =   240
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "199"
         Height          =   225
         Left            =   5430
         TabIndex        =   21
         Top             =   885
         Width           =   675
      End
      Begin VB.Image Image4 
         BorderStyle     =   1  'Fixed Single
         Height          =   540
         Left            =   5550
         Picture         =   "DelExec.frx":0342
         Top             =   1170
         Width           =   540
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "提示窗口图标"
         Height          =   180
         Left            =   135
         TabIndex        =   18
         Top             =   1335
         Width           =   1080
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "分"
         Height          =   180
         Left            =   4950
         TabIndex        =   15
         Top             =   285
         Width           =   180
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "时"
         Height          =   180
         Left            =   3435
         TabIndex        =   11
         Top             =   285
         Width           =   180
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "在何时关机?(精确度:分)"
         Height          =   360
         Left            =   135
         TabIndex        =   10
         Top             =   285
         Width           =   2070
      End
   End
   Begin VB.CommandButton Command4 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   435
      Left            =   4860
      TabIndex        =   5
      Top             =   3765
      Width           =   1290
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00C5791D&
      Caption         =   "强制关闭没有响应的进程(&F)"
      Height          =   375
      Left            =   165
      TabIndex        =   4
      Top             =   3795
      Width           =   2655
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C000&
      Caption         =   "想要计算机做什么?"
      ForeColor       =   &H00000000&
      Height          =   1740
      Left            =   1110
      TabIndex        =   0
      Top             =   30
      Width           =   5115
      Begin VB.CommandButton Command1 
         BackColor       =   &H000000FF&
         Caption         =   "关机(&S)"
         Height          =   1200
         Left            =   210
         MaskColor       =   &H00000000&
         Picture         =   "DelExec.frx":0784
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
         Left            =   1897
         MaskColor       =   &H00000000&
         Picture         =   "DelExec.frx":2BC6
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   345
         UseMaskColor    =   -1  'True
         Width           =   1365
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FFFF00&
         Caption         =   "注销(&L)"
         Height          =   1200
         Left            =   3585
         MaskColor       =   &H00000000&
         Picture         =   "DelExec.frx":4708
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
      Picture         =   "DelExec.frx":6B4A
      Top             =   150
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   2
      Left            =   0
      Picture         =   "DelExec.frx":6F8C
      Top             =   0
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   1
      Left            =   570
      Picture         =   "DelExec.frx":73CE
      Top             =   30
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   0
      Left            =   195
      Picture         =   "DelExec.frx":7810
      Top             =   165
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00C5791D&
      FillStyle       =   0  'Solid
      Height          =   705
      Left            =   0
      Top             =   3675
      Width           =   7995
   End
   Begin VB.Image Image3 
      Height          =   75
      Index           =   0
      Left            =   -195
      Picture         =   "DelExec.frx":7C52
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   6615
   End
   Begin VB.Image Image3 
      Height          =   480
      Index           =   1
      Left            =   480
      Picture         =   "DelExec.frx":80BE
      Top             =   555
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   165
      Picture         =   "DelExec.frx":83C8
      Top             =   225
      Width           =   720
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim KeyFlag As Boolean
Private Type TimeData
HourValue As Integer
MinuteValue As Integer
TimeUserSet As String
TimeSystem As String
ExecuteDelay As Integer
End Type
Dim TimeVar As TimeData
Private Sub Check1_GotFocus()
On Error Resume Next
KeyFlag = False
End Sub
Private Sub Check2_Click()
On Error Resume Next
If Check2.Value = 1 Then
Dim ans As Integer
ans = MsgBox("警告:不推荐强制结束进程,因为这样可能导致不可知的问题,继续?", vbExclamation + vbYesNo, "Alert")
If ans = vbYes Then
Exit Sub
Else
Check2.Value = 0
End If
End If
End Sub
Private Sub Check2_GotFocus()
On Error Resume Next
KeyFlag = False
End Sub
Private Sub Command1_Click()
On Error Resume Next
If Val(Text1.Text) >= 0 And Val(Text1.Text) <= 24 And Val(Text2.Text) >= 0 And Val(Text2.Text) <= 59 Then
Me.Hide
delShutdown.Show
Else
MsgBox "定时器时间输入不正确,请检查!", vbCritical, "Error"
Exit Sub
End If
End Sub
Private Sub Command1_GotFocus()
On Error Resume Next
KeyFlag = False
End Sub
Private Sub Command2_Click()
On Error Resume Next
If Val(Text1.Text) >= 0 And Val(Text1.Text) <= 24 And Val(Text2.Text) >= 0 And Val(Text2.Text) <= 59 Then
Me.Hide
delReboot.Show
Else
MsgBox "定时器时间输入不正确,请检查!", vbCritical, "Error"
Exit Sub
End If
End Sub
Private Sub Command2_GotFocus()
On Error Resume Next
KeyFlag = False
End Sub
Private Sub Command3_Click()
On Error Resume Next
If Val(Text1.Text) >= 0 And Val(Text1.Text) <= 24 And Val(Text2.Text) >= 0 And Val(Text2.Text) <= 59 Then
Me.Hide
delLogoff.Show
Else
MsgBox "定时器时间输入不正确,请检查!", vbCritical, "Error"
Exit Sub
End If
End Sub
Private Sub Command3_GotFocus()
On Error Resume Next
KeyFlag = False
End Sub
Private Sub Command4_Click()
On Error Resume Next
Unload Me
Form1.Show
End Sub
Private Sub Command4_GotFocus()
On Error Resume Next
KeyFlag = False
End Sub
Private Sub Command5_Click()
On Error Resume Next
If Val(Text1.Text) <= 0 Then
Text1.Text = 0
Exit Sub
End If
Text1.Text = Val(Text1.Text) - 1
End Sub
Private Sub Command5_GotFocus()
On Error Resume Next
KeyFlag = False
End Sub
Private Sub Command6_Click()
On Error Resume Next
If Val(Text1.Text) >= 24 Then
Text1.Text = 24
Exit Sub
End If
Text1.Text = Val(Text1.Text) + 1
End Sub
Private Sub Command6_GotFocus()
On Error Resume Next
KeyFlag = False
End Sub
Private Sub Command7_Click()
On Error Resume Next
If Val(Text2.Text) >= 59 Then
Text2.Text = 59
Exit Sub
End If
Text2.Text = Val(Text2.Text) + 1
End Sub
Private Sub Command7_GotFocus()
On Error Resume Next
KeyFlag = False
End Sub
Private Sub Command8_Click()
On Error Resume Next
If Val(Text2.Text) <= 0 Then
Text2.Text = 0
Exit Sub
End If
Text2.Text = Val(Text2.Text) - 1
End Sub
Private Sub Command8_GotFocus()
On Error Resume Next
KeyFlag = False
End Sub
Private Sub Form_Activate()
On Error Resume Next
KeyFlag = False
End Sub
Private Sub Form_GotFocus()
On Error Resume Next
KeyFlag = False
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyFlag = False Then
Select Case KeyAscii
Case 115
On Error Resume Next
If Val(Text1.Text) >= 0 And Val(Text1.Text) <= 24 And Val(Text2.Text) >= 0 And Val(Text2.Text) <= 59 Then
Me.Hide
delShutdown.Show
Else
MsgBox "定时器时间输入不正确,请检查!", vbCritical, "Error"
Exit Sub
End If
Case 114
On Error Resume Next
If Val(Text1.Text) >= 0 And Val(Text1.Text) <= 24 And Val(Text2.Text) >= 0 And Val(Text2.Text) <= 59 Then
Me.Hide
delReboot.Show
Else
MsgBox "定时器时间输入不正确,请检查!", vbCritical, "Error"
Exit Sub
End If
Case 108
On Error Resume Next
If Val(Text1.Text) >= 0 And Val(Text1.Text) <= 24 And Val(Text2.Text) >= 0 And Val(Text2.Text) <= 59 Then
Me.Hide
delLogoff.Show
Else
MsgBox "定时器时间输入不正确,请检查!", vbCritical, "Error"
Exit Sub
End If
End Select
End If
End Sub
Private Sub Form_Load()
On Error Resume Next
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
With TimeVar
.HourValue = Hour(Now)
.MinuteValue = Minute(Now)
Me.Text1.Text = .HourValue
Me.Text2.Text = .MinuteValue
End With
With Me
.Combo1.ListIndex = 2
.Image4.Picture = .Image2(2).Picture
End With
Me.KeyPreview = True
Me.Check1.Value = 1
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
End Select
End Sub
Private Sub Form_Unload(Cancel As Integer)
Form1.Show
End Sub
Private Sub Label4_Click()
On Error Resume Next
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
Private Sub Text1_GotFocus()
On Error Resume Next
KeyFlag = True
End Sub
Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = vbKeyUp Then
If Val(Text1.Text) >= 24 Then
Text1.Text = 24
Exit Sub
End If
Text1.Text = Val(Text1.Text) + 1
Err.Clear
ElseIf KeyCode = vbKeyRight Then
If Val(Text1.Text) >= 24 Then
Text1.Text = 24
Exit Sub
End If
Text1.Text = Val(Text1.Text) + 1
Err.Clear
ElseIf KeyCode = vbKeyRight Then
If Val(Text1.Text) <= 0 Then
Text1.Text = 0
Err.Clear
Exit Sub
End If
Text1.Text = Val(Text1.Text) - 1
Err.Clear
ElseIf KeyCode = vbKeyDown Then
If Val(Text1.Text) <= 0 Then
Text1.Text = 0
Err.Clear
Exit Sub
End If
Text1.Text = Val(Text1.Text) - 1
Err.Clear
Else
Exit Sub
End If
End Sub
Private Sub Text1_LostFocus()
On Error Resume Next
KeyFlag = False
End Sub
Private Sub Text2_GotFocus()
On Error Resume Next
KeyFlag = True
End Sub
Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = vbKeyUp Then
If Val(Text2.Text) >= 59 Then
Text2.Text = 59
Exit Sub
End If
Text2.Text = Val(Text2.Text) + 1
Err.Clear
ElseIf KeyCode = vbKeyRight Then
If Val(Text2.Text) >= 59 Then
Text2.Text = 59
Exit Sub
End If
Text2.Text = Val(Text2.Text) + 1
Err.Clear
ElseIf KeyCode = vbKeyRight Then
If Val(Text2.Text) <= 0 Then
Text2.Text = 0
Err.Clear
Exit Sub
End If
Text2.Text = Val(Text2.Text) - 1
Err.Clear
ElseIf KeyCode = vbKeyDown Then
If Val(Text2.Text) <= 0 Then
Text2.Text = 0
Err.Clear
Exit Sub
End If
Text2.Text = Val(Text2.Text) - 1
Err.Clear
Else
Exit Sub
End If
End Sub
Private Sub Text2_LostFocus()
On Error Resume Next
KeyFlag = False
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii < 48 Or KeyAscii > 57 Then
If KeyAscii <> 8 Then
KeyAscii = 0
End If
End If
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii < 48 Or KeyAscii > 57 Then
If KeyAscii <> 8 Then
KeyAscii = 0
End If
End If
End Sub
Private Sub Combo1_Change()
On Error Resume Next
Me.Image4.Picture = Me.Image2(Combo1.ListIndex).Picture
End Sub
Private Sub Combo1_Click()
On Error Resume Next
Me.Image4.Picture = Me.Image2(Combo1.ListIndex).Picture
End Sub
Private Sub Combo1_GotFocus()
On Error Resume Next
KeyFlag = False
End Sub
Private Sub Check3_Click()
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
End Sub
Private Sub Check3_GotFocus()
On Error Resume Next
KeyFlag = False
End Sub
Private Sub HScroll1_Change()
On Error Resume Next
Label4.Caption = Me.HScroll1.Value
End Sub
Private Sub HScroll1_GotFocus()
On Error Resume Next
KeyFlag = False
End Sub
