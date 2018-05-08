VERSION 5.00
Begin VB.Form Form4 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "新建任务(运行)"
   ClientHeight    =   1905
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5055
   Icon            =   "Form4.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1905
   ScaleWidth      =   5055
   Begin VB.ComboBox Text1 
      Height          =   300
      Left            =   60
      TabIndex        =   3
      Top             =   1080
      Width           =   4950
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   390
      Left            =   3780
      TabIndex        =   2
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "运行(&R)"
      Enabled         =   0   'False
      Height          =   390
      Left            =   2505
      TabIndex        =   1
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "    请输入应用程序/资源的名称,敲击回车键或单击'运行'按钮,Windows将为您打开它"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1110
      TabIndex        =   0
      Top             =   105
      Width           =   3825
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   300
      Picture         =   "Form4.frx":0442
      Stretch         =   -1  'True
      Top             =   285
      Width           =   495
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Sub Command1_Click()
On Error GoTo ep
Dim APIRetVal As Long
Dim cunt As Integer
Dim forvrt As Integer
If Trim(Text1.Text) = "" Then Exit Sub
APIRetVal = ShellExecute(Me.hwnd, "Open", Text1.Text, vbNullString, vbNullString, 3)
If APIRetVal <= 32 Then
Shell Text1.Text, vbNormalFocus
End If
If Text1.ListCount = 0 Then
Text1.AddItem Text1.Text
Else
For forvrt = 0 To Text1.ListCount - 1
If Text1.Text <> Text1.List(forvrt) Then
cunt = cunt + 1
If cunt = Text1.ListCount Then
Text1.AddItem Text1.Text
End If
Else
Exit For
End If
Next
End If
Exit Sub
ep:
MsgBox "发生错误:" & Chr(13) & Err.Description & Chr(13) & "请检查输入的路径/程序或文件名有效且有权限访问", vbCritical, "Error"
End Sub
Private Sub Command2_Click()
On Error Resume Next
Unload Me
End Sub
Private Sub Command3_Click()
On Error GoTo ep
Text1.SetFocus
Exit Sub
ep:
If Err.Number = 32755 Then Exit Sub
MsgBox "发生错误" & Chr(13) & Err.Description & Chr(13) & "也许是不正当的操作造成的", vbCritical, "Error"
End Sub
Private Sub Form_Activate()
On Error Resume Next
Text1.SetFocus
End Sub
Private Sub Form_Load()
On Error Resume Next
Form1.Visible = False
Command1.Default = True
Command2.Cancel = True
End Sub
Private Sub Text1_Change()
On Error Resume Next
If Trim(Text1.Text) = "" Then
Command1.Enabled = False
Else
Command1.Enabled = True
End If
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
On Error GoTo ep
If KeyAscii = vbKeyReturn Then
Dim APIRetVal As Long
Dim cunt As Integer
Dim forvrt As Integer
If Trim(Text1.Text) = "" Then Exit Sub
APIRetVal = ShellExecute(Me.hwnd, "Open", Text1.Text, vbNullString, vbNullString, 3)
If APIRetVal <= 32 Then
Shell Text1.Text, vbNormalFocus
End If
If Text1.ListCount = 0 Then
Text1.AddItem Text1.Text
Else
For forvrt = 0 To Text1.ListCount - 1
If Text1.Text <> Text1.List(forvrt) Then
cunt = cunt + 1
If cunt = Text1.ListCount Then
Text1.AddItem Text1.Text
End If
Else
Exit For
End If
Next
End If
End If
Exit Sub
ep:
MsgBox "发生错误:" & Chr(13) & Err.Description & Chr(13) & "请检查输入的路径/程序或文件名有效且有权限访问", vbCritical, "Error"
End Sub
