VERSION 5.00
Begin VB.Form Form8 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Regsvr32参数说明"
   ClientHeight    =   2340
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7245
   LinkTopic       =   "Form8"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2340
   ScaleWidth      =   7245
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "确定"
      Default         =   -1  'True
      Height          =   510
      Left            =   5520
      TabIndex        =   0
      Top             =   1680
      Width           =   1590
   End
   Begin VB.Image Image1 
      Height          =   630
      Left            =   120
      Picture         =   "Form8.frx":0000
      Top             =   165
      Width           =   660
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Regsvr32.exe Help"
      Height          =   1320
      Left            =   960
      TabIndex        =   1
      Top             =   120
      Width           =   6135
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   1035
      Left            =   0
      Top             =   1560
      Width           =   7860
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command1_Click()
On Error Resume Next
Unload Me
Form1.Show
End Sub
Private Sub Form_Activate()
On Error Resume Next
Me.Command1.SetFocus
End Sub
Private Sub Form_Load()
On Error Resume Next
Form1.Visible = False
Me.Icon = LoadPicture("")
Label1.Caption = "Regsvr32.exe参数说明" & vbCrLf & "用法: regsvr32 [/u] [/s] [/n] [/i[:cmdline]] dllname" & vbCrLf & "/u -    解除服务器注册" & vbCrLf & "/s -    无声；不显示消息框" & vbCrLf & "/i -    调用 DllInstall，给其传递一个可选 [cmdline]；跟 /u 一起使用时，卸载 dll" & vbCrLf & "/n -    不要调用 DllRegisterServer；这个选项必须跟 /i 一起使用"
With Me.Command1
.Default = True
.Cancel = True
End With
End Sub
