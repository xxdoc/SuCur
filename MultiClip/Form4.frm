VERSION 5.00
Begin VB.Form Form4 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "请选择导入文件的方式"
   ClientHeight    =   1530
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4335
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1530
   ScaleWidth      =   4335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "确定(&O)"
      CausesValidation=   0   'False
      Default         =   -1  'True
      Height          =   390
      Left            =   3090
      TabIndex        =   3
      Top             =   1080
      Width           =   1200
   End
   Begin VB.OptionButton Option2 
      Caption         =   "将整个文件作为一个记录导入待复制列表(&W)"
      Height          =   315
      Left            =   60
      TabIndex        =   2
      Top             =   615
      Value           =   -1  'True
      Width           =   4020
   End
   Begin VB.OptionButton Option1 
      Caption         =   "把每一行作为一个记录导入待复制列表(&E)"
      Height          =   315
      Left            =   60
      TabIndex        =   1
      Top             =   255
      Width           =   3645
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "请选择导入文件的方式"
      Height          =   180
      Left            =   60
      TabIndex        =   0
      Top             =   15
      Width           =   1800
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command1_Click()
On Error Resume Next
If Option1.Value = True Then
Form1.Tag = "Line"
ElseIf Option2.Value = True Then
Form1.Tag = "Whole"
Else
Form1.Tag = "Whole"
End If
Unload Me
End Sub
Private Sub Form_Load()
On Error Resume Next
Command1.Default = True
End Sub
