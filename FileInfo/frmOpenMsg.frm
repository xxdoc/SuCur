VERSION 5.00
Begin VB.Form frmOpenMsg 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "正在打開文件..."
   ClientHeight    =   1155
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5340
   ControlBox      =   0   'False
   LinkTopic       =   "Form11"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1155
   ScaleWidth      =   5340
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "正在打開文件 %FilePath% ,請等待...."
      Height          =   1005
      Left            =   705
      TabIndex        =   0
      Top             =   60
      Width           =   4530
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   105
      Picture         =   "frmOpenMsg.frx":0000
      Top             =   75
      Width           =   480
   End
End
Attribute VB_Name = "frmOpenMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

