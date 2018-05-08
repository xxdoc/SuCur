VERSION 5.00
Begin VB.Form frmParam 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Param Info"
   ClientHeight    =   4965
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8865
   Icon            =   "frmParam.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4965
   ScaleWidth      =   8865
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   4920
      Left            =   45
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "frmParam.frx":030A
      Top             =   0
      Width           =   8775
   End
End
Attribute VB_Name = "frmParam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Load()
On Error Resume Next
With Text1
.Left = 0
.Top = 0
.Height = Me.ScaleHeight
.Width = Me.ScaleWidth
.BackColor = RGB(0, 0, 0)
.ForeColor = RGB(255, 255, 255)
.Locked = True
End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
With Form1
.Visible = True
End With
End Sub
