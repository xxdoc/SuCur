VERSION 5.00
Begin VB.Form Form5 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Image Full Preview"
   ClientHeight    =   8565
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15270
   Icon            =   "Form5.frx":0000
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8565
   ScaleWidth      =   15270
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Height          =   255
      Left            =   14985
      TabIndex        =   3
      ToolTipText     =   "退出预览窗口"
      Top             =   8280
      Width           =   270
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      LargeChange     =   666
      Left            =   0
      SmallChange     =   100
      TabIndex        =   2
      Top             =   8280
      Width           =   14985
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   8265
      LargeChange     =   666
      Left            =   14985
      SmallChange     =   100
      TabIndex        =   1
      Top             =   0
      Width           =   270
   End
   Begin VB.PictureBox Picture1 
      Height          =   8250
      Left            =   0
      ScaleHeight     =   8190
      ScaleWidth      =   14910
      TabIndex        =   0
      Top             =   0
      Width           =   14970
      Begin VB.Image Image1 
         Height          =   6495
         Left            =   0
         Top             =   0
         Width           =   11055
      End
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Sub Command1_Click()
Unload Me
End Sub
Private Sub Form_Load()
On Error Resume Next
HWND_TOPMOST = -1
SWP_NOSIZE = &H1
SWP_NOREDRAW = &H8
SWP_NOMOVE = &H2
SetWindowPos Form5.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
End Sub
Private Sub HScroll1_Change()
On Error Resume Next
Image1.Left = -HScroll1.Value
If -HScroll1.Value > 0 Then
Image1.Left = HScroll1.Value
End If
End Sub
Private Sub VScroll1_Change()
On Error Resume Next
Image1.Top = -VScroll1.Value
If -VScroll1.Value > 0 Then
Image1.Top = VScroll1.Value
End If
End Sub
