VERSION 5.00
Begin VB.Form Form10 
   BorderStyle     =   0  'None
   Caption         =   "Form10"
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form10"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   1740
      Top             =   1350
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Const WS_EX_LAYERED = &H80000
Private Const GWL_EXSTYLE = (-20)
Private Const LWA_ALPHA = &H2
Private Const LWA_COLORKEY = &H1
Private Sub Form_Load()
On Error Resume Next
Unload Me
If 1 = 2 Then
Dim rtn     As Long
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 1, LWA_ALPHA
HWND_TOPMOST = -1
SWP_NOSIZE = &H1
SWP_NOREDRAW = &H8
SWP_NOMOVE = &H2
SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Form10
.Top = frmMain.WebBrowser1.Top + frmMain.Top + 650
.Left = frmMain.WebBrowser1.Left + frmMain.Left
.Height = frmMain.WebBrowser1.Height
.Width = frmMain.WebBrowser1.Width - 200
.Enabled = False
.Visible = True
.Show
End With
End If
Unload Me
End Sub
Private Sub Timer1_Timer()
On Error Resume Next
Unload Me
Exit Sub
If 1 = 2 Then
If frmMain.act = True Then
Unload Me
Exit Sub
End If
If frmMain.Visible = False Then
Unload Me
Exit Sub
End If
With Form10
.Top = frmMain.WebBrowser1.Top + frmMain.Top
.Left = frmMain.WebBrowser1.Left + frmMain.Left
.Move frmMain.WebBrowser1.Left + frmMain.Left, frmMain.WebBrowser1.Top + frmMain.Top + 651
.Height = frmMain.WebBrowser1.Height
.Width = frmMain.WebBrowser1.Width - 200
.Enabled = False
End With
End If
End Sub
