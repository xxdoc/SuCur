Attribute VB_Name = "Balloon"
Option Explicit


Public Const WM_USER = &H400
Public Const WM_NOTIFYICON = WM_USER + 1               '   自定义消息


   
'   关于气球提示的自定义消息,   2000下不产生这些消息
Public Const NIN_BALLOONSHOW = (WM_USER + &H2)               '   当   Balloon   Tips   弹出时执行
Public Const NIN_BALLOONHIDE = (WM_USER + &H3)               '   当   Balloon   Tips   消失时执行（如   SysTrayIcon   被删除），
'   但指定的   TimeOut   时间到或鼠标点击   Balloon   Tips   后的消失不发送此消息
Public Const NIN_BALLOONTIMEOUT = (WM_USER + &H4)               '   当   Balloon   Tips   的   TimeOut   时间到时执行
Public Const NIN_BALLOONUSERCLICK = (WM_USER + &H5)               '   当鼠标点击   Balloon   Tips   时执行。
'   注意:在XP下执行时   Balloon   Tips   上有个关闭按钮,
'   如果鼠标点在按钮上将接收到   NIN_BALLOONTIMEOUT   消息。
   
Public preWndProc     As Long
   
'   Form1   窗口入口函数
Function WindowProc(ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'   拦截   WM_NOTIFYICON   消息
If msg = WM_NOTIFYICON Then
Select Case lParam
Case WM_RBUTTONUP
FormMain.PopupMenu FormMain.mnuTrayMenu
'   右键单击图标是运行这里的代码,   可以在这里添加弹出右键菜单的代码
Case WM_LBUTTONDBLCLK
FormMain.Show
On Error Resume Next
With FormMain
.Show
End With
'   删除托盘区图标
Dim IconData     As NOTIFYICONDATA
With IconData
.cbSize = Len(IconData)
.hwnd = FormMain.hwnd
.uID = 0
.uFlags = NIF_TIP Or NIF_ICON Or NIF_MESSAGE
.uCallbackMessage = WM_NOTIFYICON
.szTip = "Super Cursor - 雙擊顯示窗口，右擊顯示菜單"
.hIcon = FormMain.Icon.Handle
End With
Shell_NotifyIcon NIM_DELETE, IconData
SetWindowLong FormMain.hwnd, GWL_WNDPROC, preWndProc
'With FormMain.cSysTray1
'.InTray = False
'.TrayTip = "Super Cursor - 雙擊還原主窗口,右擊顯示菜單"
'End With
Case NIN_BALLOONSHOW
Debug.Print "显示气球提示"
Case NIN_BALLOONHIDE
Debug.Print "删除托盘图标"
Case NIN_BALLOONTIMEOUT
Debug.Print "气球提示消失"
Case NIN_BALLOONUSERCLICK
Debug.Print "单击气球提示"
FormMain.Show
On Error Resume Next
With FormMain
.Show
End With
'   删除托盘区图标
With IconData
.cbSize = Len(IconData)
.hwnd = FormMain.hwnd
.uID = 0
.uFlags = NIF_TIP Or NIF_ICON Or NIF_MESSAGE
.uCallbackMessage = WM_NOTIFYICON
.szTip = "Super Cursor - 雙擊顯示窗口，右擊顯示菜單"
.hIcon = FormMain.Icon.Handle
End With
Shell_NotifyIcon NIM_DELETE, IconData
SetWindowLong FormMain.hwnd, GWL_WNDPROC, preWndProc
'With FormMain.cSysTray1
'.InTray = False
'.TrayTip = "Super Cursor - 雙擊還原主窗口,右擊顯示菜單"
'End With
End Select
End If
WindowProc = CallWindowProc(preWndProc, FormMain.hwnd, msg, wParam, lParam)
End Function
