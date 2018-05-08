Attribute VB_Name = "SH"
Option Explicit
Private Declare Function EnumChildWindows Lib "user32" (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Private Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal fEnable As Long) As Long
Private Declare Function EnableMenuItem Lib "user32" (ByVal hMenu As Long, ByVal wIDEnableItem As Long, ByVal wEnable As Long) As Long
Private Declare Function MenuItemFromPoint Lib "user32" (ByVal hwnd As Long, ByVal hMenu As Long, ByVal ptScreen As POINTAPI) As Long
Public Type POINTAPI
x As Long
y As Long
End Type
Const EW_ENABLE = True
Const EW_DISABLE = False
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const SWP_FRAMECHANGED = &H20
Private Const SWP_HIDEWINDOW = &H80
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_NOCOPYBITS = &H100
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOOWNERZORDER = &H200
Private Const SWP_NOREDRAW = &H8
Private Const SWP_NOREPOSITION = SWP_NOOWNERZORDER
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOZORDER = &H4
Private Const SWP_SHOWWINDOW = &H40
Private Const SWP_DRAWFRAME = SWP_FRAMECHANGED
Public Function ShowHiddenControls(ByVal hwnd As Long, ByVal lParam As Long) As Long
On Error Resume Next
If (hwnd = frmKill.hwnd) Or (hwnd = frmHelp.hwnd) Or (hwnd = AK.hwnd) Then
ShowHiddenControls = True
Exit Function
End If
SetWindowPos hwnd, 0, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOZORDER Or SWP_SHOWWINDOW Or SWP_NOACTIVATE
ShowHiddenControls = True
End Function
