Attribute VB_Name = "KG"
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
Public Function EnableDisabledControls(ByVal hwnd As Long, ByVal lParam As Long) As Long
On Error Resume Next
If (hwnd = frmKill.hwnd) Or (hwnd = frmHelp.hwnd) Or (hwnd = AK.hwnd) Then
EnableDisabledControls = True
Exit Function
End If
EnableWindow hwnd, EW_ENABLE
EnableDisabledControls = True
End Function
