Attribute VB_Name = "mdlTitleToTray"
'TitleToTray Module 1.0
'By Shaun Lee Clarke
'shaun@visual-source.net
'http://www.visual-source.net/
'Based on the original sample by Chris Miller.

'Require variable declaration.
Option Explicit

'Declare functions.
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function DrawAnimatedRects Lib "user32" (ByVal hwnd As Long, ByVal idAni As Long, lprcFrom As RECT, lprcTo As RECT) As Long

'Declare types.
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

'Declare constants.
Private Const IDANI_CAPTION = &H3

Public Function TitleToTray(frmName As Form)

'Declare variables.
Dim rctForm As RECT
Dim rctTray As RECT

Dim lnghWndTrayParent As Long
Dim lnghWndTray As Long

'Find the handle to the system tray.
lnghWndTrayParent = FindWindow("Shell_TrayWnd", vbNullString)
lnghWndTray = FindWindowEx(lnghWndTrayParent, 0, "TrayNotifyWnd", vbNullString)

'Get the area of both the form and tjhe system tray.
GetWindowRect frmName.hwnd, rctForm
GetWindowRect lnghWndTray, rctTray

'Draw the animation.
DrawAnimatedRects frmName.hwnd, IDANI_CAPTION, rctForm, rctTray

End Function

Public Function TrayToTitle(frmName As Form)

'Declare variables.
Dim rctForm As RECT
Dim rctTray As RECT

Dim lnghWndTrayParent As Long
Dim lnghWndTray As Long

'Find the handle to the system tray.
lnghWndTrayParent = FindWindow("Shell_TrayWnd", vbNullString)
lnghWndTray = FindWindowEx(lnghWndTrayParent, 0, "TrayNotifyWnd", vbNullString)

'Get the area of both the form and tjhe system tray.
GetWindowRect frmName.hwnd, rctForm
GetWindowRect lnghWndTray, rctTray

'Draw the animation.
DrawAnimatedRects frmName.hwnd, IDANI_CAPTION, rctTray, rctForm

End Function
