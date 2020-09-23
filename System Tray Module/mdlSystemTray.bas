Attribute VB_Name = "mdlSystemTray"
' System Tray Module 1.0
' By Shaun Lee Clarke
' shaun@visual-source.net
' http://www.visual-source.net/
' Based on the original sample by (unknown).

' Usage:
' To add an icon:
'     AddIcon picSystemTray.hwnd, 1, picSystemTrayIcon.Picture.Handle, "Tooltip"
'             Handler Handle      ID Handle to Icon                    Tooltip
'
' To modify an icon:
'     ModifyIcon picSystemTray.hwnd, 1, picSystemTrayIcon2.Picture.Handle, "Standby"
'                Handler Handle      ID Handle to Icon                     Tooltip
'
' To remove an icon:
'     RemoveIcon picSystemTray.hwnd, 1
'                Handler Handle      ID
'
' Note: To use multiple icons, simple use a different ID number for each icon.
'       For example, if you want to add a second icon, simply set the ID as a
'       different number as the original such as 2. An example is shown below:
'
'           AddIcon picSystemTray.hwnd, 1, picSystemTrayIcon1.Picture.Handle, "Icon 1"
'           AddIcon picSystemTray.hwnd, 2, picSystemTrayIcon2.Picture.Handle, "Icon 2"
'
'       This code would add two icons to the tray, simple use the ID number if
'       you wish to modify or delete the icon.
'
' If you do not delete the Icon when your program closes, you will find that the
' icon will remain in the System Tray unless it receives an event, such as
' moving the mouse over it. To try this, run the sample application and stop it
' with the IDE's stop button.

'Require variable declaration.
Option Explicit

'Declare functions.
Public Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long

'Declare types.
Public Type NOTIFYICONDATA
        cbSize As Long
        hwnd As Long
        uID As Long
        uFlags As Long
        uCallbackMessage As Long
        hIcon As Long
        szTip As String * 64
End Type

'Declare constants.
Public Const NIF_MESSAGE = &H1
Public Const NIF_TIP = &H4
Public Const NIF_ICON = &H2

Public Const NIM_ADD = &H0
Public Const NIM_DELETE = &H2
Public Const NIM_MODIFY = &H1

Public Const WM_MOUSEMOVE = &H200

Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202

Public Const WM_MBUTTONDBLCLK = &H209
Public Const WM_MBUTTONDOWN = &H207
Public Const WM_MBUTTONUP = &H208

Public Const WM_RBUTTONDBLCLK = &H206
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205

Public Function AddIcon(lngParenthWnd As Long, lnguID As Long, lngIcon As Long, strToolTip As String)

'Declare objects.
Dim IconData As NOTIFYICONDATA

IconData.cbSize = Len(IconData)

'Construct the information.
IconData.hwnd = lngParenthWnd
IconData.uID = lnguID
IconData.uFlags = NIF_MESSAGE Or NIF_ICON Or NIF_TIP
IconData.uCallbackMessage = WM_MOUSEMOVE
IconData.hIcon = lngIcon
IconData.szTip = strToolTip & vbNullChar

'Call the funciton to add the icon.
Shell_NotifyIcon NIM_ADD, IconData

End Function

Public Function RemoveIcon(lngParenthWnd As Long, lnguID As Long)

'Declare objects.
Dim IconData As NOTIFYICONDATA

IconData.cbSize = Len(IconData)

'Construct the information.
IconData.hwnd = lngParenthWnd
IconData.uID = lnguID
'IconData.uFlags = NIF_MESSAGE Or NIF_ICON Or NIF_TIP
'IconData.uCallbackMessage = WM_MOUSEMOVE
'IconData.hIcon = lngIcon
'IconData.szTip = strToolTip & vbNullChar

'Call the funciton to remove the icon.
Shell_NotifyIcon NIM_DELETE, IconData

End Function

Public Function ModifyIcon(lngParenthWnd As Long, lnguID As Long, lngIcon As Long, strToolTip As String)

'Declare objects.
Dim IconData As NOTIFYICONDATA

IconData.cbSize = Len(IconData)

'Construct the information.
IconData.hwnd = lngParenthWnd
IconData.uID = lnguID
IconData.uFlags = NIF_MESSAGE Or NIF_ICON Or NIF_TIP
IconData.uCallbackMessage = WM_MOUSEMOVE
IconData.hIcon = lngIcon
IconData.szTip = strToolTip & vbNullChar

'Call the funciton to modify the icon.
Shell_NotifyIcon NIM_MODIFY, IconData

End Function
