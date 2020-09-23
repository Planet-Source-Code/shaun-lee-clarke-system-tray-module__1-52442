VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "System Tray Sample Application"
   ClientHeight    =   2295
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   5175
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   5175
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picSystemTrayIcon2 
      AutoSize        =   -1  'True
      Height          =   300
      Left            =   960
      Picture         =   "frmMain.frx":27A2
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   5
      Top             =   600
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.CommandButton cmdAddIcon 
      Caption         =   "&Add Icon"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdRemoveIcon 
      Caption         =   "&Remove Icon"
      Height          =   375
      Left            =   2760
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdModifyIcon 
      Caption         =   "&Modify Icon"
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.PictureBox picSystemTray 
      Height          =   375
      Left            =   120
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   1
      Top             =   600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picSystemTrayIcon1 
      AutoSize        =   -1  'True
      Height          =   300
      Left            =   600
      Picture         =   "frmMain.frx":2E64
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   0
      Top             =   600
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Menu mnuPopupMenu 
      Caption         =   "Popup Menu"
      Visible         =   0   'False
      Begin VB.Menu mnuPopupMenuShowForm 
         Caption         =   "&Show/Hide Form"
      End
      Begin VB.Menu mnuPopupMenuSeperator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopupMenuExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' System Tray Sample Application 1.0
' By Shaun Lee Clarke
' shaun@visual-source.net
' http://www.visual-source.net/

'Require variable declaration.
Option Explicit

Private Sub cmdAddIcon_Click()

'Add the system tray icon.
AddIcon picSystemTray.hwnd, 1, picSystemTrayIcon1.Picture.Handle, "System Tray Sample Application"

End Sub

Private Sub cmdModifyIcon_Click()

'Modify the system tray icon.
ModifyIcon picSystemTray.hwnd, 1, picSystemTrayIcon2.Picture.Handle, "New Icon and Tooltip!"

End Sub

Private Sub cmdRemoveIcon_Click()

'Remove the system tray icon.
RemoveIcon picSystemTray.hwnd, 1

End Sub
Private Sub Form_Load()

'Hide the application.
frmMain.Hide

'Add the system tray icon.
AddIcon picSystemTray.hwnd, 1, picSystemTrayIcon1.Picture.Handle, "System Tray Sample Application"

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

'This code makes the form minimise back to the system tray when the 'X' button
'is clicked. To quit the program, select 'Exit' from the context menu of the
'icon. This code can be easily removed if you do not want this feature, it's
'just here as an example.

'Abort form closing.
Cancel = 1

'Display Title To Tray animation.
TitleToTray frmMain

1 'Hide the form.
frmMain.Hide

End Sub

Private Sub mnuPopupMenuExit_Click()

'Remove the System Tray icon.
RemoveIcon picSystemTray.hwnd, 1

'Close the program.
End

End Sub

Private Sub mnuPopupMenuShowForm_Click()

'Display Tray To Title animation if the form is not visible.
If frmMain.Visible = False Then
    TrayToTitle frmMain
End If

'Show the form.
frmMain.Show

End Sub

Private Sub picSystemTray_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

x = x / Screen.TwipsPerPixelX

Select Case x
    Case WM_LBUTTONDOWN
        'lblLastEvent.Caption = "Left button down"
    Case WM_LBUTTONUP
        'lblLastEvent.Caption = "Left button up"
    Case WM_LBUTTONDBLCLK
        If frmMain.Visible = False Then
            TrayToTitle frmMain
            frmMain.Show
        ElseIf frmMain.Visible = True Then
            TitleToTray frmMain
            frmMain.Hide
        End If
        'lblLastEvent.Caption = "Left button double click"
    Case WM_RBUTTONDOWN
        'lblLastEvent.Caption = "Right button down"
    Case WM_RBUTTONUP
        'lblLastEvent.Caption = "Right button up"
        Me.PopupMenu mnuPopupMenu, , , , mnuPopupMenuShowForm
    Case WM_RBUTTONDBLCLK
        'lblLastEvent.Caption = "Right button double click"
    Case WM_MBUTTONDOWN
        'lblLastEvent.Caption = "Middle button down"
    Case WM_MBUTTONUP
        'lblLastEvent.Caption = "Middle button up"
    Case WM_MBUTTONDBLCLK
        'lblLastEvent.Caption = "Middle button double click"
End Select

End Sub
