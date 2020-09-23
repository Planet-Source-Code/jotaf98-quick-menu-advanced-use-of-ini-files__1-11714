Attribute VB_Name = "mTrayIcon"

'
'Tray Icon Demo Project by Fox (Fox_McCloud@gmx.net)
'

Option Explicit

'Consts
    Public Const ICON_MESSAGE = 1
    Public Const ICON_ICON = 2
    Public Const ICON_TIP = 4
    
    Public Const ADD_ICON = 0
    Public Const MODIFY_ICON = 1
    Public Const DELETE_ICON = 2
    
    Public Const WM_LBUTTONDOWN = &H201
    Public Const WM_RBUTTONDOWN = &H204

'API
    Public Declare Function Shell_NotifyIcon Lib "shell32.dll" (ByVal dwMessage As Long, IpData As NOTIFYICONDATA) As Long

'Types
    Type NOTIFYICONDATA
        cbSize As Long
        hWnd As Long
        uID As Long
        uFlags As Long
        uCallbackMessage As Long
        hIcon As Long
        szTip As String * 128
    End Type

'Vars
    Private IconData As NOTIFYICONDATA

Public Sub CreateTrayIcon(ihWnd As Long, iPicture As StdPicture, iToolTipText As String)
    Dim Temp As Long
    
    'Set icon data
    IconData.uID = vbNull
    IconData.cbSize = Len(IconData)
    IconData.uFlags = ICON_MESSAGE Or ICON_ICON Or ICON_TIP
    
    'Important
    IconData.hWnd = ihWnd 'The form which handles the clicks
    IconData.uCallbackMessage = WM_RBUTTONDOWN
    
    'Now the visual part
    IconData.szTip = iToolTipText & vbNullChar 'The ToolTipText
    IconData.hIcon = iPicture 'The icon
    
    'Delete and create new icon
    Temp = Shell_NotifyIcon(DELETE_ICON, IconData)
    Temp = Shell_NotifyIcon(ADD_ICON, IconData)
End Sub


Public Sub RemoveTrayIcon()
    Dim Temp As Long

    'Remove tray icon
    Temp = Shell_NotifyIcon(DELETE_ICON, IconData)
End Sub


