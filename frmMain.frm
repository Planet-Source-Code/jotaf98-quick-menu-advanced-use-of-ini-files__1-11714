VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Quick Menu Options"
   ClientHeight    =   4590
   ClientLeft      =   150
   ClientTop       =   150
   ClientWidth     =   5670
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   306
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   378
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdDown 
      Caption         =   "u"
      BeginProperty Font 
         Name            =   "Marlett"
         Size            =   11.25
         Charset         =   2
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   15
      ToolTipText     =   "Move Down Shortcut"
      Top             =   2880
      Width           =   255
   End
   Begin VB.TextBox txtIconPath 
      Height          =   285
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   13
      Text            =   "Current Icon: (Default)"
      Top             =   240
      Width           =   5175
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "-"
      Height          =   255
      Left            =   240
      TabIndex        =   12
      ToolTipText     =   "Delete Shortcut"
      Top             =   2640
      Width           =   255
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "+"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      ToolTipText     =   "New Shortcut"
      Top             =   2400
      Width           =   255
   End
   Begin VB.ComboBox cmbModify 
      Height          =   315
      ItemData        =   "frmMain.frx":0E42
      Left            =   4200
      List            =   "frmMain.frx":0E4C
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "About..."
      Height          =   375
      Left            =   360
      TabIndex        =   8
      Top             =   3960
      Width           =   1335
   End
   Begin VB.PictureBox picPreview 
      AutoSize        =   -1  'True
      BackColor       =   &H80000010&
      BorderStyle     =   0  'None
      Height          =   450
      Left            =   240
      Picture         =   "frmMain.frx":0E5F
      ScaleHeight     =   450
      ScaleWidth      =   5160
      TabIndex        =   7
      Top             =   1320
      Width           =   5160
      Begin VB.Image picIcon 
         Height          =   255
         Left            =   4230
         Picture         =   "frmMain.frx":8791
         Stretch         =   -1  'True
         Top             =   120
         Width           =   255
      End
   End
   Begin VB.TextBox txtShortcutProp 
      Height          =   285
      Left            =   240
      TabIndex        =   6
      Top             =   3240
      Width           =   3855
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   3960
      TabIndex        =   5
      Top             =   3960
      Width           =   1335
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Height          =   375
      Left            =   2160
      TabIndex        =   4
      Top             =   3960
      Width           =   1335
   End
   Begin VB.ListBox lstShortcuts 
      Height          =   1035
      ItemData        =   "frmMain.frx":88DB
      Left            =   600
      List            =   "frmMain.frx":88DD
      TabIndex        =   3
      Top             =   2160
      Width           =   4815
   End
   Begin VB.CommandButton cmdDefaultIcon 
      Caption         =   "Use Default"
      Height          =   375
      Left            =   4200
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin VB.PictureBox picDefaultIcon 
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   2400
      Picture         =   "frmMain.frx":88DF
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   1
      Top             =   600
      Visible         =   0   'False
      Width           =   240
   End
   Begin MSComDlg.CommonDialog dlgIcon 
      Left            =   2880
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Open New Icon"
      Filter          =   "Icon files (*.ico;*.cur)|*.ico;*.cur|Image files (*.bmp;*.gif;*.jpg)|*.bmp;*.gif;*.jpg"
   End
   Begin VB.CommandButton cmdBrowseIcon 
      Caption         =   "Browse..."
      Height          =   375
      Left            =   2880
      TabIndex        =   0
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton cmdUp 
      Caption         =   "t"
      BeginProperty Font 
         Name            =   "Marlett"
         Size            =   11.25
         Charset         =   2
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   14
      ToolTipText     =   "Move Up Shortcut"
      Top             =   2160
      Width           =   255
   End
   Begin VB.Label lblPreview 
      Caption         =   "Preview:"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   1080
      Width           =   735
   End
   Begin VB.Line lneSeparator1 
      BorderColor     =   &H80000016&
      Index           =   1
      X1              =   16
      X2              =   360
      Y1              =   129
      Y2              =   129
   End
   Begin VB.Line lneSeparator1 
      BorderColor     =   &H80000010&
      Index           =   0
      X1              =   16
      X2              =   360
      Y1              =   128
      Y2              =   128
   End
   Begin VB.Line lneSeparator2 
      BorderColor     =   &H80000010&
      Index           =   0
      X1              =   16
      X2              =   360
      Y1              =   248
      Y2              =   248
   End
   Begin VB.Line lneSeparator2 
      BorderColor     =   &H80000016&
      Index           =   1
      X1              =   16
      X2              =   360
      Y1              =   249
      Y2              =   249
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "Popup"
      Visible         =   0   'False
      Begin VB.Menu mnuQuickMenu 
         Caption         =   "Quick Menu"
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSeparator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOptionsMain 
         Caption         =   "Options"
         Begin VB.Menu mnuOptions 
            Caption         =   "Options..."
         End
         Begin VB.Menu mnuSeparator2 
            Caption         =   "-"
         End
         Begin VB.Menu mnuAbout 
            Caption         =   "About..."
         End
         Begin VB.Menu mnuSeparator3 
            Caption         =   "-"
         End
         Begin VB.Menu mnuClose 
            Caption         =   "Close"
         End
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'This API will execute a command like if in "Start -> Execute..."
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const SW_SHOWNORMAL = 1

'The current icon's path
Dim IconPath As String

'This type keeps a shortcut's name and its executed command
Private Type Shortcut
    Name As String
    Command As String
End Type

'Required to delete unused shortcuts
Dim InitialShortcutsUBound As Integer

'All shortcuts will be stored in this array
Dim Shortcuts() As Shortcut

'Counter
Dim i As Integer

Private Sub Form_Load()
    'Set the icon path according to the INI file
    IconPath = ReadINI("Options", "IconPath", App.Path & "\QuickMenu.ini")
    txtIconPath.Text = "Current Icon: " & IconPath
    
    'User didn't specify a drive: assume it's a
    'file in this application's folder
    If Mid(IconPath, 1, 1) <> "\" And IconPath <> "(Default)" Then
        IconPath = App.Path & "\" & IconPath
    End If
    
    'Load the picture only if it exists
    If Dir(IconPath) <> "" Then picIcon.Picture = LoadPicture(dlgIcon.FileName)
    
    
    ' - Load the shortcuts -
    
    Dim ExitLoop As Boolean 'Will be True when loading is done
    Dim TempName As String 'Temp name (shortcut property)
    Dim TempCommand As String 'Temp command (shortcut property)
    
    
    'This loop will try all possible numbers starting at 1, and
    'try to load that shortcut from the INI file. It will stop
    'when it doesn't find a shortcut (so the previous one was
    'the last). Have a look at QuickMenu.ini to see what I mean!
    Do
        'Increase counter
        i = i + 1
        
        'Attempt to read this shortcut
        TempName = ReadINI("Shortcuts", Str(i) & "Name", App.Path & "\QuickMenu.ini")
        TempCommand = ReadINI("Shortcuts", Str(i) & "Command", App.Path & "\QuickMenu.ini")
        
        If TempName = "" Or TempCommand = "" Then
            'Failed reading this shortcut, so it doesn't exist
            '(we've reached the last one) - exit this loop
            ExitLoop = True
        Else 'Shortcut exists - add it to the array
            
            'Redim array to hold this one too
            ReDim Preserve Shortcuts(1 To i)
            
            'Load the shortcut's name/command into array
            Shortcuts(i).Name = TempName
            Shortcuts(i).Command = TempCommand
        End If
    Loop Until ExitLoop
    
    'Set "InitialShortcutsUBound" (so we know the shortcuts
    'we'll have to delete)
    InitialShortcutsUBound = i - 1
    
    'Fill the list
    FillListBox
    
    'Select the first item in the combo
    cmbModify.ListIndex = 0
    
    'If Quick Menu was launched with " -NoOptions" in the
    'command line, hide and display only the menu
    If LCase(Trim(Command)) = "-nooptions" Then
        cmdOk_Click
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Remove the tray icon
    RemoveTrayIcon
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (X + Y ^ 16) - 1 = WM_RBUTTONDOWN Or (X + Y ^ 16) - 1 = WM_LBUTTONDOWN Then
        'The menu that will pop up when the user
        'right-clicks the Tray Icon
        frmMain.PopupMenu mnuPopup
    End If
End Sub

Private Sub lstShortcuts_Click()
    If cmbModify.Text = "Name" Then
        'Name selected - display it in the text box
        txtShortcutProp.Text = Shortcuts(lstShortcuts.ListIndex + 1).Name
    Else
        'Command selected - display it in the text box
        txtShortcutProp.Text = Shortcuts(lstShortcuts.ListIndex + 1).Command
    End If
End Sub

Private Sub cmbModify_Click()
    If cmbModify.Text = "Name" Then
        'The user chose Name - display it
        txtShortcutProp.Text = Shortcuts(lstShortcuts.ListIndex + 1).Name
    Else
        'The user chose Command - display it
        txtShortcutProp.Text = Shortcuts(lstShortcuts.ListIndex + 1).Command
    End If
End Sub

'Choose a new icon
Private Sub cmdBrowseIcon_Click()
    On Error GoTo UserCanceled
    
    'Set flags and show dialog
    dlgIcon.Flags = cdlOFNFileMustExist And cdlOFNNoChangeDir
    dlgIcon.ShowOpen
    
    'Load the icon
    picIcon.Picture = LoadPicture(dlgIcon.FileName)
    
    'Set the icon path
    IconPath = dlgIcon.FileName
    txtIconPath.Text = "Current Icon: " & IconPath
    
UserCanceled:
End Sub

Private Sub cmdDefaultIcon_Click()
    'Set icon path to (Default)
    txtIconPath.Text = "Current Icon: " & "(Default)"
    IconPath = "(Default)"
    
    'Load the default picture
    picIcon.Picture = picDefaultIcon.Picture
End Sub

Private Sub txtShortcutProp_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Or KeyAscii = 10 Then 'User pressed "Enter"
        'Setting KeyAscii to 0 here means that no
        'character will be inserted
        KeyAscii = 0
        
        If cmbModify.Text = "Name" Then
            'Name selected - set this shortcut's name
            Shortcuts(lstShortcuts.ListIndex + 1).Name = txtShortcutProp.Text
            
            '...and reset the listbox
            FillListBox
        Else
            'Command selected - set this shortcut's command
            Shortcuts(lstShortcuts.ListIndex + 1).Command = txtShortcutProp.Text
        End If
    End If
End Sub

Private Sub cmdNew_Click()
    Dim NewShortcutIndex As Integer 'The index of the new shortcut
    
    'Set the new shortcut's index
    NewShortcutIndex = UBound(Shortcuts) + 1
    
    'Add an item to the array
    ReDim Preserve Shortcuts(1 To NewShortcutIndex)
    
    'Set its properties
    Shortcuts(NewShortcutIndex).Name = "(New)"
    Shortcuts(NewShortcutIndex).Command = "(New Command)"
    
    'Refresh the list box
    FillListBox
End Sub

Private Sub cmdDelete_Click()
    'Prompt the user to delete the shortcut
    If MsgBox("Are you sure you want to delete the shortcut '" & Shortcuts(lstShortcuts.ListIndex + 1).Name & "'?", vbYesNo Or vbQuestion, "Delete Shortcut") = vbYes Then
        'Move trough all shortcuts after the one that we'll
        'delete and move them one position down
        For i = lstShortcuts.ListIndex + 2 To UBound(Shortcuts)
            Shortcuts(i - 1).Name = Shortcuts(i).Name
            Shortcuts(i - 1).Command = Shortcuts(i).Command
        Next i
        
        'Delete the last shortcut
        ReDim Preserve Shortcuts(1 To UBound(Shortcuts) - 1)
        
        'Refresh the list box
        FillListBox
    End If
End Sub

Private Sub cmdUp_Click()
    'If the selected shortcut can't be moved up,
    'don't do it
    If lstShortcuts.ListIndex = 0 Then Exit Sub
    
    'Selected shortcut
    Dim Temp1 As Shortcut
    'Shortcut directly above it
    Dim Temp2 As Shortcut
    
    'Set them
    Temp1 = Shortcuts(lstShortcuts.ListIndex + 1)
    Temp2 = Shortcuts(lstShortcuts.ListIndex)
    
    'Now, switch them
    Shortcuts(lstShortcuts.ListIndex) = Temp1
    Shortcuts(lstShortcuts.ListIndex + 1) = Temp2
    
    'Select it
    lstShortcuts.ListIndex = lstShortcuts.ListIndex - 1
    
    'Refresh the list box
    FillListBox
End Sub

Private Sub cmdDown_Click()
    'If the selected shortcut can't be moved down,
    'don't do it
    If lstShortcuts.ListIndex = UBound(Shortcuts) - 1 Then Exit Sub
    
    'Selected shortcut
    Dim Temp1 As Shortcut
    'Shortcut directly below it
    Dim Temp2 As Shortcut
    
    'Set them
    Temp1 = Shortcuts(lstShortcuts.ListIndex + 1)
    Temp2 = Shortcuts(lstShortcuts.ListIndex + 2)
    
    'Now, switch them
    Shortcuts(lstShortcuts.ListIndex + 2) = Temp1
    Shortcuts(lstShortcuts.ListIndex + 1) = Temp2
    
    'Select it
    lstShortcuts.ListIndex = lstShortcuts.ListIndex + 1
    
    'Refresh the list box
    FillListBox
End Sub

Private Sub cmdAbout_Click()
    'Show the About form
    frmAbout.Show vbModal, Me
End Sub

Private Sub cmdOk_Click()
    ' - Load shortcuts into the menu -
    
    On Error Resume Next
    
    'Loop trough all shortcuts
    For i = 1 To UBound(Shortcuts)
        'Load this menu
        Load mnuQuickMenu(i)
        
        'Set its properties
        mnuQuickMenu(i).Caption = Shortcuts(i).Name
        mnuQuickMenu(i).Visible = True
        
        'Save it to the INI file
        WriteINI "Shortcuts", Str(i) & "Name", Shortcuts(i).Name, App.Path & "\QuickMenu.ini"
        WriteINI "Shortcuts", Str(i) & "Command", Shortcuts(i).Command, App.Path & "\QuickMenu.ini"
    Next i
    
    'This will delete unused shortcuts (by setting
    'their values to a null string)
    If InitialShortcutsUBound > UBound(Shortcuts) Then
        For i = UBound(Shortcuts) + 1 To InitialShortcutsUBound
            WriteINI "Shortcuts", Str(i) & "Name", vbNullString, App.Path & "\QuickMenu.ini"
            WriteINI "Shortcuts", Str(i) & "Command", vbNullString, App.Path & "\QuickMenu.ini"
        Next i
    End If
    
    'Create the tray icon and hide this form
    CreateTrayIcon Me.hWnd, picIcon.Picture, "Quick Menu"
    Me.Hide
End Sub

Private Sub cmdExit_Click()
    'Prompt the user to save before exiting
    If MsgBox("Save changes to the Shortcuts before exiting?", vbYesNo Or vbQuestion, "Quick Menu") = vbYes Then
        'Loop trough all shortcuts
        For i = 1 To UBound(Shortcuts)
            'Save it to the INI file
            WriteINI "Shortcuts", Str(i) & "Name", Shortcuts(i).Name, App.Path & "\QuickMenu.ini"
            WriteINI "Shortcuts", Str(i) & "Command", Shortcuts(i).Command, App.Path & "\QuickMenu.ini"
        Next i
        
        'This will delete unused shortcuts (by setting
        'their values to a null string)
        If InitialShortcutsUBound > UBound(Shortcuts) Then
            For i = UBound(Shortcuts) + 1 To InitialShortcutsUBound
                WriteINI "Shortcuts", Str(i) & "Name", vbNullString, App.Path & "\QuickMenu.ini"
                WriteINI "Shortcuts", Str(i) & "Command", vbNullString, App.Path & "\QuickMenu.ini"
            Next i
        End If
    End If
    
    'Close
    Unload Me
End Sub

Private Sub mnuQuickMenu_Click(Index As Integer)
    'Execute this shortcut's command
    ExecuteShortcut Index
End Sub

Private Sub mnuOptions_Click()
    'Show this form (the options dialog) and
    'remove the tray icon
    Me.Show
    RemoveTrayIcon
End Sub

Private Sub mnuAbout_Click()
    'Show the About form
    frmAbout.Show vbModal, Me
End Sub

Private Sub mnuClose_Click()
    'Close
    Unload Me
End Sub



'This will fill the list box with all shortcuts in the array
Private Sub FillListBox()
    On Error GoTo SelectLastItem
    'The item that was selected before clearing the list
    Dim LastListIndex As Integer
    
    'If an item is selected, set the last selected list item
    If lstShortcuts.ListIndex <> -1 Then LastListIndex = lstShortcuts.ListIndex
    
    'Clear the list
    lstShortcuts.Clear
    
    'Loop trough all shortcuts, and add them to the list
    For i = 1 To UBound(Shortcuts)
        lstShortcuts.AddItem Shortcuts(i).Name
    Next i
    
    'Select the item that was previously selected
    lstShortcuts.ListIndex = LastListIndex
    
    Exit Sub
    
SelectLastItem:
    lstShortcuts.ListIndex = UBound(Shortcuts) - 1
End Sub

'This will execute a command (shortcut)
Public Sub ExecuteShortcut(ShortcutIndex As Integer)
    Dim TempStr As String 'Temorary string
    
    'If the command starts with "QM:", add "[QuickMenu's
    'path]\Shortcuts" to the start of the command
    If Left(Shortcuts(ShortcutIndex).Command, 3) = "QM:" Then
        TempStr = App.Path & "\Shortcuts" & Mid(Shortcuts(ShortcutIndex).Command, 4)
    Else 'Else, execute it as it is
        TempStr = Shortcuts(ShortcutIndex).Command
    End If
    
    'Execute the command "TempStr"
    ShellExecute 0&, vbNullString, TempStr, vbNullString, "C:\", SW_SHOWNORMAL
End Sub
