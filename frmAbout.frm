VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "About Quick Menu"
   ClientHeight    =   2070
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3270
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2070
   ScaleWidth      =   3270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtEMail 
      BackColor       =   &H8000000D&
      ForeColor       =   &H8000000E&
      Height          =   285
      Left            =   1440
      TabIndex        =   4
      Text            =   "jotaf98@hotmail.com"
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label lblEMail 
      Alignment       =   1  'Right Justify
      Caption         =   "E-mail:"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label lblName 
      Caption         =   "(João F. S. Henriques)"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   2775
   End
   Begin VB.Label lblNick 
      Caption         =   "Coded by Jotaf98"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   2775
   End
   Begin VB.Label lblQuickMenu 
      Caption         =   "QuickMenu"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   0
      Top             =   240
      Width           =   1935
   End
   Begin VB.Line lneSeparator1 
      BorderColor     =   &H80000010&
      Index           =   0
      X1              =   240
      X2              =   3000
      Y1              =   735
      Y2              =   735
   End
   Begin VB.Line lneSeparator1 
      BorderColor     =   &H80000016&
      Index           =   1
      X1              =   240
      X2              =   3000
      Y1              =   750
      Y2              =   750
   End
   Begin VB.Image imgLogo 
      Height          =   480
      Left            =   240
      Picture         =   "frmAbout.frx":0000
      Top             =   240
      Width           =   480
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "Popup"
      Visible         =   0   'False
      Begin VB.Menu mnuSendEMail 
         Caption         =   "Send e-mail now?"
      End
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'To send an e-mail...
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const SW_SHOWNORMAL = 1

Private Sub txtEMail_Click()
    txtEMail.SelStart = 0
    txtEMail.SelLength = Len(txtEMail.Text)
    
    PopupMenu mnuPopup
End Sub

Private Sub mnuSendEMail_Click()
    'Send an e-mail to me
    ShellExecute 0&, vbNullString, "mailto:jotaf98@hotmail.com", vbNullString, "C:\", SW_SHOWNORMAL
End Sub
