VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmRegister 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2355
   ClientLeft      =   2550
   ClientTop       =   8595
   ClientWidth     =   5865
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   HelpContextID   =   800
   Icon            =   "FRMREGIS.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2355
   ScaleWidth      =   5865
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox CheckResult 
      Alignment       =   1  'Right Justify
      Caption         =   "Result"
      Height          =   315
      Left            =   30
      TabIndex        =   6
      Top             =   2400
      Visible         =   0   'False
      Width           =   1965
   End
   Begin VB.TextBox txtRegCode 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   1080
      TabIndex        =   1
      Top             =   1380
      Width           =   2355
   End
   Begin VB.TextBox txtRegUser 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   1080
      TabIndex        =   0
      Top             =   990
      Width           =   4635
   End
   Begin Threed.SSCommand cmdOK 
      Default         =   -1  'True
      Height          =   330
      Left            =   3480
      TabIndex        =   2
      Top             =   1380
      Width           =   1095
      _Version        =   65536
      _ExtentX        =   1931
      _ExtentY        =   582
      _StockProps     =   78
      Caption         =   "&OK"
      ForeColor       =   -2147483630
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Threed.SSCommand cmdCancel 
      Cancel          =   -1  'True
      Height          =   330
      Left            =   4620
      TabIndex        =   3
      Top             =   1380
      Width           =   1095
      _Version        =   65536
      _ExtentX        =   1931
      _ExtentY        =   582
      _StockProps     =   78
      Caption         =   "C&ancel"
      ForeColor       =   -2147483630
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label labelHelpInfo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000D&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Press 'F1' for information on registering GSR and obtaining your user code."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   450
      Left            =   60
      TabIndex        =   9
      Top             =   1860
      Width           =   5730
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Code :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   3
      Left            =   60
      TabIndex        =   8
      Top             =   1470
      Width           =   975
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User name :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   1
      Left            =   60
      TabIndex        =   5
      Top             =   1050
      Width           =   990
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Please enter your registered user name and code below"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   255
      Index           =   0
      Left            =   60
      TabIndex        =   4
      Top             =   30
      Width           =   5715
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"FRMREGIS.frx":000C
      ForeColor       =   &H80000002&
      Height          =   615
      Index           =   2
      Left            =   60
      TabIndex        =   7
      Top             =   300
      Width           =   5715
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()

    '[SET RESULT OF BUTTON CHOICE]
    frmRegister.CheckResult = vbUnchecked
    '[CLOSE FORM / ONLY UNLOAD THOUGH, LET CALLING ROUTINE UNLOAD FORM]
    frmRegister.Hide

End Sub


Private Sub cmdCancel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Cancel registration and continue."
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub cmdOK_Click()

    '[SET RESULT OF BUTTON CHOICE]
    frmRegister.CheckResult = vbChecked
    '[CLOSE FORM / ONLY UNLOAD THOUGH, LET CALLING ROUTINE UNLOAD FORM]
    frmRegister.Hide

End Sub


Private Sub cmdOK_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Accept the displayed registration information and continue."
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub Form_Load()
    
    '[center form on screen]
    Top = (Screen.Height / 2) - (Height / 2)
    Left = (Screen.Width / 2) - (Width / 2)

End Sub


Private Sub txtRegCode_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Registration validation code."
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub txtRegUser_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Registered user name."
    '[---------------------------------------------------------------------------------]

End Sub


