VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About GSR"
   ClientHeight    =   5280
   ClientLeft      =   5355
   ClientTop       =   4365
   ClientWidth     =   4755
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
   ForeColor       =   &H00808080&
   HelpContextID   =   600
   Icon            =   "FRMABOUT.frx":0000
   LinkTopic       =   "frmAbout"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5280
   ScaleWidth      =   4755
   Begin VB.TextBox TextWeb 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000012&
      Height          =   255
      Left            =   660
      Locked          =   -1  'True
      TabIndex        =   10
      Text            =   "Web Address"
      Top             =   4980
      Width           =   3270
   End
   Begin VB.TextBox TextEmail 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000012&
      Height          =   255
      Left            =   660
      Locked          =   -1  'True
      TabIndex        =   9
      Text            =   "Email Address"
      Top             =   4680
      Width           =   3270
   End
   Begin VB.Frame frameBorder 
      Height          =   1815
      Left            =   0
      TabIndex        =   4
      Top             =   -90
      Width           =   4755
      Begin VB.Image ImageLogo 
         Height          =   1575
         Left            =   3120
         Picture         =   "FRMABOUT.frx":0E42
         Top             =   180
         Width           =   1560
      End
      Begin VB.Label LabelVersion 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Version"
         Height          =   210
         Left            =   60
         TabIndex        =   8
         Top             =   1560
         Width           =   570
      End
      Begin VB.Label lblTitle_1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Generic"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   60
         TabIndex        =   7
         Top             =   240
         Width           =   3000
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(C) 1998, David Gilbert"
         Height          =   210
         Left            =   1440
         TabIndex        =   6
         Top             =   1560
         Width           =   1635
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Staff Roster"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   495
         Left            =   60
         TabIndex        =   5
         Top             =   900
         Width           =   3000
      End
   End
   Begin Threed.SSCommand cmdReturn 
      Cancel          =   -1  'True
      Default         =   -1  'True
      Height          =   360
      Left            =   4020
      TabIndex        =   0
      Top             =   4860
      Width           =   705
      _Version        =   65536
      _ExtentX        =   1244
      _ExtentY        =   635
      _StockProps     =   78
      Caption         =   "C&lose"
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
      AutoSize        =   2
   End
   Begin VB.Image ImageWeb 
      Height          =   480
      Left            =   30
      Picture         =   "FRMABOUT.frx":3D2C
      Top             =   3060
      Width           =   480
   End
   Begin VB.Label Label_std 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " This copy of GSR is licensed to :"
      ForeColor       =   &H80000009&
      Height          =   255
      Index           =   1
      Left            =   30
      TabIndex        =   3
      Top             =   1740
      Width           =   4700
   End
   Begin VB.Image imageInfo 
      Height          =   480
      Left            =   30
      Picture         =   "FRMABOUT.frx":45F6
      Tag             =   "0"
      Top             =   2460
      Width           =   480
   End
   Begin VB.Label LabelInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Generic Staff Roster,  (C) 1997 David Gilbert."
      Height          =   2175
      Left            =   660
      TabIndex        =   2
      Tag             =   "0"
      Top             =   2460
      Width           =   4035
      WordWrap        =   -1  'True
   End
   Begin VB.Label labelRegUser 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Registered Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   435
      Left            =   30
      TabIndex        =   1
      Top             =   1980
      Width           =   4695
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'[----------------------------------------------]
'[frmAbout          About program and author    ]
'[----------------------------------------------]
'[GSR                                           ]
'[generic staff roster system for Windows       ]
'[(C) ---------- David Gilbert, 1997            ]
'[----------------------------------------------]
'[version                           2.1         ]
'[initial date                      21/07/1997  ]
'[----------------------------------------------]



Private Sub cmdReturn_Click()

    '[CLOSE FORM]
    Unload frmAbout
    mdiMain.ZOrder

End Sub

Private Sub cmdReturn_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Close this window and return to the roster form."
    '[---------------------------------------------------------------------------------]

End Sub



Private Sub Form_Load()

    '[center form on screen]
    Top = (Screen.Height / 2) - (Height / 2)
    Left = (Screen.Width / 2) - (Width / 2)
        
End Sub



Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Information about Generic Staff Roster."
    '[---------------------------------------------------------------------------------]

End Sub





Private Sub ImageInfo_Click()

    '[ANIMATE BUTTON]
    ImageInfo.BorderStyle = 1
    Delay vbDelay
    ImageInfo.BorderStyle = 0

    Select Case ImageInfo.Tag
    Case 0
        labelInfo.Caption = "Generic Staff Roster was designed as a simple alternative to pencil and paper staff roster creation, allowing the maintainance of ten seperate weekly rosters.  GSR incorporates several roster failsafes, such as duplicate shift allocations."
        labelInfo.Caption = labelInfo.Caption & strBreak & strBreak & "The concept was spawned from various comments about the difficulty of creating multiple weekly rosters on paper."
        ImageInfo.Tag = 1
    Case 1
        labelInfo.Caption = "David Gilbert lives in Toowoomba, Australia and works as a freelance programmer.  He is " & Format(Date - CDate("25/12/1967"), "yy") & " years old and spends most of his time throwing a tennis ball for his two dogs."
        labelInfo.Caption = labelInfo.Caption & strBreak & strBreak & "GSR is his first non-commissioned software release and was written during a frenzied two month period in late 1997 and updated in April 1998."
        ImageInfo.Tag = 2
    Case 2
        frmAbout.labelInfo = "You have been using Generic Staff Roster for " & sinDaysUsed & " days."
        ImageInfo.Tag = 3
    Case Else
        frmAbout.labelInfo.Caption = "A Generic Staff Rostering software system for small businesses."
        frmAbout.labelInfo.Caption = frmAbout.labelInfo.Caption & strBreak & strBreak & "Generic Staff Roster was designed to assist small business owners/operators in allocating their available human resources easily and efficiently."
        If DsDefault!RegCode = "" Or IsNull(DsDefault!RegCode) Then frmAbout.labelInfo.Caption = frmAbout.labelInfo.Caption & strBreak & strBreak & "For more information on registering Generic Staff Roster, see the help file section entitled 'Registering GSR'."
        ImageInfo.Tag = 0
    End Select

End Sub

Private Sub ImageInfo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Click here for even more information .."
    '[---------------------------------------------------------------------------------]

End Sub





Private Sub ImageWeb_Click()

    '[ANIMATE BUTTON]
    ImageWeb.BorderStyle = 1
    Delay vbDelay
    ImageWeb.BorderStyle = 0

    '[SHOW STARTUP LOG]
    Dim Result
    Result = Shell("NOTEPAD.EXE startup.log", 1)

End Sub

Private Sub ImageWeb_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Click here to display the startup log."
    '[---------------------------------------------------------------------------------]

End Sub



Private Sub labelInfo_Click()

    Unload frmAbout

End Sub


Private Sub labelRegUser_Click()

    Unload frmAbout

End Sub





