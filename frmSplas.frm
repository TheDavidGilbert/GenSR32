VERSION 5.00
Begin VB.Form frmSplash 
   ClientHeight    =   2280
   ClientLeft      =   6060
   ClientTop       =   5235
   ClientWidth     =   6030
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillColor       =   &H00E0E0E0&
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FRMSPLAS.frx":0000
   LinkTopic       =   "frmSplash"
   NegotiateMenus  =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2280
   ScaleWidth      =   6030
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.Label lblVerInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "Generic Staff Roster Version"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   60
      TabIndex        =   0
      Top             =   1920
      Width           =   5925
   End
   Begin VB.Shape gaugeBar 
      BackColor       =   &H00008080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00400040&
      FillColor       =   &H00E0E0E0&
      Height          =   120
      Left            =   30
      Tag             =   "5950"
      Top             =   2130
      Width           =   5955
   End
   Begin VB.Image ImageLogo 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   2280
      Left            =   0
      Picture         =   "FRMSPLAS.frx":000C
      Top             =   0
      Width           =   6030
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'[----------------------------------------------]
'[frmSplash         Splash startup screen       ]
'[----------------------------------------------]
'[GSR                                           ]
'[generic staff roster system for Windows       ]
'[(C) ---------- David Gilbert, 1998            ]
'[----------------------------------------------]

Private Sub Form_Load()
    
    '[center form on screen]
    Top = (Screen.Height - frmSplash.Height) / 2
    Left = (Screen.Width - frmSplash.Width) / 2
    
    '[SET STARTUP (C) MESSAGE]
    frmSplash.lblVerInfo.Caption = "Generic Staff Roster " & constVersion & Chr$(vbKeyReturn) & " (C) 1998, David Gilbert."
    
End Sub

