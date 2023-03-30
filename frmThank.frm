VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmThank 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3930
   ClientLeft      =   3915
   ClientTop       =   6570
   ClientWidth     =   4755
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
   Icon            =   "FRMTHANK.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3930
   ScaleWidth      =   4755
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frameBorder 
      Height          =   1815
      Left            =   0
      TabIndex        =   2
      Top             =   -90
      Width           =   4755
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
         TabIndex        =   6
         Top             =   900
         Width           =   3000
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(C) 1998, David Gilbert"
         Height          =   210
         Left            =   1440
         TabIndex        =   5
         Top             =   1560
         Width           =   1635
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
         TabIndex        =   4
         Top             =   240
         Width           =   3000
      End
      Begin VB.Label LabelVersion 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Version"
         Height          =   210
         Left            =   60
         TabIndex        =   3
         Top             =   1560
         Width           =   570
      End
      Begin VB.Image ImageLogo 
         Height          =   1575
         Left            =   3120
         Picture         =   "FRMTHANK.frx":000C
         Top             =   180
         Width           =   1560
      End
   End
   Begin Threed.SSCommand cmdReturn 
      Cancel          =   -1  'True
      Default         =   -1  'True
      Height          =   360
      Left            =   4020
      TabIndex        =   0
      Top             =   3540
      Width           =   705
      _Version        =   65536
      _ExtentX        =   1244
      _ExtentY        =   635
      _StockProps     =   78
      Caption         =   "C&lose"
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
   Begin VB.Label LabelInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BackStyle       =   0  'Transparent
      Caption         =   "Generic Staff Roster,  (C) 1997 David Gilbert."
      Height          =   1815
      Left            =   60
      TabIndex        =   1
      Top             =   1740
      Width           =   4695
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmThank"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdReturn_Click()

    '[UNLOAD FORM AND RETURN CONTROL]
    Unload frmThank
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

    frmThank.LabelVersion.Caption = constVersion
    '[BETA VERSION]
    If flagBeta = True Then frmThank.LabelVersion.Caption = frmThank.LabelVersion.Caption + " beta"

    '[SETUP TEXT MESSAGE BOX]
    frmThank.labelInfo = DsDefault("RegUser") & ", thank you for taking the time to register this copy of Generic Staff Roster." & strBreak
    frmThank.labelInfo = frmThank.labelInfo & strBreak & "We appreciate your support and hope that you find GSR to be an invaluable tool." & strBreak
    frmThank.labelInfo = frmThank.labelInfo & strBreak & "Your name has been added to our customer database and you will receive regular updates on any changes to GSR." & strBreak
    
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Thank you for registering Generic Staff Roster."
    '[---------------------------------------------------------------------------------]

End Sub


