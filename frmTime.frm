VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmTime 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select Time"
   ClientHeight    =   390
   ClientLeft      =   6150
   ClientTop       =   10260
   ClientWidth     =   3285
   ControlBox      =   0   'False
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FRMTIME.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   390
   ScaleWidth      =   3285
   Begin VB.CheckBox CheckResult 
      Alignment       =   1  'Right Justify
      Caption         =   "Result"
      Height          =   315
      Left            =   30
      TabIndex        =   4
      Top             =   900
      Value           =   1  'Checked
      Width           =   1965
   End
   Begin VB.ComboBox ComboHour 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   330
      ItemData        =   "FRMTIME.frx":000C
      Left            =   30
      List            =   "FRMTIME.frx":0066
      TabIndex        =   0
      Text            =   "ComboHour"
      Top             =   30
      Width           =   795
   End
   Begin VB.ComboBox ComboMinute 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   330
      ItemData        =   "FRMTIME.frx":00CA
      Left            =   870
      List            =   "FRMTIME.frx":00DA
      TabIndex        =   1
      Text            =   "ComboMinute"
      Top             =   30
      Width           =   795
   End
   Begin Threed.SSCommand cmdCancel 
      Cancel          =   -1  'True
      Height          =   330
      Left            =   2520
      TabIndex        =   3
      Top             =   36
      Width           =   756
      _Version        =   65536
      _ExtentX        =   1323
      _ExtentY        =   582
      _StockProps     =   78
      Caption         =   "C&ancel"
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
   Begin Threed.SSCommand cmdOK 
      Height          =   330
      Left            =   1716
      TabIndex        =   2
      Top             =   36
      Width           =   756
      _Version        =   65536
      _ExtentX        =   1323
      _ExtentY        =   582
      _StockProps     =   78
      Caption         =   "&OK"
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
End
Attribute VB_Name = "frmTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()

    '[SET RESULT OF BUTTON CHOICE]
    frmTime.CheckResult = vbUnchecked
    '[CLOSE FORM / ONLY UNLOAD THOUGH, LET CALLING ROUTINE UNLOAD FORM]
    frmTime.Hide

End Sub

Private Sub cmdCancel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Abandon any changes made and return."
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub cmdOK_Click()

    '[SET RESULT OF BUTTON CHOICE]
    frmTime.CheckResult = vbChecked
    '[CLOSE FORM / ONLY UNLOAD THOUGH, LET CALLING ROUTINE UNLOAD FORM]
    frmTime.Hide

End Sub


Private Sub cmdOK_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Accept any changes made and return."
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub Form_Load()

    '[CENTER FORM]
    frmTime.Left = (Screen.Width / 2) - (frmTime.Width / 2)
    frmTime.Top = (Screen.Height / 2) - (frmTime.Height / 2)

End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Select the required hour/minute."
    '[---------------------------------------------------------------------------------]

End Sub


