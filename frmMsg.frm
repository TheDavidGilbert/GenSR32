VERSION 5.00
Begin VB.Form frmMsg 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Generic Staff Roster - Message"
   ClientHeight    =   3945
   ClientLeft      =   2685
   ClientTop       =   2850
   ClientWidth     =   5250
   ControlBox      =   0   'False
   FillColor       =   &H000000FF&
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
   Icon            =   "FRMMSG.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3945
   ScaleWidth      =   5250
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2610
      TabIndex        =   3
      Top             =   3450
      Width           =   945
   End
   Begin VB.CommandButton cmdcancel 
      Caption         =   "C&ancel"
      Height          =   375
      Left            =   3630
      TabIndex        =   4
      Top             =   3450
      Width           =   945
   End
   Begin VB.CommandButton cmdNo 
      Caption         =   "&No"
      Height          =   375
      Left            =   1590
      TabIndex        =   2
      Top             =   3450
      Width           =   945
   End
   Begin VB.TextBox TextNote 
      Height          =   2535
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   690
      Visible         =   0   'False
      Width           =   5100
   End
   Begin VB.CommandButton cmdYes 
      Caption         =   "&Yes"
      Height          =   375
      Left            =   570
      TabIndex        =   1
      Top             =   3450
      Width           =   945
   End
   Begin VB.Shape gaugeBar 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   225
      Left            =   140
      Tag             =   "4950"
      Top             =   3630
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.Shape gaugeBorder 
      BorderWidth     =   2
      Height          =   345
      Left            =   60
      Top             =   3570
      Visible         =   0   'False
      Width           =   5100
   End
   Begin VB.Label LabelMessage 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "Label Message"
      Height          =   2300
      Left            =   150
      TabIndex        =   6
      Top             =   800
      Width           =   4935
      WordWrap        =   -1  'True
   End
   Begin VB.Line LineDivider 
      BorderWidth     =   2
      X1              =   30
      X2              =   5220
      Y1              =   630
      Y2              =   630
   End
   Begin VB.Label LabelTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Title Label"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   60
      TabIndex        =   5
      Top             =   150
      Width           =   5100
   End
   Begin VB.Image ImageInfo 
      Height          =   480
      Index           =   0
      Left            =   30
      Picture         =   "FRMMSG.frx":000C
      Top             =   90
      Width           =   480
   End
   Begin VB.Label LabelInfo 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   140
      TabIndex        =   7
      Top             =   3300
      Visible         =   0   'False
      Width           =   4950
   End
End
Attribute VB_Name = "frmMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'[----------------------------------------------]
'[frmMsg            GSR Message Form            ]
'[----------------------------------------------]
'[GSR                                           ]
'[generic staff roster system for Windows       ]
'[(C) ---------- David Gilbert, 1997            ]
'[----------------------------------------------]
'[version                           2.1         ]
'[initial date                      21/07/1997  ]
'[----------------------------------------------]



Private Sub cmdCancel_Click()
    
    '[SAVE RETURN CODE AND UNLOAD FORM]
    gsrReturn = vbCancel
    Unload frmMsg

End Sub

Private Sub cmdCancel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Cancel this operation and return."
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub cmdNo_Click()
    
    '[SAVE RETURN CODE AND UNLOAD FORM]
    gsrReturn = vbNo
    Unload frmMsg

End Sub

Private Sub cmdNo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Answer NO to the question and continue."
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub cmdOK_Click()
    
    '[SAVE RETURN CODE AND UNLOAD FORM]
    gsrReturn = vbOK
    '[IF THIS IS A TEXT GET MESSAGE, SAVE THE CHANGED TEXT]
    If frmMsg.TextNote.Visible = True Then gsrNote = frmMsg.TextNote.Text
    Unload frmMsg

End Sub

Private Sub cmdOK_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Continue once you have read the message."
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub cmdYes_Click()

    '[SAVE RETURN CODE AND UNLOAD FORM]
    gsrReturn = vbYes
    Unload frmMsg
    

End Sub


Private Sub cmdYes_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Answer YES to the question and continue."
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub Form_Load()

    '[center form on screen]
    frmMsg.Top = (Screen.Height / 2) - (frmMsg.Height / 2)
    frmMsg.Left = (Screen.Width / 2) - (frmMsg.Width / 2)

End Sub




Private Sub Form_LostFocus()

    '[GIVE THE FORM THE FOCUS AGAIN]
    On Error Resume Next
    frmMsg.SetFocus

End Sub


Private Sub LabelMessage_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Displays relevant messages about GSR's activites."
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub TextNote_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Enter notes for the selected roster."
    '[---------------------------------------------------------------------------------]

End Sub


