VERSION 5.00
Begin VB.Form frmValidate 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Validation Code Generation"
   ClientHeight    =   1755
   ClientLeft      =   4875
   ClientTop       =   6885
   ClientWidth     =   5715
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FRMVALID.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1755
   ScaleWidth      =   5715
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtRegUser 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   1020
      TabIndex        =   4
      Top             =   360
      Width           =   4575
   End
   Begin VB.TextBox txtRegCode 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   1020
      TabIndex        =   3
      Top             =   750
      Width           =   1875
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   345
      Left            =   4740
      TabIndex        =   2
      Top             =   750
      Width           =   885
   End
   Begin VB.TextBox txtModifier 
      Height          =   315
      Left            =   3750
      TabIndex        =   0
      Text            =   "129"
      Top             =   750
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Please enter the registering user code below :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   4
      Left            =   60
      TabIndex        =   8
      Top             =   30
      Width           =   3915
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User name :"
      Height          =   210
      Index           =   1
      Left            =   60
      TabIndex        =   7
      Top             =   420
      Width           =   870
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Please check to ensure the registering user name matches that supplied by the user."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Index           =   0
      Left            =   60
      TabIndex        =   6
      Top             =   1170
      Width           =   5655
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Code :"
      Height          =   210
      Index           =   3
      Left            =   60
      TabIndex        =   5
      Top             =   810
      Width           =   465
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Modifier :"
      Height          =   210
      Index           =   2
      Left            =   3060
      TabIndex        =   1
      Top             =   810
      Visible         =   0   'False
      Width           =   660
   End
End
Attribute VB_Name = "frmValidate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Function Validate(strValidate) As String

    '[FUNCTION TO VALIDATE A STRING AND RETURN THE VALIDATION CODE]
    Dim strRegCode      As String       '[RETURNED HEX VALIDATION CODE]
    Dim sinValue        As Single       '[ACCUMULATED VALUE]
    Dim intCounter      As Integer      '[COUNTER FOR LENGTH]

    If IsNull(strValidate) Or Len(strValidate) = 0 Then
        '[NO CODE TO VALIDATE SO RETURN NULL]
        strRegCode = ""
    Else
        For intCounter = 1 To Len(strValidate)
        '[CYCLE THROUGH VALIDATION STRING AND ACCUMULATE VALUES]
            sinValue = sinValue + (Asc(Mid$(strValidate, intCounter, 1)) * Val(txtModifier))
        Next intCounter
        strRegCode = Hex(sinValue)
    End If
    
    '[RETURN CODE VALUE]
    Validate = strRegCode

End Function


Private Sub cmdDDE_Click()

    txtRegUser.LinkTopic = "GSR|frmRegister"       ' Set link topic.
    txtRegUser.LinkItem = "txtRegUser"          ' Set link item.
    txtRegUser.LinkMode = 1
    txtRegUser.LinkPoke                         ' Poke value to cell.
    
    txtRegCode.LinkTopic = "GSR|frmRegister"       ' Set link topic.
    txtRegCode.LinkItem = "txtRegCode"          ' Set link item.
    txtRegCode.LinkMode = 1
    txtRegCode.LinkPoke                         ' Poke value to cell.
    
        
End Sub

Private Sub cmdOK_Click()

    '[CLOSE VALIDATE FORM]
    Unload frmValidate
    End
    
End Sub

Private Sub Form_Load()
    
    '[center form on screen]
    Top = (Screen.Height / 2) - (Height / 2)
    Left = (Screen.Width / 2) - (Width / 2)


End Sub

Private Sub txtRegUser_Change()

    '[SHOW VALIDATION CODE]
    txtRegCode.Text = Validate(txtRegUser.Text)

End Sub


