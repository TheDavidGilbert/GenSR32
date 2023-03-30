VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmSearch 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Search"
   ClientHeight    =   1305
   ClientLeft      =   5370
   ClientTop       =   6090
   ClientWidth     =   5400
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   7.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   51
   Icon            =   "FRMSEARC.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1305
   ScaleWidth      =   5400
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frameOptions 
      Caption         =   "Search Options"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3120
      TabIndex        =   9
      Top             =   0
      Width           =   2265
      Begin VB.CheckBox checkReplace 
         Alignment       =   1  'Right Justify
         Caption         =   "Search and Replace"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   60
         TabIndex        =   2
         Top             =   270
         Width           =   2100
      End
      Begin VB.CheckBox checkConstraints 
         Alignment       =   1  'Right Justify
         Caption         =   "Ignore Constraints"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   60
         TabIndex        =   3
         Top             =   540
         Visible         =   0   'False
         Width           =   2100
      End
   End
   Begin VB.ComboBox comboReplaceName 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1300
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   510
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.ComboBox comboSearchName 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1300
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   150
      Width           =   1800
   End
   Begin Threed.SSCommand cmdReplaceAll 
      Height          =   360
      Left            =   3096
      TabIndex        =   7
      Top             =   936
      Visible         =   0   'False
      Width           =   1008
      _Version        =   65536
      _ExtentX        =   1778
      _ExtentY        =   635
      _StockProps     =   78
      Caption         =   "Replace &All"
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
   Begin Threed.SSCommand cmdReplace 
      Height          =   360
      Left            =   2076
      TabIndex        =   6
      Top             =   936
      Visible         =   0   'False
      Width           =   1008
      _Version        =   65536
      _ExtentX        =   1778
      _ExtentY        =   635
      _StockProps     =   78
      Caption         =   "R&eplace"
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
   Begin Threed.SSCommand cmdFindNext 
      Height          =   360
      Left            =   1056
      TabIndex        =   5
      Top             =   936
      Width           =   1008
      _Version        =   65536
      _ExtentX        =   1778
      _ExtentY        =   635
      _StockProps     =   78
      Caption         =   "Find &Next"
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
   Begin Threed.SSCommand cmdFindFirst 
      Default         =   -1  'True
      Height          =   360
      Left            =   36
      TabIndex        =   4
      Top             =   936
      Width           =   1008
      _Version        =   65536
      _ExtentX        =   1778
      _ExtentY        =   635
      _StockProps     =   78
      Caption         =   "F&ind First"
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
   Begin Threed.SSCommand cmdReturn 
      Cancel          =   -1  'True
      Height          =   360
      Left            =   4680
      TabIndex        =   8
      Top             =   930
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
   Begin VB.Label labelHeading 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Replace With"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   315
      Index           =   1
      Left            =   15
      TabIndex        =   11
      Top             =   510
      Visible         =   0   'False
      Width           =   1250
   End
   Begin VB.Label labelHeading 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Search For"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   315
      Index           =   0
      Left            =   15
      TabIndex        =   10
      Top             =   150
      Width           =   1250
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'[FIND (AND REPLACE) STAFF NAMES WITHIN THE CURRENT ROSTER]
Option Explicit

Dim intRow              As Integer      '[STARTING ROW]
Dim intCol              As Integer      '[STARTING COL]



Private Sub checkConstraints_GotFocus()
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Ignore staff availability restrictions ?"
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub checkConstraints_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Ignore staff availability restrictions ?"
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub checkReplace_Click()

    Select Case checkReplace.Value
    Case vbChecked      '[REPLACE FOUND NAME WITH NAME IN SECOND LIST]
        frmSearch.comboReplaceName.Visible = True
        frmSearch.labelHeading(1).Visible = True
        frmSearch.cmdReplace.Visible = True
        frmSearch.cmdReplaceAll.Visible = True
        frmSearch.checkConstraints.Visible = True
    Case vbUnchecked    '[SEARCH ONLY]
        frmSearch.comboReplaceName.Visible = False
        frmSearch.labelHeading(1).Visible = False
        frmSearch.cmdReplace.Visible = False
        frmSearch.cmdReplaceAll.Visible = False
        frmSearch.checkConstraints.Visible = True
    Case Else
    End Select
    
End Sub


Private Sub checkReplace_GotFocus()
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "'Search' or 'Search and Replace' ?"
    '[---------------------------------------------------------------------------------]

End Sub

Private Sub checkReplace_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "'Search' or 'Search and Replace' ?"
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub cmdFindFirst_Click()

    '[FIND FIRST OCCURRENCE OF THIS NAME WITHIN THE CURRENT ROSTER]
    Dim flagFound       As Boolean
    Dim intRowCount     As Integer
    Dim intColCount     As Integer
    Dim Msg As String
    Dim Style
    Dim Response
    Dim Title
    
    intCol = frmRoster.gridRoster.Col
    intRow = frmRoster.gridRoster.Row
    flagFound = False
    
    For intRowCount = 1 To (frmRoster.gridRoster.Rows - 1)
        frmRoster.gridRoster.Row = intRowCount
        For intColCount = 2 To 8
            frmRoster.gridRoster.Col = intColCount
            '[CHECK TO SEE IF NAME IS IN CELL]
            If InStr(frmRoster.gridRoster.Text, frmSearch.comboSearchName) > 0 Then
                '[NAME FOUND - EXIT SUBROUTINE]
                Call ShowRowCol(intRowCount, intColCount)
                Exit Sub
            End If
        Next intColCount
    Next intRowCount
    
    '[NAME NOT FOUND - POPUP MESSAGE]
    Msg = "Could not locate the staff member (" & frmSearch.comboSearchName & ") within this roster."
    Style = vbOKOnly ' Define buttons.
    Title = "Not Found"  ' Define title.
    Response = gsrMsg(Msg, Style, Title)
    '[RESTORE CELL POSITION]
    frmRoster.gridRoster.Col = intCol
    frmRoster.gridRoster.Row = intRow
    
End Sub

Private Sub cmdFindFirst_GotFocus()
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Find the first occurence of the selected staff name in this roster."
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub cmdFindFirst_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Find the first occurence of the selected staff name in this roster."
    '[---------------------------------------------------------------------------------]

End Sub

Private Sub cmdFindNext_Click()
    
    '[FIND NEXT OCCURRENCE OF THIS NAME WITHIN THE CURRENT ROSTER]
    Dim flagFound       As Boolean
    Dim intRowCount     As Integer
    Dim intColCount     As Integer
    Dim Msg As String
    Dim Style
    Dim Response
    Dim Title
    
    intCol = frmRoster.gridRoster.Col
    intRow = frmRoster.gridRoster.Row
    flagFound = False
    
    For intRowCount = intRow To (frmRoster.gridRoster.Rows - 1)
        frmRoster.gridRoster.Row = intRowCount
        For intColCount = 2 To 8
            frmRoster.gridRoster.Col = intColCount
            '[CHECK FOR SAME ROW/COLS]
            If Not (intRowCount = intRow And frmRoster.gridRoster.Col <= intCol) Then
                '[CHECK TO SEE IF NAME IS IN CELL]
                If InStr(frmRoster.gridRoster.Text, frmSearch.comboSearchName) > 0 Then
                    '[NAME FOUND - EXIT SUBROUTINE]
                    Call ShowRowCol(intRowCount, intColCount)
                    Exit Sub
                End If
            End If
        Next intColCount
    Next intRowCount
    
    '[NAME NOT FOUND - POPUP MESSAGE]
    Msg = "Could not locate the next occurence of staff member (" & frmSearch.comboSearchName & ") within this roster."
    Style = vbOKOnly ' Define buttons.
    Title = "Not Found"  ' Define title.
    Response = gsrMsg(Msg, Style, Title)
    '[RESTORE CELL POSITION]
    frmRoster.gridRoster.Col = intCol
    frmRoster.gridRoster.Row = intRow
    
End Sub

Private Sub cmdFindNext_GotFocus()
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Find the next occurence of the selected staff name in this roster."
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub cmdFindNext_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Find the next occurence of the selected staff name in this roster."
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub cmdReplace_Click()

    Dim Msg As String
    Dim Style
    Dim Response
    Dim Title
    Dim flagConstraints     As Boolean
    Dim flagResult          As Boolean
    
    '[REPLACE OCCURENCE OF SEARCH NAME WITH REPLACE NAME]
    If frmSearch.checkConstraints.Value = vbChecked Then flagConstraints = False Else flagConstraints = True
    
    '[EXIT IF NAMES MATCH]
    If frmSearch.comboSearchName.Text = frmSearch.comboReplaceName.Text Then
        '[REPLACE NAME ALREADY IN CELL]
        Msg = "You have selected identical names for search and replace (" & frmSearch.comboSearchName & ")." & strBreak & strBreak & "The search name and the replace name must be different."
        Style = vbOKOnly ' Define buttons.
        Title = "Cannot Replace - Names Identical"  ' Define title.
        Response = gsrMsg(Msg, Style, Title)
        Exit Sub
    End If
    
    If InStr(frmRoster.gridRoster.Text, frmSearch.comboSearchName.Text) > 0 Then
        If InStr(frmRoster.gridRoster.Text, frmSearch.comboReplaceName.Text) > 0 Then
            '[REPLACE NAME ALREADY IN CELL]
            Msg = "The replacement name (" & frmSearch.comboReplaceName & ") is already in this roster cell." & strBreak & strBreak & "Replacements can only be made when the replacement name doesn't exist within the selected cell."
            Style = vbOKOnly ' Define buttons.
            Title = "Cannot Replace - Name Exists"  ' Define title.
            Response = gsrMsg(Msg, Style, Title)
        Else
            Call TransferToRoster(frmSearch.comboReplaceName.Text, flagConstraints, flagResult, flagContinue)
            '[ONLY REMOVE IF SUCCESSFUL PLACEMENT]
            If flagResult = True Then Call RemoveFromRoster(frmSearch.comboSearchName.Text)
        End If
    End If
    
End Sub

Private Sub cmdReplace_GotFocus()
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Replace this occurence of the selected staff name."
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub cmdReplace_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Replace this occurence of the selected staff name."
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub cmdReplaceAll_Click()

    '[SEARCH FOR AND REPLACE ALL OCCURENCES OF THE SEARCH NAME WITHIN THIS ROSTER]
    Dim intRowCount     As Integer
    Dim intColCount     As Integer
    Dim Msg As String
    Dim Style
    Dim Response
    Dim Title
    Dim flagResult      As Boolean
    Dim intReplaceCount As Integer
    Dim intFoundCount   As Integer
    
    '[EXIT IF NAMES MATCH]
    If frmSearch.comboSearchName.Text = frmSearch.comboReplaceName.Text Then
        '[REPLACE NAME ALREADY IN CELL]
        Msg = "You have selected identical names for search and replace (" & frmSearch.comboSearchName & ")." & strBreak & strBreak & "The search name and the replace name must be different."
        Style = vbOKOnly ' Define buttons.
        Title = "Cannot Replace - Names Identical"  ' Define title.
        Response = gsrMsg(Msg, Style, Title)
        Exit Sub
    End If
    
    intCol = frmRoster.gridRoster.Col
    intRow = frmRoster.gridRoster.Row
    
    For intRowCount = 1 To (frmRoster.gridRoster.Rows - 1)
        frmRoster.gridRoster.Row = intRowCount
        For intColCount = 2 To 8
            frmRoster.gridRoster.Col = intColCount
            If frmSearch.checkConstraints.Value = vbChecked Then flagConstraints = False Else flagConstraints = True
            '[CHECK TO SEE IF NAME IS IN CELL]
            If InStr(frmRoster.gridRoster.Text, frmSearch.comboSearchName) > 0 Then
            '[REPLACE OCCURENCE OF SEARCH NAME WITH REPLACE NAME]
                intFoundCount = intFoundCount + 1
                '[MOVE TO CELL]
                Call ShowRowCol(intRowCount, intColCount)
                If InStr(frmRoster.gridRoster.Text, frmSearch.comboReplaceName.Text) Then
                    '[REPLACE NAME ALREADY IN CELL]
                    Msg = "The replacement name (" & frmSearch.comboReplaceName & ") is already in this roster cell." & strBreak & strBreak & "Replacements can only be made when the replacement name doesn't exist within the selected cell."
                    Style = vbOKCancel                      ' Define buttons.
                    Title = "Cannot Replace - Name Exists"  ' Define title.
                    Response = gsrMsg(Msg, Style, Title)
                    '[EXIT SUB IF CANCEL PRESSED]
                    If Response = vbCancel Then Exit Sub
                Else
                    Call TransferToRoster(frmSearch.comboReplaceName.Text, flagConstraints, flagResult, flagContinue)
                    '[ONLY REMOVE IF SUCCESSFUL PLACEMENT]
                    If flagResult = True Then
                        Call RemoveFromRoster(frmSearch.comboSearchName.Text)
                        intReplaceCount = intReplaceCount + 1
                    End If
                End If
            End If
        Next intColCount
    Next intRowCount
    
    '[REPLACE FINISHED - POPUP MESSAGE]
    If intFoundCount = 1 Then
        Msg = "Located 1 occurence of " & frmSearch.comboSearchName & "." & strBreak & "Replaced " & intReplaceCount & " with " & frmSearch.comboReplaceName & "."
    Else
        Msg = "Located " & intFoundCount & " occurences of " & frmSearch.comboSearchName & "." & strBreak & "Replaced " & intReplaceCount & " with " & frmSearch.comboReplaceName & "."
    End If
    Style = vbOKOnly ' Define buttons.
    Title = "Search and Replace Complete"  ' Define title.
    Response = gsrMsg(Msg, Style, Title)
    
    '[RESTORE CELL POSITION]
    frmRoster.gridRoster.Col = intCol
    frmRoster.gridRoster.Row = intRow

End Sub

Private Sub cmdReplaceAll_GotFocus()
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Replace ALL occurences of the selected staff name."
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub cmdReplaceAll_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Replace ALL occurences of the selected staff name."
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub cmdReturn_Click()

    '[HIDE FORM AND RETURN TO ROSTER FORM]
    frmSearch.Hide
    frmRoster.ZOrder

End Sub

Private Sub cmdReturn_GotFocus()
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Close the search and replace form."
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub cmdReturn_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Close the search and replace form."
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub comboReplaceName_GotFocus()
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Select the staff name to replace the found name with."
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub comboSearchname_GotFocus()
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Select the staff name to search for."
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub Form_Load()
    
    '[CENTER FORM]
    frmSearch.Left = (Screen.Width / 2) - (frmSearch.Width / 2)
    frmSearch.Top = (Screen.Height / 2) - (frmSearch.Height / 2)

End Sub


