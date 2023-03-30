VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{A8B3B723-0B5A-101B-B22E-00AA0037B2FC}#1.0#0"; "GRID32.OCX"
Begin VB.Form frmScratch 
   Caption         =   "Scratch Roster"
   ClientHeight    =   4290
   ClientLeft      =   4635
   ClientTop       =   4935
   ClientWidth     =   6615
   ClipControls    =   0   'False
   FillColor       =   &H00FFFFFF&
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
   HelpContextID   =   60
   Icon            =   "FRMSCRAT.frx":0000
   LinkTopic       =   "frmRoster"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4290
   ScaleWidth      =   6615
   Visible         =   0   'False
   Begin MSGrid.Grid GridRoster 
      Height          =   3915
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   6615
      _Version        =   65536
      _ExtentX        =   11668
      _ExtentY        =   6906
      _StockProps     =   77
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Cols            =   9
      FixedCols       =   0
   End
   Begin Threed.SSPanel PanelToolBar 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      Negotiate       =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   6615
      _Version        =   65536
      _ExtentX        =   11668
      _ExtentY        =   635
      _StockProps     =   15
      ForeColor       =   -2147483641
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   1
      BevelOuter      =   0
      RoundedCorners  =   0   'False
      FloodType       =   1
      FloodColor      =   -2147483646
      FloodShowPct    =   0   'False
      Alignment       =   0
      MouseIcon       =   "FRMSCRAT.frx":08CA
      Begin VB.Image cmdRebuild 
         Height          =   300
         Left            =   3660
         Picture         =   "FRMSCRAT.frx":11A4
         Top             =   30
         Width           =   300
      End
      Begin VB.Image cmdRoster 
         Height          =   300
         Left            =   360
         Picture         =   "FRMSCRAT.frx":12F6
         Top             =   30
         Width           =   300
      End
      Begin VB.Image cmdInsertRow 
         Height          =   300
         Left            =   2760
         Picture         =   "FRMSCRAT.frx":15C8
         Top             =   30
         Width           =   300
      End
      Begin VB.Image cmdRemoveRow 
         Height          =   300
         Left            =   2460
         Picture         =   "FRMSCRAT.frx":170A
         Top             =   30
         Width           =   300
      End
      Begin VB.Image cmdExpand 
         Height          =   300
         Left            =   3360
         Picture         =   "FRMSCRAT.frx":184C
         Top             =   30
         Width           =   300
      End
      Begin VB.Image cmdResize 
         Height          =   300
         Left            =   3060
         Picture         =   "FRMSCRAT.frx":199E
         Top             =   30
         Width           =   300
      End
      Begin VB.Image cmdPaste 
         Height          =   240
         Left            =   1860
         Picture         =   "FRMSCRAT.frx":1AF0
         Top             =   30
         Width           =   240
      End
      Begin VB.Image cmdCopy 
         Height          =   300
         Left            =   1560
         Picture         =   "FRMSCRAT.frx":1BF2
         Top             =   30
         Width           =   300
      End
      Begin VB.Image cmdDelete 
         Height          =   300
         Left            =   1260
         Picture         =   "FRMSCRAT.frx":1D44
         Top             =   30
         Width           =   300
      End
   End
End
Attribute VB_Name = "frmScratch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'[----------------------------------------------]
'[frmScratch         Temp Roster Creation form  ]
'[----------------------------------------------]
'[GSR                                           ]
'[generic staff roster system for Windows       ]
'[(C) ---------- David Gilbert, 1997            ]
'[----------------------------------------------]
'[version                           2.1         ]
'[initial date                      21/07/1997  ]
'[----------------------------------------------]

Private Sub cmdExpand_Click()

    '[ANIMATE BUTTON]
    cmdExpand.BorderStyle = 1
    Delay vbDelay
    cmdExpand.BorderStyle = 0

    
    '[EXPAND CELLS TO MATCH CONTENTS]
    flagScratch = True
    Call ExpandCells
    flagScratch = False
    
End Sub


Private Sub cmdExpand_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Expand the roster cells to fit the contents."
    '[---------------------------------------------------------------------------------]

End Sub




Private Sub cmdHide_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Hide the scratch roster form."
    '[---------------------------------------------------------------------------------]

End Sub



Private Sub cmdLoad_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Open a file and load into the scratch roster."
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub cmdRebuild_Click()
    
    '[ANIMATE BUTTON]
    cmdRebuild.BorderStyle = 1
    Delay vbDelay
    cmdRebuild.BorderStyle = 0

    '[IF CHANGES HAVE BEEN MADE AND NOT SAVED, CHECK WITH USER]
    '[CHECK FOR CHANGED DATA AND NOTIFY]
    Dim Msg As String
    Dim Style
    Dim Response
    Dim Title

    '[REBUILDING ROSTER, POPUP YES/NO DIALOG]
    Msg = "Caution - This action will rebuild this roster using the settings on the Control Form - start time, end time and increment." & strBreak & strBreak & "This will result in this roster being cleared of all data currently entered." & strBreak & strBreak & "Do you wish to continue and rebuild this roster ?"
    Style = vbYesNo             ' Define buttons.
    Title = "Rebuilding Roster"  ' Define title.
    Response = gsrMsg(Msg, Style, Title)
    If Response = vbNo Then    ' User chose No.
        Exit Sub
    End If

    '[REBUILD ROSTER]
    flagScratch = True
    RebuildRoster
    flagScratch = False
    
    '[MOVE TO ROSTER FORM IF IT IS VISIBLE]
    If frmScratch.WindowState = 1 Then frmScratch.WindowState = 0
    frmScratch.ZOrder
    
End Sub

Private Sub cmdRebuild_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Rebuild the scratch roster using the defaults specified on the settings form."
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub cmdResize_Click()

    '[ANIMATE BUTTON]
    cmdResize.BorderStyle = 1
    Delay vbDelay
    cmdResize.BorderStyle = 0


    '[RESIZE ALL COLUMN WIDTHS TO MATCH GRID WIDTH]
    flagScratch = True
    Call FitRosterToGrid
    flagScratch = False
    
End Sub


Private Sub cmdResize_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Resize the roster columns to fit the roster grid."
    '[---------------------------------------------------------------------------------]

End Sub




Private Sub cmdRoster_Click()

    '[ANIMATE BUTTON]
    cmdRoster.BorderStyle = 1
    Delay vbDelay
    cmdRoster.BorderStyle = 0

    '[DISPLAY THE ROSTER FORM]
    frmRoster.ZOrder

End Sub

Private Sub cmdRoster_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "View the Main Roster Form."
    '[---------------------------------------------------------------------------------]

End Sub

Private Sub Form_GotFocus()

    '[SET SCRATCH FLAG]
    flagScratch = True

End Sub

Private Sub Form_Load()

    '[HIDE FORM WHILE LOADING]
    frmScratch.Visible = False
    flagScratch = True
    Call SetGridTitles
    flagScratch = False

End Sub



Private Sub cmdCopy_Click()

    '[ANIMATE BUTTON]
    cmdCopy.BorderStyle = 1
    Delay vbDelay
    cmdCopy.BorderStyle = 0

    '[ADD CLIP AREA OF ROSTER TO strCLIP]
    flagScratch = True
    strClip = CopyCells
    flagScratch = False

End Sub


Private Sub cmdCopy_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Copy selected roster cells to the clipboard."
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub cmdDelete_Click()

    '[ANIMATE BUTTON]
    cmdDelete.BorderStyle = 1
    Delay vbDelay
    cmdDelete.BorderStyle = 0

    '[CLEAR SELECTED CELLS IN ROSTER]
    flagScratch = True
    Call DeleteCellContents(frmScratch)
    flagScratch = False
    
End Sub


Private Sub cmdDelete_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Delete the selected cell contents from the roster."
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub cmdInsertRow_Click()

    '[ANIMATE BUTTON]
    cmdInsertRow.BorderStyle = 1
    Delay vbDelay
    cmdInsertRow.BorderStyle = 0


    '[INSERT ROW INTO ROSTER]
    Dim intRow As Integer
    If frmScratch.GridRoster.Row = 0 Then
        '[SET ROW TO FIRST ROW]
        intRow = 1
    Else
        intRow = frmScratch.GridRoster.Row + 1
    End If
    
    '[ADD NEW ROW ITEM TO THE GRID]
    frmScratch.GridRoster.AddItem Format(DsDefault("StartTime"), "Medium Time") & Chr$(vbKeyTab) & Format(DsDefault("EndTime"), "Medium Time"), intRow
    '[OLD TIME FORMAT] -> "hh:mm AMPM"
    
    '[MOVE TO THIS NEW ROW]
    frmScratch.GridRoster.Row = intRow

End Sub

Private Sub cmdInsertRow_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Insert a new row into the roster grid at the highlighted position."
    '[---------------------------------------------------------------------------------]

End Sub




Private Sub cmdPaste_Click()

    '[ANIMATE BUTTON]
    cmdPaste.BorderStyle = 1
    Delay vbDelay
    cmdPaste.BorderStyle = 0

    '[PUT strClip IN ROSTER]
    If Len(strClip) = 0 Then Exit Sub
    flagScratch = True
    PasteCells
    flagScratch = False
    
End Sub

Private Sub cmdPaste_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Paste the contents of the clipboard into the selected cells in the roster."
    '[---------------------------------------------------------------------------------]

End Sub






Private Sub cmdRemoveRow_Click()

    '[ANIMATE BUTTON]
    cmdRemoveRow.BorderStyle = 1
    Delay vbDelay
    cmdRemoveRow.BorderStyle = 0


    '[REMOVE SELECTED ROW FROM GRID]
    Dim intCounter          As Integer
    Dim intCol              As Integer
    Dim intRow              As Integer

    '[POPUP YES/NO DIALOG TO CONFIRM DELETION]
    Dim Msg As String
    Dim Style
    Dim Response
    Dim Title
    
    If frmScratch.GridRoster.Row = 0 Then Exit Sub
    intCol = frmScratch.GridRoster.Col
    
    '[CHECK DELETION FLAG]
    If flagDeleteConfirm Then
        Msg = "This action will remove this time slot from the roster." & strBreak & strBreak & "Do you wish to continue ?"
        Style = vbYesNo                     ' Define buttons.
        Title = "Confirmation Required"     ' Define title.
        Response = gsrMsg(Msg, Style, Title)
    Else
        Response = vbYes
    End If
    
    If Response = vbYes Then    ' User chose Yes.
        '[REMOVE ROW FROM ROSTER]
        intRow = frmScratch.GridRoster.Row
        If frmScratch.GridRoster.Rows = 2 Then
            For intCounter = 0 To 8
                frmScratch.GridRoster.Col = intCounter
                frmScratch.GridRoster.Text = ""
            Next intCounter
        Else
            frmScratch.GridRoster.RemoveItem intRow
        End If
        '[MOVE TO A NEW ROW]
        If intRow = frmScratch.GridRoster.Rows Then frmScratch.GridRoster.Row = intRow - 1
    End If

    frmScratch.GridRoster.Col = intCol


End Sub

Private Sub cmdRemoveRow_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Remove the selected row from the roster grid."
    '[---------------------------------------------------------------------------------]

End Sub



Private Sub cmdSave_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Save the scratch roster to a file."
    '[---------------------------------------------------------------------------------]

End Sub






Private Sub Form_LostFocus()

    '[SET SCRATCH FLAG]
    flagScratch = False
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "The roster form is used to create and modify your rosters."
    '[---------------------------------------------------------------------------------]
    
End Sub


Private Sub Form_Resize()

    '[TEMP VARIABLES SO WE CAN CATCH ILLEGAL WIDTHS]
    Dim sinWidth        As Single
    Dim sinHeight       As Single
    
    '[IF FORM IS MINIMISED THEN EXIT THIS ROUTINE]
    If frmScratch.WindowState = 1 Then Exit Sub
    '[RESIZE GRID AND ARRANGE CONTROLS ON FORM]
    
    '[ROSTER GRID]
    'frmScratch.GridRoster.Top = frmScratch.PanelToolBar.Height + 60
    sinWidth = frmScratch.Width - 300
    frmScratch.GridRoster.Left = 100
    
    If sinWidth > 0 Then frmScratch.GridRoster.Width = sinWidth
    sinHeight = frmScratch.Height - frmScratch.GridRoster.Top - 500
    If sinHeight > 0 Then frmScratch.GridRoster.Height = sinHeight
    
End Sub


Private Sub GridRoster_DblClick()
    
    '[ALLOW USER TO CHANGE THE SHIFT START/END TIME]
    Dim Msg As String
    Dim Style
    Dim Response
    Dim Title

    If frmScratch.GridRoster.Row = 0 Then Exit Sub
    
    Select Case frmScratch.GridRoster.Col
    Case 0, 1     '[SHIFT START TIME/END TIME]
        Load frmTime
        '[SET TIME ON FORM]
        If frmScratch.GridRoster.Text = "" Then
            If frmScratch.GridRoster.Col = 0 Then
                '[SHIFT START TIME]
                frmTime.Caption = "Select Start Time"
                frmTime.ComboHour = Format(Hour(DsDefault("StartTime")), "0#")
                frmTime.ComboMinute = Format(Minute(DsDefault("StartTime")), "0#")
                frmTime.Refresh
            Else
                '[SHIFT END TIME]
                frmTime.Caption = "Select End Time"
                frmTime.ComboHour = Format(Hour(DsDefault("EndTime")), "0#")
                frmTime.ComboMinute = Format(Minute(DsDefault("EndTime")), "0#")
                frmTime.Refresh
            End If
        Else
            If IsDate(frmScratch.GridRoster.Text) Then
                frmTime.ComboHour = Format(Hour(CDate(frmScratch.GridRoster.Text)), "0#")
                frmTime.ComboMinute = Format(Minute(CDate(frmScratch.GridRoster.Text)), "0#")
            Else
                frmTime.ComboHour = "00"
                frmTime.ComboMinute = "00"
            End If
        End If
        frmTime.Show 1
        '[PROCESS RESULT OF FORM]
        If frmTime.CheckResult = vbChecked Then
            '[ENSURE TIME RETURNED IS A TIME]
            If IsDate(frmTime.ComboHour.Text & ":" & frmTime.ComboMinute.Text) Then
                frmScratch.GridRoster.Text = Format(frmTime.ComboHour.Text & ":" & frmTime.ComboMinute.Text, "Medium Time")
                '[OLD FORMAT COMMAND] -> "hh:mm AMPM"
                '[CHECK FOR CELL SIZE AND RESIZE IF NECESSARY]
                '[ALLOW 10% MARGIN FOR TEXT ADJUSTMENT]
                If frmScratch.GridRoster.ColWidth(frmScratch.GridRoster.Col) < (TextWidth(frmScratch.GridRoster.Text) * 1.1) Then frmScratch.GridRoster.ColWidth(frmScratch.GridRoster.Col) = (TextWidth(frmScratch.GridRoster.Text) * 1.1)
            End If
        End If
        '[REMOVE FORM]
        'Unload frmTime

    Case Else   '[SHOW FULL CELL CONTENTS]
        If Len(Trim(frmScratch.GridRoster.Text)) = 0 Then Exit Sub
        Msg = frmScratch.GridRoster.Text
        Style = vbOKOnly                        ' Define buttons.
        Title = "Cell Contents"                 ' Define title.
        Response = gsrMsg(Msg, Style, Title)
    
    End Select

End Sub


Private Sub GridRoster_KeyUp(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
    Case vbKeyDelete
        '[CAPTURE DELETE KEY]
        Call cmdDelete_Click
    Case Else
    End Select
    
End Sub

Private Sub GridRoster_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    If X > (frmScratch.GridRoster.ColPos(2)) Then
        StatusBar "Double-click to display the names which appear in this cell."
    Else
        '[DISPLAY DIFFERENT MESSAGES FOR LOCKED OR UNLOCKED]
        If frmScratch.GridRoster.FixedCols = 0 Then
            StatusBar "Double-click to modify the shift start/finish time."
        Else
            StatusBar "Select 'Unlock Roster Columns' from the 'Options' menu to modify the Start/Finish times."
        End If
    End If
    '[---------------------------------------------------------------------------------]

End Sub

Private Sub GridRoster_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[POPUP COPY/PASTE/DELETE MENU]
    If Button = vbRightButton Then
        '[DISPLAY CONTEXT MENU ON RIGHT MOUSE CLICK]
        PopupMenu mdiMain.mnuScratchEdit, vbPopupMenuRightButton, , , mdiMain.mnuScratchCopy
    End If

End Sub


