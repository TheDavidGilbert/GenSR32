VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{A8B3B723-0B5A-101B-B22E-00AA0037B2FC}#1.0#0"; "GRID32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmRoster 
   AutoRedraw      =   -1  'True
   Caption         =   "Roster"
   ClientHeight    =   6300
   ClientLeft      =   3135
   ClientTop       =   4245
   ClientWidth     =   8790
   ClipControls    =   0   'False
   ControlBox      =   0   'False
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
   HelpContextID   =   50
   Icon            =   "frmRoster.frx":0000
   LinkTopic       =   "frmRoster"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6300
   ScaleWidth      =   8790
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin MSFlexGridLib.MSFlexGrid GridRoster 
      CausesValidation=   0   'False
      Height          =   5835
      Left            =   1860
      TabIndex        =   5
      Top             =   420
      Width           =   6915
      _ExtentX        =   12197
      _ExtentY        =   10292
      _Version        =   393216
      Cols            =   9
      FixedCols       =   2
      WordWrap        =   -1  'True
      GridLinesFixed  =   1
      MergeCells      =   1
      AllowUserResizing=   3
      Appearance      =   0
      FormatString    =   ""
   End
   Begin MSGrid.Grid GridRoster_ 
      Height          =   2235
      Left            =   1860
      TabIndex        =   4
      Top             =   4020
      Visible         =   0   'False
      Width           =   6795
      _Version        =   65536
      _ExtentX        =   11986
      _ExtentY        =   3942
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
      FixedCols       =   2
   End
   Begin VB.ComboBox ComboClass 
      Height          =   330
      Left            =   15
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   420
      Width           =   1800
   End
   Begin VB.ListBox ListStaff 
      Height          =   5520
      Left            =   15
      MultiSelect     =   2  'Extended
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   780
      Width           =   1800
   End
   Begin Threed.SSPanel PanelToolBar 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      Negotiate       =   -1  'True
      TabIndex        =   3
      Top             =   0
      Width           =   8790
      _Version        =   65536
      _ExtentX        =   15505
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
      RoundedCorners  =   0   'False
      FloodType       =   1
      FloodColor      =   -2147483646
      FloodShowPct    =   0   'False
      Alignment       =   4
      Autosize        =   2
      MouseIcon       =   "frmRoster.frx":08CA
      Begin MSMask.MaskEdBox MaskDate 
         DataField       =   "DateHired"
         Height          =   300
         Left            =   4620
         TabIndex        =   2
         Top             =   30
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   529
         _Version        =   393216
         ClipMode        =   1
         Appearance      =   0
         BackColor       =   16777215
         ForeColor       =   0
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         Format          =   "Short Date"
         PromptChar      =   "_"
      End
      Begin VB.Image cmdScratch 
         Height          =   300
         Left            =   360
         Picture         =   "frmRoster.frx":11A4
         Top             =   30
         Width           =   300
      End
      Begin VB.Image cmdPick 
         Height          =   300
         Left            =   4260
         Picture         =   "frmRoster.frx":1476
         Top             =   30
         Width           =   300
      End
      Begin VB.Image cmdInsertRow 
         Height          =   300
         Left            =   2760
         Picture         =   "frmRoster.frx":15C8
         Top             =   30
         Width           =   300
      End
      Begin VB.Image cmdRemoveRow 
         Height          =   300
         Left            =   2460
         Picture         =   "frmRoster.frx":170A
         Top             =   30
         Width           =   300
      End
      Begin VB.Image cmdToday 
         Height          =   300
         Left            =   6240
         Picture         =   "frmRoster.frx":184C
         Top             =   30
         Width           =   300
      End
      Begin VB.Image cmdSearch 
         Height          =   300
         Left            =   3960
         Picture         =   "frmRoster.frx":198E
         Top             =   30
         Width           =   300
      End
      Begin VB.Image cmdRebuild 
         Height          =   300
         Left            =   3660
         Picture         =   "frmRoster.frx":1AD0
         Top             =   30
         Width           =   300
      End
      Begin VB.Image cmdExpand 
         Height          =   300
         Left            =   3360
         Picture         =   "frmRoster.frx":1C22
         Top             =   30
         Width           =   300
      End
      Begin VB.Image cmdResize 
         Height          =   300
         Left            =   3060
         Picture         =   "frmRoster.frx":1D74
         Top             =   30
         Width           =   300
      End
      Begin VB.Image cmdPaste 
         Height          =   240
         Left            =   1860
         Picture         =   "frmRoster.frx":1EC6
         Top             =   30
         Width           =   240
      End
      Begin VB.Image cmdCopy 
         Height          =   300
         Left            =   1560
         Picture         =   "frmRoster.frx":1FC8
         Top             =   30
         Width           =   300
      End
      Begin VB.Image cmdDelete 
         Height          =   300
         Left            =   1260
         Picture         =   "frmRoster.frx":211A
         Top             =   30
         Width           =   300
      End
      Begin VB.Image cmdSave 
         Height          =   300
         Left            =   60
         Picture         =   "frmRoster.frx":226C
         Top             =   30
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Image cmdTransfer 
         Height          =   300
         Left            =   960
         Picture         =   "frmRoster.frx":23BE
         Top             =   30
         Width           =   300
      End
      Begin VB.Image cmdRemove 
         Height          =   300
         Left            =   660
         Picture         =   "frmRoster.frx":2510
         Top             =   30
         Width           =   300
      End
   End
   Begin VB.Image imageDrop 
      Appearance      =   0  'Flat
      DragIcon        =   "frmRoster.frx":2662
      Height          =   480
      Left            =   600
      Picture         =   "frmRoster.frx":2F2C
      Top             =   4500
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imageDrag 
      Appearance      =   0  'Flat
      DragIcon        =   "frmRoster.frx":37F6
      Height          =   480
      Left            =   60
      Picture         =   "frmRoster.frx":40C0
      Top             =   4500
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "frmRoster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'[----------------------------------------------]
'[frmRoster         Roster Creation form        ]
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

    '[GRID REDRAW]
    GridRedraw False
    
    '[EXPAND CELLS TO MATCH CONTENTS]
    Call ExpandCells
    
    '[GRID REDRAW]
    GridRedraw True
    
End Sub


Private Sub cmdExpand_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Expand the roster cells to fit the contents."
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub cmdPick_Click()

    '[ANIMATE BUTTON]
    cmdPick.BorderStyle = 1
    Delay vbDelay
    cmdPick.BorderStyle = 0

    '[PICK STAFF TO FIT]
    Call procPickStaff

End Sub

Private Sub cmdPick_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Pick staff members who can fill the selected time slot."
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub cmdResize_Click()

    '[ANIMATE BUTTON]
    cmdResize.BorderStyle = 1
    Delay vbDelay
    cmdResize.BorderStyle = 0

    '[GRID REDRAW]
    GridRedraw False

    '[RESIZE ALL COLUMN WIDTHS TO MATCH GRID WIDTH]
    Call FitRosterToGrid
    
    '[GRID REDRAW]
    GridRedraw True
    
End Sub


Private Sub cmdResize_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Resize the roster columns to fit the roster grid."
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub cmdScratch_Click()

    '[ANIMATE BUTTON]
    cmdScratch.BorderStyle = 1
    Delay vbDelay
    cmdScratch.BorderStyle = 0

    '[DISPLAY THE SCRATCH ROSTER]
    '[SELECT WINDOWSTATE]
    If frmRoster.WindowState = 2 Then
        frmScratch.ZOrder
        frmScratch.WindowState = 2
    Else
        frmScratch.ZOrder
        frmScratch.WindowState = 0
    End If

End Sub


Private Sub cmdScratch_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "View the Scratch Roster Form."
    '[---------------------------------------------------------------------------------]

End Sub



Private Sub cmdSearch_Click()

    '[ANIMATE BUTTON]
    cmdSearch.BorderStyle = 1
    Delay vbDelay
    cmdSearch.BorderStyle = 0


    '[DECLARATIONS]
    Dim intCounter          As Integer
    Dim strSearchName       As String
    Dim strReplaceName      As String
    
    '[SHOW SEARCH FORM AND REFRESH STAFF LISTS]
    frmSearch.Show
    '[REFRESH STAFF LISTS]
    strSearchName = frmRoster.ListStaff.Text
    strReplaceName = frmSearch.comboReplaceName.Text
    If strReplaceName = "" Then strReplaceName = frmRoster.ListStaff.Text
    
    
    frmSearch.comboSearchName.Clear
    frmSearch.comboReplaceName.Clear
    
    '[FILL WITH DATA FROM STAFF LIST ON STAFF FORM]
    For intCounter = 0 To frmSet.ListStaff.ListCount - 1
        frmSearch.comboSearchName.AddItem frmSet.ListStaff.List(intCounter)
        frmSearch.comboReplaceName.AddItem frmSet.ListStaff.List(intCounter)
    Next intCounter
    
    frmSearch.comboSearchName.Text = strSearchName
    frmSearch.comboReplaceName.Text = strReplaceName

End Sub


Private Sub cmdSearch_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Search for a staff member within this roster."
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub cmdToday_Click()

    '[ANIMATE BUTTON]
    cmdToday.BorderStyle = 1
    Delay vbDelay
    cmdToday.BorderStyle = 0

    '[SET MSK DATE TO TODAYS DATE]
    frmRoster.MaskDate.Text = Format(Date, strDateFormat)

End Sub



Private Sub cmdToday_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Set the 'Roster Start Date' to todays date (" & Format(Now, strDateFormat) & ")."

End Sub



Private Sub Form_GotFocus()

    '[SET SCRATCH FLAG]
    flagScratch = False
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If flagTerminate = False Then
        frmRoster.Hide
        Cancel = True
    End If

End Sub

Private Sub GridRoster_DragDrop(Source As Control, X As Single, Y As Single)

    If Source.Name = "imageDrag" And Source.DragIcon = frmRoster.imageDrop.DragIcon Then
        '[SOURCE OF DRAG WAS STAFF LIST BOX]
        Call cmdTransfer_Click
        intLastCheckRow = -1
        intLastCheckCol = -1
    End If

End Sub

Private Sub GridRoster_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    
    '[IF NO STATE THEN CHANGE CURSOR TO NO DROP]
    Dim intRowCounter           As Integer
    Dim intColCounter           As Integer
    Dim flagResult              As Integer
    Dim intResult               As Integer
    Dim dateStart               As Date
    Dim dateEnd                 As Date
    Dim intRow                  As Integer
    Dim intCol                  As Integer
    Dim Msg                     As String
    
    If State = 1 Then
        Source.DragIcon = frmRoster.imageDrop.Picture
        intLastCheckRow = -1
        intLastCheckCol = -1
        Exit Sub
    End If
    
    If Source.Name = "imageDrag" Then
        '[CHANGE TARGET CELL]
        For intColCounter = 2 To (frmRoster.GridRoster.Cols - 1)
            If X > (frmRoster.GridRoster.ColPos(intColCounter)) And X < (frmRoster.GridRoster.ColPos(intColCounter) + frmRoster.GridRoster.ColWidth(intColCounter)) Then
                frmRoster.GridRoster.Col = intColCounter
'                frmRoster.GridRoster.SelStartCol = intColCounter
'                frmRoster.GridRoster.SelEndCol = intColCounter
                frmRoster.GridRoster.Col = intColCounter
                frmRoster.GridRoster.ColSel = intColCounter
                Exit For
            End If
        Next intColCounter
        For intRowCounter = 1 To (frmRoster.GridRoster.Rows - 1)
            If Y > (frmRoster.GridRoster.RowPos(intRowCounter)) And Y < (frmRoster.GridRoster.RowPos(intRowCounter) + frmRoster.GridRoster.RowHeight(intRowCounter)) Then
                frmRoster.GridRoster.Row = intRowCounter
'                frmRoster.GridRoster.SelStartRow = intRowCounter
'                frmRoster.GridRoster.SelEndRow = intRowCounter
                frmRoster.GridRoster.Row = intRowCounter
                frmRoster.GridRoster.RowSel = intRowCounter
                Exit For
            End If
        Next intRowCounter
        '[CHECK STAFF AVAILABLE ON THIS DAY]
        If frmRoster.ListStaff.SelCount = 1 Then
            intRow = frmRoster.GridRoster.Row
            intCol = frmRoster.GridRoster.Col

            frmRoster.GridRoster.Col = 0: If IsDate(frmRoster.GridRoster.Text) Then dateStart = frmRoster.GridRoster.Text Else Exit Sub
            frmRoster.GridRoster.Col = 1: If IsDate(frmRoster.GridRoster.Text) Then dateEnd = frmRoster.GridRoster.Text Else Exit Sub
            frmRoster.GridRoster.Row = intRow
            frmRoster.GridRoster.Col = intCol
            '[CHECK IF ROW AND COL HAVE CHANGED]
            If intRow = intLastCheckRow And intCol = intLastCheckCol Then
                '[NOTHING TO DO - SAME ROW AND COL]
            Else
                flagResult = CheckStaffDay(frmRoster.ListStaff.List(frmRoster.ListStaff.ListIndex), (intCol - 1), dateStart, dateEnd, intResult, "")
                If flagResult = False Then
                    Source.DragIcon = frmRoster.imageDrop.Picture
                    Select Case intResult
                    Case 0, 1   '[NOT THIS DAY]
                        Msg = ArrayWeek(DayOfWeek(frmRoster.GridRoster.Col - 1)).LongDay & ", " & Format(DsDefault!StartDate + frmRoster.GridRoster.Col - 2, strDateFormat) & " : This staff member is marked as not being available on " & ArrayWeek(DayOfWeek(frmRoster.GridRoster.Col - 1)).LongDay & "s."
                        '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
                        StatusBar Msg
                        '[---------------------------------------------------------------------------------]
                    Case vbInside       '[INSIDE]
                        Msg = ArrayWeek(DayOfWeek(frmRoster.GridRoster.Col - 1)).LongDay & ", " & Format(DsDefault!StartDate + frmRoster.GridRoster.Col - 2, strDateFormat) & " : This staff member is marked as not being available for these hours."
                        '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
                        StatusBar Msg
                        '[---------------------------------------------------------------------------------]
                    Case vbOutside      '[OUTSIDE]
                        Msg = ArrayWeek(DayOfWeek(frmRoster.GridRoster.Col - 1)).LongDay & ", " & Format(DsDefault!StartDate + frmRoster.GridRoster.Col - 2, strDateFormat) & " : This staff member is marked as not being available for these hours."
                        '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
                        StatusBar Msg
                        '[---------------------------------------------------------------------------------]
                    Case vbHoliday      '[HOLIDAYS]
                        Msg = ArrayWeek(DayOfWeek(frmRoster.GridRoster.Col - 1)).LongDay & ", " & Format(DsDefault!StartDate + frmRoster.GridRoster.Col - 2, strDateFormat) & " : This staff member is marked as being on holidays on this day."
                        '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
                        StatusBar Msg
                        '[---------------------------------------------------------------------------------]
                    Case vbNotInClass   '[NOT IN CLASS]
                        Msg = ArrayWeek(DayOfWeek(frmRoster.GridRoster.Col - 1)).LongDay & ", " & Format(DsDefault!StartDate + frmRoster.GridRoster.Col - 2, strDateFormat) & " : This staff member is marked as not being available to this roster class (" & frmRoster.ComboClass.Text & ")."
                        '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
                        StatusBar Msg
                        '[---------------------------------------------------------------------------------]
                    End Select
                Else
                    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
                    StatusBar "Release the mouse button to transfer the selected staff name/s to this roster cell."
                    '[---------------------------------------------------------------------------------]
                    Source.DragIcon = frmRoster.imageDrop.DragIcon
                End If
                intLastCheckRow = intRow
                intLastCheckCol = intCol
            End If
        End If
    End If

End Sub


Private Sub GridRoster_KeyUp(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
    Case vbKeyDelete
        '[CAPTURE DELETE KEY]
        Call cmdDelete_Click
    Case Else
    End Select
    
End Sub

Private Sub GridRoster_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[POPUP COPY/PASTE/DELETE MENU]
    If Button = vbRightButton Then
        '[DISPLAY CONTEXT MENU ON RIGHT MOUSE CLICK]
        PopupMenu mdiMain.mnuRosterEdit, vbPopupMenuRightButton, , , mdiMain.mnuRosterCopy
    End If

End Sub

Private Sub ListStaff_GotFocus()
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Double-click a name to transfer to or remove from the current roster -/- Right-click for a context menu."
    '[---------------------------------------------------------------------------------]

End Sub

Private Sub ListStaff_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    imageDrag.DragIcon = imageDrop.DragIcon
    
    If Button = vbLeftButton And X <> sinLastMouseX And Y <> sinLastMouseY Then
        '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
        StatusBar "Drag the selected name/s to the roster and drop to transfer."
        '[---------------------------------------------------------------------------------]
        If sinLastMouseX = -1 And sinLastMouseY = -1 Then
            sinLastMouseX = X
            sinLastMouseY = Y
        Else
            Dim DY
            DY = imageDrag.Height       '[HEIGHT OF IMAGE]
            imageDrag.Move X, ListStaff.Top + Y - DY / 2, imageDrag.Width, DY
            imageDrag.Drag              '[DRAG IMAGE OUTLINE]
            sinLastMouseX = X
            sinLastMouseY = Y
        End If
    End If
    
End Sub

Private Sub ListStaff_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Double-click a name to transfer to or remove from the current roster -/- Right-click for a context menu."
    '[---------------------------------------------------------------------------------]


End Sub

Private Sub ListStaff_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    '[DECLARE VARIABLES]
    Dim intStaffCounter         As Integer
    Dim intStaffIndex           As Integer
    Dim strFullname         As String

    '[POPUP STAFF INFORMATION SCREEN IF RIGHT MOUSE BUTTON WAS CLICKED]
    If Button = vbRightButton Then
    
        '[SELECT STAFF MEMBER TO PERFORM OPERATION ON]
        intStaffIndex = -1
        '[EXIT IF NO RECORDS]
        If frmRoster.ListStaff.ListCount = 0 Then Exit Sub
        For intStaffCounter = 0 To (frmRoster.ListStaff.ListCount - 1)
            If frmRoster.ListStaff.Selected(intStaffCounter) Then
                intStaffIndex = intStaffCounter
                Exit For
            End If
        Next intStaffCounter
        
        '[SELECT FIRST NAME IN LIST IF NONE ARE SELECTED]
        If intStaffIndex = -1 Then
            intStaffIndex = 0
            frmRoster.ListStaff.Selected(intStaffIndex) = True
        End If
        strFullname = frmRoster.ListStaff.List(intStaffIndex)
        mdiMain.mnuStaffName.Caption = strFullname
        
        '[DISPLAY CONTEXT MENU ON RIGHT MOUSE CLICK]
        PopupMenu mdiMain.mnuStaffList, vbPopupMenuRightButton, , , mdiMain.mnuStaffName

    Else
    
        sinLastMouseX = X
        sinLastMouseY = Y
        intLastCheckRow = -1
        intLastCheckCol = -1

    End If
    
    
End Sub




Private Sub MaskDate_Change()

    '[IF VALID DATE, SAVE TO DSDEFAULT DYNASET]
    If IsDate(frmRoster.MaskDate.Text) Then

        If DsDefault!StartDate = CDate(frmRoster.MaskDate.Text) Then Exit Sub
        DsDefault.Edit
            DsDefault("StartDate") = frmRoster.MaskDate.Text
        DsDefault.Update
        
        '[APPLY DAY LABELS AND RESET DAY-DATES]
        SetDayLabels
        
    End If

End Sub



Private Sub cmdCopy_Click()

    '[ANIMATE BUTTON]
    cmdCopy.BorderStyle = 1
    Delay vbDelay
    cmdCopy.BorderStyle = 0

    '[ADD CLIP AREA OF ROSTER TO strCLIP]
    strClip = CopyCells

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
    Call DeleteCellContents(frmRoster)

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
    If frmRoster.GridRoster.Row = 0 Then
        '[SET ROW TO FIRST ROW]
        intRow = 1
    Else
        intRow = frmRoster.GridRoster.Row + 1
    End If
    
    '[ADD NEW ROW ITEM TO THE GRID]
    frmRoster.GridRoster.AddItem Format(DsDefault("StartTime"), "Medium Time") & Chr$(vbKeyTab) & Format(DsDefault("EndTime"), "Medium Time"), intRow
    '[OLD TIME FORMAT] -> "hh:mm AMPM"
    
    '[MOVE TO THIS NEW ROW]
    frmRoster.GridRoster.Row = intRow
    
    '[MAKE SAVE BUTTON VISIBLE]
    frmRoster.cmdSave.Visible = True

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
    PasteCells
    '[MAKE SAVE BUTTON VISIBLE]
    frmRoster.cmdSave.Visible = True
    
End Sub

Private Sub cmdPaste_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Paste the contents of the clipboard into the selected cells in the roster."
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

    If frmRoster.cmdSave.Visible = True Then
        '[ROSTER IS UNSAVED, POPUP YES/NO DIALOG]
        Msg = "You have made changes to this roster (" & frmRoster.ComboClass.Text & ") but have not saved these changes." & strBreak & strBreak & "If you choose not to save now, any changes you have made since your last save will be lost." & strBreak & strBreak & "Do you wish to save these changes before you rebuild this roster ?"
        Style = vbYesNoCancel ' Define buttons.
        Title = "Roster Changes Not Saved"  ' Define title.
        Response = gsrMsg(Msg, Style, Title)
        If Response = vbYes Then    ' User chose Yes.
            SaveRosterGrid
            '[MAKE SAVE BUTTON VISIBLE]
            frmRoster.cmdSave.Visible = True
        ElseIf Response = vbCancel Then
            Exit Sub
        End If
    Else
        '[REBUILDING ROSTER, POPUP YES/NO DIALOG]
        Msg = "Caution - This action will rebuild this roster using the settings on the Control Form - start time, end time and increment." & strBreak & strBreak & "This will result in this roster being cleared of all data currently entered." & strBreak & strBreak & "Do you wish to continue and rebuild this roster ?"
        Style = vbYesNo             ' Define buttons.
        Title = "Rebuilding Roster"  ' Define title.
        Response = gsrMsg(Msg, Style, Title)
        If Response = vbNo Then    ' User chose No.
            Exit Sub
        End If
    End If

    '[REBUILD ROSTER]
    RebuildRoster
    
    '[MOVE TO ROSTER FORM IF IT IS VISIBLE]
    If frmRoster.WindowState = 1 Then frmRoster.WindowState = 0
    frmRoster.ZOrder
    
    '[MAKE SAVE BUTTON VISIBLE]
    frmRoster.cmdSave.Visible = True

End Sub



Private Sub cmdRebuild_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Rebuild this roster using the defaults specified on the control panel."
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub cmdRemove_Click()

    '[ANIMATE BUTTON]
    cmdRemove.BorderStyle = 1
    Delay vbDelay
    cmdRemove.BorderStyle = 0

    '[REMOVE ALL SELECTED STAFF MEMBERS FROM THE SELECTED CELL/CELLS IN THE ROSTER GRID]
    Dim intCounter          As Integer
    Dim intTrCount          As Integer
    Dim intSelCount         As Integer
    
    '[SELECTION COUNT - ITEMS SELECTED]
    intSelCount = frmRoster.ListStaff.SelCount
    If intSelCount = 0 Then Exit Sub
    intTrCount = 1
    
    For intCounter = 0 To (frmRoster.ListStaff.ListCount - 1)
        If frmRoster.ListStaff.Selected(intCounter) Then
            '[ITEM IS SELECTED SO TRANSFER TO ROSTER]
            RemoveFromRoster (frmRoster.ListStaff.List(intCounter))
            '[TURN SAVE BUTTON ON]
            frmRoster.cmdSave.Visible = True
            '[PROGRESS BAR]
            intTrCount = intTrCount + 1
        End If
    Next intCounter

End Sub



Private Sub cmdRemove_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Remove the selected staff member(s) from the selected cell(s) in the roster."
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
    
    If frmRoster.GridRoster.Row = 0 Then Exit Sub
    intCol = frmRoster.GridRoster.Col
    
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
        intRow = frmRoster.GridRoster.Row
        If frmRoster.GridRoster.Rows = 2 Then
            For intCounter = 0 To 8
                frmRoster.GridRoster.Col = intCounter
                frmRoster.GridRoster.Text = ""
            Next intCounter
        Else
            frmRoster.GridRoster.RemoveItem intRow
        End If
        '[MOVE TO A NEW ROW]
        If intRow = frmRoster.GridRoster.Rows Then frmRoster.GridRoster.Row = intRow - 1
        '[MAKE SAVE BUTTON VISIBLE]
        frmRoster.cmdSave.Visible = True
    End If

    frmRoster.GridRoster.Col = intCol


End Sub

Private Sub cmdRemoveRow_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Remove the selected row from the roster grid."
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub cmdSave_Click()

    '[ANIMATE BUTTON]
    cmdSave.BorderStyle = 1
    Delay vbDelay
    cmdSave.BorderStyle = 0

    '[SAVE DATA IN ROSTER GRID TO DYNASET]
    SaveRosterGrid
    
    '[HIDE SAVE BUTTON]
    frmRoster.cmdSave.Visible = False

End Sub


Private Sub cmdSave_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Save any changes made to this roster."
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub cmdTransfer_Click()

    '[ANIMATE BUTTON]
    cmdTransfer.BorderStyle = 1
    Delay vbDelay
    cmdTransfer.BorderStyle = 0

    '[TRANSFER ALL SELECTED STAFF MEMBERS TO THE SELECTED CELL/CELLS IN THE ROSTER GRID]
    Dim intCounter          As Integer
    Dim intTrCount          As Integer
    Dim intSelCount         As Integer
    Dim flagResult          As Boolean
    
    '[SELECTION COUNT - ITEMS SELECTED]
    intSelCount = frmRoster.ListStaff.SelCount
    If intSelCount = 0 Then Exit Sub
    intTrCount = 1
    flagContinue = vbOK
    
    For intCounter = 0 To (frmRoster.ListStaff.ListCount - 1)
        If frmRoster.ListStaff.Selected(intCounter) Then
            '[ITEM IS SELECTED SO TRANSFER TO ROSTER]
            Call TransferToRoster(frmRoster.ListStaff.List(intCounter), True, flagResult, flagContinue)
            '[PROGRESS BAR]
            intTrCount = intTrCount + 1
            '[EXIT IF CANCEL PRESSED]
            If flagContinue = vbCancel Then Exit Sub
        End If
    Next intCounter
    
End Sub


Private Sub cmdTransfer_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Transfer the selected staff member(s) to the selected cell(s) in the roster."
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub comboClass_Click()
    
    '[IF FORM HAS NOT YET LOADED, EXIT]
    If flagLoaded = False Then Exit Sub
    
    '[CHECK FOR CHANGED DATA AND NOTIFY]
    If frmRoster.cmdSave.Visible = True Then
        '[ROSTER IS UNSAVED, POPUP YES/NO DIALOG]
        Dim Msg As String
        Dim Style
        Dim Response
        Dim Title
        
        Msg = "You have made changes to this roster (" & frmRoster.ComboClass.Text & ") but have not saved these changes." & strBreak & strBreak & "If you choose not to save now, any changes you have made since your last save will be lost." & strBreak & strBreak & "Do you wish to save these changes before you move to another roster ?"
        Style = vbYesNoCancel ' Define buttons.
        Title = "Roster Changes Not Saved"  ' Define title.
        Response = gsrMsg(Msg, Style, Title)
        If Response = vbYes Then    ' User chose Yes.
            SaveRosterGrid
        ElseIf Response = vbCancel Then
            Exit Sub
        End If
    End If
    
    '[IF BLANK, MOVE TO FIRST ITEM IN LIST]
    If frmRoster.ComboClass = "" Or frmRoster.ComboClass.ListIndex = -1 Then frmRoster.ComboClass.ListIndex = 0
    
    '[HIDE SAVE BUTTON]
    frmRoster.cmdSave.Visible = False
    
    '[LOCATE CLASS IN DSCLASS DYNASET]
    LocateClass (frmRoster.ComboClass.Text)
    
    '[APPLY THE NEW CLASS ID TO THE PUBLIC VARIABLE]
    intRosterClass = frmRoster.ComboClass.ItemData(frmRoster.ComboClass.ListIndex)
    
    '[CHANGE CAPTION ON FORM TO CLASS NAME]
    frmRoster.Caption = "Roster - " & frmRoster.ComboClass.Text
    
    '[FILL STAFF ROSTER LIST WITH THOSE STAFF THAT MATCH]
    FillStaffRosterList
    
    '[REBUILD DSROSTER DYNASET WITH DATA FROM DATABASE]
    BuildRosterDynaset (intRosterClass)
    
    '[FILL GRID WITH ROSTER FROM DSROSTER ARRAY]
    FillRosterGrid
    
    '[PLACE ROSTER DATE IN START DATE BOX]
    frmRoster.MaskDate.Text = Format(DsDefault("StartDate"), strDateFormat)

    '[SELECT FIRST ROW AND COLUMN OF THE ROSTER GRID]
    frmRoster.GridRoster.Row = 1
    frmRoster.GridRoster.Col = 2
    frmRoster.GridRoster.RowSel = 1 '.SelStartRow = 1
    frmRoster.GridRoster.ColSel = 2 '.SelStartCol = 2
    'frmRoster.GridRoster.SelEndRow = 1
    'frmRoster.GridRoster.SelEndCol = 2
    
    
End Sub


Private Sub ComboClass_GotFocus()
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "List of available rosters (those checked as active on the settings form)."
    '[---------------------------------------------------------------------------------]

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
    If frmRoster.WindowState = 1 Then Exit Sub
    
    '[RESIZE GRID AND LIST BOXES AND ARRANGE CONTROLS ON FORM]
    
    '[STAFF LIST]
    sinHeight = frmRoster.Height - frmRoster.ListStaff.Top - PanelToolBar.Height - 180
    If sinHeight > 0 Then frmRoster.ListStaff.Height = sinHeight
    
    '[ROSTER GRID]
    sinWidth = frmRoster.Width - (frmRoster.ListStaff.Width + 200)
    If sinWidth > 0 Then frmRoster.GridRoster.Width = sinWidth
    sinHeight = frmRoster.Height - frmRoster.GridRoster.Top - 500
    If sinHeight > 0 Then frmRoster.GridRoster.Height = sinHeight
    
End Sub


Private Sub GridRoster_DblClick()
    
    '[ALLOW USER TO CHANGE THE SHIFT START/END TIME]
    Dim Msg As String
    Dim Style
    Dim Response
    Dim Title

    If frmRoster.GridRoster.Row = 0 Then Exit Sub
    
    Select Case frmRoster.GridRoster.Col
    Case 0, 1     '[SHIFT START TIME/END TIME]
        Load frmTime
        '[SET TIME ON FORM]
        If frmRoster.GridRoster.Text = "" Then
            If frmRoster.GridRoster.Col = 0 Then
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
            If IsDate(frmRoster.GridRoster.Text) Then
                frmTime.ComboHour = Format(Hour(CDate(frmRoster.GridRoster.Text)), "0#")
                frmTime.ComboMinute = Format(Minute(CDate(frmRoster.GridRoster.Text)), "0#")
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
                frmRoster.GridRoster.Text = Format(frmTime.ComboHour.Text & ":" & frmTime.ComboMinute.Text, "Medium Time")
                '[OLD FORMAT COMMAND] -> "hh:mm AMPM"
                '[CHECK FOR CELL SIZE AND RESIZE IF NECESSARY]
                '[ALLOW 10% MARGIN FOR TEXT ADJUSTMENT]
                If frmRoster.GridRoster.ColWidth(frmRoster.GridRoster.Col) < (TextWidth(frmRoster.GridRoster.Text) * 1.1) Then frmRoster.GridRoster.ColWidth(frmRoster.GridRoster.Col) = (TextWidth(frmRoster.GridRoster.Text) * 1.1)
                '[MAKE SAVE BUTTON VISIBLE]
                frmRoster.cmdSave.Visible = True
            End If
        End If
        '[REMOVE FORM]
        'Unload frmTime

    Case Else   '[SHOW FULL CELL CONTENTS]
        If Len(Trim(frmRoster.GridRoster.Text)) = 0 Then Exit Sub
        Msg = frmRoster.GridRoster.Text
        Style = vbOKOnly                        ' Define buttons.
        Title = "Cell Contents"                 ' Define title.
        Response = gsrMsg(Msg, Style, Title)
    
    End Select

End Sub

Private Sub GridRoster_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    If X > (frmRoster.GridRoster.ColPos(2)) Then
        StatusBar "Double-click to display the names which appear in this cell."
    Else
        '[DISPLAY DIFFERENT MESSAGES FOR LOCKED OR UNLOCKED]
        If frmRoster.GridRoster.FixedCols = 0 Then
            StatusBar "Double-click to modify the shift start/finish time."
        Else
            StatusBar "Select 'Unlock Roster Columns' from the 'Options' menu to modify the Start/Finish times."
        End If
    End If
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub ListStaff_Click()

    '[CHECK NUMBER OF SELECTED ITEMS AND CHANGE ICON ON TRANSFER BUTTON]
    If frmRoster.ListStaff.SelCount = 0 Then
        frmRoster.cmdTransfer.Enabled = False
    Else
        frmRoster.cmdTransfer.Enabled = True
        '[DECLARATIONS]
        Dim intCounter          As Integer
        If frmRoster.ListStaff.ListCount = 0 Then Exit Sub
        
        '[RELOCATE STAFF NAME]
        For intCounter = 0 To (frmSet.ListStaff.ListCount - 1)
            If frmSet.ListStaff.List(intCounter) = frmRoster.ListStaff.Text Then
                '[NAME FOUND]
                frmSet.tabSet.Tab = 1
                frmSet.ListStaff.ListIndex = intCounter
                Exit For
            End If
        Next intCounter
    End If

End Sub


Private Sub ListStaff_DblClick()

    '[TRANSFER STAFF MEMBER ON DOUBLE CLICK -IF- NAME ISN'T IN CELL]
    If InStr(frmRoster.GridRoster.Text, frmRoster.ListStaff.List(frmRoster.ListStaff.ListIndex)) > 0 Then
        '[NAME IS IN CELL, REMOVE NAME]
        cmdRemove_Click
    Else
        '[NAME ISN'T IN CELL, ADD NAME]
        cmdTransfer_Click
    End If

End Sub


Private Sub MaskDate_GotFocus()
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Set the roster starting date."
    '[---------------------------------------------------------------------------------]

End Sub


