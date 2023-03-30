VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.MDIForm mdiMain 
   Appearance      =   0  'Flat
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000F&
   Caption         =   "Generic Staff Roster"
   ClientHeight    =   6345
   ClientLeft      =   2610
   ClientTop       =   2805
   ClientWidth     =   9360
   HelpContextID   =   20
   Icon            =   "MDIMAIN.frx":0000
   LinkTopic       =   "mdiMain"
   LockControls    =   -1  'True
   Begin Threed.SSPanel PanelToolBar 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      Negotiate       =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   9360
      _Version        =   65536
      _ExtentX        =   16510
      _ExtentY        =   635
      _StockProps     =   15
      ForeColor       =   0
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   7.5
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
      MouseIcon       =   "MDIMAIN.frx":0CFA
      Begin VB.Image cmdRosterReport 
         Height          =   300
         Left            =   2460
         Picture         =   "MDIMAIN.frx":15D4
         Top             =   30
         Width           =   300
      End
      Begin VB.Image cmdAllStaffRosters 
         Height          =   300
         Left            =   3900
         Picture         =   "MDIMAIN.frx":1746
         Top             =   30
         Width           =   300
      End
      Begin VB.Image cmdSelectedStaffRoster 
         Height          =   300
         Left            =   3600
         Picture         =   "MDIMAIN.frx":1998
         Top             =   30
         Width           =   300
      End
      Begin VB.Image cmdPrintMaster 
         Height          =   300
         Left            =   1560
         Picture         =   "MDIMAIN.frx":1B0A
         Top             =   30
         Width           =   300
      End
      Begin VB.Image cmdGraph 
         Height          =   300
         Left            =   3060
         Picture         =   "MDIMAIN.frx":1C5C
         Top             =   30
         Width           =   300
      End
      Begin VB.Image cmdStaffReport 
         Height          =   300
         Left            =   2760
         Picture         =   "MDIMAIN.frx":1DAE
         Top             =   30
         Width           =   300
      End
      Begin VB.Image cmdExceptionReport 
         Height          =   300
         Left            =   2160
         Picture         =   "MDIMAIN.frx":2380
         Top             =   30
         Width           =   300
      End
      Begin VB.Image cmdControl 
         Height          =   300
         Left            =   60
         Picture         =   "MDIMAIN.frx":24D2
         Top             =   30
         Width           =   300
      End
      Begin VB.Image cmdPrint 
         Height          =   300
         Left            =   1260
         Picture         =   "MDIMAIN.frx":2644
         Top             =   30
         Width           =   300
      End
      Begin VB.Image cmdFonts 
         Height          =   300
         Left            =   660
         Picture         =   "MDIMAIN.frx":2796
         Top             =   30
         Width           =   300
      End
   End
   Begin Threed.SSPanel panelStatusBar 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   1
      Top             =   6045
      Width           =   9360
      _Version        =   65536
      _ExtentX        =   16510
      _ExtentY        =   529
      _StockProps     =   15
      Caption         =   " Status Bar Message Area"
      ForeColor       =   -2147483634
      BackColor       =   -2147483635
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   0
      RoundedCorners  =   0   'False
      FloodColor      =   -2147483648
      FloodShowPct    =   0   'False
      Alignment       =   1
   End
   Begin Crystal.CrystalReport Report 
      Left            =   480
      Top             =   360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      WindowLeft      =   0
      WindowTop       =   0
      WindowWidth     =   640
      WindowHeight    =   480
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
   End
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   0
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      FontName        =   "Arial"
      HelpContext     =   10
      HelpFile        =   "gsr.hlp"
      HelpKey         =   "F1"
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuLoadFromFile 
         Caption         =   "&Load Roster from File"
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuSavetoFile 
         Caption         =   "&Save Roster to File"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSpacer1 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnuPrintSetup 
         Caption         =   "Pri&nter Setup"
      End
      Begin VB.Menu mnuPrintGrid 
         Caption         =   "&Print Current Roster"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuMasterRoster 
         Caption         =   "Print &Master Roster"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnuFileSpacer2 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnuFile_Quit 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuFile_Backup 
         Caption         =   "Exit and &Backup"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuSettings 
         Caption         =   "&Control Panel"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuScratch 
         Caption         =   "&Scratch Roster"
      End
      Begin VB.Menu mnuOptionsSpace1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuToolbarState 
         Caption         =   "Hide &Toolbar"
      End
      Begin VB.Menu mnuStatusBarState 
         Caption         =   "Hide St&atus Bar"
      End
      Begin VB.Menu mnuLockRoster 
         Caption         =   "&Lock Roster Columns"
      End
   End
   Begin VB.Menu mnuReportS 
      Caption         =   "&Reports"
      Begin VB.Menu mnuException 
         Caption         =   "E&xception"
      End
      Begin VB.Menu mnuRosterRpt 
         Caption         =   "&Roster Details"
      End
      Begin VB.Menu mnuStaffRpt 
         Caption         =   "&Staff Details"
      End
      Begin VB.Menu mnuGraph 
         Caption         =   "Roster &Breakdown"
      End
      Begin VB.Menu mnuReportSpacer1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuStaffListGeneral 
         Caption         =   "General Staff List"
      End
      Begin VB.Menu mnuStaffListDetailed 
         Caption         =   "Detailed Staff List"
      End
      Begin VB.Menu mnuReportSpacer2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSingleTimesheet 
         Caption         =   "S&ingle Timesheet"
      End
      Begin VB.Menu mnuAllTimesheets 
         Caption         =   "&Multiple  Timesheets"
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      WindowList      =   -1  'True
      Begin VB.Menu mnuMaximise 
         Caption         =   "&Maximise"
      End
      Begin VB.Menu mnuRestore 
         Caption         =   "&Restore"
      End
      Begin VB.Menu mnuCascade 
         Caption         =   "&Cascade"
      End
      Begin VB.Menu mnuTileHor 
         Caption         =   "Tile Hori&zontally"
      End
      Begin VB.Menu mnuTileVer 
         Caption         =   "Tile &Vertically"
      End
      Begin VB.Menu mnuWindowSpacer 
         Caption         =   "-"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpContents 
         Caption         =   "Contents"
         HelpContextID   =   10
         Shortcut        =   ^{F1}
      End
      Begin VB.Menu mnuHelp_About 
         Caption         =   "&About"
      End
      Begin VB.Menu mnuGuide 
         Caption         =   "&Quick Guide"
      End
      Begin VB.Menu mnuRegister 
         Caption         =   "Re&gister"
      End
   End
   Begin VB.Menu mnuStaffList 
      Caption         =   "&StaffList"
      NegotiatePosition=   3  'Right
      Visible         =   0   'False
      Begin VB.Menu mnuStaffName 
         Caption         =   "StaffName"
      End
      Begin VB.Menu mnuNameSpacer 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAddStaff 
         Caption         =   "Transfer to Roster"
      End
      Begin VB.Menu mnuRemoveStaff 
         Caption         =   "Remove from Roster"
      End
   End
   Begin VB.Menu mnuRosterEdit 
      Caption         =   "&RosterEdit"
      NegotiatePosition=   3  'Right
      Visible         =   0   'False
      Begin VB.Menu mnuRosterCopy 
         Caption         =   "Copy"
      End
      Begin VB.Menu mnuRosterPaste 
         Caption         =   "Paste"
      End
      Begin VB.Menu mnuRosterEditSpacer 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRosterDelete 
         Caption         =   "Delete"
      End
      Begin VB.Menu mnuRosterEditSpacer2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRosterPick 
         Caption         =   "Pick Staff"
      End
   End
   Begin VB.Menu mnuScratchEdit 
      Caption         =   "&ScratchEdit"
      NegotiatePosition=   3  'Right
      Visible         =   0   'False
      Begin VB.Menu mnuScratchCopy 
         Caption         =   "Copy"
      End
      Begin VB.Menu mnuScratchPaste 
         Caption         =   "Paste"
      End
      Begin VB.Menu mnuScratchEditSpacer 
         Caption         =   "-"
      End
      Begin VB.Menu mnuScratchDelete 
         Caption         =   "Delete"
      End
   End
End
Attribute VB_Name = "mdiMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'[----------------------------------------------]
'[mdiMain           Main Parent Form            ]
'[----------------------------------------------]
'[GSR                                           ]
'[generic staff roster system for Windows       ]
'[(C) ---------- David Gilbert, 1997            ]
'[----------------------------------------------]
'[version                           3.00.xx     ]
'[initial date                      21/07/1997  ]
'[----------------------------------------------]
'[last revision                     31          ]
'[                                  03/07/98    ]
'[----------------------------------------------]








Private Sub cmdAllStaffRosters_Click()
    
    '[ANIMATE BUTTON]
    cmdAllStaffRosters.BorderStyle = 1
    Delay vbDelay
    cmdAllStaffRosters.BorderStyle = 0
    
    '[CALL ROUTINE TO PROCESS ALL STAFF RECORDS AND PRINT TIME SHEETS]
    Call procAllStaffRosters

End Sub

Private Sub cmdAllStaffRosters_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Print the weekly roster for all staff members in the list who have been allocated to rosters."
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub cmdControl_Click()

    '[ANIMATE BUTTON]
    cmdControl.BorderStyle = 1
    Delay vbDelay
    cmdControl.BorderStyle = 0
    
    '[SHOW THE SETTINGS FORM NONMODAL]
    frmSet.Show

End Sub

Private Sub cmdControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Display the GSR Control Panel."
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub cmdExceptionReport_Click()
    
    '[ANIMATE BUTTON]
    cmdExceptionReport.BorderStyle = 1
    Delay vbDelay
    cmdExceptionReport.BorderStyle = 0

    '[CALL SUBROUTINE WHICH PROCESSES THE EXCEPTION REPORT]
    Call procExceptionReport
    
End Sub



Private Sub cmdExceptionReport_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Produce the Exception Report."
    '[---------------------------------------------------------------------------------]

End Sub

Private Sub cmdFonts_Click()

    '[ANIMATE BUTTON]
    cmdFonts.BorderStyle = 1
    Delay vbDelay
    cmdFonts.BorderStyle = 0

    '[ERROR HANDLER]
    On Error Resume Next
    
    '[GET CURRENT FONTS FROM ROSTER FORM]
    mdiMain.CommonDialog.FontName = frmRoster.GridRoster.Font.Name
    mdiMain.CommonDialog.FontSize = frmRoster.GridRoster.Font.Size
    mdiMain.CommonDialog.FontBold = frmRoster.GridRoster.Font.Bold
    mdiMain.CommonDialog.FontItalic = frmRoster.GridRoster.Font.Italic
    
    '[SET FONT DIALOG TO CURRENT OPTIONS]
    
    '[SHOW FONT DIALOG]
    mdiMain.CommonDialog.Flags = cdlCFPrinterFonts
    mdiMain.CommonDialog.ShowFont
    
    '[APPLY FONT TO CURRENT FORM FOR GRID RESIZING]
    ActiveForm.Font.Name = mdiMain.CommonDialog.FontName
    ActiveForm.Font.Size = mdiMain.CommonDialog.FontSize
    ActiveForm.Font.Bold = mdiMain.CommonDialog.FontBold
    ActiveForm.Font.Italic = mdiMain.CommonDialog.FontItalic
    
    '[APPLY TO ALL OBJECTS ON THE FORM]
    ActiveForm.ActiveControl.Font.Name = mdiMain.CommonDialog.FontName
    ActiveForm.ActiveControl.Font.Size = mdiMain.CommonDialog.FontSize
    ActiveForm.ActiveControl.Font.Bold = mdiMain.CommonDialog.FontBold
    ActiveForm.ActiveControl.Font.Italic = mdiMain.CommonDialog.FontItalic
    
    '[NOW CHECK FOR FIRST ROW OF REPORT GRIDS]
    If ActiveForm.Name = "frmRoster" Then frmRoster.GridRoster.RowHeight(0) = frmRoster.TextHeight("A")
    
End Sub

Private Sub cmdFonts_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Change the display font."
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub GotoRoster()
    
    '[MOVE ROSTER FORM TO THE FRONT]
    If frmRoster.WindowState = 1 Then frmRoster.WindowState = 0
    frmRoster.ZOrder

End Sub



Private Sub cmdGraph_Click()

    '[ANIMATE BUTTON]
    cmdGraph.BorderStyle = 1
    Delay vbDelay
    cmdGraph.BorderStyle = 0

    '[PRODUCE THE ROSTER BREAKDOWN REPORT]
    ProcRosterBreakdown
        
End Sub

Private Sub cmdGraph_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Produce the Roster Breakdown Report."
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub cmdPrint_Click()

    '[ANIMATE BUTTON]
    cmdPrint.BorderStyle = 1
    Delay vbDelay
    cmdPrint.BorderStyle = 0

    '[CALL PROCEEDURE TO PRINT CURRENT GRID]
    Call procCurrentRoster
    
End Sub

Private Sub cmdPrint_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Print the " & frmRoster.ComboClass.Text & " Roster."
    '[---------------------------------------------------------------------------------]

End Sub









Private Sub cmdPrintMaster_Click()

    '[ANIMATE BUTTON]
    cmdPrintMaster.BorderStyle = 1
    Delay vbDelay
    cmdPrintMaster.BorderStyle = 0
    
    '[CALL PROCEEDURE TO PRINT MASTER ROSTER]
    Call procMasterRoster

End Sub

Private Sub cmdPrintMaster_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Print the Master Roster (all active rosters combined)."
    '[---------------------------------------------------------------------------------]

End Sub




Private Sub cmdRosterReport_Click()

    '[ANIMATE BUTTON]
    cmdRosterReport.BorderStyle = 1
    Delay vbDelay
    cmdRosterReport.BorderStyle = 0

    '[CALL SUBROUTINE WHICH PROCESSES THE Roster Details Report]
    Call procRosterReport
    
End Sub

Private Sub cmdRosterReport_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Produce the Roster Details Report."
    '[---------------------------------------------------------------------------------]

End Sub




Private Sub cmdRosterRep_Click()

End Sub

Private Sub cmdSelectedStaffRoster_Click()

    '[ANIMATE BUTTON]
    cmdSelectedStaffRoster.BorderStyle = 1
    Delay vbDelay
    cmdSelectedStaffRoster.BorderStyle = 0


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

    '[CALL ROUTINE TO PROCESS SINGLE STAFF RECORD AND PRINT TIME SHEET]
    Call procSelectedStaffRoster

End Sub

Private Sub cmdSelectedStaffRoster_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    If frmRoster.ListStaff.ListCount = 0 Then
        StatusBar "Click here to print the weekly roster for the selected staff member."
    Else
        StatusBar "Print the weekly roster for " & frmRoster.ListStaff.List(frmRoster.ListStaff.ListIndex) & "."
    End If
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub cmdStaffReport_Click()

    '[ANIMATE BUTTON]
    cmdStaffReport.BorderStyle = 1
    Delay vbDelay
    cmdStaffReport.BorderStyle = 0

    '[PROCESS THE STAFF REPORT]
    procStaffReport

End Sub

Private Sub cmdStaffReport_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Produce the Staff Details Report."
    '[---------------------------------------------------------------------------------]

End Sub



Private Sub MDIForm_Load()
    
    '[center form on screen]
    Top = (Screen.Height / 2) - (Height / 2)
    Left = (Screen.Width / 2) - (Width / 2)
   
    '[DEBUG]
    Call LogToFile("=-MAIN FORM - LOAD SUB-------------------=")
    
    '[initialise dynasets and other variables]
    Call Initialise
    
    '[DEBUG]
    Call LogToFile("setting toolbar state")
    '[SET TOOLBAR STATE]
    If DsDefault("ToolBarState").Value = 0 Then
        '[HIDE TOOLBAR]
        Call mnuToolbarState_Click
    End If
    
    '[DEBUG]
    Call LogToFile("setting statusbar state")
    '[SET STATUSBAR STATE]
    If DsDefault("StatusBarState") = 0 Then
        '[HIDE STATUSBAR]
        Call mnuStatusBarState_Click
    End If
    
    '[DEBUG]
    Call LogToFile("setting roster column state")
    '[SET ROSTER COLUMN STATE]
    If DsDefault("RosterLocked") > 0 Then
        '[LOCK ROSTER]
        Call mnuLockRoster_Click
    End If
    
End Sub

Private Sub MDIForm_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Information on various controls and the general use of GSR can be found here !"
    '[---------------------------------------------------------------------------------]
    
End Sub


Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    '[CALL TERMINATE ROUTINE]
    Dim Response
    If flagTerminate = False Then Call Terminate(Response)
    If Response = vbCancel Then Cancel = True

End Sub

Private Sub mnuAddStaff_Click()

        '[DECLARATIONS]
        Dim flagResult          As Boolean
        Dim flagContinue        As Boolean

        '[TRANSFER HIGHLIGHTED NAME TO ROSTER]
        Call TransferToRoster(mdiMain.mnuStaffName.Caption, True, flagResult, flagContinue)
        
End Sub

Private Sub mnuAllTimesheets_Click()
    
    '[CALL ROUTINE TO PROCESS ALL STAFF RECORDS AND PRINT TIME SHEETS]
    Call procAllStaffRosters

End Sub

Private Sub mnuCascade_Click()

    '[ARRANGE VISIBLE WINDOWS IN CASCADE FORMAT]
    mdiMain.Arrange vbCascade

End Sub


Private Sub mnuException_Click()
    
    '[CALL SUBROUTINE WHICH PROCESSES THE EXCEPTION REPORT]
    Call procExceptionReport

End Sub

Private Sub mnuFile_Backup_Click()

    '[CALL TERMINATE SUBROUTINE]
    Dim Response
    flagTerminate = True
    flagBackup = True
    Call Terminate(Response)

End Sub

Private Sub mnuFile_Quit_Click()

    '[CALL TERMINATE SUBROUTINE]
    Dim Response
    flagTerminate = True
    flagBackup = False
    Call Terminate(Response)
    
End Sub


Private Sub mnuGraph_Click()

    '[PRODUCE BREAKDOWN GRAPH]
    ProcRosterBreakdown
    
End Sub

Private Sub mnuGuide_Click()

    '[DISPLAY THE USER'S GUIDE]
    Dim intOldContext       As Integer
    intOldContext = ActiveForm.ActiveControl.HelpContextID
    ActiveForm.ActiveControl.HelpContextID = 100
    SendKeys ("{F1}"), True
    ActiveForm.ActiveControl.HelpContextID = intOldContext
    
End Sub

Private Sub mnuHelp_About_Click()

    '[load fmrAbout and set labelRegUser to the Registered User Value in the database]
    Load frmAbout
    frmAbout.labelRegUser.Caption = DsDefault("RegUser")
    frmAbout.TextEmail.Text = constEmail
    frmAbout.TextWeb.Text = constWeb
    
    frmAbout.LabelVersion.Caption = constVersion
    '[BETA VERSION]
    If flagBeta = True Then frmAbout.LabelVersion.Caption = frmAbout.LabelVersion.Caption + " beta"
    
    frmAbout.labelInfo.Caption = "A Generic Staff Rostering software system for small businesses."
    frmAbout.labelInfo.Caption = frmAbout.labelInfo.Caption & strBreak & strBreak & "Generic Staff Roster was designed to assist small business owners/operators in allocating their available human resources easily and efficiently."
    If DsDefault!RegCode = "" Or IsNull(DsDefault!RegCode) Then frmAbout.labelInfo.Caption = frmAbout.labelInfo.Caption & strBreak & strBreak & "For more information on registering Generic Staff Roster, see the help file section entitled 'Registering GSR'."
    frmAbout.ImageInfo.Tag = 0
    frmAbout.Show 1
    mdiMain.Show
    
End Sub








Private Sub mnuHelpContents_Click()
    
    '[DISPLAY THE CONTENTS PAGE]
    Dim intOldContext       As Integer
    intOldContext = ActiveForm.ActiveControl.HelpContextID
    ActiveForm.ActiveControl.HelpContextID = 10
    SendKeys ("{F1}"), True
    ActiveForm.ActiveControl.HelpContextID = intOldContext

End Sub

Private Sub mnuLoadFromFile_Click()

    '[SET CALL TO ERRORHANDLING ROUTINE]
    On Error GoTo ErrorHandler
    Dim intFileHandle           As Integer
    
    '[DETERMINE FORM AND FILENAME TO USE TO LOAD ROSTER FILE]
    If ActiveForm.Name = "frmScratch" Then
        flagScratch = True
    Else
        flagScratch = False
    End If
    
    '[SET DEFAULTS FOR FILE DIALOG]
    FileSetFilter
    
    '[DIALOG BOX TO OPEN A .FPL FIELD PLAN FILE]
    mdiMain.CommonDialog.Flags = cdlOFNFileMustExist Or cdlOFNPathMustExist
    If flagScratch = True Then
        mdiMain.CommonDialog.DialogTitle = "Open A Scratch Roster File"
    Else
        mdiMain.CommonDialog.DialogTitle = "Open A Roster File"
    End If
    mdiMain.CommonDialog.CancelError = True
    mdiMain.CommonDialog.ShowOpen

    Select Case mdiMain.CommonDialog.FileName
        Case "" '[NO FILE SELECTED]
            '[NO CHANGES TO STATUS QUO]
        Case Else '[OPEN SELECTED FILE]
            If flagScratch = True Then
                strScratchFile = mdiMain.CommonDialog.FileName
            Else
                strRosterFile = mdiMain.CommonDialog.FileName
            End If
            Call FileRead(mdiMain.CommonDialog.FileName)
            frmRoster.GridRoster.Row = 1
            frmRoster.GridRoster.Col = 0
    End Select

ErrorHandler:
    If Err.Number > 0 Then
        '[CANCEL WAS PRESSED ON THE SAVE FORM - NO PROCESSING REQUIRED]
    End If


End Sub

Private Sub mnuLockRoster_Click()

    '[LOCK/UNLOCK THE FIRST TWO ROSTER COLUMNS]
    Select Case mnuLockRoster.Caption
    Case "&Lock Roster Columns"
        frmRoster.GridRoster.FixedCols = 2
        mnuLockRoster.Caption = "Un&lock Roster Columns"
    Case "Un&lock Roster Columns"
        frmRoster.GridRoster.FixedCols = 0
        mnuLockRoster.Caption = "&Lock Roster Columns"
    End Select

End Sub

Private Sub mnuMasterRoster_Click()
    
    '[REV: 3.00.28]
    '[CALL PROCEEDURE TO PRINT MASTER ROSTER]
    Call procMasterRoster
    
End Sub

Private Sub mnuMaximise_Click()

    '[MAXIMISE CURRENT WINDOW]
    On Error Resume Next
    ActiveForm.WindowState = 2

End Sub

Private Sub mnuPrintGrid_Click()

    '[CALL PROCEEDURE TO PRINT CURRENTLY HIGHLIGHTED GRID]
    Call procCurrentRoster

End Sub

Private Sub mnuPrintSetup_Click()

    '[LOAD PRINTER DIALOG]
    '[MESSAGE VARIABLES]
    Dim Msg
    Dim Style
    Dim Title
    Dim Response
    '[CHECK TO SEE IF A PRINTER IS ATTACHED]
    If Printers.Count = 0 Then
        Msg = "There is no default printer attached to this computer." & strBreak & strBreak & "GSR cannot change the printer settings."
        Style = vbOKOnly                     ' Define buttons.
        Title = "No Printer Attached"
        Response = gsrMsg(Msg, Style, Title)
        Exit Sub
    End If
  
    mdiMain.CommonDialog.Flags = cdlPDPrintSetup    '[JUST DISPLAY THE PRINT SETUP BOX]
    mdiMain.CommonDialog.ShowPrinter

End Sub

Private Sub mnuRegister_Click()

    '[DECLARE VARIABLES]
    Dim Msg
    Dim Style
    Dim Title
    Dim Response

    '[SHOW REGISTER FORM AND VALIDATE REGISTRATION CODE]
    Load frmRegister
    '[PLACE VALUES IN FORM TEXT BOXES]
    frmRegister.txtRegUser = DsDefault("RegUser")
    frmRegister.txtRegCode = DsDefault("RegCode")
    
    '[SHOW FORM]
    frmRegister.Show 1
    
    '[PROCESS RESULT OF FORM]
    If frmRegister.CheckResult = vbChecked Then
        If Validate(frmRegister.txtRegUser) = frmRegister.txtRegCode And frmRegister.txtRegCode > "" Then
            '[CODE MATCHES, PLACE NEW NAME AND CODE INTO THE DATABASE]
            If DsDefault("RegUser") <> frmRegister.txtRegUser And DsDefault("RegCode") <> frmRegister.txtRegCode Then
                DsDefault.Edit
                    DsDefault("RegUser") = frmRegister.txtRegUser
                    DsDefault("RegCode") = frmRegister.txtRegCode
                DsDefault.Update
                '[SHOW THANKYOU FORM]
                frmThank.Show 1
            End If
        Else
            '[CODE DOESN'T MATCH, PLACE UNREGISTERED DETAILS INTO DATABASE]
            DsDefault.Edit
                DsDefault("RegUser") = "Unregistered Version"
                DsDefault("RegCode") = ""
            DsDefault.Update
            
            '[SHOW MESSAGE FORM]
            Msg = "Your registration validation code does not match your registered user name.  Please check both the name and the validation code and re-enter." & strBreak & strBreak & "Name : " & frmRegister.txtRegUser & strBreak & strBreak & "Code : " & frmRegister.txtRegCode
            Style = vbOKOnly                     ' Define buttons.
            Title = "Registration Incorrect"
            Response = gsrMsg(Msg, Style, Title)
        End If
    End If
    
    '[REMOVE FORM]
    Unload frmRegister

End Sub

Private Sub mnuRemoveStaff_Click()

    '[REMOVE STAFF FROM ROSTER]
    Call RemoveFromRoster(mdiMain.mnuStaffName.Caption)

End Sub

Private Sub mnuRestore_Click()

    '[RESTORE CURRENT WINDOW]
    On Error Resume Next
    ActiveForm.WindowState = 0

End Sub




Private Sub mnuRosterCopy_Click()

    '[CALL COPY ROUTINES]
    '[ADD CLIP AREA OF ROSTER TO strCLIP]
    strClip = CopyCells


End Sub

Private Sub mnuRosterDelete_Click()

    '[CLEAR SELECTED CELLS IN ROSTER]
    Call DeleteCellContents(frmRoster)
    
End Sub


Private Sub mnuRosterPaste_Click()
    
    '[CALL PASTE ROUTINES]
    '[PUT strClip IN ROSTER]
    If Len(strClip) = 0 Then Exit Sub
    PasteCells
    '[MAKE SAVE BUTTON VISIBLE]
    frmRoster.cmdSave.Visible = True

End Sub

Private Sub mnuRosterPick_Click()

    '[PICK STAFF TO FILL THE SELECTED ROSTER SPOT]
    Call procPickStaff

End Sub

Private Sub mnuRosterRpt_Click()

    '[CALL SUBROUTINE WHICH PROCESSES THE Roster Details Report]
    Call procRosterReport

End Sub

Private Sub mnuSavetoFile_Click()

    '[SAVE THIS ROSTER TO A FILE]
    '[FILENAME IS PREALLOCATED BUT MAY BE CHANGED]
    Dim strSaveFile             As String
    Dim intClassCounter         As Integer
    Dim strBookmark             As String
    Dim intSlashFound           As Integer
    Dim strDate                 As String
    
    '[DETERMINE FORM AND FILENAME TO USE TO LOAD ROSTER FILE]
    If ActiveForm.Name = "frmScratch" Then
        flagScratch = True
        strSaveFile = strScratchFile
    Else
        flagScratch = False
        strSaveFile = strRosterFile
    End If
    
    '[CREATE FILENAME FOR FILE SAVE]
    If Len(strSaveFile) = 0 Then    '[CREATE IF NO SAVE FILE AVAILABLE]
        strDate = Format(frmRoster.MaskDate, "ddmmmyy")
        strSaveFile = LCase(Trim(strDate) & "." & Trim(DsClass("Code")))
    
        '[REMOVE SLASHES FROM FILENAME]
        strSaveFile = Replace(strSaveFile, "/", "-")
        strSaveFile = Replace(strSaveFile, "\", "-")
        strSaveFile = Replace(strSaveFile, " ", "_")
    End If
        
    '[SET DEFAULTS FOR FILE DIALOG]
    FileSetFilter
    
    '[CALL SAVE FILE ROUTINE]
    Call FileSaveAs(strSaveFile)
    
End Sub

Private Sub mnuScratch_Click()
    
    '[DISPLAY THE SCRATCH ROSTER]
    frmScratch.Show

End Sub

Private Sub mnuScratchCopy_Click()
    
    '[ADD CLIP AREA OF ROSTER TO strCLIP]
    flagScratch = True
    strClip = CopyCells
    flagScratch = False

End Sub

Private Sub mnuScratchDelete_Click()
    
    '[CLEAR SELECTED CELLS IN ROSTER]
    flagScratch = True
    Call DeleteCellContents(frmScratch)
    flagScratch = False

End Sub

Private Sub mnuScratchPaste_Click()

    '[PUT strClip IN ROSTER]
    If Len(strClip) = 0 Then Exit Sub
    flagScratch = True
    PasteCells
    flagScratch = False
    
End Sub


Private Sub mnuSettings_Click()

    '[SHOW THE CONTROL FORM]
    Call GotoControl

End Sub




Private Sub mnuSingleTimesheet_Click()

    '[CALL ROUTINE TO PROCESS SINGLE STAFF RECORD AND PRINT TIME SHEET]
    Call procSelectedStaffRoster

End Sub

Private Sub mnuStaffListDetailed_Click()

    Call procDetailedStaffList
    
End Sub

Private Sub mnuStaffListGeneral_Click()

    Call procGeneralStaffList

End Sub

Private Sub mnuStaffName_Click()

    '[DECLARATIONS]
    Dim intCounter          As Integer
    '[RELOCATE STAFF NAME]
    For intCounter = 0 To (frmSet.ListStaff.ListCount - 1)
        If frmSet.ListStaff.List(intCounter) = mdiMain.mnuStaffName.Caption Then
            '[NAME FOUND - SHOW FORM]
            frmSet.Show
            frmSet.tabSet.Tab = 1
            frmSet.ListStaff.ListIndex = intCounter
            Exit For
        End If
    Next intCounter

End Sub


Private Sub mnuStaffRpt_Click()
    
    '[PROCESS THE STAFF REPORT]
    procStaffReport

End Sub

Private Sub mnuStatusBarState_Click()

    '[SHOW OR HIDE THE STATUS BAR]
    Select Case mnuStatusBarState.Caption
    Case "Show St&atus Bar"
        mdiMain.panelStatusBar.Visible = True
        mdiMain.panelStatusBar.Top = 0
        mnuStatusBarState.Caption = "Hide St&atus Bar"
    Case "Hide St&atus Bar"
        mdiMain.panelStatusBar.Visible = False
        mnuStatusBarState.Caption = "Show St&atus Bar"
    End Select

End Sub

Private Sub mnuTileHor_Click()

    '[ARRANGE VISIBLE WINDOWS IN HORIZONTAL TILE]
    mdiMain.Arrange vbTileHorizontal
    
End Sub

Private Sub mnuTileVer_Click()

    '[ARRANGE VISIBLE WINDOWS IN VERTICAL TILE]
    mdiMain.Arrange vbTileVertical
    
End Sub

Private Sub mnuToolbarState_Click()

    Select Case mnuToolbarState.Caption
    Case "Show &Toolbar"
        mdiMain.PanelToolBar.Visible = True
        mnuToolbarState.Caption = "Hide &Toolbar"
    Case "Hide &Toolbar"
        mdiMain.PanelToolBar.Visible = False
        mnuToolbarState.Caption = "Show &Toolbar"
    End Select
    
End Sub

Private Sub panelStatusBar_DblClick()

    '[SIMULATE STATUS BAR TURNING OFF]
    Call mnuStatusBarState_Click
    
End Sub

Private Sub panelStatusBar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Select 'Hide Status Bar' from the 'Options' menu to remove this bar from the screen."
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub PanelToolBar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Select 'Hide Toolbar' from the 'Options' menu to remove this bar from the screen."
    '[---------------------------------------------------------------------------------]

End Sub


