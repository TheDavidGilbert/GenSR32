Attribute VB_Name = "modProcs"
Option Explicit
'[----------------------------------------------]
'[modProcs.bas      Basic Sub Modules           ]
'[----------------------------------------------]
'[GSR                                           ]
'[generic staff roster system for Windows       ]
'[ David Gilbert, 1997                          ]
'[----------------------------------------------]
'[version                           3.00.xx     ]
'[initial date                      21/07/1997  ]
'[----------------------------------------------]
Public DsStaff          As Dynaset
Public DsClass          As Dynaset
Public DsRoster         As Dynaset
Public DsDefault        As Dynaset
Public DsReport         As Dynaset
Public DsException      As Dynaset
Public DsRosterDetail   As Dynaset
Public DsStaffRoster    As Dynaset
Public DBMain           As Database
Public strReport        As String
Public intLastCheckCol  As Integer
Public intLastCheckRow  As Integer
Public sinLastMouseX    As Single
Public sinLastMouseY    As Single
Public GSRWorkspace     As Workspace
    
'[FLAG FOR BETA OR FULL VERSION]
Public flagBeta         As Boolean
Public flagTerminate    As Boolean
Public flagScratch      As Boolean
Public flagBackup       As Boolean
Public flagLoad         As Boolean      '[Flag for form load state]

Public constVersion As String           '[REV32 - CHANGE CONST VERSION TO STRING]
Public Const constEmail = "gensr@usa.net"
Public Const constWeb = "http://members.tripod.com/~GenSR"
Public Const constFileOut = 1     '[INTEGER FOR FILE ACCESS TYPE INPUT/OUTPUT]
Public Const constFileIn = 0
Public Const vbPRNone = 0

'[WARNING LEVELS FOR EXCEPTION REPORT]
Public Const constCritical = 2
Public Const constSerious = 1
Public Const constWarning = 0

'[PAY RATE TYPES]
Public Const vbHourly = 0
Public Const vbWeekly = 1

'[REPORT TYPES FOR TAG ON REPORT GRID]
Public Const vbExceptionReport = 1
Public Const vbRosterReport = 2
Public Const vbStaffReport = 3

'[CONSTANTS FOR INSIDE/OUTSIDE VALUES ON STAFF DAY AVAILABILITY]
Public Const vbNotAvail = 1
Public Const vbInside = 2
Public Const vbOutside = 3

'[CONSTANT FOR STAFF UNAVAILABLE DUE TO HOLIDAY]
Public Const vbHoliday = 4
Public Const vbNotInClass = 5

'[DELAY FOR BUTTON ANIMATION]
Public Const vbDelay = 0.15

Type WeekType
    ShortDay            As String
    LongDay             As String
End Type
    
'[STAFF TYPE FOR STAFF DETAILS]
Type StaffType
    DayDate         As Date
    Roster          As String
    Minutes         As Single
    Rate            As Single
    Amount          As Single
    StartTime       As Date
    EndTime         As Date
End Type

Type StaffReportType
    FullName        As String
    Roster          As String
    Day(1 To 7)     As String
End Type

'[GRAPH TYPE FOR WEEK BREAKDOWN GRAPH]
Type GraphType
    Roster          As String
    Active          As Boolean
    Time(1 To 7)    As Single
    Cost(1 To 7)    As Single
End Type

'[REV: 3.00.28]
'[ARRAY TYPE FOR STAFF WEEKLY RATE CALCULATIONS]
Type GraphStaffType
    RosterHours(1 To 10, 1 To 7)    As Single   '[HOURS ASSIGNED FOR EACH ROSTER FOR EACH DAY]
    PayType                         As Integer  '[TYPE OF PAY]
    PayRate                         As Single   '[PAY RATE]
    RosterTotal                     As Single   '[TOTAL MINUTES WORKED BY THIS STAFF MEMBER]
End Type

Public ArrayWeek(7)         As WeekType
Public flagDeleteConfirm    As Boolean      'flag for deletion confirmation
Public flagSounds           As Boolean      'flag for sounds
Public flagAllShifts        As Boolean      'flag for all shifts required to be filled
Public flagLoaded           As Boolean
Public intRosterClass       As Integer      'roster class id number (1 - 10)
Public gsrReturn            As Double       'return code from gsr msg box
Public gsrNote              As String       'string for roster note
Public strClip              As String       'string for holding clip
Public strFileName          As String       'string for filename
Public strRosterFile        As String       'string for roster filename
Public strGraphFile         As String       'string for graph/data filename
Public strScratchFile       As String       'string for scratch roster filename
Public sinDaysUsed          As Single       'number of days program has been installed for
Public sinModifier          As Single       'modifier for validation code
Public strBreak             As String       'cr string for text display in messages
Public strDateFormat        As String       'format for displayed dates
Public flagContinue         As Boolean      'flag to continue if cancel pressed in search/replace/transfer
Public flagConstraints      As Boolean      'follow day/hour contraints set on staff form for search/replace
Public intLogHandle         As Integer      'integer for log file handle
Public Const strLogFile = "startup.log"     'name of log file at program start-up
Public intStartPercent      As Integer      'percentage of startup routine completed

Sub CompactDatabase()

    '[SUB ROUTINE TO COMPACT DATABASE AND MAKE A BACKUP]
    '[VER: 3.00.33]
    '[ERROR HANDLER ROUTINE]
    On Error GoTo ErrorHandler
    Dim intStage            As Integer
    
    '[DISPLAY MESSAGE]
    Dim Msg
    Dim Style
    Dim Title
    Dim Response

    '[SHOW GSR MESSAGE FORM]
    Msg = "Please wait while GSR compacts the roster database and create a backup copy of your data." & strBreak & strBreak & "The backup roster information will be stored in gsr_back.dat."
    Style = vbInformation            ' Define buttons.
    Title = "Compacting and Performing Backup"
    Response = gsrMsg(Msg, Style, Title)
    frmMsg.gaugeBar.Visible = True
    frmMsg.gaugeBorder.Visible = True
    frmMsg.labelInfo.Visible = True
    frmMsg.ZOrder
    frmMsg.Refresh
    
    intStage = 1
    If Dir("gsr_comp.dat") > "" Then Kill "gsr_comp.dat"
    DBEngine.CompactDatabase "gsr.dat", "gsr_comp.dat"
    Call ReportProgressBar(50)
    intStage = 2
    If Dir("gsr_back.dat") > "" Then Kill "gsr_back.dat"
    intStage = 3
    Name "gsr.dat" As "gsr_back.dat"
    Call ReportProgressBar(75)
    intStage = 4
    Name "gsr_comp.dat" As "gsr.dat"
    Call ReportProgressBar(100)
    
    '[HIDE GSR MESSAGE FORM]
    Unload frmMsg

ErrorHandler:
    If Err.Number > 0 Then
        '[ERROR WHILE COMPACTING AND BACKING UP]
        Unload frmMsg
        
        Select Case intStage
        Case 1
            Msg = "Error: GSR encountered an error while compacting the GSR database." & strBreak & strBreak & "Please check that the GSR data files are not marked 'read-only'." & strBreak & strBreak & "Error Code: " & Err.Number
        Case 2
            Msg = "Error: GSR encountered an error while deleting the existing GSR backup file." & strBreak & strBreak & "Please check that the file 'gsr_back.dat' file is not marked 'read-only'." & strBreak & strBreak & "Error Code: " & Err.Number
        Case 3
            Msg = "Error: GSR encountered an error while backing up the GSR database." & strBreak & strBreak & "Please check that the file 'gsr_back.dat' is not marked 'read-only'." & strBreak & strBreak & "Error Code: " & Err.Number
        Case 4
            Msg = "Error: GSR encountered an error while backing up the GSR database." & strBreak & strBreak & "Please check that the file 'gsr.dat' is not marked 'read-only'." & strBreak & strBreak & "Error Code: " & Err.Number
        End Select
        Msg = Msg & strBreak & strBreak & "[" & Err.Description & " in " & Err.Source & "]"
        
        Style = vbOKOnly                     ' Define buttons.
        Title = "Error While Compacting/Backing Up"
        Response = gsrMsg(Msg, Style, Title)
        Exit Sub
    End If
    
End Sub

Function CopyCells() As String

    '[FUNCTION TO COPY THE SELECTED RANGE OF CELLS AND RETURN THE CLIP VALUE IN A STRING]
    '[THIS FUNCTION REPLACES vbKeyReturn WITH CHR$(254) STRING AND CONCAT'S THE STRING]

    Dim intTempRow          As Integer
    Dim intTempCol          As Integer
    Dim frmTemp             As Form
    Dim strCells            As String
    Dim strTempCell         As String
    Dim intCol              As Integer
    Dim intRow              As Integer
        
    '[SET ACTIVE FORM]
    If flagScratch = True Then
        Set frmTemp = frmScratch
    Else
        Set frmTemp = frmRoster
    End If

    '[SAVE POSITION]
    intTempRow = frmTemp.GridRoster.Row
    intTempCol = frmTemp.GridRoster.Col
    
    '[REV: 3.00.30]
    '[DON'T ALLOW COPYING FROM FIXED ROWS/COLS]
    '[COLUMNS]
    If frmTemp.GridRoster.SelStartCol + 1 <= frmTemp.GridRoster.FixedCols Then
        frmTemp.GridRoster.SelStartCol = frmTemp.GridRoster.FixedCols
    End If
    If frmTemp.GridRoster.SelEndCol + 1 <= frmTemp.GridRoster.FixedCols Then
        frmTemp.GridRoster.SelEndCol = frmTemp.GridRoster.FixedCols
    End If
    '[ROWS]
    If frmTemp.GridRoster.SelStartRow + 1 <= frmTemp.GridRoster.FixedRows Then
        frmTemp.GridRoster.SelStartRow = frmTemp.GridRoster.FixedRows
    End If
    If frmTemp.GridRoster.SelEndRow + 1 <= frmTemp.GridRoster.FixedRows Then
        frmTemp.GridRoster.SelEndRow = frmTemp.GridRoster.FixedRows
    End If
    
    '[DETERMINE WHETHER IT IS A MULTISELECT OR SINGLE CELL FILL]
     If frmTemp.GridRoster.SelStartCol = -1 Or frmTemp.GridRoster.SelEndCol = -1 Then
         '[NO CELLS TO COPY]
         strCells = ""
     Else
         '[MULTI CELL FILL]
         strCells = ""
         For intRow = frmTemp.GridRoster.SelStartRow To frmTemp.GridRoster.SelEndRow
             frmTemp.GridRoster.Row = intRow
             For intCol = frmTemp.GridRoster.SelStartCol To frmTemp.GridRoster.SelEndCol
                 frmTemp.GridRoster.Col = intCol
                 '[COPY AND REPLACE ENTERS IN CELL CONTENTS]
                 strTempCell = frmTemp.GridRoster.Text
                 strTempCell = Replace(strTempCell, Chr$(vbKeyReturn), Chr$(254))
                 strCells = strCells & strTempCell & Chr$(vbKeyTab)
             Next intCol
            strCells = strCells & Chr$(vbKeyReturn)
         Next intRow
     End If

    '[SAVE POSITION]
    frmTemp.GridRoster.Row = intTempRow
    frmTemp.GridRoster.Col = intTempCol
    
    '[RETURN VALUE]
    CopyCells = strCells
    
End Function


Sub Delay(sinDelay)

    '[LOOP FOR LENGTH OF DELAY]
    Dim sinTimer    As Single
    sinTimer = Timer + sinDelay
    Do While sinTimer > Timer: Loop
    
End Sub

Sub DeleteCellContents(frmTemp As Form)
    
    '[DELETE CONTENTS OF ROSTER CELLS]
    Dim intCounter          As Integer
    Dim intCol              As Integer
    Dim intRow              As Integer

    '[POPUP YES/NO DIALOG TO CONFIRM DELETION]
    Dim Msg As String
    Dim Style
    Dim Response
    Dim Title
    
    If frmTemp.GridRoster.Col <= 1 Or frmTemp.GridRoster.Row = 0 Then Exit Sub
    
    '[CHECK DELETION FLAG]
    If flagDeleteConfirm Then
        Msg = "This action will delete the contents of the selected cell/cells." & strBreak & strBreak & "Do you wish to continue ?"
        Style = vbYesNo                     ' Define buttons.
        Title = "Confirmation Required"     ' Define title.
        Response = gsrMsg(Msg, Style, Title)
    Else
        Response = vbYes
    End If
    
    If Response = vbYes Then    ' User chose Yes.

        '[DETERMINE WHETHER IT IS A MULTISELECT OR SINGLE CELL FILL]
        If frmTemp.GridRoster.SelStartCol = -1 Or frmTemp.GridRoster.SelEndCol = -1 Then
            '[SINGLE CELL FILL]
            '[CLEAR CELL CONTENTS]
            frmTemp.GridRoster.Text = ""
        Else
            '[MULTI CELL FILL]
            For intCol = frmTemp.GridRoster.SelStartCol To frmTemp.GridRoster.SelEndCol
                frmTemp.GridRoster.Col = intCol
                For intRow = frmTemp.GridRoster.SelStartRow To frmTemp.GridRoster.SelEndRow
                    frmTemp.GridRoster.Row = intRow
                    '[CLEAR CELL CONTENTS]
                    If frmTemp.GridRoster.Col <= 1 Or frmTemp.GridRoster.Row = 0 Then Exit For
                    frmTemp.GridRoster.Text = ""
                Next intRow
            Next intCol
        End If
        
        '[MAKE SAVE BUTTON VISIBLE]
        If flagScratch = False Then frmTemp.cmdSave.Visible = True
    
    End If

End Sub

Sub GridRedraw(flagState As Boolean)

    '[SET GRID REDRAW STATE]
    frmRoster.GridRoster.Redraw = flagState
    
End Sub

Sub PasteCells()
    
    '[SUB TO PASTE THE strClip STRING INTO THE CURRENT ROSTER]
    '[THIS FUNCTION REPLACES CHR$(254) WITH vbKeyReturn IN EACH SELECTED CELL]

    Dim intTempRow          As Integer
    Dim intTempCol          As Integer
    Dim frmTemp             As Form
    Dim strTempCell         As String
    Dim intCol              As Integer
    Dim intRow              As Integer
        
    '[SET ACTIVE FORM]
    If flagScratch = True Then
        Set frmTemp = frmScratch
    Else
        Set frmTemp = frmRoster
    End If

    '[SAVE POSITION]
    intTempRow = frmTemp.GridRoster.Row
    intTempCol = frmTemp.GridRoster.Col
    
    '[REV: 3.00.30]
    '[DON'T ALLOW PASTING INTO FIXED ROWS/COLS]
    '[COLUMNS]
    If frmTemp.GridRoster.SelStartCol + 1 <= frmTemp.GridRoster.FixedCols Then
        frmTemp.GridRoster.SelStartCol = frmTemp.GridRoster.FixedCols
    End If
    If frmTemp.GridRoster.SelEndCol + 1 <= frmTemp.GridRoster.FixedCols Then
        frmTemp.GridRoster.SelEndCol = frmTemp.GridRoster.FixedCols
    End If
    '[ROWS]
    If frmTemp.GridRoster.SelStartRow + 1 <= frmTemp.GridRoster.FixedRows Then
        frmTemp.GridRoster.SelStartRow = frmTemp.GridRoster.FixedRows
    End If
    If frmTemp.GridRoster.SelEndRow + 1 <= frmTemp.GridRoster.FixedRows Then
        frmTemp.GridRoster.SelEndRow = frmTemp.GridRoster.FixedRows
    End If
    
    '[DETERMINE WHETHER IT IS A MULTISELECT OR SINGLE CELL FILL]
     If frmTemp.GridRoster.SelStartCol = -1 Or frmTemp.GridRoster.SelEndCol = -1 Then
         '[NO CELLS TO PASTE]
         Exit Sub
     Else
        '[PASTE STRING INTO CLIP]
        frmTemp.GridRoster.Clip = strClip
         '[MULTI CELL FILL]
         For intRow = frmTemp.GridRoster.SelStartRow To frmTemp.GridRoster.SelEndRow
             frmTemp.GridRoster.Row = intRow
             For intCol = frmTemp.GridRoster.SelStartCol To frmTemp.GridRoster.SelEndCol
                 frmTemp.GridRoster.Col = intCol
                 '[COPY AND REPLACE ENTERS IN CELL CONTENTS]
                 strTempCell = frmTemp.GridRoster.Text
                 '[PLACE CONTENTS BACK INTO CELL]
                 frmTemp.GridRoster.Text = Replace(strTempCell, Chr$(254), Chr$(vbKeyReturn))
             Next intCol
         Next intRow
     End If

    '[SAVE POSITION]
    frmTemp.GridRoster.Row = intTempRow
    frmTemp.GridRoster.Col = intTempCol

End Sub


Function prnColour() As Integer
    
    '[HANDLER TO CATCH NOT COLOUR PRINTER ERROR]]
    On Error GoTo ErrorHandler:

    '[CHECK PRINTER COLOUR CAPABILITIES AND RETURN CODE AS RESULT]
    Dim intResult           As Integer
        
    If Printers.Count = 0 Then
        '[RETURN 0 - NO PRINTER]
        intResult = 0
    Else
        '[RETURN PRINTER COLOR]
        intResult = Printer.ColorMode
    End If
    prnColour = intResult
    
ErrorHandler:
    If Err.Number = 483 Then    '[NOT A COLOR PRINTER]
        prnColour = 1
    End If
    
End Function

Sub prnMasterRoster()

    '[BUILD ROSTER DYNASET WITH ALL RECORDS]
    Dim SQLStmt         As String
    Dim strBookmark     As String
    Dim strFullname     As String
    Dim strDayKey       As String
    Dim intCounter      As Integer
    Dim intDayCount     As Integer
    Dim strClassDesc    As String
    Dim DsRosterReport  As Recordset
    Dim strNames        As String
    
    '[DSROSTER MUST CONTAIN RECORDS ELSE SUB WOULD NOT BE CALLED}
    DsRoster.MoveFirst
    
    '[DELETE EXISTING ROSTER REPORT RECORDS]
    SQLStmt = "DELETE * FROM [RosterReport]"
    DBMain.Execute SQLStmt, dbFailOnError
    
    '[OPEN ROSTER REPORT DYNASET]
    Set DsRosterReport = DBMain.OpenRecordset("RosterReport", dbOpenDynaset)

    If Not (DsRoster.EOF And DsRoster.BOF) Then
        DsRoster.MoveFirst
        Do While Not DsRoster.EOF
            '[CHECK TO SEE IF THE ROSTER IS ACTIVE]
            DsClass.AbsolutePosition = (DsRoster!Class - 1)
            If DsClass("Active") = vbChecked Then
                With DsRosterReport
                    .AddNew
                    !Class = DsRoster!Class
                    !Roster = DsClass!Description
                    !Day_1 = Trim(DsRoster!Day_1 & "")
                    !Day_2 = Trim(DsRoster!Day_2 & "")
                    !Day_3 = Trim(DsRoster!Day_3 & "")
                    !Day_4 = Trim(DsRoster!Day_4 & "")
                    !Day_5 = Trim(DsRoster!Day_5 & "")
                    !Day_6 = Trim(DsRoster!Day_6 & "")
                    !Day_7 = Trim(DsRoster!Day_7 & "")
                    !ShiftStart = Format(DsRoster!ShiftStart, "Medium Time")
                    !ShiftEnd = Format(DsRoster!ShiftEnd, "Medium Time")
                    !TimeFull = Format(DsRoster!ShiftStart, "Medium Time") & " to " & Format(DsRoster!ShiftEnd, "Medium Time")
                    .Update
                End With
            End If
            DsRoster.MoveNext
        Loop
    End If
    
    '[MOVE BACK TO START]
    DsRoster.MoveFirst
    '[CLOSE DYNASET]
    DsRosterReport.Close

End Sub

Sub procMasterRoster()

    '[REV: 3.00.28]
    '[ERROR HANDLER ROUTINE]
    On Error GoTo ErrorHandler
    
    '[USE THE REPORT FORM TO ADDRESS ANY OF THESE PROBLEMS]
    Dim Msg
    Dim Style
    Dim Title
    Dim Response
    
    '[REBUILD ROSTER DYNASET USING ALL ROSTERS]
    Dim SQLStmt         As String
    
    '[SELECT APPROPRIATE RECORDS]
    SQLStmt = "SELECT * FROM Roster ORDER BY SHIFTSTART, SHIFTEND, CLASS"
    Set DsRoster = DBMain.OpenRecordset(SQLStmt, dbOpenSnapshot)
       
    
    '[CANNOT PRINT IF ROSTER IS EMPTY]
    If DsRoster.EOF And DsRoster.BOF Then
        Msg = "The roster database does not contain any information or you have not saved any changes to the current roster." & strBreak & strBreak & "GSR cannot print a blank roster."
        Style = vbOKOnly                     ' Define buttons.
        Title = "Cannot Print Blank Roster"
        Response = gsrMsg(Msg, Style, Title)
        '{REBUILD ROSTER DYNASET]
        Call BuildRosterDynaset(intRosterClass)
        Exit Sub
    End If
    
    '[SHOW WARNING FORM]
    Msg = "Continue and print the master roster ?"
    Style = vbYesNo                     ' Define buttons.
    Title = "Confirmation Required"
    Response = gsrMsg(Msg, Style, Title)
    If Response = vbYes Then
        '[SHOW GSR MESSAGE FORM]
        Msg = "GSR is now printing the master roster." & strBreak & strBreak & "Please wait."
        Style = vbInformation            ' Define buttons.
        Title = "Printing Master Roster"
        Response = gsrMsg(Msg, Style, Title)
        frmMsg.gaugeBar.Visible = True
        frmMsg.gaugeBorder.Visible = True
        frmMsg.labelInfo.Visible = True
        frmMsg.ZOrder
        frmMsg.Refresh
        
        '[PRINT ROSTER TO REPORT FORM HERE]
        Call prnMasterRoster
        
        '[HIDE GSR MESSAGE FORM]
        Unload frmMsg
    
        '[SET REPORT TITLE]
        Select Case prnOrient
        Case vbPRNone           '[NO PRINTER ATTACHED]
            strReport = "p_master.rpt"
        Case vbPRORPortrait     '[PORTRAIT STYLE]
            strReport = "p_master.rpt"
        Case vbPRORLandscape    '[LANDSCAPE STYLE]
            strReport = "l_master.rpt"
        Case Else
            strReport = "p_master.rpt"
        End Select
        
        mdiMain.Report.ReportFileName = strReport
    
    
        mdiMain.Report.Formulas(0) = ""
        mdiMain.Report.WindowTitle = "Master Roster, Week Starting " & Format(DsDefault("StartDate"), strDateFormat)
        mdiMain.Report.SortFields(0) = ""
        mdiMain.Report.Action = 1
        
    End If
    
    '{REBUILD ROSTER DYNASET]
    Call BuildRosterDynaset(intRosterClass)
    
ErrorHandler:
    If Err.Number > 0 Then
        '[CANNOT FIND ROSTER FILE ?]
        Msg = "Error: GSR cannot produce this report." & strBreak & strBreak & "The required report file (" & strReport & ") may be missing from the GSR directory." & strBreak & strBreak & "Error Code: " & Err.Number
        If Err.Number = 20504 Then
            Msg = Msg & strBreak & strBreak & "[Could not find required report file]"
        Else
            Msg = Msg & strBreak & strBreak & "[" & Err.Description & " in " & Err.Source & "]"
        End If
        Style = vbOKOnly                     ' Define buttons.
        Title = "Cannot Produce Report"
        Response = gsrMsg(Msg, Style, Title)
        '{REBUILD ROSTER DYNASET]
        Call BuildRosterDynaset(intRosterClass)
        Exit Sub
    End If

End Sub

Sub procPickStaff()
    
    '[CHECK SELECTED ROSTER CELL AND PICK ALL STAFF MEMBERS WHO CAN FILL THIS TIME SLOT]]
    Dim intCounter          As Integer
    Dim intTrCount          As Integer
    Dim intSelCount         As Integer
    Dim flagResult          As Boolean
    Dim dateStart           As Date
    Dim strFullname     As String
    Dim dateEnd         As Date
    
    flagResult = False
    
    '[CLEAR ALL SELECTIONS FROM STAFF LIST]
    For intCounter = 0 To (frmRoster.ListStaff.ListCount - 1)
        strFullname = frmRoster.ListStaff.List(intCounter)
        frmRoster.ListStaff.Selected(intCounter) = False
        Call PickStaff(flagResult, strFullname)
        If flagResult = True Then frmRoster.ListStaff.Selected(intCounter) = True
    Next intCounter

End Sub

Sub ProcRosterBreakdown()
    
    '[COMMAND TO PRODUCE A ROSTER BREAKDOWN GRAPH]
    Dim Msg
    Dim Style
    Dim Title
    Dim Response
    Dim flagResult          As Boolean
    
    '[CANNOT PRODUCE IF NO STAFF ARE DEFINED]
    If (DsStaff.EOF And DsStaff.BOF) Or (DsStaff.RecordCount = 1 And DsStaff!LastName = "*LastName") Then
        Msg = "There are no staff records entered at present." & strBreak & strBreak & "Please enter staff details on the staff form before producing this graph."
        Style = vbOKOnly                     ' Define buttons.
        Title = "Cannot Create Graph"
        Response = gsrMsg(Msg, Style, Title)
        Exit Sub
    End If
    
    '[SHOW WARNING FORM]
    Msg = "The roster breakdown graph details the contribution of each roster to your daily staff costs." & strBreak & strBreak & "Because this routine has to perform multiple comparisions and searches, it may take a few minutes to complete, depending upon the number of rosters, staff and the speed of your computer." & strBreak & strBreak & "Do you wish to continue and produce the report ?"
    Style = vbYesNo                     ' Define buttons.
    Title = "Confirmation Required"
    Response = gsrMsg(Msg, Style, Title)
    If Response = vbYes Then
        '[PRODUCE THE GSR ROSTER GRAPH BREAKDOWN]
        Unload frmGraph
        frmGraph.Show
        Call frmGraph.procRosterGraph(flagResult)
    End If

End Sub

Function Replace(strString, strFind, strReplace)

    '[REPLACE ALL OCCURENCES OF strFind IN strString WITH strReplace]
    Dim intCounter          As Integer
    For intCounter = 1 To Len(strString)
        If Mid$(strString, intCounter, 1) = strFind Then
            Mid$(strString, intCounter, 1) = strReplace
        End If
    Next intCounter
    
    '[RETURN VALUE]
    Replace = strString
    
End Function



Function ReplaceStr(TextIn, ByVal SearchStr As String, ByVal Replacement As String, ByVal CompMode As Integer)
    Dim WorkText As String, Pointer As Integer
    If IsNull(TextIn) Then
        ReplaceStr = Null
    Else
        WorkText = TextIn
        Pointer = InStr(1, WorkText, SearchStr, CompMode)
        Do While Pointer > 0
            WorkText = Left(WorkText, Pointer - 1) & Replacement & Mid(WorkText, Pointer + Len(SearchStr))
            Pointer = InStr(Pointer + Len(Replacement), WorkText, SearchStr, CompMode)
        Loop
        ReplaceStr = WorkText
    End If
End Function

Function SQLFixup(TextIn)
    SQLFixup = ReplaceStr(TextIn, "'", "''", 0)
End Function
Function prnOrient() As Integer

    '[CHECK PRINTER ORIENTATION AND RETURN ORIENT CODE AS RESULT]
    On Error GoTo ErrorHandler
    Dim intResult           As Integer
        
    If Printers.Count = 0 Then
        '[RETURN 0 - NO PRINTER]
        intResult = vbPRNone
    Else
        '[RETURN ORIENTATION]
        intResult = Printer.Orientation
    End If
    prnOrient = intResult

ErrorHandler:
    If Err.Number > 0 Then
        '[ERROR GETTING ORIENTATION - SET TO STANDARD PORTRAIT]
        prnOrient = vbPRORPortrait
    End If

End Function

Sub SaveScratchToFile()

    '[SAVE SCRATCH ROSTER TO A FILE]
    '[FILENAME IS PREALLOCATED BUT MAY BE CHANGED]
    Dim strSaveFile             As String
    Dim intClassCounter         As Integer
    Dim strBookmark             As String
    Dim intSlashFound           As Integer
    Dim strDate                 As String
    Dim flagSave                As Boolean
    Dim intCounter              As Integer
    
    '[CREATE FILENAME FOR FILE SAVE]
    If mdiMain.CommonDialog.FileName > "" Then
        strSaveFile = mdiMain.CommonDialog.FileName
    Else
        If Len(DsDefault("StartDate")) > 8 Then
            strDate = LCase(Left(DsDefault("StartDate"), 6) & Right(DsDefault("StartDate"), 2))
        Else
            strDate = Trim(LCase(DsDefault("StartDate")))
        End If
        
        strSaveFile = LCase(Trim(strDate) & "." & Trim(DsClass("Code")))
    
        '[REMOVE SLASHES FROM FILENAME]
        strSaveFile = Replace(strSaveFile, "/", "-")
        strSaveFile = Replace(strSaveFile, "\", "-")
        strSaveFile = Replace(strSaveFile, " ", "_")
        
    End If
    
    '[SET DEFAULTS FOR FILE DIALOG]
    FileSetFilter
    
    '[CALL SAVE FILE ROUTINE]
    flagSave = FileSaveAs(strSaveFile)

    '[SET CAPTION]
    If flagSave = True Then frmScratch.Caption = "Scratch Roster : " & mdiMain.CommonDialog.FileName
    
End Sub

Sub CloseStaffRosterSet()

    '[CLOSE STAFF ROSTER DYNASET]
    DsStaffRoster.Close

End Sub


Sub ExpandCells()
    
    '[RESIZE ALL CELL WIDTH TO MATCH CONTENTS]
    Dim intColCounter   As Integer
    Dim intRowCounter   As Integer
    Dim intCols         As Integer
    Dim intRows         As Integer
    Dim intRow          As Integer
    Dim intCol          As Integer
    Dim frmTemp         As Form
    
    '[SET ACTIVE FORM]
    If flagScratch = True Then
        Set frmTemp = frmScratch
    Else
        Set frmTemp = frmRoster
    End If
    
    '[SAVE STARTING GRID POSITION]
    intCol = frmTemp.GridRoster.Col
    intRow = frmTemp.GridRoster.Row
    intCols = (frmTemp.GridRoster.Cols - 1)
    intRows = (frmTemp.GridRoster.Rows - 1)
    For intRowCounter = 1 To intRows
        frmTemp.GridRoster.Row = intRowCounter
        frmTemp.GridRoster.RowHeight(frmTemp.GridRoster.Row) = (frmTemp.TextHeight("A") * 1.1)
        
        For intColCounter = 0 To intCols
            frmTemp.GridRoster.Col = intColCounter
            Call ResizeRosterCell(frmTemp.GridRoster.Text)
        Next intColCounter
    Next intRowCounter
    '[RESTORE STARTING GRID POSITION]
    frmTemp.GridRoster.Col = intCol
    frmTemp.GridRoster.Row = intRow

End Sub


Sub InitStaffRosterSet()

    '[CLEAR STAFF ROSTER SET FOR PRINTING]
    Dim SQLStmt         As String
    
    '[DELETE EXISTING STAFF ROSTER REPORT RECORDS]
    SQLStmt = "DELETE * FROM [StaffRoster]"
    DBMain.Execute SQLStmt, dbFailOnError
    
    '[OPEN STAFF ROSTER REPORT DYNASET]
    Set DsStaffRoster = DBMain.OpenRecordset("StaffRoster", dbOpenDynaset)

End Sub

Sub Main()

    '[DECLARATIONS]
    Dim Response        As Integer
    Dim strDataFile     As String
    
    '[REV32 - SET CONST VERSION INFO HERE]
    constVersion = App.Major & "." & Format(App.Minor, "00") & "." & Format(App.Revision, "00")
    
    '[MAIN SUBROUTINE FOR GSR LOADUP]
    '[Rev: 3.00.29]
    '[OPEN LOG FILE FOR STARTUP]
    intLogHandle = OpenFile(strLogFile, constFileOut)
    
    '[DEBUG]
    Call LogToFile("=------------------------GSR STARTUP LOG-=")
    Call LogToFile("Generic Staff Roster")
    Call LogToFile(constVersion)
    Call LogToFile(Date & " at " & Time)
    Call LogToFile("=----------------------------------------=")

    '[GRID REDRAW]
    GridRedraw False
    
    '[Rev: 3.00.35]
    '[NEED TO CHECK FOR EXISTENCE OF GSR.DAT FILE]
    '[DEBUG]
    Call LogToFile("checking GSR.DAT file exists")
    strDataFile = Dir("gsr.dat")
    '[IF GSR.DAT FILE EXISTS, NO ACTION SHOULD BE TAKEN]
    '[IF GSR.DAT FILE DOES NOT EXIST, POPUP WARNING]
    If IsNull(strDataFile) Or strDataFile = "" Then
       '[IF NEITHER FILE EXISTS, POPUP WARNING AND EXIT]
        Response = MsgBox("The Generic Staff Roster database file (GSR.DAT) is missing from the current directory." & Chr$(vbKeyReturn) & Chr$(vbKeyReturn) & "Please Re-Install Generic Staff Roster.", vbCritical & vbOKOnly, "Critical Error - Data File Missing")
        '[DEBUG]
        Call LogToFile("* GSR.DAT file not found - TERMINATING")
        Close
        End
    End If

    '[DEBUG]
    Call LogToFile("LOADING FORM: SPLASH")
    
    '[LOAD SPLASH FORM IN FRONT]
    frmSplash.Show
    frmSplash.Refresh
    
    '[DEBUG]
    Call LogToFile("LOADING FORM: MAIN")
    Load mdiMain
    
    '[SHOW MAIN FORM]
    '[Rev: 3.00.32]
    '[DEBUG]
    Call LogToFile("showing child forms")
    mdiMain.AutoShowChildren = True
    mdiMain.Show
    frmRoster.Show
    
    '[DEBUG]
    Call LogToFile("resizing roster columns")
    Call SizeRosterColumnsToGrid
    
    '[DEBUG]
    Call LogToFile("=-------------------------------COMPLETE-=")
    
    '[UNLOAD SPLASH FORM]
    Unload frmSplash
    
    '[Rev: 3.00.29]
    '[CLOSE GSR LOG FILE]
    Close intLogHandle
    
    '[GRID REDRAW]
    GridRedraw True

End Sub

Sub LogToFile(strMessage)

    '[Rev: 3.00.29]
    '[SUB ROUTINE TO PRINT MESSAGE TO LOGFILE WITH FORMATTED TIME]
    Print #intLogHandle, Time, strMessage
    
    '[REV:3.00.34]
    '[ALLOW OTHER ACTIVITIES TO OCCUR]
    DoEvents
    
    '[REV:3.00.34]
    '[SHOW PROGRESS ON SPLASH FORM]
    If frmSplash.Visible Then
        intStartPercent = intStartPercent + 3
        Call StartupProgressBar(intStartPercent)
    End If
    
End Sub


Sub PickStaff(flagResult, strFullString)

    '[SUB-ROUTINE TO PICK STAFF FROM THE STAFF LIST WHO CAN FILL THE DESIRED ROSTER SLOT]
    Dim intCounter          As Integer
    Dim strLastName         As String
    Dim strFirstName        As String
    Dim intDelimiter        As Integer
    Dim intCol              As Integer
    Dim intRow              As Integer
    Dim intTimeCol          As Integer
    Dim intResult           As Integer
    Dim dateStart           As Date
    Dim dateEnd             As Date
    Dim strMessage          As String
    
    '[DETERMINE WHETHER IT IS A MULTISELECT OR SINGLE CELL FILL]
    flagResult = False
    'If (frmRoster.GridRoster.SelStartCol = -1) Or (frmRoster.GridRoster.SelEndCol = -1) Or (frmRoster.GridRoster.SelStartCol = frmRoster.GridRoster.SelEndCol And frmRoster.GridRoster.SelStartRow = frmRoster.GridRoster.SelEndRow) Then
    If (frmRoster.GridRoster.Col = -1) Or (frmRoster.GridRoster.ColSel = -1) Or (frmRoster.GridRoster.Col = frmRoster.GridRoster.ColSel And frmRoster.GridRoster.Row = frmRoster.GridRoster.RowSel) Then
        '[SINGLE CELL FILL]
        '[CALL ROUTINE TO PLACE NAME IN CELL]
        If frmRoster.GridRoster.Col <= 1 Then Exit Sub
        
        '[LOCATE START AND FINISH TIME]
        intTimeCol = frmRoster.GridRoster.Col
        frmRoster.GridRoster.Col = 0
        If frmRoster.GridRoster.Text = "" Then
            Exit Sub
        End If
        dateStart = frmRoster.GridRoster.Text
        frmRoster.GridRoster.Col = 1
        If frmRoster.GridRoster.Text = "" Then
            Exit Sub
        End If
        dateEnd = frmRoster.GridRoster.Text
        '[ADD 24 HOURS IF NEXT DAY]
        If dateEnd < dateStart Then dateEnd = dateEnd + CDate("12:00") + CDate("12:00")
        frmRoster.GridRoster.Col = intTimeCol
            
        If CheckStaffDay(strFullString, frmRoster.GridRoster.Col - 1, dateStart, dateEnd, intResult, strMessage) Or flagConstraints = False Then
            flagResult = True
        Else
            flagResult = False
        End If
    Else
        '[MULTI CELL FILL]
            intCol = frmRoster.GridRoster.Col '.SelStartCol
            intRow = frmRoster.GridRoster.Row '.SelStartRow
            If intCol > 1 And intRow > 0 Then
            
                '[LOCATE START AND FINISH TIME]
                intTimeCol = frmRoster.GridRoster.Col
                frmRoster.GridRoster.Col = 0
                dateStart = frmRoster.GridRoster.Text
                frmRoster.GridRoster.Col = 1
                dateEnd = frmRoster.GridRoster.Text
                '[ADD 24 HOURS IF NEXT DAY]
                If dateEnd < dateStart Then dateEnd = dateEnd + CDate("12:00") + CDate("12:00")
                frmRoster.GridRoster.Col = intTimeCol
                
                frmRoster.GridRoster.Col = intCol
                frmRoster.GridRoster.Row = intRow
                
                '[CHECK STAFF MEMBER IS AVAILABLE FOR THIS DAY]
                If CheckStaffDay(strFullString, frmRoster.GridRoster.Col - 1, dateStart, dateEnd, intResult, strMessage) Or flagConstraints = False Then
                    flagResult = True
                Else
                    flagResult = False
                End If
            End If
    End If

End Sub


Sub prnRoster()

    '[BUILD ROSTER DYNASET WITH ALL RECORDS]
    Dim SQLStmt         As String
    Dim strBookmark     As String
    Dim strFullname     As String
    Dim strDayKey       As String
    Dim intCounter      As Integer
    Dim intDayCount     As Integer
    Dim strClassDesc    As String
    Dim DsRosterReport  As Recordset
    Dim strNames        As String
    
    '[DSROSTER MUST CONTAIN RECORDS ELSE SUB WOULD NOT BE CALLED}
    DsRoster.MoveFirst
    
    '[DELETE EXISTING ROSTER REPORT RECORDS]
    SQLStmt = "DELETE * FROM [RosterReport]"
    DBMain.Execute SQLStmt, dbFailOnError
        
    '[OPEN ROSTER REPORT DYNASET]
    Set DsRosterReport = DBMain.OpenRecordset("RosterReport", dbOpenTable)

    If Not (DsRoster.EOF And DsRoster.BOF) Then
        DsRoster.MoveFirst
        Do While Not DsRoster.EOF
            With DsRosterReport
                .AddNew
                !Class = DsRoster!Class
                !Roster = frmRoster.ComboClass.Text & " Roster"
                !Day_1 = Trim(DsRoster!Day_1 & "")
                !Day_2 = Trim(DsRoster!Day_2 & "")
                !Day_3 = Trim(DsRoster!Day_3 & "")
                !Day_4 = Trim(DsRoster!Day_4 & "")
                !Day_5 = Trim(DsRoster!Day_5 & "")
                !Day_6 = Trim(DsRoster!Day_6 & "")
                !Day_7 = Trim(DsRoster!Day_7 & "")
                !ShiftStart = Format(DsRoster!ShiftStart, "Medium Time")
                !ShiftEnd = Format(DsRoster!ShiftEnd, "Medium Time")
                .Update
            End With
            DsRoster.MoveNext
        Loop
    End If
    
    '[MOVE BACK TO START]
    DsRoster.MoveFirst
    '[CLOSE DYNASET]
    DsRosterReport.Close
    
End Sub

Sub procDetailedStaffList()
    
    '[ERROR HANDLER ROUTINE]
    On Error GoTo ErrorHandler
    
    '[MESSAGE VARIABLES]
    Dim Msg
    Dim Style
    Dim Title
    Dim Response
    
    '[CANNOT PRINT IF NO STAFF ARE DEFINED]
    If (DsStaff.EOF And DsStaff.BOF) Or (DsStaff.RecordCount = 1 And DsStaff!LastName = "*LastName") Then
        Msg = "There are no staff records entered at present." & strBreak & strBreak & "Please enter staff details on the staff form before executing this report."
        Style = vbOKOnly                     ' Define buttons.
        Title = "Cannot Create Report"
        Response = gsrMsg(Msg, Style, Title)
        Exit Sub
    End If

    '[SET REPORT TITLE]
    Dim intOrient           As Integer
    intOrient = vbPRNone    '[SET TO NO PRINTER]
    
    Select Case prnOrient
    Case vbPRNone           '[NO PRINTER ATTACHED]
        strReport = "p_detstf.rpt"
    Case vbPRORPortrait     '[PORTRAIT STYLE]
        strReport = "p_detstf.rpt"
    Case vbPRORLandscape    '[LANDSCAPE STYLE]
        intOrient = prnOrient
        Printer.Orientation = vbPRORPortrait
        strReport = "p_detstf.rpt"
    Case Else
        intOrient = prnOrient
        Printer.Orientation = vbPRORPortrait
        strReport = "p_detstf.rpt"
    End Select
    
    mdiMain.Report.ReportFileName = strReport
   
    '[SET REPORT TITLE]
    mdiMain.Report.Formulas(0) = ""


    mdiMain.Report.SortFields(0) = "+{Staff.StaffID}"
    mdiMain.Report.WindowTitle = "Detailed Staff List"
    mdiMain.Report.Action = 1

    '[RESTORE PRINTER ORIENTATION IF CHANGED]
    If intOrient > 0 Then Printer.Orientation = intOrient

ErrorHandler:
    If Err.Number > 0 Then
        '[CANNOT FIND ROSTER FILE ?]
        Msg = "Error: GSR cannot produce this report." & strBreak & strBreak & "The required report file (" & strReport & ") may be missing from the GSR directory." & strBreak & strBreak & "Error Code: " & Err.Number
        If Err.Number = 20504 Then
            Msg = Msg & strBreak & strBreak & "[Could not find required report file]"
        Else
            Msg = Msg & strBreak & strBreak & "[" & Err.Description & " in " & Err.Source & "]"
        End If
        Style = vbOKOnly                     ' Define buttons.
        Title = "Cannot Produce Report"
        Response = gsrMsg(Msg, Style, Title)
        Exit Sub
    End If
    
End Sub

Sub procGeneralStaffList()

    '[ERROR HANDLER ROUTINE]
    On Error GoTo ErrorHandler
    
    '[MESSAGE VARIABLES]
    Dim Msg
    Dim Style
    Dim Title
    Dim Response
    
    '[CANNOT PRINT IF NO STAFF ARE DEFINED]
    If (DsStaff.EOF And DsStaff.BOF) Or (DsStaff.RecordCount = 1 And DsStaff!LastName = "*LastName") Then
        Msg = "There are no staff records entered at present." & strBreak & strBreak & "Please enter staff details on the staff form before executing this report."
        Style = vbOKOnly                     ' Define buttons.
        Title = "Cannot Create Report"
        Response = gsrMsg(Msg, Style, Title)
        Exit Sub
    End If
    
    '[SET REPORT TITLE]
    Dim intOrient           As Integer
    intOrient = vbPRNone    '[SET TO NO PRINTER]
    
    Select Case prnOrient
    Case vbPRNone           '[NO PRINTER ATTACHED]
        strReport = "p_genstf.rpt"
    Case vbPRORPortrait     '[PORTRAIT STYLE]
        strReport = "p_genstf.rpt"
    Case vbPRORLandscape    '[LANDSCAPE STYLE]
        intOrient = prnOrient
        Printer.Orientation = vbPRORPortrait
        strReport = "p_genstf.rpt"
    Case Else
        intOrient = prnOrient
        Printer.Orientation = vbPRORPortrait
        strReport = "p_genstf.rpt"
    End Select
    
    mdiMain.Report.ReportFileName = strReport
      
    '[SET REPORT TITLE]
    Printer.Orientation = vbPRORPortrait


    mdiMain.Report.Formulas(0) = ""
    mdiMain.Report.SortFields(0) = "+{Staff.LastName}"
    mdiMain.Report.WindowTitle = "General Staff List"
    mdiMain.Report.Action = 1
    
    '[RESTORE PRINTER ORIENTATION IF CHANGED]
    If intOrient > 0 Then Printer.Orientation = intOrient

ErrorHandler:
    If Err.Number > 0 Then
        '[CANNOT FIND ROSTER FILE ?]
        Msg = "Error: GSR cannot produce this report." & strBreak & strBreak & "The required report file (" & strReport & ") may be missing from the GSR directory." & strBreak & strBreak & "Error Code: " & Err.Number
        If Err.Number = 20504 Then
            Msg = Msg & strBreak & strBreak & "[Could not find required report file]"
        Else
            Msg = Msg & strBreak & strBreak & "[" & Err.Description & " in " & Err.Source & "]"
        End If
        Style = vbOKOnly                     ' Define buttons.
        Title = "Cannot Produce Report"
        Response = gsrMsg(Msg, Style, Title)
        Exit Sub
    End If

End Sub

Sub procStaffReport()

    '[ERROR HANDLER ROUTINE]
    On Error GoTo ErrorHandler

    '[PRODUCE STAFF DETAILS REPORT, LISTING STAFF MEMBER, DAYS AVAILABLE AND ROSTERS AVAILABLE]
    Dim intCounter As Integer
    Dim Msg
    Dim Style
    Dim Title
    Dim Response
    
    '[CANNOT PRINT IF NO STAFF ARE DEFINED]
    If (DsStaff.EOF And DsStaff.BOF) Or (DsStaff.RecordCount = 1 And DsStaff!LastName = "*LastName") Then
        Msg = "There are no staff records entered at present." & strBreak & strBreak & "Please enter staff details on the staff form before executing this report."
        Style = vbOKOnly                     ' Define buttons.
        Title = "Cannot Create Report"
        Response = gsrMsg(Msg, Style, Title)
        Exit Sub
    End If
    
    '[SHOW WARNING FORM]
    Msg = "The Staff Details Report lists staff availability for each week day for every active roster.  Use this report to help create your staff rosters." & strBreak & strBreak & "All times are shown in 24 hour format because of space limitations.  This report should not take long to produce." & strBreak & strBreak & "Do you wish to continue and produce the report ?"
    Style = vbYesNo                     ' Define buttons.
    Title = "Confirmation Required"
    Response = gsrMsg(Msg, Style, Title)
    If Response = vbYes Then
    
        '[SHOW PLEASE WAIT MESSAGE]
        '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
        StatusBar "Please wait while this report is processed."
        '[---------------------------------------------------------------------------------]
        mdiMain.panelStatusBar.Refresh
        
        '[SHOW GSR MESSAGE FORM]
        Msg = "GSR is now processing your staff records and producing the Staff Detail Report." & strBreak & strBreak & "The report will list day and roster availabilty all staff members." & strBreak & strBreak & "This report should not take long to process."
        Style = vbInformation            ' Define buttons.
        Title = "Staff Details Report"
        Response = gsrMsg(Msg, Style, Title)
        frmMsg.gaugeBar.Visible = True
        frmMsg.gaugeBorder.Visible = True
        frmMsg.labelInfo.Visible = True
        frmMsg.ZOrder
        frmMsg.Refresh
        
        '[CALL STAFF Details Report SUBROUTINE]
        StaffReport
        
        '[HIDE GSR MESSAGE FORM]
        Unload frmMsg
    
        '[SET REPORT TITLE]
        Select Case prnOrient
        Case vbPRNone           '[NO PRINTER ATTACHED]
            strReport = "p_stfdet.rpt"
        Case vbPRORPortrait     '[PORTRAIT STYLE]
            strReport = "p_stfdet.rpt"
        Case vbPRORLandscape    '[LANDSCAPE STYLE]
            strReport = "l_stfdet.rpt"
        Case Else
            strReport = "p_stfdet.rpt"
        End Select
        
        mdiMain.Report.ReportFileName = strReport
        
        '[SET REPORT TITLE]
    
    
        mdiMain.Report.Formulas(0) = ""
        mdiMain.Report.WindowTitle = "Staff Detail Report, Week Starting " & Format(DsDefault("StartDate"), strDateFormat)
        mdiMain.Report.SortFields(0) = "+{StaffRoster.FullName}"
        mdiMain.Report.SortFields(0) = "+{StaffRoster.Roster}"
        mdiMain.Report.Action = 1

    End If

ErrorHandler:
    If Err.Number > 0 Then
        '[CANNOT FIND ROSTER FILE ?]
        Msg = "Error: GSR cannot produce this report." & strBreak & strBreak & "The required report file (" & strReport & ") may be missing from the GSR directory." & strBreak & strBreak & "Error Code: " & Err.Number
        If Err.Number = 20504 Then
            Msg = Msg & strBreak & strBreak & "[Could not find required report file]"
        Else
            Msg = Msg & strBreak & strBreak & "[" & Err.Description & " in " & Err.Source & "]"
        End If
        Style = vbOKOnly                     ' Define buttons.
        Title = "Cannot Produce Report"
        Response = gsrMsg(Msg, Style, Title)
        Exit Sub
    End If

End Sub


Sub ShowRowCol(intRow, intCol)

    '[MAKE SELECTED ROSTER CELL VISIBLE]
            
    '[CHECK FOR ZERO VALUE]
    If intRow = 0 Then intRow = 1
    frmRoster.GridRoster.Row = intRow
    frmRoster.GridRoster.Col = intCol
    frmRoster.GridRoster.LeftCol = frmRoster.GridRoster.FixedCols
    frmRoster.GridRoster.TopRow = 1
    Do While Not frmRoster.GridRoster.RowIsVisible(intRow)
        frmRoster.GridRoster.TopRow = frmRoster.GridRoster.TopRow + 1
    Loop
    Do While Not frmRoster.GridRoster.ColIsVisible(intCol)
        frmRoster.GridRoster.LeftCol = frmRoster.GridRoster.LeftCol + 1
    Loop
    
    '[SELECT CELL]
    frmRoster.GridRoster.Row = intRow
    frmRoster.GridRoster.Col = intCol
    frmRoster.GridRoster.RowSel = intRow
    frmRoster.GridRoster.ColSel = intCol

End Sub


Public Sub GotoControl()
    
    '[MOVE CONTROL FORM TO THE FRONT]
    frmSet.Show
    frmSet.ZOrder
        
End Sub


Public Sub GotoStaff()

    '[MOVE STAFF FORM TO THE FRONT]
    frmSet.Show
    frmSet.ZOrder

End Sub





Sub AddFieldToTable(DsDynaset, strTableName, strFieldName, intFieldType, intFieldSize, varFieldValue)

    '[THIS ROUTINE ADDS THE DESIGNATED FIELD TO THE PASSED DYNASET]
    '[THE DYNASET NEEDS TO BE CLOSED AND REINITIALISED EVERY TIME]
    Dim fieldNew        As Field
    Dim tableNew        As TableDef
    
    Set tableNew = DBMain.TableDefs(strTableName)
    DsDynaset.Close                         '[CLOSE DYNASET]
    Set fieldNew = tableNew.CreateField(strFieldName, intFieldType, intFieldSize)
    tableNew.Fields.Append fieldNew         '[ADD NEW FIELD]
    '[REOPEN DYNASET]
    Set DsDynaset = DBMain.OpenRecordset(strTableName, dbOpenDynaset)
    If DsDynaset.EOF And DsDynaset.BOF Then
        '[NO RECORDS - NOTHING TO DO HERE]
        '[REV: 3.00.28]
    Else
        DsDynaset.Edit                          '[PLACE NEW VALUE]
            DsDynaset(strFieldName).Value = varFieldValue
        DsDynaset.Update
    End If
    
    '[DEBUG]
    Call LogToFile("* adding field to table : " & strTableName & ":" & strFieldName)
    
End Sub

Sub CheckDatabaseStructure()

    '[SUBROUTINE TO CHECK FOR ESSENTIAL FIELDS IN THE DATABASE AND CREATE THEM]
    '[IF THEY ARE NOT PRESENT - THIS ROUTINE WAS INSTALLED FOR CHANGES BETWEEN]
    '[VERSIONS]
    '[ADD THE DESIGNATED FIELD TO THE DESIGNATED TABLE]
    Dim fieldNew            As Field
    Dim tableNew            As TableDef
    Dim intCounter          As Integer
    Dim strDay              As String
    Dim strClass            As String
    Dim flagRate            As Boolean
    Dim strField            As String
    Dim sinRate             As Single
    
    flagRate = False
    
    '[OPEN REQUIRED DYNASETS]
    Dim DsRosterReport  As Recordset
    Set DsRosterReport = DBMain.OpenRecordset("RosterReport", dbOpenDynaset)
    
    '[DYNASET - Staff    ]
    '[FIELDS  - Holiday, HolStart, HolEnd-----------------]
    '[REV: 3.00.27]
    If FieldInDynaset(DsStaff, "Holiday") = False Then Call AddFieldToTable(DsStaff, "Staff", "Holiday", dbInteger, 0, 0)
    If FieldInDynaset(DsStaff, "HolStart") = False Then Call AddFieldToTable(DsStaff, "Staff", "HolStart", dbDate, 0, Format(Now, "Short Date"))
    If FieldInDynaset(DsStaff, "HolEnd") = False Then Call AddFieldToTable(DsStaff, "Staff", "HolEnd", dbDate, 0, Format(Now, "Short Date"))
    '[REV: 3.00.28]
    If FieldInDynaset(DsStaff, "PayType") = False Then Call AddFieldToTable(DsStaff, "Staff", "PayType", dbInteger, 0, 0)
    If FieldInDynaset(DsDefault, "AllShifts") = False Then Call AddFieldToTable(DsDefault, "Defaults", "AllShifts", dbInteger, 0, 0)
    If FieldInDynaset(DsRosterReport, "TimeFull") = False Then Call AddFieldToTable(DsRosterReport, "RosterReport", "TimeFull", dbText, 25, " ")
    '[REV: 3.00.34]
    If FieldInDynaset(DsStaff, "StaffNote") = False Then Call AddFieldToTable(DsStaff, "Staff", "StaffNote", dbText, 255, " ")
    
    '[CLOSE DYNASETS]
    DsRosterReport.Close
    
End Sub


Function CountReturns(strText) As Integer

    '[FUNCTION TO COUNT THE NUMBER OF RETURN CHARACTERS IN A STRING]
    Dim intCounter          As Integer
    Dim intCount            As Integer
    
    intCount = 0
    
    For intCounter = 1 To Len(strText)
        If Mid$(strText, intCounter, 1) = Chr$(vbKeyReturn) Then intCount = intCount + 1
    Next intCounter

    '[RETURN NUMBER OF RETURNS FOUND]
    CountReturns = intCount

End Function

Function FieldInDynaset(DsCheck As Dynaset, strFieldName) As Boolean
    
    '[FUNCTION TO CHECK IF A FIELD NAME IS PRESENT IN THE PASSED DYNASET]
    Dim intCounter          As Integer
    Dim boolResult          As Boolean
    
    boolResult = False
    
    For intCounter = 0 To (DsCheck.Fields.Count - 1)
        If DsCheck.Fields(intCounter).Name = strFieldName Then boolResult = True
    Next intCounter
    
    FieldInDynaset = boolResult

End Function


Sub FillRosterList()

    '[REFILL COMBO LIST BOX WITH ALL AVAILABLE ROSTERS]
    Dim strClass        As String
    Dim intCounter      As Integer
    Dim intClassIndex   As Integer      '[LOCATION OF SELECTED ITEM IN COMBOCLASS LIST]
    Dim strBookmark     As String
    Dim flagFound       As Boolean
    Dim strFoundClass   As String
    Dim Result
    
    '[SAVE CURRENTLY DISPLAYED CLASS ITEM INDEX]
    strClass = frmRoster.ComboClass.Text
    strFoundClass = ""
    flagFound = True
    
    '[IF NOTHING SELECTED, CHOOSE THE FIRST ITEM]
    If intClassIndex = -1 Then intClassIndex = 0
    
    '[MOVE TO FIRST DYNASET RECORD]
    DsClass.MoveFirst
    
    '[CLEAR COMBO BOX]
    frmRoster.ComboClass.Clear
    
    '[CYCLE THROUGH DYNASET AND FILL COMBO LIST]
    Do While Not DsClass.EOF
        If DsClass("Active") = vbChecked Then
            frmRoster.ComboClass.AddItem DsClass("Description")
            frmRoster.ComboClass.ItemData(frmRoster.ComboClass.NewIndex) = (DsClass.AbsolutePosition + 1)
            If strFoundClass = "" Then strFoundClass = DsClass("Description")
        Else
            If DsClass("Description") = strClass Then flagFound = False
        End If
        DsClass.MoveNext
    Loop
    
    '[RESTORE CURRENTLY DISPLAYED CLASS ITEM INDEX]
    If flagFound = False And strFoundClass = "" Then
        '[NO ACTIVE RECORDS]
        '[SET FIRST RECORD TO ACTIVE]
        DsClass.MoveFirst
        DsClass.Edit
            DsClass("Active") = vbChecked
        DsClass.Update
        DsClass.MoveFirst
        '[ADD RECORD TO LIST AND MOVE TO THIS RECORD]
        frmRoster.ComboClass.AddItem DsClass("Description")
        frmRoster.ComboClass.ItemData(frmRoster.ComboClass.NewIndex) = (DsClass.AbsolutePosition + 1)
        frmRoster.ComboClass.ListIndex = 0
        
        '[DISPLAY MESSAGE BOX]
        Result = MsgBox("GSR requires that one roster is active at all times.  The first roster (" & DsClass("Description") & ") will now be made active.", vbOKOnly & vbExclamation, "Active Roster Required")
        
        '[RESET PICTURE IN CLASS GRID]
        frmSet.GridClass.Row = 1
        frmSet.GridClass.Text = vbChecked
        frmSet.GridClass.Picture = frmSet.ImageSwitch(constWarning).Picture
    ElseIf flagFound = True And strFoundClass > "" Then
        '[FOUND LAST DISPLAYED CLASS]
        Call LocateClass(strClass)
    ElseIf flagFound = False Then
        Call LocateClass(strFoundClass)
    End If
    
    '[NOW LOCATE DYNASET RECORD IN CLASS LIST]
    For intCounter = 0 To (frmRoster.ComboClass.ListCount - 1)
        If frmRoster.ComboClass.List(intCounter) = DsClass("Description") Then frmRoster.ComboClass.ListIndex = intCounter
    Next intCounter
    
End Sub

Sub ShowStaffInfo()

    '[DAYS EMPLOYED AND AGE]
    Dim strFormat As String
    
    frmSet.labelInfo.Caption = ""
    If DsStaff("BirthDate") > 0 Then frmSet.labelInfo.Caption = DsStaff("FirstName") & " " & DsStaff("LastName") & " is " & Format(Date - DsStaff("BirthDate"), "yy \y\e\a\r\s \a\n\d mm \m\o\n\t\h\s") & " old"
    
    If frmSet.labelInfo.Caption = "" Then
        '[EMPLOYMENT PERIOD]
        If DsStaff("DateHired") > 0 Then frmSet.labelInfo.Caption = frmSet.labelInfo.Caption & DsStaff("FirstName") & " " & DsStaff("LastName") & " has been employed for " & Format(Date - DsStaff("DateHired"), "yy \y\e\a\r\s \a\n\d mm \m\o\n\t\h\s") & "."
    Else
        If DsStaff("DateHired") > 0 Then
            frmSet.labelInfo.Caption = frmSet.labelInfo.Caption & " and has been employed for " & Format(Date - DsStaff("DateHired"), "yy \y\e\a\r\s \a\n\d mm \m\o\n\t\h\s") & "."
        Else
            frmSet.labelInfo.Caption = frmSet.labelInfo.Caption & "."
        End If
    End If

End Sub

Sub SizeRosterColumnsToGrid()
    
    '[RESIZE ALL COLUMN WIDTHS TO MATCH GRID WIDTH]
    Dim sinWidth    As Single
    Dim intCounter  As Integer
    Dim intCols     As Integer
    
    '[GRID REDRAW]
    GridRedraw False
    
    sinWidth = frmRoster.GridRoster.Width
    intCols = (frmRoster.GridRoster.Cols - 1)
    
    For intCounter = 0 To intCols
        frmRoster.GridRoster.ColWidth(intCounter) = sinWidth / (intCols + 1)
    Next intCounter
    
    '[GRID REDRAW]
    GridRedraw True

End Sub

Sub StaffReport()

    '[MACHINERY OF PRODUCING STAFF REPORT]

    '[THIS IS THE STAFF DETAILS REPORT SUBROUTINE.]
    '[IT CYCLES THROUGH    ALL STAFF RECORDS AND PRODUCES DETAILS FOR EACH]
    Dim SQLStmt         As String
    Dim strStaffID      As String
    Dim intDayKey       As String
    Dim strBookmark     As String
    Dim strFullname     As String
    Dim flagStart       As Boolean
    Dim intDayCount     As Integer
    Dim intRosterCount  As Integer
    Dim strData         As String
    Dim intLineCount    As Integer
    Dim strDayKey       As String
    Dim strRosterKey    As String
    Dim strField        As String
    Dim DsStaffReport       As Dynaset
        
    '[DELETE EXISTING ROSTER REPORT RECORDS]
    SQLStmt = "DELETE * FROM [StaffRoster]"
    DBMain.Execute SQLStmt, dbFailOnError
        
    '[OPEN STAFF ROSTER  DETAIL REPORT DYNASET]
    Set DsStaffReport = DBMain.OpenRecordset("StaffRoster", dbOpenTable)
    
    '[SET UP ARRAY FOR HOLDING STAFF WAGE DATA]
    Dim StaffReport As StaffReportType
    
    '[STORE CURRENT STAFF LOCATION]
    strBookmark = DsStaff.Bookmark
            
    '[*STAFF LOOP**************************************************************]
    '[MOVE TO FIRST STAFF RECORD]
    DsStaff.MoveFirst
    '[CYCLE THROUGH STAFF LIST]
    Do While Not DsStaff.EOF
        
        strFullname = DsStaff("LastName") & ", " & DsStaff("FirstName")
        strStaffID = DsStaff!StaffID
        '[SHOW PROGRESS REPORT]
        Call ReportInfo(strFullname, 0)
        '[PROGRESS BAR]
        Call ReportProgressBar(((DsStaff.AbsolutePosition + 1) / DsStaff.RecordCount) * 100)
        
        '[ADD LINE TO REPORT GRID]
        For intRosterCount = 1 To 10
            '[FETCH NAME FROM ROSTER]
            DsClass.AbsolutePosition = (intRosterCount - 1)
            '[REV: 3.00.28]
            '[ONLY PROCESS ACTIVE ROSTERS]
            If DsClass!Active = vbChecked Then
                '[CHECK STAFF AVAILABILITY FOR THIS ROSTER NUMBER]
                strRosterKey = "Class_" & Trim(Str(intRosterCount))
                '[ONLY PROCESS IF DAY IS AVAILABLE]
                If DsStaff(strRosterKey).Value = vbChecked Then
                    '[CLEAR DATA LINE]
                    StaffReport.FullName = ""
                    StaffReport.Roster = ""
                    For intDayCount = 1 To 7
                        StaffReport.Day(intDayCount) = ""
                    Next intDayCount
                    
                    '[ADVANCE LINE COUNTER]
                    intLineCount = intLineCount + 1
                    
                    StaffReport.FullName = strFullname
                    StaffReport.Roster = DsClass("Description")
                    
                    For intDayCount = 1 To 7
                        strDayKey = "Day_" & Trim(Str(intDayCount))
                        Select Case DsStaff(strDayKey).Value
                        Case vbChecked
                           StaffReport.Day(intDayCount) = "Yes"
                        Case vbUnchecked
                            StaffReport.Day(intDayCount) = "---"
                        Case vbInside
                            StaffReport.Day(intDayCount) = ">"
                            '[DAY START TIME]
                            strField = "Start_" & Trim(Str(intDayCount))
                            If Not IsNull(DsStaff(strField).Value) Then StaffReport.Day(intDayCount) = StaffReport.Day(intDayCount) & Format(DsStaff(strField), "Short Time")
                            StaffReport.Day(intDayCount) = StaffReport.Day(intDayCount) & "-"
                            '[DAY FINISH TIME]
                            strField = "Finish_" & Trim(Str(intDayCount))
                            If Not IsNull(DsStaff(strField).Value) Then StaffReport.Day(intDayCount) = StaffReport.Day(intDayCount) & Format(DsStaff(strField), "Short Time")
                            StaffReport.Day(intDayCount) = StaffReport.Day(intDayCount) & "<"
                        Case vbOutside
                            StaffReport.Day(intDayCount) = StaffReport.Day(intDayCount) & "<"
                            '[DAY START TIME]
                            strField = "Start_" & Trim(Str(intDayCount))
                            If Not IsNull(DsStaff(strField).Value) Then StaffReport.Day(intDayCount) = StaffReport.Day(intDayCount) & Format(DsStaff(strField), "Short Time")
                            StaffReport.Day(intDayCount) = StaffReport.Day(intDayCount) & "-"
                            '[DAY FINISH TIME]
                            strField = "Finish_" & Trim(Str(intDayCount))
                            If Not IsNull(DsStaff(strField).Value) Then StaffReport.Day(intDayCount) = StaffReport.Day(intDayCount) & Format(DsStaff(strField), "Short Time")
                            StaffReport.Day(intDayCount) = StaffReport.Day(intDayCount) & ">"
                        Case Else
                        End Select
                    Next intDayCount
                    
                    '[ADD LINE TO REPORT]
                    DsStaffReport.AddNew
                        DsStaffReport!FullName = StaffReport.FullName
                        DsStaffReport!Roster = StaffReport.Roster
                        DsStaffReport!StaffID = strStaffID
                        For intDayCount = 1 To 7
                            strDayKey = "Day_" & Trim(Str(intDayCount))
                            DsStaffReport(strDayKey) = StaffReport.Day(intDayCount)
                        Next intDayCount
                    DsStaffReport.Update
                End If
            End If
        Next intRosterCount
        
        '[=====================================================================]
        '[MOVE TO NEXT STAFF RECORD]
        DsStaff.MoveNext
        '[ADD BLANK LINE TO REPORT]
        intLineCount = intLineCount + 1
        'frmReport.GridReport.AddItem "", intLineCount
        
    Loop
    
    '[CLEAR REPORT INFO]
    Call ReportInfo("", 0)
    '[CLEAR PROGRESS BAR]
    Call ReportProgressBar(0)
    
    '[RETURN TO STAFF BOOKMARK]
    DsStaff.Bookmark = strBookmark
    
    '[CLOSE STAFF ROSTER DETAILS REPORT DYNASET]
    DsStaffReport.Close

End Sub

Sub StartupProgressBar(intPercent)

    '[CATCH EXTREME VALUES]
    Dim intCurrPerc     As Integer
    intCurrPerc = Int((frmSplash.gaugeBar.Width / Val(frmSplash.gaugeBar.Tag)) * 100)
    
    '[VER: 3.00.30]
    If intPercent < 0 Then intPercent = 0
    If intPercent > 100 Then intPercent = 100
    
    '[EXIT IF NOT BIG ENOUGH INCREMENT TO SHOW]
    If (intCurrPerc < intPercent - 2) Or (intCurrPerc > intPercent + 2) Then
        frmSplash.gaugeBar.Width = Val(frmSplash.gaugeBar.Tag) * (intPercent / 100)
    End If
    
End Sub

Sub TrimNames(strNames, intWidth)

    '[ROUTINE TO TRIM NAMES PASSED TO THE DESIRED WIDTH]
    Dim strRight            As String
    Dim intLeftPos          As Integer
    Dim intRightPos         As Integer
    Dim intLength           As Integer
    Dim intCounter          As Integer
    Dim intStartPos         As Integer
    Dim strTempNames        As String
    intStartPos = 1
    
    '[CHECKING FOR OTHER NAMES AND PREVIOUS INSTANCE OF NAME]
    If strNames = "" Then
        '[STRING IS EMPTY - EXIT ROUTINE]
        Exit Sub
    Else
        '[STRING IS NOT EMPTY - CHECK FOR NAMES]
        For intCounter = 1 To Len(strNames)
            '[BREAK UP INTO NAMES AND ADD TO strTempNames]
            If Mid$(strNames, intCounter, 1) = Chr$(vbKeyReturn) Then
                If intCounter - intStartPos > intWidth Then
                        strTempNames = strTempNames & Trim(Left$(Mid$(strNames, intStartPos, intCounter), intWidth)) & "."
                Else
                        strTempNames = strTempNames & Trim(Mid$(strNames, intStartPos, intCounter))
                End If
                intStartPos = intCounter
            End If
        Next intCounter
    End If
    '[REMOVE TRAILING RETURN CHARACTER]
    If Right$(strTempNames, 1) = Chr$(vbKeyReturn) Then strTempNames = Left$(strTempNames, Len(strTempNames) - 1)
    strNames = strTempNames
    
End Sub

Function Validate(strValidate) As String

    '[FUNCTION TO VALIDATE A STRING AND RETURN THE VALIDATION CODE]
    Dim strRegCode      As String       '[RETURNED HEX VALIDATION CODE]
    Dim sinValue        As Single       '[ACCUMULATED VALUE]
    Dim intCounter      As Integer      '[COUNTER FOR LENGTH]

    If IsNull(strValidate) Or Len(strValidate) = 0 Then
        '[NO CODE TO VALIDATE SO RETURN BLANK]
        strRegCode = ""
    Else
        For intCounter = 1 To Len(strValidate)
        '[CYCLE THROUGH VALIDATION STRING AND ACCUMULATE VALUES]
            sinValue = sinValue + (Asc(Mid$(strValidate, intCounter, 1)) * sinModifier)
        Next intCounter
        strRegCode = Hex(sinValue)
    End If
    
    '[RETURN CODE VALUE]
    Validate = strRegCode

End Function

Sub exCheckNameInStaffList(strCelltext, intDayCount)

    '[DECLARE VARIABLES REQUIRED]
    Dim strBookmark         As String
    Dim strFullname         As String
    Dim strLastName         As String
    Dim strFirstName        As String
    Dim intStart            As Integer
    Dim intFinish           As Integer
    Dim intBreak            As Integer
    Dim SQLStmt             As String
    Dim flagNameFound       As Boolean
    Dim intClass            As Integer
    Dim strTime             As String
    
    '[SAVE BOOKMARK]
    strBookmark = DsStaff.Bookmark
    
    '[SET FLAG]
    flagNameFound = True
    If Len(Trim(strCelltext)) = 0 Or IsNull(strCelltext) Then Exit Sub
    
    Do While flagNameFound
        '[EXTRACT NAME FROM CELL TEXT AND CHECK FOR IT IN THE STAFF LIST]
        intBreak = InStr(strCelltext, strBreak)
        Select Case intBreak
        Case 0          '[NO BREAK FOUND]
            If Len(Trim(strCelltext)) = 0 Then
                '[RESTORE BOOKMARK]
                DsStaff.Bookmark = strBookmark
                Exit Sub
            End If
            strFullname = Trim$(strCelltext)
            strCelltext = ""
        Case Else       '[BREAK FOUND]
            strFullname = Trim$(Left$(strCelltext, intBreak - 1))
            strCelltext = Trim$(Mid$(strCelltext, intBreak + 1))
        End Select
        
        If InStr(strFullname, ",") > 0 Then strLastName = Trim(Left$(strFullname, InStr(strFullname, ",") - 1))
        If InStr(strFullname, ",") > 0 Then strFirstName = Trim(Mid$(strFullname, InStr(strFullname, ",") + 1))
        
        '[LOCATE LASTNAME, FIRSTNAME IN DYNASET]
        'SQLStmt = "LastName = '" & strLastName & "' AND FirstName = '" & strFirstName & "'"
        'DsStaff.FindFirst SQLStmt
        If LocateStaffName(strLastName, strFirstName) = False Then
        'If DsStaff.NoMatch Then
            '[NAME NOT FOUND IN STAFF LIST, ADD TO EXCEPTION REPORT]
            intClass = DsReport("Class")
            strTime = DsReport("ShiftStart")
            Call exAddNewException(5, intClass, intDayCount, strTime, strFullname, "Name not found in staff list")
        End If
    
    Loop
    
    '[RESTORE BOOKMARK]
    DsStaff.Bookmark = strBookmark

End Sub

Sub FileRead(ReadFileName As String)

    '[CALL ERROR HANDLING ROUTINE]
    On Error GoTo ErrorHandler
    Dim Response            As Integer
    Dim Msg                 As String
    Dim Style               As Integer
    Dim Title               As String
    Dim intRows             As Integer
    Dim intCols             As Integer
    Dim intRowCounter       As Integer
    Dim intColCounter       As Integer
    Dim FileHandle          As Integer
    Dim strClassDesc        As String
    Dim strClassID          As String
    Dim strStartDate        As String
    Dim txtDummy            As String
    Dim varDummy
    Dim frmTemp         As Form
    
    '[SET ACTIVE FORM]
    If flagScratch = True Then
        Set frmTemp = frmScratch
    Else
        Set frmTemp = frmRoster
    End If
    
    '[ALLOCATE FILEHANDLE AND READ SELECTED FILE DETAILS]
    FileHandle = OpenFile(ReadFileName, constFileIn)
        
    '[READ DETAILS FROM INPUT FILE]
    Input #FileHandle, strClassDesc                     '[CLASS DESCRIPTION]
    Input #FileHandle, strClassID                       '[CLASS CODE]
    Input #FileHandle, strStartDate                     '[STARTING DATE OF ROSTER]
    
    '[===================================================================================]
    '[NOW SEE IF WE CAN CHANGE TO THIS ROSTER, OTHERWISE JUST PLACE IN THE CURRENT ROSTER]
    If flagScratch = True Then
        '[NO PLACEMENT FOR SCRATCH ROSTER]
    Else
        If LocateClass(strClassDesc) = True Then
            '[DESCRIPTION FOUND - MOVE TO ROSTER IF ACTIVE]
            If DsClass!Active = 1 Then  '[ROSTER IS ACTIVE]
                frmTemp.ComboClass = strClassDesc                               '[SET DESCRIPTION]
                If IsDate(strStartDate) Then frmTemp.MaskDate = Format(strStartDate, strDateFormat)   '[SET DATE]
            End If
        End If
    End If
    '[===================================================================================]
    
    '[GRID SIZE]
    Input #FileHandle, intRows, intCols

    If intRows < 2 Then intRows = 2
    If intCols < 9 Then intCols = 9

    frmTemp.GridRoster.Rows = intRows
    frmTemp.GridRoster.Cols = intCols
    varDummy = 0
    
    '[WRITE GRID]
    For intRowCounter = 1 To (frmTemp.GridRoster.Rows - 1)
        frmTemp.GridRoster.Row = intRowCounter   '[SET ROW POSITION]
        For intColCounter = 0 To (frmTemp.GridRoster.Cols - 1)
            varDummy = varDummy + 1
            frmTemp.GridRoster.Col = intColCounter   '[SET COL POSITION]
            Input #FileHandle, txtDummy: frmTemp.GridRoster.Text = txtDummy
            '[RESIZE THE ROSTER CELL]
            ResizeRosterCell (txtDummy)
        Next intColCounter
    Next intRowCounter
        
    '[SET SAVE VISIBLE]
    If Not flagScratch Then
        frmRoster.cmdSave.Visible = True
        strRosterFile = ReadFileName
    Else
        frmScratch.Caption = "Scratch Roster - " & ReadFileName
        strScratchFile = ReadFileName
    End If
    
ErrorHandler:
    If Err.Number > 0 Then
        '[ERROR WITH FILE READ]
        Msg = "Error: an error occurred while reading roster data from the file : " & mdiMain.CommonDialog.FileName & strBreak & strBreak & "Error Code: " & Err.Number
        Msg = Msg & strBreak & strBreak & "[" & Err.Description & " in " & Err.Source & "]"
        Style = vbOKOnly                     ' Define buttons.
        Title = "Error Reading from Roster File"
        Response = gsrMsg(Msg, Style, Title)
    End If
    
    '[CLOSE FILE]
    Close #FileHandle
    
End Sub

Public Sub FileSave(SaveFileName As String)
    
    '[CALL ERROR HANDLING ROUTINE]
    On Error GoTo ErrorHandler
    Dim varDummy
    Dim frmTemp         As Form
    Dim Response        As Integer
    Dim Msg             As String
    Dim Style           As Integer
    Dim Title           As String
    
    '[SET ACTIVE FORM]
    If flagScratch = True Then
        Set frmTemp = frmScratch
    Else
        Set frmTemp = frmRoster
    End If
    
    '[SAVE DATA AS PASSED SaveFileName]
    Dim FileHandle      As Integer
    Dim intRowCounter   As Integer
    Dim intColCounter   As Integer
    
    '[GET FILE HANDLE TO OPEN]
    FileHandle = OpenFile(SaveFileName, constFileOut)

    '[WRITE DETAILS TO OUTPUT FILE]
    Write #FileHandle, DsClass("Description")           '[CLASS DESCRIPTION]
    Write #FileHandle, DsClass("Code")                  '[CLASS CODE]
    Write #FileHandle, Str$(DsDefault("StartDate"))       '[STARTING DATE OF ROSTER]
    '[GRID SIZE]
    Write #FileHandle, frmTemp.GridRoster.Rows, frmTemp.GridRoster.Cols
    '[WRITE GRID]
    varDummy = 0
    For intRowCounter = 1 To (frmTemp.GridRoster.Rows - 1)
        frmTemp.GridRoster.Row = intRowCounter   '[SET ROW POSITION]
        For intColCounter = 0 To (frmTemp.GridRoster.Cols - 1)
            varDummy = varDummy + 1
            frmTemp.GridRoster.Col = intColCounter   '[SET COL POSITION]
            If intColCounter < (frmTemp.GridRoster.Cols - 1) Then
                Write #FileHandle, frmTemp.GridRoster.Text;
            Else
                Write #FileHandle, frmTemp.GridRoster.Text
            End If
        Next intColCounter
    Next intRowCounter
    
    '[SET SCRATCH ROSTER CAPTION]
    If flagScratch Then
        frmScratch.Caption = "Scratch Roster - " & SaveFileName
        strScratchFile = SaveFileName
    Else
        strRosterFile = SaveFileName
    End If
        
ErrorHandler:
    If Err.Number > 0 Then
        '[ERROR WITH FILE SAVE]
        Msg = "Error: GSR cannot save to the file : " & mdiMain.CommonDialog.FileName & strBreak & strBreak & "Error Code: " & Err.Number
        Msg = Msg & strBreak & strBreak & "[" & Err.Description & " in " & Err.Source & "]"
        Style = vbOKOnly                     ' Define buttons.
        Title = "Cannot Save Roster to File"
        Response = gsrMsg(Msg, Style, Title)
    End If

    '[CLOSE FILE]
    Close #FileHandle

End Sub

Sub FileSetFilter()

    Dim intClassCounter         As Integer
    Dim strBookmark             As String

    '[default file extension]
    mdiMain.CommonDialog.DefaultExt = Trim(DsClass("Code"))
    
    '[save class bookmark]
    strBookmark = DsClass.Bookmark
    
    '[Set Filter Property]
    mdiMain.CommonDialog.Filter = ""
    For intClassCounter = 0 To 9
        DsClass.AbsolutePosition = intClassCounter
        '[ONLY ALLOW SAVE/LOAD TO ACTIVE ROSTERS]
        If DsClass("Active") = vbChecked Then
            '[ADD ACTIVE CLASSES TO FILTER LIST}
            mdiMain.CommonDialog.Filter = mdiMain.CommonDialog.Filter & Trim(DsClass("Description")) & " (*." & Trim(DsClass("Code")) & ")"
            mdiMain.CommonDialog.Filter = mdiMain.CommonDialog.Filter & "|*." & Trim(DsClass("Code")) & "|"
            '[SET FILTER INDEX IF A MATCH IS FOUND]
            If mdiMain.CommonDialog.DefaultExt = Trim(DsClass("Code")) Then mdiMain.CommonDialog.FilterIndex = (intClassCounter + 1)
        End If
    Next intClassCounter
    
    '[add standard file types to end of filter list]
    mdiMain.CommonDialog.Filter = mdiMain.CommonDialog.Filter & "Text files (*.txt)|*.txt|All files (*.*)|*.*"
    '[restore class bookmark]
    DsClass.Bookmark = strBookmark

End Sub

Function LineCount(strText) As Integer

    '[FUNCTION TO COUNT NUMBER OF RETURN CHARACTERS WITHIN A GIVEN STRING]
    Dim intLineCount            As Integer
    Dim intPos                  As Integer
    
    intPos = 1
    '[STOP NULL ERROR]
    strText = Trim(strText & "")
    If Len(strText) = 0 Then intLineCount = 0

    Do While intPos > 0
        intPos = InStr(intPos + 1, strText, strBreak)
        intLineCount = intLineCount + 1
    Loop
    
    '[RETURN VALUE]
    LineCount = intLineCount
    
End Function

Public Function OpenFile(FileName As String, FileMode As Integer) As Integer

    '[FUNCTION TO OPEN A FILE AND RETURN THE FILE HANDLE ASSOCIATED WITH THE FILE]
    '[FILEMODE 0=INPUT]
    '[FILEMODE 1=OUTPUT]
    Dim FileHandle As Integer '[Next Free File Handle]
    Dim Result
    FileHandle = FreeFile(0)  '[Allocate free handle 1-255]

    Select Case FileMode
        Case constFileIn
            Open FileName For Input As #FileHandle      '[OPEN FILE FOR INPUT]
        Case constFileOut
            Open FileName For Output As #FileHandle     '[OPEN FILE FOR OUTPUT]
        Case Else
    End Select

    OpenFile = FileHandle         '[RETURN FILE HANDLE]

End Function

Public Function FileSaveAs(NewFileName As String) As Boolean
    
    '[CALL ERROR HANDLING ROUTINE]
    On Error GoTo ErrorHandler

    '[FUNCTION TO SAVE FILE AS A NEW NAME AND RETURN CODE INDICATING WHETHER CANCEL WAS PRESSED]

    '[SAVE CURRENT DATA AS A NEW FILE]
    mdiMain.CommonDialog.DialogTitle = "Save Roster As"
    mdiMain.CommonDialog.FileName = NewFileName
    mdiMain.CommonDialog.CancelError = True
    
    '[SET FILE DIALOG FLAGS]
    mdiMain.CommonDialog.Flags = cdlOFNOverwritePrompt Or cdlOFNPathMustExist Or cdlOFNHideReadOnly
    mdiMain.CommonDialog.ShowSave
    
    '[PROCESS COMMON DIALOG SAVE FORM]
    Call FileSave(mdiMain.CommonDialog.FileName)
    FileSaveAs = True

ErrorHandler:
    If Err.Number = cdlCancel Then
        FileSaveAs = False
    End If

End Function

Sub procAllStaffRosters()
    
    '[ERROR HANDLER ROUTINE]
    On Error GoTo ErrorHandler
    
    '[THIS ROUTINE WILL CALL THE APPROPRIATE SUBROUTINE FOR ALL STAFF MEMBERS]
    '[AND DISPLAY ANY WARNING MESSAGES]
    Dim Msg
    Dim Style
    Dim Title
    Dim Response
    Dim strFullname         As String
    Dim strBookmark         As String
    Dim intCounter          As Integer
    Dim flagFound           As Boolean
    
    '[CANNOT PRINT IF NO STAFF ARE DEFINED]
    If (DsStaff.EOF And DsStaff.BOF) Or (DsStaff.RecordCount = 1 And DsStaff!LastName = "*LastName") Then
        Msg = "There are no staff records entered at present." & strBreak & strBreak & "Please enter staff details before producing timesheets."
        Style = vbOKOnly                     ' Define buttons.
        Title = "Cannot Create Timesheet"
        Response = gsrMsg(Msg, Style, Title)
        Exit Sub
    End If

    '[SAVE STAFF DYNASET POSITION]
    strBookmark = DsStaff.Bookmark
    
    '[MOVE TO FIRST RECORD]
    DsStaff.MoveFirst
    
    '[SHOW WARNING FORM]
    Msg = "The full staff roster report prints weekly timesheets for all staff members (" & DsStaff.RecordCount & " records).  You will be able to print the weekly timesheets from the report screen." & strBreak & strBreak & strBreak & strBreak & "Do you wish to continue and print the reports ?"
    Style = vbYesNo                     ' Define buttons.
    Title = "Confirmation Required"
    Response = gsrMsg(Msg, Style, Title)
    If Response = vbYes Then
        
        '[SHOW GSR MESSAGE FORM]
        Msg = "GSR is now processing your roster files and producing the required staff timesheets for all staff who appear in active rosters." & strBreak & strBreak & "This may take a few minutes."
        Style = vbInformation            ' Define buttons.
        Title = "Staff Roster"
        Response = gsrMsg(Msg, Style, Title)
        frmMsg.gaugeBar.Visible = True
        frmMsg.gaugeBorder.Visible = True
        frmMsg.labelInfo.Visible = True
        frmMsg.ZOrder
        frmMsg.Refresh
        
        '[INITIALISE STAFF ROSTER SET]
        InitStaffRosterSet

        '[SET UP LOOP]
        Do While Not DsStaff.EOF
            '[INCREMENT COUNTER]
            intCounter = intCounter + 1
            '[APPLY STAFF NAME TO FULL STRING]
            strFullname = DsStaff("LastName") & ", " & DsStaff("FirstName")
            '[SHOW PROGRESS REPORT]
            Call ReportInfo(strFullname, 0)
            '[PROGRESS BAR]
            Call ReportProgressBar((intCounter / DsStaff.RecordCount) * 100)
            
            '[CALL STAFF ROSTER SUBROUTINE FOR THE HIGHLIGHTED STAFF MEMBER]
            Call prnStaffRoster(strFullname, flagFound)
            
            '[MOVE TO NEXT RECORD]
            DsStaff.MoveNext
        Loop
        
        '[CLOSE STAFF ROSTER SET]
        CloseStaffRosterSet
        
        '[HIDE GSR MESSAGE FORM]
        Unload frmMsg
        
        '[SET REPORT TITLE]
        Select Case prnOrient
        Case vbPRNone           '[NO PRINTER ATTACHED]
            strReport = "p_stftim.rpt"
        Case vbPRORPortrait     '[PORTRAIT STYLE]
            strReport = "p_stftim.rpt"
        Case vbPRORLandscape    '[LANDSCAPE STYLE]
            strReport = "l_stftim.rpt"
        Case Else
            strReport = "p_stftim.rpt"
        End Select
        
        mdiMain.Report.ReportFileName = strReport
        
        '[SET REPORT TITLE]
    
    
        mdiMain.Report.WindowTitle = "Weekly Staff Rosters for All Staff, Week Starting " & Format(DsDefault("StartDate"), strDateFormat)
        mdiMain.Report.SortFields(0) = "+{StaffRoster.FullName}"
        mdiMain.Report.Formulas(0) = ""
        mdiMain.Report.Action = 1
        
    End If

    '[RESTORE STAFF DYNASET POSITION]
    DsStaff.Bookmark = strBookmark

ErrorHandler:
    If Err.Number > 0 Then
        '[CANNOT FIND ROSTER FILE ?]
        Msg = "Error: GSR cannot produce this report." & strBreak & strBreak & "The required report file (" & strReport & ") may be missing from the GSR directory." & strBreak & strBreak & "Error Code: " & Err.Number
        If Err.Number = 20504 Then
            Msg = Msg & strBreak & strBreak & "[Could not find required report file]"
        Else
            Msg = Msg & strBreak & strBreak & "[" & Err.Description & " in " & Err.Source & "]"
        End If
        Style = vbOKOnly                     ' Define buttons.
        Title = "Cannot Produce Report"
        Response = gsrMsg(Msg, Style, Title)
        Exit Sub
    End If

End Sub

Sub procExceptionReport()
        
    '[ERROR HANDLER]
    On Error GoTo ErrorHandler
    
    '[COMMAND TO CHECK ROSTER FOR -ANY- IRREGULARITIES AND REPORT THEM TO THE REPORT FORM]
    '[USE THE REPORT FORM TO ADDRESS ANY OF THESE PROBLEMS]
    Dim Msg
    Dim Style
    Dim Title
    Dim Response

    '[CANNOT PRINT IF NO STAFF ARE DEFINED]
    If (DsStaff.EOF And DsStaff.BOF) Or (DsStaff.RecordCount = 1 And DsStaff!LastName = "*LastName") Then
        Msg = "There are no staff records entered at present." & strBreak & strBreak & "Please enter staff details on the staff form before executing this report."
        Style = vbOKOnly                     ' Define buttons.
        Title = "Cannot Create Report"
        Response = gsrMsg(Msg, Style, Title)
        Exit Sub
    End If

    '[SHOW WARNING FORM]
    Msg = "The exception report details any problems which may occur within your currently defined and active rosters." & strBreak & strBreak & "Because this routine has to perform multiple comparisions and searches, it may take a few minutes to complete, depending upon the number of rosters, staff and the speed of your computer." & strBreak & strBreak & "Do you wish to continue and produce the report ?"
    Style = vbYesNo                     ' Define buttons.
    Title = "Confirmation Required"
    Response = gsrMsg(Msg, Style, Title)
    If Response = vbYes Then
        '[SHOW PLEASE WAIT MESSAGE]
        '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
        StatusBar "Please wait while this report is processed."
        '[---------------------------------------------------------------------------------]
        mdiMain.panelStatusBar.Refresh
        
        '[SHOW GSR MESSAGE FORM]
        Msg = "GSR is now processing your roster files and producing an exception report." & strBreak & strBreak & "The report will detail which problems (if any) have been found with your rosters." & strBreak & strBreak & "This may take a few minutes."
        Style = vbInformation            ' Define buttons.
        Title = "Exception Report"
        Response = gsrMsg(Msg, Style, Title)
        frmMsg.gaugeBar.Visible = True
        frmMsg.gaugeBorder.Visible = True
        frmMsg.labelInfo.Visible = True
        frmMsg.ZOrder
        frmMsg.Refresh
        
        '[CALL EXCEPTION REPORT SUBROUTINE]
        ExceptionReport
        
        '[HIDE GSR MESSAGE FORM]
        Unload frmMsg
        
            
        '[SET REPORT TITLE]
        Dim intOrient           As Integer
        intOrient = vbPRNone    '[SET TO NO PRINTER]
        
        Select Case prnOrient
        Case vbPRNone           '[NO PRINTER ATTACHED]
            strReport = "p_except.rpt"
        Case vbPRORPortrait     '[PORTRAIT STYLE]
            strReport = "p_except.rpt"
        Case vbPRORLandscape    '[LANDSCAPE STYLE]
            intOrient = prnOrient
            Printer.Orientation = vbPRORPortrait
            strReport = "p_except.rpt"
        Case Else
            intOrient = prnOrient
            Printer.Orientation = vbPRORPortrait
            strReport = "p_except.rpt"
        End Select
        
        mdiMain.Report.ReportFileName = strReport
          
        '[SET REPORT TITLE]
        Printer.Orientation = vbPRORPortrait
    
    
        mdiMain.Report.Formulas(0) = ""
        mdiMain.Report.WindowTitle = "Exception Report, Week Starting " & Format(DsDefault("StartDate"), strDateFormat)
        mdiMain.Report.SortFields(0) = "+{Exception.FullName}"
        mdiMain.Report.Action = 1
        
        '[RESTORE PRINTER ORIENTATION IF CHANGED]
        If intOrient > 0 Then Printer.Orientation = intOrient

    End If

ErrorHandler:
    If Err.Number > 0 Then
        '[CANNOT FIND ROSTER FILE ?]
        Msg = "Error: GSR cannot produce this report." & strBreak & strBreak & "The required report file (" & strReport & ") may be missing from the GSR directory." & strBreak & strBreak & "Error Code: " & Err.Number
        If Err.Number = 20504 Then
            Msg = Msg & strBreak & strBreak & "[Could not find required report file]"
        Else
            Msg = Msg & strBreak & strBreak & "[" & Err.Description & " in " & Err.Source & "]"
        End If
        Style = vbOKOnly                     ' Define buttons.
        Title = "Cannot Produce Report"
        Response = gsrMsg(Msg, Style, Title)
        Exit Sub
    End If

End Sub

Sub procCurrentRoster()
    
    '[ERROR HANDLER ROUTINE]
    'On Error GoTo ErrorHandler
    
    '[USE THE REPORT FORM TO ADDRESS ANY OF THESE PROBLEMS]
    Dim Msg
    Dim Style
    Dim Title
    Dim Response
    
    '[CANNOT PRINT IF ROSTER IS EMPTY]
    If DsRoster.EOF And DsRoster.BOF Then
        Msg = frmRoster.ComboClass.Text & " does not contain any information or you have not saved any changes to the current roster." & strBreak & strBreak & "GSR cannot print a blank roster."
        Style = vbOKOnly                     ' Define buttons.
        Title = "Cannot Print Blank Roster"
        Response = gsrMsg(Msg, Style, Title)
        Exit Sub
    End If
    
    '[SHOW WARNING FORM]
    Msg = "Continue and print the " & frmRoster.ComboClass.Text & " roster ?"
    Style = vbYesNo                     ' Define buttons.
    Title = "Confirmation Required"
    Response = gsrMsg(Msg, Style, Title)
    If Response = vbYes Then
        '[SHOW GSR MESSAGE FORM]
        Msg = "GSR is now printing your roster." & strBreak & strBreak & "Please wait."
        Style = vbInformation            ' Define buttons.
        Title = "Printing Roster"
        Response = gsrMsg(Msg, Style, Title)
        frmMsg.gaugeBar.Visible = True
        frmMsg.gaugeBorder.Visible = True
        frmMsg.labelInfo.Visible = True
        frmMsg.ZOrder
        frmMsg.Refresh
        
        '[PRINT ROSTER TO REPORT FORM HERE]
        Call prnRoster
        
        '[HIDE GSR MESSAGE FORM]
        Unload frmMsg
    
        '[SET REPORT TITLE]
        Select Case prnOrient
        Case vbPRNone           '[NO PRINTER ATTACHED]
            strReport = "p_roster.rpt"
        Case vbPRORPortrait     '[PORTRAIT STYLE]
            strReport = "p_roster.rpt"
        Case vbPRORLandscape    '[LANDSCAPE STYLE]
            strReport = "l_roster.rpt"
        Case Else
            strReport = "p_roster.rpt"
        End Select
        
        mdiMain.Report.ReportFileName = strReport
        mdiMain.Report.Formulas(0) = ""
    
    
        mdiMain.Report.WindowTitle = frmRoster.ComboClass.Text & " Roster, Week Starting " & Format(DsDefault("StartDate"), strDateFormat)
        mdiMain.Report.SortFields(0) = ""
        mdiMain.Report.Action = 1
        
    End If
    
ErrorHandler:
    If Err.Number > 0 Then
        '[CANNOT FIND ROSTER FILE ?]
        Msg = "Error: GSR cannot produce this report." & strBreak & strBreak & "The required report file (" & strReport & ") may be missing from the GSR directory." & strBreak & strBreak & "Error Code: " & Err.Number
        If Err.Number = 20504 Then
            Msg = Msg & strBreak & strBreak & "[Could not find required report file]"
        Else
            Msg = Msg & strBreak & strBreak & "[" & Err.Description & " in " & Err.Source & "]"
        End If
        Style = vbOKOnly                     ' Define buttons.
        Title = "Cannot Produce Report"
        Response = gsrMsg(Msg, Style, Title)
        Exit Sub
    End If

End Sub

Sub procSelectedStaffRoster()

    '[ERROR HANDLER ROUTINE]
    On Error GoTo ErrorHandler

    '[THIS ROUTINE WILL CALL THE APPROPRIATE SUBROUTINE FOR A SINGLE STAFF MEMBER]
    '[AND DISPLAY ANY WARNING MESSAGES]
    Dim Msg
    Dim Style
    Dim Title
    Dim Response
    Dim strFullname         As String
    Dim flagFound           As Boolean
    
    '[CANNOT PRINT IF NO STAFF ARE DEFINED]
    If (DsStaff.EOF And DsStaff.BOF) Or (DsStaff.RecordCount = 1 And DsStaff!LastName = "*LastName") Then
        Msg = "There are no staff records entered at present." & strBreak & strBreak & "Please enter staff details before producing timesheets."
        Style = vbOKOnly                     ' Define buttons.
        Title = "Cannot Create Timesheet"
        Response = gsrMsg(Msg, Style, Title)
        Exit Sub
    End If
    
    '[APPLY STAFF NAME TO FULL STRING]
    strFullname = DsStaff("LastName") & ", " & DsStaff("FirstName")
    
    '[SHOW WARNING FORM]
    Msg = "The staff roster report prints weekly timesheet details for the selected staff member (" & strFullname & ").  You will be able to print the weekly timesheet for " & strFullname & " from the report screen." & strBreak & strBreak & strBreak & strBreak & "Do you wish to continue and print the report ?"
    Style = vbYesNo                     ' Define buttons.
    Title = "Confirmation Required"
    Response = gsrMsg(Msg, Style, Title)
    If Response = vbYes Then
        
        '[SHOW GSR MESSAGE FORM]
        Msg = "GSR is now processing your roster files and producing the required staff timesheet for " & strFullname & "." & strBreak & strBreak & "This may take a few minutes."
        Style = vbInformation            ' Define buttons.
        Title = "Staff Roster"
        Response = gsrMsg(Msg, Style, Title)
        frmMsg.gaugeBorder.Visible = True
        frmMsg.gaugeBar.Visible = True
        frmMsg.labelInfo.Visible = True
        frmMsg.ZOrder
        frmMsg.Refresh
        
        '[INITIALISE STAFF ROSTER SET]
        InitStaffRosterSet
        
        '[CALL STAFF ROSTER SUBROUTINE FOR THE HIGHLIGHTED STAFF MEMBER]
        Call prnStaffRoster(strFullname, flagFound)
        
        '[CLOSE STAFF ROSTER SET]
        CloseStaffRosterSet
        
        '[HIDE GSR MESSAGE FORM]
        Unload frmMsg
    
        If flagFound = True Then
            '[SET REPORT TITLE]
            Select Case prnOrient
            Case vbPRNone           '[NO PRINTER ATTACHED]
                strReport = "p_stftim.rpt"
            Case vbPRORPortrait     '[PORTRAIT STYLE]
                strReport = "p_stftim.rpt"
            Case vbPRORLandscape    '[LANDSCAPE STYLE]
                strReport = "l_stftim.rpt"
            Case Else
                strReport = "p_stftim.rpt"
            End Select
            
            mdiMain.Report.ReportFileName = strReport

            '[SET REPORT TITLE]
        
        
            mdiMain.Report.WindowTitle = "Weekly Staff Roster for " & strFullname & ", Week Starting " & Format(DsDefault("StartDate"), strDateFormat)
            mdiMain.Report.SortFields(0) = "+{StaffRoster.FullName}"
            mdiMain.Report.Formulas(0) = ""
            mdiMain.Report.Action = 1
        Else
            Msg = "No roster allocations were found for the selected staff member (" & strFullname & "). " & strBreak & strBreak & "No timesheet will be produced." & strBreak & strBreak & "Click OK to continue."
            Style = vbOKOnly                     ' Define buttons.
            Title = "No Roster Allocation Found"
            Response = gsrMsg(Msg, Style, Title)
        End If
    End If


ErrorHandler:
    If Err.Number > 0 Then
        '[CANNOT FIND ROSTER FILE ?]
        Msg = "Error: GSR cannot produce this report." & strBreak & strBreak & "The required report file (" & strReport & ") may be missing from the GSR directory." & strBreak & strBreak & "Error Code: " & Err.Number
        If Err.Number = 20504 Then
            Msg = Msg & strBreak & strBreak & "[Could not find required report file]"
        Else
            Msg = Msg & strBreak & strBreak & "[" & Err.Description & " in " & Err.Source & "]"
        End If
        Style = vbOKOnly                     ' Define buttons.
        Title = "Cannot Produce Report"
        Response = gsrMsg(Msg, Style, Title)
        Exit Sub
    End If

End Sub

Sub prnStaffRoster(strFullname, flagFound)

    '[THIS IS THE STAFF ROSTER SUBROUTINE. IT WILL CREATE A TEMP DYNASET        ]
    '[TO CONTAIN ALL OF THE ROSTER RECORDS AND CYCLE THROUGH USING SQL          ]
    '[STATEMENTS (HOPEFULLY).                                                   ]
    
    '[BUILD ROSTER DYNASET WITH ALL RECORDS]
    Dim intAddedCount   As Integer
    Dim SQLStmt         As String
    Dim strDayKey       As String
    Dim strCriteria     As String
    Dim intCounter      As Integer
    Dim intDayCount     As Integer
    Dim strClassDesc    As String
    Dim sinMinsWorked   As Single
    Dim sinIncrement    As Single
    Dim intInterval     As Integer
    Dim intRosterCount  As Integer
    Dim intStaffCount   As Integer
    Dim dateStart       As Date
    Dim dateEnd         As Date
    Dim dateDay         As Date
    Dim strStaffID      As String
    Dim strDayStart     As String
    Dim strDayFinish    As String
    Dim strName         As String
    Dim strRoster       As String
    Dim sinMinutes      As Single
    Dim sinTotal        As Single
    Dim sinTotalAmount  As Single
    Dim sinTotalMinutes As Single
    Dim strShortDay     As String
    Dim strLastDay      As String
    Dim strNote         As String
    Dim intPlace        As Integer
    Dim Result
    
    '[SET UP ARRAY FOR HOLDING STAFF ROSTER DATA]
    flagFound = False
    '[(DAY * SHIFT)]
    Dim arrayRoster(7, 10)  As StaffType
    '[ARRAY FOR COUNTING SHIFTS ON THIS DAY]
    Dim arrayCount(7)       As Integer
    
    '[SELECT APPROPRIATE RECORDS]
    SQLStmt = "SELECT * FROM Roster ORDER BY CLASS, SHIFTSTART"
    Set DsReport = DBMain.OpenRecordset(SQLStmt, dbOpenDynaset)
    
    '[CHECK CURRENT DYNASET AND PREPARE]
    If DsRoster.EOF And DsRoster.BOF Then
        DsReport.Close
        Exit Sub
    End If

    '[COUNT NUMBER OF TIMES STAFF MEMBER APPEARS IN ROSTER]
    intStaffCount = CountStaffInRoster(strFullname, DsReport)
    
    '[*STAFF LOOP**************************************************************]
    '[RESET MINUTES WORKED]
    Erase arrayRoster
    Erase arrayCount
    '[SHOW PROGRESS REPORT]
    Call ReportInfo(strFullname, 0)
                    
    '[=ROSTER LOOP=========================================================]
    '[MOVE TO FIRST RECORD]
    DsReport.MoveFirst
    '[CYCLE THROUGH ROSTER]
    Do While Not DsReport.EOF
        
        '[CHECK TO SEE IF THE ROSTER IS ACTIVE]
        DsClass.AbsolutePosition = (DsReport("Class") - 1)
        If DsClass("Active") = vbChecked Then
            For intDayCount = 1 To 7
                '[-=-NOW CHECK EACH DAY TO SEE IF STAFF MEMBER IS INCLUDED IN ANY DAY-=-]
                strDayKey = "Day_" & Trim(Str(intDayCount))
                
                If InStr(DsReport(strDayKey), strFullname) > 0 Then
                    '[SET FLAG TO TRUE]
                    flagFound = True
                    '[INCREMENT COUNTER FOR THIS DAY]
                    arrayCount(intDayCount) = arrayCount(intDayCount) + 1
                    '[RESET INCREMENT - ALLOWS FOR ROSTERS WITH DIFFERING INCREMENTS]
                    If IsNull(DsReport("ShiftStart")) Then
                        DsReport.Edit
                            DsReport("ShiftStart") = DsDefault("StartTime")
                        DsReport.Update
                    End If
                    dateStart = DsReport("ShiftStart")
                    If IsNull(DsReport("ShiftEnd")) Then
                        DsReport.Edit
                            DsReport("ShiftEnd") = DsDefault("EndTime")
                        DsReport.Update
                    End If
                    dateEnd = DsReport("ShiftEnd")
                    '[ALLOW FOR NEXT DAY TIMES]
                    If dateEnd < dateStart Then dateEnd = dateEnd + CDate("12:00") + CDate("12:00")
                    '[ALLOW FOR 24 HOUR TIMES]
                    If dateEnd = dateStart Then dateEnd = dateStart + CDate("12:00") + CDate("12:00")
                    '[CALCULATE INCREMENT]
                    sinIncrement = (dateEnd - dateStart) * (24 * 60)
                    
                    '[CHECK TO SEE IF INCREMENT IS > WORK BLOCK, IF SO SUBTRACT BREAK TIME]
                    '[ALSO CHECK FOR NO WORK PERIOD SET (00:00)]
                    If ((DsDefault("BlockHour") * 60) + DsDefault("BlockMin")) > 0 And sinIncrement >= ((DsDefault("BlockHour") * 60) + DsDefault("BlockMin")) Then
                        sinIncrement = sinIncrement - ((DsDefault("BreakHour") * 60) + DsDefault("BreakMin"))
                    End If
                    
                    '[ADD MINUTES TO STAFF ROSTER ARRAY]
                    arrayRoster(intDayCount, arrayCount(intDayCount)).Minutes = arrayRoster(intDayCount, arrayCount(intDayCount)).Minutes + sinIncrement
                    '[ADD START TIME, END TIME TO STAFF ROSTER ARRAY]
                    arrayRoster(intDayCount, arrayCount(intDayCount)).StartTime = DsReport("ShiftStart")
                    arrayRoster(intDayCount, arrayCount(intDayCount)).EndTime = DsReport("ShiftEnd")
                    '[ADD ROSTER DESCRIPTION]
                    arrayRoster(intDayCount, arrayCount(intDayCount)).Roster = DsClass("Description")
                    '[ADD STARTING DATE]
                    arrayRoster(intDayCount, arrayCount(intDayCount)).DayDate = DsDefault("StartDate") + (intDayCount - 1)
                End If
                
            Next intDayCount
        End If
        
        '[MOVE TO NEXT ROSTER RECORD]
        DsReport.MoveNext
        
    Loop
    '[ADD STAFF ROSTER RECORDS TO DYNASET]
    strStaffID = DsStaff("StaffID")
   
    '[=====================================================================]
    '[IF NO DATA FOUND, EXIT THE ROUTINE]
    If flagFound = False Then Exit Sub
   
    '[=====================================================================]
    intAddedCount = 0
    
        '[CYCLE THROUGH ROSTERS]
        For intDayCount = 1 To 7
            
            '[ASSIGN ARRAY DATA TO SET VARIABLES]
            strDayKey = "Day_" & Trim(Str(intDayCount))
            
            For intCounter = 1 To arrayCount(intDayCount)
                
                '[CYCLE THROUGH SHIFTS]
                sinMinutes = 0
                sinMinutes = arrayRoster(intDayCount, intCounter).Minutes
                strShortDay = ArrayWeek(DayOfWeek(intDayCount)).ShortDay
                
                dateDay = arrayRoster(intDayCount, intCounter).DayDate
            
                strRoster = arrayRoster(intDayCount, intCounter).Roster
                strCriteria = "[FullName]='" & strFullname & "' AND [Roster]='" & strRoster & "' AND [" & strDayKey & "]=NULL"
                
                '[LOCATE FIRST OCCURENCE OF THIS NAME AND ROSTER]
                DsStaffRoster.FindFirst strCriteria
                
                '[CHECK TO SEE IF DAY ALREADY HAS A TIME]
                If DsStaffRoster.NoMatch Then
                    intAddedCount = intAddedCount + 1
                    DsStaffRoster.AddNew
                    DsStaffRoster("Fullname") = strFullname
                    DsStaffRoster("StaffID") = strStaffID
                    DsStaffRoster("Roster") = strRoster
                    DsStaffRoster!Hours = sinMinutes / 60
                Else
                    DsStaffRoster.Edit
                    DsStaffRoster!Hours = DsStaffRoster!Hours + sinMinutes / 60
                End If
                '[UPDATE RECORD WITH CHANGES]
                DsStaffRoster(strDayKey) = Format(arrayRoster(intDayCount, intCounter).StartTime, "Medium Time") & "-" & Format(arrayRoster(intDayCount, intCounter).EndTime, "Medium Time")
                        
                DsStaffRoster.Update
            Next intCounter
            
        Next intDayCount
    
    For intCounter = intAddedCount To 9
        DsStaffRoster.AddNew
        DsStaffRoster("Fullname") = strFullname
        DsStaffRoster("StaffID") = strStaffID
        DsStaffRoster.Update
    Next intCounter
    '[=====================================================================]
    
    '[CLOSE TEMPORARY DYNASET]
    DsReport.Close

End Sub

Function CountStaffInRoster(strFullname, DsReport) As Integer

    '[COUNT OCCURENCES OF STAFF IN ACTIVE ROSTERS AND RETURN]
    Dim intStaffCount   As Integer
    Dim intDayCount     As Integer
    Dim strDayKey       As String
    
    intStaffCount = 0
    
    DsReport.MoveFirst
    Do While Not DsReport.EOF
        
        '[CHECK TO SEE IF THE ROSTER IS ACTIVE]
        DsClass.AbsolutePosition = (DsReport("Class") - 1)
        If DsClass("Active") = vbChecked Then
            For intDayCount = 1 To 7
                '[-=-NOW CHECK EACH DAY TO SEE IF STAFF MEMBER IS INCLUDED IN ANY DAY-=-]
                strDayKey = "Day_" & Trim(Str(intDayCount))
                If InStr(DsReport(strDayKey), strFullname) > 0 Then
                    intStaffCount = intStaffCount + 1
                End If
            Next intDayCount
        End If
        
        '[MOVE TO NEXT ROSTER RECORD]
        DsReport.MoveNext
    Loop
        
    DsReport.MoveFirst
    CountStaffInRoster = intStaffCount
        
End Function

Sub procRosterReport()
    
    '[ERROR HANDLER ROUTINE]
    On Error GoTo ErrorHandler
    
    '[COMMAND TO PROCESS EACH STAFF MEMBER AND COLLECT INFO ABOUT EACH
    Dim Msg
    Dim Style
    Dim Title
    Dim Response
    
    '[CANNOT PRINT IF NO STAFF ARE DEFINED]
    If (DsStaff.EOF And DsStaff.BOF) Or (DsStaff.RecordCount = 1 And DsStaff!LastName = "*LastName") Then
        Msg = "There are no staff records entered at present." & strBreak & strBreak & "Please enter staff details on the staff form before executing this report."
        Style = vbOKOnly                     ' Define buttons.
        Title = "Cannot Create Report"
        Response = gsrMsg(Msg, Style, Title)
        Exit Sub
    End If
    
    '[SHOW WARNING FORM]
    Msg = "The Roster Details Report lists roster, hour and cost figures for each staff member in your staff list." & strBreak & strBreak & "Because this routine has to perform multiple comparisions and searches, it may take a few minutes to complete, depending upon the number of rosters, staff and the speed of your computer." & strBreak & strBreak & "Do you wish to continue and produce the report ?"
    Style = vbYesNo                     ' Define buttons.
    Title = "Confirmation Required"
    Response = gsrMsg(Msg, Style, Title)
    If Response = vbYes Then
        '[SHOW PLEASE WAIT MESSAGE]
        '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
        StatusBar "Please wait while this report is processed."
        '[---------------------------------------------------------------------------------]
        mdiMain.panelStatusBar.Refresh
        
        '[SHOW GSR MESSAGE FORM]
        Msg = "GSR is now processing your roster files and producing the roster detail report." & strBreak & strBreak & "The report will list hours and costs for all staff members who have been assigned to rosters." & strBreak & strBreak & "This may take a few minutes."
        Style = vbInformation            ' Define buttons.
        Title = "Roster Details Report"
        Response = gsrMsg(Msg, Style, Title)
        frmMsg.gaugeBar.Visible = True
        frmMsg.gaugeBorder.Visible = True
        frmMsg.labelInfo.Visible = True
        frmMsg.ZOrder
        frmMsg.Refresh
        
        '[CALL Roster Details Report SUBROUTINE]
        RosterReport
        
        '[HIDE GSR MESSAGE FORM]
        Unload frmMsg
            
        '[SET REPORT TITLE]
        Dim intOrient           As Integer
        intOrient = vbPRNone    '[SET TO NO PRINTER]
        
        Select Case prnOrient
        Case vbPRNone           '[NO PRINTER ATTACHED]
            strReport = "p_rosdet.rpt"
        Case vbPRORPortrait     '[PORTRAIT STYLE]
            strReport = "p_rosdet.rpt"
        Case vbPRORLandscape    '[LANDSCAPE STYLE]
            intOrient = prnOrient
            Printer.Orientation = vbPRORPortrait
            strReport = "p_rosdet.rpt"
        Case Else
            intOrient = prnOrient
            Printer.Orientation = vbPRORPortrait
            strReport = "p_rosdet.rpt"
        End Select
        
        mdiMain.Report.ReportFileName = strReport
          
        '[SET REPORT TITLE]
        Printer.Orientation = vbPRORPortrait
        mdiMain.Report.Formulas(0) = ""
    
    
        mdiMain.Report.WindowTitle = "Roster Detail Report, Week Starting " & Format(DsDefault("StartDate"), strDateFormat)
        mdiMain.Report.SortFields(0) = "+{RosterDetail.FullName}"
        mdiMain.Report.Action = 1
        
        '[RESTORE PRINTER ORIENTATION IF CHANGED]
        If intOrient > 0 Then Printer.Orientation = intOrient
        
    End If


ErrorHandler:
    If Err.Number > 0 Then
        '[CANNOT FIND ROSTER FILE ?]
        Msg = "Error: GSR cannot produce this report." & strBreak & strBreak & "The required report file (" & strReport & ") may be missing from the GSR directory." & strBreak & strBreak & "Error Code: " & Err.Number
        If Err.Number = 20504 Then
            Msg = Msg & strBreak & strBreak & "[Could not find required report file]"
        Else
            Msg = Msg & strBreak & strBreak & "[" & Err.Description & " in " & Err.Source & "]"
        End If
        Style = vbOKOnly                     ' Define buttons.
        Title = "Cannot Produce Report"
        Response = gsrMsg(Msg, Style, Title)
        Exit Sub
    End If

End Sub



Function LocateClass(strClassDesc) As Boolean

    '[LOCATE PASSED DESC IN CLASS DYNASET]
    Dim SQLStmt
    Dim strBookmark        As String
    
    SQLStmt = "[Description]='" & strClassDesc & "'"
    If DsClass.EOF Or DsClass.BOF Then DsClass.MoveFirst
    strBookmark = DsClass.Bookmark
    
    DsClass.FindFirst SQLStmt
    If DsClass.NoMatch Then
        LocateClass = False
        DsClass.Bookmark = strBookmark
    Else
        LocateClass = True
    End If

End Function
Function CheckStaffDay(strFullString, intDay, dateStart, dateEnd, intResult, strMessage)

    '[CHECK THAT THE STAFF MEMBER PASSED IS AVAILABLE ON THE SELECTED DAY]
    Dim SQLStmt         As String
    Dim strSurName      As String
    Dim strFirstName    As String
    Dim intDelimiter    As Integer
    Dim strBookmark     As String
    Dim strDayKey       As String
    Dim flagResult      As Boolean
    Dim strStartKey     As String
    Dim strEndKey       As String
    Dim dateStaffStart  As Date
    Dim dateStaffEnd    As Date
    Dim dateHolCheck    As Date
    Dim strClassKey     As String
    
    '[STORE CURRENT STAFF LOCATION]
    strBookmark = DsStaff.Bookmark
    
    intDelimiter = InStr(strFullString, ",")
    strSurName = Trim(Left(strFullString, intDelimiter - 1))
    strFirstName = Trim(Mid(strFullString, intDelimiter + 1))
    If intDay < 1 Or intDay > 7 Then
        flagResult = False
        CheckStaffDay = flagResult
        Exit Function
    End If
    flagResult = True
    
    '[LOCATE LASTNAME, FIRSTNAME IN DYNASET]
    If LocateStaffName(strSurName, strFirstName) = False Then
        flagResult = True
        intResult = vbUnchecked
    Else
        '[CHECK FOR RESULT OF SEARCH]
        strDayKey = "Day_" & Trim(Str$(intDay))
        strStartKey = "Start_" & Trim(Str$(intDay))
        strEndKey = "Finish_" & Trim(Str$(intDay))
        
        '[CHECK FOR STAFF IN THIS CLASS]
        strClassKey = "Class_" & Trim(Str(intRosterClass))
        If DsStaff(strClassKey) = vbUnchecked Then
            flagResult = False
            intResult = vbNotInClass
        ElseIf DsStaff(strDayKey) = vbChecked Then
            flagResult = True
            intResult = vbChecked
        ElseIf DsStaff(strDayKey) = vbUnchecked Then
            flagResult = False
            intResult = vbUnchecked
        Else
            '[NOW CHECK FOR INSIDE/OUTSIDE START/FINISH TIMES]
            '[SELECT TIMES]
            dateStaffStart = DsStaff(strStartKey)
            dateStaffEnd = DsStaff(strEndKey)
            '[ADJUST FOR NEXT DAY]
            If dateStaffStart > dateStaffEnd Then
                dateStaffEnd = dateStaffEnd + CDate("12:00") + CDate("12:00")
            End If
            strMessage = Format(DsStaff(strStartKey), "Medium Time") & " to " & Format(dateStaffEnd, "Medium Time")
        
            '[INSIDE-------------------------------------]
            If DsStaff(strDayKey) = vbInside Then
                If Not (dateStart >= dateStaffStart And dateEnd <= dateStaffEnd) Then
                    flagResult = False
                    intResult = vbInside
                End If
            End If
            '[OUTSIDE------------------------------------]
            If DsStaff(strDayKey) = vbOutside Then
                If Not ((dateStart < dateStaffStart And dateEnd <= DsStaff(strStartKey)) Or (dateStart >= dateStaffEnd And dateEnd > dateStaffEnd)) Then
                    flagResult = False
                    intResult = vbOutside
                End If
            End If
        End If
            
        '[HOLIDAYS-----------------------------------]
        If DsStaff!Holiday = vbChecked Then
            dateHolCheck = DsDefault!StartDate + (intDay - 1)
            If IsDate(DsStaff!HolStart) And IsDate(DsStaff!HolEnd) And dateHolCheck >= DsStaff!HolStart And dateHolCheck <= DsStaff!HolEnd Then
                flagResult = False
                strMessage = Format(DsStaff!HolStart, strDateFormat) & " to " & Format(DsStaff!HolEnd, strDateFormat)
                intResult = vbHoliday
            ElseIf IsDate(DsStaff!HolStart) And dateHolCheck >= DsStaff!HolStart Then
                '[REV: 3.00.36 - ADDED CHECK FOR ONLY FIRST DATE IN HOLIDAY TO ALLOW FOR UNSPECIFIED LENGTH OF ABSCENCE]
                flagResult = False
                strMessage = Format(DsStaff!HolStart, strDateFormat)
                intResult = vbHoliday
            End If
        End If

    End If

    '[RESTORE CURRENT STAFF LOCATION]
    DsStaff.Bookmark = strBookmark

    CheckStaffDay = flagResult

End Function


Sub ExceptionReport()

    '[THIS IS THE EXCEPTION REPORT SUBROUTINE. IT WILL CREATE A TEMP DYNASET    ]
    '[TO CONTAIN ALL OF THE ROSTER RECORDS AND CYCLE THROUGH USING SQL          ]
    '[STATEMENTS (HOPEFULLY).                                                   ]
   
    '[BUILD ROSTER DYNASET WITH ALL RECORDS]
    Dim SQLStmt         As String
    Dim strBookmark     As String
    Dim strFullname     As String
    Dim strDayKey       As String
    Dim intCounter      As Integer
    Dim intDayCount     As Integer
    Dim strClassDesc    As String
    Dim sinMinsWorked   As Single
    Dim sinIncrement    As Single
    Dim intInterval     As Integer
    Dim intTotal        As Integer
    Dim intClass        As Integer
    Dim strTime         As String
    Dim flagCheckStaffExist   As Boolean
    
    '[STORE CURRENT STAFF LOCATION]
    strBookmark = DsStaff.Bookmark
    flagCheckStaffExist = False
    
    '[DELETE EXISTING EXCEPTION REPORT RECORDS]
    SQLStmt = "DELETE * FROM [Exception]"
    DBMain.Execute SQLStmt, dbFailOnError
    
    '[SELECT APPROPRIATE RECORDS]
    SQLStmt = "SELECT * FROM Roster ORDER BY [Class], [ShiftStart]"
    Set DsReport = DBMain.OpenRecordset(SQLStmt, dbOpenSnapshot)
    
    '[OPEN EXCEPTION REPORT DYNASET AND CLEAR]
    Set DsException = DBMain.OpenRecordset("Exception", dbOpenTable)
    
    '[CHECK CURRENT DYNASET AND PREPARE]
    If (DsReport.EOF And DsReport.EOF) Then
        DsReport.Close
        DsException.Close
        Exit Sub
    End If
    
    '[*STAFF LOOP**************************************************************]
    DsStaff.MoveLast
    DsReport.MoveLast
    
    '[MOVE TO FIRST STAFF RECORD]
    DsStaff.MoveFirst
    intTotal = DsStaff.RecordCount * DsReport.RecordCount
    intCounter = 1
    DsReport.MoveFirst
    
    '[CYCLE THROUGH STAFF LIST]
    Do While Not DsStaff.EOF
        
        '[RESET MINUTES WORKED]
        sinMinsWorked = 0
        '[DETERMINE FULL NAME]
        strFullname = DsStaff("LastName") & ", " & DsStaff("FirstName")
        '[SHOW PROGRESS REPORT]
        Call ReportInfo(strFullname, 0)
        '[=ROSTER LOOP=========================================================]
        '[MOVE TO FIRST RECORD]
        DsReport.MoveFirst
      
        '[CYCLE THROUGH ROSTER]
        Do While Not DsReport.EOF
            '[CHECK TO SEE IF THE ROSTER IS ACTIVE]
            DsClass.AbsolutePosition = DsReport("Class") - 1
            '[PROGRESS BAR]
            Call ReportProgressBar((intCounter / intTotal) * 100)
            If DsClass("Active") = vbChecked Then
                '[-=-FIRST CHECK - STAFF NOT AVAIABLE-=-]
                Call exCheckStaffAvailable(sinMinsWorked, sinIncrement)
            End If
            If flagCheckStaffExist = False Then
                '[CHECK ALL STAFF EXITS IN DATABASE]
                '[CYCLE ACROSS DAYS]
                For intDayCount = 1 To 7
                    '[SET DAY KEY]
                    strDayKey = "Day_" & Trim(Str$(intDayCount))
                    '[CHECK ALL NAMES IN CELL AGAINST NAMES IN STAFF LIST]
                    If Len(Trim(DsReport(strDayKey))) > 0 Then Call exCheckNameInStaffList(DsReport(strDayKey), intDayCount)
                Next intDayCount
            End If
            DsReport.MoveNext
            intCounter = intCounter + 1
        Loop
        '[=====================================================================]
        '[-=-THIRD CHECKCHECK MINUTES WORKED-=-]
        If ((sinMinsWorked / 60) > DsStaff("MaxHours")) And DsStaff("maxHours") > 0 Then Call exAddNewException(constWarning, 0, 0, 0, strFullname, Str$(Format(sinMinsWorked / 60, "#0.00")) & " hours allocated/" & Str$(DsStaff("MaxHours")) & " hours allowed.")
        If ((sinMinsWorked / 60) < DsStaff("MinHours")) Then Call exAddNewException(constWarning, 0, 0, 0, strFullname, Str$(Format(sinMinsWorked / 60, "#0.00")) & " hours allocated/" & Str$(DsStaff("MinHours")) & " hours required.")
        flagCheckStaffExist = True
        DsStaff.MoveNext
    Loop
    
    '[CHECK FOR NON-FILLED ROSTER SHIFTS]
    '[REV: 3.00.28]
    If flagAllShifts = True Then
        '[SHOW INFO]
        Call ReportInfo("Please wait - checking all roster shifts are filled", 0)
        '[=ROSTER LOOP=========================================================]
        '[MOVE TO FIRST RECORD]
        DsReport.MoveFirst
        '[CYCLE THROUGH ROSTER]
        Do While Not DsReport.EOF
            If Not (IsNull(DsReport("ShiftStart")) Or IsNull(DsReport("ShiftEnd"))) Then
                '[CHECK TO SEE IF THE ROSTER IS ACTIVE]
                DsClass.AbsolutePosition = DsReport("Class") - 1
                If DsClass("Active") = vbChecked Then
                    '[CYCLE ACROSS DAYS]
                    For intDayCount = 1 To 7
                        '[SET DAY KEY]
                        strDayKey = "Day_" & Trim(Str$(intDayCount))
                        '[CHECK THERE IS SOME TEXT IN THIS CELL]
                        If Len(Trim(DsReport(strDayKey))) = 0 Then
                            '[NO TEXT - REPORT EXCEPTION]
                            intClass = DsReport("Class")
                            strTime = DsReport("ShiftStart")
                            Call exAddNewException(5, intClass, intDayCount, strTime, " ", "No staff assigned to this shift")
                        End If
                    Next intDayCount
                End If
            End If
            DsReport.MoveNext
        Loop
        '[=====================================================================]
    End If
    '[*************************************************************************]
    
    '[CLEAR REPORT INFO]
    Call ReportInfo("", 0)
    
    '[CLEAR PROGRESS BAR]
    Call ReportProgressBar(0)
    
    '[RETURN TO STAFF BOOKMARK]
    DsStaff.Bookmark = strBookmark
    
    '[CLOSE TEMPORARY DYNASET]
    DsReport.Close
    DsException.Close
    
End Sub


Sub exCheckRosterConflict(strFullname, intDayCount, sinIncrement)

    '[SUBROUTINE TO CHECK FOR CONFLICTS IN THE CURRENT ROSTER]
    '[THIS SUBROUTINE WILL TAKE A LONG TIME TO PROCESS AS IT HAS TO CHECK EVERY RECORD]
    Dim SQLStmt         As String
    Dim xSQLStmt        As String
    Dim strDayKey       As String
    Dim intCounter      As Integer
    Dim strClassDesc    As String
    Dim intClass        As Integer
    Dim intStartClass   As Integer
    Dim intRow          As Integer
    Dim strTime         As String
    Dim strPeriod       As String
    Dim dateTemp        As Date
    Dim strBookmark     As String
    Dim dateStart       As Date
    Dim dateEnd         As Date
    Dim dateShiftStart  As Date
    Dim dateShiftEnd    As Date
    
    '[EXIT IF RECORD IS NULL]
    If IsNull(DsReport("ShiftStart")) Or IsNull(DsReport("ShiftEnd")) Then Exit Sub
    '[SAVE RECORD POSITION]
    strBookmark = DsReport.Bookmark
    '[SET DAY KEY STRING]
    strDayKey = "Day_" & Trim(Str(intDayCount))
    '[SET TIME TO CHECK AGAINST]
    strTime = Format(DsReport("ShiftStart"), "Short Time") & "-" & Format(DsReport("ShiftEnd"), "Short Time")
    dateStart = CDate(DsReport("ShiftStart"))
    dateEnd = CDate(DsReport("ShiftEnd"))
    '[ALLOW FOR NEXT DAY TIMES]
    If dateEnd < dateStart Then dateEnd = dateEnd + CDate("12:00") + CDate("12:00")
    '[ALLOW FOR 24 HOUR TIMES]
    If dateEnd = dateStart Then dateEnd = dateStart + CDate("12:00") + CDate("12:00")
    '[SET CLASS ID]
    intClass = DsReport!Class
    intStartClass = DsReport!Class
    '[SET SQL STATEMENT FOR LOCATING RECORDS]
    SQLStmt = "[" & strDayKey & "] > ''"
    '[ALREADY CHECKED ALL PREVIOUS RECORDS SO NO NEED TO MOVE TO START]
    DsReport.FindNext SQLStmt
    
    Do While Not DsReport.NoMatch
        '[CHECK TO SEE IF THE ROSTER IS ACTIVE]
        DsClass.AbsolutePosition = DsReport("Class") - 1
        strClassDesc = DsClass("Description")
        '[CONTINUE IF TIMES ARE NOT NULL AND ROSTER IS ACTIVE AND DAY CELL CONTAINS STAFF NAME]
        If Not (IsNull(dateShiftStart) Or IsNull(DsReport("ShiftEnd"))) And DsClass("Active") = vbChecked And InStr(DsReport(strDayKey), strFullname) > 0 Then
            '[FIND OUT IF THE TIME MATCHES OR FALLS WITHIN THE INCREMENT ALLOWED]
            '[dateStart     = start of original roster shift                    ]
            '[dateEnd       = end of original roster shift                      ]
            '[ShiftStart    = start of checking roster shift                    ]
            '[ShiftEnd      = end of checking roster shift                      ]
            '[RULES ------------------------------------------------------------]
            '[1. ShiftStart >= dateStart and ShiftEnd <= dateEnd                ]
            '[2. ShiftStart <= dateStart and ShiftEnd > dateStart               ]
            '[3. ShiftStart < dateEnd and ShiftEnd >= dateEnd                   ]
            strPeriod = "(" & Format(DsReport("ShiftStart"), "Short Time") & "-" & Format(DsReport("ShiftEnd"), "Short Time") & ")"
            '[ASSIGN SHIFT START/END FOR CHECKING]
            dateShiftStart = DsReport!ShiftStart
            dateShiftEnd = DsReport!ShiftEnd
            '[ALLOW FOR NEXT DAY TIMES]
            If dateShiftEnd < dateShiftStart Then dateShiftEnd = dateShiftEnd + CDate("12:00") + CDate("12:00")
            '[ALLOW FOR 24 HOUR TIMES]
            If dateShiftEnd = dateShiftStart Then dateShiftEnd = dateShiftStart + CDate("12:00") + CDate("12:00")
            
            If dateShiftStart = dateStart And dateShiftEnd = dateEnd And DsReport!Class = intStartClass Then
                '[NOTHING TO DO HERE - SAME START AND FINISH ON SAME ROSTER]
            ElseIf dateShiftStart >= dateStart And dateShiftEnd <= dateEnd Then
                If DsReport!Class = intStartClass Then
                    '[SAME DAY CONFLICT]
                    Call exAddNewException(constCritical + 1, intStartClass, intDayCount, strTime, strFullname, "Conflict: [same roster] " & strPeriod & ".")
                End If
            ElseIf dateShiftStart <= dateStart And dateShiftEnd > dateStart Then
                Call exAddNewException(constCritical + 1, intStartClass, intDayCount, strTime, strFullname, "Conflict: [" & strClassDesc & "] " & strPeriod & ".")
            ElseIf dateShiftStart < dateEnd And dateShiftEnd >= dateEnd Then
                Call exAddNewException(constCritical + 1, intStartClass, intDayCount, strTime, strFullname, "Conflict: [" & strClassDesc & "] " & strPeriod & ".")
            End If
        End If
        '[FIND NEXT RECORD]
        DsReport.FindNext SQLStmt
    Loop
   
    '[RESTORE RECORD POSITION]
    DsReport.Bookmark = strBookmark
    
End Sub

Sub exCheckStaffAvailable(sinMinutesWorked, sinIncrement)

    '[SUBROUTINE TO CHECK THAT THE STAFF MEMBER LISTED IN THE ROSTER IS AVAILABLE]
    Dim strFullname     As String
    Dim strDayKey       As String
    Dim intCounter      As Integer
    Dim intDayCount     As Integer
    Dim strClassDesc    As String
    Dim intClass        As Integer
    Dim strTime         As String
    Dim dateTemp        As Date
    Dim intBookmark     As Integer
    Dim strLastSearchName       As String
    Dim strBookmark     As String
    Dim dateStart       As Date
    Dim dateEnd         As Date
    Dim intResult           As Integer
    Dim strMessage          As String
    Dim flagResult      As Boolean
    
    '[DETERMINE FULL NAME]
    strFullname = DsStaff("LastName") & ", " & DsStaff("FirstName")
    '[SAVE BOOKMARK POSITION]
    strBookmark = DsReport.Bookmark
    
    '[RESET INCREMENT - ALLOWS FOR ROSTERS WITH DIFFERING INCREMENTS]
    If IsNull(DsReport("ShiftStart")) Or IsNull(DsReport("ShiftEnd")) Then
        dateStart = Format(Now, strDateFormat)
        dateEnd = Format(Now, strDateFormat)
    Else
        dateStart = DsReport("ShiftStart")
        dateEnd = DsReport("ShiftEnd")
    End If
    
    '[ALLOW FOR NEXT DAY TIMES]
    If dateEnd < dateStart Then dateEnd = dateEnd + CDate("12:00") + CDate("12:00")
    '[ALLOW FOR 24 HOUR TIMES]
    If dateEnd = dateStart Then dateEnd = dateStart + CDate("12:00") + CDate("12:00")
    '[CALCULATE INCREMENT]
    sinIncrement = (dateEnd - dateStart) * (24 * 60)
    '[CHECK BREAKS]
    If ((DsDefault("BlockHour") * 60) + DsDefault("BlockMin")) > 0 And sinIncrement >= ((DsDefault("BlockHour") * 60) + DsDefault("BlockMin")) Then
        sinIncrement = sinIncrement - ((DsDefault("BreakHour") * 60) + DsDefault("BreakMin"))
    End If
    
    '[CYCLE ACROSS DAYS]
    For intDayCount = 1 To 7
    
        '[SET DAY KEY]
        strDayKey = "Day_" & Trim(Str$(intDayCount))
        If InStr(DsReport(strDayKey), strFullname) > 0 Then
            flagResult = CheckStaffDay(strFullname, intDayCount, dateStart, dateEnd, intResult, strMessage)
            If flagResult = False Then
                Select Case intResult
                Case 0, 1   '[NOT THIS DAY]
                    '[NAME FOUND, STAFF NOT AVAILABLE THOUGH, ADD TO EXCEPTION REPORT]
                    intClass = DsReport("Class")
                    strTime = Format(DsReport("ShiftStart"), "hh:mm") & "-" & Format(DsReport("ShiftEnd"), "hh:mm")
                    Call exAddNewException(vbNotAvail, intClass, intDayCount, strTime, strFullname, "Staff not available on this day.")
                Case 2      '[INSIDE]
                    '[NAME FOUND, STAFF NOT AVAILABLE THOUGH, ADD TO EXCEPTION REPORT]
                    intClass = DsReport("Class")
                    strTime = Format(DsReport("ShiftStart"), "hh:mm") & "-" & Format(DsReport("ShiftEnd"), "hh:mm")
                    Call exAddNewException(vbInside, intClass, intDayCount, strTime, strFullname, "Staff not available outside " & strMessage & ".")
                Case 3      '[OUTSIDE]
                    '[NAME FOUND, STAFF NOT AVAILABLE THOUGH, ADD TO EXCEPTION REPORT]
                    intClass = DsReport("Class")
                    strTime = Format(DsReport("ShiftStart"), "hh:mm") & "-" & Format(DsReport("ShiftEnd"), "hh:mm")
                    Call exAddNewException(vbOutside, intClass, intDayCount, strTime, strFullname, "Staff not available between " & strMessage & ".")
                Case 4      '[HOLIDAY]
                    '[NAME FOUND, STAFF NOT AVAILABLE THOUGH, ADD TO EXCEPTION REPORT]
                    intClass = DsReport("Class")
                    strTime = Format(DsReport("ShiftStart"), "hh:mm") & "-" & Format(DsReport("ShiftEnd"), "hh:mm")
                    Call exAddNewException(vbHoliday, intClass, intDayCount, strTime, strFullname, "Holiday: " & strMessage & ".")
                End Select
            End If
            '[MINUTES WORKED]
            sinMinutesWorked = sinMinutesWorked + sinIncrement
            Call exCheckStaffInClass(strFullname, intDayCount)
            Call exCheckRosterConflict(strFullname, intDayCount, sinIncrement)
        End If
    Next intDayCount

End Sub


Sub exAddNewException(intWarningLevel, intClass, intDayCount, strTime, strFullname, strMessage)

    '[SUBROUTINE TO PLACE NEW MESSAGE IN REPORT GRID]
    Dim strClassDescription     As String
    Dim strDayName              As String
    Dim intStartDay             As Integer
    Dim intCounter              As Integer
    Dim SQLStmt                 As String
    
    intStartDay = DsDefault("StartDay")
        
    '[SET DAY NAME - LONG FORMAT]
    If intDayCount = 0 Then
        strDayName = " "
    Else
        If intStartDay + (intDayCount - 1) > 7 Then
            strDayName = ArrayWeek(intDayCount - 7 + intStartDay - 1).LongDay
        Else
            strDayName = ArrayWeek(intStartDay + (intDayCount - 1)).LongDay
        End If
    End If
    
    '[SET ROSTER NAME]
    If intClass > 0 Then
        SQLStmt = "ID = " & intClass
        DsClass.FindFirst SQLStmt
        strClassDescription = DsClass("Description")
    Else
        strClassDescription = " "
    End If
    
    '[TIME FORMAT]
    If strTime = "0" Then
        strTime = " "
    Else
        strTime = Format(strTime, "hh:mm")
        '[OLD TIME FORMAT] -> "hh:mm AMPM"
    End If
    
    '[TRIM MESSAGE]
    strMessage = Trim(strMessage)
    
    '[ADD EXCEPTION TO DATABASE RECORDSET OBJECT]
    DsException.AddNew
        DsException("Roster") = strClassDescription
        DsException("Day") = strDayName
        DsException("Time") = strTime
        DsException("FullName") = strFullname
        DsException("Message") = strMessage
    DsException.Update

End Sub

Sub exCheckStaffInClass(strFullname, intDayCount)
    
    Dim strClassKey     As String
    Dim intClass        As Integer
    Dim strTime         As String
    
    '[CHECK NULL START/END TIME]
    If IsNull(DsReport("ShiftStart")) Or IsNull(DsReport("ShiftEnd")) Then Exit Sub
    
    '[-=-SECOND CHECK - STAFF NOT IN THIS CLASS-=-]
    strClassKey = "Class_" & Trim(Str$(DsReport("Class")))
    intClass = DsReport("Class")
    strTime = DsReport("ShiftStart")
    If DsStaff(strClassKey) = vbUnchecked Then
        '[STAFF MEMBER NOT AVAILABLE IN THIS CLASS]
        Call exAddNewException(constSerious, intClass, intDayCount, strTime, strFullname, "Staff member not in this class")
    End If

End Sub

Function gsrMsg(strMsg, Style, strTitle)

    '[PROCESS PASSED VARIABLES AND SETUP MSG FORM]
    Load frmMsg
    
    '[SET TITLE AND MESSAGE]
    frmMsg.LabelTitle = strTitle
    frmMsg.LabelMessage = Trim(strMsg & " ")
    
    '[SET STYLE OF BUTTONS]
    Select Case Style
    Case vbYesNo
        frmMsg.cmdYes.Visible = True
        frmMsg.cmdNo.Visible = True
        frmMsg.cmdCancel.Visible = False
        frmMsg.cmdOK.Visible = False
        
        frmMsg.cmdYes.Default = True
        frmMsg.cmdNo.Cancel = True
        
        frmMsg.cmdYes.Left = (frmMsg.Width / 3) - (frmMsg.cmdYes.Width / 2)
        frmMsg.cmdNo.Left = ((frmMsg.Width / 3) * 2) - (frmMsg.cmdNo.Width / 2)
        
    Case vbYesNoCancel
        frmMsg.cmdYes.Visible = True
        frmMsg.cmdNo.Visible = True
        frmMsg.cmdCancel.Visible = True
        frmMsg.cmdOK.Visible = False
    
        frmMsg.cmdYes.Default = True
'        frmMsg.cmdCancel.Cancel = True
        
        frmMsg.cmdYes.Left = (frmMsg.Width / 4) - (frmMsg.cmdYes.Width / 2)
        frmMsg.cmdNo.Left = ((frmMsg.Width / 4) * 2) - (frmMsg.cmdNo.Width / 2)
        frmMsg.cmdCancel.Left = ((frmMsg.Width / 4) * 3) - (frmMsg.cmdCancel.Width / 2)
    
    Case vbOKOnly
        frmMsg.cmdYes.Visible = False
        frmMsg.cmdNo.Visible = False
        frmMsg.cmdCancel.Visible = False
        frmMsg.cmdOK.Visible = True
        
        frmMsg.cmdOK.Default = True
        frmMsg.cmdOK.Cancel = True
       
        frmMsg.cmdOK.Left = (frmMsg.Width / 2) - (frmMsg.cmdOK.Width / 2)
    
    Case vbInformation
        frmMsg.cmdYes.Visible = False
        frmMsg.cmdNo.Visible = False
        frmMsg.cmdCancel.Visible = False
        frmMsg.cmdOK.Visible = False
    
    Case vbOKCancel
        frmMsg.cmdYes.Visible = False
        frmMsg.cmdNo.Visible = False
        frmMsg.cmdOK.Visible = True
        frmMsg.cmdCancel.Visible = True
        
        frmMsg.cmdCancel.Cancel = False
        frmMsg.cmdOK.Default = True
        
        frmMsg.cmdCancel.Left = ((frmMsg.Width / 3) * 2) - (frmMsg.cmdCancel.Width / 2)
        frmMsg.cmdOK.Left = (frmMsg.Width / 3) - (frmMsg.cmdOK.Width / 2)
        
        
    Case vbQuestion
        frmMsg.cmdYes.Visible = False
        frmMsg.cmdNo.Visible = False
        frmMsg.cmdOK.Visible = True
        frmMsg.cmdCancel.Visible = True
        frmMsg.TextNote.FontName = frmRoster.GridRoster.FontName
        frmMsg.TextNote.FontBold = frmRoster.GridRoster.FontBold
        frmMsg.TextNote.FontItalic = frmRoster.GridRoster.FontItalic
        frmMsg.TextNote.Visible = True
        frmMsg.TextNote.Text = Trim(strMsg & " ")
    
        frmMsg.cmdOK.Left = (frmMsg.Width / 3) - (frmMsg.cmdOK.Width / 2)
        frmMsg.cmdCancel.Left = ((frmMsg.Width / 3) * 2) - (frmMsg.cmdCancel.Width / 2)
    
        frmMsg.cmdOK.Default = True
    
    Case Else
        frmMsg.cmdYes.Visible = True
        frmMsg.cmdNo.Visible = True
        frmMsg.cmdCancel.Visible = False
        frmMsg.cmdOK.Visible = False
        
        frmMsg.cmdYes.Default = True
        
        frmMsg.cmdYes.Left = (frmMsg.Width / 3) - (frmMsg.cmdYes.Width / 2)
        frmMsg.cmdNo.Left = ((frmMsg.Width / 3) * 2) - (frmMsg.cmdNo.Width / 2)
    
    End Select
    
    '[BEEP]
    If flagSounds = True Then Beep
    '[SHOW FORM MODAL]
    If Style = vbInformation Then
        frmMsg.Show
    Else
        If frmMsg.Visible Then  '[FORM ALREADY LOADED - ERROR MESSAGE ?]
            frmMsg.Show
            frmMsg.labelInfo.Visible = False
            frmMsg.gaugeBar.Visible = False
            frmMsg.gaugeBorder.Visible = False
        Else
            frmMsg.Show 1
        End If
    End If
    '[RETURN VALUE IN gsrReturn]
    gsrMsg = gsrReturn

End Function

Sub ReportInfo(strDisplayText, sinForeColor)

    '[SET CAPTION ON INFO LABEL TO PASSED TEXT AND COLOR TO PASSED COLOR]
    If sinForeColor <= 0 Then sinForeColor = vbBlack
    frmMsg.labelInfo.ForeColor = sinForeColor
    frmMsg.labelInfo.Caption = strDisplayText
    frmMsg.labelInfo.Refresh

End Sub

Sub BuildRosterDynaset(intClass)

    '[REBUILD ROSTER DYNASET BASED UPON THE CLASS PASSED]
    Dim SQLStmt         As String
    
    '[SELECT APPROPRIATE RECORDS]
    SQLStmt = "SELECT * FROM Roster WHERE Class = " & Str(intClass) & " ORDER BY CLASS, SHIFTSTART, SHIFTEND"
    Set DsRoster = DBMain.OpenRecordset(SQLStmt, dbOpenDynaset)
    
End Sub

Sub FillRosterGrid()

    '[PLACE LIST ITEMS IN ROSTER DYNASET]
    Dim intCol          As Integer
    Dim intRow          As Integer
    Dim intCounter      As Integer
        
    '[GRID OFF]
    GridRedraw False
           
    '[CLEAR CURRENT GRID AND PREPARE]
    frmRoster.GridRoster.Rows = 2
    
    '[NEED TO CLEAR LAST ROW OF GRID]
    frmRoster.GridRoster.Row = 1
    For intCol = 0 To (frmRoster.GridRoster.Cols - 1)
        frmRoster.GridRoster.Col = intCol
        frmRoster.GridRoster.Text = ""
    Next intCol
    
    intCounter = 1
    
    If DsRoster.RecordCount > 0 Then
        '[CYLCE THROUGH DYNASET AND PLACE IN GRID]
        DsRoster.MoveFirst
        Do While Not DsRoster.EOF
            frmRoster.GridRoster.AddItem Format(DsRoster("ShiftStart"), "Medium Time") & Chr$(vbKeyTab) & Format(DsRoster("ShiftEnd"), "Medium Time") & Chr$(vbKeyTab) & DsRoster("Day_1") & Chr$(vbKeyTab) & DsRoster("Day_2") & Chr$(vbKeyTab) & DsRoster("Day_3") & Chr$(vbKeyTab) & DsRoster("Day_4") & Chr$(vbKeyTab) & DsRoster("Day_5") & Chr$(vbKeyTab) & DsRoster("Day_6") & Chr$(vbKeyTab) & DsRoster("Day_7"), intCounter
            '[OLD TIME FORMAT] -> "hh:mm AMPM"
            frmRoster.GridRoster.Row = intCounter
            intCounter = intCounter + 1
            DsRoster.MoveNext
        Loop
    End If
    
    '[REMOVE LAST ROW]
    If frmRoster.GridRoster.Rows > 2 Then frmRoster.GridRoster.Rows = frmRoster.GridRoster.Rows - 1

    '[RESIZE ALL COLUMN WIDTHS TO MATCH GRID WIDTH]
    Call FitRosterToGrid
    
    '[GRID ON]
    GridRedraw True
    
End Sub


Sub ReportProgressBar(intPercent)

    '[CATCH EXTREME VALUES]
    Dim intCurrPerc     As Integer
    intCurrPerc = Int((frmMsg.gaugeBar.Width / Val(frmMsg.gaugeBar.Tag)) * 100)
    
    '[VER: 3.00.30]
    If intPercent < 0 Then intPercent = 0
    If intPercent > 100 Then intPercent = 100
    
    '[EXIT IF NOT BIG ENOUGH INCREMENT TO SHOW]
    If (intCurrPerc < intPercent - 2) Or (intCurrPerc > intPercent + 2) Then
        frmMsg.gaugeBar.Width = Val(frmMsg.gaugeBar.Tag) * (intPercent / 100)
        '[REV: 3.00.33]
        '[ALLOW OTHER EVENTS TO OCCUR]
        DoEvents
    End If
    
End Sub

Sub PutNameInCell(strFullString)

    Dim sinColWidth         As Single
    Dim sinRowHeight        As Single
    Dim intEnterCount       As Single
    Dim intStartPos         As Single
    
    intStartPos = 1
    intEnterCount = 0
    
    '[CHECK FOR PLACEMENTS INTO TITLE ROWS/COLS]
    If frmRoster.GridRoster.Col <= 1 Or frmRoster.GridRoster.Row = 0 Then Exit Sub
    
    '[ROUTINE TO PLACE THE PASSED NAME INTO THE ROSTER CELL]
    '[CHECKING FOR OTHER NAMES AND PREVIOUS INSTANCE OF NAME]
    If Len(Trim(frmRoster.GridRoster.Text)) = 0 Then
        '[CELL IS EMPTY - PLACE NAME]
        frmRoster.GridRoster.Text = Trim(strFullString)
    Else
        '[CELL IS NOT EMPTY - CHECK FOR NAME]
        If InStr(frmRoster.GridRoster.Text, strFullString) > 0 Then
            '[NAME IS ALREADY IN CELL, EXIT ROUTINE]
            Exit Sub
        Else
            '[NAME NOT IN CELL, ADD AT END OF CELL]
            frmRoster.GridRoster.Text = Trim(frmRoster.GridRoster.Text) & strBreak & Trim(strFullString)
        End If
    End If

    ResizeRosterCell (strFullString)

End Sub
Sub RemoveFromRoster(strFullString)
    
    '[REMOVE THE PASSED STAFF NAME FROM ALL SELECTED CELLS IN THE ROSTER GRID]
    Dim intCounter          As Integer
    Dim strLastName         As String
    Dim strFirstName        As String
    Dim intDelimiter        As Integer
    Dim intCol              As Integer
    Dim intRow              As Integer
    
    '[Delimit list item, break into Surname, FirstName]
    intDelimiter = InStr(strFullString, ",")
    strLastName = Trim(Left(strFullString, intDelimiter - 1))
    strFirstName = Trim(Mid(strFullString, intDelimiter + 1))

    '[DETERMINE WHETHER IT IS A MULTISELECT OR SINGLE CELL FILL]
    'If frmRoster.GridRoster.SelStartCol = -1 Or frmRoster.GridRoster.SelEndCol = -1 Then
    If frmRoster.GridRoster.Col = -1 Or frmRoster.GridRoster.ColSel = -1 Then
        '[SINGLE CELL REMOVE]
        '[CALL ROUTINE TO REMOVE NAME FROM CELL]
        RemoveNameFromCell (strFullString)
    Else
        '[MULTI CELL REMOVE]
        'For intCol = frmRoster.GridRoster.SelStartCol To frmRoster.GridRoster.SelEndCol
        For intCol = frmRoster.GridRoster.Col To frmRoster.GridRoster.ColSel
            frmRoster.GridRoster.Col = intCol
            'For intRow = frmRoster.GridRoster.SelStartRow To frmRoster.GridRoster.SelEndRow
            For intRow = frmRoster.GridRoster.Row To frmRoster.GridRoster.RowSel
                frmRoster.GridRoster.Row = intRow
                '[CALL ROUTINE TO REMOVE NAME FROM CELL]
                RemoveNameFromCell (strFullString)
            Next intRow
        Next intCol
    End If

End Sub

Sub RemoveNameFromCell(strFullString)
    
    Dim sinColWidth         As Single
    Dim sinRowHeight        As Single
    Dim intEnterCount       As Single
    Dim intStartPos         As Single
    Dim strLeft             As String
    Dim strRight            As String
    Dim intLeftPos          As Integer
    Dim intRightPos         As Integer
    Dim intLength           As Integer
    
    intStartPos = 1
    intEnterCount = 0
    
    '[ROUTINE TO REMOVE THE PASSED NAME FROM THE ROSTER CELL]
    '[CHECKING FOR OTHER NAMES AND PREVIOUS INSTANCE OF NAME]
    If frmRoster.GridRoster.Text = "" Then
        '[CELL IS EMPTY - EXIT ROUTINE]
        Exit Sub
    Else
        '[CELL IS NOT EMPTY - CHECK FOR NAME]
        If InStr(frmRoster.GridRoster.Text, strFullString) > 0 Then
            '[NAME IS IN CELL, REMOVE FROM CELL]
            '[FIND STRING LEFT OF PASSED NAME]
            intLeftPos = InStr(frmRoster.GridRoster.Text, strFullString)
            If intLeftPos = 1 Then
                strLeft = ""
            Else
                strLeft = Left(frmRoster.GridRoster.Text, intLeftPos - 2)
            End If
            '[FIND STRING RIGHT OF PASSED NAME]
            strRight = Trim(Mid$(frmRoster.GridRoster.Text, Len(strLeft) + Len(strFullString) + 2))
            '[APPLY NEW CONCATENATED STRING TO CELL]
            frmRoster.GridRoster.Text = strLeft & strRight
            '[TURN SAVE BUTTON ON]
            frmRoster.cmdSave.Visible = True
        Else
            '[NAME NOT IN CELL, EXIT ROUTINE]
            Exit Sub
        End If
    End If

    '[RESIZE ROW AND COL TO SUIT]
    sinRowHeight = frmRoster.GridRoster.RowHeight(frmRoster.GridRoster.Row)
    sinColWidth = frmRoster.GridRoster.ColWidth(frmRoster.GridRoster.Col)
    
    If sinColWidth < (frmRoster.TextWidth(strFullString) * 1.25) Then frmRoster.GridRoster.ColWidth(frmRoster.GridRoster.Col) = (frmRoster.TextWidth(strFullString) * 1.25)
            
    '[LOOK FOR ENTER CHARACTERS IN THE CURRENT CELL AND ADJUST SIZE]
    Do While InStr(intStartPos, frmRoster.GridRoster.Text, strBreak) > 0
        intEnterCount = intEnterCount + 1
        intStartPos = InStr(intStartPos, frmRoster.GridRoster.Text, strBreak) + 1
    Loop

    If sinRowHeight < (frmRoster.TextHeight("A") * (intEnterCount + 1) * 1.25) Then frmRoster.GridRoster.RowHeight(frmRoster.GridRoster.Row) = (frmRoster.TextHeight("A") * (intEnterCount + 1) * 1.25)

End Sub

Sub ResizeRosterCell(strFullString)
    
    '[DECLARATIONS]
    Dim sinRowHeight        As Single
    Dim sinColWidth         As Single
    Dim intEnterCount       As Integer
    Dim intStartPos         As Integer
    Dim frmTemp             As Form
    Dim sinMod              As Single
    
    '[SET ACTIVE FORM]
    If flagScratch = True Then
        Set frmTemp = frmScratch
    Else
        Set frmTemp = frmRoster
    End If
    
    '[SET START POSITION]
    sinMod = 1.1
    intStartPos = 1
    intEnterCount = 0
    
    '[RETURN IF CELL IS EMPTY]
    If Len(frmTemp.GridRoster.Text) <= 0 Then Exit Sub
    
    '[RESIZE ROW AND COL TO SUIT]
    sinColWidth = frmTemp.GridRoster.ColWidth(frmTemp.GridRoster.Col)
    sinRowHeight = frmTemp.GridRoster.RowHeight(frmTemp.GridRoster.Row)
        
    If sinColWidth < (frmTemp.TextWidth(strFullString) * sinMod) Then frmTemp.GridRoster.ColWidth(frmTemp.GridRoster.Col) = (frmTemp.TextWidth(strFullString) * sinMod)
            
    '[LOOK FOR ENTER CHARACTERS IN THE CURRENT CELL AND ADJUST SIZE]
    intEnterCount = CountReturns(frmTemp.GridRoster.Text)
    
    '[RESIZE ROSTER GRID ROW HEIGHT]
    If sinRowHeight < (frmTemp.TextHeight("A") * (intEnterCount + 1) * sinMod) Then frmTemp.GridRoster.RowHeight(frmTemp.GridRoster.Row) = (frmTemp.TextHeight("A") * (intEnterCount + 1) * sinMod)

End Sub



Sub SaveRosterGrid()

    '[PLACE LIST ITEMS IN ROSTER DYNASET]
    Dim intCol          As Integer
    Dim intRow          As Integer
    Dim intCounter      As Integer
    Dim strDayKey       As String
    Dim intCurrRow      As Integer
    Dim intCurrCol      As Integer
    Dim intCurrStartRow      As Integer
    Dim intCurrStartCol      As Integer
    Dim intCurrEndRow      As Integer
    Dim intCurrEndCol      As Integer
    
    '[CLEAR CURRENT DYNASET AND PREPARE]
    Do While DsRoster.RecordCount > 0
        DsRoster.MoveFirst
        DsRoster.Delete
    Loop
    intCounter = 1
    
    '[SAVE CURRENT COLUMN AND ROW POSITION]
    intCurrRow = frmRoster.GridRoster.Row
    intCurrCol = frmRoster.GridRoster.Col
'    intCurrStartRow = frmRoster.GridRoster.SelStartRow
'    intCurrStartCol = frmRoster.GridRoster.SelStartCol
'    intCurrEndRow = frmRoster.GridRoster.SelEndRow
'    intCurrEndCol = frmRoster.GridRoster.SelEndCol
    intCurrStartRow = frmRoster.GridRoster.Row
    intCurrStartCol = frmRoster.GridRoster.Col
    intCurrEndRow = frmRoster.GridRoster.RowSel
    intCurrEndCol = frmRoster.GridRoster.ColSel
    If intCurrStartRow = 0 Then intCurrStartRow = 1
    If intCurrEndRow = 0 Then intCurrEndRow = 1
    
    For intRow = 1 To (frmRoster.GridRoster.Rows - 1)
        frmRoster.GridRoster.Row = intRow
        frmRoster.GridRoster.Col = 0
        '[ADD NEW ITEM TO DYNASET]
        DsRoster.AddNew
            DsRoster("Class") = intRosterClass
            If IsDate(frmRoster.GridRoster.Text) Then DsRoster("Time") = frmRoster.GridRoster.Text
            If IsDate(frmRoster.GridRoster.Text) Then DsRoster("ShiftStart") = frmRoster.GridRoster.Text
            frmRoster.GridRoster.Col = 1
            DsRoster("ShiftEnd") = frmRoster.GridRoster.Text
            DsRoster("Row") = intRow
        
        For intCol = 2 To (frmRoster.GridRoster.Cols - 1)
            frmRoster.GridRoster.Col = intCol
            '[CREATE DAY KEY]
            strDayKey = "Day_" & Trim(Str(intCol - 1))
            '[TRIM AND REMOVE LEADING CR'S FROM TEXT]
            frmRoster.GridRoster.Text = Trim(frmRoster.GridRoster.Text)
            If Left$(frmRoster.GridRoster.Text, 1) = strBreak Then frmRoster.GridRoster.Text = Mid(frmRoster.GridRoster.Text, 2)
            DsRoster(strDayKey) = (frmRoster.GridRoster.Text & " ")
        Next intCol
        
        '[UPDATE DYNASET ITEM]
        DsRoster.Update
        
    Next intRow

    '[RESTORE PREVIOUS COLUMN AND ROW POSITION]
    frmRoster.GridRoster.Row = intCurrRow
    frmRoster.GridRoster.Col = intCurrCol
    frmRoster.GridRoster.Row = intCurrStartRow
    frmRoster.GridRoster.Col = intCurrStartCol
    frmRoster.GridRoster.RowSel = intCurrEndRow
    frmRoster.GridRoster.ColSel = intCurrEndCol


End Sub

Sub RosterReport()

    '[THIS IS THE ROSTER DETAILS REPORT SUBROUTINE. IT WILL CREATE A TEMP DYNASET]
    '[TO CONTAIN ALL OF THE ROSTER RECORDS AND CYCLE THROUGH USING SQL           ]
    '[STATEMENTS (HOPEFULLY).                                                    ]
    
    '[REV: 3.00.28]
    '[- MAJOR CHANGES - (1) determine format and structure for adding weekly pay rate to report]
    '[                  (2) determine code and placing of code for weekly pay rate adjustments]
    
    '[BUILD ROSTER DYNASET WITH ALL RECORDS]
    Dim SQLStmt         As String
    Dim strBookmark     As String
    Dim strFullname     As String
    Dim strDayKey       As String
    Dim intCounter      As Integer
    Dim intDayCount     As Integer
    Dim strClassDesc    As String
    Dim sinMinsWorked   As Single
    Dim sinIncrement    As Single
    Dim intInterval     As Integer
    Dim intRosterCount  As Integer
    Dim dateStart       As Date
    Dim dateEnd         As Date
    Dim flagStart       As Boolean
    Dim strStaffID      As String
    Dim strName         As String
    Dim strRoster       As String
    Dim sinMinutes      As Single
    Dim sinTotal        As Single
    Dim sinTotalAmount  As Single
    Dim sinTotalMinutes As Single
    Dim strClass        As String
    Dim intSRC          As Integer
    Dim sinSMT          As Single
    
    '[DELETE EXISTING ROSTER DETAIL REPORT RECORDS]
    SQLStmt = "DELETE * FROM [RosterDetail]"
    DBMain.Execute SQLStmt, dbFailOnError
    
    '[OPEN ROSTER DETAIL REPORT DYNASET]
    Set DsRosterDetail = DBMain.OpenRecordset("RosterDetail", dbOpenTable)
    
    '[SET UP ARRAY FOR HOLDING STAFF WAGE DATA]
    Dim arrayStaff(10) As StaffType
    Dim arrayRoster(10) As StaffType
    
    '[STORE CURRENT STAFF LOCATION]
    strBookmark = DsStaff.Bookmark
    
    '[SELECT APPROPRIATE RECORDS]
    SQLStmt = "SELECT * FROM Roster ORDER BY CLASS, SHIFTSTART"
    Set DsReport = DBMain.OpenRecordset(SQLStmt, dbOpenDynaset)
            
    '[CHECK CURRENT DYNASET AND PREPARE]
    If (DsReport.EOF And DsReport.BOF) Then
        DsReport.Close
        Exit Sub
    End If
    
    '[*STAFF LOOP**************************************************************]
    '[MOVE TO FIRST STAFF RECORD]
    DsStaff.MoveFirst
    '[CYCLE THROUGH STAFF LIST]
    Do While Not DsStaff.EOF
        
        '[RESET MINUTES WORKED]
        Erase arrayStaff
        '[DETERMINE FULL NAME]
        strFullname = DsStaff("LastName") & ", " & DsStaff("FirstName")
        '[SHOW PROGRESS REPORT]
        Call ReportInfo(strFullname, 0)
        '[PROGRESS BAR]
        Call ReportProgressBar(((DsStaff.AbsolutePosition + 1) / DsStaff.RecordCount) * 100)
                        
        '[=ROSTER LOOP=========================================================]
        '[MOVE TO FIRST RECORD]
        DsReport.MoveFirst
        '[CYCLE THROUGH ROSTER]
        Do While Not DsReport.EOF
            
            '[CHECK TO SEE IF THE ROSTER IS ACTIVE]
            DsClass.AbsolutePosition = (DsReport("Class") - 1)
            
            If DsClass("Active") = vbChecked And Not IsNull(DsReport("ShiftStart")) And Not IsNull(DsReport("ShiftEnd")) Then
                For intDayCount = 1 To 7
                    '[-=-NOW CHECK EACH DAY TO SEE IF STAFF MEMBER IS INCLUDED IN ANY DAY-=-]
                    strDayKey = "Day_" & Trim(Str(intDayCount))
                    If InStr(DsReport(strDayKey), strFullname) > 0 Then
                        '[RESET INCREMENT - ALLOWS FOR ROSTERS WITH DIFFERING INCREMENTS]
                        dateStart = DsReport("ShiftStart")
                        dateEnd = DsReport("ShiftEnd")
                        '[ALLOW FOR NEXT DAY TIMES]
                        If dateEnd < dateStart Then dateEnd = dateEnd + CDate("12:00") + CDate("12:00")
                        '[ALLOW FOR 24 HOUR TIMES]
                        If dateEnd = dateStart Then dateEnd = dateStart + CDate("12:00") + CDate("12:00")
                        '[CALCULATE INCREMENT]
                        sinIncrement = (dateEnd - dateStart) * (24 * 60)
                        '[CHECK TO SEE IF INCREMENT IS > WORK BLOCK, IF SO SUBTRACT BREAK TIME]
                        '[ALSO CHECK FOR NO WORK PERIOD SET (00:00)]
                        If ((DsDefault("BlockHour") * 60) + DsDefault("BlockMin")) > 0 And sinIncrement >= ((DsDefault("BlockHour") * 60) + DsDefault("BlockMin")) Then
                            sinIncrement = sinIncrement - ((DsDefault("BreakHour") * 60) + DsDefault("BreakMin"))
                        End If
                        '[ADD MINUTES TO STAFF ARRAY]
                        arrayStaff(DsReport("Class")).Minutes = arrayStaff(DsReport("Class")).Minutes + sinIncrement
                    End If
                Next intDayCount
                
                '[PLACE RATE IN RATE FIELD]
                strClass = "Rate_" & Trim(Str(DsReport("Class")))
                If IsNull(DsStaff(strClass)) Then
                    arrayStaff(DsReport("Class")).Rate = 0
                Else
                    arrayStaff(DsReport("Class")).Rate = DsStaff(strClass)
                End If
                
            End If
            '[MOVE TO NEXT ROSTER RECORD]
            DsReport.MoveNext
        Loop
        '[=====================================================================]
        flagStart = False
        sinMinsWorked = 0
        sinTotalMinutes = 0
        sinTotalAmount = 0
        
        '[COUNT NUMBER OF ROSTERS ASSIGNED TO]
        intSRC = 0
        sinSMT = 0
        For intCounter = 1 To 10
            If arrayStaff(intCounter).Minutes > 0 Then
                intSRC = intSRC + 1
                sinSMT = sinSMT + arrayStaff(intCounter).Minutes
            End If
        Next intCounter
        
        For intCounter = 1 To 10
            '[CHECK TO SEE IF THERE ARE HOURS ALLOCATED FOR THIS RECORD]
            If arrayStaff(intCounter).Minutes > 0 Then
                '[ASSIGN STAFF ID AND NAMES]
                strStaffID = DsStaff("StaffID")
                strName = strFullname
                '[ASSIGN OTHER VALUES - ROSTER NAME, HOURS AND TOTAL]
                sinMinutes = arrayStaff(intCounter).Minutes / 60
                If arrayStaff(intCounter).Rate = 0 Then sinTotal = 0 Else sinTotal = arrayStaff(intCounter).Rate * sinMinutes
                '[SET ROSTER NAME]
                DsClass.AbsolutePosition = (intCounter - 1)
                strRoster = DsClass("Description")
                '[ADD MINUTES TO TOTAL]
                sinTotalMinutes = sinTotalMinutes + sinMinutes
                sinTotalAmount = sinTotalAmount + sinTotal
                
                '[ADD MINUTES TO TOTAL ROSTER ARRAY]
                arrayRoster(intCounter).Minutes = arrayRoster(intCounter).Minutes + sinMinutes

                '[CHECK HERE FOR HOURLY PAY RATE TYPE]
                '[REV: 3.00.28]
                If DsStaff!PayType = vbWeekly Then
                    arrayStaff(intCounter).Rate = (DsStaff!HourRate / sinSMT) * 60
                    sinTotal = sinMinutes * arrayStaff(intCounter).Rate
                    arrayRoster(intCounter).Amount = arrayRoster(intCounter).Amount + sinTotal
                Else
                    '[ADD CURRENCY AMOUNT TO TOTAL ROSTER ARRAY]
                    arrayRoster(intCounter).Amount = arrayRoster(intCounter).Amount + sinTotal
                End If
                
                '[ADD ROSTER DETAILS TO DATABASE RECORDSET OBJECT]
                DsRosterDetail.AddNew
                    DsRosterDetail("StaffID") = strStaffID
                    DsRosterDetail("FullName") = strName
                    DsRosterDetail("Roster") = strRoster
                    DsRosterDetail("Rate") = arrayStaff(intCounter).Rate
                    DsRosterDetail("Hours") = sinMinutes
                    DsRosterDetail("Amount") = sinTotal
                DsRosterDetail.Update
            End If
        Next intCounter
        
        
        '[=====================================================================]
        '[MOVE TO NEXT STAFF RECORD]
        DsStaff.MoveNext
    Loop
    '[*************************************************************************]
    '[PLACE STAFF ROSTER TOTALS]
    sinTotalMinutes = 0
    sinTotalAmount = 0

    For intCounter = 1 To 10
        '[CHECK TO SEE IF THERE ARE HOURS ALLOCATED FOR THIS RECORD]
        If arrayRoster(intCounter).Minutes > 0 Then
            strStaffID = ""
            strName = ""
            '[ASSIGN OTHER VALUES - ROSTER NAME, HOURS AND TOTAL]
            sinMinutes = arrayRoster(intCounter).Minutes
            sinTotal = arrayRoster(intCounter).Amount
            '[SET ROSTER NAME]
            DsClass.AbsolutePosition = (intCounter - 1)
            strRoster = DsClass("Description")
            '[ADD MINUTES TO TOTAL]
            sinTotalMinutes = sinTotalMinutes + sinMinutes
            sinTotalAmount = sinTotalAmount + sinTotal
            '[ADD LINE TO REPORT GRID]
            DsRosterDetail.AddNew
                DsRosterDetail("StaffID") = " "
                DsRosterDetail("FullName") = "[ Totals ]"
                DsRosterDetail("Roster") = strRoster
                DsRosterDetail("Rate") = sinTotal / sinMinutes
                DsRosterDetail("Hours") = sinMinutes
                DsRosterDetail("Amount") = sinTotal
            DsRosterDetail.Update
        End If
    Next intCounter
    '[*************************************************************************]

    '[CLEAR REPORT INFO]
    Call ReportInfo("", 0)
    '[CLEAR PROGRESS BAR]
    Call ReportProgressBar(0)
    
    '[RETURN TO STAFF BOOKMARK]
    DsStaff.Bookmark = strBookmark
    
    '[CLOSE TEMPORARY DYNASET]
    DsReport.Close

End Sub

Sub StatusBar(strMessage)

    '[PLACE PASSED MESSAGE ON THE STATUS BAR]
    If mdiMain.panelStatusBar.Visible = False Then Exit Sub
    If mdiMain.panelStatusBar.Caption <> " " + strMessage Then mdiMain.panelStatusBar.Caption = " " + strMessage

End Sub

Sub Terminate(Response)

    '[CHECK FOR PRESENCE OF SAVE BUTTONS]
    Dim Msg As String
    Dim Style
    Dim Title
    Dim intFlag         As Integer
    Dim intDateFormat   As Integer
    
    '[CHECK FOR CHANGED DATA AND NOTIFY]
    If frmRoster.cmdSave.Visible = True Then
        '[ROSTER IS UNSAVED, POPUP YES/NO DIALOG]
        Msg = "You have made changes to this roster (" & frmRoster.ComboClass.Text & ") but have not saved these changes." & strBreak & strBreak & "If you choose not to save now, any changes you have made since your last save will be lost." & strBreak & strBreak & "Do you wish to save these changes before you exit ?"
        Style = vbYesNoCancel ' Define buttons.
        Title = "Roster Changes Not Saved"  ' Define title.
        Response = gsrMsg(Msg, Style, Title)
        If Response = vbYes Then    ' User chose Yes.
            SaveRosterGrid
        ElseIf Response = vbCancel Then
            Exit Sub
        End If
    End If
    
    '[CHECK TO SEE IF WE ARE MOVING FROM AN UNSAVED RECORD]
    If frmSet.cmdSave.Visible = True Then
        '[RECORD IS UNSAVED, POPUP YES/NO DIALOG]
        Msg = "You have made changes to this staff record (" & DsStaff("Lastname") & ", " & DsStaff("FirstName") & ") but have not saved these changes." & strBreak & strBreak & "If you choose not to save now, any changes you have made since your last save will be lost." & strBreak & strBreak & "Do you wish to save these changes before you exit ?"
        Style = vbYesNoCancel ' Define buttons.
        Title = "Staff Changes Not Saved"  ' Define title.
        Response = gsrMsg(Msg, Style, Title)
        If Response = vbYes Then    ' User chose Yes.
            '[COMMIT RECORD CHANGES TO THE DYNASET]
            SaveStaffDetails
        ElseIf Response = vbCancel Then
            Exit Sub
        End If
    End If

    '[place closing/termination statements here]
    Dim intStaffState       As Integer
    Dim intRosterState      As Integer
    Dim intControlState     As Integer
    Dim intMainState        As Integer
    
    '[SET DEFAULT DATE FORMAT]
    intDateFormat = 0
    
    '[DATE FORMAT]
    If frmSet.OptionDateFormat(0).Value = True Then intDateFormat = 0
    If frmSet.OptionDateFormat(1).Value = True Then intDateFormat = 1
    If frmSet.OptionDateFormat(2).Value = True Then intDateFormat = 2
        
    DsDefault.Edit
    
        '[SAVE WINDOW STATES HERE]
        DsDefault("RosterState") = frmRoster.WindowState
        DsDefault("MainState") = mdiMain.WindowState
        DsDefault("ControlTop") = frmSet.Top
        DsDefault("ControlLeft") = frmSet.Left
        If mdiMain.WindowState = 0 Then
            If mdiMain.Width > 100 Then DsDefault!StaffState = mdiMain.Width
            If mdiMain.Height > 100 Then DsDefault!ControlState = mdiMain.Height
            If mdiMain.Left > 100 Then DsDefault!StaffLeft = mdiMain.Left
            If mdiMain.Top > 100 Then DsDefault!StaffTop = mdiMain.Top
        End If
       
        '[FLAGS FOR DELETE AND SOUNDS]
        intFlag = 0
        '[DELETE]
        If flagDeleteConfirm = True Then intFlag = intFlag + (1 ^ 2)
        '[SOUNDS]
        If flagSounds = True Then intFlag = intFlag + (2 ^ 2)
        DsDefault("DeleteConfirm") = intFlag
        '[ALL SHIFTS FILLED]
        '[REV: 3.00.28]
        If flagAllShifts = True Then DsDefault!AllShifts = 1 Else DsDefault!AllShifts = 0
        
        '[SAVE ROSTER GRID FONTS]
        DsDefault("RosterFontName") = frmRoster.GridRoster.Font.Name
        DsDefault("RosterFontBold") = frmRoster.GridRoster.Font.Bold
        DsDefault("RosterFontItalic") = frmRoster.GridRoster.Font.Italic
        DsDefault("RosterFontSize") = frmRoster.GridRoster.Font.Size
        
        '[SAVE LAST ROSTER ID]
        DsDefault("RosterID") = frmRoster.ComboClass.ListIndex
        
        '[SET TOOLBAR STATE]
        If mdiMain.PanelToolBar.Visible = True Then
            DsDefault("ToolBarState") = 1
        Else
            DsDefault("ToolBarState") = 0
        End If
        
        '[SET STATUSBAR STATE]
        If mdiMain.panelStatusBar.Visible = True Then
            DsDefault("StatusBarState") = 1
        Else
            DsDefault("StatusBarState") = 0
        End If
        
        '[SAVE LOCKED COLUMNS]
        DsDefault("RosterLocked") = frmRoster.GridRoster.FixedCols
        
        '[DATE FORMAT AND CUSTOM DATE FIELDS]
        DsDefault("DateFormat") = intDateFormat
        DsDefault("CustomDate") = frmSet.maskDateMask.Text
        
    DsDefault.Update

    '[VER: 3.00.32]
    '[CLOSE ALL OPEN DATABASE OBJECTS]
    DsRoster.Close
    DsStaff.Close
    DsDefault.Close
    DBMain.Close
    GSRWorkspace.Close
    Close
    
    '[VER: 3.00.33]
    '[PERFORM BACKUP/COMPACTION IF APPROPRIATE MENU ITEM WAS SELECTED]
    If flagBackup = True Then Call CompactDatabase
    Unload mdiMain
    End

End Sub

Sub TransferToRoster(strFullString, flagConstraints, flagResult, flagCancel)

    '[TRANSFER THE PASSED STAFF NAME TO ALL SELECTED CELLS IN THE ROSTER GRID]
    Dim intCounter          As Integer
    Dim strLastName         As String
    Dim strFirstName        As String
    Dim intDelimiter        As Integer
    Dim intCol              As Integer
    Dim intRow              As Integer
    Dim Msg As String
    Dim Style
    Dim Response
    Dim Title
    Dim intTimeCol          As Integer
    Dim intResult           As Integer
    Dim dateStart           As Date
    Dim dateEnd             As Date
    Dim strMessage          As String
    
    '[Delimit list item, break into Surname, FirstName]
    intDelimiter = InStr(strFullString, ",")
    strLastName = Trim(Left(strFullString, intDelimiter - 1))
    strFirstName = Trim(Mid(strFullString, intDelimiter + 1))
    flagResult = False
    
    '[DETERMINE WHETHER IT IS A MULTISELECT OR SINGLE CELL FILL]
    'If (frmRoster.GridRoster.SelStartCol = -1) Or (frmRoster.GridRoster.SelEndCol = -1) Or (frmRoster.GridRoster.SelStartCol = frmRoster.GridRoster.SelEndCol And frmRoster.GridRoster.SelStartRow = frmRoster.GridRoster.SelEndRow) Then
    If (frmRoster.GridRoster.Col = -1) Or (frmRoster.GridRoster.ColSel = -1) Or (frmRoster.GridRoster.Col = frmRoster.GridRoster.ColSel And frmRoster.GridRoster.Row = frmRoster.GridRoster.RowSel) Then
        '[SINGLE CELL FILL]
        '[CALL ROUTINE TO PLACE NAME IN CELL]
        If frmRoster.GridRoster.Col <= 1 Then Exit Sub
        
        '[LOCATE START AND FINISH TIME]
        intTimeCol = frmRoster.GridRoster.Col
        frmRoster.GridRoster.Col = 0
        dateStart = frmRoster.GridRoster.Text
        frmRoster.GridRoster.Col = 1
        dateEnd = frmRoster.GridRoster.Text
        '[ADD 24 HOURS IF NEXT DAY]
        If dateEnd < dateStart Then dateEnd = dateEnd + CDate("12:00") + CDate("12:00")
        frmRoster.GridRoster.Col = intTimeCol
            
        If CheckStaffDay(strFullString, frmRoster.GridRoster.Col - 1, dateStart, dateEnd, intResult, strMessage) Or flagConstraints = False Then
            PutNameInCell (strFullString)
            '[TURN SAVE BUTTON ON]
            frmRoster.cmdSave.Visible = True
            flagResult = True
        Else
            '[SHOW MESSAGE]
            Select Case intResult
            Case 0, 1   '[NOT THIS DAY]
                Msg = ArrayWeek(DayOfWeek(frmRoster.GridRoster.Col - 1)).LongDay & ", " & Format(DsDefault!StartDate + frmRoster.GridRoster.Col - 2, strDateFormat) & " : This staff member (" & strFullString & ") is marked as not being available on " & ArrayWeek(DayOfWeek(frmRoster.GridRoster.Col - 1)).LongDay & "s." & strBreak & strBreak & "If you require this staff member to be available on this day, you should click the appropriate check box on the roster details form."
                Style = vbOKCancel                          ' Define buttons.
                Title = "Staff Member Not Available"        ' Define title.
            Case vbInside       '[INSIDE]
                Msg = ArrayWeek(DayOfWeek(frmRoster.GridRoster.Col - 1)).LongDay & ", " & Format(DsDefault!StartDate + frmRoster.GridRoster.Col - 2, strDateFormat) & " : This staff member (" & strFullString & ") is marked as not being available outside the hours of " & strMessage & " on " & ArrayWeek(DayOfWeek(frmRoster.GridRoster.Col - 1)).LongDay & "s." & strBreak & strBreak & "If you require this staff member to be available outside these hours, you should change the hours available on their staff form."
                Style = vbOKCancel                          ' Define buttons.
                Title = "Staff Member Not Available"        ' Define title.
            Case vbOutside      '[OUTSIDE]
                Msg = ArrayWeek(DayOfWeek(frmRoster.GridRoster.Col - 1)).LongDay & ", " & Format(DsDefault!StartDate + frmRoster.GridRoster.Col - 2, strDateFormat) & " : This staff member (" & strFullString & ") is marked as not being available between the hours of " & strMessage & " on " & ArrayWeek(DayOfWeek(frmRoster.GridRoster.Col - 1)).LongDay & "s." & strBreak & strBreak & "If you require this staff member to be available within these hours, you should change the hours available on their staff form."
                Style = vbOKCancel                          ' Define buttons.
                Title = "Staff Member Not Available"        ' Define title.
            Case vbHoliday      '[HOLIDAYS]
                Msg = ArrayWeek(DayOfWeek(frmRoster.GridRoster.Col - 1)).LongDay & ", " & Format(DsDefault!StartDate + frmRoster.GridRoster.Col - 2, strDateFormat) & " : This staff member (" & strFullString & ") is marked as being on holidays from :" & strBreak & strBreak & strMessage & "." & strBreak & strBreak & "You will not be able to assign this staff member to any rosters until the specified holiday period has passed. If the staff member is not on holidays/leave of absence, you should change the holiday details on the staff form."
                Style = vbOKCancel                          ' Define buttons.
                Title = "Staff Member Not Available"        ' Define title.
            Case vbNotInClass   '[NOT IN CLASS]
                Msg = ArrayWeek(DayOfWeek(frmRoster.GridRoster.Col - 1)).LongDay & ", " & Format(DsDefault!StartDate + frmRoster.GridRoster.Col - 2, strDateFormat) & " : This staff member (" & strFullString & ") is marked as not being available to this roster class (" & frmRoster.ComboClass.Text & ")." & strBreak & strBreak & "If you require this staff member to be available for this roster, you should change the staff roster details on the staff form."
                Style = vbOKCancel                          ' Define buttons.
                Title = "Staff Member Not Available"        ' Define title.
            End Select
            Response = gsrMsg(Msg, Style, Title)
            flagResult = False
            flagContinue = Response
            If Response = vbCancel Then Exit Sub
        End If
    Else
        '[MULTI CELL FILL]
        'For intCol = frmRoster.GridRoster.SelStartCol To frmRoster.GridRoster.SelEndCol
        For intCol = frmRoster.GridRoster.Col To frmRoster.GridRoster.ColSel
            If intCol <= 1 Then Exit For
            
            '[LOCATE START AND FINISH TIME]
            If frmRoster.GridRoster.Row = 0 Then Exit For '[EXIT IF PROCESSING FIRST ROW OF GRID]
            intTimeCol = frmRoster.GridRoster.Col
            frmRoster.GridRoster.Col = 0
            dateStart = frmRoster.GridRoster.Text
            frmRoster.GridRoster.Col = 1
            dateEnd = frmRoster.GridRoster.Text
            '[ADD 24 HOURS IF NEXT DAY]
            If dateEnd < dateStart Then dateEnd = dateEnd + CDate("12:00") + CDate("12:00")
            frmRoster.GridRoster.Col = intTimeCol
            
            frmRoster.GridRoster.Col = intCol
            'For intRow = frmRoster.GridRoster.SelStartRow To frmRoster.GridRoster.SelEndRow
            For intRow = frmRoster.GridRoster.Row To frmRoster.GridRoster.RowSel
                frmRoster.GridRoster.Row = intRow
                '[CHECK STAFF MEMBER IS AVAILABLE FOR THIS DAY]
                If CheckStaffDay(strFullString, frmRoster.GridRoster.Col - 1, dateStart, dateEnd, intResult, strMessage) Or flagConstraints = False Then
                    '[CALL ROUTINE TO PLACE NAME IN CELL]
                    PutNameInCell (strFullString)
                    '[TURN SAVE BUTTON ON]
                    frmRoster.cmdSave.Visible = True
                    flagResult = True
                Else
                    '[SHOW MESSAGE]
                    Select Case intResult
                    Case 0, 1   '[NOT THIS DAY]
                        Msg = ArrayWeek(DayOfWeek(frmRoster.GridRoster.Col - 1)).LongDay & ", " & Format(DsDefault!StartDate + frmRoster.GridRoster.Col - 1, strDateFormat) & " : This staff member (" & strFullString & ") is marked as not being available on " & ArrayWeek(DayOfWeek(frmRoster.GridRoster.Col - 1)).LongDay & "s." & strBreak & strBreak & "If you require this staff member to be available on this day, you should click the appropriate check box on the roster details form."
                        Style = vbOKCancel                          ' Define buttons.
                        Title = "Staff Member Not Available"        ' Define title.
                    Case vbInside       '[INSIDE]
                        Msg = ArrayWeek(DayOfWeek(frmRoster.GridRoster.Col - 1)).LongDay & ", " & Format(DsDefault!StartDate + frmRoster.GridRoster.Col - 1, strDateFormat) & " : This staff member (" & strFullString & ") is marked as not being available outside the hours of " & strMessage & " on " & ArrayWeek(DayOfWeek(frmRoster.GridRoster.Col - 1)).LongDay & "s." & strBreak & strBreak & "If you require this staff member to be available outside these hours, you should change the hours available on their staff form."
                        Style = vbOKCancel                          ' Define buttons.
                        Title = "Staff Member Not Available"        ' Define title.
                    Case vbOutside      '[OUTSIDE]
                        Msg = ArrayWeek(DayOfWeek(frmRoster.GridRoster.Col - 1)).LongDay & ", " & Format(DsDefault!StartDate + frmRoster.GridRoster.Col - 1, strDateFormat) & " : This staff member (" & strFullString & ") is marked as not being available between the hours of " & strMessage & " on " & ArrayWeek(DayOfWeek(frmRoster.GridRoster.Col - 1)).LongDay & "s." & strBreak & strBreak & "If you require this staff member to be available within these hours, you should change the hours available on their staff form."
                        Style = vbOKCancel                          ' Define buttons.
                        Title = "Staff Member Not Available"        ' Define title.
                    Case vbHoliday      '[HOLIDAYS]
                        Msg = ArrayWeek(DayOfWeek(frmRoster.GridRoster.Col - 1)).LongDay & ", " & Format(DsDefault!StartDate + frmRoster.GridRoster.Col - 1, strDateFormat) & " : This staff member (" & strFullString & ") is marked as being on holidays from :" & strBreak & strBreak & strMessage & "." & strBreak & strBreak & "You will not be able to assign this staff member to any rosters until the specified holiday period has passed. If the staff member is not on holidays/leave of absence, you should change the holiday details on the staff form."
                        Style = vbOKCancel                          ' Define buttons.
                        Title = "Staff Member Not Available"        ' Define title.
                    End Select
                    Response = gsrMsg(Msg, Style, Title)
                    flagResult = False
                    flagContinue = Response
                    If Response = vbCancel Then Exit Sub
                    '[EXIT THE ROUTINE]
                    If intResult < vbInside Then Exit For
                End If
            Next intRow
        Next intCol
    End If

End Sub
Sub AddNewStaff()

    Dim strDisplayname      As String
    Dim intCounter          As Integer
    Dim SQLStmt             As String
    
    '[REV: 3.00.30]
    '[ALLOCATE STAFF ID]
    Dim intNewID            As Single
    intNewID = 0
    SQLStmt = "[StaffID]='*" & Trim(Str$(intNewID)) & "'"
    DsStaff.FindFirst SQLStmt
    Do While Not DsStaff.NoMatch
        intNewID = intNewID + 1
        SQLStmt = "[StaffID]='*" & Trim(Str$(intNewID)) & "'"
        DsStaff.FindFirst SQLStmt
    Loop
    
    '[ADD A NEW STAFF MEMBER TO THE LIST]
    DsStaff.AddNew
        DsStaff("StaffID") = "*" & Trim(Str$(intNewID))
        DsStaff("LastName") = "*LastName"
        DsStaff("FirstName") = "*FirstName"
        DsStaff!PayType = 0
        DsStaff("Day_1") = 1
        DsStaff("Day_2") = 1
        DsStaff("Day_3") = 1
        DsStaff("Day_4") = 1
        DsStaff("Day_5") = 1
        DsStaff("Day_6") = 1
        DsStaff("Day_7") = 1
    DsStaff.Update
    
    '[MOVE TO FIRST RECORD IF NO RECORD LOCATED]
    DsStaff.Bookmark = DsStaff.LastModified
    
    If DsStaff.EOF Or DsStaff.BOF Then DsStaff.MoveFirst
    
    strDisplayname = DsStaff("LastName") & ", " & DsStaff("FirstName")
    
    '[REFILL STAFF LIST FOR ORDER]
    FillStaffList
                    
    '[RELOCATE STAFF NAME]
    For intCounter = 0 To (frmSet.ListStaff.ListCount - 1)
        If frmSet.ListStaff.List(intCounter) = strDisplayname Then frmSet.ListStaff.ListIndex = intCounter
    Next intCounter

End Sub


Sub FillStaffList()
    
    '[EXIT IF FORM NOT LOADED]
    If flagLoad = False Then Exit Sub
    
    '[PLACE VALUES FROM THE STAFF DYNASET INTO THE STAFF FORM LIST]
    Dim intCounter As Integer
    '[IF NO STAFF MEMBERS, ADD A NEW DEFAULT STAFF MEMBER]
    If DsStaff.RecordCount = 0 Then AddNewStaff
    DsStaff.MoveFirst
    
    '[CLEAR ALL CURRENT CONTENTS IN THE LIST]
    frmSet.ListStaff.Clear
    
    '[MOVE THROUGH STAFF DYNASET AND FILL LIST]
    Do While Not DsStaff.EOF
        frmSet.ListStaff.AddItem DsStaff("LastName") & ", " & DsStaff("FirstName")
        DsStaff.MoveNext
    Loop
    
End Sub

Sub FillStaffRosterList()
    
    '[PLACE VALUES FROM THE STAFF DYNASET INTO THE STAFF FORM LIST]
    Dim intCounter          As Integer
    Dim intClassChoice      As Integer
    Dim strClassKey         As String
    Dim DsBookmark
    
    '[SAVE CURRENT LOCATION IN DYNASET]
    If DsStaff.EOF Or DsStaff.BOF Then
        '[MOVE TO FIRST POSITION OFF BEGINNING/END OF TABLE]
        DsStaff.MoveFirst
    End If
    DsBookmark = DsStaff.Bookmark
        
    If frmRoster.ComboClass.ListIndex = -1 Then Exit Sub
    
    '[MAKE CLASS CHOICE]
    intClassChoice = frmRoster.ComboClass.ItemData(frmRoster.ComboClass.ListIndex)
    strClassKey = "Class_" & Trim(Str(intClassChoice))
    
    If DsStaff.RecordCount = 0 Then Exit Sub
    DsStaff.MoveFirst
    
    '[CLEAR ALL CURRENT CONTENTS IN THE LIST]
    frmRoster.ListStaff.Clear
    
    '[MOVE THROUGH STAFF DYNASET AND FILL LIST]
    Do While Not DsStaff.EOF
        '[ADD STAFF IF CLASS MATCHES AND STAFF MEMBER IS AVAILABLE]
        If (DsStaff(strClassKey) = vbChecked) Then frmRoster.ListStaff.AddItem DsStaff("LastName") & ", " & DsStaff("FirstName")
        DsStaff.MoveNext
    Loop

    '[RESTORE STAFF POSITION]
    DsStaff.Bookmark = DsBookmark
    
End Sub
Sub LocateStaff()

    '[FIND STAFF MEMBER IN DYNASET AND UPDATE ALL STAFF TEXT BOXES]
    Dim strFullString   As String   '[Temporary full string lastname, firstname]
    Dim strSurName      As String   '[Temporary surname string]
    Dim strFirstName    As String   '[Temporary firstname string]
    Dim intDelimiter    As Integer  '[Temporary location of surname, firstname delimiter]
    Dim SQLStmt         As String   '[Search string]
    Dim intCounter      As Integer
    Dim strField        As String   '[String representation of Class]
    
    '[IF LIST IS EMPTY THEN EXIT]
    If frmSet.ListStaff.ListIndex < 0 Then Exit Sub
    
    '[CHECK TO SEE IF WE ARE MOVING FROM AN UNSAVED RECORD]
    If frmSet.cmdSave.Visible = True Then
        '[RECORD IS UNSAVED, POPUP YES/NO DIALOG]
        Dim Msg As String
        Dim Style
        Dim Response
        Dim Title
        
        Msg = "You have made changes to this staff record (" & DsStaff("Lastname") & ", " & DsStaff("FirstName") & ") but have not saved these changes." & strBreak & strBreak & "If you choose not to save now, any changes you have made since your last save will be lost." & strBreak & strBreak & "Do you wish to save these changes before you move to another record ?"
        Style = vbYesNo ' Define buttons.
        Title = "Staff Changes Not Saved"  ' Define title.
        Response = gsrMsg(Msg, Style, Title)
        If Response = vbYes Then    ' User chose Yes.
        
            Dim strDisplayname      As String
        
            '[COMMIT RECORD CHANGES TO THE DYNASET]
            SaveStaffDetails
            
            '[SAVE DISPLAYED NAME TO TEMPORARY STRING]
            strDisplayname = frmSet.ListStaff.List(frmSet.ListStaff.ListIndex)
            
            '[FILL STAFF LIST SO WE GET ORDER]
            FillStaffList
                
            '[RELOCATE STAFF NAME]
            For intCounter = 0 To (frmSet.ListStaff.ListCount - 1)
                If frmSet.ListStaff.List(intCounter) = strDisplayname Then frmSet.ListStaff.ListIndex = intCounter
            Next intCounter
        
        End If
        
    End If
    
    '[Delimit list item, break into Surname, FirstName]
    strFullString = frmSet.ListStaff.List(frmSet.ListStaff.ListIndex)
    If strFullString = "" Then Exit Sub
    
    intDelimiter = InStr(strFullString, ",")
    strSurName = Trim(Left(strFullString, intDelimiter - 1))
    strFirstName = Trim(Mid(strFullString, intDelimiter + 1))
    
    '[LOCATE LASTNAME, FIRSTNAME IN DYNASET]
'    SQLStmt = "LastName = '" & strSurname & "' AND FirstName = '" & strFirstName & "'"
'    DsStaff.FindFirst SQLStmt

    If LocateStaffName(strSurName, strFirstName) Then
        '[UPDATE STAFF DETAILS ON FORM]
        ShowStaffDetails
    End If
    
End Sub

Function LocateStaffName(strSurName, strFirstName) As Boolean

    '[REV: 3.00.34]
    '[ALTERNATE STAFF NAME LOCATE ROUTINE]
    Dim strCriteria         As String
    strCriteria = "[LastName] = '" & SQLFixup(strSurName) & "' AND [FirstName] = '" & SQLFixup(strFirstName) & "'"
    DsStaff.FindFirst strCriteria
    
    '[STAFF MEMBER FOUND ?]
    If DsStaff.NoMatch Then LocateStaffName = False Else LocateStaffName = True
    
End Function

Sub ProcessDayLength()

    '[PROCESS CHANGES TO COMBO BOXES ON THE CONTROL FORM]
    Select Case frmSet.OptionTime(0).Value
    Case True
        DsDefault.Edit
            DsDefault("StartTime") = frmSet.ComboHour.Text & ":" & frmSet.ComboMinute.Text
        DsDefault.Update
    Case False
        DsDefault.Edit
            DsDefault("EndTime") = frmSet.ComboHour.Text & ":" & frmSet.ComboMinute.Text
        DsDefault.Update
    End Select
        
    '[CALCULATE DAY LENGTH AND DISPLAY]
    If DsDefault("StartTime") > DsDefault("EndTime") Then
        frmSet.MaskDayLength.Text = CDate("12:00") + CDate("12:00") + (DsDefault("EndTime") - DsDefault("StartTime"))
    Else
        frmSet.MaskDayLength.Text = DsDefault("StartTime") - DsDefault("EndTime")
    End If
    
End Sub

Sub FillClassGrid()
    
    '[GENERAL PUBLIC VARIABLE VALUE ASSIGNMENT]
    flagConstraints = True
    flagContinue = True
    
    '[PLACE VALUES FROM THE CLASS DYNASET INTO THE MAIN FORM CLASS GRID]
    Dim intCounter As Integer
    DsClass.MoveFirst
    
    frmSet.GridClass.Row = 0
    frmSet.GridClass.Col = 1
    frmSet.GridClass.Text = "Code"
    frmSet.GridClass.Col = 2
    frmSet.GridClass.Text = "Description"
    frmSet.GridClass.Col = 3
    frmSet.GridClass.Text = "Notes"
    frmSet.GridClass.Col = 0
    frmSet.GridClass.Text = ""
    
    '[Make sure only 10 editable rows are included]
    frmSet.GridClass.Rows = 11
    
    For intCounter = 1 To 10
        
        frmSet.GridClass.Row = intCounter
        frmSet.GridClass.Col = 1
        frmSet.GridClass.Text = Trim(DsClass("Code") & " ")
        frmSet.GridClass.Col = 2
        frmSet.GridClass.Text = Trim(DsClass("Description") & " ")
        frmSet.GridClass.Col = 3
        frmSet.GridClass.Text = Trim(DsClass("Note") & " ")
        frmSet.GridClass.Col = 0
        Select Case DsClass("Active")
        Case vbChecked
            frmSet.GridClass.Picture = frmSet.ImageSwitch(constWarning).Picture
        Case vbUnchecked
            frmSet.GridClass.Picture = frmSet.ImageSwitch(constCritical).Picture
        Case Else
            frmSet.GridClass.Picture = frmSet.ImageSwitch(constSerious).Picture
        End Select
        frmSet.GridClass.Text = DsClass("Active")
        
        DsClass.MoveNext
        
    Next intCounter
    
    
End Sub

Sub Initialise()
    
    '[ERROR HANDLER ROUTINE]
    On Error GoTo ErrorHandler
    
    Dim strDataFile             As String
    Dim StaffIndex
    Dim StaffField
    
    '[MESSAGE VARIABLES]
    Dim Msg
    Dim Style
    Dim Title
    Dim Response
    
    '[SET PUBLIC VARIABLE VALUES]
    flagLoaded = False
    flagContinue = True
    flagConstraints = True
    flagTerminate = False
    
    '[SET BREAK STRING VARIABLE]
    strBreak = Chr$(vbKeyReturn)
    
    '[SET LAST MOUSE CO-ORDS FOR STAFF LIST]
    sinLastMouseX = -1
    sinLastMouseY = -1
    
    '[Assign dynasets to the public variables]
    Dim SQLStmt             As String
    Dim strRegCode          As String
    
    '[OPEN GSR DATABASE]
    '[DEBUG]
    Call LogToFile("setting GSR database")
    Set GSRWorkspace = DBEngine.Workspaces(0)
    
    '[DEBUG]
    strDataFile = Dir("gsr.dat")
    Call LogToFile("opening GSR database - " & strDataFile)
    Set DBMain = GSRWorkspace.OpenDatabase(strDataFile, True, False)
    
    '[SQLSTMT TO ORDER STAFF DYNASET]
    SQLStmt = "SELECT * FROM Staff ORDER BY LastName, FirstName;"
        
    '[DEBUG]
    Call LogToFile("opening recordsets")
    
    '[OPEN RECORDSETS]
    '[DEBUG]
    Call LogToFile("- staff")
    
    '[OPEN STAFF TABLE AS DYNASET]
    Set DsStaff = DBMain.OpenRecordset(SQLStmt, dbOpenDynaset)
    
    '[DEBUG]
    Call LogToFile("- class")
    Set DsClass = DBMain.OpenRecordset("Class", dbOpenDynaset)
    '[DEBUG]
    Call LogToFile("- roster")
    Set DsRoster = DBMain.OpenRecordset("Roster", dbOpenDynaset)
    '[DEBUG]
    Call LogToFile("- defaults")
    Set DsDefault = DBMain.OpenRecordset("Defaults", dbOpenTable)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    Call LogToFile("checking database structure")
    '[---------------------------------------------------------------------------------]
    
    '[CHECK HERE FOR STRUCTURE CHANGES - LATER VERSIONS OF GSR]
    Call CheckDatabaseStructure  '[REACTIVATED 11may98]
    
    '[Change to first record in default dynaset]
    DsDefault.MoveFirst
    
    '[DEBUG]
    Call LogToFile("counting records")
    '[MOVE TO LAST RECORD IN DYNASETS]
    If Not (DsStaff.EOF And DsStaff.BOF) Then
        DsStaff.MoveLast
        DsStaff.MoveFirst
    End If
    '[DEBUG]
    Call LogToFile("= " & DsStaff.RecordCount & " staff")
    If Not (DsClass.EOF And DsClass.BOF) Then
        DsClass.MoveLast
        DsClass.MoveFirst
    End If
    '[DEBUG]
    Call LogToFile("= " & DsClass.RecordCount & " classes")
    If Not (DsRoster.EOF And DsRoster.BOF) Then
        DsRoster.MoveLast
        DsRoster.MoveFirst
    End If
    '[DEBUG]
    Call LogToFile("= " & DsRoster.RecordCount & " roster lines")
    
    '[DEBUG]
    Call LogToFile("setting day of week array")
    '[Fill Day of Week Array]
    ArrayWeek(1).LongDay = "Sunday":        ArrayWeek(1).ShortDay = "Sun"
    ArrayWeek(2).LongDay = "Monday":        ArrayWeek(2).ShortDay = "Mon"
    ArrayWeek(3).LongDay = "Tuesday":       ArrayWeek(3).ShortDay = "Tue"
    ArrayWeek(4).LongDay = "Wednesday":     ArrayWeek(4).ShortDay = "Wed"
    ArrayWeek(5).LongDay = "Thursday":      ArrayWeek(5).ShortDay = "Thu"
    ArrayWeek(6).LongDay = "Friday":        ArrayWeek(6).ShortDay = "Fri"
    ArrayWeek(7).LongDay = "Saturday":      ArrayWeek(7).ShortDay = "Sat"

    '[LOAD FORMS]
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    Call LogToFile("LOADING FORM: GSR CONTROL PANEL")
    '[---------------------------------------------------------------------------------]
    flagLoad = False
    Load frmSet
    flagLoad = True
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    Call LogToFile("LOADING FORM: ROSTER")
    StatusBar "Loading roster form."
    '[---------------------------------------------------------------------------------]
    Load frmRoster
            
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    Call LogToFile("applying font - " & DsDefault("RosterFontName"))
    '[---------------------------------------------------------------------------------]
    '[APPLY FONT TO ROSTER GRID]
    If DsDefault("RosterFontName") > "" Then
        frmRoster.GridRoster.Font.Name = DsDefault("RosterFontName")
        frmRoster.GridRoster.Font.Bold = DsDefault("RosterFontBold")
        frmRoster.GridRoster.Font.Italic = DsDefault("RosterFontItalic")
        frmRoster.GridRoster.Font.Size = DsDefault("RosterFontSize")
        '[APPLY FONT TO CONTROL GRID]
        Set frmSet.GridClass.Font = frmRoster.GridRoster.Font
        '[APPLY FONT TO WHOLE ROSTER FORM]
        Set frmRoster.Font = frmRoster.GridRoster.Font
    End If
    
    '[SET FIRST ROW HEIGHT]
    frmRoster.GridRoster.RowHeight(0) = frmRoster.TextHeight("A")

    '[CHECK FIRST INSTALLED DATE AND IF ISNULL THEN PLACED CURRENT DATE IN THE FIELD]
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    '[---------------------------------------------------------------------------------]
    If IsNull(DsDefault("InstallDate")) Or Not (IsDate(DsDefault("InstallDate"))) Then
        DsDefault.Edit
            DsDefault("InstallDate") = Format(Now, strDateFormat)
        DsDefault.Update
    End If
    
    '[SET DAYS USED PUBLIC VARIABLE]
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    Call LogToFile("calculating days used")
    '[---------------------------------------------------------------------------------]
    sinDaysUsed = CDate(Format(Now, "dd/mm/yyyy")) - CDate(Format(DsDefault("InstallDate"), "dd/mm/yyyy"))
    '[DEBUG]
    Call LogToFile("= " & sinDaysUsed)
    
    '[SET MODIFIER AND CALL VALIDATION ROUTINE]
    sinModifier = 129#
    '[------------------------------------------------------------------------------------------]
    '[-VALIDATION ROUTINES HERE-----------------------------------------------------------------]
    '[------------------------------------------------------------------------------------------]
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    Call LogToFile("validating registration")
    '[---------------------------------------------------------------------------------]
    
    strRegCode = Validate(DsDefault("RegUser"))
    If strRegCode = "" Or Not (strRegCode = DsDefault("RegCode")) Or IsNull(DsDefault("RegCode")) Then
        '[CODE DOESN'T MATCH - REPLACE CODE IN DATABASE]
        DsDefault.Edit
            If sinDaysUsed > 45 Then
                DsDefault("RegUser") = "Unregistered Version"
                DsDefault("RegCode") = ""
            Else
                DsDefault("RegUser") = "Shareware Evaluation Version"
                DsDefault("RegCode") = ""
            End If
        DsDefault.Update
    End If
    '[DEBUG]
    Call LogToFile("= " & DsDefault!RegUser & "-" & DsDefault!RegCode)
    
    '[DEBUG]
    Call LogToFile("filling roster list")
    '[FILL ROSTER COMBO LIST]
    FillRosterList
    '[SET SO NO ROSTER IS SELECTED]
    frmRoster.ComboClass.ListIndex = -1
    '[SET PROGRAM LOADED FLAG]
    flagLoaded = True
    
    '[SET ROSTER TO THAT LAST WORKED ON]
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    Call LogToFile("restoring last roster")
    
    '[---------------------------------------------------------------------------------]
    If (frmRoster.ComboClass.ListCount - 1) < DsDefault("RosterID") Then
        frmRoster.ComboClass.ListIndex = (frmRoster.ComboClass.ListCount - 1)
    Else
        frmRoster.ComboClass.ListIndex = DsDefault("RosterID")
    End If
    '[DEBUG]
    Call LogToFile("= " & frmRoster.ComboClass.Text)
    
    '[SET LOCKED COLUMNS]
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    Call LogToFile("setting locked columns")
    '[---------------------------------------------------------------------------------]
    frmRoster.GridRoster.FixedCols = DsDefault("RosterLocked")
    
    
    '[FIRST ACTIVE COLUMN]
    '[FIRST ACTIVE ROW]
    frmRoster.GridRoster.Col = 2
    frmRoster.GridRoster.ColSel = 2 ' SelStartCol = 2
    'frmRoster.GridRoster.SelEndCol = 2
    frmRoster.GridRoster.Row = 1
    frmRoster.GridRoster.RowSel = 1 '.SelStartRow = 1
    'frmRoster.GridRoster.SelEndRow = 1
    
    '[RESTORE WINDOW STATE FROM DEFAULT DYNASET]
    '[0 = HIDDEN]
    '[1 = NORMAL]
    '[2 = MAXIMISED]
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    '[DEBUG]
    Call LogToFile("restoring windows sizes")
    '[---------------------------------------------------------------------------------]
    mdiMain.WindowState = DsDefault("MainState")
    If mdiMain.WindowState = 0 And DsDefault!StaffState > 100 And DsDefault!ControlState > 100 And DsDefault!StaffLeft > 100 And DsDefault!StaffTop > 100 Then
        mdiMain.Width = DsDefault!StaffState
        mdiMain.Height = DsDefault!ControlState
        mdiMain.Left = DsDefault!StaffTop
        mdiMain.Top = DsDefault!StaffTop
    End If

    '[CHECK PRINTER CAPABILITY]
    If Printers.Count > 0 Then
        '[DEBUG]
        Call LogToFile("Default printer : " & Printer.DeviceName & " on " & Printer.Port)
        Call LogToFile("         Driver : " & Printer.DriverName)
    End If

    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Generic Staff Roster ready."
    '[---------------------------------------------------------------------------------]
    Call LogToFile("completed initialisation")
   
ErrorHandler:
    If Err.Number > 0 Then
        '[ERRORS IN THE INITIALISATION ROUTINE ?]
        '[PROCESS BY ERROR CODE]
        Msg = "Error: GSR has experienced a critical error while initialising." & strBreak & strBreak & "Please contact the author with details of this error.  GSR will now terminate." & strBreak & strBreak & "Error Code: " & Err.Number
        If Err.Number = 3356 Then
            Msg = Msg & ".  It appears as though the GSR database file (" & strDataFile & ") is marked 'read-only'."
        End If
        Msg = Msg & strBreak & strBreak & "[" & Err.Description & " in " & Err.Source & "]"
        Style = vbOKOnly                     ' Define buttons.
        Title = "Error in GSR Initialisation"
        Response = gsrMsg(Msg, Style, Title)
        
        If Workspaces(0).Databases.Count > 0 Then
            '[CLOSE GSR WORKSPACE]
            DBMain.Close
            GSRWorkspace.Close
            '[ATTEMPT TO REPAIR GSR DATABASE]
            If Err.Number <> 3356 And strDataFile > "" Then
                DBEngine.RepairDatabase strDataFile
            End If
        End If
        Close
        End         '[END PROGRAM]
    End If
      
End Sub

Sub FitRosterToGrid()
    
    '[RESIZE ALL COLUMN WIDTHS TO MATCH GRID WIDTH]
    Dim sinWidth        As Single
    Dim sinHeight       As Single
    Dim intColCounter   As Integer
    Dim intRowCounter   As Integer
    Dim intCols         As Integer
    Dim intRows         As Integer
    Dim intRow          As Integer
    Dim intCol          As Integer
    Dim intLines        As Integer
    Dim frmTemp         As Form
    Dim sinMod          As Single
    
    '[SET ACTIVE FORM]
    If flagScratch = True Then
        Set frmTemp = frmScratch
    Else
        Set frmTemp = frmRoster
    End If
    
    '[SAVE STARTING GRID POSITION]
    intCol = frmTemp.GridRoster.Col
    intRow = frmTemp.GridRoster.Row
    sinMod = 2
    sinWidth = frmTemp.GridRoster.Width - 480
    intCols = (frmTemp.GridRoster.Cols - 1)
    intRows = (frmTemp.GridRoster.Rows - 1)
    
    For intColCounter = 0 To intCols
        '[NOW RESIZE COLUMN HEIGHTS]
        sinHeight = 0
        '[COLUMN WIDTHS]
        frmTemp.GridRoster.ColWidth(intColCounter) = (sinWidth / (intCols + 1))
        
        For intRowCounter = 1 To intRows
            frmTemp.GridRoster.Col = intColCounter
            frmTemp.GridRoster.Row = intRowCounter
            intLines = LineCount(frmTemp.GridRoster.Text)
            If intLines = 0 Then intLines = 1
            
            '[FIND CURRENT TEXT HEIGHT]
            sinHeight = frmTemp.TextHeight("A") * intLines
            If sinHeight = 0 Or sinHeight = frmTemp.TextHeight("A") Then
                sinHeight = frmTemp.TextHeight("A") * sinMod
            End If
            If frmTemp.TextWidth(frmTemp.GridRoster.Text) > frmTemp.GridRoster.ColWidth(intColCounter) Then
                sinHeight = frmTemp.TextHeight("A") * intLines * 2
            End If
            If sinHeight > frmTemp.GridRoster.RowHeight(intRowCounter) Then
                frmTemp.GridRoster.RowHeight(intRowCounter) = sinHeight
            End If
            
        Next intRowCounter
    Next intColCounter
    
    '[RESTORE STARTING GRID POSITION]
    frmTemp.GridRoster.Col = intCol
    frmTemp.GridRoster.Row = intRow

End Sub

Sub RebuildRoster()

    '[REBUILD ROSTER USING START-TIME, END-TIME AND INTERVAL]
    Dim intRow          As Integer
    Dim intCol          As Integer
    Dim dateTemp        As Date
    Dim dateStart       As Date
    Dim dateEnd         As Date
    Dim intInterval     As Integer
    Dim flagContinue    As Boolean
    Dim strInterval     As String
    Dim flagNextDay     As Boolean
    Dim frmTemp         As Form
    
    '[SET ACTIVE FORM]
    If flagScratch = True Then
        Set frmTemp = frmScratch
    Else
        Set frmTemp = frmRoster
    End If
    
    dateStart = DsDefault("StartTime")
    dateTemp = DsDefault("StartTime")
    dateEnd = DsDefault("EndTime")
    
    '[ADD 24 HOURS IF END TIME IS THE NEXT DAY]
    If dateEnd < dateStart Then dateEnd = dateEnd + CDate("12:00") + CDate("12:00")
    flagContinue = True
    
    '[DETERMINE INTERVAL]
    intInterval = frmSet.ComboIncrement.ItemData(DsDefault("Increment") - 1)
    
    Select Case intInterval
    Case 30
        strInterval = "00:30"
    Case 15
        strInterval = "00:15"
    Case Else
        strInterval = Str(Format(intInterval, "0#")) & ":00"
    End Select
    
    '[REDUCE ROWS TO TWO AND REBUILD]
    frmTemp.GridRoster.Rows = 2
    frmTemp.GridRoster.Row = 1
    
    '[BUILD ROWS]
    Do While flagContinue = True
        frmTemp.GridRoster.AddItem Format(dateTemp, "Medium Time") & Chr$(vbKeyTab) & Format(dateTemp + CDate(strInterval), "Medium Time"), (frmTemp.GridRoster.Rows - 1)
        '[OLD TIME FORMAT] -> "hh:mm AMPM"
        dateTemp = dateTemp + CDate(strInterval)
        If dateTemp > dateEnd Then flagContinue = False
    Loop
    
    '[REMOVE LAST BLANK ROW]
    frmTemp.GridRoster.Rows = (frmTemp.GridRoster.Rows - 1)

    

End Sub

Sub SetClassLabels()
    
    '[SET CLASS DESCRIPTIONS TO MATCH CLASS CHECK BOX LABELS]
    Dim intCounter      As Integer
    Dim intClassIndex   As Integer      '[LOCATION OF SELECTED ITEM IN COMBOCLASS LIST]
    Dim strBookmark     As String
    
    '[MOVE THROUGH DYNASET]
    If DsClass.EOF Or DsClass.BOF Then DsClass.MoveFirst
    strBookmark = DsClass.Bookmark      '[SAVE BOOKMARK]
    DsClass.MoveFirst
    intCounter = 0
    
    Do While Not DsClass.EOF
        frmSet.CheckClass(intCounter).Caption = DsClass("Description")
        intCounter = intCounter + 1
        DsClass.MoveNext
    Loop
    
    '[RESTORE BOOKMARK]
    DsClass.Bookmark = strBookmark
    
End Sub


Sub SetDayLabels()

    '[SET LABELS ON STAFF FORM TO THE ORDER SPECIFIED BY STARTDAY]
    Dim intCounter      As Integer
    Dim DsDayNames      As Recordset
    Dim strDay          As String
    
    Set DsDayNames = DBMain.OpenRecordset("DayNames", dbOpenTable)
    DsDayNames.MoveFirst
    DsDayNames.Edit
    
    For intCounter = 1 To 7
        frmSet.labelDay(intCounter - 1).Caption = ArrayWeek(DayOfWeek(intCounter)).LongDay
        strDay = "Day_" & Trim(Str(intCounter))
        '[UPDATE DAYS OF WEEK IN DAYNAMES DYNASET]
        DsDayNames(strDay) = ArrayWeek(DayOfWeek(intCounter)).LongDay
        strDay = "Date_" & Trim(Str(intCounter))
        DsDayNames(strDay) = DsDefault("StartDate") + (intCounter - 1)
    Next intCounter
    
    DsDayNames.Update
    DsDayNames.Close
    
End Sub

Sub SetGridTitles()

    '[SET GRID TITLES ON THE ROSTER GRID ACCORDING TO SELECTED START DAY]
    Dim intCounter      As Integer
    Dim intCol          As Integer
    Dim intRow          As Integer
    Dim intStartDay     As Integer
    Dim frmTemp         As Form
    
    '[SET ACTIVE FORM]
    If flagScratch = True Then
        Set frmTemp = frmScratch
    Else
        Set frmTemp = frmRoster
    End If
    
    '[SAVE CURRENT GRID POSITION]
    intRow = frmTemp.GridRoster.Row
    intCol = frmTemp.GridRoster.Col

    frmTemp.GridRoster.Row = 0
    
    '[CYCLE THROUGH COLUMNS AND SET TITLES]
    frmTemp.GridRoster.Col = 0
    frmTemp.GridRoster.Text = "Start"
    frmTemp.GridRoster.Col = 1
    frmTemp.GridRoster.Text = "End"
    
    For intCounter = 2 To 8
        frmTemp.GridRoster.Col = intCounter
        frmTemp.GridRoster.Text = ArrayWeek(DayOfWeek(intCounter - 1)).ShortDay
    Next intCounter

    '[RESTORE GRID POSITION]
    frmTemp.GridRoster.Row = intRow
    frmTemp.GridRoster.Col = intCol
    
End Sub

Function DayOfWeek(intDayNumber)

    '[FUNCTION TO RETURN INTEGER FOR DAY OF WEEK ARRAY]
    Dim intStartDay             As Integer
    Dim intReturnValue          As Integer
    
    intStartDay = DsDefault("StartDay")
    
    If intStartDay + (intDayNumber - 1) > 7 Then
        intReturnValue = (intDayNumber - 7) + (intStartDay - 1)
    Else
        intReturnValue = (intStartDay + (intDayNumber - 1))
    End If

    DayOfWeek = intReturnValue

End Function
Sub ShowStaffDetails()
        
    '[SHOW STAFF DETAILS FOR THE CURRENTLY SELECTED STAFF MEMBER IN THE DYNASET]
    If flagLoad = False Then Exit Sub
    Dim intCounter      As Integer
    Dim strField        As String
    Dim flagLabels      As Boolean
    Dim sinRate         As Single
    
    '[FLAG FOR DISPLAYING START/FINISH LABELS]
    flagLabels = False
    
    '[TEXT BOXES]
    frmSet.TextStaffID.Text = DsStaff("StaffID")
    '[THE " " IS REQUIRED IN CASE OF NULL VALUES IN THE DATABASE (ALTHOUGH THIS SHOULDN'T OCCUR)]
    frmSet.TextLastName.Text = Trim(DsStaff("LastName") & " ")
    frmSet.TextFirstName.Text = Trim(DsStaff("FirstName") & " ")
    frmSet.TextMiddleName.Text = Trim(DsStaff("MiddleName") & " ")
    
    frmSet.MaskHomePhone.Text = Trim(DsStaff("HomePhone") & " ")
    frmSet.MaskHourRate.Text = Trim(DsStaff("HourRate") & " ")
    If IsNull(DsStaff!PayType) Then frmSet.OptionPay(0).Value = True Else frmSet.OptionPay(DsStaff!PayType).Value = True
    
    frmSet.MaskBirthDate.Text = Trim(DsStaff("BirthDate") & " ")
    frmSet.MaskDateHired.Text = Trim(DsStaff("DateHired") & " ")
    
    frmSet.TextNote.Text = Trim(DsStaff("Note") & " ")
    frmSet.textStaffNotes.Text = Trim(DsStaff("StaffNote") & " ")
    
    '[MAX/MIN HOURS]
    If DsStaff("MinHours") > 0 Then frmSet.MaskMinHours.Text = DsStaff("MinHours") Else frmSet.MaskMinHours.Text = ""
    If DsStaff("MaxHours") > 0 Then frmSet.MaskMaxHours.Text = DsStaff("MaxHours") Else frmSet.MaskMaxHours.Text = ""

    '[STAFF AVAILABILITY DAYS]
    For intCounter = 1 To 7
        strField = "Day_" & Trim(Str(intCounter))
        '[CHANGED ----- If IsNull(DsStaff(strField).Value) Then frmSet.CheckDay(intCounter - 1).Value = vbChecked Else frmSet.CheckDay(intCounter - 1).Value = DsStaff(strField).Value
        '[NOW USE COMBO BOX WITH FOUR SETTINGS]
        If IsNull(DsStaff(strField).Value) Then
            frmSet.comboDayStatus(intCounter - 1).ListIndex = vbChecked
        Else
            frmSet.comboDayStatus(intCounter - 1).ListIndex = DsStaff(strField).Value
            '[IF INSIDE OR OUTSIDE, SET LABELS ON]
            If DsStaff(strField).Value > 1 Then flagLabels = True
        End If
        '[DAY START TIME]
        strField = "Start_" & Trim(Str(intCounter))
        If IsNull(DsStaff(strField).Value) Then frmSet.textStart(intCounter - 1).Text = Format(DsDefault("StartTime"), "Medium Time") Else frmSet.textStart(intCounter - 1).Text = Format(DsStaff(strField), "Medium Time")
        '[DAY END TIME]
        strField = "Finish_" & Trim(Str(intCounter))
        If IsNull(DsStaff(strField).Value) Then frmSet.textFinish(intCounter - 1).Text = Format(DsDefault("EndTime"), "Medium Time") Else frmSet.textFinish(intCounter - 1).Text = Format(DsStaff(strField).Value, "Medium Time")
    Next intCounter
    
    '[SHOW STAFF START/FINISH LABELS]
    If flagLabels = True Then
        frmSet.lbl_std(1).Visible = True
        frmSet.lbl_std(2).Visible = True
    Else
        frmSet.lbl_std(1).Visible = False
        frmSet.lbl_std(2).Visible = False
    End If
    
    '[STAFF AGE AND EMPLOYMENT PERIOD]
    ShowStaffInfo
    
    '[STAFF CLASSIFICATION CHECK BOXES]
    flagLabels = False
    For intCounter = 1 To 10
        strField = "Class_" & Trim(Str(intCounter))
        frmSet.CheckClass(intCounter - 1).Value = DsStaff(strField).Value
        If DsStaff(strField).Value = vbChecked Then flagLabels = True
        '[HOURLY RATE PER ROSTER]
        strField = "Rate_" & Trim(Str(intCounter))
        If IsNull(DsStaff(strField)) Then frmSet.maskRosterRate(intCounter - 1).Text = 0 Else frmSet.maskRosterRate(intCounter - 1).Text = DsStaff(strField).Value
    Next intCounter

    '[HOLIDAY INFORMATION]
    If IsDate(DsStaff!HolStart) Then frmSet.MaskHolStart = Format(DsStaff!HolStart, strDateFormat) Else frmSet.MaskHolStart = ""
    If IsDate(DsStaff!HolEnd) Then frmSet.MaskHolEnd = Format(DsStaff!HolEnd, strDateFormat) Else frmSet.MaskHolEnd = ""
    If IsNull(DsStaff!Holiday) Or DsStaff!Holiday.Value < 0 Then
        frmSet.CheckHoliday.Value = vbUnchecked
    Else
        frmSet.CheckHoliday.Value = DsStaff!Holiday.Value
    End If

    '[HIDE LABEL]
    If flagLabels = False Then
        frmSet.lbl_hr_rate.Visible = False
    Else
        frmSet.lbl_hr_rate.Visible = True
    End If

    '[SET SAVE BUTTON TO DISABLED]
    frmSet.cmdSave.Visible = False

End Sub


Sub SaveStaffDetails()

    '[SAVE STAFF DETAILS TO THE DYNASET FOR THE CURRENTLY DISPLAYED STAFF MEMBER]
    Dim intCounter      As Integer
    Dim strField        As String
    
    '[OPEN DYNASET FOR EDITING]
    DsStaff.Edit
    
        '[TEXT BOXES]
        '[CHECK TO SEE IF THE STAFF ID IS BLANK - IF SO, REPLACE WITH OLD STAFF ID]
        If frmSet.TextStaffID <> "" Then DsStaff("StaffID") = frmSet.TextStaffID.Text
        If frmSet.TextLastName.Text <> "" Then DsStaff("LastName") = frmSet.TextLastName.Text
        If frmSet.TextFirstName.Text <> "" Then DsStaff("FirstName") = frmSet.TextFirstName.Text
        DsStaff("MiddleName") = frmSet.TextMiddleName.Text
    
        DsStaff("HomePhone") = frmSet.MaskHomePhone.Text
        DsStaff("HourRate") = Val(frmSet.MaskHourRate.Text)
        If frmSet.OptionPay(0).Value = True Then DsStaff!PayType = 0 Else DsStaff!PayType = 1
        If IsDate(frmSet.MaskBirthDate.Text) Then DsStaff("BirthDate") = frmSet.MaskBirthDate.Text
        If IsDate(frmSet.MaskDateHired.Text) Then DsStaff("DateHired") = frmSet.MaskDateHired.Text
        If Len(Trim(frmSet.TextNote.Text)) > 254 Then frmSet.TextNote.Text = Left(frmSet.TextNote.Text, 254)
        DsStaff!Note = Trim(frmSet.TextNote.Text) & " "
        If Len(Trim(frmSet.textStaffNotes.Text)) > 254 Then frmSet.textStaffNotes.Text = Left(frmSet.textStaffNotes.Text, 254)
        DsStaff!StaffNote = Trim(frmSet.textStaffNotes.Text) & " "
                
        '[STAFF MAX/MIN HOURS]
        DsStaff("MinHours") = Val(frmSet.MaskMinHours.Text)
        DsStaff("MaxHours") = Val(frmSet.MaskMaxHours.Text)
                
        '[STAFF AVAILABILITY DAYS]
        For intCounter = 1 To 7
            strField = "Day_" & Trim(Str(intCounter))
            '[CHANGE ----- DsStaff(strField) = (frmSet.CheckDay(intCounter - 1).Value)
            '[NOW USES COMBO BOX WITH FOUR SETTINGS]
            If frmSet.comboDayStatus(intCounter - 1).ListIndex > -1 Then
                DsStaff(strField) = (frmSet.comboDayStatus(intCounter - 1).ListIndex)
            Else
                DsStaff(strField) = 1
            End If
            
            '[DAY START TIME AND DAY END TIME - ONLY FOR INSIDE/OUTSIDE SELECTIONS]
            If frmSet.comboDayStatus(intCounter - 1).ListIndex > 1 Then
            
                '[DAY START TIME]
                strField = "Start_" & Trim(Str(intCounter))
                If IsDate(frmSet.textStart(intCounter - 1).Text) Then
                    DsStaff(strField).Value = Format(frmSet.textStart(intCounter - 1).Text, "Medium Time")
                Else
                    DsStaff(strField).Value = Format(DsDefault("StartTime"), "Medium Time")
                End If
                
                '[DAY END TIME]
                strField = "Finish_" & Trim(Str(intCounter))
                If IsDate(frmSet.textFinish(intCounter - 1).Text) Then
                    DsStaff(strField).Value = Format(frmSet.textFinish(intCounter - 1).Text, "Medium Time")
                Else
                    DsStaff(strField).Value = Format(DsDefault("EndTime"), "Medium Time")
                End If
                
            End If
            
        Next intCounter
                
        '[STAFF CLASSIFICATION CHECK BOXES]
        For intCounter = 1 To 10
            strField = "Class_" & Trim(Str(intCounter))
            DsStaff(strField) = (frmSet.CheckClass(intCounter - 1).Value)
            strField = "Rate_" & Trim(Str(intCounter))
            DsStaff(strField).Value = frmSet.maskRosterRate(intCounter - 1).Text
        Next intCounter
    
        '[HOLIDAYS]
        If IsDate(frmSet.MaskHolStart) Then DsStaff!HolStart = frmSet.MaskHolStart
        If IsDate(frmSet.MaskHolEnd) Then DsStaff!HolEnd = frmSet.MaskHolEnd
        DsStaff!Holiday.Value = frmSet.CheckHoliday.Value
    
    '[COMMIT CHANGES TO DYNASET]
    DsStaff.Update

    '[SET SAVE BUTTON TO DISABLED]
    frmSet.cmdSave.Visible = False

End Sub
