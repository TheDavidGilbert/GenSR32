VERSION 4.00
Begin VB.Form frmControl 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Settings"
   ClientHeight    =   4485
   ClientLeft      =   1695
   ClientTop       =   1470
   ClientWidth     =   4530
   ControlBox      =   0   'False
   FillStyle       =   0  'Solid
   BeginProperty Font 
      name            =   "Arial"
      charset         =   1
      weight          =   400
      size            =   8.25
      underline       =   0   'False
      italic          =   0   'False
      strikethrough   =   0   'False
   EndProperty
   Height          =   4890
   HelpContextID   =   30
   Icon            =   "FRMCONTR.frx":0000
   Left            =   1635
   LinkTopic       =   "frmControl"
   MaxButton       =   0   'False
   ScaleHeight     =   4485
   ScaleWidth      =   4530
   ShowInTaskbar   =   0   'False
   Top             =   1125
   Width           =   4650
   Begin TabDlg.SSTab tabRoster 
      Height          =   4500
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4545
      _version        =   65536
      _extentx        =   8017
      _extenty        =   7938
      _stockprops     =   15
      caption         =   "Roster Details"
      forecolor       =   0
      tabsperrow      =   3
      tab             =   0
      taborientation  =   0
      tabs            =   3
      style           =   1
      tabmaxwidth     =   0
      tabheight       =   529
      tabcaption(0)   =   "Roster Details"
      tab(0).controlcount=   2
      tab(0).controlenabled=   -1  'True
      tab(0).control(0)=   "cmdReturn(1)"
      tab(0).control(1)=   "FrameStandard(4)"
      tabcaption(1)   =   "Roster Settings"
      tab(1).controlcount=   4
      tab(1).controlenabled=   0   'False
      tab(1).control(0)=   "FrameStandard(3)"
      tab(1).control(1)=   "FrameStandard(1)"
      tab(1).control(2)=   "FrameStandard(2)"
      tab(1).control(3)=   "cmdReturn(0)"
      tabcaption(2)   =   "Program Settings"
      tab(2).controlcount=   4
      tab(2).controlenabled=   0   'False
      tab(2).control(0)=   "FrameStandard(0)"
      tab(2).control(1)=   "frameDateFormat"
      tab(2).control(2)=   "FrameStandard(5)"
      tab(2).control(3)=   "cmdReturn(2)"
      Begin VB.Frame FrameStandard 
         Height          =   615
         Index           =   3
         Left            =   -74940
         TabIndex        =   43
         Top             =   360
         Width           =   4395
         Begin VB.ComboBox ComboStartDay 
            Height          =   330
            ItemData        =   "FRMCONTR.frx":08CA
            Left            =   2760
            List            =   "FRMCONTR.frx":08E3
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   180
            Width           =   1575
         End
         Begin VB.Label Label_std 
            Appearance      =   0  'Flat
            BackColor       =   &H80000002&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Starting Day of Roster Week"
            ForeColor       =   &H80000009&
            Height          =   255
            Index           =   0
            Left            =   60
            TabIndex        =   44
            Top             =   210
            Width           =   2655
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Frame FrameStandard 
         Height          =   1935
         Index           =   1
         Left            =   -74940
         TabIndex        =   31
         Top             =   960
         Width           =   4395
         Begin VB.OptionButton OptionTime 
            Caption         =   "End"
            Height          =   210
            Index           =   1
            Left            =   3540
            TabIndex        =   5
            Top             =   480
            Width           =   735
         End
         Begin VB.OptionButton OptionTime 
            Caption         =   "Start"
            Height          =   210
            Index           =   0
            Left            =   2760
            TabIndex        =   4
            Top             =   480
            Value           =   -1  'True
            Width           =   795
         End
         Begin VB.ComboBox ComboMinute 
            Height          =   330
            ItemData        =   "FRMCONTR.frx":0927
            Left            =   3570
            List            =   "FRMCONTR.frx":0937
            TabIndex        =   7
            Text            =   "ComboMinute"
            Top             =   750
            Width           =   765
         End
         Begin VB.ComboBox ComboHour 
            Height          =   330
            ItemData        =   "FRMCONTR.frx":094B
            Left            =   2760
            List            =   "FRMCONTR.frx":09A5
            TabIndex        =   6
            Text            =   "ComboHour"
            Top             =   750
            Width           =   765
         End
         Begin VB.ComboBox ComboIncrement 
            Height          =   330
            ItemData        =   "FRMCONTR.frx":0A09
            Left            =   2760
            List            =   "FRMCONTR.frx":0A3C
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   1530
            Width           =   1545
         End
         Begin VB.Label Label_std 
            AutoSize        =   -1  'True
            BackColor       =   &H80000004&
            BackStyle       =   0  'Transparent
            Caption         =   $"FRMCONTR.frx":0ABE
            Height          =   630
            Index           =   6
            Left            =   60
            TabIndex        =   36
            Top             =   480
            Width           =   2730
            WordWrap        =   -1  'True
         End
         Begin MSMask.MaskEdBox MaskDayLength 
            Height          =   330
            Left            =   2760
            TabIndex        =   35
            TabStop         =   0   'False
            Top             =   1140
            Width           =   1545
            _version        =   65536
            _extentx        =   2731
            _extenty        =   572
            _stockprops     =   109
            forecolor       =   -2147483646
            backcolor       =   -2147483644
            borderstyle     =   1
            enabled         =   0   'False
            maxlength       =   5
            format          =   "hh:mm"
         End
         Begin VB.Label Label_std 
            AutoSize        =   -1  'True
            BackColor       =   &H80000004&
            BackStyle       =   0  'Transparent
            Caption         =   "Length of working day (hh:mm)"
            Height          =   210
            Index           =   2
            Left            =   60
            TabIndex        =   34
            Top             =   1230
            Width           =   2700
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label_std 
            AutoSize        =   -1  'True
            BackColor       =   &H80000004&
            BackStyle       =   0  'Transparent
            Caption         =   "Roster increment (or shift length)"
            Height          =   210
            Index           =   1
            Left            =   60
            TabIndex        =   33
            Top             =   1590
            Width           =   2700
            WordWrap        =   -1  'True
         End
         Begin VB.Label labelHeading 
            Appearance      =   0  'Flat
            BackColor       =   &H80000002&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Autobuild Settings"
            ForeColor       =   &H80000009&
            Height          =   255
            Index           =   2
            Left            =   60
            TabIndex        =   32
            Top             =   150
            Width           =   4275
         End
      End
      Begin VB.Frame FrameStandard 
         Height          =   1215
         Index           =   0
         Left            =   -74940
         TabIndex        =   28
         Top             =   360
         Width           =   4395
         Begin VB.CheckBox CheckAllShifts 
            Alignment       =   1  'Right Justify
            Caption         =   "Require All Shifts Filled"
            Height          =   210
            Left            =   2220
            TabIndex        =   16
            Top             =   960
            Width           =   2100
         End
         Begin VB.CheckBox CheckDelete 
            Alignment       =   1  'Right Justify
            Caption         =   "Confirm Deletes"
            Height          =   210
            Left            =   2220
            TabIndex        =   15
            Top             =   720
            Width           =   2100
         End
         Begin VB.CheckBox CheckSounds 
            Alignment       =   1  'Right Justify
            Caption         =   "Sounds"
            Height          =   210
            Left            =   2220
            TabIndex        =   14
            Top             =   480
            Width           =   2100
         End
         Begin VB.Label Label_std 
            AutoSize        =   -1  'True
            BackColor       =   &H80000004&
            BackStyle       =   0  'Transparent
            Caption         =   "These settings affect the way GSR responds to your actions."
            Height          =   675
            Index           =   7
            Left            =   90
            TabIndex        =   30
            Top             =   480
            Width           =   2085
            WordWrap        =   -1  'True
         End
         Begin VB.Label labelHeading 
            Appearance      =   0  'Flat
            BackColor       =   &H80000002&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Feedback"
            ForeColor       =   &H80000009&
            Height          =   255
            Index           =   3
            Left            =   60
            TabIndex        =   29
            Top             =   180
            Width           =   4275
         End
      End
      Begin VB.Frame frameDateFormat 
         Height          =   1692
         Left            =   -74940
         TabIndex        =   25
         Top             =   1500
         Width           =   4392
         Begin VB.OptionButton OptionDateFormat 
            Caption         =   "Custom Format"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   19
            Top             =   1290
            Width           =   1995
         End
         Begin VB.OptionButton OptionDateFormat 
            Height          =   255
            Index           =   1
            Left            =   2196
            TabIndex        =   18
            Top             =   960
            Width           =   2055
         End
         Begin VB.OptionButton OptionDateFormat 
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   17
            Top             =   960
            Width           =   1995
         End
         Begin VB.Label labelHeading 
            Appearance      =   0  'Flat
            BackColor       =   &H80000002&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Displayed Date Format"
            ForeColor       =   &H80000009&
            Height          =   255
            Index           =   1
            Left            =   60
            TabIndex        =   27
            Top             =   180
            Width           =   4275
         End
         Begin VB.Label Label_std 
            AutoSize        =   -1  'True
            BackColor       =   &H80000004&
            BackStyle       =   0  'Transparent
            Caption         =   $"FRMCONTR.frx":0B1A
            Height          =   450
            Index           =   3
            Left            =   90
            TabIndex        =   26
            Top             =   480
            Width           =   4215
            WordWrap        =   -1  'True
         End
         Begin MSMask.MaskEdBox maskDateMask 
            Height          =   300
            Left            =   2196
            TabIndex        =   20
            Top             =   1260
            Width           =   2088
            _version        =   65536
            _extentx        =   3678
            _extenty        =   529
            _stockprops     =   109
            forecolor       =   -2147483640
            backcolor       =   -2147483643
            borderstyle     =   1
         End
      End
      Begin VB.Frame FrameStandard 
         Height          =   1155
         Index           =   2
         Left            =   -74940
         TabIndex        =   37
         Top             =   2880
         Width           =   4395
         Begin VB.ComboBox ComboBreakHour 
            Height          =   330
            ItemData        =   "FRMCONTR.frx":0B70
            Left            =   2760
            List            =   "FRMCONTR.frx":0BCA
            TabIndex        =   11
            Text            =   "ComboHour"
            Top             =   780
            Width           =   765
         End
         Begin VB.ComboBox ComboBreakMin 
            Height          =   330
            ItemData        =   "FRMCONTR.frx":0C2E
            Left            =   3570
            List            =   "FRMCONTR.frx":0C3E
            TabIndex        =   12
            Text            =   "ComboMinute"
            Top             =   780
            Width           =   765
         End
         Begin VB.ComboBox ComboBlockHour 
            Height          =   330
            ItemData        =   "FRMCONTR.frx":0C52
            Left            =   2760
            List            =   "FRMCONTR.frx":0CAC
            TabIndex        =   9
            Text            =   "ComboHour"
            Top             =   420
            Width           =   765
         End
         Begin VB.ComboBox ComboBlockMin 
            Height          =   330
            ItemData        =   "FRMCONTR.frx":0D10
            Left            =   3570
            List            =   "FRMCONTR.frx":0D20
            TabIndex        =   10
            Text            =   "ComboMinute"
            Top             =   420
            Width           =   765
         End
         Begin VB.Label Label_std 
            AutoSize        =   -1  'True
            BackColor       =   &H80000004&
            BackStyle       =   0  'Transparent
            Caption         =   "Time before break is applied."
            Height          =   210
            Index           =   8
            Left            =   60
            TabIndex        =   42
            Top             =   480
            Width           =   2205
            WordWrap        =   -1  'True
         End
         Begin VB.Label labelHeading 
            Appearance      =   0  'Flat
            BackColor       =   &H80000002&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Staff Breaks"
            ForeColor       =   &H80000009&
            Height          =   255
            Index           =   5
            Left            =   60
            TabIndex        =   41
            Top             =   150
            Width           =   2655
         End
         Begin VB.Label Label_std 
            AutoSize        =   -1  'True
            BackColor       =   &H80000004&
            BackStyle       =   0  'Transparent
            Caption         =   "Length of break."
            Height          =   210
            Index           =   9
            Left            =   60
            TabIndex        =   40
            Top             =   840
            Width           =   1995
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label_std 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000002&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Minutes"
            ForeColor       =   &H80000009&
            Height          =   255
            Index           =   13
            Left            =   3540
            TabIndex        =   39
            Top             =   150
            Width           =   765
         End
         Begin VB.Label Label_std 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000002&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Hours"
            ForeColor       =   &H80000009&
            Height          =   255
            Index           =   4
            Left            =   2760
            TabIndex        =   38
            Top             =   150
            Width           =   765
         End
      End
      Begin VB.Frame FrameStandard 
         Height          =   3675
         Index           =   4
         Left            =   60
         TabIndex        =   45
         Top             =   360
         Width           =   4395
         Begin MSGrid.Grid GridClass 
            Height          =   3195
            Left            =   60
            TabIndex        =   1
            Top             =   420
            Width           =   4275
            _version        =   65536
            _extentx        =   7541
            _extenty        =   5636
            _stockprops     =   77
            forecolor       =   -2147483640
            backcolor       =   -2147483643
            cols            =   3
            fixedcols       =   0
            scrollbars      =   0
            mouseicon       =   "FRMCONTR.frx":0D34
         End
         Begin VB.Label Label_std 
            Appearance      =   0  'Flat
            BackColor       =   &H80000002&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Roster Definitions"
            ForeColor       =   &H80000009&
            Height          =   255
            Index           =   10
            Left            =   60
            TabIndex        =   46
            Top             =   150
            Width           =   4275
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Frame FrameStandard 
         Height          =   855
         Index           =   5
         Left            =   -74940
         TabIndex        =   47
         Top             =   3180
         Width           =   4395
         Begin VB.Label Label_std 
            AutoSize        =   -1  'True
            BackColor       =   &H80000004&
            BackStyle       =   0  'Transparent
            Caption         =   "Click to expose NUKE command button."
            Height          =   210
            Index           =   11
            Left            =   60
            TabIndex        =   50
            Top             =   480
            Width           =   3615
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label_std 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Delete all roster shifts from the GSR database"
            ForeColor       =   &H80000009&
            Height          =   240
            Index           =   5
            Left            =   60
            TabIndex        =   48
            Top             =   210
            Width           =   3630
            WordWrap        =   -1  'True
         End
         Begin Threed.SSCommand cmdNukeCover 
            Height          =   600
            Left            =   3720
            TabIndex        =   21
            Top             =   180
            Width           =   600
            _version        =   65536
            _extentx        =   1058
            _extenty        =   1058
            _stockprops     =   78
            autosize        =   2
            picture         =   "FRMCONTR.frx":0D50
         End
         Begin Threed.SSCommand cmdNuke 
            Height          =   600
            Left            =   3720
            TabIndex        =   49
            Top             =   180
            Width           =   600
            _version        =   65536
            _extentx        =   1058
            _extenty        =   1058
            _stockprops     =   78
            autosize        =   2
            picture         =   "FRMCONTR.frx":162A
         End
      End
      Begin Threed.SSCommand cmdReturn 
         Height          =   360
         Index           =   2
         Left            =   -71220
         TabIndex        =   22
         Top             =   4080
         Width           =   705
         _version        =   65536
         _extentx        =   1244
         _extenty        =   635
         _stockprops     =   78
         caption         =   "C&lose"
         autosize        =   2
      End
      Begin Threed.SSCommand cmdReturn 
         Height          =   360
         Index           =   0
         Left            =   -71220
         TabIndex        =   13
         Top             =   4080
         Width           =   705
         _version        =   65536
         _extentx        =   1244
         _extenty        =   635
         _stockprops     =   78
         caption         =   "C&lose"
         autosize        =   2
      End
      Begin Threed.SSCommand cmdReturn 
         Height          =   360
         Index           =   1
         Left            =   3780
         TabIndex        =   2
         Top             =   4080
         Width           =   705
         _version        =   65536
         _extentx        =   1244
         _extenty        =   635
         _stockprops     =   78
         caption         =   "C&lose"
         autosize        =   2
      End
      Begin VB.Label Label_std 
         BackColor       =   &H80000004&
         BackStyle       =   0  'Transparent
         Caption         =   "Maximum Hours Allowed"
         Height          =   285
         Index           =   22
         Left            =   -74910
         TabIndex        =   24
         Top             =   810
         Width           =   3400
      End
      Begin VB.Label Label_std 
         BackColor       =   &H80000004&
         BackStyle       =   0  'Transparent
         Caption         =   "Minimum Hours Required"
         Height          =   315
         Index           =   21
         Left            =   -74910
         TabIndex        =   23
         Top             =   480
         Width           =   3400
      End
   End
   Begin VB.Image ImageSwitch 
      Height          =   240
      Index           =   2
      Left            =   510
      Picture         =   "FRMCONTR.frx":1944
      Top             =   4650
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image ImageSwitch 
      Height          =   240
      Index           =   1
      Left            =   300
      Picture         =   "FRMCONTR.frx":1A46
      Top             =   4650
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image ImageSwitch 
      Height          =   240
      Index           =   0
      Left            =   60
      Picture         =   "FRMCONTR.frx":1B48
      Top             =   4650
      Visible         =   0   'False
      Width           =   240
   End
End
Attribute VB_Name = "frmControl"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Option Explicit
'[----------------------------------------------]
'[frmControl        Control Panel form          ]
'[----------------------------------------------]
'[GSR                                           ]
'[generic staff roster system for Windows       ]
'[(C) ---------- David Gilbert, 1997            ]
'[----------------------------------------------]
'[version                           2.1         ]
'[initial date                      21/07/1997  ]
'[----------------------------------------------]



Private Sub CheckAllShifts_Click()

    '[REV: 3.00.28]
    '[CHANGE PUBLIC VARIABLE ALLSHIFTS]
    If frmControl.CheckAllShifts.Value = 1 Then
        flagAllShifts = True
    Else
        flagAllShifts = False
    End If

End Sub


Private Sub CheckAllShifts_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Exception will be produced if all roster shifts are not filled."

End Sub


Private Sub CheckDelete_Click()
    
    '[CHANGE PUBLIC VARIABLE DELETE CONFIRMATION]
    If frmControl.CheckDelete.Value = 1 Then
        flagDeleteConfirm = True
    Else
        flagDeleteConfirm = False
    End If

End Sub

Private Sub CheckDelete_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Enable/disable the delete confirmation dialog."

End Sub


Private Sub CheckSounds_Click()

    '[CHANGE PUBLIC VARIABLE SOUNDS]
    If frmControl.CheckSounds.Value = 1 Then
        flagSounds = True
    Else
        flagSounds = False
    End If

End Sub

Private Sub CheckSounds_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Enable/disable message sounds."

End Sub


Private Sub cmdNuke_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Permanently delete ALL roster records."

End Sub

Private Sub cmdNukeCover_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Reveal the 'NUKE' command button."

End Sub


Private Sub cmdReturn_Click(Index As Integer)

    '[HIDE CONTROL FORM]
    frmControl.Hide

End Sub

Private Sub cmdReturn_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Close this window and return to the roster form."
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub ComboBlockHour_Click()

    '[SET DYNASET VALUE]
    DsDefault.Edit
        DsDefault("BlockHour") = frmControl.ComboBlockHour.Text
    DsDefault.Update


End Sub


Private Sub ComboBlockMin_Click()
    
    '[SET DYNASET VALUE]
    DsDefault.Edit
        DsDefault("BlockMin") = frmControl.ComboBlockMin.Text
    DsDefault.Update

End Sub


Private Sub ComboBreakHour_Click()
    
    '[SET DYNASET VALUE]
    DsDefault.Edit
        DsDefault("BreakHour") = frmControl.ComboBreakHour.Text
    DsDefault.Update

End Sub


Private Sub ComboBreakMin_Click()
    
    '[SET DYNASET VALUE]
    DsDefault.Edit
        DsDefault("BreakMin") = frmControl.ComboBreakMin.Text
    DsDefault.Update

End Sub


Private Sub GridClass_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii
    
    Case vbKeyReturn
        '[CAPTURE ENTER KEY PRESSED]
        Call GridClass_DblClick
    Case Else
    
    End Select
    
End Sub

Private Sub GridClass_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    If X > frmControl.GridClass.ColPos(2) Then
        StatusBar "Activate/deactivate the selected roster.  Only active roster can be modified and processed."
    ElseIf X > frmControl.GridClass.ColPos(1) Then
        StatusBar "Modify the roster description, a 20 character identifier for the selected roster."
    Else
        StatusBar "Modify the roster identifier, a three character short form for the roster name."
    End If
    '[---------------------------------------------------------------------------------]

End Sub

Private Sub Label_std_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    Select Case Index
    Case 0
        StatusBar "Select the starting day for your roster period."
    Case 1
        StatusBar "Select the appropriate roster increment for autobuilding rosters."
    Case 2
        StatusBar "Displays the day length calculated from the start time and end time."
    Case 3
        StatusBar "Select the desired display format for dates."
    Case 5
        StatusBar "Permanently delete ALL roster records."
    Case 6
        StatusBar "Select the start/end time for your rosters."
    Case 7
        StatusBar "Modify notification aspects of GSR."
    Case Else
    End Select
    
End Sub


Private Sub cmdNuke_Click()

    '[***********************************************************************]
    '[WARNING - DELETE ALL ROSTERS CURRENTLY STORED IN THE DATABASE - WARNING]
    '[***********************************************************************]
    '[THIS COMMAND IS NON-REVERSABLE, POPUP YES/NO DIALOG]
    Dim Msg As String
    Dim Style
    Dim Response
    Dim Title
    
    Msg = "Caution - This action will permanently delete and reset all rosters." & strBreak & strBreak & "This action is not reversible and should be used with extreme care." & strBreak & strBreak & "Do you wish to continue and delete all rosters ?"
    Style = vbYesNo                         ' Define buttons.
    Title = "Confirm Roster Deletion"     ' Define title.
    Response = gsrMsg(Msg, Style, Title)
    If Response = vbYes Then    ' User chose Yes.
        '[DELETE ROSTERS HERE]
        '[REBUILD ROSTER DYNASET WITH ALL RECORDS]
        Dim SQLStmt         As String
        
        '[SELECT APPROPRIATE RECORDS]
        SQLStmt = "SELECT * FROM Roster"
        Set DsRoster = DBMain.OpenRecordset(SQLStmt, dbOpenDynaset)
                
        '[CLEAR CURRENT DYNASET AND PREPARE]
        If DsRoster.EOF And DsRoster.BOF Then
            '[RESIZE NUKECOVER BUTTON]
            frmControl.cmdNukeCover.Visible = True
            frmControl.cmdNuke.Visible = False
            Exit Sub
        End If
        
        DsRoster.MoveLast
        Do While DsRoster.RecordCount > 0
            DsRoster.MoveFirst
            DsRoster.Delete
        Loop
        
        '[MOVE TO FIRST ITEM IN LIST]
        frmRoster.ComboClass.ListIndex = 0
        
    End If
    
    '[RESIZE NUKECOVER BUTTON]
    frmControl.cmdNukeCover.Visible = True
    frmControl.cmdNuke.Visible = False
    '[***********************************************************************]
    '[WARNING - DELETE ALL ROSTERS CURRENTLY STORED IN THE DATABASE - WARNING]
    '[***********************************************************************]
    
End Sub
Private Sub cmdNukeCover_Click()

    '[HIDE NUKE COVER]
    frmControl.cmdNukeCover.Visible = False
    frmControl.cmdNuke.Visible = True
    
End Sub




Private Sub ComboHour_Click()

    ProcessDayLength
    
End Sub


Private Sub ComboIncrement_Click()

    '[TAKE CHANGES MADE BY USER TO INCREMENT AND APPLY TO DEFAULT DYNASET]
    If ComboIncrement.ListIndex <> DsDefault("Increment") - 1 And ComboIncrement.ListIndex > -1 Then
        DsDefault.Edit
            DsDefault("Increment") = ComboIncrement.ListIndex + 1
        DsDefault.Update
    End If

End Sub


Private Sub ComboMinute_Click()

    ProcessDayLength
    
End Sub

Private Sub ComboStartDay_Click()

    '[TAKE CHANGES MADE BY USER TO START DAY AND APPLY TO DEFAULT DYNASET]
    If ComboStartDay.ListIndex <> DsDefault("StartDay") - 1 And ComboStartDay.ListIndex > -1 Then
        DsDefault.Edit
            DsDefault("StartDay") = ComboStartDay.ListIndex + 1
        DsDefault.Update
    End If
    
    '[SET GRID TITLES ON ROSTER GRID]
    SetGridTitles
    
    '[APPLY DAY LABELS]
    SetDayLabels
    
End Sub


Private Sub Form_Load()

      '[DEBUG]
    Call LogToFile("=-CONTROL FORM - LOAD SUB----------------=")

    '[SET STARTING POSITION FOR FORM]
    frmControl.Top = DsDefault("ControlTop")
    frmControl.Left = DsDefault("ControlLeft")
    '[CHECK FOR > WIDTH, HEIGHT]
    If frmControl.Top > Screen.Height Or frmControl.Left > Screen.Width Then
        '[center form on screen]
        frmControl.Top = (Screen.Height / 2) - (frmControl.Height / 2)
        frmControl.Left = (Screen.Width / 2) - (frmControl.Width / 2)
    End If

    '[Resize controls and grids to match]
    frmControl.GridClass.ColWidth(2) = frmControl.ImageSwitch(0).Width
    frmControl.GridClass.ColWidth(0) = (GridClass.Width - frmControl.GridClass.ColWidth(2)) * 0.25
    frmControl.GridClass.ColWidth(1) = (GridClass.Width - frmControl.GridClass.ColWidth(2)) * 0.75
    
    '[DEBUG]
    Call LogToFile("- filling class grid")
    '[CALL ROUTINE TO FILL GRID]
    FillClassGrid
    
    '[DEBUG]
    Call LogToFile("- setting start day and increments")
    '[PLACE DEFAULT VALUES IN COMBOBOXES AND MASK BOXES]
    ComboStartDay.ListIndex = DsDefault("StartDay") - 1
    ComboIncrement.ListIndex = DsDefault("Increment") - 1
    
    '[DEBUG]
    Call LogToFile("- setting start hour and minutes")
    '[SET START HOUR AND MINUTE]
    frmControl.ComboHour = Format(Hour(DsDefault("StartTime")), "0#")
    frmControl.ComboMinute = Format(Minute(DsDefault("StartTime")), "0#")
    ProcessDayLength
    
    '[DEBUG]
    Call LogToFile("- setting date captions")
    '[SET DATE FORMAT MASK]
    frmControl.optionDateFormat(0).Caption = Format(Date, strDateFormat)
    frmControl.optionDateFormat(1).Caption = Format(Date, "Medium Date")
    '[PLACE CUSTOM DATE MASK IN CONTROL]
    frmControl.maskDateMask = DsDefault("CustomDate")
    
End Sub


Private Sub GridClass_DblClick()

    '[USER DOUBLE CLICKED ON CLASS GRID, POPUP INPUT BOX FOR NEW VALUE]
    Dim strTemp             As String   '[Temporary Variable for Cell Content]
    Dim strMessage          As String
    Dim strTitle            As String
    Dim strOldClass         As String
    Dim SQLStmt             As String
    Dim intColCounter          As Integer
    Dim intCol, intRow      As Integer
    
    '[SAVE CURRENT LOCATION]
    intCol = GridClass.Col
    intRow = GridClass.Row
    
    '[EXIT IF TITLE ROW CLICKED]
    If intRow = 0 Then Exit Sub
    
    Select Case GridClass.Col
    Case 0
        strMessage = "The roster code is used as a short form to identify which rosters an employee belongs to.  You are allowed to use up to three alpha-numeric characters to define this roster."
    Case 1
        strMessage = "The roster description is used on reports to provide a longer identifier for each roster.  You may use up to 20 alpha-numeric characters in this field."
    End Select
    
    strTitle = "Class Definitions"
    strOldClass = GridClass.Text
    strTemp = GridClass.Text
    
    If GridClass.Col = 2 Then
        '[ENABLED/DISABLED]
        If GridClass.Text = vbChecked Then
            GridClass.Text = vbUnchecked
            frmControl.GridClass.Picture = frmControl.ImageSwitch(constCritical).Picture
        Else
            GridClass.Text = vbChecked
            frmControl.GridClass.Picture = frmControl.ImageSwitch(constWarning).Picture
        End If
    Else
        '[GET USER INPUT]
        strTemp = InputBox(strMessage, strTitle, strTemp)
        '[PROCESS USER INPUT]
        If Not strTemp = "" Or strTemp = GridClass.Text Then
            If GridClass.Col = 0 Then
            '[CHECK FOR EXISTANCE OF NEW CODE]
                SQLStmt = "Code = '" & strTemp & "'"
                DsClass.FindFirst SQLStmt
                If Not DsClass.NoMatch Then
                    '[RESTORE LOCATION IN GRID]
                    GridClass.Col = intCol
                    GridClass.Row = intRow
                    Exit Sub
                End If
                If Len(strTemp) > 3 Then GridClass.Text = Left$(UCase$(strTemp), 3) Else GridClass.Text = UCase$(strTemp)
            Else
                '[CHECK FOR EXISTANCE OF NEW CODE]
                SQLStmt = "Description = '" & strTemp & "'"
                DsClass.FindFirst SQLStmt
                If Not DsClass.NoMatch Then
                    '[RESTORE LOCATION IN GRID]
                    GridClass.Col = intCol
                    GridClass.Row = intRow
                    Exit Sub
                End If
                If Len(strTemp) > 20 Then GridClass.Text = Left$(strTemp, 3) Else GridClass.Text = strTemp
            End If
        Else
            '[RESTORE LOCATION IN GRID]
            GridClass.Col = intCol
            GridClass.Row = intRow
            Exit Sub
        End If
    End If
    
    '[PLACE VALUES IN DYNASET]
    DsClass.MoveFirst
    For intColCounter = 1 To 10
        frmControl.GridClass.Row = intColCounter
        frmControl.GridClass.Col = 0
        DsClass.Edit
            DsClass("Code") = frmControl.GridClass.Text
            frmControl.GridClass.Col = 1
            DsClass("Description") = frmControl.GridClass.Text
            frmControl.GridClass.Col = 2
            DsClass("Active") = frmControl.GridClass.Text
        DsClass.Update
        DsClass.MoveNext
    Next intColCounter
    
    '[SET CLASS LABELS ON STAFF FORM]
    Call SetClassLabels
    '[FILL ROSTER LIST]
    Call FillRosterList
    
    '[RESTORE LOCATION IN GRID]
    GridClass.Col = intCol
    GridClass.Row = intRow
    
End Sub




Private Sub maskDateMask_Change()

    '[CHECK TO SEE IF THIS IS A DATE]
    If Not IsDate(Format(Now, frmControl.maskDateMask)) Then
        '[NOT DATE FORMAT, DISABLE OPTION 3]
        frmControl.optionDateFormat(2).Enabled = False
        '[IF OPTION 3 WAS CHECKED, MOVE TO FIRST OPTION]
        If frmControl.optionDateFormat(2).Value = True Then frmControl.optionDateFormat(0).Value = True
    Else
        '[IS DATE FORMAT, ENABLE OPTION 3]
        frmControl.optionDateFormat(2).Enabled = True
    End If

End Sub

Private Sub mnuClose_Click()

    '[HIDE FORM]
    frmControl.Hide
    mdiMain.ZOrder

End Sub



Private Sub optionDateFormat_Click(Index As Integer)

    '[SET GLOBAL VARIABLE strDateFormat TO THE SELECTED DATE FORMAT]
    Dim flagSave        As Boolean
    
    Select Case Index
    Case 0      '[SHORT DATE]
        strDateFormat = "Short Date"
    Case 1      '[MEDIUM DATE]
        strDateFormat = "Medium Date"
    Case 2      '[CUSTOM DATE]
        strDateFormat = maskDateMask.Text
    Case Else
    End Select


    '[CHANGE ALL DISPLAYED DATE FORMATS ON CONTROLS]
    frmRoster.MaskDate.Format = strDateFormat
    frmRoster.MaskDate.Text = Format(frmRoster.MaskDate.Text, strDateFormat)
    '[CHECK TO SEE IF STAFF SAVE COMMAND IS VISIBLE]
    If frmStaff.cmdSave.Visible = True Then flagSave = True Else flagSave = False
    frmStaff.MaskBirthDate.Format = strDateFormat
    frmStaff.MaskBirthDate.Text = Format(frmStaff.MaskBirthDate.Text, strDateFormat)
    frmStaff.MaskDateHired.Format = strDateFormat
    frmStaff.MaskDateHired.Text = Format(frmStaff.MaskDateHired.Text, strDateFormat)
    frmStaff.MaskHolStart.Format = strDateFormat
    frmStaff.MaskHolStart.Text = Format(frmStaff.MaskHolStart.Text, strDateFormat)
    frmStaff.MaskHolEnd.Format = strDateFormat
    frmStaff.MaskHolEnd.Text = Format(frmStaff.MaskHolEnd.Text, strDateFormat)
    '[RESET STAFF SAVE COMMAND BUTTON STATE]
    If flagSave = False Then frmStaff.cmdSave.Visible = False
    

End Sub

Private Sub optionDateFormat_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Select the display format for dates."

End Sub





Private Sub OptionTIme_Click(Index As Integer)
    
    Select Case Index
    Case 0
        ComboHour = Format(Hour(DsDefault("StartTime")), "0#")
        ComboMinute = Format(Minute(DsDefault("StartTime")), "0#")
    Case 1
        ComboHour = Format(Hour(DsDefault("EndTime")), "0#")
        ComboMinute = Format(Minute(DsDefault("EndTime")), "0#")
    End Select

End Sub

Private Sub OptionTIme_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    Select Case Index
    Case 0
        StatusBar "Display/modify the roster START time."
    Case 1
        StatusBar "Display/modify the roster END time."
    Case Else
    End Select

End Sub


