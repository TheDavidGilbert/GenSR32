VERSION 4.00
Begin VB.Form frmSet 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "GSR Control Panel"
   ClientHeight    =   5190
   ClientLeft      =   3690
   ClientTop       =   4275
   ClientWidth     =   7395
   ControlBox      =   0   'False
   BeginProperty Font 
      name            =   "Arial"
      charset         =   1
      weight          =   400
      size            =   8.25
      underline       =   0   'False
      italic          =   0   'False
      strikethrough   =   0   'False
   EndProperty
   Height          =   5595
   Icon            =   "FRMSET_A.frx":0000
   Left            =   3630
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   7395
   Top             =   3930
   Width           =   7515
   Begin VB.PictureBox picSetBar 
      Align           =   3  'Align Left
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   4905
      Left            =   0
      ScaleHeight     =   4905
      ScaleWidth      =   585
      TabIndex        =   86
      Top             =   0
      Width           =   585
      Begin Threed.SSCommand cmdPrintSet 
         Height          =   570
         Left            =   0
         TabIndex        =   4
         Top             =   2460
         Width           =   570
         _version        =   65536
         _extentx        =   1005
         _extenty        =   1005
         _stockprops     =   78
         outline         =   0   'False
         autosize        =   2
         picture         =   "FRMSET_A.frx":08CA
      End
      Begin Threed.SSCommand cmdSetClose 
         Cancel          =   -1  'True
         Height          =   570
         Left            =   0
         TabIndex        =   5
         Top             =   4260
         Width           =   570
         _version        =   65536
         _extentx        =   1005
         _extenty        =   1005
         _stockprops     =   78
         outline         =   0   'False
         autosize        =   2
         picture         =   "FRMSET_A.frx":11A4
      End
      Begin Threed.SSCommand cmdProgSet 
         Height          =   570
         Left            =   0
         TabIndex        =   3
         Top             =   1860
         Width           =   570
         _version        =   65536
         _extentx        =   1005
         _extenty        =   1005
         _stockprops     =   78
         outline         =   0   'False
         autosize        =   2
         picture         =   "FRMSET_A.frx":15F6
      End
      Begin Threed.SSCommand cmdRosSet 
         Height          =   570
         Left            =   0
         TabIndex        =   2
         Top             =   1260
         Width           =   570
         _version        =   65536
         _extentx        =   1005
         _extenty        =   1005
         _stockprops     =   78
         outline         =   0   'False
         autosize        =   2
         picture         =   "FRMSET_A.frx":2300
      End
      Begin Threed.SSCommand cmdRosters 
         Height          =   570
         Left            =   0
         TabIndex        =   1
         Top             =   660
         Width           =   570
         _version        =   65536
         _extentx        =   1005
         _extenty        =   1005
         _stockprops     =   78
         outline         =   0   'False
         autosize        =   2
         picture         =   "FRMSET_A.frx":2BDA
      End
      Begin Threed.SSCommand cmdStaff 
         Height          =   570
         Left            =   0
         TabIndex        =   0
         Top             =   60
         Width           =   570
         _version        =   65536
         _extentx        =   1005
         _extenty        =   1005
         _stockprops     =   78
         autosize        =   2
         picture         =   "FRMSET_A.frx":34B4
      End
   End
   Begin Threed.SSPanel panelStatusBar 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   85
      Top             =   4905
      Width           =   7395
      _version        =   65536
      _extentx        =   13044
      _extenty        =   503
      _stockprops     =   15
      caption         =   " Status Bar Message Area"
      forecolor       =   16777215
      backcolor       =   8388608
      bevelwidth      =   0
      borderwidth     =   0
      bevelouter      =   0
      roundedcorners  =   0   'False
      outline         =   -1  'True
      floodcolor      =   -2147483648
      floodshowpct    =   0   'False
      alignment       =   2
   End
   Begin VB.Frame frameStaff 
      Caption         =   "Staff Details"
      Height          =   4875
      Left            =   600
      TabIndex        =   87
      Top             =   0
      Width           =   6795
      Begin VB.ListBox ListStaff 
         Appearance      =   0  'Flat
         Height          =   4230
         Left            =   60
         Sorted          =   -1  'True
         TabIndex        =   6
         Top             =   540
         Width           =   1935
      End
      Begin TabDlg.SSTab tabStaff 
         Height          =   4575
         HelpContextID   =   40
         Left            =   2040
         TabIndex        =   88
         Top             =   180
         Width           =   4695
         _version        =   65536
         _extentx        =   8281
         _extenty        =   8070
         _stockprops     =   15
         caption         =   "Staff Details"
         tabsperrow      =   4
         tab             =   0
         taborientation  =   0
         tabs            =   4
         style           =   1
         tabmaxwidth     =   0
         tabheight       =   529
         wordwrap        =   0   'False
         tabcaption(0)   =   "Staff Details"
         tab(0).controlcount=   21
         tab(0).controlenabled=   -1  'True
         tab(0).control(0)=   "ShapeBorder"
         tab(0).control(1)=   "labelInfo"
         tab(0).control(2)=   "MaskHomePhone"
         tab(0).control(3)=   "MaskBirthDate"
         tab(0).control(4)=   "MaskDateHired"
         tab(0).control(5)=   "MaskHourRate"
         tab(0).control(6)=   "label_std(1)"
         tab(0).control(7)=   "label_std(2)"
         tab(0).control(8)=   "label_std(3)"
         tab(0).control(9)=   "label_std(0)"
         tab(0).control(10)=   "label_std(7)"
         tab(0).control(11)=   "label_std(6)"
         tab(0).control(12)=   "label_std(5)"
         tab(0).control(13)=   "label_std(4)"
         tab(0).control(14)=   "OptionPay(1)"
         tab(0).control(15)=   "OptionPay(0)"
         tab(0).control(16)=   "TextLastName"
         tab(0).control(17)=   "TextFirstName"
         tab(0).control(18)=   "TextMiddleName"
         tab(0).control(19)=   "TextStaffID"
         tab(0).control(20)=   "TextNote"
         tabcaption(1)   =   "Availability"
         tab(1).controlcount=   38
         tab(1).controlenabled=   0   'False
         tab(1).control(0)=   "labelDay(0)"
         tab(1).control(1)=   "labelDay(1)"
         tab(1).control(2)=   "labelDay(2)"
         tab(1).control(3)=   "labelDay(3)"
         tab(1).control(4)=   "labelDay(4)"
         tab(1).control(5)=   "labelDay(5)"
         tab(1).control(6)=   "labelDay(6)"
         tab(1).control(7)=   "lbl_std(1)"
         tab(1).control(8)=   "lbl_std(2)"
         tab(1).control(9)=   "MaskMinHours"
         tab(1).control(10)=   "MaskMaxHours"
         tab(1).control(11)=   "label_std(31)"
         tab(1).control(12)=   "label_std(32)"
         tab(1).control(13)=   "label_std(33)"
         tab(1).control(14)=   "label_std(34)"
         tab(1).control(15)=   "label_std(35)"
         tab(1).control(16)=   "label_std(36)"
         tab(1).control(17)=   "comboDayStatus(0)"
         tab(1).control(18)=   "comboDayStatus(1)"
         tab(1).control(19)=   "comboDayStatus(2)"
         tab(1).control(20)=   "comboDayStatus(3)"
         tab(1).control(21)=   "comboDayStatus(4)"
         tab(1).control(22)=   "comboDayStatus(5)"
         tab(1).control(23)=   "comboDayStatus(6)"
         tab(1).control(24)=   "textStart(0)"
         tab(1).control(25)=   "textStart(1)"
         tab(1).control(26)=   "textStart(2)"
         tab(1).control(27)=   "textStart(3)"
         tab(1).control(28)=   "textStart(4)"
         tab(1).control(29)=   "textStart(5)"
         tab(1).control(30)=   "textStart(6)"
         tab(1).control(31)=   "textFinish(0)"
         tab(1).control(32)=   "textFinish(1)"
         tab(1).control(33)=   "textFinish(2)"
         tab(1).control(34)=   "textFinish(3)"
         tab(1).control(35)=   "textFinish(4)"
         tab(1).control(36)=   "textFinish(5)"
         tab(1).control(37)=   "textFinish(6)"
         tabcaption(2)   =   "Rosters "
         tab(2).controlcount=   32
         tab(2).controlenabled=   0   'False
         tab(2).control(0)=   "cmdBaseRate(9)"
         tab(2).control(1)=   "cmdBaseRate(8)"
         tab(2).control(2)=   "cmdBaseRate(7)"
         tab(2).control(3)=   "cmdBaseRate(6)"
         tab(2).control(4)=   "cmdBaseRate(5)"
         tab(2).control(5)=   "cmdBaseRate(4)"
         tab(2).control(6)=   "cmdBaseRate(3)"
         tab(2).control(7)=   "cmdBaseRate(2)"
         tab(2).control(8)=   "cmdBaseRate(1)"
         tab(2).control(9)=   "cmdBaseRate(0)"
         tab(2).control(10)=   "lbl_hr_rate"
         tab(2).control(11)=   "maskRosterRate(9)"
         tab(2).control(12)=   "maskRosterRate(8)"
         tab(2).control(13)=   "maskRosterRate(7)"
         tab(2).control(14)=   "maskRosterRate(6)"
         tab(2).control(15)=   "maskRosterRate(5)"
         tab(2).control(16)=   "maskRosterRate(4)"
         tab(2).control(17)=   "maskRosterRate(3)"
         tab(2).control(18)=   "maskRosterRate(2)"
         tab(2).control(19)=   "maskRosterRate(1)"
         tab(2).control(20)=   "maskRosterRate(0)"
         tab(2).control(21)=   "label_std(37)"
         tab(2).control(22)=   "CheckClass(9)"
         tab(2).control(23)=   "CheckClass(8)"
         tab(2).control(24)=   "CheckClass(7)"
         tab(2).control(25)=   "CheckClass(6)"
         tab(2).control(26)=   "CheckClass(5)"
         tab(2).control(27)=   "CheckClass(4)"
         tab(2).control(28)=   "CheckClass(3)"
         tab(2).control(29)=   "CheckClass(2)"
         tab(2).control(30)=   "CheckClass(1)"
         tab(2).control(31)=   "CheckClass(0)"
         tabcaption(3)   =   "Holidays"
         tab(3).controlcount=   18
         tab(3).controlenabled=   0   'False
         tab(3).control(0)=   "txtHolDays"
         tab(3).control(1)=   "CheckHoliday"
         tab(3).control(2)=   "label_std(41)"
         tab(3).control(3)=   "label_std(40)"
         tab(3).control(4)=   "label_std(39)"
         tab(3).control(5)=   "label_std(38)"
         tab(3).control(6)=   "MaskHolStart"
         tab(3).control(7)=   "MaskHolEnd"
         tab(3).control(8)=   "cmdHolStartToday"
         tab(3).control(9)=   "cmdHolEndToday"
         tab(3).control(10)=   "cmdHolStartPlus"
         tab(3).control(11)=   "cmdHolStartMinus"
         tab(3).control(12)=   "cmdholEndPlus"
         tab(3).control(13)=   "cmdHolEndMinus"
         tab(3).control(14)=   "cmdStartWeekPlus"
         tab(3).control(15)=   "cmdStartWeekMinus"
         tab(3).control(16)=   "cmdendWeekPlus"
         tab(3).control(17)=   "cmdEndWeekMinus"
         Begin VB.TextBox TextNote 
            DataField       =   "Note"
            Height          =   1515
            Left            =   60
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   17
            Text            =   "FRMSET_A.frx":3D8E
            Top             =   2940
            Width           =   4515
         End
         Begin VB.TextBox TextStaffID 
            DataField       =   "StaffID"
            Height          =   315
            Left            =   60
            MaxLength       =   10
            TabIndex        =   7
            Text            =   "StaffID"
            Top             =   670
            Width           =   1060
         End
         Begin VB.TextBox TextMiddleName 
            DataField       =   "MiddleName"
            Height          =   315
            Left            =   3360
            MaxLength       =   25
            TabIndex        =   10
            Text            =   "MiddleName"
            Top             =   670
            Width           =   1200
         End
         Begin VB.TextBox TextFirstName 
            DataField       =   "FirstName"
            Height          =   315
            Left            =   2260
            MaxLength       =   25
            TabIndex        =   9
            Text            =   "FirstName"
            Top             =   670
            Width           =   1060
         End
         Begin VB.TextBox TextLastName 
            DataField       =   "LastName"
            Height          =   315
            Left            =   1160
            MaxLength       =   25
            TabIndex        =   8
            Text            =   "LastName"
            Top             =   670
            Width           =   1060
         End
         Begin VB.TextBox textFinish 
            Height          =   315
            Index           =   6
            Left            =   -71520
            TabIndex        =   38
            Top             =   2910
            Width           =   1000
         End
         Begin VB.TextBox textFinish 
            Height          =   315
            Index           =   5
            Left            =   -71535
            TabIndex        =   35
            Top             =   2550
            Width           =   1000
         End
         Begin VB.TextBox textFinish 
            Height          =   315
            Index           =   4
            Left            =   -71535
            TabIndex        =   32
            Top             =   2190
            Width           =   1000
         End
         Begin VB.TextBox textFinish 
            Height          =   315
            Index           =   3
            Left            =   -71535
            TabIndex        =   29
            Top             =   1830
            Width           =   1000
         End
         Begin VB.TextBox textFinish 
            Height          =   315
            Index           =   2
            Left            =   -71535
            TabIndex        =   26
            Top             =   1470
            Width           =   1000
         End
         Begin VB.TextBox textFinish 
            Height          =   315
            Index           =   1
            Left            =   -71535
            TabIndex        =   23
            Top             =   1110
            Width           =   1000
         End
         Begin VB.TextBox textFinish 
            Height          =   315
            Index           =   0
            Left            =   -71535
            TabIndex        =   20
            Top             =   750
            Width           =   1000
         End
         Begin VB.TextBox textStart 
            Height          =   315
            Index           =   6
            Left            =   -72600
            TabIndex        =   37
            Top             =   2910
            Width           =   1000
         End
         Begin VB.TextBox textStart 
            Height          =   315
            Index           =   5
            Left            =   -72600
            TabIndex        =   34
            Top             =   2550
            Width           =   1000
         End
         Begin VB.TextBox textStart 
            Height          =   315
            Index           =   4
            Left            =   -72600
            TabIndex        =   31
            Top             =   2190
            Width           =   1000
         End
         Begin VB.TextBox textStart 
            Height          =   315
            Index           =   3
            Left            =   -72600
            TabIndex        =   28
            Top             =   1830
            Width           =   1000
         End
         Begin VB.TextBox textStart 
            Height          =   315
            Index           =   2
            Left            =   -72600
            TabIndex        =   25
            Top             =   1470
            Width           =   1000
         End
         Begin VB.TextBox textStart 
            Height          =   315
            Index           =   1
            Left            =   -72600
            TabIndex        =   22
            Top             =   1110
            Width           =   1000
         End
         Begin VB.TextBox textStart 
            Height          =   315
            Index           =   0
            Left            =   -72600
            TabIndex        =   19
            Top             =   750
            Width           =   1000
         End
         Begin VB.ComboBox comboDayStatus 
            Height          =   330
            Index           =   6
            ItemData        =   "FRMSET_A.frx":3D97
            Left            =   -73845
            List            =   "FRMSET_A.frx":3DA7
            Style           =   2  'Dropdown List
            TabIndex        =   36
            Top             =   2910
            Width           =   1155
         End
         Begin VB.ComboBox comboDayStatus 
            Height          =   330
            Index           =   5
            ItemData        =   "FRMSET_A.frx":3DC5
            Left            =   -73845
            List            =   "FRMSET_A.frx":3DD5
            Style           =   2  'Dropdown List
            TabIndex        =   33
            Top             =   2550
            Width           =   1155
         End
         Begin VB.ComboBox comboDayStatus 
            Height          =   330
            Index           =   4
            ItemData        =   "FRMSET_A.frx":3DF3
            Left            =   -73845
            List            =   "FRMSET_A.frx":3E03
            Style           =   2  'Dropdown List
            TabIndex        =   30
            Top             =   2190
            Width           =   1155
         End
         Begin VB.ComboBox comboDayStatus 
            Height          =   330
            Index           =   3
            ItemData        =   "FRMSET_A.frx":3E21
            Left            =   -73845
            List            =   "FRMSET_A.frx":3E31
            Style           =   2  'Dropdown List
            TabIndex        =   27
            Top             =   1830
            Width           =   1155
         End
         Begin VB.ComboBox comboDayStatus 
            Height          =   330
            Index           =   2
            ItemData        =   "FRMSET_A.frx":3E4F
            Left            =   -73845
            List            =   "FRMSET_A.frx":3E5F
            Style           =   2  'Dropdown List
            TabIndex        =   24
            Top             =   1470
            Width           =   1155
         End
         Begin VB.ComboBox comboDayStatus 
            Height          =   330
            Index           =   1
            ItemData        =   "FRMSET_A.frx":3E7D
            Left            =   -73845
            List            =   "FRMSET_A.frx":3E8D
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   1110
            Width           =   1155
         End
         Begin VB.ComboBox comboDayStatus 
            Height          =   330
            Index           =   0
            ItemData        =   "FRMSET_A.frx":3EAB
            Left            =   -73845
            List            =   "FRMSET_A.frx":3EBB
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Top             =   750
            Width           =   1155
         End
         Begin VB.TextBox txtHolDays 
            Alignment       =   2  'Center
            Height          =   315
            Left            =   -70935
            MaxLength       =   25
            TabIndex        =   64
            Top             =   750
            Width           =   500
         End
         Begin VB.OptionButton OptionPay 
            Caption         =   "Hourly"
            Height          =   255
            Index           =   0
            Left            =   1140
            TabIndex        =   15
            Top             =   1740
            Width           =   855
         End
         Begin VB.OptionButton OptionPay 
            Caption         =   "Weekly"
            Height          =   255
            Index           =   1
            Left            =   1140
            TabIndex        =   16
            Top             =   2040
            Width           =   855
         End
         Begin VB.CheckBox CheckHoliday 
            Alignment       =   1  'Right Justify
            Caption         =   "On/Off"
            Height          =   210
            Left            =   -74940
            TabIndex        =   61
            Top             =   840
            Width           =   855
         End
         Begin VB.CheckBox CheckClass 
            Caption         =   "Roster"
            Height          =   255
            Index           =   0
            Left            =   -74880
            TabIndex        =   41
            Top             =   810
            Width           =   2895
         End
         Begin VB.CheckBox CheckClass 
            Caption         =   "Roster"
            Height          =   255
            Index           =   1
            Left            =   -74880
            TabIndex        =   43
            Top             =   1140
            Width           =   2895
         End
         Begin VB.CheckBox CheckClass 
            Caption         =   "Roster"
            Height          =   255
            Index           =   2
            Left            =   -74880
            TabIndex        =   45
            Top             =   1470
            Width           =   2895
         End
         Begin VB.CheckBox CheckClass 
            Caption         =   "Roster"
            Height          =   255
            Index           =   3
            Left            =   -74880
            TabIndex        =   47
            Top             =   1800
            Width           =   2895
         End
         Begin VB.CheckBox CheckClass 
            Caption         =   "Roster"
            Height          =   255
            Index           =   4
            Left            =   -74880
            TabIndex        =   49
            Top             =   2130
            Width           =   2895
         End
         Begin VB.CheckBox CheckClass 
            Caption         =   "Roster"
            Height          =   255
            Index           =   5
            Left            =   -74880
            TabIndex        =   51
            Top             =   2460
            Width           =   2895
         End
         Begin VB.CheckBox CheckClass 
            Caption         =   "Roster"
            Height          =   255
            Index           =   6
            Left            =   -74880
            TabIndex        =   53
            Top             =   2790
            Width           =   2895
         End
         Begin VB.CheckBox CheckClass 
            Caption         =   "Roster"
            Height          =   255
            Index           =   7
            Left            =   -74880
            TabIndex        =   55
            Top             =   3120
            Width           =   2895
         End
         Begin VB.CheckBox CheckClass 
            Caption         =   "Roster"
            Height          =   255
            Index           =   8
            Left            =   -74880
            TabIndex        =   57
            Top             =   3450
            Width           =   2895
         End
         Begin VB.CheckBox CheckClass 
            Caption         =   "Roster"
            Height          =   255
            Index           =   9
            Left            =   -74880
            TabIndex        =   59
            Top             =   3780
            Width           =   2895
         End
         Begin VB.Label label_std 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000002&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "End Date"
            ForeColor       =   &H80000009&
            Height          =   255
            Index           =   41
            Left            =   -72480
            TabIndex        =   155
            Top             =   420
            Width           =   1470
         End
         Begin VB.Label label_std 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000002&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Start Date"
            ForeColor       =   &H80000009&
            Height          =   255
            Index           =   40
            Left            =   -74040
            TabIndex        =   154
            Top             =   420
            Width           =   1470
         End
         Begin VB.Label label_std 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000002&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Days"
            ForeColor       =   &H80000009&
            Height          =   255
            Index           =   39
            Left            =   -70935
            TabIndex        =   153
            Top             =   420
            Width           =   495
         End
         Begin VB.Label label_std 
            AutoSize        =   -1  'True
            BackColor       =   &H80000004&
            BackStyle       =   0  'Transparent
            Caption         =   $"FRMSET_A.frx":3ED9
            Height          =   2430
            Index           =   38
            Left            =   -74880
            TabIndex        =   152
            Top             =   1560
            Width           =   4440
            WordWrap        =   -1  'True
         End
         Begin VB.Label label_std 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000002&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Roster Availability"
            ForeColor       =   &H80000009&
            Height          =   255
            Index           =   37
            Left            =   -74880
            TabIndex        =   151
            Top             =   420
            Width           =   2895
         End
         Begin VB.Label label_std 
            AutoSize        =   -1  'True
            BackColor       =   &H80000004&
            BackStyle       =   0  'Transparent
            Caption         =   "Minimum && maximum hours per week required by this staff member."
            Height          =   690
            Index           =   36
            Left            =   -72720
            TabIndex        =   150
            Top             =   3660
            Width           =   2280
            WordWrap        =   -1  'True
         End
         Begin VB.Label label_std 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000D&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Hours Required per Week"
            ForeColor       =   &H8000000E&
            Height          =   255
            Index           =   35
            Left            =   -74880
            TabIndex        =   149
            Top             =   3360
            Width           =   4365
         End
         Begin VB.Label label_std 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000002&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Maximum"
            ForeColor       =   &H80000009&
            Height          =   255
            Index           =   34
            Left            =   -73800
            TabIndex        =   148
            Top             =   3660
            Width           =   1005
         End
         Begin VB.Label label_std 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000002&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Minimum"
            ForeColor       =   &H80000009&
            Height          =   255
            Index           =   33
            Left            =   -74880
            TabIndex        =   147
            Top             =   3660
            Width           =   1005
         End
         Begin VB.Label label_std 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000002&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Day"
            ForeColor       =   &H80000009&
            Height          =   255
            Index           =   32
            Left            =   -74880
            TabIndex        =   146
            Top             =   420
            Width           =   990
         End
         Begin VB.Label label_std 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000002&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Status"
            ForeColor       =   &H80000009&
            Height          =   255
            Index           =   31
            Left            =   -73830
            TabIndex        =   145
            Top             =   420
            Width           =   1155
         End
         Begin VB.Label label_std 
            AutoSize        =   -1  'True
            BackColor       =   &H80000004&
            BackStyle       =   0  'Transparent
            Caption         =   "Minimum && maximum hours per week required by this staff member."
            Height          =   630
            Index           =   11
            Left            =   -72720
            TabIndex        =   118
            Top             =   3540
            Width           =   2280
            WordWrap        =   -1  'True
         End
         Begin VB.Label label_std 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000002&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Birth Date"
            ForeColor       =   &H80000009&
            Height          =   255
            Index           =   4
            Left            =   60
            TabIndex        =   117
            Top             =   1050
            Width           =   1460
         End
         Begin VB.Label label_std 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000002&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Employed"
            ForeColor       =   &H80000009&
            Height          =   255
            Index           =   5
            Left            =   1580
            TabIndex        =   116
            Top             =   1050
            Width           =   1460
         End
         Begin VB.Label label_std 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000002&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Base Rate"
            ForeColor       =   &H80000009&
            Height          =   255
            Index           =   6
            Left            =   60
            TabIndex        =   115
            Top             =   1710
            Width           =   1065
         End
         Begin VB.Label label_std 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000002&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Home Phone"
            ForeColor       =   &H80000009&
            Height          =   255
            Index           =   7
            Left            =   3100
            TabIndex        =   114
            Top             =   1050
            Width           =   1460
         End
         Begin VB.Label label_std 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000D&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Staff ID"
            ForeColor       =   &H8000000E&
            Height          =   255
            Index           =   0
            Left            =   60
            TabIndex        =   113
            Top             =   390
            Width           =   1060
         End
         Begin VB.Label label_std 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000002&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Middle Name"
            ForeColor       =   &H80000009&
            Height          =   255
            Index           =   3
            Left            =   3360
            TabIndex        =   112
            Top             =   390
            Width           =   1200
         End
         Begin VB.Label label_std 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000002&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "First Name"
            ForeColor       =   &H80000009&
            Height          =   255
            Index           =   2
            Left            =   2260
            TabIndex        =   111
            Top             =   390
            Width           =   1060
         End
         Begin VB.Label label_std 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000002&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Last Name"
            ForeColor       =   &H80000009&
            Height          =   255
            Index           =   1
            Left            =   1160
            TabIndex        =   110
            Top             =   396
            Width           =   1060
         End
         Begin VB.Label label_std 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000D&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Hours Required per Week"
            ForeColor       =   &H8000000E&
            Height          =   255
            Index           =   10
            Left            =   -74880
            TabIndex        =   109
            Top             =   3240
            Width           =   4360
         End
         Begin MSMask.MaskEdBox MaskMaxHours 
            DataField       =   "HourRate"
            Height          =   315
            Left            =   -73800
            TabIndex        =   40
            Top             =   4005
            Width           =   1005
            _version        =   65536
            _extentx        =   1773
            _extenty        =   556
            _stockprops     =   109
            forecolor       =   -2147483640
            backcolor       =   -2147483643
            borderstyle     =   1
            autotab         =   -1  'True
            promptinclude   =   0   'False
            clipmode        =   1
            format          =   "#0.00"
         End
         Begin MSMask.MaskEdBox MaskMinHours 
            DataField       =   "HourRate"
            Height          =   315
            Left            =   -74880
            TabIndex        =   39
            Top             =   4005
            Width           =   1005
            _version        =   65536
            _extentx        =   1773
            _extenty        =   556
            _stockprops     =   109
            forecolor       =   -2147483640
            backcolor       =   -2147483643
            borderstyle     =   1
            autotab         =   -1  'True
            promptinclude   =   0   'False
            clipmode        =   1
            format          =   "#0.00"
         End
         Begin MSMask.MaskEdBox MaskHourRate 
            DataField       =   "HourRate"
            Height          =   315
            Left            =   60
            TabIndex        =   14
            Top             =   1995
            Width           =   1065
            _version        =   65536
            _extentx        =   1879
            _extenty        =   556
            _stockprops     =   109
            forecolor       =   -2147483640
            backcolor       =   -2147483643
            borderstyle     =   1
            autotab         =   -1  'True
            promptinclude   =   0   'False
            clipmode        =   1
            format          =   "Currency"
         End
         Begin MSMask.MaskEdBox MaskDateHired 
            DataField       =   "DateHired"
            Height          =   315
            Left            =   1580
            TabIndex        =   12
            Top             =   1335
            Width           =   1460
            _version        =   65536
            _extentx        =   2566
            _extenty        =   556
            _stockprops     =   109
            forecolor       =   -2147483640
            backcolor       =   -2147483643
            borderstyle     =   1
            autotab         =   -1  'True
            promptinclude   =   0   'False
            clipmode        =   1
            format          =   "Short Date"
         End
         Begin MSMask.MaskEdBox MaskBirthDate 
            DataField       =   "Birthdate"
            Height          =   315
            Left            =   60
            TabIndex        =   11
            Top             =   1330
            Width           =   1460
            _version        =   65536
            _extentx        =   2566
            _extenty        =   556
            _stockprops     =   109
            forecolor       =   -2147483640
            backcolor       =   -2147483643
            borderstyle     =   1
            autotab         =   -1  'True
            promptinclude   =   0   'False
            clipmode        =   1
            format          =   "Short Date"
         End
         Begin MSMask.MaskEdBox MaskHomePhone 
            DataField       =   "HomePhone"
            Height          =   315
            Left            =   3100
            TabIndex        =   13
            Top             =   1330
            Width           =   1460
            _version        =   65536
            _extentx        =   2566
            _extenty        =   556
            _stockprops     =   109
            forecolor       =   -2147483640
            backcolor       =   -2147483643
            borderstyle     =   1
            autotab         =   -1  'True
            promptinclude   =   0   'False
            clipmode        =   1
            maxlength       =   14
            mask            =   "(##9) #####999"
         End
         Begin MSMask.MaskEdBox maskRosterRate 
            DataField       =   "HourRate"
            Height          =   315
            Index           =   0
            Left            =   -71940
            TabIndex        =   42
            Top             =   750
            Visible         =   0   'False
            Width           =   1005
            _version        =   65536
            _extentx        =   1773
            _extenty        =   556
            _stockprops     =   109
            forecolor       =   -2147483640
            backcolor       =   -2147483643
            borderstyle     =   1
            autotab         =   -1  'True
            promptinclude   =   0   'False
            clipmode        =   1
            format          =   "Currency"
         End
         Begin MSMask.MaskEdBox maskRosterRate 
            DataField       =   "HourRate"
            Height          =   315
            Index           =   1
            Left            =   -71940
            TabIndex        =   44
            Top             =   1080
            Visible         =   0   'False
            Width           =   1005
            _version        =   65536
            _extentx        =   1773
            _extenty        =   556
            _stockprops     =   109
            forecolor       =   -2147483640
            backcolor       =   -2147483643
            borderstyle     =   1
            autotab         =   -1  'True
            promptinclude   =   0   'False
            clipmode        =   1
            format          =   "Currency"
         End
         Begin MSMask.MaskEdBox maskRosterRate 
            DataField       =   "HourRate"
            Height          =   315
            Index           =   2
            Left            =   -71940
            TabIndex        =   46
            Top             =   1410
            Visible         =   0   'False
            Width           =   1005
            _version        =   65536
            _extentx        =   1773
            _extenty        =   556
            _stockprops     =   109
            forecolor       =   -2147483640
            backcolor       =   -2147483643
            borderstyle     =   1
            autotab         =   -1  'True
            promptinclude   =   0   'False
            clipmode        =   1
            format          =   "Currency"
         End
         Begin MSMask.MaskEdBox maskRosterRate 
            DataField       =   "HourRate"
            Height          =   315
            Index           =   3
            Left            =   -71940
            TabIndex        =   48
            Top             =   1740
            Visible         =   0   'False
            Width           =   1005
            _version        =   65536
            _extentx        =   1773
            _extenty        =   556
            _stockprops     =   109
            forecolor       =   -2147483640
            backcolor       =   -2147483643
            borderstyle     =   1
            autotab         =   -1  'True
            promptinclude   =   0   'False
            clipmode        =   1
            format          =   "Currency"
         End
         Begin MSMask.MaskEdBox maskRosterRate 
            DataField       =   "HourRate"
            Height          =   315
            Index           =   4
            Left            =   -71940
            TabIndex        =   50
            Top             =   2070
            Visible         =   0   'False
            Width           =   1005
            _version        =   65536
            _extentx        =   1773
            _extenty        =   556
            _stockprops     =   109
            forecolor       =   -2147483640
            backcolor       =   -2147483643
            borderstyle     =   1
            autotab         =   -1  'True
            promptinclude   =   0   'False
            clipmode        =   1
            format          =   "Currency"
         End
         Begin MSMask.MaskEdBox maskRosterRate 
            DataField       =   "HourRate"
            Height          =   315
            Index           =   5
            Left            =   -71940
            TabIndex        =   52
            Top             =   2400
            Visible         =   0   'False
            Width           =   1005
            _version        =   65536
            _extentx        =   1773
            _extenty        =   556
            _stockprops     =   109
            forecolor       =   -2147483640
            backcolor       =   -2147483643
            borderstyle     =   1
            autotab         =   -1  'True
            promptinclude   =   0   'False
            clipmode        =   1
            format          =   "Currency"
         End
         Begin MSMask.MaskEdBox maskRosterRate 
            DataField       =   "HourRate"
            Height          =   315
            Index           =   6
            Left            =   -71940
            TabIndex        =   54
            Top             =   2730
            Visible         =   0   'False
            Width           =   1005
            _version        =   65536
            _extentx        =   1773
            _extenty        =   556
            _stockprops     =   109
            forecolor       =   -2147483640
            backcolor       =   -2147483643
            borderstyle     =   1
            autotab         =   -1  'True
            promptinclude   =   0   'False
            clipmode        =   1
            format          =   "Currency"
         End
         Begin MSMask.MaskEdBox maskRosterRate 
            DataField       =   "HourRate"
            Height          =   315
            Index           =   7
            Left            =   -71940
            TabIndex        =   56
            Top             =   3060
            Visible         =   0   'False
            Width           =   1005
            _version        =   65536
            _extentx        =   1773
            _extenty        =   556
            _stockprops     =   109
            forecolor       =   -2147483640
            backcolor       =   -2147483643
            borderstyle     =   1
            autotab         =   -1  'True
            promptinclude   =   0   'False
            clipmode        =   1
            format          =   "Currency"
         End
         Begin MSMask.MaskEdBox maskRosterRate 
            DataField       =   "HourRate"
            Height          =   315
            Index           =   8
            Left            =   -71940
            TabIndex        =   58
            Top             =   3390
            Visible         =   0   'False
            Width           =   1005
            _version        =   65536
            _extentx        =   1773
            _extenty        =   556
            _stockprops     =   109
            forecolor       =   -2147483640
            backcolor       =   -2147483643
            borderstyle     =   1
            autotab         =   -1  'True
            promptinclude   =   0   'False
            clipmode        =   1
            format          =   "Currency"
         End
         Begin MSMask.MaskEdBox maskRosterRate 
            DataField       =   "HourRate"
            Height          =   315
            Index           =   9
            Left            =   -71940
            TabIndex        =   60
            Top             =   3720
            Visible         =   0   'False
            Width           =   1005
            _version        =   65536
            _extentx        =   1773
            _extenty        =   556
            _stockprops     =   109
            forecolor       =   -2147483640
            backcolor       =   -2147483643
            borderstyle     =   1
            autotab         =   -1  'True
            promptinclude   =   0   'False
            clipmode        =   1
            format          =   "Currency"
         End
         Begin VB.Label label_std 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000002&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Roster Availability"
            ForeColor       =   &H80000009&
            Height          =   255
            Index           =   14
            Left            =   -74880
            TabIndex        =   108
            Top             =   420
            Width           =   2900
         End
         Begin VB.Label lbl_hr_rate 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000002&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Hourly Rate"
            ForeColor       =   &H80000009&
            Height          =   255
            Left            =   -71940
            TabIndex        =   107
            Top             =   420
            Width           =   1335
         End
         Begin VB.Label label_std 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000002&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Maximum"
            ForeColor       =   &H80000009&
            Height          =   255
            Index           =   13
            Left            =   -73800
            TabIndex        =   106
            Top             =   3540
            Width           =   1005
         End
         Begin VB.Label label_std 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000002&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Minimum"
            ForeColor       =   &H80000009&
            Height          =   255
            Index           =   12
            Left            =   -74880
            TabIndex        =   105
            Top             =   3540
            Width           =   1005
         End
         Begin VB.Label label_std 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000002&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Day"
            ForeColor       =   &H80000009&
            Height          =   255
            Index           =   8
            Left            =   -74880
            TabIndex        =   104
            Top             =   420
            Width           =   996
         End
         Begin VB.Label lbl_std 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000002&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Finish Time"
            ForeColor       =   &H80000009&
            Height          =   255
            Index           =   2
            Left            =   -71532
            TabIndex        =   103
            Top             =   420
            Visible         =   0   'False
            Width           =   1000
         End
         Begin VB.Label lbl_std 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000002&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Start Time"
            ForeColor       =   &H80000009&
            Height          =   255
            Index           =   1
            Left            =   -72600
            TabIndex        =   102
            Top             =   420
            Visible         =   0   'False
            Width           =   1000
         End
         Begin VB.Label label_std 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000002&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Status"
            ForeColor       =   &H80000009&
            Height          =   255
            Index           =   9
            Left            =   -73824
            TabIndex        =   101
            Top             =   420
            Width           =   1155
         End
         Begin VB.Label labelDay 
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "Day Name"
            Height          =   210
            Index           =   6
            Left            =   -74880
            TabIndex        =   100
            Top             =   2985
            Width           =   990
         End
         Begin VB.Label labelDay 
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "Day Name"
            Height          =   210
            Index           =   5
            Left            =   -74880
            TabIndex        =   99
            Top             =   2625
            Width           =   990
         End
         Begin VB.Label labelDay 
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "Day Name"
            Height          =   210
            Index           =   4
            Left            =   -74880
            TabIndex        =   98
            Top             =   2265
            Width           =   990
         End
         Begin VB.Label labelDay 
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "Day Name"
            Height          =   210
            Index           =   3
            Left            =   -74880
            TabIndex        =   97
            Top             =   1905
            Width           =   990
         End
         Begin VB.Label labelDay 
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "Day Name"
            Height          =   210
            Index           =   2
            Left            =   -74880
            TabIndex        =   96
            Top             =   1545
            Width           =   990
         End
         Begin VB.Label labelDay 
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "Day Name"
            Height          =   210
            Index           =   1
            Left            =   -74880
            TabIndex        =   95
            Top             =   1200
            Width           =   990
         End
         Begin VB.Label labelDay 
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "Day Name"
            Height          =   210
            Index           =   0
            Left            =   -74880
            TabIndex        =   94
            Top             =   840
            Width           =   990
         End
         Begin MSMask.MaskEdBox MaskHolStart 
            DataField       =   "DateHired"
            Height          =   315
            Left            =   -74040
            TabIndex        =   62
            Top             =   750
            Width           =   1470
            _version        =   65536
            _extentx        =   2593
            _extenty        =   556
            _stockprops     =   109
            forecolor       =   0
            backcolor       =   16777215
            borderstyle     =   1
            enabled         =   0   'False
            autotab         =   -1  'True
            promptinclude   =   0   'False
            clipmode        =   1
            format          =   "Short Date"
         End
         Begin MSMask.MaskEdBox MaskHolEnd 
            DataField       =   "DateHired"
            Height          =   315
            Left            =   -72480
            TabIndex        =   63
            Top             =   750
            Width           =   1470
            _version        =   65536
            _extentx        =   2593
            _extenty        =   556
            _stockprops     =   109
            forecolor       =   0
            backcolor       =   16777215
            borderstyle     =   1
            enabled         =   0   'False
            autotab         =   -1  'True
            promptinclude   =   0   'False
            clipmode        =   1
            format          =   "Short Date"
         End
         Begin VB.Label label_std 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000002&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "End Date"
            ForeColor       =   &H80000009&
            Height          =   255
            Index           =   16
            Left            =   -72480
            TabIndex        =   93
            Top             =   420
            Width           =   1470
         End
         Begin VB.Label label_std 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000002&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Start Date"
            ForeColor       =   &H80000009&
            Height          =   255
            Index           =   15
            Left            =   -74040
            TabIndex        =   92
            Top             =   420
            Width           =   1470
         End
         Begin VB.Label label_std 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000002&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Days"
            ForeColor       =   &H80000009&
            Height          =   255
            Index           =   17
            Left            =   -70935
            TabIndex        =   91
            Top             =   420
            Width           =   495
         End
         Begin VB.Image cmdHolStartToday 
            Height          =   300
            Left            =   -73440
            Picture         =   "FRMSET_A.frx":3F63
            Top             =   1140
            Width           =   300
         End
         Begin VB.Image cmdHolEndToday 
            Height          =   300
            Left            =   -71880
            Picture         =   "FRMSET_A.frx":4537
            Top             =   1140
            Width           =   300
         End
         Begin VB.Label label_std 
            AutoSize        =   -1  'True
            BackColor       =   &H80000004&
            BackStyle       =   0  'Transparent
            Caption         =   $"FRMSET_A.frx":4B0B
            Height          =   2550
            Index           =   20
            Left            =   -74880
            TabIndex        =   90
            Top             =   1440
            Width           =   4440
            WordWrap        =   -1  'True
         End
         Begin VB.Image cmdHolStartPlus 
            Height          =   300
            Left            =   -73140
            Picture         =   "FRMSET_A.frx":4B95
            Top             =   1140
            Width           =   300
         End
         Begin VB.Image cmdHolStartMinus 
            Height          =   300
            Left            =   -73740
            Picture         =   "FRMSET_A.frx":5169
            Top             =   1140
            Width           =   300
         End
         Begin VB.Image cmdholEndPlus 
            Height          =   300
            Left            =   -71580
            Picture         =   "FRMSET_A.frx":573D
            Top             =   1140
            Width           =   300
         End
         Begin VB.Image cmdHolEndMinus 
            Height          =   300
            Left            =   -72180
            Picture         =   "FRMSET_A.frx":5D11
            Top             =   1140
            Width           =   300
         End
         Begin VB.Image cmdStartWeekPlus 
            Height          =   300
            Left            =   -72840
            Picture         =   "FRMSET_A.frx":62E5
            Top             =   1140
            Width           =   300
         End
         Begin VB.Image cmdStartWeekMinus 
            Height          =   300
            Left            =   -74040
            Picture         =   "FRMSET_A.frx":68B9
            Top             =   1140
            Width           =   300
         End
         Begin VB.Image cmdendWeekPlus 
            Height          =   300
            Left            =   -71280
            Picture         =   "FRMSET_A.frx":6E8D
            Top             =   1140
            Width           =   300
         End
         Begin VB.Image cmdEndWeekMinus 
            Height          =   300
            Left            =   -72480
            Picture         =   "FRMSET_A.frx":7461
            Top             =   1140
            Width           =   300
         End
         Begin VB.Label labelInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            ForeColor       =   &H80000008&
            Height          =   1035
            Left            =   2100
            TabIndex        =   89
            Top             =   1800
            Width           =   2415
            WordWrap        =   -1  'True
         End
         Begin VB.Shape ShapeBorder 
            BackColor       =   &H80000002&
            BorderColor     =   &H80000002&
            BorderWidth     =   2
            Height          =   1155
            Left            =   2040
            Top             =   1740
            Width           =   2535
         End
         Begin VB.Image cmdBaseRate 
            Height          =   300
            Index           =   0
            Left            =   -70920
            Picture         =   "FRMSET_A.frx":7A35
            Top             =   765
            Visible         =   0   'False
            Width           =   300
         End
         Begin VB.Image cmdBaseRate 
            Height          =   300
            Index           =   1
            Left            =   -70920
            Picture         =   "FRMSET_A.frx":8009
            Top             =   1095
            Visible         =   0   'False
            Width           =   300
         End
         Begin VB.Image cmdBaseRate 
            Height          =   300
            Index           =   2
            Left            =   -70920
            Picture         =   "FRMSET_A.frx":85DD
            Top             =   1425
            Visible         =   0   'False
            Width           =   300
         End
         Begin VB.Image cmdBaseRate 
            Height          =   300
            Index           =   3
            Left            =   -70920
            Picture         =   "FRMSET_A.frx":8BB1
            Top             =   1755
            Visible         =   0   'False
            Width           =   300
         End
         Begin VB.Image cmdBaseRate 
            Height          =   300
            Index           =   4
            Left            =   -70920
            Picture         =   "FRMSET_A.frx":9185
            Top             =   2085
            Visible         =   0   'False
            Width           =   300
         End
         Begin VB.Image cmdBaseRate 
            Height          =   300
            Index           =   5
            Left            =   -70920
            Picture         =   "FRMSET_A.frx":9759
            Top             =   2415
            Visible         =   0   'False
            Width           =   300
         End
         Begin VB.Image cmdBaseRate 
            Height          =   300
            Index           =   6
            Left            =   -70920
            Picture         =   "FRMSET_A.frx":9D2D
            Top             =   2745
            Visible         =   0   'False
            Width           =   300
         End
         Begin VB.Image cmdBaseRate 
            Height          =   300
            Index           =   7
            Left            =   -70920
            Picture         =   "FRMSET_A.frx":A301
            Top             =   3075
            Visible         =   0   'False
            Width           =   300
         End
         Begin VB.Image cmdBaseRate 
            Height          =   300
            Index           =   8
            Left            =   -70920
            Picture         =   "FRMSET_A.frx":A8D5
            Top             =   3405
            Visible         =   0   'False
            Width           =   300
         End
         Begin VB.Image cmdBaseRate 
            Height          =   300
            Index           =   9
            Left            =   -70920
            Picture         =   "FRMSET_A.frx":AEA9
            Top             =   3735
            Visible         =   0   'False
            Width           =   300
         End
      End
      Begin VB.Image cmdDelete 
         Height          =   300
         Left            =   360
         Picture         =   "FRMSET_A.frx":B47D
         Top             =   180
         Width           =   300
      End
      Begin VB.Image cmdSave 
         Height          =   300
         Left            =   60
         Picture         =   "FRMSET_A.frx":BA51
         Top             =   180
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Image cmdNew 
         Height          =   300
         Left            =   660
         Picture         =   "FRMSET_A.frx":C025
         Top             =   180
         Width           =   300
      End
      Begin VB.Image cmdSelectedStaffRoster 
         Height          =   300
         Left            =   960
         Picture         =   "FRMSET_A.frx":C5F9
         Top             =   180
         Width           =   300
      End
      Begin VB.Image cmdAllStaffRosters 
         Height          =   300
         Left            =   1260
         Picture         =   "FRMSET_A.frx":CBCD
         Top             =   180
         Width           =   300
      End
   End
   Begin VB.Frame frameRosters 
      Caption         =   "Roster Definitions"
      Height          =   4875
      Left            =   600
      TabIndex        =   119
      Top             =   0
      Width           =   6795
      Begin MSGrid.Grid GridClass 
         Height          =   4515
         Left            =   120
         TabIndex        =   65
         Top             =   240
         Width           =   6555
         _version        =   65536
         _extentx        =   11562
         _extenty        =   7964
         _stockprops     =   77
         forecolor       =   -2147483640
         backcolor       =   -2147483643
         cols            =   4
         fixedcols       =   0
         scrollbars      =   0
         mousepointer    =   14
         mouseicon       =   "FRMSET_A.frx":D1A1
      End
   End
   Begin VB.Frame frameProgSet 
      Caption         =   "Program Settings"
      Height          =   4875
      Left            =   600
      TabIndex        =   135
      Top             =   0
      Width           =   6795
      Begin VB.Frame FrameStandard 
         Height          =   855
         Index           =   5
         Left            =   120
         TabIndex        =   142
         Top             =   3840
         Width           =   6555
         Begin Threed.SSCommand cmdNukeCover 
            Height          =   600
            Left            =   5880
            TabIndex        =   73
            Top             =   180
            Width           =   600
            _version        =   65536
            _extentx        =   1058
            _extenty        =   1058
            _stockprops     =   78
            autosize        =   2
            picture         =   "FRMSET_A.frx":D1BD
         End
         Begin VB.Label label_std 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Delete all roster shifts from the GSR database"
            ForeColor       =   &H80000009&
            Height          =   240
            Index           =   30
            Left            =   60
            TabIndex        =   144
            Top             =   210
            Width           =   5730
            WordWrap        =   -1  'True
         End
         Begin VB.Label label_std 
            AutoSize        =   -1  'True
            BackColor       =   &H80000004&
            BackStyle       =   0  'Transparent
            Caption         =   "Click to expose NUKE command button."
            Height          =   210
            Index           =   29
            Left            =   60
            TabIndex        =   143
            Top             =   540
            Width           =   5715
            WordWrap        =   -1  'True
         End
         Begin Threed.SSCommand cmdNuke 
            Height          =   600
            Left            =   5880
            TabIndex        =   74
            Top             =   180
            Width           =   600
            _version        =   65536
            _extentx        =   1058
            _extenty        =   1058
            _stockprops     =   78
            autosize        =   2
            picture         =   "FRMSET_A.frx":DA97
         End
      End
      Begin VB.Frame frameDateFormat 
         Height          =   1695
         Left            =   120
         TabIndex        =   139
         Top             =   1440
         Width           =   6555
         Begin VB.OptionButton OptionDateFormat 
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   69
            Top             =   960
            Width           =   1995
         End
         Begin VB.OptionButton OptionDateFormat 
            Height          =   255
            Index           =   1
            Left            =   2196
            TabIndex        =   70
            Top             =   960
            Width           =   2055
         End
         Begin VB.OptionButton OptionDateFormat 
            Caption         =   "Custom Format"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   71
            Top             =   1290
            Width           =   1995
         End
         Begin MSMask.MaskEdBox maskDateMask 
            Height          =   300
            Left            =   2196
            TabIndex        =   72
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
         Begin VB.Label label_std 
            AutoSize        =   -1  'True
            BackColor       =   &H80000004&
            BackStyle       =   0  'Transparent
            Caption         =   $"FRMSET_A.frx":DDB1
            Height          =   450
            Index           =   28
            Left            =   90
            TabIndex        =   141
            Top             =   480
            Width           =   4215
            WordWrap        =   -1  'True
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
            TabIndex        =   140
            Top             =   180
            Width           =   4275
         End
      End
      Begin VB.Frame FrameStandard 
         Height          =   1215
         Index           =   0
         Left            =   120
         TabIndex        =   136
         Top             =   180
         Width           =   6555
         Begin VB.CheckBox CheckSounds 
            Alignment       =   1  'Right Justify
            Caption         =   "Sounds"
            Height          =   210
            Left            =   2220
            TabIndex        =   66
            Top             =   480
            Width           =   2100
         End
         Begin VB.CheckBox CheckDelete 
            Alignment       =   1  'Right Justify
            Caption         =   "Confirm Deletes"
            Height          =   210
            Left            =   2220
            TabIndex        =   67
            Top             =   720
            Width           =   2100
         End
         Begin VB.CheckBox CheckAllShifts 
            Alignment       =   1  'Right Justify
            Caption         =   "Require All Shifts Filled"
            Height          =   210
            Left            =   2220
            TabIndex        =   68
            Top             =   960
            Width           =   2100
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
            TabIndex        =   138
            Top             =   180
            Width           =   4275
         End
         Begin VB.Label label_std 
            AutoSize        =   -1  'True
            BackColor       =   &H80000004&
            BackStyle       =   0  'Transparent
            Caption         =   "These settings affect the way GSR responds to your actions."
            Height          =   675
            Index           =   27
            Left            =   90
            TabIndex        =   137
            Top             =   480
            Width           =   2085
            WordWrap        =   -1  'True
         End
      End
   End
   Begin VB.Frame frameRosSet 
      Caption         =   "Roster Settings"
      Height          =   4875
      Left            =   600
      TabIndex        =   120
      Top             =   0
      Width           =   6795
      Begin VB.Frame FrameStandard 
         Height          =   1875
         Index           =   2
         Left            =   120
         TabIndex        =   129
         Top             =   2880
         Width           =   6555
         Begin VB.ComboBox ComboBlockMin 
            Height          =   330
            ItemData        =   "FRMSET_A.frx":DE07
            Left            =   3570
            List            =   "FRMSET_A.frx":DE17
            TabIndex        =   82
            Text            =   "ComboMinute"
            Top             =   420
            Width           =   765
         End
         Begin VB.ComboBox ComboBlockHour 
            Height          =   330
            ItemData        =   "FRMSET_A.frx":DE2B
            Left            =   2760
            List            =   "FRMSET_A.frx":DE85
            TabIndex        =   81
            Text            =   "ComboHour"
            Top             =   420
            Width           =   765
         End
         Begin VB.ComboBox ComboBreakMin 
            Height          =   330
            ItemData        =   "FRMSET_A.frx":DEE9
            Left            =   3570
            List            =   "FRMSET_A.frx":DEF9
            TabIndex        =   84
            Text            =   "ComboMinute"
            Top             =   780
            Width           =   765
         End
         Begin VB.ComboBox ComboBreakHour 
            Height          =   330
            ItemData        =   "FRMSET_A.frx":DF0D
            Left            =   2760
            List            =   "FRMSET_A.frx":DF67
            TabIndex        =   83
            Text            =   "ComboHour"
            Top             =   780
            Width           =   765
         End
         Begin VB.Label label_std 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000002&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Hours"
            ForeColor       =   &H80000009&
            Height          =   255
            Index           =   26
            Left            =   2760
            TabIndex        =   134
            Top             =   150
            Width           =   765
         End
         Begin VB.Label label_std 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000002&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Minutes"
            ForeColor       =   &H80000009&
            Height          =   255
            Index           =   25
            Left            =   3540
            TabIndex        =   133
            Top             =   150
            Width           =   765
         End
         Begin VB.Label label_std 
            AutoSize        =   -1  'True
            BackColor       =   &H80000004&
            BackStyle       =   0  'Transparent
            Caption         =   "Length of break."
            Height          =   210
            Index           =   24
            Left            =   60
            TabIndex        =   132
            Top             =   840
            Width           =   1995
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
            TabIndex        =   131
            Top             =   150
            Width           =   2655
         End
         Begin VB.Label label_std 
            AutoSize        =   -1  'True
            BackColor       =   &H80000004&
            BackStyle       =   0  'Transparent
            Caption         =   "Time before break is applied."
            Height          =   210
            Index           =   23
            Left            =   60
            TabIndex        =   130
            Top             =   480
            Width           =   2205
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Frame FrameStandard 
         Height          =   1995
         Index           =   1
         Left            =   120
         TabIndex        =   123
         Top             =   840
         Width           =   6555
         Begin VB.ComboBox ComboIncrement 
            Height          =   330
            ItemData        =   "FRMSET_A.frx":DFCB
            Left            =   2760
            List            =   "FRMSET_A.frx":DFFE
            Style           =   2  'Dropdown List
            TabIndex        =   80
            Top             =   1530
            Width           =   1545
         End
         Begin VB.ComboBox ComboHour 
            Height          =   330
            ItemData        =   "FRMSET_A.frx":E080
            Left            =   2760
            List            =   "FRMSET_A.frx":E0DA
            TabIndex        =   78
            Text            =   "ComboHour"
            Top             =   750
            Width           =   765
         End
         Begin VB.ComboBox ComboMinute 
            Height          =   330
            ItemData        =   "FRMSET_A.frx":E13E
            Left            =   3570
            List            =   "FRMSET_A.frx":E14E
            TabIndex        =   79
            Text            =   "ComboMinute"
            Top             =   750
            Width           =   765
         End
         Begin VB.OptionButton OptionTime 
            Caption         =   "Start"
            Height          =   210
            Index           =   0
            Left            =   2760
            TabIndex        =   76
            Top             =   480
            Value           =   -1  'True
            Width           =   795
         End
         Begin VB.OptionButton OptionTime 
            Caption         =   "End"
            Height          =   210
            Index           =   1
            Left            =   3540
            TabIndex        =   77
            Top             =   480
            Width           =   735
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
            TabIndex        =   128
            Top             =   150
            Width           =   4275
         End
         Begin VB.Label label_std 
            AutoSize        =   -1  'True
            BackColor       =   &H80000004&
            BackStyle       =   0  'Transparent
            Caption         =   "Roster increment (or shift length)"
            Height          =   210
            Index           =   22
            Left            =   60
            TabIndex        =   127
            Top             =   1590
            Width           =   2700
            WordWrap        =   -1  'True
         End
         Begin VB.Label label_std 
            AutoSize        =   -1  'True
            BackColor       =   &H80000004&
            BackStyle       =   0  'Transparent
            Caption         =   "Length of working day (hh:mm)"
            Height          =   210
            Index           =   21
            Left            =   60
            TabIndex        =   126
            Top             =   1230
            Width           =   2700
            WordWrap        =   -1  'True
         End
         Begin MSMask.MaskEdBox MaskDayLength 
            Height          =   330
            Left            =   2760
            TabIndex        =   125
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
         Begin VB.Label label_std 
            AutoSize        =   -1  'True
            BackColor       =   &H80000004&
            BackStyle       =   0  'Transparent
            Caption         =   $"FRMSET_A.frx":E162
            Height          =   630
            Index           =   19
            Left            =   60
            TabIndex        =   124
            Top             =   480
            Width           =   2730
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Frame FrameStandard 
         Height          =   615
         Index           =   3
         Left            =   120
         TabIndex        =   121
         Top             =   180
         Width           =   6555
         Begin VB.ComboBox ComboStartDay 
            Height          =   330
            ItemData        =   "FRMSET_A.frx":E1BE
            Left            =   2760
            List            =   "FRMSET_A.frx":E1D7
            Style           =   2  'Dropdown List
            TabIndex        =   75
            Top             =   180
            Width           =   1575
         End
         Begin VB.Label label_std 
            Appearance      =   0  'Flat
            BackColor       =   &H80000002&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Starting Day of Roster Week"
            ForeColor       =   &H80000009&
            Height          =   255
            Index           =   18
            Left            =   60
            TabIndex        =   122
            Top             =   210
            Width           =   2655
            WordWrap        =   -1  'True
         End
      End
   End
   Begin VB.Image ImageSwitch 
      Height          =   240
      Index           =   0
      Left            =   0
      Picture         =   "FRMSET_A.frx":E21B
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image ImageSwitch 
      Height          =   240
      Index           =   1
      Left            =   240
      Picture         =   "FRMSET_A.frx":E31D
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image ImageSwitch 
      Height          =   240
      Index           =   2
      Left            =   450
      Picture         =   "FRMSET_A.frx":E41F
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
End
Attribute VB_Name = "frmSet"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Option Explicit
'[----------------------------------------------]
'[frmSet             GSR Control Panel Form     ]
'[----------------------------------------------]
'[GSR                                           ]
'[generic staff roster system for Windows       ]
'[(C) ---------- David Gilbert, 1997            ]
'[----------------------------------------------]
'[REV: 3.00.32                                  ]
'[----------------------------------------------]

Dim Shared intStartDay As Integer       '[Start time - week day]
Dim Shared intFinishDay As Integer      '[Finish time - week day]

Sub StatusBar(strMessage)

    '[PLACE PASSED MESSAGE ON THE STATUS BAR]
    If frmSet.panelStatusBar.Caption <> " " + strMessage Then frmSet.panelStatusBar.Caption = " " + strMessage

End Sub


Private Sub cmdPrintSet_Click()

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

Private Sub cmdPrintSet_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Change printer settings."
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub cmdProgSet_Click()
    
    '[MOVE PROGRAM SETTINGS FORM TO THE FRONT]
    frmSet.cmdStaff.Outline = False
    frmSet.cmdRosters.Outline = False
    frmSet.cmdRosSet.Outline = False
    frmSet.cmdProgSet.Outline = True
    
    frameProgSet.ZOrder 0

End Sub

Private Sub cmdProgSet_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "View the program settings section."
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub cmdRosSet_Click()
        
    '[MOVE ROSTER SETTINGS FORM TO THE FRONT]
    frmSet.cmdStaff.Outline = False
    frmSet.cmdRosters.Outline = False
    frmSet.cmdRosSet.Outline = True
    frmSet.cmdProgSet.Outline = False
    
    frameRosSet.ZOrder 0

End Sub

Private Sub cmdRosSet_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "View the roster settings section."
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub cmdRosters_Click()
    
    '[MOVE ROSTER FORM TO THE FRONT]
    frmSet.cmdStaff.Outline = False
    frmSet.cmdRosters.Outline = True
    frmSet.cmdRosSet.Outline = False
    frmSet.cmdProgSet.Outline = False
    
    frameRosters.ZOrder 0
    
End Sub

Private Sub cmdRosters_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "View the rosters detail section."
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub cmdSetClose_Click()

    '[HIDE CONTROL FORM]
    frmSet.Hide

End Sub

Private Sub cmdSetClose_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Close this form and return to GSR."
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub cmdStaff_Click()

    '[MOVE STAFF FORM TO THE FRONT]
    frmSet.cmdStaff.Outline = True
    frmSet.cmdRosters.Outline = False
    frmSet.cmdRosSet.Outline = False
    frmSet.cmdProgSet.Outline = False
    
    frameStaff.ZOrder 0

End Sub

Private Sub cmdStaff_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "View the staff details section."
    '[---------------------------------------------------------------------------------]

End Sub







Private Sub CheckClass_Click(Index As Integer)

    '[MAKE SAVE BUTTON VISIBLE]
    If frmSet.cmdSave.Visible = False Then frmSet.cmdSave.Visible = True
    
    '[MAKE RATE VISIBLE IF HOURLY PAY RATE IS SET AND CLASS IS CHECKED]
    If frmSet.CheckClass(Index).Value = vbChecked And frmSet.optionPay(vbHourly).Value = True Then
        frmSet.maskRosterRate(Index).Visible = True
        frmSet.cmdBaseRate(Index).Visible = True
        frmSet.lbl_hr_Rate.Visible = True
    Else
        frmSet.maskRosterRate(Index).Visible = False
        frmSet.cmdBaseRate(Index).Visible = False
    End If
    

End Sub

Private Sub CheckClass_GotFocus(Index As Integer)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Use the space bar to toggle the availability of this staff member to this roster."
    '[---------------------------------------------------------------------------------]

End Sub

Private Sub CheckClass_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Select which roster this staff member is available to."
    '[---------------------------------------------------------------------------------]

End Sub





Private Sub CheckHoliday_Click()

    '[REV: 3.00.27]
    '[CHANGE VALUE / ENABLE/DISABLE DATE BOXES]
    '[ENABLE SAVE BUTTON IF A CHANGE HAS BEEN MADE]
    If frmSet.cmdSave.Visible = False Then frmSet.cmdSave.Visible = True
    
    '[ENABLE/DISABLE DATES]
    If frmSet.CheckHoliday.Value = vbChecked Then
        frmSet.MaskHolStart.Enabled = True
        frmSet.MaskHolEnd.Enabled = True
        If IsDate(frmSet.MaskHolStart.Text) And IsDate(frmSet.MaskHolEnd.Text) Then frmSet.txtHolDays = (Format(CDate(frmSet.MaskHolEnd) - CDate(frmSet.MaskHolStart), "###0")) + 1
    Else
        frmSet.MaskHolStart.Enabled = False
        frmSet.MaskHolEnd.Enabled = False
        frmSet.txtHolDays = ""
    End If

End Sub

Private Sub CheckHoliday_GotFocus()
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Activate or deactivate holiday checking for this staff member."
    '[---------------------------------------------------------------------------------]

End Sub

Private Sub CheckHoliday_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Activate or deactivate holiday checking for this staff member."
    '[---------------------------------------------------------------------------------]

End Sub


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




Private Sub cmdBaseRate_Click(Index As Integer)

    '[PLACE CURRENT BASE HOURLY RATE INTO HOURLY RATE MASK BOX]
    frmSet.maskRosterRate(Index).Text = frmSet.MaskHourRate.Text

    '[ENABLE SAVE BUTTON IF A CHANGE HAS BEEN MADE]
    If frmSet.cmdSave.Visible = False Then frmSet.cmdSave.Visible = True

End Sub

Private Sub cmdBaseRate_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Set the hourly rate for this roster to the base rate."
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub cmdDelete_Click()

    '[ANIMATE BUTTON]
    cmdDelete.BorderStyle = 1
    Delay vbDelay
    cmdDelete.BorderStyle = 0

    '[*********************************************************************]
    '[DON'T GET TOO COMPLICATED IN HERE, OTHERWISE OTHER EVENTS MAY TRIGGER]
    '[*********************************************************************]
    
    '[DELETE A RECORD FROM THE STAFF LIST]
    Dim strDisplayname      As String
    Dim intColCounter          As Integer
    
        '[POPUP YES/NO DIALOG TO CONFIRM DELETION]
        Dim Msg As String
        Dim Style
        Dim Response
        Dim Title
        
        If frmSet.ListStaff.ListIndex = -1 Then Exit Sub
        
        '[CHECK DELETION FLAG]
        If flagDeleteConfirm Then
            Msg = "You have chosen to delete this staff member - " & DsStaff("LastName") & ", " & DsStaff("FirstName") & "." & strBreak & strBreak & "This will permanently remove this staff member from the list." & strBreak & strBreak & "Do you wish to continue ?"
            Style = vbYesNo ' Define buttons.
            Title = "Confirmation Required"  ' Define title.
            Response = gsrMsg(Msg, Style, Title)
        Else
            Response = vbYes
        End If
        
        If Response = vbYes Then    ' User chose Yes.
            
            '[DELETE STAFF MEMBER FROM DYNASET]
            DsStaff.Delete
            
            '[SAVE DISPLAYED NAME TO TEMPORARY STRING]
            If frmSet.ListStaff.ListCount > 1 Then
                strDisplayname = frmSet.ListStaff.List(frmSet.ListStaff.ListIndex - 1)
            Else
                strDisplayname = ""
            End If
            
            '[FILL STAFF LIST SO WE GET ORDER]
            FillStaffList
                
            '[RELOCATE STAFF NAME]
            For intColCounter = 0 To (frmSet.ListStaff.ListCount - 1)
                If frmSet.ListStaff.List(intColCounter) = strDisplayname Then
                    frmSet.ListStaff.ListIndex = intColCounter
                    Exit For
                End If
            Next intColCounter
        
            '[CHECK FOR NO STAFF MEMBERS]
            If DsStaff.RecordCount = 0 Then
                '[CALL ROUTINE TO ADD NEW STAFF MEMBER AND REPOSITION LIST]
                AddNewStaff
            End If
        
            '[MOVE TO FIRST ITEM IF NO ITEM IS SELECTED]
            If frmSet.ListStaff.ListIndex < 0 Then frmSet.ListStaff.ListIndex = 0
        
        End If
    

End Sub

Private Sub cmdDelete_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Delete the currently selected staff member (" & frmSet.ListStaff.List(frmSet.ListStaff.ListIndex) & ") from the list."
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub cmdEndWeekMinus_Click()
    
    '[ANIMATE BUTTON]
    frmSet.cmdEndWeekMinus.BorderStyle = 1
    Delay vbDelay
    frmSet.cmdEndWeekMinus.BorderStyle = 0
    
    '[CHANGE HOLEND DATE]
    If IsDate(frmSet.MaskHolEnd) Then frmSet.MaskHolEnd = Format(CDate(frmSet.MaskHolEnd) - 7, strDateFormat)

End Sub

Private Sub cmdEndWeekMinus_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Decrease the 'Holiday End Date' by 7 days."
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub cmdendWeekPlus_Click()
    
    '[ANIMATE BUTTON]
    frmSet.cmdendWeekPlus.BorderStyle = 1
    Delay vbDelay
    frmSet.cmdendWeekPlus.BorderStyle = 0
    
    '[CHANGE HOLEND DATE]
    If IsDate(frmSet.MaskHolEnd) Then frmSet.MaskHolEnd = Format(CDate(frmSet.MaskHolEnd) + 7, strDateFormat)

End Sub

Private Sub cmdendWeekPlus_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Increase the 'Holiday End Date' by 7 days."
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub cmdHolEndMinus_Click()
    
    '[ANIMATE BUTTON]
    frmSet.cmdHolEndMinus.BorderStyle = 1
    Delay vbDelay
    frmSet.cmdHolEndMinus.BorderStyle = 0
    
    '[CHANGE HOLEND DATE]
    If IsDate(frmSet.MaskHolEnd) Then frmSet.MaskHolEnd = Format(CDate(frmSet.MaskHolEnd) - 1, strDateFormat)

End Sub

Private Sub cmdHolEndMinus_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Decrease the 'Holiday End Date' by 1 day."
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub cmdholEndPlus_Click()
    
    '[ANIMATE BUTTON]
    frmSet.cmdholEndPlus.BorderStyle = 1
    Delay vbDelay
    frmSet.cmdholEndPlus.BorderStyle = 0
    
    '[CHANGE HOLEND DATE]
    If IsDate(frmSet.MaskHolEnd) Then frmSet.MaskHolEnd = Format(CDate(frmSet.MaskHolEnd) + 1, strDateFormat)
    
End Sub

Private Sub cmdholEndPlus_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Increase the 'Holiday End Date' by 1 day."
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub cmdHolEndToday_Click()
    
    '[ANIMATE BUTTON]
    cmdHolEndToday.BorderStyle = 1
    Delay vbDelay
    cmdHolEndToday.BorderStyle = 0

    '[SET MSK DATE TO TODAYS DATE]
    frmSet.MaskHolEnd.Text = Format(Date, strDateFormat)

End Sub

Private Sub cmdHolEndToday_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Set the 'Holiday End Date' to todays date (" & Format(Now, strDateFormat) & ")."

End Sub


Private Sub cmdHolStartMinus_Click()
    
    '[ANIMATE BUTTON]
    frmSet.cmdHolStartMinus.BorderStyle = 1
    Delay vbDelay
    frmSet.cmdHolStartMinus.BorderStyle = 0
    
    '[CHANGE HOLEND DATE]
    If IsDate(frmSet.MaskHolStart) Then frmSet.MaskHolStart = Format(CDate(frmSet.MaskHolStart) - 1, strDateFormat)

End Sub

Private Sub cmdHolStartMinus_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Decrease the 'Holiday Start Date' by 1 day."
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub cmdHolStartPlus_Click()
    
    '[ANIMATE BUTTON]
    frmSet.cmdHolStartPlus.BorderStyle = 1
    Delay vbDelay
    frmSet.cmdHolStartPlus.BorderStyle = 0
    
    '[CHANGE HOLEND DATE]
    If IsDate(frmSet.MaskHolStart) Then frmSet.MaskHolStart = Format(CDate(frmSet.MaskHolStart) + 1, strDateFormat)

End Sub

Private Sub cmdHolStartPlus_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Increase the 'Holiday Start Date' by 1 day."
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub cmdHolStartToday_Click()
    
    '[ANIMATE BUTTON]
    cmdHolStartToday.BorderStyle = 1
    Delay vbDelay
    cmdHolStartToday.BorderStyle = 0

    '[SET MSK DATE TO TODAYS DATE]
    frmSet.MaskHolStart.Text = Format(Date, strDateFormat)

End Sub

Private Sub cmdHolStartToday_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Set the 'Holiday Start Date' to todays date (" & Format(Now, strDateFormat) & ")."

End Sub


Private Sub cmdNew_Click()

    '[ANIMATE BUTTON]
    cmdNew.BorderStyle = 1
    Delay vbDelay
    cmdNew.BorderStyle = 0

    Dim Msg As String
    Dim Style
    Dim Response
    Dim Title
    '[ADD A NEW STAFF MEMBER TO THE LIST]
    '[CHECK TO SEE IF WE ARE MOVING FROM AN UNSAVED RECORD]
    If frmSet.cmdSave.Visible = True Then
        '[RECORD IS UNSAVED, POPUP YES/NO DIALOG]
        Msg = "You have made changes to this staff record (" & DsStaff("Lastname") & ", " & DsStaff("FirstName") & ") but have not saved these changes." & strBreak & strBreak & "If you choose not to save now, any changes you have made since your last save will be lost." & strBreak & strBreak & "Do you wish to save these changes before you add another staff member ?"
        Style = vbYesNoCancel              ' Define buttons.
        Title = "Staff Changes Not Saved"  ' Define title.
        Response = gsrMsg(Msg, Style, Title)
        If Response = vbYes Then    ' User chose Yes.
            '[COMMIT RECORD CHANGES TO THE DYNASET]
            SaveStaffDetails
        ElseIf Response = vbCancel Then
            Exit Sub
        End If
    End If
    
    '[CALL ROUTINE TO ADD NEW STAFF MEMBER AND REPOSITION LIST]
    AddNewStaff

End Sub

Private Sub cmdNew_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Add a new staff member to the list."
    '[---------------------------------------------------------------------------------]

End Sub





Private Sub cmdSave_Click()

    '[ANIMATE BUTTON]
    cmdSave.BorderStyle = 1
    Delay vbDelay
    cmdSave.BorderStyle = 0

    Dim strDisplayname      As String
    Dim intColCounter          As Integer

    '[COMMIT RECORD CHANGES TO THE DYNASET]
    SaveStaffDetails
    
    '[CONVERT DYNASET LASTNAME, FIRSTNAME TO DISPLAY FORMAT]
    strDisplayname = DsStaff("LastName") & ", " & DsStaff("FirstName")
    
    '[IF NAME HASN'T CHANGED, EXIT THIS SUB NOW]
    If frmSet.ListStaff.List(frmSet.ListStaff.ListIndex) = strDisplayname Then Exit Sub
    
    '[FILL STAFF LIST SO WE GET ORDER]
    FillStaffList
        
    '[RELOCATE STAFF NAME]
    For intColCounter = 0 To (frmSet.ListStaff.ListCount - 1)
        If frmSet.ListStaff.List(intColCounter) = strDisplayname Then frmSet.ListStaff.ListIndex = intColCounter
    Next intColCounter
    
    '[SHOW STAFF INFO]
    ShowStaffInfo
    
End Sub

Private Sub cmdSave_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Save any changes made to the current staff record."
    '[---------------------------------------------------------------------------------]

End Sub





Private Sub cmdSelectedStaffRoster_Click()

    '[ANIMATE BUTTON]
    cmdSelectedStaffRoster.BorderStyle = 1
    Delay vbDelay
    cmdSelectedStaffRoster.BorderStyle = 0

    '[CALL ROUTINE TO PROCESS SINGLE STAFF RECORD AND PRINT TIME SHEET]
    Call procSelectedStaffRoster

End Sub

Private Sub cmdSelectedStaffRoster_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Print the weekly roster for " & frmSet.ListStaff.List(frmSet.ListStaff.ListIndex) & "."
    '[---------------------------------------------------------------------------------]

End Sub




Private Sub cmdStartWeekMinus_Click()
    
    '[ANIMATE BUTTON]
    frmSet.cmdStartWeekMinus.BorderStyle = 1
    Delay vbDelay
    frmSet.cmdStartWeekMinus.BorderStyle = 0
    
    '[CHANGE HOLEND DATE]
    If IsDate(frmSet.MaskHolStart) Then frmSet.MaskHolStart = Format(CDate(frmSet.MaskHolStart) - 7, strDateFormat)

End Sub

Private Sub cmdStartWeekMinus_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Decrease the 'Holiday Start Date' by 7 days."
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub cmdStartWeekPlus_Click()
    
    '[ANIMATE BUTTON]
    frmSet.cmdStartWeekPlus.BorderStyle = 1
    Delay vbDelay
    frmSet.cmdStartWeekPlus.BorderStyle = 0
    
    '[CHANGE HOLEND DATE]
    If IsDate(frmSet.MaskHolStart) Then frmSet.MaskHolStart = Format(CDate(frmSet.MaskHolStart) + 7, strDateFormat)

End Sub

Private Sub cmdStartWeekPlus_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Increase the 'Holiday Start Date' by 7 days."
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub comboDayStatus_Click(Index As Integer)
    
    '[CHANGE VISIBILITY FOR START AND END TIME TEXT BOXES]
    If frmSet.comboDayStatus(Index).ListIndex <= 1 Then
        frmSet.textStart(Index).Visible = False
        frmSet.textFinish(Index).Visible = False
    Else
        frmSet.textStart(Index).Visible = True
        frmSet.textFinish(Index).Visible = True
        frmSet.lbl_std(1).Visible = True
        frmSet.lbl_std(2).Visible = True
    End If
    
    '[MAKE SAVE BUTTON VISIBLE]
    If frmSet.cmdSave.Visible = False Then frmSet.cmdSave.Visible = True

End Sub


Private Sub comboDayStatus_GotFocus(Index As Integer)

    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Select the desired staff availability for this day."
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub Form_Load()

    Dim intFlag             As Integer
    
    '[DEBUG]
    Call LogToFile("=-STAFF SECT - LOAD SUB------------------=")

    '[SET STAFF FRAME FIRST]
    frmSet.frameStaff.ZOrder

    '[DEBUG]
    Call LogToFile("- filling staff list")
    '[FILL STAFF LIST FROM DYNASET]
    Call SplashProgress("Filling staff list.", 50)
    FillStaffList
    '[DEBUG]
    Call LogToFile("- setting class labels")
    '[APPLY CLASS LABELS]
    Call SplashProgress("Setting class labels.", 52)
    SetClassLabels
    '[DEBUG]
    Call LogToFile("- setting day labels")
    '[APPLY DAY LABELS]
    Call SplashProgress("Setting day labels.", 54)
    SetDayLabels
    '[DEBUG]
    Call LogToFile("- displaying staff details")
    '[LOCATE FIRST STAFF MEMBER]
    If frmSet.ListStaff.ListCount > 0 Then frmSet.ListStaff.ListIndex = 0
        
      '[DEBUG]
    Call LogToFile("=-CONTROL SECT - LOAD SUB----------------=")

    '[Resize controls and grids to match]
    frmSet.GridClass.ColWidth(3) = frmSet.ImageSwitch(0).Width
    frmSet.GridClass.ColWidth(0) = (GridClass.Width - frmSet.GridClass.ColWidth(3)) * 0.12
    frmSet.GridClass.ColWidth(1) = (GridClass.Width - frmSet.GridClass.ColWidth(3)) * 0.3
    frmSet.GridClass.ColWidth(2) = (GridClass.Width - frmSet.GridClass.ColWidth(3)) * 0.55
    
    '[DEBUG]
    Call LogToFile("- filling class grid")
    '[CALL ROUTINE TO FILL GRID]
    Call SplashProgress("Filling class grid.", 60)
    FillClassGrid
    
    '[DEBUG]
    Call LogToFile("- setting start day and increments")
    '[PLACE DEFAULT VALUES IN COMBOBOXES AND MASK BOXES]
    ComboStartDay.ListIndex = DsDefault("StartDay") - 1
    ComboIncrement.ListIndex = DsDefault("Increment") - 1
    
    '[DEBUG]
    Call LogToFile("- setting start hour and minutes")
    '[SET START HOUR AND MINUTE]
    frmSet.ComboHour = Format(Hour(DsDefault("StartTime")), "0#")
    frmSet.ComboMinute = Format(Minute(DsDefault("StartTime")), "0#")
    ProcessDayLength
    
    '[DEBUG]
    Call LogToFile("- setting date captions")
    '[SET DATE FORMAT MASK]
    Call SplashProgress("Setting control panel options.", 62)
    frmSet.optionDateFormat(0).Caption = Format(Date, strDateFormat)
    frmSet.optionDateFormat(1).Caption = Format(Date, "Medium Date")
    '[PLACE CUSTOM DATE MASK IN CONTROL]
    frmSet.maskDateMask = DsDefault("CustomDate")
 
    '[------------------------------------------------------------------------------------------]
    '[SET CONTROL FORM DATE TYPE]
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    '[DEBUG]
    Call LogToFile("- setting date display format")
    Call SplashProgress("Setting date display format.", 65)
    '[---------------------------------------------------------------------------------]
    frmSet.optionDateFormat(DsDefault("DateFormat").Value).Value = True
    '[DEBUG]
    Call LogToFile("= " & Format(Date$, strDateFormat))
    
    '[RESTORE DELETE CONFIRM AND SOUND FLAGS]
    '[DELETE IS 1]
    
    '[SOUNDS IS 2]
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    '[DEBUG]
    Call LogToFile("- restoring default options")
    Call SplashProgress("Restoring default options.", 67)
    
    '[---------------------------------------------------------------------------------]
    '[ALLSHIFTS EXCEPTION CHECK]
    '[REV: 3.00.28]
    If DsDefault!AllShifts = 0 Then
        frmSet.CheckAllShifts.Value = 0
    Else
        frmSet.CheckAllShifts.Value = 1
    End If
    intFlag = DsDefault("DeleteConfirm")
    If intFlag > (2 ^ 2) Then
        frmSet.CheckSounds.Value = 1
        frmSet.CheckDelete.Value = 1
    ElseIf intFlag = (2 ^ 2) Then
        frmSet.CheckDelete.Value = 0
        frmSet.CheckSounds.Value = 1
    ElseIf intFlag = (1 ^ 2) Then
        frmSet.CheckDelete.Value = 1
        frmSet.CheckSounds.Value = 0
    Else
        frmSet.CheckSounds.Value = 0
        frmSet.CheckDelete.Value = 0
    End If
    
    '[SET BREAK TIME AND WORK PERIOD ON CONTROL FORM]
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    Call LogToFile("- setting break times")
    Call SplashProgress("Setting break times.", 70)
    '[---------------------------------------------------------------------------------]
    frmSet.ComboBlockHour.Text = Format(DsDefault("BlockHour"), "0#")
    frmSet.ComboBlockMin.Text = Format(DsDefault("BlockMin"), "0#")
    frmSet.ComboBreakHour.Text = Format(DsDefault("BreakHour"), "0#")
    frmSet.ComboBreakMin.Text = Format(DsDefault("BreakMin"), "0#")
 
    '[SET STARTING POSITION FOR FORM]
    frmSet.Top = DsDefault("ControlTop")
    frmSet.Left = DsDefault("ControlLeft")
    '[CHECK FOR > WIDTH, HEIGHT]
    If frmSet.Top > Screen.Height Or frmSet.Left > Screen.Width Then
        '[center form on screen]
        frmSet.Top = (Screen.Height / 2) - (frmSet.Height / 2)
        frmSet.Left = (Screen.Width / 2) - (frmSet.Width / 2)
    End If
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "The staff form allows you to add, delete and modify staff records."
    '[---------------------------------------------------------------------------------]

End Sub



Private Sub Label_std_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Select Case Index
    Case 4
        '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
        StatusBar "Birth date of the staff member."
        '[---------------------------------------------------------------------------------]
    Case 5
        '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
        StatusBar "Employment date of the staff member."
        '[---------------------------------------------------------------------------------]
    Case 6
        '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
        StatusBar "Hourly rate of the staff member."
        '[---------------------------------------------------------------------------------]
    Case 7
        '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
        StatusBar "Home phone number of the staff member (not essential)."
        '[---------------------------------------------------------------------------------]
    Case 12
        '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
        StatusBar "Minimum hours a week to be worked ."
        '[---------------------------------------------------------------------------------]
    Case 13
        '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
        StatusBar "Maximum hours a week to be worked."
        '[---------------------------------------------------------------------------------]
    Case Else
    End Select
    
End Sub



Private Sub ListStaff_Click()

    '[LOCATE SELECTED STAFF MEMBER]
    LocateStaff

End Sub


Private Sub ListStaff_KeyUp(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
    Case vbKeyDelete
        '[CAPTURE DELETE KEY]
        Call cmdDelete_Click
    Case Else
    End Select
    
End Sub


Private Sub ListStaff_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Display the details for the selected staff member."
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub MaskBirthDate_Change()

    '[ENABLE SAVE BUTTON IF A CHANGE HAS BEEN MADE]
    If frmSet.cmdSave.Visible = False Then frmSet.cmdSave.Visible = True

End Sub

Private Sub MaskDateHired_Change()

    '[ENABLE SAVE BUTTON IF A CHANGE HAS BEEN MADE]
    If frmSet.cmdSave.Visible = False Then frmSet.cmdSave.Visible = True

End Sub


Private Sub MaskHolEnd_Change()
    
    '[IF VALID DATE, SET DAYS TO VALID FIGURE AND SHOW SAVE BUTTON]
    If IsDate(frmSet.MaskHolStart.Text) And IsDate(frmSet.MaskHolEnd.Text) Then
        frmSet.txtHolDays = Format((CDate(frmSet.MaskHolEnd) - CDate(frmSet.MaskHolStart)) + 1, "###0")
        If CDate(frmSet.MaskHolEnd) <> DsStaff!HolEnd Then
            If frmSet.cmdSave.Visible = False Then frmSet.cmdSave.Visible = True
        End If
    End If

End Sub

Private Sub MaskHolEnd_GotFocus()
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Set the 'Holiday End Date'."
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub MaskHolStart_Change()

    '[IF VALID DATE, SET DAYS TO VALID FIGURE AND SHOW SAVE BUTTON]
    If IsDate(frmSet.MaskHolStart.Text) And IsDate(frmSet.MaskHolEnd.Text) Then
        frmSet.txtHolDays = Format((CDate(frmSet.MaskHolEnd) - CDate(frmSet.MaskHolStart)) + 1, "###0")
        If CDate(frmSet.MaskHolStart) <> DsStaff!HolStart Then
            If frmSet.cmdSave.Visible = False Then frmSet.cmdSave.Visible = True
        End If
    Else
        frmSet.txtHolDays = ""
    End If

End Sub

Private Sub MaskHolStart_GotFocus()
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Set the 'Holiday Start Date'."
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub MaskHomePhone_Change()


    '[ENABLE SAVE BUTTON IF A CHANGE HAS BEEN MADE]
    If frmSet.cmdSave.Visible = False Then frmSet.cmdSave.Visible = True

End Sub

Private Sub MaskHourRate_Change()

    '[ENABLE SAVE BUTTON IF A CHANGE HAS BEEN MADE]
    If frmSet.cmdSave.Visible = False Then frmSet.cmdSave.Visible = True

End Sub



Private Sub MaskMaxHours_Change()
    
    '[ENABLE SAVE BUTTON IF A CHANGE HAS BEEN MADE]
    If frmSet.cmdSave.Visible = False Then frmSet.cmdSave.Visible = True

End Sub

Private Sub MaskMaxHours_GotFocus()

    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Enter the maximum hours required per week by this staff member."
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub MaskMinHours_Change()
    
    '[ENABLE SAVE BUTTON IF A CHANGE HAS BEEN MADE]
    If frmSet.cmdSave.Visible = False Then frmSet.cmdSave.Visible = True

End Sub

Private Sub MaskMinHours_GotFocus()

    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Enter the minimum hours required per week by this staff member."
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub maskRosterRate_Change(Index As Integer)

    '[ENABLE SAVE BUTTON IF A CHANGE HAS BEEN MADE]
    If frmSet.cmdSave.Visible = False Then frmSet.cmdSave.Visible = True

End Sub







Private Sub maskRosterRate_GotFocus(Index As Integer)

    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Enter the hourly rate earned by this staff member when allocated to this roster."
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub optionPay_Click(Index As Integer)

    '[REV: 3.00.28]
    '[ENABLE SAVE BUTTON IF A CHANGE HAS BEEN MADE]
    If frmSet.cmdSave.Visible = False Then frmSet.cmdSave.Visible = True

    '[CYCLE THROUGH AND HIDE/SHOW HOURLY RATES ON THE ROSTER TAB]
    Dim intCounter          As Integer
    '[MAKE RATE VISIBLE IF HOURLY PAY RATE IS SET AND CLASS IS CHECKED]
    For intCounter = 0 To 9
        If frmSet.optionPay(vbHourly).Value = True Then
            If frmSet.CheckClass(intCounter).Value = vbChecked Then
                frmSet.maskRosterRate(intCounter).Visible = True
                frmSet.cmdBaseRate(intCounter).Visible = True
                frmSet.lbl_hr_Rate.Visible = True
            End If
        Else
            frmSet.maskRosterRate(intCounter).Visible = False
            frmSet.cmdBaseRate(intCounter).Visible = False
        End If
    Next intCounter
    
    

End Sub

Private Sub picsetBar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "The GSR Control Panel allows you to change staff, roster and program settings."
    '[---------------------------------------------------------------------------------]

End Sub



Private Sub textFinish_Change(Index As Integer)

    Dim strField            As String
    strField = "Start_" & Trim(Str(Index + 1))
    If IsDate(textStart(Index).Text) And DsStaff(strField).Value <> Format(textStart(Index).Text, "Medium Time") Then
        '[DATA HAS CHANGED FROM THAT STORED - SHOW SAVE BUTTON]
        If frmSet.cmdSave.Visible = False Then frmSet.cmdSave.Visible = True
    End If

End Sub

Private Sub textFinish_DblClick(Index As Integer)
    
    '[CHANGE STARTING TIME]
    Load frmTime
    '[SET TIME ON FORM]
    If textFinish(Index).Text = "" Then
        frmTime.Caption = "Select Finish Time"
        frmTime.ComboHour = Format(Hour(DsDefault("EndTime")), "0#")
        frmTime.ComboMinute = Format(Minute(DsDefault("EndTime")), "0#")
        frmTime.Refresh
    Else
        If IsDate(textFinish(Index).Text) Then
            frmTime.ComboHour = Format(Hour(CDate(textFinish(Index).Text)), "0#")
            frmTime.ComboMinute = Format(Minute(CDate(textFinish(Index).Text)), "0#")
        Else
            frmTime.ComboHour = Format(Hour(DsDefault("EndTime")), "0#")
            frmTime.ComboMinute = Format(Minute(DsDefault("EndTime")), "0#")
        End If
    End If
    frmTime.Show 1
    '[PROCESS RESULT OF FORM]
    If frmTime.CheckResult = vbChecked Then
        '[ENSURE TIME RETURNED IS A TIME]
        If IsDate(frmTime.ComboHour.Text & ":" & frmTime.ComboMinute.Text) Then
            textFinish(Index).Text = Format(frmTime.ComboHour.Text & ":" & frmTime.ComboMinute.Text, "Medium Time")
        End If
    End If

End Sub


Private Sub textFinish_GotFocus(Index As Integer)

    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Enter the desired finish time or double-click for the time popup form."
    '[---------------------------------------------------------------------------------]

End Sub

Private Sub textFinish_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Enter the desired finish time or double-click for the time popup form."
    '[---------------------------------------------------------------------------------]

End Sub

Private Sub TextFirstName_Change()


    '[ENABLE SAVE BUTTON IF A CHANGE HAS BEEN MADE]
    If frmSet.cmdSave.Visible = False Then frmSet.cmdSave.Visible = True

End Sub

Private Sub TextFirstName_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "First name of the staff member (or initial if more space is required)."
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub TextLastName_Change()

    '[ENABLE SAVE BUTTON IF A CHANGE HAS BEEN MADE]
    If frmSet.cmdSave.Visible = False Then frmSet.cmdSave.Visible = True

End Sub

Private Sub TextLastName_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Staff members last name."
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub TextMiddleName_Change()


    '[ENABLE SAVE BUTTON IF A CHANGE HAS BEEN MADE]
    If frmSet.cmdSave.Visible = False Then frmSet.cmdSave.Visible = True

End Sub

Private Sub TextMiddleName_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Middle name of the staff member (not essential)."
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub TextNote_Change()

    '[ENABLE SAVE BUTTON IF A CHANGE HAS BEEN MADE]
    If frmSet.cmdSave.Visible = False Then frmSet.cmdSave.Visible = True

End Sub

Private Sub TextNote_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Enter notes for this staff member (will be printed on the staff timesheet) - up to 255 characters."
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub TextStaffID_Change()

    '[ENABLE SAVE BUTTON IF A CHANGE HAS BEEN MADE]
    If frmSet.cmdSave.Visible = False Then frmSet.cmdSave.Visible = True

End Sub


Private Sub TextStaffID_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Unique staff identification code."
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub textStart_Change(Index As Integer)

    Dim strField            As String
    strField = "Start_" & Trim(Str(Index + 1))
    If IsDate(textStart(Index).Text) And DsStaff(strField).Value <> Format(textStart(Index).Text, "Medium Time") Then
        '[DATA HAS CHANGED FROM THAT STORED - SHOW SAVE BUTTON]
        If frmSet.cmdSave.Visible = False Then frmSet.cmdSave.Visible = True
    End If

End Sub

Private Sub textStart_DblClick(Index As Integer)

    '[CHANGE STARTING TIME]
    Load frmTime
    '[SET TIME ON FORM]
    If textStart(Index).Text = "" Then
        frmTime.Caption = "Select Start Time"
        frmTime.ComboHour = Format(Hour(DsDefault("StartTime")), "0#")
        frmTime.ComboMinute = Format(Minute(DsDefault("StartTime")), "0#")
        frmTime.Refresh
    Else
        If IsDate(textStart(Index).Text) Then
            frmTime.ComboHour = Format(Hour(CDate(textStart(Index).Text)), "0#")
            frmTime.ComboMinute = Format(Minute(CDate(textStart(Index).Text)), "0#")
        Else
            frmTime.ComboHour = Format(Hour(DsDefault("StartTime")), "0#")
            frmTime.ComboMinute = Format(Minute(DsDefault("StartTime")), "0#")
        End If
    End If
    frmTime.Show 1
    '[PROCESS RESULT OF FORM]
    If frmTime.CheckResult = vbChecked Then
        '[ENSURE TIME RETURNED IS A TIME]
        If IsDate(frmTime.ComboHour.Text & ":" & frmTime.ComboMinute.Text) Then
            textStart(Index).Text = Format(frmTime.ComboHour.Text & ":" & frmTime.ComboMinute.Text, "Medium Time")
        End If
    End If

End Sub


Private Sub textStart_GotFocus(Index As Integer)

    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Enter the desired start time or double-click for the time popup form."
    '[---------------------------------------------------------------------------------]

End Sub

Private Sub textStart_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Enter the desired start time or double-click for the time popup form."
    '[---------------------------------------------------------------------------------]


End Sub


Private Sub txtHolDays_Change()

    '[ADD NUMBER OF DAYS TO END DATE]
    If IsDate(frmSet.MaskHolStart) And Val(txtHolDays) > 0 And Val(txtHolDays) < 36500 Then
        If Not (IsDate(frmSet.MaskHolEnd)) Then
            frmSet.MaskHolEnd = Format(CDate(frmSet.MaskHolStart) + (Val(txtHolDays) - 1), strDateFormat)
        Else
            If CDate(frmSet.MaskHolEnd) <> CDate(frmSet.MaskHolStart) + (Val(txtHolDays) - 1) Then
                frmSet.MaskHolEnd = Format(CDate(frmSet.MaskHolStart) + (Val(txtHolDays) - 1), strDateFormat)
            End If
        End If
    End If

End Sub


Private Sub txtHolDays_GotFocus()
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Manually set the number of days this staff member will be absent for."
    '[---------------------------------------------------------------------------------]

End Sub

Private Sub txtHolDays_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Manually set the number of days this staff member will be absent for."
    '[---------------------------------------------------------------------------------]

End Sub




Private Sub CheckAllShifts_Click()

    '[REV: 3.00.28]
    '[CHANGE PUBLIC VARIABLE ALLSHIFTS]
    If frmSet.CheckAllShifts.Value = 1 Then
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
    If frmSet.CheckDelete.Value = 1 Then
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
    If frmSet.CheckSounds.Value = 1 Then
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




Private Sub ComboBlockHour_Click()

    '[SET DYNASET VALUE]
    DsDefault.Edit
        DsDefault("BlockHour") = frmSet.ComboBlockHour.Text
    DsDefault.Update


End Sub


Private Sub ComboBlockMin_Click()
    
    '[SET DYNASET VALUE]
    DsDefault.Edit
        DsDefault("BlockMin") = frmSet.ComboBlockMin.Text
    DsDefault.Update

End Sub


Private Sub ComboBreakHour_Click()
    
    '[SET DYNASET VALUE]
    DsDefault.Edit
        DsDefault("BreakHour") = frmSet.ComboBreakHour.Text
    DsDefault.Update

End Sub


Private Sub ComboBreakMin_Click()
    
    '[SET DYNASET VALUE]
    DsDefault.Edit
        DsDefault("BreakMin") = frmSet.ComboBreakMin.Text
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
    If X > frmSet.GridClass.ColPos(3) Then
        StatusBar "Activate/deactivate the selected roster.  Only active rosters can be modified and processed."
    ElseIf X > frmSet.GridClass.ColPos(2) Then
        StatusBar "Enter notes which will be printed with the weekly roster."
    ElseIf X > frmSet.GridClass.ColPos(1) Then
        StatusBar "Modify the roster description, a 20 character identifier for the selected roster."
    Else
        StatusBar "Modify the roster identifier, a three character short form for the roster name."
    End If
    '[---------------------------------------------------------------------------------]

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
            frmSet.cmdNukeCover.Visible = True
            frmSet.cmdNuke.Visible = False
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
    frmSet.cmdNukeCover.Visible = True
    frmSet.cmdNuke.Visible = False
    '[***********************************************************************]
    '[WARNING - DELETE ALL ROSTERS CURRENTLY STORED IN THE DATABASE - WARNING]
    '[***********************************************************************]
    
End Sub
Private Sub cmdNukeCover_Click()

    '[HIDE NUKE COVER]
    frmSet.cmdNukeCover.Visible = False
    frmSet.cmdNuke.Visible = True
    
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




Private Sub GridClass_DblClick()

    '[USER DOUBLE CLICKED ON CLASS GRID, POPUP INPUT BOX FOR NEW VALUE]
    Dim strTemp             As String   '[Temporary Variable for Cell Content]
    Dim strMessage          As String
    Dim strTitle            As String
    Dim strOldClass         As String
    Dim SQLStmt             As String
    Dim intColCounter       As Integer
    Dim intCol, intRow      As Integer
    Dim intClass            As Integer
    
    '[SAVE CURRENT LOCATION]
    intCol = GridClass.Col
    intRow = GridClass.Row
    
    '[EXIT IF TITLE ROW CLICKED]
    If intRow = 0 Then Exit Sub
    '[ROSTER CLASS]
    intClass = GridClass.Row
    
    Select Case GridClass.Col
    Case 0
        strMessage = "The roster code is used as a short form to identify which rosters an employee belongs to.  You are allowed to use up to three alpha-numeric characters to define this roster."
    Case 1
        strMessage = "The roster description is used on reports to provide a longer identifier for each roster.  You may use up to 20 alpha-numeric characters in this field."
    End Select
    
    strTitle = "Class Definitions"
    strOldClass = GridClass.Text
    strTemp = GridClass.Text
    
    If GridClass.Col = 3 Then
        '[ENABLED/DISABLED]
        If GridClass.Text = vbChecked Then
            GridClass.Text = vbUnchecked
            frmSet.GridClass.Picture = frmSet.ImageSwitch(constCritical).Picture
        Else
            GridClass.Text = vbChecked
            frmSet.GridClass.Picture = frmSet.ImageSwitch(constWarning).Picture
        End If
        GridClass.Refresh
    ElseIf GridClass.Col = 2 Then
        '[PLACE ROSTER TEXT IN THE GRID]
        Dim flagResult      As Boolean
        
        gsrNote = GridClass.Text
        GridClass.Col = 1
        flagResult = gsrMsg(gsrNote, vbQuestion, "Enter " & GridClass.Text & " Roster Notes")
        GridClass.Col = 2
        Select Case flagResult
        Case vbOK
            If Len(Trim(gsrNote)) > 255 Then gsrNote = Left(gsrNote, 255)
            GridClass.Text = gsrNote
        Case Else
        End Select
    Else
        '[GET USER INPUT]
        strTemp = InputBox(strMessage, strTitle, strTemp)
        '[EXIT IF NOTHING HAS CHANGED]
        If strTemp = strOldClass Then Exit Sub
        
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
                If Len(strTemp) > 20 Then GridClass.Text = Left$(strTemp, 20) Else GridClass.Text = strTemp
            End If
        Else
            '[RESTORE LOCATION IN GRID]
            GridClass.Col = intCol
            GridClass.Row = intRow
            Exit Sub
        End If
    End If
    
    '[PLACE VALUES IN DYNASET]
    DsClass.FindFirst ("[ID]=" & intClass)
    If Not DsClass.NoMatch Then
        DsClass.Edit
            frmSet.GridClass.Col = 0
            DsClass("Code") = frmSet.GridClass.Text
            frmSet.GridClass.Col = 1
            DsClass("Description") = frmSet.GridClass.Text
            frmSet.GridClass.Col = 2
            DsClass("Note") = frmSet.GridClass.Text
            frmSet.GridClass.Col = 3
            DsClass("Active") = frmSet.GridClass.Text
        DsClass.Update
    End If
    
    '[ONLY UPDATE IF NAME OR STATE HAS CHANGED]
    If intCol = 1 Or intCol = 3 Then
        '[SET CLASS LABELS ON STAFF FORM]
        Call SetClassLabels
        
        '[FILL ROSTER LIST]
        Call FillRosterList
    End If
    
    '[RESTORE LOCATION IN GRID]
    GridClass.Col = intCol
    GridClass.Row = intRow
    
End Sub




Private Sub maskDateMask_Change()

    '[CHECK TO SEE IF THIS IS A DATE]
    If Not IsDate(Format(Now, frmSet.maskDateMask)) Then
        '[NOT DATE FORMAT, DISABLE OPTION 3]
        frmSet.optionDateFormat(2).Enabled = False
        '[IF OPTION 3 WAS CHECKED, MOVE TO FIRST OPTION]
        If frmSet.optionDateFormat(2).Value = True Then frmSet.optionDateFormat(0).Value = True
    Else
        '[IS DATE FORMAT, ENABLE OPTION 3]
        frmSet.optionDateFormat(2).Enabled = True
    End If

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
    If frmSet.cmdSave.Visible = True Then flagSave = True Else flagSave = False
    frmSet.MaskBirthDate.Format = strDateFormat
    frmSet.MaskBirthDate.Text = Format(frmSet.MaskBirthDate.Text, strDateFormat)
    frmSet.MaskDateHired.Format = strDateFormat
    frmSet.MaskDateHired.Text = Format(frmSet.MaskDateHired.Text, strDateFormat)
    frmSet.MaskHolStart.Format = strDateFormat
    frmSet.MaskHolStart.Text = Format(frmSet.MaskHolStart.Text, strDateFormat)
    frmSet.MaskHolEnd.Format = strDateFormat
    frmSet.MaskHolEnd.Text = Format(frmSet.MaskHolEnd.Text, strDateFormat)
    '[RESET STAFF SAVE COMMAND BUTTON STATE]
    If flagSave = False Then frmSet.cmdSave.Visible = False
    

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





