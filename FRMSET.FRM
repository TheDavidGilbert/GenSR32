VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{A8B3B723-0B5A-101B-B22E-00AA0037B2FC}#1.0#0"; "GRID32.OCX"
Begin VB.Form frmSet 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "GSR Control Panel"
   ClientHeight    =   5595
   ClientLeft      =   3600
   ClientTop       =   3585
   ClientWidth     =   7380
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FRMSET.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5595
   ScaleWidth      =   7380
   Begin Threed.SSPanel panelStatusBar 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   15
      Top             =   5295
      Width           =   7380
      _Version        =   65536
      _ExtentX        =   13018
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
      BorderWidth     =   1
      RoundedCorners  =   0   'False
      FloodColor      =   -2147483648
      FloodShowPct    =   0   'False
      Alignment       =   2
   End
   Begin TabDlg.SSTab tabSet 
      Height          =   5295
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7395
      _ExtentX        =   13044
      _ExtentY        =   9340
      _Version        =   327681
      Style           =   1
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   529
      TabCaption(0)   =   "Roster &Details"
      TabPicture(0)   =   "FRMSET.frx":08CA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblRosterDetailsInfo"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdClose(4)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "GridClass"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "&Staff Details"
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdClose(3)"
      Tab(1).Control(1)=   "ListStaff"
      Tab(1).Control(2)=   "tabStaff"
      Tab(1).Control(3)=   "panelStaffToolbar"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "&Roster Settings"
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "ShapeBorder(2)"
      Tab(2).Control(1)=   "label_std(18)"
      Tab(2).Control(2)=   "label_std(22)"
      Tab(2).Control(3)=   "label_std(21)"
      Tab(2).Control(4)=   "label_std(19)"
      Tab(2).Control(5)=   "label_std(26)"
      Tab(2).Control(6)=   "label_std(25)"
      Tab(2).Control(7)=   "label_std(24)"
      Tab(2).Control(8)=   "labelHeading(5)"
      Tab(2).Control(9)=   "label_std(23)"
      Tab(2).Control(10)=   "lblRosterSettingsInfo"
      Tab(2).Control(11)=   "label_std(44)"
      Tab(2).Control(12)=   "labelHeading(2)"
      Tab(2).Control(13)=   "SSPanel1"
      Tab(2).Control(14)=   "ComboStartDay"
      Tab(2).Control(15)=   "ComboBlockMin"
      Tab(2).Control(16)=   "ComboBlockHour"
      Tab(2).Control(17)=   "ComboBreakMin"
      Tab(2).Control(18)=   "ComboBreakHour"
      Tab(2).Control(19)=   "cmdClose(2)"
      Tab(2).ControlCount=   20
      TabCaption(3)   =   "&Program Settings"
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "cmdNuke"
      Tab(3).Control(1)=   "ShapeBorder(1)"
      Tab(3).Control(2)=   "labelHeading(3)"
      Tab(3).Control(3)=   "label_std(27)"
      Tab(3).Control(4)=   "maskDateMask"
      Tab(3).Control(5)=   "label_std(28)"
      Tab(3).Control(6)=   "labelHeading(1)"
      Tab(3).Control(7)=   "cmdPrintSet"
      Tab(3).Control(8)=   "cmdNukeCover"
      Tab(3).Control(9)=   "label_std(30)"
      Tab(3).Control(10)=   "label_std(29)"
      Tab(3).Control(11)=   "label_std(45)"
      Tab(3).Control(12)=   "CheckSounds"
      Tab(3).Control(13)=   "CheckDelete"
      Tab(3).Control(14)=   "CheckAllShifts"
      Tab(3).Control(15)=   "OptionDateFormat(0)"
      Tab(3).Control(16)=   "OptionDateFormat(1)"
      Tab(3).Control(17)=   "OptionDateFormat(2)"
      Tab(3).Control(18)=   "cmdClose(1)"
      Tab(3).ControlCount=   19
      TabCaption(4)   =   "&Close the Control Panel"
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "cmdClose(0)"
      Tab(4).ControlCount=   1
      Begin MSGrid.Grid GridClass 
         Height          =   3375
         Left            =   120
         TabIndex        =   156
         Top             =   480
         Width           =   7215
         _Version        =   65536
         _ExtentX        =   12726
         _ExtentY        =   5953
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
         Cols            =   4
         FixedCols       =   0
      End
      Begin VB.CommandButton cmdClose 
         Cancel          =   -1  'True
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   7080
         TabIndex        =   144
         Top             =   30
         Width           =   315
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   -67920
         TabIndex        =   143
         Top             =   30
         Width           =   315
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   -67920
         TabIndex        =   142
         Top             =   30
         Width           =   315
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   -67920
         TabIndex        =   141
         Top             =   30
         Width           =   315
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   -67920
         TabIndex        =   140
         Top             =   30
         Width           =   315
      End
      Begin VB.OptionButton OptionDateFormat 
         Caption         =   "Custom Format"
         Height          =   315
         Index           =   2
         Left            =   -74820
         TabIndex        =   74
         Top             =   2880
         Width           =   1995
      End
      Begin VB.OptionButton OptionDateFormat 
         Height          =   315
         Index           =   1
         Left            =   -74820
         TabIndex        =   73
         Top             =   2520
         Width           =   2055
      End
      Begin VB.OptionButton OptionDateFormat 
         Height          =   315
         Index           =   0
         Left            =   -74820
         TabIndex        =   72
         Top             =   2160
         Width           =   1995
      End
      Begin VB.CheckBox CheckAllShifts 
         Alignment       =   1  'Right Justify
         Caption         =   "Require All Shifts Filled"
         Height          =   315
         Left            =   -74820
         TabIndex        =   71
         Top             =   1380
         Width           =   2400
      End
      Begin VB.CheckBox CheckDelete 
         Alignment       =   1  'Right Justify
         Caption         =   "Confirm Deletes"
         Height          =   315
         Left            =   -74820
         TabIndex        =   70
         Top             =   1080
         Width           =   2400
      End
      Begin VB.CheckBox CheckSounds 
         Alignment       =   1  'Right Justify
         Caption         =   "Sounds"
         Height          =   315
         Left            =   -74820
         TabIndex        =   69
         Top             =   780
         Width           =   2400
      End
      Begin VB.ComboBox ComboBreakHour 
         Height          =   330
         ItemData        =   "FRMSET.frx":08E6
         Left            =   -74040
         List            =   "FRMSET.frx":0940
         TabIndex        =   67
         Text            =   "ComboHour"
         Top             =   4140
         Width           =   900
      End
      Begin VB.ComboBox ComboBreakMin 
         Height          =   330
         ItemData        =   "FRMSET.frx":09A4
         Left            =   -73080
         List            =   "FRMSET.frx":09B4
         TabIndex        =   68
         Text            =   "ComboMinute"
         Top             =   4140
         Width           =   900
      End
      Begin VB.ComboBox ComboBlockHour 
         Height          =   330
         ItemData        =   "FRMSET.frx":09C8
         Left            =   -74040
         List            =   "FRMSET.frx":0A22
         TabIndex        =   65
         Text            =   "ComboHour"
         Top             =   3780
         Width           =   900
      End
      Begin VB.ComboBox ComboBlockMin 
         Height          =   330
         ItemData        =   "FRMSET.frx":0A86
         Left            =   -73080
         List            =   "FRMSET.frx":0A96
         TabIndex        =   66
         Text            =   "ComboMinute"
         Top             =   3780
         Width           =   900
      End
      Begin VB.ComboBox ComboStartDay 
         Height          =   330
         ItemData        =   "FRMSET.frx":0AAA
         Left            =   -74880
         List            =   "FRMSET.frx":0AC3
         Style           =   2  'Dropdown List
         TabIndex        =   64
         Top             =   780
         Width           =   2775
      End
      Begin VB.ListBox ListStaff 
         Appearance      =   0  'Flat
         Height          =   4440
         Left            =   -74940
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   780
         Width           =   1935
      End
      Begin TabDlg.SSTab tabStaff 
         Height          =   4755
         HelpContextID   =   40
         Left            =   -72960
         TabIndex        =   3
         Top             =   420
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   8387
         _Version        =   327681
         Style           =   1
         Tabs            =   4
         TabsPerRow      =   4
         TabHeight       =   529
         WordWrap        =   0   'False
         TabCaption(0)   =   "Staff Details"
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "ShapeBorder(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "labelInfo"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "label_std(1)"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "label_std(2)"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "label_std(3)"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "label_std(0)"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "label_std(7)"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "label_std(6)"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "label_std(5)"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "label_std(4)"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "labelHeading(0)"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).Control(11)=   "MaskHomePhone"
         Tab(0).Control(11).Enabled=   0   'False
         Tab(0).Control(12)=   "MaskBirthDate"
         Tab(0).Control(12).Enabled=   0   'False
         Tab(0).Control(13)=   "MaskDateHired"
         Tab(0).Control(13).Enabled=   0   'False
         Tab(0).Control(14)=   "MaskHourRate"
         Tab(0).Control(14).Enabled=   0   'False
         Tab(0).Control(15)=   "OptionPay(1)"
         Tab(0).Control(15).Enabled=   0   'False
         Tab(0).Control(16)=   "OptionPay(0)"
         Tab(0).Control(16).Enabled=   0   'False
         Tab(0).Control(17)=   "TextLastName"
         Tab(0).Control(17).Enabled=   0   'False
         Tab(0).Control(18)=   "TextFirstName"
         Tab(0).Control(18).Enabled=   0   'False
         Tab(0).Control(19)=   "TextMiddleName"
         Tab(0).Control(19).Enabled=   0   'False
         Tab(0).Control(20)=   "TextStaffID"
         Tab(0).Control(20).Enabled=   0   'False
         Tab(0).Control(21)=   "TextNote"
         Tab(0).Control(21).Enabled=   0   'False
         Tab(0).ControlCount=   22
         TabCaption(1)   =   "Availability"
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "labelDay(0)"
         Tab(1).Control(1)=   "labelDay(1)"
         Tab(1).Control(2)=   "labelDay(2)"
         Tab(1).Control(3)=   "labelDay(3)"
         Tab(1).Control(4)=   "labelDay(4)"
         Tab(1).Control(5)=   "labelDay(5)"
         Tab(1).Control(6)=   "labelDay(6)"
         Tab(1).Control(7)=   "lbl_std(1)"
         Tab(1).Control(8)=   "lbl_std(2)"
         Tab(1).Control(9)=   "MaskMinHours"
         Tab(1).Control(10)=   "MaskMaxHours"
         Tab(1).Control(11)=   "label_std(31)"
         Tab(1).Control(12)=   "label_std(32)"
         Tab(1).Control(13)=   "label_std(33)"
         Tab(1).Control(14)=   "label_std(34)"
         Tab(1).Control(15)=   "label_std(35)"
         Tab(1).Control(16)=   "label_std(36)"
         Tab(1).Control(17)=   "comboDayStatus(0)"
         Tab(1).Control(18)=   "comboDayStatus(1)"
         Tab(1).Control(19)=   "comboDayStatus(2)"
         Tab(1).Control(20)=   "comboDayStatus(3)"
         Tab(1).Control(21)=   "comboDayStatus(4)"
         Tab(1).Control(22)=   "comboDayStatus(5)"
         Tab(1).Control(23)=   "comboDayStatus(6)"
         Tab(1).Control(24)=   "textStart(0)"
         Tab(1).Control(25)=   "textStart(1)"
         Tab(1).Control(26)=   "textStart(2)"
         Tab(1).Control(27)=   "textStart(3)"
         Tab(1).Control(28)=   "textStart(4)"
         Tab(1).Control(29)=   "textStart(5)"
         Tab(1).Control(30)=   "textStart(6)"
         Tab(1).Control(31)=   "textFinish(0)"
         Tab(1).Control(32)=   "textFinish(1)"
         Tab(1).Control(33)=   "textFinish(2)"
         Tab(1).Control(34)=   "textFinish(3)"
         Tab(1).Control(35)=   "textFinish(4)"
         Tab(1).Control(36)=   "textFinish(5)"
         Tab(1).Control(37)=   "textFinish(6)"
         Tab(1).ControlCount=   38
         TabCaption(2)   =   "Rosters "
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "cmdBaseRate(9)"
         Tab(2).Control(1)=   "cmdBaseRate(8)"
         Tab(2).Control(2)=   "cmdBaseRate(7)"
         Tab(2).Control(3)=   "cmdBaseRate(6)"
         Tab(2).Control(4)=   "cmdBaseRate(5)"
         Tab(2).Control(5)=   "cmdBaseRate(4)"
         Tab(2).Control(6)=   "cmdBaseRate(3)"
         Tab(2).Control(7)=   "cmdBaseRate(2)"
         Tab(2).Control(8)=   "cmdBaseRate(1)"
         Tab(2).Control(9)=   "cmdBaseRate(0)"
         Tab(2).Control(10)=   "lbl_hr_rate"
         Tab(2).Control(11)=   "maskRosterRate(9)"
         Tab(2).Control(12)=   "maskRosterRate(8)"
         Tab(2).Control(13)=   "maskRosterRate(7)"
         Tab(2).Control(14)=   "maskRosterRate(6)"
         Tab(2).Control(15)=   "maskRosterRate(5)"
         Tab(2).Control(16)=   "maskRosterRate(4)"
         Tab(2).Control(17)=   "maskRosterRate(3)"
         Tab(2).Control(18)=   "maskRosterRate(2)"
         Tab(2).Control(19)=   "maskRosterRate(1)"
         Tab(2).Control(20)=   "maskRosterRate(0)"
         Tab(2).Control(21)=   "label_std(37)"
         Tab(2).Control(22)=   "CheckClass(9)"
         Tab(2).Control(23)=   "CheckClass(8)"
         Tab(2).Control(24)=   "CheckClass(7)"
         Tab(2).Control(25)=   "CheckClass(6)"
         Tab(2).Control(26)=   "CheckClass(5)"
         Tab(2).Control(27)=   "CheckClass(4)"
         Tab(2).Control(28)=   "CheckClass(3)"
         Tab(2).Control(29)=   "CheckClass(2)"
         Tab(2).Control(30)=   "CheckClass(1)"
         Tab(2).Control(31)=   "CheckClass(0)"
         Tab(2).ControlCount=   32
         TabCaption(3)   =   "Holidays and Notes"
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "ShapeBorder(3)"
         Tab(3).Control(1)=   "ShapeBorder(4)"
         Tab(3).Control(2)=   "cmdEndWeekMinus"
         Tab(3).Control(3)=   "cmdendWeekPlus"
         Tab(3).Control(4)=   "cmdStartWeekMinus"
         Tab(3).Control(5)=   "cmdStartWeekPlus"
         Tab(3).Control(6)=   "cmdHolEndMinus"
         Tab(3).Control(7)=   "cmdholEndPlus"
         Tab(3).Control(8)=   "cmdHolStartMinus"
         Tab(3).Control(9)=   "cmdHolStartPlus"
         Tab(3).Control(10)=   "cmdHolEndToday"
         Tab(3).Control(11)=   "cmdHolStartToday"
         Tab(3).Control(12)=   "MaskHolEnd"
         Tab(3).Control(13)=   "MaskHolStart"
         Tab(3).Control(14)=   "label_std(38)"
         Tab(3).Control(15)=   "label_std(39)"
         Tab(3).Control(16)=   "label_std(40)"
         Tab(3).Control(17)=   "label_std(41)"
         Tab(3).Control(18)=   "label_std(46)"
         Tab(3).Control(19)=   "CheckHoliday"
         Tab(3).Control(20)=   "txtHolDays"
         Tab(3).Control(21)=   "textStaffNotes"
         Tab(3).ControlCount=   22
         Begin VB.TextBox textStaffNotes 
            DataField       =   "Note"
            Height          =   2055
            Left            =   -74940
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   63
            Text            =   "FRMSET.frx":0B07
            Top             =   2640
            Width           =   5175
         End
         Begin VB.TextBox TextNote 
            DataField       =   "Note"
            Height          =   1515
            Left            =   60
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   14
            Text            =   "FRMSET.frx":0B10
            Top             =   3180
            Width           =   5175
         End
         Begin VB.TextBox TextStaffID 
            DataField       =   "StaffID"
            Height          =   315
            Left            =   1260
            MaxLength       =   10
            TabIndex        =   4
            Text            =   "StaffID"
            Top             =   360
            Width           =   1155
         End
         Begin VB.TextBox TextMiddleName 
            DataField       =   "MiddleName"
            Height          =   315
            Left            =   1260
            MaxLength       =   25
            TabIndex        =   7
            Text            =   "MiddleName"
            Top             =   1440
            Width           =   1155
         End
         Begin VB.TextBox TextFirstName 
            DataField       =   "FirstName"
            Height          =   315
            Left            =   1260
            MaxLength       =   25
            TabIndex        =   6
            Text            =   "FirstName"
            Top             =   1080
            Width           =   1155
         End
         Begin VB.TextBox TextLastName 
            DataField       =   "LastName"
            Height          =   315
            Left            =   1260
            MaxLength       =   25
            TabIndex        =   5
            Text            =   "LastName"
            Top             =   720
            Width           =   1155
         End
         Begin VB.TextBox textFinish 
            Height          =   315
            Index           =   6
            Left            =   -71085
            TabIndex        =   36
            Top             =   2970
            Width           =   1200
         End
         Begin VB.TextBox textFinish 
            Height          =   315
            Index           =   5
            Left            =   -71085
            TabIndex        =   33
            Top             =   2610
            Width           =   1200
         End
         Begin VB.TextBox textFinish 
            Height          =   315
            Index           =   4
            Left            =   -71085
            TabIndex        =   30
            Top             =   2250
            Width           =   1200
         End
         Begin VB.TextBox textFinish 
            Height          =   315
            Index           =   3
            Left            =   -71085
            TabIndex        =   27
            Top             =   1890
            Width           =   1200
         End
         Begin VB.TextBox textFinish 
            Height          =   315
            Index           =   2
            Left            =   -71085
            TabIndex        =   24
            Top             =   1530
            Width           =   1200
         End
         Begin VB.TextBox textFinish 
            Height          =   315
            Index           =   1
            Left            =   -71085
            TabIndex        =   21
            Top             =   1170
            Width           =   1200
         End
         Begin VB.TextBox textFinish 
            Height          =   315
            Index           =   0
            Left            =   -71085
            TabIndex        =   18
            Top             =   810
            Width           =   1200
         End
         Begin VB.TextBox textStart 
            Height          =   315
            Index           =   6
            Left            =   -72390
            TabIndex        =   35
            Top             =   2970
            Width           =   1200
         End
         Begin VB.TextBox textStart 
            Height          =   315
            Index           =   5
            Left            =   -72390
            TabIndex        =   32
            Top             =   2610
            Width           =   1200
         End
         Begin VB.TextBox textStart 
            Height          =   315
            Index           =   4
            Left            =   -72390
            TabIndex        =   29
            Top             =   2250
            Width           =   1200
         End
         Begin VB.TextBox textStart 
            Height          =   315
            Index           =   3
            Left            =   -72390
            TabIndex        =   26
            Top             =   1890
            Width           =   1200
         End
         Begin VB.TextBox textStart 
            Height          =   315
            Index           =   2
            Left            =   -72390
            TabIndex        =   23
            Top             =   1530
            Width           =   1200
         End
         Begin VB.TextBox textStart 
            Height          =   315
            Index           =   1
            Left            =   -72390
            TabIndex        =   20
            Top             =   1170
            Width           =   1200
         End
         Begin VB.TextBox textStart 
            Height          =   315
            Index           =   0
            Left            =   -72390
            TabIndex        =   17
            Top             =   810
            Width           =   1200
         End
         Begin VB.ComboBox comboDayStatus 
            Height          =   330
            Index           =   6
            ItemData        =   "FRMSET.frx":0B19
            Left            =   -73815
            List            =   "FRMSET.frx":0B29
            Style           =   2  'Dropdown List
            TabIndex        =   34
            Top             =   2970
            Width           =   1300
         End
         Begin VB.ComboBox comboDayStatus 
            Height          =   330
            Index           =   5
            ItemData        =   "FRMSET.frx":0B47
            Left            =   -73815
            List            =   "FRMSET.frx":0B57
            Style           =   2  'Dropdown List
            TabIndex        =   31
            Top             =   2610
            Width           =   1300
         End
         Begin VB.ComboBox comboDayStatus 
            Height          =   330
            Index           =   4
            ItemData        =   "FRMSET.frx":0B75
            Left            =   -73815
            List            =   "FRMSET.frx":0B85
            Style           =   2  'Dropdown List
            TabIndex        =   28
            Top             =   2250
            Width           =   1300
         End
         Begin VB.ComboBox comboDayStatus 
            Height          =   330
            Index           =   3
            ItemData        =   "FRMSET.frx":0BA3
            Left            =   -73815
            List            =   "FRMSET.frx":0BB3
            Style           =   2  'Dropdown List
            TabIndex        =   25
            Top             =   1890
            Width           =   1300
         End
         Begin VB.ComboBox comboDayStatus 
            Height          =   330
            Index           =   2
            ItemData        =   "FRMSET.frx":0BD1
            Left            =   -73815
            List            =   "FRMSET.frx":0BE1
            Style           =   2  'Dropdown List
            TabIndex        =   22
            Top             =   1530
            Width           =   1300
         End
         Begin VB.ComboBox comboDayStatus 
            Height          =   330
            Index           =   1
            ItemData        =   "FRMSET.frx":0BFF
            Left            =   -73815
            List            =   "FRMSET.frx":0C0F
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   1170
            Width           =   1300
         End
         Begin VB.ComboBox comboDayStatus 
            Height          =   330
            Index           =   0
            ItemData        =   "FRMSET.frx":0C2D
            Left            =   -73815
            List            =   "FRMSET.frx":0C3D
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   810
            Width           =   1300
         End
         Begin VB.TextBox txtHolDays 
            Alignment       =   2  'Center
            Height          =   315
            Left            =   -70785
            MaxLength       =   25
            TabIndex        =   62
            Top             =   810
            Width           =   500
         End
         Begin VB.OptionButton OptionPay 
            Caption         =   "Hourly"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   2520
            TabIndex        =   12
            Top             =   1970
            Value           =   -1  'True
            Width           =   1000
         End
         Begin VB.OptionButton OptionPay 
            Caption         =   "Weekly"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   3720
            TabIndex        =   13
            Top             =   1970
            Width           =   1000
         End
         Begin VB.CheckBox CheckHoliday 
            Alignment       =   1  'Right Justify
            Caption         =   "On/Off"
            Height          =   315
            Left            =   -74880
            TabIndex        =   59
            Top             =   840
            Width           =   885
         End
         Begin VB.CheckBox CheckClass 
            Caption         =   "Roster"
            Height          =   315
            Index           =   0
            Left            =   -74880
            TabIndex        =   39
            Top             =   810
            Width           =   3400
         End
         Begin VB.CheckBox CheckClass 
            Caption         =   "Roster"
            Height          =   315
            Index           =   1
            Left            =   -74880
            TabIndex        =   41
            Top             =   1140
            Width           =   3400
         End
         Begin VB.CheckBox CheckClass 
            Caption         =   "Roster"
            Height          =   315
            Index           =   2
            Left            =   -74880
            TabIndex        =   43
            Top             =   1500
            Width           =   3400
         End
         Begin VB.CheckBox CheckClass 
            Caption         =   "Roster"
            Height          =   315
            Index           =   3
            Left            =   -74880
            TabIndex        =   45
            Top             =   1860
            Width           =   3400
         End
         Begin VB.CheckBox CheckClass 
            Caption         =   "Roster"
            Height          =   315
            Index           =   4
            Left            =   -74880
            TabIndex        =   47
            Top             =   2220
            Width           =   3400
         End
         Begin VB.CheckBox CheckClass 
            Caption         =   "Roster"
            Height          =   315
            Index           =   5
            Left            =   -74880
            TabIndex        =   49
            Top             =   2580
            Width           =   3400
         End
         Begin VB.CheckBox CheckClass 
            Caption         =   "Roster"
            Height          =   315
            Index           =   6
            Left            =   -74880
            TabIndex        =   51
            Top             =   2940
            Width           =   3400
         End
         Begin VB.CheckBox CheckClass 
            Caption         =   "Roster"
            Height          =   315
            Index           =   7
            Left            =   -74880
            TabIndex        =   53
            Top             =   3300
            Width           =   3400
         End
         Begin VB.CheckBox CheckClass 
            Caption         =   "Roster"
            Height          =   315
            Index           =   8
            Left            =   -74880
            TabIndex        =   55
            Top             =   3660
            Width           =   3400
         End
         Begin VB.CheckBox CheckClass 
            Caption         =   "Roster"
            Height          =   315
            Index           =   9
            Left            =   -74880
            TabIndex        =   57
            Top             =   4020
            Width           =   3400
         End
         Begin MSMask.MaskEdBox MaskMaxHours 
            DataField       =   "HourRate"
            Height          =   315
            Left            =   -73560
            TabIndex        =   38
            Top             =   4140
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   556
            _Version        =   327681
            ClipMode        =   1
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            Format          =   "#0.00"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskMinHours 
            DataField       =   "HourRate"
            Height          =   315
            Left            =   -73560
            TabIndex        =   37
            Top             =   3780
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   556
            _Version        =   327681
            ClipMode        =   1
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            Format          =   "#0.00"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskHourRate 
            DataField       =   "HourRate"
            Height          =   315
            Left            =   1260
            TabIndex        =   11
            Top             =   1920
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            _Version        =   327681
            ClipMode        =   1
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            Format          =   "Currency"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskDateHired 
            DataField       =   "DateHired"
            Height          =   315
            Left            =   3720
            TabIndex        =   9
            Top             =   1080
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   327681
            ClipMode        =   1
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            Format          =   "Short Date"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskBirthDate 
            DataField       =   "Birthdate"
            Height          =   315
            Left            =   3720
            TabIndex        =   8
            Top             =   720
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   327681
            ClipMode        =   1
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            Format          =   "Short Date"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskHomePhone 
            DataField       =   "HomePhone"
            Height          =   315
            Left            =   3720
            TabIndex        =   10
            Top             =   1440
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   327681
            ClipMode        =   1
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            MaxLength       =   14
            Mask            =   "(##9) #####999"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox maskRosterRate 
            DataField       =   "HourRate"
            Height          =   315
            Index           =   0
            Left            =   -71400
            TabIndex        =   40
            Top             =   780
            Visible         =   0   'False
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   556
            _Version        =   327681
            ClipMode        =   1
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            Format          =   "Currency"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox maskRosterRate 
            DataField       =   "HourRate"
            Height          =   315
            Index           =   1
            Left            =   -71400
            TabIndex        =   42
            Top             =   1140
            Visible         =   0   'False
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   556
            _Version        =   327681
            ClipMode        =   1
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            Format          =   "Currency"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox maskRosterRate 
            DataField       =   "HourRate"
            Height          =   315
            Index           =   2
            Left            =   -71400
            TabIndex        =   44
            Top             =   1500
            Visible         =   0   'False
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   556
            _Version        =   327681
            ClipMode        =   1
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            Format          =   "Currency"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox maskRosterRate 
            DataField       =   "HourRate"
            Height          =   315
            Index           =   3
            Left            =   -71400
            TabIndex        =   46
            Top             =   1860
            Visible         =   0   'False
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   556
            _Version        =   327681
            ClipMode        =   1
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            Format          =   "Currency"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox maskRosterRate 
            DataField       =   "HourRate"
            Height          =   315
            Index           =   4
            Left            =   -71400
            TabIndex        =   48
            Top             =   2220
            Visible         =   0   'False
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   556
            _Version        =   327681
            ClipMode        =   1
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            Format          =   "Currency"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox maskRosterRate 
            DataField       =   "HourRate"
            Height          =   315
            Index           =   5
            Left            =   -71400
            TabIndex        =   50
            Top             =   2580
            Visible         =   0   'False
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   556
            _Version        =   327681
            ClipMode        =   1
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            Format          =   "Currency"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox maskRosterRate 
            DataField       =   "HourRate"
            Height          =   315
            Index           =   6
            Left            =   -71400
            TabIndex        =   52
            Top             =   2940
            Visible         =   0   'False
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   556
            _Version        =   327681
            ClipMode        =   1
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            Format          =   "Currency"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox maskRosterRate 
            DataField       =   "HourRate"
            Height          =   315
            Index           =   7
            Left            =   -71400
            TabIndex        =   54
            Top             =   3300
            Visible         =   0   'False
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   556
            _Version        =   327681
            ClipMode        =   1
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            Format          =   "Currency"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox maskRosterRate 
            DataField       =   "HourRate"
            Height          =   315
            Index           =   8
            Left            =   -71400
            TabIndex        =   56
            Top             =   3660
            Visible         =   0   'False
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   556
            _Version        =   327681
            ClipMode        =   1
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            Format          =   "Currency"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox maskRosterRate 
            DataField       =   "HourRate"
            Height          =   315
            Index           =   9
            Left            =   -71400
            TabIndex        =   58
            Top             =   4020
            Visible         =   0   'False
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   556
            _Version        =   327681
            ClipMode        =   1
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            Format          =   "Currency"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskHolStart 
            DataField       =   "DateHired"
            Height          =   315
            Left            =   -73890
            TabIndex        =   60
            Top             =   810
            Width           =   1470
            _ExtentX        =   2593
            _ExtentY        =   556
            _Version        =   327681
            ClipMode        =   1
            BackColor       =   16777215
            ForeColor       =   0
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            Enabled         =   0   'False
            Format          =   "Short Date"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskHolEnd 
            DataField       =   "DateHired"
            Height          =   315
            Left            =   -72330
            TabIndex        =   61
            Top             =   810
            Width           =   1470
            _ExtentX        =   2593
            _ExtentY        =   556
            _Version        =   327681
            ClipMode        =   1
            BackColor       =   16777215
            ForeColor       =   0
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            Enabled         =   0   'False
            Format          =   "Short Date"
            PromptChar      =   "_"
         End
         Begin VB.Label label_std 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000D&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Staff Notes"
            ForeColor       =   &H8000000E&
            Height          =   315
            Index           =   46
            Left            =   -74940
            TabIndex        =   154
            Top             =   2280
            Width           =   5175
         End
         Begin VB.Label labelHeading 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "< [ Staff ID must be unique identifier ]"
            Height          =   255
            Index           =   0
            Left            =   2520
            TabIndex        =   120
            Top             =   420
            Width           =   2655
         End
         Begin VB.Label label_std 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000002&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "End Date"
            ForeColor       =   &H80000009&
            Height          =   315
            Index           =   41
            Left            =   -72330
            TabIndex        =   119
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
            Height          =   315
            Index           =   40
            Left            =   -73890
            TabIndex        =   118
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
            Height          =   315
            Index           =   39
            Left            =   -70785
            TabIndex        =   117
            Top             =   420
            Width           =   495
         End
         Begin VB.Label label_std 
            AutoSize        =   -1  'True
            BackColor       =   &H80000004&
            BackStyle       =   0  'Transparent
            Caption         =   $"FRMSET.frx":0C5B
            Height          =   570
            Index           =   38
            Left            =   -74880
            TabIndex        =   116
            Top             =   1620
            Width           =   5070
            WordWrap        =   -1  'True
         End
         Begin VB.Label label_std 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000002&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Roster Availability"
            ForeColor       =   &H80000009&
            Height          =   315
            Index           =   37
            Left            =   -74880
            TabIndex        =   115
            Top             =   420
            Width           =   3400
         End
         Begin VB.Label label_std 
            AutoSize        =   -1  'True
            BackColor       =   &H80000004&
            BackStyle       =   0  'Transparent
            Caption         =   "Minimum && maximum hours per week required by this staff member."
            Height          =   750
            Index           =   36
            Left            =   -72390
            TabIndex        =   114
            Top             =   3780
            Width           =   2580
            WordWrap        =   -1  'True
         End
         Begin VB.Label label_std 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000D&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Hours Required per Week"
            ForeColor       =   &H8000000E&
            Height          =   315
            Index           =   35
            Left            =   -74880
            TabIndex        =   113
            Top             =   3420
            Width           =   5055
         End
         Begin VB.Label label_std 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000002&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Maximum"
            ForeColor       =   &H80000009&
            Height          =   315
            Index           =   34
            Left            =   -74880
            TabIndex        =   112
            Top             =   4140
            Width           =   1275
         End
         Begin VB.Label label_std 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000002&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Minimum"
            ForeColor       =   &H80000009&
            Height          =   315
            Index           =   33
            Left            =   -74880
            TabIndex        =   111
            Top             =   3780
            Width           =   1275
         End
         Begin VB.Label label_std 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000002&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Day"
            ForeColor       =   &H80000009&
            Height          =   315
            Index           =   32
            Left            =   -74880
            TabIndex        =   110
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
            Height          =   315
            Index           =   31
            Left            =   -73800
            TabIndex        =   109
            Top             =   420
            Width           =   1300
         End
         Begin VB.Label label_std 
            AutoSize        =   -1  'True
            BackColor       =   &H80000004&
            BackStyle       =   0  'Transparent
            Caption         =   "Minimum && maximum hours per week required by this staff member."
            Height          =   630
            Index           =   11
            Left            =   -72720
            TabIndex        =   108
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
            Height          =   315
            Index           =   4
            Left            =   2520
            TabIndex        =   107
            Top             =   720
            Width           =   1065
         End
         Begin VB.Label label_std 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000002&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Employed"
            ForeColor       =   &H80000009&
            Height          =   315
            Index           =   5
            Left            =   2520
            TabIndex        =   106
            Top             =   1080
            Width           =   1065
         End
         Begin VB.Label label_std 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000002&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Base Rate"
            ForeColor       =   &H80000009&
            Height          =   315
            Index           =   6
            Left            =   120
            TabIndex        =   105
            Top             =   1920
            Width           =   1065
         End
         Begin VB.Label label_std 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000002&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Home Phone"
            ForeColor       =   &H80000009&
            Height          =   315
            Index           =   7
            Left            =   2520
            TabIndex        =   104
            Top             =   1440
            Width           =   1065
         End
         Begin VB.Label label_std 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000D&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Staff ID"
            ForeColor       =   &H8000000E&
            Height          =   315
            Index           =   0
            Left            =   120
            TabIndex        =   103
            Top             =   360
            Width           =   1065
         End
         Begin VB.Label label_std 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000002&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Middle Name"
            ForeColor       =   &H80000009&
            Height          =   315
            Index           =   3
            Left            =   120
            TabIndex        =   102
            Top             =   1440
            Width           =   1065
         End
         Begin VB.Label label_std 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000002&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "First Name"
            ForeColor       =   &H80000009&
            Height          =   315
            Index           =   2
            Left            =   120
            TabIndex        =   101
            Top             =   1080
            Width           =   1065
         End
         Begin VB.Label label_std 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000002&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Last Name"
            ForeColor       =   &H80000009&
            Height          =   315
            Index           =   1
            Left            =   120
            TabIndex        =   100
            Top             =   720
            Width           =   1065
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
            TabIndex        =   99
            Top             =   3240
            Width           =   4360
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
            TabIndex        =   98
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
            Height          =   315
            Left            =   -71400
            TabIndex        =   97
            Top             =   420
            Width           =   1200
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
            TabIndex        =   96
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
            TabIndex        =   95
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
            TabIndex        =   94
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
            Height          =   315
            Index           =   2
            Left            =   -71085
            TabIndex        =   93
            Top             =   420
            Visible         =   0   'False
            Width           =   1200
         End
         Begin VB.Label lbl_std 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000002&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Start Time"
            ForeColor       =   &H80000009&
            Height          =   315
            Index           =   1
            Left            =   -72390
            TabIndex        =   92
            Top             =   420
            Visible         =   0   'False
            Width           =   1200
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
            TabIndex        =   91
            Top             =   420
            Width           =   1155
         End
         Begin VB.Label labelDay 
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "Day Name"
            Height          =   210
            Index           =   6
            Left            =   -74850
            TabIndex        =   90
            Top             =   3045
            Width           =   990
         End
         Begin VB.Label labelDay 
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "Day Name"
            Height          =   210
            Index           =   5
            Left            =   -74850
            TabIndex        =   89
            Top             =   2685
            Width           =   990
         End
         Begin VB.Label labelDay 
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "Day Name"
            Height          =   210
            Index           =   4
            Left            =   -74850
            TabIndex        =   88
            Top             =   2325
            Width           =   990
         End
         Begin VB.Label labelDay 
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "Day Name"
            Height          =   210
            Index           =   3
            Left            =   -74850
            TabIndex        =   87
            Top             =   1965
            Width           =   990
         End
         Begin VB.Label labelDay 
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "Day Name"
            Height          =   210
            Index           =   2
            Left            =   -74850
            TabIndex        =   86
            Top             =   1605
            Width           =   990
         End
         Begin VB.Label labelDay 
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "Day Name"
            Height          =   210
            Index           =   1
            Left            =   -74850
            TabIndex        =   85
            Top             =   1260
            Width           =   990
         End
         Begin VB.Label labelDay 
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "Day Name"
            Height          =   210
            Index           =   0
            Left            =   -74850
            TabIndex        =   84
            Top             =   900
            Width           =   990
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
            TabIndex        =   83
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
            TabIndex        =   82
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
            TabIndex        =   81
            Top             =   420
            Width           =   495
         End
         Begin VB.Image cmdHolStartToday 
            Height          =   300
            Left            =   -73290
            Picture         =   "FRMSET.frx":0CE5
            Top             =   1200
            Width           =   300
         End
         Begin VB.Image cmdHolEndToday 
            Height          =   300
            Left            =   -71730
            Picture         =   "FRMSET.frx":0E27
            Top             =   1200
            Width           =   300
         End
         Begin VB.Label label_std 
            AutoSize        =   -1  'True
            BackColor       =   &H80000004&
            BackStyle       =   0  'Transparent
            Caption         =   $"FRMSET.frx":0F69
            Height          =   2550
            Index           =   20
            Left            =   -74880
            TabIndex        =   80
            Top             =   1440
            Width           =   4440
            WordWrap        =   -1  'True
         End
         Begin VB.Image cmdHolStartPlus 
            Height          =   300
            Left            =   -72990
            Picture         =   "FRMSET.frx":0FF3
            Top             =   1200
            Width           =   300
         End
         Begin VB.Image cmdHolStartMinus 
            Height          =   300
            Left            =   -73590
            Picture         =   "FRMSET.frx":1135
            Top             =   1200
            Width           =   300
         End
         Begin VB.Image cmdholEndPlus 
            Height          =   300
            Left            =   -71430
            Picture         =   "FRMSET.frx":1277
            Top             =   1200
            Width           =   300
         End
         Begin VB.Image cmdHolEndMinus 
            Height          =   300
            Left            =   -72030
            Picture         =   "FRMSET.frx":13B9
            Top             =   1200
            Width           =   300
         End
         Begin VB.Image cmdStartWeekPlus 
            Height          =   300
            Left            =   -72690
            Picture         =   "FRMSET.frx":14FB
            Top             =   1200
            Width           =   300
         End
         Begin VB.Image cmdStartWeekMinus 
            Height          =   300
            Left            =   -73890
            Picture         =   "FRMSET.frx":164D
            Top             =   1200
            Width           =   300
         End
         Begin VB.Image cmdendWeekPlus 
            Height          =   300
            Left            =   -71130
            Picture         =   "FRMSET.frx":179F
            Top             =   1200
            Width           =   300
         End
         Begin VB.Image cmdEndWeekMinus 
            Height          =   300
            Left            =   -72330
            Picture         =   "FRMSET.frx":18F1
            Top             =   1200
            Width           =   300
         End
         Begin VB.Label labelInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Staff Information"
            Height          =   735
            Left            =   120
            TabIndex        =   79
            Top             =   2340
            Width           =   5055
            WordWrap        =   -1  'True
         End
         Begin VB.Image cmdBaseRate 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Index           =   0
            Left            =   -70110
            Picture         =   "FRMSET.frx":1A43
            Top             =   780
            Visible         =   0   'False
            Width           =   330
         End
         Begin VB.Image cmdBaseRate 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Index           =   1
            Left            =   -70110
            Picture         =   "FRMSET.frx":1B95
            Top             =   1140
            Visible         =   0   'False
            Width           =   330
         End
         Begin VB.Image cmdBaseRate 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Index           =   2
            Left            =   -70110
            Picture         =   "FRMSET.frx":1CE7
            Top             =   1500
            Visible         =   0   'False
            Width           =   330
         End
         Begin VB.Image cmdBaseRate 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Index           =   3
            Left            =   -70110
            Picture         =   "FRMSET.frx":1E39
            Top             =   1860
            Visible         =   0   'False
            Width           =   330
         End
         Begin VB.Image cmdBaseRate 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Index           =   4
            Left            =   -70110
            Picture         =   "FRMSET.frx":1F8B
            Top             =   2220
            Visible         =   0   'False
            Width           =   330
         End
         Begin VB.Image cmdBaseRate 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Index           =   5
            Left            =   -70110
            Picture         =   "FRMSET.frx":20DD
            Top             =   2580
            Visible         =   0   'False
            Width           =   330
         End
         Begin VB.Image cmdBaseRate 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Index           =   6
            Left            =   -70110
            Picture         =   "FRMSET.frx":222F
            Top             =   2940
            Visible         =   0   'False
            Width           =   330
         End
         Begin VB.Image cmdBaseRate 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Index           =   7
            Left            =   -70110
            Picture         =   "FRMSET.frx":2381
            Top             =   3300
            Visible         =   0   'False
            Width           =   330
         End
         Begin VB.Image cmdBaseRate 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Index           =   8
            Left            =   -70110
            Picture         =   "FRMSET.frx":24D3
            Top             =   3660
            Visible         =   0   'False
            Width           =   330
         End
         Begin VB.Image cmdBaseRate 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Index           =   9
            Left            =   -70110
            Picture         =   "FRMSET.frx":2625
            Top             =   4020
            Visible         =   0   'False
            Width           =   330
         End
         Begin VB.Shape ShapeBorder 
            BackColor       =   &H00E0E0E0&
            BorderColor     =   &H80000012&
            BorderWidth     =   2
            Height          =   1275
            Index           =   0
            Left            =   60
            Top             =   1860
            Width           =   5175
         End
         Begin VB.Shape ShapeBorder 
            BackColor       =   &H00C0C0C0&
            BackStyle       =   1  'Opaque
            BorderWidth     =   2
            Height          =   315
            Index           =   4
            Left            =   -72330
            Top             =   1200
            Width           =   1515
         End
         Begin VB.Shape ShapeBorder 
            BackColor       =   &H00C0C0C0&
            BackStyle       =   1  'Opaque
            BorderWidth     =   2
            Height          =   315
            Index           =   3
            Left            =   -73890
            Top             =   1200
            Width           =   1515
         End
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   1455
         Left            =   -74880
         TabIndex        =   145
         Top             =   1560
         Width           =   2775
         _Version        =   65536
         _ExtentX        =   4895
         _ExtentY        =   2566
         _StockProps     =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         BorderWidth     =   0
         BevelOuter      =   0
         Begin VB.ComboBox ComboIncrement 
            Height          =   315
            ItemData        =   "FRMSET.frx":2777
            Left            =   840
            List            =   "FRMSET.frx":27AA
            Style           =   2  'Dropdown List
            TabIndex        =   150
            Top             =   720
            Width           =   1875
         End
         Begin VB.ComboBox ComboHour 
            Height          =   330
            ItemData        =   "FRMSET.frx":282C
            Left            =   840
            List            =   "FRMSET.frx":2886
            TabIndex        =   149
            Text            =   "ComboHour"
            Top             =   360
            Width           =   900
         End
         Begin VB.ComboBox ComboMinute 
            Height          =   330
            ItemData        =   "FRMSET.frx":28EA
            Left            =   1800
            List            =   "FRMSET.frx":28FA
            TabIndex        =   148
            Text            =   "ComboMinute"
            Top             =   360
            Width           =   900
         End
         Begin VB.OptionButton OptionTime 
            Caption         =   "Start"
            Height          =   315
            Index           =   0
            Left            =   0
            TabIndex        =   147
            Top             =   0
            Value           =   -1  'True
            Width           =   795
         End
         Begin VB.OptionButton OptionTime 
            Caption         =   "End"
            Height          =   315
            Index           =   1
            Left            =   0
            TabIndex        =   146
            Top             =   360
            Width           =   735
         End
         Begin MSMask.MaskEdBox MaskDayLength 
            Height          =   315
            Left            =   840
            TabIndex        =   153
            TabStop         =   0   'False
            Top             =   1080
            Width           =   1875
            _ExtentX        =   3307
            _ExtentY        =   556
            _Version        =   327681
            BackColor       =   -2147483644
            ForeColor       =   -2147483646
            Enabled         =   0   'False
            MaxLength       =   5
            Format          =   "hh:mm"
            PromptChar      =   "_"
         End
         Begin VB.Label label_std 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000D&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Hours"
            ForeColor       =   &H8000000E&
            Height          =   315
            Index           =   42
            Left            =   840
            TabIndex        =   152
            Top             =   0
            Width           =   900
         End
         Begin VB.Label label_std 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000D&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Minutes"
            ForeColor       =   &H8000000E&
            Height          =   315
            Index           =   43
            Left            =   1800
            TabIndex        =   151
            Top             =   0
            Width           =   900
         End
      End
      Begin Threed.SSPanel panelStaffToolbar 
         Height          =   375
         Left            =   -74940
         TabIndex        =   155
         Top             =   360
         Width           =   1935
         _Version        =   65536
         _ExtentX        =   3413
         _ExtentY        =   661
         _StockProps     =   15
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
         Begin VB.Image cmdAllStaffRosters 
            Height          =   300
            Left            =   1260
            Picture         =   "FRMSET.frx":290E
            Top             =   60
            Width           =   300
         End
         Begin VB.Image cmdSelectedStaffRoster 
            Height          =   300
            Left            =   960
            Picture         =   "FRMSET.frx":2B60
            Top             =   30
            Width           =   300
         End
         Begin VB.Image cmdNew 
            Height          =   300
            Left            =   660
            Picture         =   "FRMSET.frx":2CD2
            Top             =   30
            Width           =   300
         End
         Begin VB.Image cmdSave 
            Height          =   300
            Left            =   60
            Picture         =   "FRMSET.frx":2E14
            Top             =   30
            Visible         =   0   'False
            Width           =   300
         End
         Begin VB.Image cmdDelete 
            Height          =   300
            Left            =   390
            Picture         =   "FRMSET.frx":2F66
            Top             =   30
            Width           =   300
         End
      End
      Begin Threed.SSCommand cmdNukeCover 
         Height          =   570
         Left            =   -72720
         TabIndex        =   77
         Top             =   4620
         Width           =   570
         _Version        =   65536
         _ExtentX        =   1005
         _ExtentY        =   1005
         _StockProps     =   78
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Outline         =   0   'False
         AutoSize        =   2
         Picture         =   "FRMSET.frx":30B8
      End
      Begin Threed.SSCommand cmdPrintSet 
         Height          =   570
         Left            =   -72720
         TabIndex        =   76
         Top             =   4020
         Width           =   570
         _Version        =   65536
         _ExtentX        =   1005
         _ExtentY        =   1005
         _StockProps     =   78
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Outline         =   0   'False
         AutoSize        =   2
         Picture         =   "FRMSET.frx":3992
      End
      Begin MSMask.MaskEdBox maskDateMask 
         Height          =   315
         Left            =   -74580
         TabIndex        =   75
         Top             =   3240
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   556
         _Version        =   327681
         PromptChar      =   "_"
      End
      Begin Threed.SSCommand cmdNuke 
         Height          =   570
         Left            =   -72720
         TabIndex        =   78
         Top             =   4620
         Width           =   570
         _Version        =   65536
         _ExtentX        =   1005
         _ExtentY        =   1005
         _StockProps     =   78
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Outline         =   0   'False
         AutoSize        =   2
         Picture         =   "FRMSET.frx":426C
      End
      Begin VB.Label labelHeading 
         Appearance      =   0  'Flat
         BackColor       =   &H80000002&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Autobuild Settings"
         ForeColor       =   &H80000009&
         Height          =   315
         Index           =   2
         Left            =   -74880
         TabIndex        =   122
         Top             =   1200
         Width           =   7155
      End
      Begin VB.Label label_std 
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         BackStyle       =   0  'Transparent
         Caption         =   "Click to change the settings for the currently selected printer."
         Height          =   420
         Index           =   45
         Left            =   -72000
         TabIndex        =   139
         Top             =   4080
         Width           =   4215
         WordWrap        =   -1  'True
      End
      Begin VB.Label label_std 
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         BackStyle       =   0  'Transparent
         Caption         =   "Click to expose the NUKE command button.  The NUKE command will delete ALL roster shifts from every roster."
         Height          =   420
         Index           =   29
         Left            =   -72000
         TabIndex        =   138
         Top             =   4680
         Width           =   4215
         WordWrap        =   -1  'True
      End
      Begin VB.Label label_std 
         Appearance      =   0  'Flat
         BackColor       =   &H80000002&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Printer Settings / Delete All Roster Shifts"
         ForeColor       =   &H80000009&
         Height          =   315
         Index           =   30
         Left            =   -74880
         TabIndex        =   137
         Top             =   3660
         Width           =   7155
         WordWrap        =   -1  'True
      End
      Begin VB.Label labelHeading 
         Appearance      =   0  'Flat
         BackColor       =   &H80000002&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Displayed Date Format"
         ForeColor       =   &H80000009&
         Height          =   315
         Index           =   1
         Left            =   -74880
         TabIndex        =   136
         Top             =   1740
         Width           =   7155
      End
      Begin VB.Label label_std 
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         BackStyle       =   0  'Transparent
         Caption         =   $"FRMSET.frx":4586
         Height          =   1050
         Index           =   28
         Left            =   -72000
         TabIndex        =   135
         Top             =   2100
         Width           =   4170
         WordWrap        =   -1  'True
      End
      Begin VB.Label label_std 
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         BackStyle       =   0  'Transparent
         Caption         =   $"FRMSET.frx":4678
         Height          =   840
         Index           =   27
         Left            =   -72000
         TabIndex        =   134
         Top             =   780
         Width           =   4200
         WordWrap        =   -1  'True
      End
      Begin VB.Label labelHeading 
         Appearance      =   0  'Flat
         BackColor       =   &H80000002&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Feedback"
         ForeColor       =   &H80000009&
         Height          =   315
         Index           =   3
         Left            =   -74880
         TabIndex        =   133
         Top             =   420
         Width           =   7155
      End
      Begin VB.Label label_std 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         BackStyle       =   0  'Transparent
         Caption         =   "Select the starting day of your roster week."
         Height          =   210
         Index           =   44
         Left            =   -71940
         TabIndex        =   132
         Top             =   840
         Width           =   4140
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblRosterSettingsInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "RosterSettings Help"
         Height          =   555
         Left            =   -74880
         TabIndex        =   131
         Top             =   4620
         Width           =   7155
      End
      Begin VB.Label label_std 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         BackStyle       =   0  'Transparent
         Caption         =   "Select the length of time a staff member has to work before a compulsory break is applied."
         Height          =   420
         Index           =   23
         Left            =   -71940
         TabIndex        =   130
         Top             =   3720
         Width           =   4140
         WordWrap        =   -1  'True
      End
      Begin VB.Label labelHeading 
         Appearance      =   0  'Flat
         BackColor       =   &H80000002&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Staff Breaks"
         ForeColor       =   &H80000009&
         Height          =   315
         Index           =   5
         Left            =   -74880
         TabIndex        =   129
         Top             =   3060
         Width           =   7155
      End
      Begin VB.Label label_std 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         BackStyle       =   0  'Transparent
         Caption         =   "Select the length of the break to be applied."
         Height          =   210
         Index           =   24
         Left            =   -71940
         TabIndex        =   128
         Top             =   4200
         Width           =   4140
         WordWrap        =   -1  'True
      End
      Begin VB.Label label_std 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Minutes"
         ForeColor       =   &H8000000E&
         Height          =   315
         Index           =   25
         Left            =   -73080
         TabIndex        =   127
         Top             =   3420
         Width           =   900
      End
      Begin VB.Label label_std 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Hours"
         ForeColor       =   &H8000000E&
         Height          =   315
         Index           =   26
         Left            =   -74040
         TabIndex        =   126
         Top             =   3420
         Width           =   900
      End
      Begin VB.Label label_std 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         BackStyle       =   0  'Transparent
         Caption         =   "Select the start and finish times for your working day."
         Height          =   210
         Index           =   19
         Left            =   -71940
         TabIndex        =   125
         Top             =   1980
         Width           =   4140
         WordWrap        =   -1  'True
      End
      Begin VB.Label label_std 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         BackStyle       =   0  'Transparent
         Caption         =   "This is the length of your working day (hh:mm)"
         Height          =   210
         Index           =   21
         Left            =   -71940
         TabIndex        =   124
         Top             =   2700
         Width           =   4140
         WordWrap        =   -1  'True
      End
      Begin VB.Label label_std 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         BackStyle       =   0  'Transparent
         Caption         =   "Select the standard length of each roster shift."
         Height          =   210
         Index           =   22
         Left            =   -71940
         TabIndex        =   123
         Top             =   2340
         Width           =   4140
         WordWrap        =   -1  'True
      End
      Begin VB.Label label_std 
         Appearance      =   0  'Flat
         BackColor       =   &H80000002&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Starting Day of Roster Week"
         ForeColor       =   &H80000009&
         Height          =   315
         Index           =   18
         Left            =   -74880
         TabIndex        =   121
         Top             =   420
         Width           =   7155
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblRosterDetailsInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Roster Details Help"
         Height          =   1335
         Left            =   120
         TabIndex        =   1
         Top             =   3900
         Width           =   7155
      End
      Begin VB.Shape ShapeBorder 
         BackColor       =   &H00E0E0E0&
         BorderColor     =   &H80000012&
         BorderWidth     =   2
         Height          =   4755
         Index           =   1
         Left            =   -72060
         Top             =   480
         Width           =   4335
      End
      Begin VB.Shape ShapeBorder 
         BackColor       =   &H00E0E0E0&
         BorderColor     =   &H80000012&
         BorderWidth     =   2
         Height          =   4095
         Index           =   2
         Left            =   -72045
         Top             =   480
         Width           =   4320
      End
   End
   Begin VB.Image ImageSwitch 
      Height          =   240
      Index           =   0
      Left            =   0
      Picture         =   "FRMSET.frx":4752
      Top             =   4080
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image ImageSwitch 
      Height          =   240
      Index           =   1
      Left            =   0
      Picture         =   "FRMSET.frx":4854
      Top             =   4320
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image ImageSwitch 
      Height          =   240
      Index           =   2
      Left            =   0
      Picture         =   "FRMSET.frx":4956
      Top             =   4560
      Visible         =   0   'False
      Width           =   240
   End
End
Attribute VB_Name = "frmSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
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
Dim intStartDay         As Integer      '[Start time - week day]
Dim intFinishDay As Integer             '[Finish time - week day]

Sub SetHelpLabels()
    
    '[SET HELP LABELS ON FORM]
    lblRosterDetailsInfo.Caption = "Double-click on a grid cell to modify the data contained within."
    lblRosterDetailsInfo.Caption = lblRosterDetailsInfo.Caption & Chr$(vbKeyReturn) & "Col 1: mark rosters as active (green light) or inactive (red light)."
    lblRosterDetailsInfo.Caption = lblRosterDetailsInfo.Caption & Chr$(vbKeyReturn) & "Col 2: specify a short 'code' for each roster (used when saving the roster to a disk file)."
    lblRosterDetailsInfo.Caption = lblRosterDetailsInfo.Caption & Chr$(vbKeyReturn) & "Col 3: set the full name (or 'description') of the roster."
    lblRosterDetailsInfo.Caption = lblRosterDetailsInfo.Caption & Chr$(vbKeyReturn) & "Col 4: contains any notes which you would like printed on the weekly roster printouts."
    '[SET HELP LABELS ON FORM]
    lblRosterSettingsInfo.Caption = "GSR will use the 'autobuild' settings to create a blank roster for you. The 'Staff Breaks' section will be used when determining roster costs, staff work hours and projected staff costs."
    
End Sub

Sub StatusBar(strMessage)

    '[PLACE PASSED MESSAGE ON THE STATUS BAR]
    If frmSet.panelStatusBar.Caption <> " " + strMessage Then frmSet.panelStatusBar.Caption = " " + strMessage

End Sub



Private Sub cmdClose_Click(Index As Integer)
    
    '[HIDE CONTROL FORM]
    frmSet.Hide

End Sub

Private Sub cmdClose_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
        
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Close the Control Panel and return to GSR."

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

















Private Sub CheckClass_Click(Index As Integer)

    '[MAKE SAVE BUTTON VISIBLE]
    If frmSet.cmdSave.Visible = False Then frmSet.cmdSave.Visible = True
    
    '[MAKE RATE VISIBLE IF HOURLY PAY RATE IS SET AND CLASS IS CHECKED]
    If frmSet.CheckClass(Index).Value = vbChecked And frmSet.OptionPay(vbHourly).Value = True Then
        frmSet.maskRosterRate(Index).Visible = True
        frmSet.cmdBaseRate(Index).Visible = True
        frmSet.lbl_hr_rate.Visible = True
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

    '[ANIMATE BUTTON]
    frmSet.cmdBaseRate(Index).BorderStyle = 1
    Delay vbDelay
    frmSet.cmdBaseRate(Index).BorderStyle = 0
    Delay vbDelay
    frmSet.cmdBaseRate(Index).BorderStyle = 1

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
    Dim strDisplayname          As String
    Dim strDeleteName           As String
    Dim intColCounter           As Integer
    Dim intDayNum               As Integer
    Dim strDayNum               As String
    Dim strLeft                 As String
    Dim strRight                As String
    Dim intLeftPos              As Integer
    Dim SQLStmt                 As String
    
    '[POPUP YES/NO DIALOG TO CONFIRM DELETION]
    Dim Msg As String
    Dim Style
    Dim Response
    Dim Title
    
    If frmSet.ListStaff.ListIndex = -1 Then Exit Sub
    '{SET DELETED NAME TEMP VARIABLE]
    strDeleteName = DsStaff("LastName") & ", " & DsStaff("FirstName")
    
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
    
        '[ASK TO DELETE FROM OTHER ROSTERS]
        Msg = "Do you wish to remove each occurence of the staff member (" & strDeleteName & ") from your rosters ?" & strBreak & strBreak & "If you choose 'Yes', you may need to refresh your currently displayed roster by reselecting it on the roster form."
        Style = vbYesNo ' Define buttons.
        Title = "Delete Staff From Rosters ?"  ' Define title.
        Response = gsrMsg(Msg, Style, Title)
        If Response = vbYes Then
            '[CYCLE THROUGH ALL ROSTERS AND REMOVE NAME FROM EACH CELL]
            '[SELECT APPROPRIATE RECORDS]
            SQLStmt = "SELECT * FROM Roster ORDER BY [Class], [ShiftStart]"
            Set DsReport = DBMain.OpenRecordset(SQLStmt, dbOpenDynaset)
            If Not (DsReport.EOF And DsReport.BOF) Then
                DsReport.MoveFirst
                Do While Not DsReport.EOF
                    For intDayNum = 1 To 7
                        strDayNum = "Day_" & Trim(Str$(intDayNum))
                        If InStr(DsReport(strDayNum), strDeleteName) > 0 Then
                            '[NAME IS IN ROSTER - REMOVE FROM ROSTER]
                            '[FIND STRING LEFT OF PASSED NAME]
                            intLeftPos = InStr(DsReport(strDayNum), strDeleteName)
                            If intLeftPos = 1 Then
                                strLeft = ""
                            Else
                                strLeft = Left(DsReport(strDayNum), intLeftPos - 2)
                            End If
                            '[FIND STRING RIGHT OF PASSED NAME]
                            strRight = Trim(Mid$(DsReport(strDayNum), Len(strLeft) + Len(strDeleteName) + 2))
                            '[APPLY NEW CONCATENATED STRING TO CELL]
                            DsReport.Edit
                                DsReport(strDayNum) = (strLeft & strRight) & " "
                            DsReport.Update
                        End If
                    Next intDayNum
                    DsReport.MoveNext
                Loop
            End If
            DsReport.Close
        End If
    
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
    flagLoad = True     '[ACTIOVATE LOADED FORM FLAG]
    
    '[SET ROSTER FRAME FIRST]
    frmSet.tabSet.Tab = 0
    '[DEBUG]
    Call LogToFile("- filling staff list")
    '[FILL STAFF LIST FROM DYNASET]
    FillStaffList
    '[DEBUG]
    '[APPLY CLASS LABELS]
    SetClassLabels
    '[DEBUG]
    Call LogToFile("- setting day labels")
    '[APPLY DAY LABELS]
    SetDayLabels
    '[DEBUG]
    Call LogToFile("- displaying staff details")
    '[LOCATE FIRST STAFF MEMBER]
    If frmSet.ListStaff.ListCount > 0 Then frmSet.ListStaff.ListIndex = 0
      '[DEBUG]
    Call LogToFile("=-CONTROL SECT - LOAD SUB----------------=")
    '[DEBUG]
    Call LogToFile("- filling class grid")
    '[CALL ROUTINE TO FILL GRID]
    FillClassGrid
    '[Resize controls and grids to match]
    frmSet.GridClass.ColWidth(0) = frmSet.ImageSwitch(0).Width
    frmSet.GridClass.ColWidth(1) = (GridClass.Width - frmSet.GridClass.ColWidth(0)) * 0.1
    frmSet.GridClass.ColWidth(2) = (GridClass.Width - frmSet.GridClass.ColWidth(0)) * 0.3
    frmSet.GridClass.ColWidth(3) = (GridClass.Width - frmSet.GridClass.ColWidth(0)) * 0.55
  
    '[Resize class grid height]
    'frmSet.GridClass.Height = (frmSet.GridClass.Rows + 1) * frmSet.GridClass.RowHeight(1)
        
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
    frmSet.OptionDateFormat(0).Caption = Format(Date, strDateFormat)
    frmSet.OptionDateFormat(1).Caption = Format(Date, "Medium Date")
    '[PLACE CUSTOM DATE MASK IN CONTROL]
    frmSet.maskDateMask = DsDefault("CustomDate")
 
    '[------------------------------------------------------------------------------------------]
    '[SET CONTROL FORM DATE TYPE]
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    '[DEBUG]
    Call LogToFile("- setting date display format")
    '[---------------------------------------------------------------------------------]
    frmSet.OptionDateFormat(DsDefault("DateFormat").Value).Value = True
    '[DEBUG]
    Call LogToFile("= " & Format(Date$, strDateFormat))
    
    '[RESTORE DELETE CONFIRM AND SOUND FLAGS]
    '[DELETE IS 1]
    
    '[SOUNDS IS 2]
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    '[DEBUG]
    Call LogToFile("- restoring default options")
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
    
    '[VER: 3.00.32]
    Call SetHelpLabels
    
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
        StatusBar "Base hourly or weekly pay rate for this staff member."
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



Private Sub lblRosterDetailsInfo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Information on the use of this form is displayed here."
    '[---------------------------------------------------------------------------------]

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



Private Sub MaskHourRate_GotFocus()
        
        '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
        Call StatusBar("Base hourly or weekly pay rate for this staff member.")
        '[---------------------------------------------------------------------------------]
        
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
        If frmSet.OptionPay(vbHourly).Value = True Then
            If frmSet.CheckClass(intCounter).Value = vbChecked Then
                frmSet.maskRosterRate(intCounter).Visible = True
                frmSet.cmdBaseRate(intCounter).Visible = True
                frmSet.lbl_hr_rate.Visible = True
            End If
        Else
            frmSet.maskRosterRate(intCounter).Visible = False
            frmSet.cmdBaseRate(intCounter).Visible = False
        End If
    Next intCounter
    
    

End Sub



Private Sub optionPay_GotFocus(Index As Integer)
        
        Select Case Index
        Case 0          '[HOURLY PAY RATE]
            '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
            Call StatusBar("Set the base pay rate to be an hourly pay amount.")
            '[---------------------------------------------------------------------------------]
        Case 1          '[WEEKLY PAY RATE]
            '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
            Call StatusBar("Set the base pay rate to be a weekly pay amount.")
            '[---------------------------------------------------------------------------------]
        End Select
        
End Sub


Private Sub tabSet_Click(PreviousTab As Integer)

    Select Case tabSet.Tab
    Case 4
        tabSet.Tab = PreviousTab
        '[HIDE CONTROL FORM]
        frmSet.Hide
    Case Else
    End Select

End Sub

Private Sub tabSet_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Use the GSR control form to change Staff, Roster and Program settings."
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
    StatusBar "Enter roster notes for this staff member - these will be printed on their weekly timesheet."
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


Private Sub textStaffNotes_Change()

    '[ENABLE SAVE BUTTON IF A CHANGE HAS BEEN MADE]
    If frmSet.cmdSave.Visible = False Then frmSet.cmdSave.Visible = True

End Sub

Private Sub textStaffNotes_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Enter private notes about this staff member - these will be printed on the detailed staff list."
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
        StatusBar "Enter notes which will be printed with the weekly roster."
    ElseIf X > frmSet.GridClass.ColPos(2) Then
        StatusBar "Modify the roster description, a 20 character identifier for the selected roster."
    ElseIf X > frmSet.GridClass.ColPos(1) Then
        StatusBar "Modify the roster identifier, a three character short form for the roster name."
    Else
        StatusBar "Activate/deactivate the selected roster.  Only active rosters can be modified and processed."
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
        
        '[DELETE EXISTING EXCEPTION REPORT RECORDS]
        SQLStmt = "DELETE * FROM [Roster]"
        DBMain.Execute SQLStmt, dbFailOnError
        
        '[SELECT APPROPRIATE RECORDS]
        SQLStmt = "SELECT * FROM Roster"
        Set DsRoster = DBMain.OpenRecordset(SQLStmt, dbOpenDynaset)
                
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
    Case 1
        strMessage = "The roster code is used as a short form to identify which rosters an employee belongs to.  You are allowed to use up to three alpha-numeric characters to define this roster."
    Case 2
        strMessage = "The roster description is used on reports to provide a longer identifier for each roster.  You may use up to 20 alpha-numeric characters in this field."
    End Select
    
    strTitle = "Class Definitions"
    strOldClass = GridClass.Text
    strTemp = GridClass.Text
    
    If GridClass.Col = 0 Then
        '[ENABLED/DISABLED]
        If GridClass.Text = vbChecked Then
            GridClass.Text = vbUnchecked
            frmSet.GridClass.Picture = frmSet.ImageSwitch(constCritical).Picture
        Else
            GridClass.Text = vbChecked
            frmSet.GridClass.Picture = frmSet.ImageSwitch(constWarning).Picture
        End If
        GridClass.Refresh
    ElseIf GridClass.Col = 3 Then
        '[PLACE ROSTER TEXT IN THE GRID]
        Dim flagResult      As Boolean
        
        gsrNote = GridClass.Text
        GridClass.Col = 2
        flagResult = gsrMsg(gsrNote, vbQuestion, "Enter " & GridClass.Text & " Roster Notes")
        GridClass.Col = 3
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
            If GridClass.Col = 1 Then
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
            frmSet.GridClass.Col = 1
            DsClass!Code = frmSet.GridClass.Text
            frmSet.GridClass.Col = 2
            DsClass!Description = frmSet.GridClass.Text
            frmSet.GridClass.Col = 3
            DsClass!Note = frmSet.GridClass.Text
            frmSet.GridClass.Col = 0
            DsClass!Active = frmSet.GridClass.Text
        DsClass.Update
    End If
    
    '[ONLY UPDATE IF NAME OR STATE HAS CHANGED]
    If intCol = 2 Or intCol = 0 Then
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
        frmSet.OptionDateFormat(2).Enabled = False
        '[IF OPTION 3 WAS CHECKED, MOVE TO FIRST OPTION]
        If frmSet.OptionDateFormat(2).Value = True Then frmSet.OptionDateFormat(0).Value = True
    Else
        '[IS DATE FORMAT, ENABLE OPTION 3]
        frmSet.OptionDateFormat(2).Enabled = True
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





