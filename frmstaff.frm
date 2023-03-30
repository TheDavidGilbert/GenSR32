VERSION 4.00
Begin VB.Form frmStaff 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Staff Details"
   ClientHeight    =   4890
   ClientLeft      =   1575
   ClientTop       =   1770
   ClientWidth     =   6585
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
   Height          =   5295
   HelpContextID   =   40
   Icon            =   "FRMSTAFF.frx":0000
   Left            =   1515
   LinkTopic       =   "frmEmpList"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4890
   ScaleWidth      =   6585
   ShowInTaskbar   =   0   'False
   Top             =   1425
   Width           =   6705
   Begin VB.ListBox ListStaff 
      Height          =   4470
      Left            =   30
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   360
      Width           =   1800
   End
   Begin Threed.SSPanel PanelToolBar 
      Align           =   1  'Align Top
      Height          =   300
      Left            =   0
      Negotiate       =   -1  'True
      TabIndex        =   62
      Top             =   0
      Width           =   6585
      _version        =   65536
      _extentx        =   11615
      _extenty        =   529
      _stockprops     =   15
      forecolor       =   -2147483641
      backcolor       =   12632256
      BeginProperty font {FB8F0823-0164-101B-84ED-08002B2EC713} 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   400
         size            =   8.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      borderwidth     =   1
      bevelouter      =   0
      floodtype       =   1
      floodcolor      =   -2147483646
      floodshowpct    =   0   'False
      alignment       =   0
      autosize        =   2
      mouseicon       =   "FRMSTAFF.frx":08CA
      Begin VB.Image cmdAllStaffRosters 
         Height          =   300
         Left            =   1260
         Picture         =   "FRMSTAFF.frx":11A4
         Top             =   0
         Width           =   300
      End
      Begin VB.Image cmdSelectedStaffRoster 
         Height          =   300
         Left            =   960
         Picture         =   "FRMSTAFF.frx":1778
         Top             =   0
         Width           =   300
      End
      Begin VB.Image cmdNew 
         Height          =   300
         Left            =   660
         Picture         =   "FRMSTAFF.frx":1D4C
         Top             =   0
         Width           =   300
      End
      Begin VB.Image cmdSave 
         Height          =   300
         Left            =   60
         Picture         =   "FRMSTAFF.frx":2320
         Top             =   0
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Image cmdDelete 
         Height          =   300
         Left            =   360
         Picture         =   "FRMSTAFF.frx":28F4
         Top             =   0
         Width           =   300
      End
   End
   Begin TabDlg.SSTab tabStaff 
      Height          =   4500
      HelpContextID   =   40
      Left            =   1920
      TabIndex        =   1
      Top             =   360
      Width           =   4635
      _version        =   65536
      _extentx        =   8176
      _extenty        =   7938
      _stockprops     =   15
      caption         =   "Staff Details"
      tabsperrow      =   4
      tab             =   0
      taborientation  =   0
      tabs            =   4
      style           =   1
      tabmaxwidth     =   0
      tabheight       =   529
      tabcaption(0)   =   "Staff Details"
      tab(0).controlcount=   22
      tab(0).controlenabled=   -1  'True
      tab(0).control(0)=   "label_std(4)"
      tab(0).control(1)=   "label_std(5)"
      tab(0).control(2)=   "label_std(6)"
      tab(0).control(3)=   "label_std(7)"
      tab(0).control(4)=   "label_std(0)"
      tab(0).control(5)=   "label_std(3)"
      tab(0).control(6)=   "label_std(2)"
      tab(0).control(7)=   "label_std(1)"
      tab(0).control(8)=   "MaskHourRate"
      tab(0).control(9)=   "MaskDateHired"
      tab(0).control(10)=   "MaskBirthDate"
      tab(0).control(11)=   "MaskHomePhone"
      tab(0).control(12)=   "cmdReturn(0)"
      tab(0).control(13)=   "labelInfo"
      tab(0).control(14)=   "ShapeBorder"
      tab(0).control(15)=   "TextNote"
      tab(0).control(16)=   "TextStaffID"
      tab(0).control(17)=   "TextMiddleName"
      tab(0).control(18)=   "TextFirstName"
      tab(0).control(19)=   "TextLastName"
      tab(0).control(20)=   "OptionPay(0)"
      tab(0).control(21)=   "OptionPay(1)"
      tabcaption(1)   =   "Availability"
      tab(1).controlcount=   39
      tab(1).controlenabled=   0   'False
      tab(1).control(0)=   "comboDayStatus(0)"
      tab(1).control(1)=   "comboDayStatus(1)"
      tab(1).control(2)=   "comboDayStatus(2)"
      tab(1).control(3)=   "comboDayStatus(3)"
      tab(1).control(4)=   "comboDayStatus(4)"
      tab(1).control(5)=   "comboDayStatus(5)"
      tab(1).control(6)=   "comboDayStatus(6)"
      tab(1).control(7)=   "textStart(0)"
      tab(1).control(8)=   "textStart(1)"
      tab(1).control(9)=   "textStart(2)"
      tab(1).control(10)=   "textStart(3)"
      tab(1).control(11)=   "textStart(4)"
      tab(1).control(12)=   "textStart(5)"
      tab(1).control(13)=   "textStart(6)"
      tab(1).control(14)=   "textFinish(0)"
      tab(1).control(15)=   "textFinish(1)"
      tab(1).control(16)=   "textFinish(2)"
      tab(1).control(17)=   "textFinish(3)"
      tab(1).control(18)=   "textFinish(4)"
      tab(1).control(19)=   "textFinish(5)"
      tab(1).control(20)=   "textFinish(6)"
      tab(1).control(21)=   "labelDay(0)"
      tab(1).control(22)=   "labelDay(1)"
      tab(1).control(23)=   "labelDay(2)"
      tab(1).control(24)=   "labelDay(3)"
      tab(1).control(25)=   "labelDay(4)"
      tab(1).control(26)=   "labelDay(5)"
      tab(1).control(27)=   "labelDay(6)"
      tab(1).control(28)=   "label_std(9)"
      tab(1).control(29)=   "lbl_std(1)"
      tab(1).control(30)=   "lbl_std(2)"
      tab(1).control(31)=   "label_std(8)"
      tab(1).control(32)=   "label_std(12)"
      tab(1).control(33)=   "label_std(13)"
      tab(1).control(34)=   "cmdReturn(1)"
      tab(1).control(35)=   "MaskMinHours"
      tab(1).control(36)=   "MaskMaxHours"
      tab(1).control(37)=   "label_std(10)"
      tab(1).control(38)=   "label_std(11)"
      tabcaption(2)   =   "Rosters "
      tab(2).controlcount=   33
      tab(2).controlenabled=   0   'False
      tab(2).control(0)=   "CheckClass(9)"
      tab(2).control(1)=   "CheckClass(8)"
      tab(2).control(2)=   "CheckClass(7)"
      tab(2).control(3)=   "CheckClass(6)"
      tab(2).control(4)=   "CheckClass(5)"
      tab(2).control(5)=   "CheckClass(4)"
      tab(2).control(6)=   "CheckClass(3)"
      tab(2).control(7)=   "CheckClass(2)"
      tab(2).control(8)=   "CheckClass(1)"
      tab(2).control(9)=   "CheckClass(0)"
      tab(2).control(10)=   "cmdBaseRate(9)"
      tab(2).control(11)=   "cmdBaseRate(8)"
      tab(2).control(12)=   "cmdBaseRate(7)"
      tab(2).control(13)=   "cmdBaseRate(6)"
      tab(2).control(14)=   "cmdBaseRate(5)"
      tab(2).control(15)=   "cmdBaseRate(4)"
      tab(2).control(16)=   "cmdBaseRate(3)"
      tab(2).control(17)=   "cmdBaseRate(2)"
      tab(2).control(18)=   "cmdBaseRate(1)"
      tab(2).control(19)=   "cmdBaseRate(0)"
      tab(2).control(20)=   "lbl_hr_rate"
      tab(2).control(21)=   "label_std(14)"
      tab(2).control(22)=   "maskRosterRate(9)"
      tab(2).control(23)=   "maskRosterRate(8)"
      tab(2).control(24)=   "maskRosterRate(7)"
      tab(2).control(25)=   "maskRosterRate(6)"
      tab(2).control(26)=   "maskRosterRate(5)"
      tab(2).control(27)=   "maskRosterRate(4)"
      tab(2).control(28)=   "maskRosterRate(3)"
      tab(2).control(29)=   "maskRosterRate(2)"
      tab(2).control(30)=   "maskRosterRate(1)"
      tab(2).control(31)=   "maskRosterRate(0)"
      tab(2).control(32)=   "cmdReturn(2)"
      tabcaption(3)   =   "Holidays"
      tab(3).controlcount=   19
      tab(3).controlenabled=   0   'False
      tab(3).control(0)=   "CheckHoliday"
      tab(3).control(1)=   "txtHolDays"
      tab(3).control(2)=   "cmdEndWeekMinus"
      tab(3).control(3)=   "cmdendWeekPlus"
      tab(3).control(4)=   "cmdStartWeekMinus"
      tab(3).control(5)=   "cmdStartWeekPlus"
      tab(3).control(6)=   "cmdHolEndMinus"
      tab(3).control(7)=   "cmdholEndPlus"
      tab(3).control(8)=   "cmdHolStartMinus"
      tab(3).control(9)=   "cmdHolStartPlus"
      tab(3).control(10)=   "label_std(20)"
      tab(3).control(11)=   "cmdReturn(3)"
      tab(3).control(12)=   "cmdHolEndToday"
      tab(3).control(13)=   "cmdHolStartToday"
      tab(3).control(14)=   "label_std(17)"
      tab(3).control(15)=   "label_std(15)"
      tab(3).control(16)=   "label_std(16)"
      tab(3).control(17)=   "MaskHolEnd"
      tab(3).control(18)=   "MaskHolStart"
      Begin VB.CheckBox CheckClass 
         Caption         =   "Roster"
         Height          =   255
         Index           =   9
         Left            =   -74880
         TabIndex        =   54
         Top             =   3720
         Width           =   2895
      End
      Begin VB.CheckBox CheckClass 
         Caption         =   "Roster"
         Height          =   255
         Index           =   8
         Left            =   -74880
         TabIndex        =   52
         Top             =   3390
         Width           =   2895
      End
      Begin VB.CheckBox CheckClass 
         Caption         =   "Roster"
         Height          =   255
         Index           =   7
         Left            =   -74880
         TabIndex        =   50
         Top             =   3060
         Width           =   2895
      End
      Begin VB.CheckBox CheckClass 
         Caption         =   "Roster"
         Height          =   255
         Index           =   6
         Left            =   -74880
         TabIndex        =   48
         Top             =   2730
         Width           =   2895
      End
      Begin VB.CheckBox CheckClass 
         Caption         =   "Roster"
         Height          =   255
         Index           =   5
         Left            =   -74880
         TabIndex        =   46
         Top             =   2400
         Width           =   2895
      End
      Begin VB.CheckBox CheckClass 
         Caption         =   "Roster"
         Height          =   255
         Index           =   4
         Left            =   -74880
         TabIndex        =   44
         Top             =   2070
         Width           =   2895
      End
      Begin VB.CheckBox CheckClass 
         Caption         =   "Roster"
         Height          =   255
         Index           =   3
         Left            =   -74880
         TabIndex        =   42
         Top             =   1740
         Width           =   2895
      End
      Begin VB.CheckBox CheckClass 
         Caption         =   "Roster"
         Height          =   255
         Index           =   2
         Left            =   -74880
         TabIndex        =   40
         Top             =   1410
         Width           =   2895
      End
      Begin VB.CheckBox CheckClass 
         Caption         =   "Roster"
         Height          =   255
         Index           =   1
         Left            =   -74880
         TabIndex        =   38
         Top             =   1080
         Width           =   2895
      End
      Begin VB.CheckBox CheckClass 
         Caption         =   "Roster"
         Height          =   255
         Index           =   0
         Left            =   -74880
         TabIndex        =   36
         Top             =   750
         Width           =   2895
      End
      Begin VB.CheckBox CheckHoliday 
         Alignment       =   1  'Right Justify
         Caption         =   "On/Off"
         Height          =   210
         Left            =   -74940
         TabIndex        =   57
         Top             =   780
         Width           =   855
      End
      Begin VB.OptionButton OptionPay 
         Caption         =   "Weekly"
         Height          =   255
         Index           =   1
         Left            =   1140
         TabIndex        =   94
         Top             =   2040
         Width           =   855
      End
      Begin VB.OptionButton OptionPay 
         Caption         =   "Hourly"
         Height          =   255
         Index           =   0
         Left            =   1140
         TabIndex        =   93
         Top             =   1740
         Width           =   855
      End
      Begin VB.TextBox txtHolDays 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   -70935
         MaxLength       =   25
         TabIndex        =   60
         Top             =   690
         Width           =   500
      End
      Begin VB.ComboBox comboDayStatus 
         Height          =   330
         Index           =   0
         ItemData        =   "FRMSTAFF.frx":2EC8
         Left            =   -73848
         List            =   "FRMSTAFF.frx":2ED8
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   696
         Width           =   1155
      End
      Begin VB.ComboBox comboDayStatus 
         Height          =   330
         Index           =   1
         ItemData        =   "FRMSTAFF.frx":2EF6
         Left            =   -73848
         List            =   "FRMSTAFF.frx":2F06
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   1056
         Width           =   1155
      End
      Begin VB.ComboBox comboDayStatus 
         Height          =   330
         Index           =   2
         ItemData        =   "FRMSTAFF.frx":2F24
         Left            =   -73848
         List            =   "FRMSTAFF.frx":2F34
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   1416
         Width           =   1155
      End
      Begin VB.ComboBox comboDayStatus 
         Height          =   330
         Index           =   3
         ItemData        =   "FRMSTAFF.frx":2F52
         Left            =   -73848
         List            =   "FRMSTAFF.frx":2F62
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   1776
         Width           =   1155
      End
      Begin VB.ComboBox comboDayStatus 
         Height          =   330
         Index           =   4
         ItemData        =   "FRMSTAFF.frx":2F80
         Left            =   -73848
         List            =   "FRMSTAFF.frx":2F90
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   2136
         Width           =   1155
      End
      Begin VB.ComboBox comboDayStatus 
         Height          =   330
         Index           =   5
         ItemData        =   "FRMSTAFF.frx":2FAE
         Left            =   -73848
         List            =   "FRMSTAFF.frx":2FBE
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   2496
         Width           =   1155
      End
      Begin VB.ComboBox comboDayStatus 
         Height          =   330
         Index           =   6
         ItemData        =   "FRMSTAFF.frx":2FDC
         Left            =   -73848
         List            =   "FRMSTAFF.frx":2FEC
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   2856
         Width           =   1155
      End
      Begin VB.TextBox textStart 
         Height          =   315
         Index           =   0
         Left            =   -72600
         TabIndex        =   13
         Top             =   696
         Width           =   1000
      End
      Begin VB.TextBox textStart 
         Height          =   315
         Index           =   1
         Left            =   -72600
         TabIndex        =   16
         Top             =   1056
         Width           =   1000
      End
      Begin VB.TextBox textStart 
         Height          =   315
         Index           =   2
         Left            =   -72600
         TabIndex        =   19
         Top             =   1416
         Width           =   1000
      End
      Begin VB.TextBox textStart 
         Height          =   315
         Index           =   3
         Left            =   -72600
         TabIndex        =   22
         Top             =   1776
         Width           =   1000
      End
      Begin VB.TextBox textStart 
         Height          =   315
         Index           =   4
         Left            =   -72600
         TabIndex        =   25
         Top             =   2136
         Width           =   1000
      End
      Begin VB.TextBox textStart 
         Height          =   315
         Index           =   5
         Left            =   -72600
         TabIndex        =   28
         Top             =   2496
         Width           =   1000
      End
      Begin VB.TextBox textStart 
         Height          =   315
         Index           =   6
         Left            =   -72600
         TabIndex        =   31
         Top             =   2856
         Width           =   1000
      End
      Begin VB.TextBox textFinish 
         Height          =   315
         Index           =   0
         Left            =   -71532
         TabIndex        =   14
         Top             =   696
         Width           =   1000
      End
      Begin VB.TextBox textFinish 
         Height          =   315
         Index           =   1
         Left            =   -71532
         TabIndex        =   17
         Top             =   1056
         Width           =   1000
      End
      Begin VB.TextBox textFinish 
         Height          =   315
         Index           =   2
         Left            =   -71532
         TabIndex        =   20
         Top             =   1416
         Width           =   1000
      End
      Begin VB.TextBox textFinish 
         Height          =   315
         Index           =   3
         Left            =   -71532
         TabIndex        =   23
         Top             =   1776
         Width           =   1000
      End
      Begin VB.TextBox textFinish 
         Height          =   315
         Index           =   4
         Left            =   -71532
         TabIndex        =   26
         Top             =   2136
         Width           =   1000
      End
      Begin VB.TextBox textFinish 
         Height          =   315
         Index           =   5
         Left            =   -71532
         TabIndex        =   29
         Top             =   2496
         Width           =   1000
      End
      Begin VB.TextBox textFinish 
         Height          =   315
         Index           =   6
         Left            =   -71520
         TabIndex        =   32
         Top             =   2850
         Width           =   1000
      End
      Begin VB.TextBox TextLastName 
         DataField       =   "LastName"
         Height          =   315
         Left            =   1160
         MaxLength       =   25
         TabIndex        =   3
         Text            =   "LastName"
         Top             =   670
         Width           =   1060
      End
      Begin VB.TextBox TextFirstName 
         DataField       =   "FirstName"
         Height          =   315
         Left            =   2260
         MaxLength       =   25
         TabIndex        =   4
         Text            =   "FirstName"
         Top             =   670
         Width           =   1060
      End
      Begin VB.TextBox TextMiddleName 
         DataField       =   "MiddleName"
         Height          =   315
         Left            =   3360
         MaxLength       =   25
         TabIndex        =   5
         Text            =   "MiddleName"
         Top             =   670
         Width           =   1200
      End
      Begin VB.TextBox TextStaffID 
         DataField       =   "StaffID"
         Height          =   315
         Left            =   60
         MaxLength       =   10
         TabIndex        =   2
         Text            =   "StaffID"
         Top             =   670
         Width           =   1060
      End
      Begin VB.TextBox TextNote 
         DataField       =   "Note"
         Height          =   1080
         Left            =   60
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Text            =   "FRMSTAFF.frx":300A
         Top             =   2940
         Width           =   4515
      End
      Begin VB.Image cmdBaseRate 
         Height          =   300
         Index           =   9
         Left            =   -70920
         Picture         =   "FRMSTAFF.frx":3013
         Top             =   3670
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Image cmdBaseRate 
         Height          =   300
         Index           =   8
         Left            =   -70920
         Picture         =   "FRMSTAFF.frx":35E7
         Top             =   3340
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Image cmdBaseRate 
         Height          =   300
         Index           =   7
         Left            =   -70920
         Picture         =   "FRMSTAFF.frx":3BBB
         Top             =   3010
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Image cmdBaseRate 
         Height          =   300
         Index           =   6
         Left            =   -70920
         Picture         =   "FRMSTAFF.frx":418F
         Top             =   2680
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Image cmdBaseRate 
         Height          =   300
         Index           =   5
         Left            =   -70920
         Picture         =   "FRMSTAFF.frx":4763
         Top             =   2350
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Image cmdBaseRate 
         Height          =   300
         Index           =   4
         Left            =   -70920
         Picture         =   "FRMSTAFF.frx":4D37
         Top             =   2020
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Image cmdBaseRate 
         Height          =   300
         Index           =   3
         Left            =   -70920
         Picture         =   "FRMSTAFF.frx":530B
         Top             =   1690
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Image cmdBaseRate 
         Height          =   300
         Index           =   2
         Left            =   -70920
         Picture         =   "FRMSTAFF.frx":58DF
         Top             =   1360
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Image cmdBaseRate 
         Height          =   300
         Index           =   1
         Left            =   -70920
         Picture         =   "FRMSTAFF.frx":5EB3
         Top             =   1030
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Image cmdBaseRate 
         Height          =   300
         Index           =   0
         Left            =   -70920
         Picture         =   "FRMSTAFF.frx":6487
         Top             =   700
         Visible         =   0   'False
         Width           =   300
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
      Begin VB.Label labelInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   1035
         Left            =   2100
         TabIndex        =   92
         Top             =   1800
         Width           =   2415
         WordWrap        =   -1  'True
      End
      Begin VB.Image cmdEndWeekMinus 
         Height          =   300
         Left            =   -72480
         Picture         =   "FRMSTAFF.frx":6A5B
         Top             =   1080
         Width           =   300
      End
      Begin VB.Image cmdendWeekPlus 
         Height          =   300
         Left            =   -71280
         Picture         =   "FRMSTAFF.frx":702F
         Top             =   1080
         Width           =   300
      End
      Begin VB.Image cmdStartWeekMinus 
         Height          =   300
         Left            =   -74040
         Picture         =   "FRMSTAFF.frx":7603
         Top             =   1080
         Width           =   300
      End
      Begin VB.Image cmdStartWeekPlus 
         Height          =   300
         Left            =   -72840
         Picture         =   "FRMSTAFF.frx":7BD7
         Top             =   1080
         Width           =   300
      End
      Begin VB.Image cmdHolEndMinus 
         Height          =   300
         Left            =   -72180
         Picture         =   "FRMSTAFF.frx":81AB
         Top             =   1080
         Width           =   300
      End
      Begin VB.Image cmdholEndPlus 
         Height          =   300
         Left            =   -71580
         Picture         =   "FRMSTAFF.frx":877F
         Top             =   1080
         Width           =   300
      End
      Begin VB.Image cmdHolStartMinus 
         Height          =   300
         Left            =   -73740
         Picture         =   "FRMSTAFF.frx":8D53
         Top             =   1080
         Width           =   300
      End
      Begin VB.Image cmdHolStartPlus 
         Height          =   300
         Left            =   -73140
         Picture         =   "FRMSTAFF.frx":9327
         Top             =   1080
         Width           =   300
      End
      Begin VB.Label label_std 
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         BackStyle       =   0  'Transparent
         Caption         =   $"FRMSTAFF.frx":98FB
         Height          =   2550
         Index           =   20
         Left            =   -74880
         TabIndex        =   90
         Top             =   1440
         Width           =   4440
         WordWrap        =   -1  'True
      End
      Begin Threed.SSCommand cmdReturn 
         Cancel          =   -1  'True
         Height          =   360
         Index           =   3
         Left            =   -71130
         TabIndex        =   61
         Top             =   4080
         Width           =   705
         _version        =   65536
         _extentx        =   1244
         _extenty        =   635
         _stockprops     =   78
         caption         =   "C&lose"
         autosize        =   2
      End
      Begin VB.Image cmdHolEndToday 
         Height          =   300
         Left            =   -71880
         Picture         =   "FRMSTAFF.frx":9985
         Top             =   1080
         Width           =   300
      End
      Begin VB.Image cmdHolStartToday 
         Height          =   300
         Left            =   -73440
         Picture         =   "FRMSTAFF.frx":9F59
         Top             =   1080
         Width           =   300
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
         TabIndex        =   89
         Top             =   420
         Width           =   495
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
         TabIndex        =   87
         Top             =   420
         Width           =   1470
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
         TabIndex        =   88
         Top             =   420
         Width           =   1470
      End
      Begin MSMask.MaskEdBox MaskHolEnd 
         DataField       =   "DateHired"
         Height          =   315
         Left            =   -72480
         TabIndex        =   59
         Top             =   690
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
      Begin MSMask.MaskEdBox MaskHolStart 
         DataField       =   "DateHired"
         Height          =   315
         Left            =   -74040
         TabIndex        =   58
         Top             =   690
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
      Begin VB.Label labelDay 
         BackColor       =   &H80000002&
         BackStyle       =   0  'Transparent
         Caption         =   "Day Name"
         Height          =   216
         Index           =   0
         Left            =   -74880
         TabIndex        =   86
         Top             =   780
         Width           =   996
      End
      Begin VB.Label labelDay 
         BackColor       =   &H80000002&
         BackStyle       =   0  'Transparent
         Caption         =   "Day Name"
         Height          =   216
         Index           =   1
         Left            =   -74880
         TabIndex        =   85
         Top             =   1140
         Width           =   996
      End
      Begin VB.Label labelDay 
         BackColor       =   &H80000002&
         BackStyle       =   0  'Transparent
         Caption         =   "Day Name"
         Height          =   216
         Index           =   2
         Left            =   -74880
         TabIndex        =   84
         Top             =   1488
         Width           =   996
      End
      Begin VB.Label labelDay 
         BackColor       =   &H80000002&
         BackStyle       =   0  'Transparent
         Caption         =   "Day Name"
         Height          =   216
         Index           =   3
         Left            =   -74880
         TabIndex        =   83
         Top             =   1848
         Width           =   996
      End
      Begin VB.Label labelDay 
         BackColor       =   &H80000002&
         BackStyle       =   0  'Transparent
         Caption         =   "Day Name"
         Height          =   216
         Index           =   4
         Left            =   -74880
         TabIndex        =   82
         Top             =   2208
         Width           =   996
      End
      Begin VB.Label labelDay 
         BackColor       =   &H80000002&
         BackStyle       =   0  'Transparent
         Caption         =   "Day Name"
         Height          =   216
         Index           =   5
         Left            =   -74880
         TabIndex        =   81
         Top             =   2568
         Width           =   996
      End
      Begin VB.Label labelDay 
         BackColor       =   &H80000002&
         BackStyle       =   0  'Transparent
         Caption         =   "Day Name"
         Height          =   216
         Index           =   6
         Left            =   -74880
         TabIndex        =   80
         Top             =   2928
         Width           =   996
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
         TabIndex        =   79
         Top             =   420
         Width           =   1155
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
         TabIndex        =   78
         Top             =   420
         Visible         =   0   'False
         Width           =   1000
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
         TabIndex        =   77
         Top             =   420
         Visible         =   0   'False
         Width           =   1000
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
         TabIndex        =   76
         Top             =   420
         Width           =   996
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
         TabIndex        =   75
         Top             =   3540
         Width           =   1005
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
         TabIndex        =   74
         Top             =   3540
         Width           =   1005
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
         TabIndex        =   73
         Top             =   420
         Width           =   1335
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
         TabIndex        =   72
         Top             =   420
         Width           =   2900
      End
      Begin MSMask.MaskEdBox maskRosterRate 
         DataField       =   "HourRate"
         Height          =   315
         Index           =   9
         Left            =   -71940
         TabIndex        =   55
         Top             =   3660
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
         TabIndex        =   53
         Top             =   3330
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
         TabIndex        =   51
         Top             =   3000
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
         TabIndex        =   49
         Top             =   2670
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
         TabIndex        =   47
         Top             =   2340
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
         TabIndex        =   45
         Top             =   2010
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
         TabIndex        =   43
         Top             =   1680
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
         TabIndex        =   41
         Top             =   1350
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
         TabIndex        =   39
         Top             =   1020
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
         Index           =   0
         Left            =   -71940
         TabIndex        =   37
         Top             =   690
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
      Begin Threed.SSCommand cmdReturn 
         Height          =   360
         Index           =   0
         Left            =   3870
         TabIndex        =   11
         Top             =   4080
         Width           =   705
         _version        =   65536
         _extentx        =   1244
         _extenty        =   635
         _stockprops     =   78
         caption         =   "C&lose"
         autosize        =   2
      End
      Begin MSMask.MaskEdBox MaskHomePhone 
         DataField       =   "HomePhone"
         Height          =   315
         Left            =   3100
         TabIndex        =   8
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
      Begin MSMask.MaskEdBox MaskBirthDate 
         DataField       =   "Birthdate"
         Height          =   315
         Left            =   60
         TabIndex        =   6
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
      Begin MSMask.MaskEdBox MaskDateHired 
         DataField       =   "DateHired"
         Height          =   315
         Left            =   1580
         TabIndex        =   7
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
      Begin MSMask.MaskEdBox MaskHourRate 
         DataField       =   "HourRate"
         Height          =   315
         Left            =   60
         TabIndex        =   9
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
      Begin Threed.SSCommand cmdReturn 
         Height          =   360
         Index           =   1
         Left            =   -71130
         TabIndex        =   35
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
         Index           =   2
         Left            =   -71130
         TabIndex        =   56
         Top             =   4080
         Width           =   705
         _version        =   65536
         _extentx        =   1244
         _extenty        =   635
         _stockprops     =   78
         caption         =   "C&lose"
         autosize        =   2
      End
      Begin MSMask.MaskEdBox MaskMinHours 
         DataField       =   "HourRate"
         Height          =   315
         Left            =   -74880
         TabIndex        =   33
         Top             =   3820
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
      Begin MSMask.MaskEdBox MaskMaxHours 
         DataField       =   "HourRate"
         Height          =   315
         Left            =   -73800
         TabIndex        =   34
         Top             =   3820
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
      Begin VB.Label label_std 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Hours Required per Week"
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   10
         Left            =   -74880
         TabIndex        =   63
         Top             =   3240
         Width           =   4360
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
         TabIndex        =   70
         Top             =   396
         Width           =   1060
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
         TabIndex        =   69
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
         TabIndex        =   68
         Top             =   390
         Width           =   1200
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
         TabIndex        =   71
         Top             =   390
         Width           =   1060
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
         TabIndex        =   67
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
         TabIndex        =   64
         Top             =   1710
         Width           =   1065
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
         TabIndex        =   65
         Top             =   1050
         Width           =   1460
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
         TabIndex        =   66
         Top             =   1050
         Width           =   1460
      End
      Begin VB.Label label_std 
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         BackStyle       =   0  'Transparent
         Caption         =   "Minimum && maximum hours per week required by this staff member."
         Height          =   630
         Index           =   11
         Left            =   -72720
         TabIndex        =   91
         Top             =   3540
         Width           =   2280
         WordWrap        =   -1  'True
      End
   End
End
Attribute VB_Name = "frmStaff"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Option Explicit
'[----------------------------------------------]
'[frmStaff          Staff Details form          ]
'[----------------------------------------------]
'[GSR                                           ]
'[generic staff roster system for Windows       ]
'[(C) ---------- David Gilbert, 1997            ]
'[----------------------------------------------]
'[version                           2.1         ]
'[initial date                      21/07/1997  ]
'[----------------------------------------------]

Dim Shared intStartDay As Integer       '[Start time - week day]
Dim Shared intFinishDay As Integer      '[Finish time - week day]





Private Sub CheckClass_Click(Index As Integer)

    '[MAKE SAVE BUTTON VISIBLE]
    If frmStaff.cmdSave.Visible = False Then frmStaff.cmdSave.Visible = True
    
    '[MAKE RATE VISIBLE IF HOURLY PAY RATE IS SET AND CLASS IS CHECKED]
    If frmStaff.CheckClass(Index).Value = vbChecked And frmStaff.optionPay(vbHourly).Value = True Then
        frmStaff.maskRosterRate(Index).Visible = True
        frmStaff.cmdBaseRate(Index).Visible = True
        frmStaff.lbl_hr_Rate.Visible = True
    Else
        frmStaff.maskRosterRate(Index).Visible = False
        frmStaff.cmdBaseRate(Index).Visible = False
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
    If frmStaff.cmdSave.Visible = False Then frmStaff.cmdSave.Visible = True
    
    '[ENABLE/DISABLE DATES]
    If frmStaff.CheckHoliday.Value = vbChecked Then
        frmStaff.MaskHolStart.Enabled = True
        frmStaff.MaskHolEnd.Enabled = True
        If IsDate(frmStaff.MaskHolStart.Text) And IsDate(frmStaff.MaskHolEnd.Text) Then frmStaff.txtHolDays = (Format(CDate(frmStaff.MaskHolEnd) - CDate(frmStaff.MaskHolStart), "###0")) + 1
    Else
        frmStaff.MaskHolStart.Enabled = False
        frmStaff.MaskHolEnd.Enabled = False
        frmStaff.txtHolDays = ""
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
    frmStaff.maskRosterRate(Index).Text = frmStaff.MaskHourRate.Text

    '[ENABLE SAVE BUTTON IF A CHANGE HAS BEEN MADE]
    If frmStaff.cmdSave.Visible = False Then frmStaff.cmdSave.Visible = True

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
        
        If frmStaff.ListStaff.ListIndex = -1 Then Exit Sub
        
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
            If frmStaff.ListStaff.ListCount > 1 Then
                strDisplayname = frmStaff.ListStaff.List(frmStaff.ListStaff.ListIndex - 1)
            Else
                strDisplayname = ""
            End If
            
            '[FILL STAFF LIST SO WE GET ORDER]
            FillStaffList
                
            '[RELOCATE STAFF NAME]
            For intColCounter = 0 To (frmStaff.ListStaff.ListCount - 1)
                If frmStaff.ListStaff.List(intColCounter) = strDisplayname Then
                    frmStaff.ListStaff.ListIndex = intColCounter
                    Exit For
                End If
            Next intColCounter
        
            '[CHECK FOR NO STAFF MEMBERS]
            If DsStaff.RecordCount = 0 Then
                '[CALL ROUTINE TO ADD NEW STAFF MEMBER AND REPOSITION LIST]
                AddNewStaff
            End If
        
            '[MOVE TO FIRST ITEM IF NO ITEM IS SELECTED]
            If frmStaff.ListStaff.ListIndex < 0 Then frmStaff.ListStaff.ListIndex = 0
        
        End If
    

End Sub

Private Sub cmdDelete_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Delete the currently selected staff member (" & frmStaff.ListStaff.List(frmStaff.ListStaff.ListIndex) & ") from the list."
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub cmdEndWeekMinus_Click()
    
    '[ANIMATE BUTTON]
    frmStaff.cmdEndWeekMinus.BorderStyle = 1
    Delay vbDelay
    frmStaff.cmdEndWeekMinus.BorderStyle = 0
    
    '[CHANGE HOLEND DATE]
    If IsDate(frmStaff.MaskHolEnd) Then frmStaff.MaskHolEnd = Format(CDate(frmStaff.MaskHolEnd) - 7, strDateFormat)

End Sub

Private Sub cmdEndWeekMinus_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Decrease the 'Holiday End Date' by 7 days."
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub cmdendWeekPlus_Click()
    
    '[ANIMATE BUTTON]
    frmStaff.cmdendWeekPlus.BorderStyle = 1
    Delay vbDelay
    frmStaff.cmdendWeekPlus.BorderStyle = 0
    
    '[CHANGE HOLEND DATE]
    If IsDate(frmStaff.MaskHolEnd) Then frmStaff.MaskHolEnd = Format(CDate(frmStaff.MaskHolEnd) + 7, strDateFormat)

End Sub

Private Sub cmdendWeekPlus_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Increase the 'Holiday End Date' by 7 days."
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub cmdHolEndMinus_Click()
    
    '[ANIMATE BUTTON]
    frmStaff.cmdHolEndMinus.BorderStyle = 1
    Delay vbDelay
    frmStaff.cmdHolEndMinus.BorderStyle = 0
    
    '[CHANGE HOLEND DATE]
    If IsDate(frmStaff.MaskHolEnd) Then frmStaff.MaskHolEnd = Format(CDate(frmStaff.MaskHolEnd) - 1, strDateFormat)

End Sub

Private Sub cmdHolEndMinus_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Decrease the 'Holiday End Date' by 1 day."
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub cmdholEndPlus_Click()
    
    '[ANIMATE BUTTON]
    frmStaff.cmdholEndPlus.BorderStyle = 1
    Delay vbDelay
    frmStaff.cmdholEndPlus.BorderStyle = 0
    
    '[CHANGE HOLEND DATE]
    If IsDate(frmStaff.MaskHolEnd) Then frmStaff.MaskHolEnd = Format(CDate(frmStaff.MaskHolEnd) + 1, strDateFormat)
    
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
    frmStaff.MaskHolEnd.Text = Format(Date, strDateFormat)

End Sub

Private Sub cmdHolEndToday_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Set the 'Holiday End Date' to todays date (" & Format(Now, strDateFormat) & ")."

End Sub


Private Sub cmdHolStartMinus_Click()
    
    '[ANIMATE BUTTON]
    frmStaff.cmdHolStartMinus.BorderStyle = 1
    Delay vbDelay
    frmStaff.cmdHolStartMinus.BorderStyle = 0
    
    '[CHANGE HOLEND DATE]
    If IsDate(frmStaff.MaskHolStart) Then frmStaff.MaskHolStart = Format(CDate(frmStaff.MaskHolStart) - 1, strDateFormat)

End Sub

Private Sub cmdHolStartMinus_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Decrease the 'Holiday Start Date' by 1 day."
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub cmdHolStartPlus_Click()
    
    '[ANIMATE BUTTON]
    frmStaff.cmdHolStartPlus.BorderStyle = 1
    Delay vbDelay
    frmStaff.cmdHolStartPlus.BorderStyle = 0
    
    '[CHANGE HOLEND DATE]
    If IsDate(frmStaff.MaskHolStart) Then frmStaff.MaskHolStart = Format(CDate(frmStaff.MaskHolStart) + 1, strDateFormat)

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
    frmStaff.MaskHolStart.Text = Format(Date, strDateFormat)

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
    If frmStaff.cmdSave.Visible = True Then
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



Private Sub cmdReturn_Click(Index As Integer)

    '[HIDE FORM]
    frmStaff.Hide
    mdiMain.ZOrder

End Sub

Private Sub cmdReturn_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Close this window and return to the roster form."
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
    If frmStaff.ListStaff.List(frmStaff.ListStaff.ListIndex) = strDisplayname Then Exit Sub
    
    '[FILL STAFF LIST SO WE GET ORDER]
    FillStaffList
        
    '[RELOCATE STAFF NAME]
    For intColCounter = 0 To (frmStaff.ListStaff.ListCount - 1)
        If frmStaff.ListStaff.List(intColCounter) = strDisplayname Then frmStaff.ListStaff.ListIndex = intColCounter
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
    StatusBar "Print the weekly roster for " & frmStaff.ListStaff.List(frmStaff.ListStaff.ListIndex) & "."
    '[---------------------------------------------------------------------------------]

End Sub




Private Sub cmdStartWeekMinus_Click()
    
    '[ANIMATE BUTTON]
    frmStaff.cmdStartWeekMinus.BorderStyle = 1
    Delay vbDelay
    frmStaff.cmdStartWeekMinus.BorderStyle = 0
    
    '[CHANGE HOLEND DATE]
    If IsDate(frmStaff.MaskHolStart) Then frmStaff.MaskHolStart = Format(CDate(frmStaff.MaskHolStart) - 7, strDateFormat)

End Sub

Private Sub cmdStartWeekMinus_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Decrease the 'Holiday Start Date' by 7 days."
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub cmdStartWeekPlus_Click()
    
    '[ANIMATE BUTTON]
    frmStaff.cmdStartWeekPlus.BorderStyle = 1
    Delay vbDelay
    frmStaff.cmdStartWeekPlus.BorderStyle = 0
    
    '[CHANGE HOLEND DATE]
    If IsDate(frmStaff.MaskHolStart) Then frmStaff.MaskHolStart = Format(CDate(frmStaff.MaskHolStart) + 7, strDateFormat)

End Sub

Private Sub cmdStartWeekPlus_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Increase the 'Holiday Start Date' by 7 days."
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub comboDayStatus_Click(Index As Integer)
    
    '[CHANGE VISIBILITY FOR START AND END TIME TEXT BOXES]
    If frmStaff.comboDayStatus(Index).ListIndex <= 1 Then
        frmStaff.textStart(Index).Visible = False
        frmStaff.textFinish(Index).Visible = False
    Else
        frmStaff.textStart(Index).Visible = True
        frmStaff.textFinish(Index).Visible = True
        frmStaff.lbl_std(1).Visible = True
        frmStaff.lbl_std(2).Visible = True
    End If
    
    '[MAKE SAVE BUTTON VISIBLE]
    If frmStaff.cmdSave.Visible = False Then frmStaff.cmdSave.Visible = True

End Sub


Private Sub comboDayStatus_GotFocus(Index As Integer)

    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Select the desired staff availability for this day."
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub Form_Load()
    
    '[DEBUG]
    Call LogToFile("=-STAFF FORM - LOAD SUB------------------=")

    '[SET STARTING POSITION FOR FORM]
    frmStaff.Top = DsDefault("StaffTop")
    frmStaff.Left = DsDefault("StaffLeft")
    '[CHECK FOR > WIDTH, HEIGHT]
    If frmStaff.Top > Screen.Height Or frmStaff.Left > Screen.Width Then
        '[center form on screen]
        frmStaff.Top = (Screen.Height / 2) - (frmStaff.Height / 2)
        frmStaff.Left = (Screen.Width / 2) - (frmStaff.Width / 2)
    End If
    
    '[DEBUG]
    Call LogToFile("- filling staff list")
    '[FILL STAFF LIST FROM DYNASET]
    FillStaffList
    '[DEBUG]
    Call LogToFile("- setting class labels")
    '[APPLY CLASS LABELS]
    SetClassLabels
    '[DEBUG]
    Call LogToFile("- setting day labels")
    '[APPLY DAY LABELS]
    SetDayLabels
    '[DEBUG]
    Call LogToFile("- displaying staff details")
    '[LOCATE FIRST STAFF MEMBER]
    If frmStaff.ListStaff.ListCount > 0 Then frmStaff.ListStaff.ListIndex = 0
    
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
    If frmStaff.cmdSave.Visible = False Then frmStaff.cmdSave.Visible = True

End Sub

Private Sub MaskDateHired_Change()

    '[ENABLE SAVE BUTTON IF A CHANGE HAS BEEN MADE]
    If frmStaff.cmdSave.Visible = False Then frmStaff.cmdSave.Visible = True

End Sub


Private Sub MaskHolEnd_Change()
    
    '[IF VALID DATE, SET DAYS TO VALID FIGURE AND SHOW SAVE BUTTON]
    If IsDate(frmStaff.MaskHolStart.Text) And IsDate(frmStaff.MaskHolEnd.Text) Then
        frmStaff.txtHolDays = Format((CDate(frmStaff.MaskHolEnd) - CDate(frmStaff.MaskHolStart)) + 1, "###0")
        If CDate(frmStaff.MaskHolEnd) <> DsStaff!HolEnd Then
            If frmStaff.cmdSave.Visible = False Then frmStaff.cmdSave.Visible = True
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
    If IsDate(frmStaff.MaskHolStart.Text) And IsDate(frmStaff.MaskHolEnd.Text) Then
        frmStaff.txtHolDays = Format((CDate(frmStaff.MaskHolEnd) - CDate(frmStaff.MaskHolStart)) + 1, "###0")
        If CDate(frmStaff.MaskHolStart) <> DsStaff!HolStart Then
            If frmStaff.cmdSave.Visible = False Then frmStaff.cmdSave.Visible = True
        End If
    Else
        frmStaff.txtHolDays = ""
    End If

End Sub

Private Sub MaskHolStart_GotFocus()
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Set the 'Holiday Start Date'."
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub MaskHomePhone_Change()


    '[ENABLE SAVE BUTTON IF A CHANGE HAS BEEN MADE]
    If frmStaff.cmdSave.Visible = False Then frmStaff.cmdSave.Visible = True

End Sub

Private Sub MaskHourRate_Change()

    '[ENABLE SAVE BUTTON IF A CHANGE HAS BEEN MADE]
    If frmStaff.cmdSave.Visible = False Then frmStaff.cmdSave.Visible = True

End Sub



Private Sub MaskMaxHours_Change()
    
    '[ENABLE SAVE BUTTON IF A CHANGE HAS BEEN MADE]
    If frmStaff.cmdSave.Visible = False Then frmStaff.cmdSave.Visible = True

End Sub

Private Sub MaskMaxHours_GotFocus()

    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Enter the maximum hours required per week by this staff member."
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub MaskMinHours_Change()
    
    '[ENABLE SAVE BUTTON IF A CHANGE HAS BEEN MADE]
    If frmStaff.cmdSave.Visible = False Then frmStaff.cmdSave.Visible = True

End Sub

Private Sub MaskMinHours_GotFocus()

    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Enter the minimum hours required per week by this staff member."
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub maskRosterRate_Change(Index As Integer)

    '[ENABLE SAVE BUTTON IF A CHANGE HAS BEEN MADE]
    If frmStaff.cmdSave.Visible = False Then frmStaff.cmdSave.Visible = True

End Sub







Private Sub maskRosterRate_GotFocus(Index As Integer)

    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Enter the hourly rate earned by this staff member when allocated to this roster."
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub optionPay_Click(Index As Integer)

    '[REV: 3.00.28]
    '[ENABLE SAVE BUTTON IF A CHANGE HAS BEEN MADE]
    If frmStaff.cmdSave.Visible = False Then frmStaff.cmdSave.Visible = True

    '[CYCLE THROUGH AND HIDE/SHOW HOURLY RATES ON THE ROSTER TAB]
    Dim intCounter          As Integer
    '[MAKE RATE VISIBLE IF HOURLY PAY RATE IS SET AND CLASS IS CHECKED]
    For intCounter = 0 To 9
        If frmStaff.optionPay(vbHourly).Value = True Then
            If frmStaff.CheckClass(intCounter).Value = vbChecked Then
                frmStaff.maskRosterRate(intCounter).Visible = True
                frmStaff.cmdBaseRate(intCounter).Visible = True
                frmStaff.lbl_hr_Rate.Visible = True
            End If
        Else
            frmStaff.maskRosterRate(intCounter).Visible = False
            frmStaff.cmdBaseRate(intCounter).Visible = False
        End If
    Next intCounter
    
    

End Sub

Private Sub PanelToolBar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "The staff form allows you to add, delete and modify staff records."
    '[---------------------------------------------------------------------------------]

End Sub



Private Sub textFinish_Change(Index As Integer)

    Dim strField            As String
    strField = "Start_" & Trim(Str(Index + 1))
    If IsDate(textStart(Index).Text) And DsStaff(strField).Value <> Format(textStart(Index).Text, "Medium Time") Then
        '[DATA HAS CHANGED FROM THAT STORED - SHOW SAVE BUTTON]
        If frmStaff.cmdSave.Visible = False Then frmStaff.cmdSave.Visible = True
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
    If frmStaff.cmdSave.Visible = False Then frmStaff.cmdSave.Visible = True

End Sub

Private Sub TextFirstName_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "First name of the staff member (or initial if more space is required)."
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub TextLastName_Change()

    '[ENABLE SAVE BUTTON IF A CHANGE HAS BEEN MADE]
    If frmStaff.cmdSave.Visible = False Then frmStaff.cmdSave.Visible = True

End Sub

Private Sub TextLastName_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Staff members last name."
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub TextMiddleName_Change()


    '[ENABLE SAVE BUTTON IF A CHANGE HAS BEEN MADE]
    If frmStaff.cmdSave.Visible = False Then frmStaff.cmdSave.Visible = True

End Sub

Private Sub TextMiddleName_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Middle name of the staff member (not essential)."
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub TextNote_Change()

    '[ENABLE SAVE BUTTON IF A CHANGE HAS BEEN MADE]
    If frmStaff.cmdSave.Visible = False Then frmStaff.cmdSave.Visible = True

End Sub

Private Sub TextNote_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Enter notes for this staff member (will be printed on the staff timesheet) - up to 255 characters."
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub TextStaffID_Change()

    '[ENABLE SAVE BUTTON IF A CHANGE HAS BEEN MADE]
    If frmStaff.cmdSave.Visible = False Then frmStaff.cmdSave.Visible = True

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
        If frmStaff.cmdSave.Visible = False Then frmStaff.cmdSave.Visible = True
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
    If IsDate(frmStaff.MaskHolStart) And Val(txtHolDays) > 0 And Val(txtHolDays) < 36500 Then
        If Not (IsDate(frmStaff.MaskHolEnd)) Then
            frmStaff.MaskHolEnd = Format(CDate(frmStaff.MaskHolStart) + (Val(txtHolDays) - 1), strDateFormat)
        Else
            If CDate(frmStaff.MaskHolEnd) <> CDate(frmStaff.MaskHolStart) + (Val(txtHolDays) - 1) Then
                frmStaff.MaskHolEnd = Format(CDate(frmStaff.MaskHolStart) + (Val(txtHolDays) - 1), strDateFormat)
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


