VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0842D103-1E19-101B-9AAF-1A1626551E7C}#1.0#0"; "GRAPH32.OCX"
Object = "{A8B3B723-0B5A-101B-B22E-00AA0037B2FC}#1.0#0"; "GRID32.OCX"
Object = "{008BBE7B-C096-11D0-B4E3-00A0C901D681}#1.0#0"; "TEECHART.OCX"
Begin VB.Form frmGraph 
   AutoRedraw      =   -1  'True
   Caption         =   "Roster Breakdown Report"
   ClientHeight    =   6315
   ClientLeft      =   2715
   ClientTop       =   3585
   ClientWidth     =   8835
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FRMGRAPH.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6315
   ScaleWidth      =   8835
   Begin TeeChart.TChart chartRoster 
      Align           =   2  'Align Bottom
      Height          =   4095
      Left            =   0
      OleObjectBlob   =   "FRMGRAPH.frx":0442
      TabIndex        =   11
      Top             =   2220
      Width           =   8835
   End
   Begin MSGrid.Grid GridGraph 
      Height          =   1575
      Left            =   60
      TabIndex        =   10
      Top             =   600
      Width           =   8715
      _Version        =   65536
      _ExtentX        =   15372
      _ExtentY        =   2778
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
   End
   Begin Threed.SSPanel PanelToolBar 
      Align           =   1  'Align Top
      Height          =   520
      Left            =   0
      Negotiate       =   -1  'True
      TabIndex        =   5
      Top             =   0
      Width           =   8835
      _Version        =   65536
      _ExtentX        =   15584
      _ExtentY        =   926
      _StockProps     =   15
      ForeColor       =   -2147483641
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   1
      FloodType       =   1
      FloodColor      =   -2147483646
      FloodShowPct    =   0   'False
      Alignment       =   0
      MouseIcon       =   "FRMGRAPH.frx":050C
      Begin VB.ComboBox ComboView 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         ItemData        =   "FRMGRAPH.frx":0DE6
         Left            =   5100
         List            =   "FRMGRAPH.frx":0DF3
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   180
         Width           =   1635
      End
      Begin VB.ComboBox ComboType 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         ItemData        =   "FRMGRAPH.frx":0E0A
         Left            =   60
         List            =   "FRMGRAPH.frx":0E0C
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   180
         Width           =   1635
      End
      Begin VB.ComboBox ComboData 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         ItemData        =   "FRMGRAPH.frx":0E0E
         Left            =   3420
         List            =   "FRMGRAPH.frx":0E1B
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   180
         Width           =   1635
      End
      Begin VB.ComboBox ComboStyle 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         ItemData        =   "FRMGRAPH.frx":0E38
         Left            =   1740
         List            =   "FRMGRAPH.frx":0E4E
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   180
         Width           =   1635
      End
      Begin VB.Label labelComboDesc 
         BackStyle       =   0  'Transparent
         Caption         =   "Display Select"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   600
         Index           =   3
         Left            =   5130
         TabIndex        =   9
         Top             =   10
         Width           =   1700
      End
      Begin VB.Label labelComboDesc 
         BackStyle       =   0  'Transparent
         Caption         =   "Graph Data"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   600
         Index           =   2
         Left            =   3450
         TabIndex        =   8
         Top             =   10
         Width           =   1700
      End
      Begin VB.Label labelComboDesc 
         BackStyle       =   0  'Transparent
         Caption         =   "Graph Style"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   600
         Index           =   1
         Left            =   1770
         TabIndex        =   7
         Top             =   10
         Width           =   1700
      End
      Begin VB.Label labelComboDesc 
         BackStyle       =   0  'Transparent
         Caption         =   "Graph Type"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   600
         Index           =   0
         Left            =   90
         TabIndex        =   6
         Top             =   10
         Width           =   1700
      End
      Begin VB.Image cmdLines 
         Height          =   300
         Left            =   7980
         Picture         =   "FRMGRAPH.frx":0E9D
         Tag             =   "ON"
         Top             =   150
         Width           =   300
      End
      Begin VB.Image cmdRoster 
         Height          =   300
         Left            =   6780
         Picture         =   "FRMGRAPH.frx":1471
         Top             =   150
         Width           =   300
      End
      Begin VB.Image cmdColor 
         Height          =   300
         Left            =   7680
         Picture         =   "FRMGRAPH.frx":1743
         Top             =   150
         Width           =   300
      End
      Begin VB.Image cmdSave 
         Height          =   300
         Left            =   7380
         Picture         =   "FRMGRAPH.frx":1895
         Top             =   150
         Width           =   300
      End
      Begin VB.Image cmdPrint 
         Height          =   300
         Left            =   7080
         Picture         =   "FRMGRAPH.frx":19E7
         Top             =   150
         Width           =   300
      End
   End
   Begin GraphLib.Graph graphRoster 
      Height          =   4065
      Left            =   60
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2220
      Width           =   8715
      _Version        =   65536
      _ExtentX        =   15372
      _ExtentY        =   7170
      _StockProps     =   96
      BorderStyle     =   1
      AutoInc         =   0
      Foreground      =   0
      GraphStyle      =   2
      GridStyle       =   1
      ImageFile       =   "graph.bmp"
      IndexStyle      =   1
      NumPoints       =   7
      PatternedLines  =   1
      PrintStyle      =   2
      RandomData      =   0
      ThickLines      =   0
      Ticks           =   0
      YAxisStyle      =   1
      YAxisTicks      =   2
      ColorData       =   1
      ExtraData       =   0
      ExtraData[]     =   0
      FontFamily      =   4
      FontFamily[0]   =   1
      FontFamily[1]   =   1
      FontFamily[2]   =   1
      FontFamily[3]   =   1
      FontSize        =   4
      FontSize[0]     =   100
      FontSize[1]     =   120
      FontSize[2]     =   100
      FontSize[3]     =   100
      FontStyle       =   4
      FontStyle[0]    =   2
      FontStyle[1]    =   2
      GraphData       =   1
      GraphData[]     =   7
      GraphData[0,0]  =   0
      GraphData[0,1]  =   0
      GraphData[0,2]  =   0
      GraphData[0,3]  =   0
      GraphData[0,4]  =   0
      GraphData[0,5]  =   0
      GraphData[0,6]  =   0
      LabelText       =   0
      LegendText      =   0
      PatternData     =   0
      SymbolData      =   1
      SymbolData[0]   =   6
      XPosData        =   0
      XPosData[]      =   0
   End
End
Attribute VB_Name = "frmGraph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'[SET UP PRIVATE ARRAYS FOR DAILY BREAKDOWN DATA]
Private sinDayCost(7)       As Single
Private sinDayHours(7)      As Single
Private sinCostHour(7)      As Single
Private arrayGraph(10)      As GraphType

Sub DisplayGraphData(intDataType)

    '[SUBROUTINE TO DISPLAY THE SELECTED SERIES OF GRAPH DATA]
    '[INTDATATYPE - 0=COST INFORMATION]
    '[              1=HOUR INFORMATION]
    Dim flagDays            As Boolean
    Dim intDayCount         As Integer
    Dim intCounter          As Integer
    Dim intSetCount         As Integer
    Dim intNumSets          As Integer
    Dim intLastRow          As Integer
    Dim intLastCol          As Integer
    '[REV: 3.00.32]-[WEEKLY TOTALS]
    Dim sinWeekTotal        As Single
    Dim sinTotalWeek        As Single
    
    '[SET FLAG FOR GRAPH PRODUCED]
    flagDays = False
    
    '[CHECK FOR NUMBER OF ACTIVE ARRAY ITEMS]
    For intCounter = 1 To 10
        If arrayGraph(intCounter).Active Then intNumSets = intNumSets + 1
    Next intCounter

    '[ERASE CONTENTS OF ARRAY]
    Erase sinDayCost
    Erase sinDayHours
    Erase sinCostHour

    '[RESET GRAPH DATA]
    If intNumSets = 0 Then intNumSets = 1
    
    '[CALL RESET GRAPH COLOURS]
    If intNumSets = 1 Then Call SetGraphColours
    
    frmGraph.graphRoster.NumSets = intNumSets
    frmGraph.graphRoster.DataReset = 9
    intSetCount = 0
    
    
    '[CLEAR GRID OF ALL DATA - SHOULDN'T BE NEEDED IF GRID RESETS]
    intLastRow = frmGraph.GridGraph.Row     '[SAVE LAST ROW POSITION]
    intLastCol = frmGraph.GridGraph.Col     '[SAVE LAST COL POSITION]
    frmGraph.GridGraph.Rows = 2
    '[FILL GRAPH AND GRID]
    frmGraph.GridGraph.Row = 0
    frmGraph.GridGraph.Col = 0
    frmGraph.GridGraph.Text = "Roster"
    frmGraph.graphRoster.ThisSet = 1
    frmGraph.graphRoster.ThisPoint = 1
    
    For intDayCount = 1 To 7
        frmGraph.graphRoster.ThisPoint = intDayCount
        frmGraph.graphRoster.LabelText = ArrayWeek(DayOfWeek(intDayCount)).ShortDay
        frmGraph.GridGraph.Col = intDayCount
        frmGraph.GridGraph.Text = ArrayWeek(DayOfWeek(intDayCount)).ShortDay
    Next intDayCount
    frmGraph.GridGraph.Col = 8
    frmGraph.GridGraph.Text = "Week Total"
    
    '[DRAW GRAPH TITLES]
    Call DrawGraphTitles(intDataType)
    
    For intCounter = 1 To 10
        sinWeekTotal = 0
        If arrayGraph(intCounter).Active Then
            '[GRAPH DATA GRID]
            intSetCount = intSetCount + 1
            frmGraph.GridGraph.Rows = frmGraph.GridGraph.Rows + 1
            frmGraph.GridGraph.Row = frmGraph.GridGraph.Rows - 2
            frmGraph.GridGraph.Col = 0
            '[SET LEGEND TEXT]
            frmGraph.GridGraph.Text = "+" & arrayGraph(intCounter).Roster
            frmGraph.graphRoster.ThisSet = intSetCount
            '[SET GRAPH COLOURS]
            frmGraph.graphRoster.LegendText = arrayGraph(intCounter).Roster
            Select Case intCounter
            Case 1
                 frmGraph.graphRoster.ColorData = 1
            Case 2
                 frmGraph.graphRoster.ColorData = 14
            Case 3
                 frmGraph.graphRoster.ColorData = 4
            Case 4
                 frmGraph.graphRoster.ColorData = 13
            Case 5
                 frmGraph.graphRoster.ColorData = 9
            Case 6
                 frmGraph.graphRoster.ColorData = 12
            Case 7
                 frmGraph.graphRoster.ColorData = 5
            Case 8
                 frmGraph.graphRoster.ColorData = 11
            Case 9
                 frmGraph.graphRoster.ColorData = 2
            Case 10
                 frmGraph.graphRoster.ColorData = 10
            End Select
            '[DAY COST VALUES FOR THE SET]
            For intDayCount = 1 To 7
                frmGraph.GridGraph.Col = intDayCount
                frmGraph.graphRoster.ThisPoint = intDayCount
                sinDayHours(intDayCount) = sinDayHours(intDayCount) + arrayGraph(intCounter).Time(intDayCount)
                sinDayCost(intDayCount) = sinDayCost(intDayCount) + arrayGraph(intCounter).Cost(intDayCount)
                Select Case intDataType
                Case 0  '[COSTS]
                    frmGraph.GridGraph.Text = Format(arrayGraph(intCounter).Cost(intDayCount), "Currency")
                    frmGraph.graphRoster.GraphData = Format(arrayGraph(intCounter).Cost(intDayCount), "Currency")
                    sinWeekTotal = sinWeekTotal + arrayGraph(intCounter).Cost(intDayCount)
                Case 1  '[HOURS]
                    frmGraph.GridGraph.Text = Format(arrayGraph(intCounter).Time(intDayCount), "####0.00")
                    frmGraph.graphRoster.GraphData = Format(arrayGraph(intCounter).Time(intDayCount), "###0.00")
                    sinWeekTotal = sinWeekTotal + arrayGraph(intCounter).Time(intDayCount)
                Case 2  '[COST/HOUR]
                    If arrayGraph(intCounter).Cost(intDayCount) > 0 And arrayGraph(intCounter).Time(intDayCount) > 0 Then
                        frmGraph.GridGraph.Text = Format(arrayGraph(intCounter).Cost(intDayCount) / arrayGraph(intCounter).Time(intDayCount), "Currency")
                        frmGraph.graphRoster.GraphData = Format(arrayGraph(intCounter).Cost(intDayCount) / arrayGraph(intCounter).Time(intDayCount), "Currency")
                        If sinDayHours(intDayCount) > 0 Then sinCostHour(intDayCount) = sinDayCost(intDayCount) / sinDayHours(intDayCount) Else sinCostHour(intDayCount) = 0
                        sinWeekTotal = sinWeekTotal + arrayGraph(intCounter).Cost(intDayCount) / arrayGraph(intCounter).Time(intDayCount)
                    Else
                        frmGraph.GridGraph.Text = Format(0, "Currency")
                        frmGraph.graphRoster.GraphData = Format(0, "Currency")
                    End If
                End Select
            Next intDayCount
            '[REV: 3.00.32]-[WEEKLY TOTALS]
            frmGraph.GridGraph.Col = 8
            Select Case intDataType
            Case 0  '[COSTS]
                frmGraph.GridGraph.Text = Format(sinWeekTotal, "Currency")
            Case 1  '[HOURS]
                frmGraph.GridGraph.Text = Format(sinWeekTotal, "####0.00")
            Case 2  '[COST/HOUR]
                sinWeekTotal = sinWeekTotal / 7
                frmGraph.GridGraph.Text = Format(sinWeekTotal, "Currency")
            End Select
            sinTotalWeek = sinTotalWeek + sinWeekTotal
            
        ElseIf arrayGraph(intCounter).Roster > "" Then
            frmGraph.GridGraph.Rows = frmGraph.GridGraph.Rows + 1
            frmGraph.GridGraph.Row = frmGraph.GridGraph.Rows - 2
            frmGraph.GridGraph.Col = 0
            frmGraph.GridGraph.Text = "-" & arrayGraph(intCounter).Roster
            '[DAY COST VALUES FOR THE SET]
            For intDayCount = 1 To 8
                frmGraph.GridGraph.Col = intDayCount
                frmGraph.GridGraph.Text = "-"
            Next intDayCount
        End If
    Next intCounter
    
    '[TOTALS]
    frmGraph.GridGraph.Row = (frmGraph.GridGraph.Rows - 1)
    frmGraph.GridGraph.Col = 0
    If intDataType <= 1 Then frmGraph.GridGraph.Text = "Totals" Else frmGraph.GridGraph.Text = "Average"
    For intCounter = 1 To 7
        frmGraph.GridGraph.Col = intCounter
        Select Case intDataType
        Case 0  '[COSTS]
            frmGraph.GridGraph.Text = Format(sinDayCost(intCounter), "Currency")
        Case 1  '[HOURS]
            frmGraph.GridGraph.Text = Format(sinDayHours(intCounter), "#####0.00")
        Case 2  '[COST/HOUR]
            frmGraph.GridGraph.Text = Format(sinCostHour(intCounter), "Currency")
        End Select
    Next intCounter
    '[REV: 3.00.32]-[WEEKLY TOTALS]
    frmGraph.GridGraph.Col = 8
    Select Case intDataType
    Case 0  '[COSTS]
        frmGraph.GridGraph.Text = Format(sinTotalWeek, "Currency")
    Case 1  '[HOURS]
        frmGraph.GridGraph.Text = Format(sinTotalWeek, "#####0.00")
    Case 2  '[COST/HOUR]
        frmGraph.GridGraph.Text = Format(sinTotalWeek / intSetCount, "Currency")
    End Select
        
    '[REFRESH GRAPH]
    Call SetGraphColours
    frmGraph.graphRoster.DrawMode = 3
    frmGraph.GridGraph.Refresh
    frmGraph.graphRoster.Refresh
    
    '[RESTORE PREVIOUS ROW POSITION]
    If frmGraph.GridGraph.Rows > intLastRow Then
        frmGraph.GridGraph.Row = intLastRow
        frmGraph.GridGraph.SelStartRow = intLastRow
        frmGraph.GridGraph.SelEndRow = intLastRow
        frmGraph.GridGraph.Col = intLastCol
    End If
    
    
End Sub

Sub DrawGraphTitles(intDataType)
    
    '[DETERMINE WHERE TITLES SHOULD BE DRAWN]
    '[DECLARE VARIABLES]
    Dim strXaxis            As String
    Dim strYaxis            As String
    
    '[DETERMINE GRAPH STYLE - HORIZONTAL OR OTHER]
    If Left$(frmGraph.ComboStyle.Text, 5) = "Horiz" Then
        If cmdLines.Tag = "ON" Then frmGraph.graphRoster.GridStyle = 2
        '[COLUMN AND POINT TITLES]
        Select Case intDataType
        Case 0  '[COST]
            strYaxis = "$"
            strXaxis = "Day"
        Case 1  '[HOURS]
            strYaxis = "Hrs"
            strXaxis = "Day"
        Case 2  '[COST/HOUR]
            strYaxis = "$/hr"
            strXaxis = "Day"
        End Select
    Else
        If cmdLines.Tag = "ON" Then frmGraph.graphRoster.GridStyle = 1
        '[COLUMN AND POINT TITLES]
        Select Case intDataType
        Case 0  '[COST]
            strXaxis = "$"
            strYaxis = "Day"
        Case 1  '[HOURS]
            strXaxis = "Hrs"
            strYaxis = "Day"
        Case 2  '[COST/HOUR]
            strXaxis = "$/hr"
            strYaxis = "Day"
        End Select
    End If

    
    '{SELECT TYPE THEN STYLE]
    frmGraph.graphRoster.LeftTitle = strXaxis
    frmGraph.graphRoster.BottomTitle = strYaxis
    
End Sub


Sub FillComboStyle(intStyleIndex)

    '[FILL GRAPH STYLE COMBO DEPENDANT UPON GRAPH TYPE]
    Select Case intStyleIndex
    Case 3      '[2D bar]
        frmGraph.ComboStyle.Clear
        frmGraph.ComboStyle.AddItem "Vertical Bars": frmGraph.ComboStyle.ItemData(frmGraph.ComboStyle.NewIndex) = 0
        frmGraph.ComboStyle.AddItem "Horizontal": frmGraph.ComboStyle.ItemData(frmGraph.ComboStyle.NewIndex) = 1
        frmGraph.ComboStyle.AddItem "Stacked": frmGraph.ComboStyle.ItemData(frmGraph.ComboStyle.NewIndex) = 2
        frmGraph.ComboStyle.AddItem "Horiz. Stacked": frmGraph.ComboStyle.ItemData(frmGraph.ComboStyle.NewIndex) = 3
        frmGraph.ComboStyle.AddItem "Stacked %": frmGraph.ComboStyle.ItemData(frmGraph.ComboStyle.NewIndex) = 4
        frmGraph.ComboStyle.AddItem "Horiz. Stacked %": frmGraph.ComboStyle.ItemData(frmGraph.ComboStyle.NewIndex) = 5
        frmGraph.ComboStyle.ListIndex = 2       '[DEFAULT TYPE]
        SetGraphColours
    Case 4      '[3D bar]
        frmGraph.ComboStyle.Clear
        frmGraph.ComboStyle.AddItem "Vertical Bars": frmGraph.ComboStyle.ItemData(frmGraph.ComboStyle.NewIndex) = 0
        frmGraph.ComboStyle.AddItem "Horizontal": frmGraph.ComboStyle.ItemData(frmGraph.ComboStyle.NewIndex) = 1
        frmGraph.ComboStyle.AddItem "Stacked": frmGraph.ComboStyle.ItemData(frmGraph.ComboStyle.NewIndex) = 2
        frmGraph.ComboStyle.AddItem "Horiz. Stacked": frmGraph.ComboStyle.ItemData(frmGraph.ComboStyle.NewIndex) = 3
        frmGraph.ComboStyle.AddItem "Stacked %": frmGraph.ComboStyle.ItemData(frmGraph.ComboStyle.NewIndex) = 4
        frmGraph.ComboStyle.AddItem "Horiz. Stacked %": frmGraph.ComboStyle.ItemData(frmGraph.ComboStyle.NewIndex) = 5
        frmGraph.ComboStyle.AddItem "Z-clustered": frmGraph.ComboStyle.ItemData(frmGraph.ComboStyle.NewIndex) = 6
        frmGraph.ComboStyle.AddItem "Horiz. Z-clustered": frmGraph.ComboStyle.ItemData(frmGraph.ComboStyle.NewIndex) = 7
        frmGraph.ComboStyle.ListIndex = 2       '[DEFAULT TYPE]
    Case 6      '[Line]
        frmGraph.ComboStyle.Clear
        frmGraph.ComboStyle.AddItem "Lines": frmGraph.ComboStyle.ItemData(frmGraph.ComboStyle.NewIndex) = 0
        frmGraph.ComboStyle.AddItem "Symbols": frmGraph.ComboStyle.ItemData(frmGraph.ComboStyle.NewIndex) = 1
        frmGraph.ComboStyle.AddItem "Sticks": frmGraph.ComboStyle.ItemData(frmGraph.ComboStyle.NewIndex) = 2
        frmGraph.ComboStyle.AddItem "Sticks, Symbols": frmGraph.ComboStyle.ItemData(frmGraph.ComboStyle.NewIndex) = 3
        frmGraph.ComboStyle.AddItem "Lines, Symbols": frmGraph.ComboStyle.ItemData(frmGraph.ComboStyle.NewIndex) = 5
        frmGraph.ComboStyle.AddItem "Lines, Sticks": frmGraph.ComboStyle.ItemData(frmGraph.ComboStyle.NewIndex) = 6
        frmGraph.ComboStyle.AddItem "Lines, Sticks, Symbols": frmGraph.ComboStyle.ItemData(frmGraph.ComboStyle.NewIndex) = 7
        frmGraph.ComboStyle.ListIndex = 0       '[DEFAULT TYPE]
    Case 8      '[Area]
        frmGraph.ComboStyle.Clear
        frmGraph.ComboStyle.AddItem "Stack": frmGraph.ComboStyle.ItemData(frmGraph.ComboStyle.NewIndex) = 0
        frmGraph.ComboStyle.AddItem "Absolute": frmGraph.ComboStyle.ItemData(frmGraph.ComboStyle.NewIndex) = 1
        frmGraph.ComboStyle.AddItem "Percentage": frmGraph.ComboStyle.ItemData(frmGraph.ComboStyle.NewIndex) = 2
        frmGraph.ComboStyle.ListIndex = 0       '[DEFAULT TYPE]
    Case Else
    End Select

End Sub

Sub FillComboType(intTypeIndex)

    '[FILL GRAPH TYPE COMBO WITH ITEMS]
    frmGraph.ComboType.Clear
    frmGraph.ComboType.AddItem "2D bar": frmGraph.ComboType.ItemData(frmGraph.ComboType.NewIndex) = 3
    frmGraph.ComboType.AddItem "3D bar": frmGraph.ComboType.ItemData(frmGraph.ComboType.NewIndex) = 4
    frmGraph.ComboType.AddItem "Line": frmGraph.ComboType.ItemData(frmGraph.ComboType.NewIndex) = 6
    frmGraph.ComboType.AddItem "Area": frmGraph.ComboType.ItemData(frmGraph.ComboType.NewIndex) = 8
    
    '[SET COMBO VALUE]
    If intTypeIndex >= 0 And intTypeIndex < 4 Then
        frmGraph.ComboType.ListIndex = intTypeIndex
    End If

End Sub

Sub PrintGraph()

    '[SET CALL TO ERROR HANDLING ROUTINE]
    'On Error GoTo ErrorHandler
    Dim strGraphFile            As String
    
    '[MESSAGE VARIABLES]
    Dim Msg
    Dim Style
    Dim Title
    Dim Response
    Dim intCounter          As Integer
    Dim sngLineHeight           '[Height of each printed line]
    Dim intLinesPerPage         '[Number of Lines which can be printed per page]
    Dim sngPageWidth            '[Width of Printable Page Area]
    Dim intLineCount            '[Number of lines already printed on this page]
    Dim intRowToPrint           '[Number of Row to Print]
    Dim intListCounter          '[File List Counter]
    Dim intPage                 '[Page Number]
    Dim flagJobStart            As Boolean
    Dim flagGraph               As Boolean
    Dim intColourStyle          As Integer
    Dim sngLeftMargin           As Single
    Dim sinRatio                As Single   '[PIC RATIO]
    Dim sinPicWidth             As Single   '[PIC WIDTH]
    Dim sinPicHeight            As Single   '[PIC HEIGHT]
    
    '[CHECK TO SEE IF A PRINTER IS ATTACHED]
    If Printers.Count = 0 Then
        Msg = "There is no default printer attached to this computer." & strBreak & strBreak & "GSR cannot print this report."
        Style = vbOKOnly                     ' Define buttons.
        Title = "No Printer Attached"
        Response = gsrMsg(Msg, Style, Title)
        Exit Sub
    ElseIf InStr(UCase(Printer.DeviceName), "GENERIC") > 0 Or InStr(UCase(Printer.DeviceName), "TEXT ONLY") > 0 Or UCase(Printer.DriverName) = "TTY" Then
        Msg = "Printer: " & Printer.DeviceName & strBreak & "Driver: " & Printer.DriverName & strBreak & strBreak & "The default printer attached to this system appears to be a Generic/Text Only printer." & strBreak & strBreak & "GSR can only output the graph and data to a graphics-capable printer."
        Style = vbOKOnly                     ' Define buttons.
        Title = "Printer Not Graphics Capable"
        Response = gsrMsg(Msg, Style, Title)
        Exit Sub
    End If
    
    '[CONFIRM PRINT OPERATION]
    flagJobStart = False
    '[SHOW WARNING FORM]
    Msg = "Do you wish to continue and print the Roster Breakdown report ?"
    Style = vbYesNo                     ' Define buttons.
    Title = "Confirmation Required"
    Response = gsrMsg(Msg, Style, Title)
    If Response = vbYes Then
            
        '[SHOW GSR MESSAGE FORM]
        Msg = "GSR is now printing the Roster Breakdown report on the designated printer :" & strBreak & Printer.DeviceName & " on " & Printer.Port & strBreak & strBreak & "Please wait ..."
        Style = vbInformation            ' Define buttons.
        Title = "Printing Roster Breakdown"
        Response = gsrMsg(Msg, Style, Title)
        frmMsg.gaugeBar.Visible = True
        frmMsg.gaugeBorder.Visible = True
        frmMsg.labelInfo.Visible = True
        frmMsg.ZOrder
        frmMsg.Refresh
        
        '[SET PRINT FONT]
        'Printer.FontName = frmGraph.GridGraph.FontName
        'Printer.FontSize = frmGraph.GridGraph.FontSize
        'Printer.FontBold = frmGraph.GridGraph.FontBold
        'Printer.FontItalic = frmGraph.GridGraph.FontItalic
    
        '[PRINTER COLOUR AND TEXT HEIGHT]
        sngLineHeight = Printer.TextHeight("Dummy String") * 1.15   '[GIVE SMALL MARGIN FOR LINE DRAWING]
        sngLeftMargin = Printer.ScaleWidth * 0.05
        sngPageWidth = Printer.ScaleWidth * 0.9
        
        '{CALCULATE IMAGE PROPORTIONS TO FIT ON PAGE]
        sinRatio = frmGraph.graphRoster.Width / frmGraph.graphRoster.Height
        If frmGraph.graphRoster.Width > sngPageWidth Then
            sinPicWidth = sngPageWidth
            sinPicHeight = frmGraph.graphRoster.Height / sinRatio
        Else
            sinPicWidth = frmGraph.graphRoster.Width
            sinPicHeight = frmGraph.graphRoster.Height
        End If
        
        intLineCount = 0
        intLinesPerPage = Int(Printer.ScaleHeight / sngLineHeight) - 3  '[NUMBER OF LINES TO PRINT PER PAGE]
        Printer.FillStyle = 1
        Printer.ForeColor = vbBlack
        intPage = 1
    
        '[CLEAR PAGE AND PRINT PAGE HEADINGS]
        Call PrintPageHeadings(sngLineHeight, intLineCount, sngPageWidth, intPage, sngLeftMargin)
        flagJobStart = True
        flagGraph = True
        
        '[SHOW PROGRESS REPORT]
        Call ReportInfo("Printing Roster Breakdown Grid", 0)

        '[------------------------------MAIN PRINTING ROUTINE------------------------------]
        For intRowToPrint = 1 To (frmGraph.GridGraph.Rows - 1)
            '[PROGRESS BAR]
            Call ReportProgressBar((intRowToPrint / frmGraph.GridGraph.Rows - 1) * 100)
            Call PrintGridRow(intRowToPrint, sngLineHeight, intLineCount, sngPageWidth, sngLeftMargin)
            '[CHECK FOR END OF PAGE]
            If intLineCount >= intLinesPerPage Then
                '[CLOSING LINE]
                Printer.Line (0, Printer.CurrentY)-(sngPageWidth, Printer.CurrentY), vbBlack
                '[MUST USE END DOC BECAUSE OF BUG? IN NEWPAGE PROC.]
                Printer.NewPage
                Printer.FontName = frmGraph.GridGraph.FontName
                Printer.FontSize = frmGraph.GridGraph.FontSize
                Printer.FontBold = frmGraph.GridGraph.FontBold
                Printer.FontItalic = frmGraph.GridGraph.FontItalic
                Printer.FillStyle = 1
                Printer.ForeColor = vbBlack
                intPage = intPage + 1
                Call PrintPageHeadings(sngLineHeight, intLineCount, sngPageWidth, intPage, sngLeftMargin)
            End If
        Next intRowToPrint
    
        '[INCREMENT LINECOUNT]
        intLineCount = intLineCount + 3
        
        If flagGraph = True Then
            '[PRINT THE CURRENT GRAPH AND DATA TO THE PRINTER]
            intColourStyle = prnColour
            If intColourStyle > 0 Then  '[PRINTER IS CONNECTED]
                If intColourStyle = vbPRCMColor Then
                    '[PRINT GRAPH IN COLOUR]
                    frmGraph.graphRoster.PrintStyle = 3 '[COLOUR WITH BORDER]
                Else
                    '[PRINT GRAPH IN BLACK AND WHITE]
                    frmGraph.graphRoster.PrintStyle = 2 '[MONOCHROME WITH BORDER]
                End If
                
                '[VER: 3.00.33]
                '[RESET FILE NAME TO GRAPH.BMP]
                strGraphFile = frmGraph.graphRoster.ImageFile
                frmGraph.graphRoster.ImageFile = "graph.bmp"
                
                frmGraph.graphRoster.DrawMode = 3
                frmGraph.graphRoster.DrawMode = 6   '[SET PICTURE OBJECT]
                
                '[VER: 3.00.33]
                '[RESET FILE NAME BACK TO ORIGINAL]
                frmGraph.graphRoster.ImageFile = strGraphFile
                
                '[DETERMINE PRINT WIDTH AND HEIGHT OF GRAPH]
                Printer.PaintPicture LoadPicture("graph.bmp"), sngLeftMargin, sngLineHeight * intLineCount, sinPicWidth, sinPicHeight
                frmGraph.graphRoster.DrawMode = 3   '[SET MODE BitBlit]
            End If
        End If
        
        '[CLOSING LINE]
        Printer.FontBold = True
        Printer.CurrentY = sngLineHeight * intLinesPerPage
        Printer.Line (sngLeftMargin, Printer.CurrentY)-(sngLeftMargin + sngPageWidth, Printer.CurrentY), vbBlack
        Printer.CurrentX = sngLeftMargin + (0)
        Printer.Print "Generic Staff Roster"                        '[PRINT PROGRAM NAME AT END OF REPORT]
        Printer.CurrentY = sngLineHeight * intLinesPerPage
        Printer.CurrentX = sngPageWidth - (Printer.TextWidth(strDateFormat))
        Printer.Print "Printed : "; Format(Date, strDateFormat)     '[PRINT CURRENT DATE ON END OF REPORT - DESIGNATED DATE FORMAT]
        Printer.FontBold = False
        
        '[CLEAR PROGRESS REPORT]
        Call ReportInfo("", 0)
        
        '[CLEAR PROGRESS BAR]
        Call ReportProgressBar(0)
        
        '[UNLOAD MESSAGE FORM]
        Unload frmMsg
    End If
    '[------------------------------MAIN PRINTING ROUTINE------------------------------]

ErrorHandler:
    If Err.Number > 0 Then
        '[DISPLAY MESSAGE FOR ERROR]
        '[CANNOT FIND ROSTER FILE ?]
        Msg = "Error: GSR cannot print this breakdown report." & strBreak & strBreak & "There may be a problem with your printer connection." & strBreak & strBreak & "Error Code: " & Err.Number
        Msg = Msg & strBreak & strBreak & "[" & Err.Description & " in " & Err.Source & "]"
        Style = vbOKOnly                     ' Define buttons.
        Title = "Cannot Produce Report"
        Response = gsrMsg(Msg, Style, Title)
    End If

    '[END PRINT JOB IF JOB STARTED]
    If flagJobStart = True Then Printer.NewPage: Printer.EndDoc


End Sub

Public Sub PrintPageHeadings(sngLineHeight, intLineCount, sngPageWidth, intPage, sngLeftMargin)

    '[--------------> X <--------------]
    '[  .
    '[  .
    '[ Y
    '[  .
    '[  .

    '[RESET LINE COUNT FOR START OF NEW PAGE]
    intLineCount = 4

    '[PRINT PAGE HEADINGS, FIORST LINE WHITE ON BLACK, REMAINING LINES STANDARD BLACK]
    Printer.FillStyle = 1
    Printer.ForeColor = vbBlack
    Printer.FontBold = True
    Printer.Line (sngLeftMargin, sngLineHeight * 3.8)-(sngLeftMargin + sngPageWidth, sngLineHeight * 5.2), vbBlack, B
    Printer.CurrentY = sngLineHeight * (intLineCount)
    Select Case frmGraph.ComboData.ItemData(frmGraph.ComboData.ListIndex)
    Case 0  '[COSTS]
        Printer.CurrentX = sngLeftMargin + (sngPageWidth / 2) - ((Printer.TextWidth("Roster Breakdown Report - Cost") / 2))          '[CENTER TITLE]
        Printer.Print "Roster Breakdown Report - Cost"
    Case 1  '[HOURS]
        Printer.CurrentX = sngLeftMargin + (sngPageWidth / 2) - ((Printer.TextWidth("Roster Breakdown Report - Hours") / 2))          '[CENTER TITLE]
        Printer.Print "Roster Breakdown Report - Hours"
    Case 2  '[COST/HOURS]
        Printer.CurrentX = sngLeftMargin + (sngPageWidth / 2) - ((Printer.TextWidth("Roster Breakdown Report - Cost/Hour") / 2))          '[CENTER TITLE]
        Printer.Print "Roster Breakdown Report - Cost/Hour"
    Case Else
    End Select
    Printer.FontBold = False
    
    '[INCREMENT LINE COUNTER AND SET X POSITION]
    intLineCount = intLineCount + 2
    '[TITLE DESCRIPTION LINES]
    '[-----REGISTERED USER-----]
    Printer.CurrentX = sngLeftMargin + (sngPageWidth - Printer.TextWidth(DsDefault!RegUser))                '[RIGHT JUSTIFY]
    Printer.CurrentY = (sngLineHeight * intLineCount)
    Printer.Print DsDefault!RegUser
    '[-----STARTING DATE OF ROSTER-----]
    Printer.CurrentX = sngLeftMargin + (0)
    Printer.CurrentY = (sngLineHeight * intLineCount)
    Printer.Print "Week Starting: " & Format(DsDefault!StartDate, strDateFormat)
    intLineCount = intLineCount + 1

    '[-----COLUMN HEADINGS]
    Call PrintGridRow(0, sngLineHeight, intLineCount, sngPageWidth, sngLeftMargin)

End Sub


Sub PrintGridRow(intRowToPrint, sngLineHeight, intLineCount, sngPageWidth, sngLeftMargin)

    Dim sngColumnWidth      '[WIDTH OF EACH COLUMN]
    Dim intColumn           '[COLUMN COUNTER]
    Dim strCelltext         '[CELL TEXT TO PRINT AND CHECK FOR LENGTH]
    Dim intStartCol         '[COLUMN TO START PRINTING (FROM frmPRINT)]
    Dim BoxState            '[FLAG FOR STATE OF PRINTING BOXES]
    
    '[PRINTING OPTIONS]
    intStartCol = 0
   
    '[PRINT THE GIVEN ROW FROM THE ROSTER BREAKDOWN GRID]
    intLineCount = intLineCount + 1

    '[SET GRAPH BREAKDOWN GRID ROW TO CORRESPOND]
    frmGraph.GridGraph.Row = intRowToPrint
    frmGraph.GridGraph.Col = 0

    '[DETERMINE COLUMN WIDTHS]
    sngColumnWidth = sngPageWidth / (frmGraph.GridGraph.Cols - intStartCol)

    '[CYCLE THROUGH COLUMNS AND PRINT]
    For intColumn = intStartCol To (frmGraph.GridGraph.Cols - 1)

        frmGraph.GridGraph.Col = intColumn

        '[BOUNDING BOX]
        Printer.FillStyle = 1
        Printer.FillColor = vbWhite
        Printer.Line (((intColumn - intStartCol) * sngColumnWidth) + sngLeftMargin, (sngLineHeight * intLineCount))-Step(sngColumnWidth, sngLineHeight), vbBlack, B
        Printer.ForeColor = vbBlack

        '[ALLOCATE CELL TEXT]
        strCelltext = frmGraph.GridGraph.Text
        '[CHECK STRING LENGTH TO MAKE SURE IT FITS INTO THE CELL.  IF NOT, TRUNCATE BY ONE CHARACTER AND TRY AGAIN]
        Do While Printer.TextWidth(strCelltext) > (sngColumnWidth * 0.95)
            strCelltext = Left$(strCelltext, Len(strCelltext) - 1)
        Loop
        '[PRINT CELL TEXT]
        Printer.CurrentX = sngLeftMargin + (((intColumn - intStartCol) + 1) * sngColumnWidth) - (sngColumnWidth * 0.05 + Printer.TextWidth(strCelltext))
        '[REV: 3.00.32]-[ADDED TOP OF LINE BUFFER SPACE]
        Printer.CurrentY = (sngLineHeight * intLineCount) + (sngLineHeight * 0.1)
        Printer.Print strCelltext

        Printer.FillStyle = 1
        Printer.ForeColor = vbBlack

    Next intColumn
    
    'Printer.Line (((intColumn - intStartCol) * sngColumnWidth), (sngLineHeight * intLineCount))-Step(0, sngLineHeight), vbBlack, B
    
End Sub


Private Sub ResizeForm()

    '[RESIZE GRAPH]
    Dim intCounter          As Integer
    If frmGraph.WindowState = 1 Or frmGraph.Width < 1600 Or frmGraph.Height < 1600 Then Exit Sub '[EXIT IF MINIMISED]
    
    Select Case frmGraph.ComboView.Text
    Case "Data"
        frmGraph.GridGraph.Left = 100
        frmGraph.GridGraph.Width = frmGraph.Width - 300
        frmGraph.GridGraph.Top = 100 + frmGraph.PanelToolBar.Height
        frmGraph.GridGraph.Height = (frmGraph.Height - 600 - frmGraph.PanelToolBar.Height)
        frmGraph.GridGraph.Visible = True
        frmGraph.graphRoster.Visible = False
        frmGraph.GridGraph.Refresh
    Case "Graph"
        frmGraph.graphRoster.Left = 100
        frmGraph.graphRoster.Width = frmGraph.Width - 300
        frmGraph.graphRoster.Top = 100 + frmGraph.PanelToolBar.Height
        frmGraph.graphRoster.Height = (frmGraph.Height - 600 - frmGraph.PanelToolBar.Height)
        frmGraph.GridGraph.Visible = False
        frmGraph.graphRoster.Visible = True
        frmGraph.graphRoster.Refresh
        frmGraph.graphRoster.DrawMode = 3
    Case "Both"
        frmGraph.GridGraph.Left = 100
        frmGraph.GridGraph.Width = frmGraph.Width - 300
        frmGraph.GridGraph.Top = 100 + frmGraph.PanelToolBar.Height
        frmGraph.GridGraph.Height = (frmGraph.Height - 600 - frmGraph.PanelToolBar.Height) * 0.33
        frmGraph.graphRoster.Left = 100
        frmGraph.graphRoster.Width = frmGraph.Width - 300
        frmGraph.graphRoster.Top = 200 + frmGraph.GridGraph.Height + frmGraph.PanelToolBar.Height
        frmGraph.graphRoster.Height = (frmGraph.Height - 600 - frmGraph.PanelToolBar.Height) * 0.66
        frmGraph.GridGraph.Visible = True
        frmGraph.graphRoster.Visible = True
        frmGraph.GridGraph.Refresh
        frmGraph.graphRoster.Refresh
        frmGraph.graphRoster.DrawMode = 3
    Case Else
    End Select
    
    For intCounter = 0 To 8
        frmGraph.GridGraph.ColWidth(intCounter) = (frmGraph.GridGraph.Width / 9.5)
        If intCounter > 0 Then
            frmGraph.GridGraph.ColAlignment(intCounter) = vbRightJustify
        End If
    Next intCounter
    
End Sub

Sub SaveGraphToFile()
    
    '[SET CALL TO ERROR HANDLING ROUTINE]
    On Error GoTo ErrorHandler

    '[MESSAGE VARIABLES]
    Dim strGraphDataFile    As String
    Dim Msg
    Dim Style
    Dim Title
    Dim Response
    Dim varDummy
    Dim FileHandle          As Integer
    Dim intRowCounter       As Integer
    Dim intColCounter       As Integer
    Dim intProgress         As Integer
    '[GET FILENAME TO SAVE TO]
    Dim strTemp             As String
    Dim strMessage          As String
    Dim strTitle            As String
    
    If strGraphFile = "" Then strGraphFile = "rb_graph"
    strMessage = "Enter the base filename (without extension).  The filename can be up to eight characters in length and should only contain alphanumeric characters." & strBreak & strBreak & "The graph image will be saved to this filename with a '.bmp' extension while the grid data will be saved to this filename with a '.txt' extension."
    strTitle = "Enter Save Filename"
    strTemp = strGraphFile
    strTemp = InputBox(strMessage, strTitle, strTemp)
    strTemp = Trim(strTemp)
    If Len(strTemp) = 0 Then Exit Sub
    If Len(strTemp) > 8 Then strTemp = Left$(strTemp, 8)
    intProgress = 1     '[FLAG FOR PROGRESS THROUGH SUBROUTINE]
    
    '[REMOVE SLASHES FROM FILENAME]
    strTemp = Replace(strTemp, "/", "-")
    strTemp = Replace(strTemp, "\", "-")
    strTemp = Replace(strTemp, " ", "_")
    strTemp = Replace(strTemp, ":", "_")
    strTemp = Replace(strTemp, "*", "_")
    strTemp = Replace(strTemp, "?", "_")
    strTemp = Replace(strTemp, "=", "_")
    strTemp = Replace(strTemp, "|", "_")
    strTemp = Replace(strTemp, ";", "_")
    strTemp = Replace(strTemp, "'", "_")
    strTemp = Replace(strTemp, "<", "_")
    strTemp = Replace(strTemp, ">", "_")
    strTemp = Replace(strTemp, "$", "_")
    strGraphFile = strTemp
    frmGraph.graphRoster.ImageFile = strGraphFile & ".bmp"
    strGraphDataFile = strGraphFile & ".txt"
    
    frmGraph.graphRoster.DrawMode = 3
    frmGraph.graphRoster.DrawMode = 6   '[SET PICTURE OBJECT - WRITE IMAGE TO DISK FILE]
    frmGraph.graphRoster.DrawMode = 3   '[SET MODE BitBlit]
    
    '[CALL ERROR HANDLING ROUTINE]
    FileHandle = OpenFile(strGraphDataFile, constFileOut)
    intProgress = 2     '[FLAG FOR PROGRESS THROUGH SUBROUTINE]
    
    '[WRITE DETAILS TO OUTPUT FILE-WRITE GRID]
    varDummy = 0
    For intRowCounter = 0 To (frmGraph.GridGraph.Rows - 1)
        frmGraph.GridGraph.Row = intRowCounter   '[SET ROW POSITION]
        For intColCounter = 0 To (frmGraph.GridGraph.Cols - 1)
            varDummy = varDummy + 1
            frmGraph.GridGraph.Col = intColCounter   '[SET COL POSITION]
            If intColCounter < (frmGraph.GridGraph.Cols - 1) Then
                Print #FileHandle, frmGraph.GridGraph.Text; ",";
            Else
                Print #FileHandle, frmGraph.GridGraph.Text
            End If
        Next intColCounter
    Next intRowCounter
    
    '[CLOSE FILE]
    Close #FileHandle
    intProgress = 3
    
    '[FILE HAS BEEN SAVED]
    Msg = "Roster Breakdown Graph saved to : " & strGraphFile & ".bmp" & strBreak & "Roster Breakdown Data saved to    : " & strGraphDataFile & strBreak & strBreak & " in " & App.Path & "." & strBreak & strBreak & "Click OK to continue."
    Style = vbOKOnly                     ' Define buttons.
    Title = "Graph and Data Saved"
    Response = gsrMsg(Msg, Style, Title)

ErrorHandler:
    If Err.Number > 0 Then
        '[DISPLAY MESSAGE FOR ERROR]
        '[CANNOT SAVE TO BITMAP FILE ?]
        If intProgress = 0 Then
            Msg = "Error: GSR cannot save the Roster Breakdown Graph to " & strGraphFile & ".bmp" & strBreak & strBreak & "You may need to free some space on your hard disk or check to see that the file is not marked read-only." & strBreak & strBreak & "Error Code: " & Err.Number
        ElseIf intProgress = 1 Then
            Msg = "Error: GSR cannot save the Roster Breakdown Data to " & strGraphDataFile & strBreak & strBreak & "You may need to free some space on your hard disk or check to see that the file is not marked read-only." & strBreak & strBreak & "Error Code: " & Err.Number
        Else
            Msg = "Error: GSR cannot save the Roster Breakdown Information" & strBreak & strBreak & "You may need to free some space on your hard disk or check to see that the files in your GSR directory are not marked read-only." & strBreak & strBreak & "Error Code: " & Err.Number
        End If
        Msg = Msg & strBreak & strBreak & "[" & Err.Description & " in " & Err.Source & "]"
        Style = vbOKOnly                     ' Define buttons.
        Title = "Cannot Save to File"
        Response = gsrMsg(Msg, Style, Title)
    End If
    
End Sub

Sub SetGraphColours()
    
    '[DECLARE FOR SET NUMBER TYPES]
    Dim intCounter          As Integer
    Dim intNumSets          As Integer
    Dim intSetCount         As Integer
    Dim strRoster           As String
    Dim intGraphStyle       As Integer
    Dim intGraphType        As Integer
    
    '[GRAPH STYLE AND TYPE]
    intGraphStyle = frmGraph.ComboStyle.ItemData(frmGraph.ComboStyle.ListIndex)
    intGraphType = frmGraph.ComboType.ItemData(frmGraph.ComboType.ListIndex)
    
    '[CHECK FOR NUMBER OF ACTIVE ARRAY ITEMS]
    For intCounter = 1 To 10
        If arrayGraph(intCounter).Active Then
            intNumSets = intNumSets + 1
            strRoster = arrayGraph(intCounter).Roster
            intSetCount = intCounter
        End If
    Next intCounter
    If intNumSets = 0 Then intNumSets = 1
    
    '[NOW DETERMINE IF SINGLE COLOURS ARE NEEDED]
    If intNumSets = 1 Then
        frmGraph.graphRoster.NumSets = 7
        For intCounter = 1 To 7
            frmGraph.graphRoster.ThisSet = intCounter
            If frmGraph.graphRoster.DrawStyle = 1 Then frmGraph.graphRoster.PatternData = 0 Else frmGraph.graphRoster.PatternData = 4
            Select Case intCounter
            Case 1
                 frmGraph.graphRoster.ColorData = 1
            Case 2
                 frmGraph.graphRoster.ColorData = 14
            Case 3
                 frmGraph.graphRoster.ColorData = 4
            Case 4
                 frmGraph.graphRoster.ColorData = 13
            Case 5
                 frmGraph.graphRoster.ColorData = 9
            Case 6
                 frmGraph.graphRoster.ColorData = 12
            Case 7
                 frmGraph.graphRoster.ColorData = 5
            End Select
            
            '[SET LEGEND TEXT]
            Select Case Str$(intGraphType) & Str$(intGraphStyle)
            Case " 3 2"
                frmGraph.graphRoster.LegendText = strRoster
            Case " 3 3"
                frmGraph.graphRoster.LegendText = strRoster
            Case " 3 4"
                frmGraph.graphRoster.LegendText = strRoster
            Case " 3 5"
                frmGraph.graphRoster.LegendText = strRoster
            Case " 4 2"
                frmGraph.graphRoster.LegendText = strRoster
            Case " 4 3"
                frmGraph.graphRoster.LegendText = strRoster
            Case " 4 4"
                frmGraph.graphRoster.LegendText = strRoster
            Case " 4 5"
                frmGraph.graphRoster.LegendText = strRoster
            Case " 4 6"
                frmGraph.graphRoster.LegendText = strRoster
            Case " 4 7"
                frmGraph.graphRoster.LegendText = strRoster
            Case Else
                frmGraph.graphRoster.LegendText = ArrayWeek(DayOfWeek(intCounter)).ShortDay
            End Select
            
        Next intCounter
        frmGraph.graphRoster.NumSets = 1
    Else
        frmGraph.graphRoster.NumSets = 10
        For intCounter = 1 To 10
            frmGraph.graphRoster.ThisSet = intCounter
            Select Case intCounter
            Case 1
                 frmGraph.graphRoster.ColorData = 1
                 If frmGraph.graphRoster.DrawStyle = 1 Then frmGraph.graphRoster.PatternData = 0 Else frmGraph.graphRoster.PatternData = 0
            Case 2
                 frmGraph.graphRoster.ColorData = 14
                 If frmGraph.graphRoster.DrawStyle = 1 Then frmGraph.graphRoster.PatternData = 0 Else frmGraph.graphRoster.PatternData = 2
            Case 3
                 frmGraph.graphRoster.ColorData = 4
                 If frmGraph.graphRoster.DrawStyle = 1 Then frmGraph.graphRoster.PatternData = 0 Else frmGraph.graphRoster.PatternData = 3
            Case 4
                 frmGraph.graphRoster.ColorData = 3
                 If frmGraph.graphRoster.DrawStyle = 1 Then frmGraph.graphRoster.PatternData = 0 Else frmGraph.graphRoster.PatternData = 4
            Case 5
                 frmGraph.graphRoster.ColorData = 2
                 If frmGraph.graphRoster.DrawStyle = 1 Then frmGraph.graphRoster.PatternData = 0 Else frmGraph.graphRoster.PatternData = 5
            Case 6
                 frmGraph.graphRoster.ColorData = 12
                 If frmGraph.graphRoster.DrawStyle = 1 Then frmGraph.graphRoster.PatternData = 0 Else frmGraph.graphRoster.PatternData = 6
            Case 7
                 frmGraph.graphRoster.ColorData = 5
                 If frmGraph.graphRoster.DrawStyle = 1 Then frmGraph.graphRoster.PatternData = 0 Else frmGraph.graphRoster.PatternData = 7
            Case 8
                 frmGraph.graphRoster.ColorData = 10
                 If frmGraph.graphRoster.DrawStyle = 1 Then frmGraph.graphRoster.PatternData = 0 Else frmGraph.graphRoster.PatternData = 2
            Case 9
                 frmGraph.graphRoster.ColorData = 7
                 If frmGraph.graphRoster.DrawStyle = 1 Then frmGraph.graphRoster.PatternData = 0 Else frmGraph.graphRoster.PatternData = 3
            Case 10
                 frmGraph.graphRoster.ColorData = 13
                 If frmGraph.graphRoster.DrawStyle = 1 Then frmGraph.graphRoster.PatternData = 0 Else frmGraph.graphRoster.PatternData = 4
            End Select
        Next intCounter
        frmGraph.graphRoster.NumSets = intNumSets
    End If
    
End Sub

Private Sub cmdColor_Click()
    
    '[ANIMATE BUTTON]
    cmdColor.BorderStyle = 1
    Delay vbDelay
    cmdColor.BorderStyle = 0

    '[TOGGLE COLOR DRAW STYLE FOR GRAPH]
    If frmGraph.graphRoster.DrawStyle = 1 Then frmGraph.graphRoster.DrawStyle = 0 Else frmGraph.graphRoster.DrawStyle = 1
    Call SetGraphColours
    frmGraph.graphRoster.DrawMode = 3
    
End Sub

Private Sub cmdColor_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Toggle between colour and monochrome display of graph data."
    '[---------------------------------------------------------------------------------]

End Sub




Private Sub cmdLines_Click()

    '[ANIMATE BUTTON]
    cmdLines.BorderStyle = 1
    Delay vbDelay
    cmdLines.BorderStyle = 0

    '[TOGGLE GRAPH GRIDLINES]
    If frmGraph.graphRoster.GridStyle = 0 Then
        '[REV: 3.00.36 - MAINTAIN LINE STATE FOR CHANGE OF GRAPH]
        If Left$(frmGraph.ComboStyle.Text, 5) = "Horiz" Then
            frmGraph.graphRoster.GridStyle = 2
        Else
            frmGraph.graphRoster.GridStyle = 1
        End If
        cmdLines.Tag = "ON"
    Else
        frmGraph.graphRoster.GridStyle = 0
        cmdLines.Tag = "OFF"
    End If
    frmGraph.graphRoster.DrawMode = 3

End Sub

Private Sub cmdLines_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Toggle graph gridlines on/off."
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub cmdPrint_Click()
    
    '[ANIMATE BUTTON]
    cmdPrint.BorderStyle = 1
    Delay vbDelay
    cmdPrint.BorderStyle = 0
    
    '[PRINT GRAPH]
    Call PrintGraph
   
End Sub

Private Sub cmdPrint_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Print the Roster Breakdown Report."
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


Private Sub cmdSave_Click()
    
    '[ANIMATE BUTTON]
    cmdSave.BorderStyle = 1
    Delay vbDelay
    cmdSave.BorderStyle = 0

    '[SAVE GRAPH TO FILE]
    SaveGraphToFile
    
End Sub

Private Sub cmdSave_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Save the displayed graph and data to disk."
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub ComboData_Click()

    '[CHANGE GRAPH DATA DISPLAY TYPE]
    Call DisplayGraphData(frmGraph.ComboData.ItemData(frmGraph.ComboData.ListIndex))
    
End Sub


Private Sub ComboData_GotFocus()
        
        '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
        StatusBar "Select the graph data set: costs, hours, or cost/hour."
        '[---------------------------------------------------------------------------------]

End Sub


Private Sub ComboStyle_Click()

    '[SET STYLE OF GRAPH]
    frmGraph.graphRoster.GraphStyle = frmGraph.ComboStyle.ItemData(frmGraph.ComboStyle.ListIndex)
    
    If frmGraph.ComboData.ListIndex < 0 Then
        Call DrawGraphTitles(0)
    Else
        Call DrawGraphTitles(frmGraph.ComboData.ItemData(frmGraph.ComboData.ListIndex))
    End If
        
    '[CHECK STYLE FOR SERIES ELSE SET GRAPH COLOURS]
    SetGraphColours
    
    '[REDRAW GRAPH]
    frmGraph.graphRoster.DrawMode = 3

End Sub


Private Sub ComboStyle_GotFocus()
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Select the graph style."
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub ComboType_Click()

    '[CHANGE GRAPH STYLE TO SUIT]
    frmGraph.graphRoster.GraphType = frmGraph.ComboType.ItemData(frmGraph.ComboType.ListIndex)
    
    '[POPULATE GRAPH STYLE COMBO RELEVANT TO INDEX ITEM SELECTED]
    Call FillComboStyle(frmGraph.ComboType.ItemData(frmGraph.ComboType.ListIndex))

End Sub

Private Sub ComboType_GotFocus()
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Select the graph type."
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub ComboView_Click()

    '[SELECT DATA ITEM TO VIEW]
    ResizeForm

End Sub


Private Sub ComboView_GotFocus()
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Select the item to view on screen - data, graph or both."
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub Form_Load()

    '[FILL INITIAL COMBO AND SET TO FIRST TYPE]
    Call FillComboType(0)
    frmGraph.ComboView.ListIndex = 2

End Sub

Private Sub Form_Resize()

    '[RESIZE GRAPH AND DATA GRID]
    Call ResizeForm
    
End Sub

Sub procRosterGraph(flagResult)

    '[SUBROUTINE TO PRODUCE A WEEKLY BREAKDOWN ROSTER GRAPH DETAILING]
    '[DAILY COSTS/HOURS PER ROSTER]
    '[REQUIRES ARRAY OF (7 DAYS) BY ([N] ROSTERS) WHERE [N] IS THE]
    '[NUMBER OF ACTIVE ROSTERS]
    
    '[REV: 3.00.28]
    '[- MAJOR CHANGES - (1) determine format and structure for adding weekly pay rate to report]
    '[                  (2) determine code and placing of code for weekly pay rate adjustments]
    
    Dim strBookmark         As String
    Dim strClassBookmark    As String
    Dim DsGraph             As Dynaset
    Dim SQLStmt             As String
    Dim intRoster           As Integer
    Dim strFullname         As String
    Dim intDay              As Integer
    Dim intStaffCount       As Integer
    Dim intStaffRecord      As Integer
    Dim strDayKey           As String
    Dim strClassKey         As String
    Dim sinHourRate         As Single
    Dim sinDayCost          As Single
    Dim sinIncrement        As Single
    Dim dateStart           As Date
    Dim dateEnd             As Date
    Dim intCounter          As Integer
    Dim intDayCount         As Integer
    Dim intRecordCount      As Integer
    Dim intClass            As Integer
    '[MESSAGE VARIABLES]
    Dim Msg
    Dim Style
    Dim Title
    Dim Response
    
    '[ERASE CONTENTS OF ARRAY]
    Erase arrayGraph
    
    '[PROCESS]
    '[declare roster dynaset (full)
    '[cycle through each roster record and check to see if the roster is active
    '[if active, all totals are added to the array index matching the roster id
    '[determine length of roster shift for rate calculation (taking into account breaks)
    '[move across days in shift, checking each staff name
    '[  check staff names for base rate for the roster
    '[  multiply staff base rate x shift length and add to array item
    '[store amount and hours for each roster for each day
    flagResult = False
    
    '[A] DECLARE FULL ROSTER DYNASET
    
    '[STORE CURRENT STAFF AND CLASS LOCATION]
    strBookmark = DsStaff.Bookmark
    strClassBookmark = DsClass.Bookmark
    
    '[REV: 3.00.28]
    '[DIM ARRAYSTAFF FOR WEEKLY PAY CALCULATIONS]
    DsStaff.MoveLast
    intStaffCount = DsStaff.RecordCount
    ReDim arrayStaff(intStaffCount) As GraphStaffType
    
    '[SELECT APPROPRIATE RECORDS]
    SQLStmt = "SELECT * FROM Roster ORDER BY CLASS, SHIFTSTART"
    Set DsGraph = DBMain.OpenRecordset(SQLStmt, dbOpenDynaset)
            
    '[CHECK CURRENT DYNASET AND PREPARE]
    If (DsGraph.EOF And DsGraph.BOF) Then
        DsGraph.Close
        flagResult = False
        Exit Sub
    End If

    '[MOVE TO FIRST RECORD IN GRAPH DYNASET]
    DsGraph.MoveLast
    DsGraph.MoveFirst
    DsStaff.MoveFirst
    
    '[SHOW GSR MESSAGE FORM]
    Msg = "GSR is now processing your roster files and producing the roster breakdown graph." & strBreak & strBreak & "The graph will detail the contribution of each roster to your daily staff costs." & strBreak & strBreak & "This may take a few minutes."
    Style = vbInformation            ' Define buttons.
    Title = "Roster Breakdown Graph"
    Response = gsrMsg(Msg, Style, Title)
    frmMsg.gaugeBar.Visible = True
    frmMsg.gaugeBorder.Visible = True
    frmMsg.labelInfo.Visible = True
    frmMsg.ZOrder
    frmMsg.Refresh
    
    '[CYCLE THROUGH DYNASET UNTIL END]
    Do While Not DsGraph.EOF
        '[CHECK TO SEE IF THE ROSTER IS ACTIVE]
        DsClass.AbsolutePosition = (DsGraph!Class - 1)
        If DsClass("Active") = vbChecked Then
            '[SET ACTIVE FLAG IN ARRAY]
            arrayGraph(DsGraph!Class).Active = True
            arrayGraph(DsGraph!Class).Roster = DsClass!Description
            
            '[SHOW REPORT INFO]
            Call ReportInfo(DsClass!Description, 0)
            strClassKey = "Rate_" & Trim(Str(DsGraph!Class))
            '[CALCULATE SHIFT LENGTH]
            If IsNull(DsGraph!ShiftStart) Then dateStart = DsDefault!StartTime Else dateStart = DsGraph!ShiftStart
            If IsNull(DsGraph!ShiftEnd) Then dateEnd = DsDefault!EndTime Else dateEnd = DsGraph!ShiftEnd
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
            '[INCREMENT COUNTER FOR PROGRESS BAR]
            intRecordCount = intRecordCount + 1
            '[CYCLE ACROSS DAYS]
            For intDayCount = 1 To 7
    
                strDayKey = "Day_" & Trim(Str(intDayCount))
                '[CHECK FOR EACH STAFF NAME IN THE ROSTER DAY CELL]
                '[CYCLE THROUGH STAFF]
                DsStaff.MoveFirst
                Do While Not DsStaff.EOF
                    intStaffRecord = DsStaff.AbsolutePosition + 1
                    strFullname = DsStaff!LastName & ", " & DsStaff!FirstName
                    If InStr(DsGraph(strDayKey), strFullname) > 0 Then
                        '[STAFF FOUND - CALCULATE HOURS AND COSTS AND ADD TO ARRAY]
                        '[SET RESULT FLAG]
                        flagResult = True
                        If Not IsNull(DsStaff(strClassKey)) Then
                            '[REV: 3.00.28]
                            '[HOURLY PAY RATE]
                            If IsNull(DsStaff!PayType) Or DsStaff!PayType = vbHourly Then
                                arrayGraph(DsGraph!Class).Cost(intDayCount) = arrayGraph(DsGraph!Class).Cost(intDayCount) + ((sinIncrement / 60) * DsStaff(strClassKey))
                            End If
                        End If
                        '[ADD TIME INCREMENT]
                        arrayGraph(DsGraph!Class).Time(intDayCount) = arrayGraph(DsGraph!Class).Time(intDayCount) + (sinIncrement / 60)
                        '[REV: 3.00.28]
                        '[CALCULATE AMOUNTS FOR STAFF ARRAY]
                        If Not (IsNull(DsStaff!PayType)) And DsStaff!PayType = vbWeekly Then
                            arrayStaff(intStaffRecord).RosterHours(DsGraph!Class, intDayCount) = arrayStaff(intStaffRecord).RosterHours(DsGraph!Class, intDayCount) + sinIncrement
                            arrayStaff(intStaffRecord).RosterTotal = arrayStaff(intStaffRecord).RosterTotal + sinIncrement
                            arrayStaff(intStaffRecord).PayRate = DsStaff!HourRate
                            arrayStaff(intStaffRecord).PayType = vbWeekly
                        End If
                    End If
                    DsStaff.MoveNext
                Loop
            Next intDayCount
        Else
            '[SET NOT ACTIVE FLAG IN ARRAY]
            arrayGraph(DsGraph!Class).Active = False
        End If
        
        '[UPDATE PROGRESS BAR]
        Call ReportProgressBar(DsGraph.PercentPosition)
        
        DsGraph.MoveNext
    Loop
    
    '[REV: 3.00.28]
    '[ADD WEEKLY WAGE ADJUSTMENTS TO THE GRAPH DATA]
    For intCounter = 1 To DsStaff.RecordCount
        If arrayStaff(intCounter).PayType = vbWeekly Then
            For intClass = 1 To 10
                For intDayCount = 1 To 7
                    sinDayCost = (arrayStaff(intCounter).RosterHours(intClass, intDayCount) / arrayStaff(intCounter).RosterTotal) * arrayStaff(intCounter).PayRate
                    arrayGraph(intClass).Cost(intDayCount) = arrayGraph(intClass).Cost(intDayCount) + sinDayCost
                Next intDayCount
            Next intClass
        End If
    Next intCounter
    
    If flagResult = True Then
        '[CHANGE GRAPH DATA DISPLAY TYPE]
        frmGraph.ComboData.ListIndex = 0
    Else    '[NO DATA TO DISPLAY - UNLOAD GRAPH]
        Unload frmGraph
    End If
    
    '[HIDE GSR MESSAGE FORM]
     Unload frmMsg
    
    '[RETURN TO SPOT IN STAFF AND CLASS DYNASETS]
    DsStaff.Bookmark = strBookmark
    DsClass.Bookmark = strClassBookmark
    
    '[CLOSE DYNASET]
    DsGraph.Close
    
    '[RETURN]

End Sub


Private Sub graphRoster_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Use the list boxes on the toolbar to change chart options."
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub GridGraph_DblClick()
    
    '[ACTIVATE/DEACTIVATE ROSTER IN GRAPH]
    Dim strTemp         As String
    Dim intRoster       As Integer
    If frmGraph.GridGraph.Row = 0 Or frmGraph.GridGraph.Rows = 3 Then Exit Sub
    
    frmGraph.GridGraph.Col = 0
    strTemp = Mid$(frmGraph.GridGraph.Text, 2)
    '[LOCATE ROSTER IN ARRAY]
    intRoster = 0
    Do While Not (arrayGraph(intRoster).Roster = strTemp)
        intRoster = intRoster + 1
        If intRoster = 10 Then Exit Do
    Loop
    '[SET TIME ON FORM]
    Select Case Left$(frmGraph.GridGraph.Text, 1)
    Case "+"   '[DEACTIVATE]
        frmGraph.GridGraph.Text = "-" & arrayGraph(intRoster).Roster
        arrayGraph(intRoster).Active = False
        '[REFRESH GRAPH DATA]
        Call DisplayGraphData(frmGraph.ComboData.ItemData(frmGraph.ComboData.ListIndex))
    Case "-"   '[ACTIVATE]
        frmGraph.GridGraph.Text = "+" & arrayGraph(intRoster).Roster
        arrayGraph(intRoster).Active = True
        '[REFRESH GRAPH DATA]
        Call DisplayGraphData(frmGraph.ComboData.ItemData(frmGraph.ComboData.ListIndex))
    End Select
    
End Sub

Private Sub GridGraph_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Double-click a grid row to toggle the display of the roster data."
    '[---------------------------------------------------------------------------------]

End Sub

Private Sub labelComboDesc_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Select Case Index
    Case 0
        '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
        StatusBar "Select the graph type."
        '[---------------------------------------------------------------------------------]
    Case 1
        '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
        StatusBar "Select the graph style."
        '[---------------------------------------------------------------------------------]
    Case 2
        '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
        StatusBar "Select the graph data set: costs, hours, or cost/hour."
        '[---------------------------------------------------------------------------------]
    Case 3
        '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
        StatusBar "Select the item to view on screen - data, graph or both."
        '[---------------------------------------------------------------------------------]
    Case Else
    End Select

End Sub


