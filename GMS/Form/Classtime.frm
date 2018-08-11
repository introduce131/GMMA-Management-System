VERSION 5.00
Object = "{B9411660-10E6-4A53-BE96-7FED334704FA}#7.0#0"; "fpSpru70.ocx"
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Begin VB.Form Classtime 
   BackColor       =   &H00FFFFFF&
   Caption         =   "GMS : [학급 일정 및 시간표]"
   ClientHeight    =   10980
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   19710
   LinkTopic       =   "Form2"
   ScaleHeight     =   10980
   ScaleWidth      =   19710
   StartUpPosition =   2  '화면 가운데
   Begin MSACAL.Calendar Show_Calendar 
      Height          =   3495
      Left            =   15960
      TabIndex        =   5
      Top             =   3960
      Width           =   3615
      _Version        =   524288
      _ExtentX        =   6376
      _ExtentY        =   6165
      _StockProps     =   1
      BackColor       =   16777215
      Year            =   2018
      Month           =   1
      Day             =   1
      DayLength       =   1
      MonthLength     =   0
      DayFontColor    =   0
      FirstDay        =   1
      GridCellEffect  =   1
      GridFontColor   =   12582912
      GridLinesColor  =   8421504
      ShowDateSelectors=   -1  'True
      ShowDays        =   -1  'True
      ShowHorizontalGrid=   -1  'True
      ShowTitle       =   -1  'True
      ShowVerticalGrid=   -1  'True
      TitleFontColor  =   10485760
      ValueIsNull     =   0   'False
      BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Ebrima"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Ebrima"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Ebrima"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton Cmd_Save 
      BackColor       =   &H00FFFFFF&
      Caption         =   "시간표 저장"
      BeginProperty Font 
         Name            =   "맑은 고딕"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   14640
      Style           =   1  '그래픽
      TabIndex        =   4
      ToolTipText     =   "시간표의 변경된 내용을 수정합니다"
      Top             =   7560
      Width           =   1215
   End
   Begin FPUSpreadADO.fpSpread MainTime 
      Height          =   6855
      Left            =   360
      TabIndex        =   2
      Top             =   400
      Width           =   15495
      _Version        =   458752
      _ExtentX        =   27331
      _ExtentY        =   12091
      _StockProps     =   64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "맑은 고딕"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   7
      MaxRows         =   7
      ScrollBars      =   0
      SpreadDesigner  =   "Classtime.frx":0000
   End
   Begin FPUSpreadADO.fpSpread Basicspread 
      Height          =   3315
      Left            =   16245
      TabIndex        =   0
      Top             =   390
      Width           =   3135
      _Version        =   458752
      _ExtentX        =   5530
      _ExtentY        =   5847
      _StockProps     =   64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "맑은 고딕"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   2
      MaxRows         =   10
      ScrollBars      =   0
      SpreadDesigner  =   "Classtime.frx":0425
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "기본 수업 시간표"
      BeginProperty Font 
         Name            =   "맑은 고딕"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   16080
      TabIndex        =   1
      Top             =   120
      Width           =   3495
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "맑은 고딕"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   7335
      Left            =   240
      TabIndex        =   3
      Top             =   90
      Width           =   15735
   End
End
Attribute VB_Name = "Classtime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Show_Calendar.Value = Format(Now, "YYYY-MM-DD")
End Sub
