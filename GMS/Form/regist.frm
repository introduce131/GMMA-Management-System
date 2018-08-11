VERSION 5.00
Object = "{B9411660-10E6-4A53-BE96-7FED334704FA}#7.0#0"; "fpSpru70.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form regist 
   BackColor       =   &H00FFFFFF&
   Caption         =   "GMS : [성적 조회/관리]"
   ClientHeight    =   9330
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   17700
   ForeColor       =   &H000000FF&
   Icon            =   "regist.frx":0000
   LinkTopic       =   "Form2"
   Picture         =   "regist.frx":0442
   ScaleHeight     =   9330
   ScaleWidth      =   17700
   StartUpPosition =   3  'Windows 기본값
   Begin VB.TextBox txt_score 
      Alignment       =   1  '오른쪽 맞춤
      Height          =   270
      Left            =   1440
      MaxLength       =   3
      TabIndex        =   31
      Top             =   2280
      Width           =   855
   End
   Begin VB.Frame Frame_Select 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   9980
      TabIndex        =   28
      Top             =   3720
      Width           =   7545
      Begin VB.CommandButton cmd_Select 
         BackColor       =   &H00FFFFFF&
         Caption         =   "조회"
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6645
         Style           =   1  '그래픽
         TabIndex        =   32
         Top             =   210
         Width           =   735
      End
   End
   Begin MSComCtl2.DTPicker dt_insert 
      Height          =   345
      Left            =   4440
      TabIndex        =   22
      Top             =   2280
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "맑은 고딕"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarBackColor=   0
      CalendarForeColor=   16777215
      CalendarTitleBackColor=   0
      CalendarTitleForeColor=   16777215
      Format          =   123863040
      CurrentDate     =   43314
      MaxDate         =   73415
      MinDate         =   2
   End
   Begin VB.ComboBox cbo_time 
      Height          =   300
      ItemData        =   "regist.frx":0594
      Left            =   7800
      List            =   "regist.frx":059E
      Style           =   2  '드롭다운 목록
      TabIndex        =   21
      Top             =   1560
      Width           =   2295
   End
   Begin VB.ComboBox cbo_seme 
      Height          =   300
      ItemData        =   "regist.frx":05BE
      Left            =   5160
      List            =   "regist.frx":05C8
      Style           =   2  '드롭다운 목록
      TabIndex        =   19
      Top             =   1560
      Width           =   1335
   End
   Begin VB.TextBox txt_name 
      Height          =   270
      Left            =   8760
      MaxLength       =   5
      TabIndex        =   17
      Top             =   840
      Width           =   1335
   End
   Begin VB.ComboBox cbo_num 
      Height          =   300
      ItemData        =   "regist.frx":05DA
      Left            =   6960
      List            =   "regist.frx":063E
      Style           =   2  '드롭다운 목록
      TabIndex        =   15
      Top             =   840
      Width           =   615
   End
   Begin VB.ComboBox cbo_subject 
      Height          =   300
      ItemData        =   "regist.frx":06B9
      Left            =   1320
      List            =   "regist.frx":071D
      Style           =   2  '드롭다운 목록
      TabIndex        =   14
      Top             =   1560
      Width           =   3015
   End
   Begin VB.ComboBox cbo_class 
      Height          =   300
      ItemData        =   "regist.frx":0966
      Left            =   5880
      List            =   "regist.frx":0988
      Style           =   2  '드롭다운 목록
      TabIndex        =   11
      Top             =   840
      Width           =   615
   End
   Begin VB.ComboBox cbo_grade 
      Height          =   300
      ItemData        =   "regist.frx":09AB
      Left            =   4560
      List            =   "regist.frx":09B8
      Style           =   2  '드롭다운 목록
      TabIndex        =   8
      Top             =   840
      Width           =   615
   End
   Begin VB.ComboBox cbo_depart 
      Height          =   300
      ItemData        =   "regist.frx":09C5
      Left            =   1320
      List            =   "regist.frx":09D8
      Style           =   2  '드롭다운 목록
      TabIndex        =   7
      Top             =   840
      Width           =   2775
   End
   Begin VB.Frame Frame_Time 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   5480
      TabIndex        =   6
      Top             =   3720
      Width           =   4455
      Begin MSComCtl2.DTPicker dt_from 
         Height          =   330
         Left            =   1160
         TabIndex        =   27
         Top             =   240
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "맑은 고딕"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   0
         CalendarForeColor=   16777215
         CalendarTitleBackColor=   0
         CalendarTitleForeColor=   16777215
         Format          =   123863041
         CurrentDate     =   43315
      End
      Begin MSComCtl2.DTPicker dt_unto 
         Height          =   330
         Left            =   2960
         TabIndex        =   24
         Top             =   240
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "맑은 고딕"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   0
         CalendarForeColor=   16777215
         CalendarTitleBackColor=   0
         CalendarTitleForeColor=   16777215
         Format          =   123863041
         CurrentDate     =   43315
      End
      Begin VB.Label lbl_design 
         BackColor       =   &H00FFFFFF&
         Caption         =   "~"
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   10
         Left            =   2690
         TabIndex        =   26
         Top             =   260
         Width           =   135
      End
      Begin VB.Label lbl_design 
         BackColor       =   &H00FFFFFF&
         Caption         =   "조회 기간"
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   160
         TabIndex        =   25
         Top             =   285
         Width           =   855
      End
   End
   Begin VB.Frame Select_Result 
      BackColor       =   &H00FFFFFF&
      Caption         =   "성적 조회"
      BeginProperty Font 
         Name            =   "맑은 고딕"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   4815
      Left            =   5470
      TabIndex        =   1
      Top             =   4440
      Width           =   12060
      Begin FPUSpreadADO.fpSpread fpSpread1 
         Height          =   4455
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   11805
         _Version        =   458752
         _ExtentX        =   20823
         _ExtentY        =   7858
         _StockProps     =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "맑은 고딕"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   11
         SpreadDesigner  =   "regist.frx":0A4D
      End
   End
   Begin VB.Frame Manage_Code 
      BackColor       =   &H00FFFFFF&
      Caption         =   "과목코드 및 학과코드"
      BeginProperty Font 
         Name            =   "맑은 고딕"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   3495
      Left            =   10440
      TabIndex        =   2
      Top             =   240
      Width           =   7095
      Begin FPUSpreadADO.fpSpread CodeSpread 
         Height          =   3135
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   6855
         _Version        =   458752
         _ExtentX        =   12091
         _ExtentY        =   5530
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
         MaxCols         =   4
         MaxRows         =   32
         SpreadDesigner  =   "regist.frx":0FB3
      End
   End
   Begin VB.Label lbl_design 
      BackColor       =   &H00FFFFFF&
      Caption         =   "점"
      BeginProperty Font 
         Name            =   "맑은 고딕"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   12
      Left            =   2400
      TabIndex        =   30
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label lbl_design 
      BackColor       =   &H00FFFFFF&
      Caption         =   "점수 입력"
      BeginProperty Font 
         Name            =   "맑은 고딕"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   11
      Left            =   360
      TabIndex        =   29
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label lbl_design 
      BackColor       =   &H00FFFFFF&
      Caption         =   "성적 등록일"
      BeginProperty Font 
         Name            =   "맑은 고딕"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   9
      Left            =   3240
      TabIndex        =   23
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label lbl_design 
      BackColor       =   &H00FFFFFF&
      Caption         =   "시험 차시"
      BeginProperty Font 
         Name            =   "맑은 고딕"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   8
      Left            =   6720
      TabIndex        =   20
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label lbl_design 
      BackColor       =   &H00FFFFFF&
      Caption         =   "학기"
      BeginProperty Font 
         Name            =   "맑은 고딕"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   7
      Left            =   4560
      TabIndex        =   18
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label lbl_design 
      BackColor       =   &H00FFFFFF&
      Caption         =   "이름"
      BeginProperty Font 
         Name            =   "맑은 고딕"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   6
      Left            =   8160
      TabIndex        =   16
      Top             =   840
      Width           =   495
   End
   Begin VB.Label lbl_design 
      BackColor       =   &H00FFFFFF&
      Caption         =   "과목코드"
      BeginProperty Font 
         Name            =   "맑은 고딕"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   5
      Left            =   360
      TabIndex        =   13
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label lbl_design 
      BackColor       =   &H00FFFFFF&
      Caption         =   "번"
      BeginProperty Font 
         Name            =   "맑은 고딕"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   4
      Left            =   7680
      TabIndex        =   12
      Top             =   840
      Width           =   255
   End
   Begin VB.Label lbl_design 
      BackColor       =   &H00FFFFFF&
      Caption         =   "반"
      BeginProperty Font 
         Name            =   "맑은 고딕"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   2
      Left            =   6600
      TabIndex        =   10
      Top             =   840
      Width           =   255
   End
   Begin VB.Label lbl_design 
      BackColor       =   &H00FFFFFF&
      Caption         =   "학년"
      BeginProperty Font 
         Name            =   "맑은 고딕"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   3
      Left            =   5280
      TabIndex        =   9
      Top             =   840
      Width           =   495
   End
   Begin VB.Label lbl_rst 
      Alignment       =   1  '오른쪽 맞춤
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "맑은 고딕"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   7440
      TabIndex        =   5
      Top             =   360
      Width           =   2775
   End
   Begin VB.Label lbl_design 
      BackColor       =   &H00FFFFFF&
      Caption         =   "학과코드"
      BeginProperty Font 
         Name            =   "맑은 고딕"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   0
      Top             =   840
      Width           =   855
   End
End
Attribute VB_Name = "regist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_Select_Click()
    Dim conn As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    Dim lngRet As Long
    Dim Strsql As String
    On Error GoTo SelectHandler
    
    DBmod.OpenConnection conn
    
    Strsql = "SELECT B.NM_DEPART, A.ST_GRADE, A.ST_CLASS, A.ST_NUMBER, A.ST_NAME, A.ST_SEME, A.ST_TIME, " & vbCrLf
    Strsql = Strsql & "C.NM_SUBJECT, A.ST_SCORE, A.DT_INSERT, A.NM_INSERT " & vbCrLf
    Strsql = Strsql & "FROM TABLE_RECORD AS A " & vbCrLf
    Strsql = Strsql & "LEFT JOIN TABLE_DEPART_CODE AS B " & vbCrLf
    Strsql = Strsql & "ON A.CD_DEPART = B.CD_DEPART " & vbCrLf
    Strsql = Strsql & "LEFT JOIN TABLE_SUBJECT_CODE AS C " & vbCrLf
    Strsql = Strsql & "ON A.CD_SUBJECT = C.CD_SUBJECT" & vbCrLf
    
    lngRet = DBmod.OpenRecordSet(rs, Strsql)
    
    Call SetRows(fpSpread1, rs)
    
    lbl_rst.Caption = "데이터 " & lngRet & "건 조회되었습니다."
    MsgBox "데이터 " & lngRet & "건 조회되었습니다."

SelectHandler:
    If Err.Number <> 0 Then MsgBox Err.Description, vbCritical, "ErrNum : " & Err.Number
End Sub

Private Sub Form_Load()
    dt_insert.Value = Format(Now, "YYYY-MM-DD") '--현재 시스템이 등록된 오늘날짜를 "YYYY-MM-DD"형식으로 변환하고 dt_insert 값에 넣는다.
    
    CodeSpread.Lock = True  '--코드 Spread 셀 수정금지
    fpSpread1.Lock = True   '--성적조회 Spread 셀 수정금지
End Sub
