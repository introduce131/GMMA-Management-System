VERSION 5.00
Object = "{B9411660-10E6-4A53-BE96-7FED334704FA}#7.0#0"; "fpSpru70.ocx"
Begin VB.Form form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  '단일 고정
   Caption         =   "GMS Main"
   ClientHeight    =   10485
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   22455
   BeginProperty Font 
      Name            =   "맑은 고딕"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10485
   ScaleWidth      =   22455
   StartUpPosition =   2  '화면 가운데
   Begin VB.Timer Timer1 
      Left            =   240
      Top             =   360
   End
   Begin FPUSpreadADO.fpSpread MenuSpread 
      Height          =   4005
      Left            =   16200
      TabIndex        =   2
      Top             =   360
      Width           =   6105
      _Version        =   458752
      _ExtentX        =   10769
      _ExtentY        =   7064
      _StockProps     =   64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "맑은 고딕"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   1
      OperationMode   =   2
      SpreadDesigner  =   "form1.frx":08FF
   End
   Begin VB.Label Label2 
      Alignment       =   1  '오른쪽 맞춤
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  '투명
      Caption         =   "12:59"
      BeginProperty Font 
         Name            =   "맑은 고딕"
         Size            =   65.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1695
      Left            =   -1440
      TabIndex        =   1
      Top             =   6600
      Width           =   5895
   End
   Begin VB.Label NameLabel 
      Alignment       =   1  '오른쪽 맞춤
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  '투명
      Caption         =   "Have a Good Day! 백종원"
      BeginProperty Font 
         Name            =   "맑은 고딕"
         Size            =   36
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   1080
      TabIndex        =   0
      Top             =   8640
      Width           =   8895
   End
   Begin VB.Image Image1 
      Height          =   10500
      Left            =   0
      Picture         =   "form1.frx":0C63
      Top             =   0
      Width           =   22500
   End
End
Attribute VB_Name = "form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load() '--Form_Load
    Timer1.Enabled = True   '--타이머 실행
    Timer1.Interval = 1000
End Sub

Private Sub Form_Unload(cancel As Integer)
    If MsgBox("종료하시겠습니까?", vbYesNo, "종료") = vbYes Then
       End
    Else
        cancel = 1
    End If
End Sub

Private Sub MenuSpread_Click(ByVal Col As Long, ByVal Row As Long)
    Select Case Row
        Case 1  '--학생성적 조회/관리 메뉴
            regist.Show
        Case 2  '--학급일정 및 시간표 메뉴
            Classtime.Show
    End Select
End Sub

Private Sub Timer1_Timer()  '--1초마다 시간을 비교하여 배경을 변경해준다.
    Dim Hour As Date
    
    Label2.Caption = "" & Format(Now, "Short Time") '--Label2.Caption에 "HH:MM" 형식으로 표시해준다.
    Hour = Format(Time, "h")    '--Format함수를 이용해 시간(H)만 자르고 Hour변수에 저장한다.
    
    If Hour >= 6 Or Hour <= 12 Then         '--아침6시 ~ 오후 1시
        Image1.Picture = LoadPicture(App.Path & "\pictures\Morning.jpg")
    End If
    If Hour > 12 And Hour <= 19 Then    '--오후 2시 ~ 오후 7시
        Image1.Picture = LoadPicture(App.Path & "\pictures\Evening.jpg")
    End If
    If Hour >= 20 Or Hour <= 6 Then        '--오후 8시 ~ 새벽5시
        Image1.Picture = LoadPicture(App.Path & "\pictures\Night.jpg")
    End If
End Sub
