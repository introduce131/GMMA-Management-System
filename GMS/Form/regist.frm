VERSION 5.00
Object = "{B9411660-10E6-4A53-BE96-7FED334704FA}#7.0#0"; "fpSpru70.ocx"
Begin VB.Form regist 
   BackColor       =   &H00FFFFFF&
   Caption         =   "GMS : 학생 성적 조회/관리"
   ClientHeight    =   9660
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   17700
   ForeColor       =   &H000000FF&
   Icon            =   "regist.frx":0000
   LinkTopic       =   "Form2"
   Picture         =   "regist.frx":0442
   ScaleHeight     =   9660
   ScaleWidth      =   17700
   StartUpPosition =   3  'Windows 기본값
   Begin VB.Frame Frame_Select 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   10440
      TabIndex        =   23
      Top             =   3720
      Width           =   7095
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
         Left            =   6120
         Style           =   1  '그래픽
         TabIndex        =   24
         Top             =   210
         Width           =   855
      End
   End
   Begin VB.TextBox Text6 
      BeginProperty Font 
         Name            =   "맑은 고딕"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   7920
      TabIndex        =   17
      Top             =   1920
      Width           =   2295
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "맑은 고딕"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   2760
      TabIndex        =   15
      Top             =   480
      Width           =   2295
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "맑은 고딕"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   7680
      TabIndex        =   13
      Top             =   2760
      Width           =   2535
   End
   Begin VB.ComboBox Combo4 
      BeginProperty Font 
         Name            =   "맑은 고딕"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   7680
      TabIndex        =   11
      Top             =   1200
      Width           =   2535
   End
   Begin VB.ComboBox Combo3 
      BeginProperty Font 
         Name            =   "맑은 고딕"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   7680
      TabIndex        =   9
      Top             =   480
      Width           =   2535
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "맑은 고딕"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   2520
      TabIndex        =   7
      Top             =   3480
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "맑은 고딕"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   2520
      TabIndex        =   6
      Top             =   2760
      Width           =   2535
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "맑은 고딕"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2520
      TabIndex        =   3
      Top             =   1920
      Width           =   2535
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "맑은 고딕"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2520
      TabIndex        =   0
      ToolTipText     =   "학년을 선택해주세요."
      Top             =   1200
      Width           =   2535
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "성적 조회"
      BeginProperty Font 
         Name            =   "맑은 고딕"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4815
      Left            =   120
      TabIndex        =   18
      Top             =   4440
      Width           =   17415
      Begin FPUSpreadADO.fpSpread fpSpread1 
         Height          =   4455
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   17175
         _Version        =   458752
         _ExtentX        =   30295
         _ExtentY        =   7858
         _StockProps     =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "맑은 고딕"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   9
         SpreadDesigner  =   "regist.frx":0594
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "과목코드 및 학과코드"
      BeginProperty Font 
         Name            =   "맑은 고딕"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   10440
      TabIndex        =   19
      Top             =   240
      Width           =   7095
      Begin FPUSpreadADO.fpSpread fpSpread2 
         Height          =   3135
         Left            =   120
         TabIndex        =   21
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
         SpreadDesigner  =   "regist.frx":0862
      End
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
      Left            =   14760
      TabIndex        =   22
      Top             =   9360
      Width           =   2775
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FFFFFF&
      Caption         =   "과목코드를 입력해주세요"
      BeginProperty Font 
         Name            =   "맑은 고딕"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   5520
      TabIndex        =   16
      Top             =   1920
      Width           =   2295
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFFFF&
      Caption         =   "학과코드를 입력해주세요"
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
      Height          =   375
      Left            =   360
      TabIndex        =   14
      Top             =   480
      Width           =   2295
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFFF&
      Caption         =   "점수를 입력해주세요"
      BeginProperty Font 
         Name            =   "맑은 고딕"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5520
      TabIndex        =   12
      Top             =   2760
      Width           =   2055
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "차시를 선택해주세요"
      BeginProperty Font 
         Name            =   "맑은 고딕"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5520
      TabIndex        =   10
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "학기를 선택해주세요"
      BeginProperty Font 
         Name            =   "맑은 고딕"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5520
      TabIndex        =   8
      Top             =   480
      Width           =   2055
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "번호를 입력해주세요"
      BeginProperty Font 
         Name            =   "맑은 고딕"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   2760
      Width           =   1935
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "이름을 입력해주세요"
      BeginProperty Font 
         Name            =   "맑은 고딕"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   3480
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "반을 선택해주세요"
      BeginProperty Font 
         Name            =   "맑은 고딕"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   1920
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "학년을 선택해주세요"
      BeginProperty Font 
         Name            =   "맑은 고딕"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   1200
      Width           =   1935
   End
End
Attribute VB_Name = "regist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()        '초기화 클릭시
    Combo1.Text = ""
    Combo2.Text = ""
    Combo3.Text = ""
    Combo4.Text = ""
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    Text5.Text = ""
    Text6.Text = ""
    MsgBox "입력창을 초기화했습니다", vbOKOnly, "초기화"
End Sub

Private Sub delete_Click()
    Dim myconnobj As New ADODB.Connection
    Dim myRectSet As New ADODB.Recordset
    Dim sqlStr As String
    
    myconnobj.Open "Provider=SQLOLEDB.1;Password=qorwhddnjs23;Persist Security Info=true;User ID=sa;Initial Catalog=Grade; Data Source=IT_09\SQLEXPRESS"
    
    sqlStr = "delete from record where  이름 = '" & Trim(Text2.Text) & "'"
    
    myRectSet.Open sqlStr, myconnobj    'lstConStr, adOpenStatic, adLockReadOnly
    
    MsgBox "삭제 완료", vbOKOnly, "성적 삭제"
    myconnobj.Close
End Sub

Private Sub cmd_Select_Click()
    Dim conn As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    Dim lngRet As Long
    Dim Strsql As String
    
    DBmod.OpenConnection conn
    
    Strsql = "SELECT * FROM record"
    
    lngRet = DBmod.OpenRecordSet(rs, Strsql)
    
    Call SetRows(fpSpread1, rs)
End Sub

Private Sub Form_Load()
    Dim i As Integer
    
    ' 셀을 클릭하면 행 전체가 선택되는 코드
    fpSpread1.OperationMode = OperationModeRow
    
    '학년을 선택해주세요
    Combo1.AddItem ("1학년")
    Combo1.AddItem ("2학년")
    Combo1.AddItem ("3학년")
    
    '반을 선택해주세요
    Combo2.AddItem ("1반")
    Combo2.AddItem ("2반")
    Combo2.AddItem ("3반")
    Combo2.AddItem ("4반")
    Combo2.AddItem ("5반")
    Combo2.AddItem ("6반")
    Combo2.AddItem ("7반")
    Combo2.AddItem ("8반")
    Combo2.AddItem ("9반")
    Combo2.AddItem ("10반")
    
    '학기를 선택해주세요
    Combo3.AddItem ("1학기")
    Combo3.AddItem ("2학기")
    
    '차시를 선택해주세요
    Combo4.AddItem ("1차 지필평가")
    Combo4.AddItem ("2차 지필평가")
    Combo4.AddItem ("수행 평가")
    
    'fpspread2를 전체 열40, 행40으로 줄인다
    fpSpread2.MaxCols = 40
    fpSpread2.MaxRows = 40
    
    'fpspread2(학과코드 및 과목코드)의 제목 열 표시
    fpSpread2.Row = 0
    fpSpread2.Col = 1:      fpSpread2.Value = "과목 코드"
    fpSpread2.Col = 2:      fpSpread2.Value = "교과목 명"
    fpSpread2.Col = 3:      fpSpread2.Value = "학과 코드"
    fpSpread2.Col = 4:      fpSpread2.Value = "학 과  명"
    
    '과목명 필드
    fpSpread2.Col = 2
    fpSpread2.Row = 1:      fpSpread2.Value = "국어"
    fpSpread2.Row = 2:      fpSpread2.Value = "수학"
    fpSpread2.Row = 3:      fpSpread2.Value = "사회"
    fpSpread2.Row = 4:      fpSpread2.Value = "과학"
    fpSpread2.Row = 5:      fpSpread2.Value = "실용 영어"
    fpSpread2.Row = 6:      fpSpread2.Value = "한국사"
    fpSpread2.Row = 7:      fpSpread2.Value = "중국어"
    fpSpread2.Row = 8:      fpSpread2.Value = "창업"
    fpSpread2.Row = 9:      fpSpread2.Value = "화법과 작문"
    fpSpread2.Row = 10:     fpSpread2.Value = "회계 원리"
    fpSpread2.Row = 11:     fpSpread2.Value = "상업 경제"
    fpSpread2.Row = 12:     fpSpread2.Value = "컴퓨터 일반"
    fpSpread2.Row = 13:     fpSpread2.Value = "운 동"
    fpSpread2.Row = 14:     fpSpread2.Value = "프로그래밍 실무"
    fpSpread2.Row = 15:     fpSpread2.Value = "회계 실무"
    fpSpread2.Row = 16:     fpSpread2.Value = "음 악"
    fpSpread2.Row = 17:     fpSpread2.Value = "미 술"
    fpSpread2.Row = 18:     fpSpread2.Value = "전자계산 실무"
    fpSpread2.Row = 19:     fpSpread2.Value = "컴퓨터 그래픽"
    fpSpread2.Row = 20:     fpSpread2.Value = "금융 일반"
    fpSpread2.Row = 21:     fpSpread2.Value = "세무 실무"
    
    ' 학과코드
    fpSpread2.Row = 1
    fpSpread2.Col = 3:      fpSpread2.Value = "HG01"
    fpSpread2.ForeColor = vbRed
    fpSpread2.Row = 2:      fpSpread2.Value = "HG02"
    fpSpread2.ForeColor = vbRed
    fpSpread2.Row = 3:      fpSpread2.Value = "HG03"
    fpSpread2.ForeColor = vbRed
    
    ' 학과 명
    fpSpread2.Row = 1
    fpSpread2.Col = 4:      fpSpread2.Value = "금융경영과"
    fpSpread2.Row = 2:      fpSpread2.Value = "세무회계과"
    fpSpread2.Row = 3:      fpSpread2.Value = "회계정보과"
    
    '열 크기 조절
    fpSpread2.ColWidth(2) = 13      ' fpspread2의 "과목 명" 열의 크기를 13으로 늘림
    fpSpread1.ColWidth(7) = 13      ' fpspread1의 "차시" 열의 크기를 13으로 늘림
    
    'fpspread1 (성적조회) 의 제목 열 표시
    fpSpread1.Row = 0
    fpSpread1.Col = 1:      fpSpread1.Value = "학과코드"
    fpSpread1.Col = 2:      fpSpread1.Value = "학년"
    fpSpread1.Col = 3:      fpSpread1.Value = "반"
    fpSpread1.Col = 4:      fpSpread1.Value = "번호"
    fpSpread1.Col = 5:      fpSpread1.Value = "이름"
    fpSpread1.Col = 6:      fpSpread1.Value = "학기"
    fpSpread1.Col = 7:      fpSpread1.Value = "차시"
    fpSpread1.Col = 8:      fpSpread1.Value = "과목코드"
    fpSpread1.Col = 9:      fpSpread1.Value = "점수"
    
    ' 과목코드 내용 표시
    For i = 1 To 21 Step 1
        fpSpread2.Col = 1
        fpSpread2.Row = i
        If i < 10 Then                          ' AA01~AA09까지 표시
            fpSpread2.Value = "AA" & "0" & i
            fpSpread2.ForeColor = vbBlue
        Else                                    ' AA10~AA21까지 표시
            fpSpread2.Value = "AA" & i
            fpSpread2.ForeColor = vbBlue
        End If
    Next i
End Sub

Private Sub insert_Click()      '추가 버튼 클릭시
    Dim myconnobj As New ADODB.Connection
    Dim myRectSet As New ADODB.Recordset
    Dim sqlStr As String
    
    myconnobj.Open "Provider=SQLOLEDB.1;Password=qorwhddnjs23;Persist Security Info=true;User ID=sa;Initial Catalog=Grade; Data Source=IT_09\SQLEXPRESS"
    
    If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Or Text5.Text = "" Or Text6.Text = "" Or Combo1.Text = "" Or Combo2.Text = "" Or Combo3.Text = "" Or Combo4.Text = "" Then
        MsgBox "모든 항목에 값을 넣어주세요", vbCritical, "추가 오류"
    Else
            sqlStr = "insert into record(학과코드, 학년, 반, 번호, 이름, 학기, 차시, 과목코드, 점수) values('" & Trim(Text3.Text) & "','" & Trim(Combo1.Text) & "','" & Trim(Combo2.Text) & "','" & Trim(Text1.Text) & "','" & Trim(Text2.Text) & "','" & Trim(Combo3.Text) & "','" & Trim(Combo4.Text) & "', '" & Trim(Text6.Text) & "', '" & Trim(Text5.Text) & "')"
            myRectSet.Open sqlStr, myconnobj    'lstConStr, adOpenStatic, adLockReadOnly
             MsgBox "추가완료", vbOKOnly, "성적추가"
    End If
    myconnobj.Close
End Sub
