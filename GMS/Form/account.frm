VERSION 5.00
Begin VB.Form account 
   BackColor       =   &H00FFFFFF&
   Caption         =   "GMS : Create Account Service"
   ClientHeight    =   8265
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11520
   Icon            =   "account.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   8265
   ScaleWidth      =   11520
   StartUpPosition =   3  'Windows 기본값
   Begin VB.TextBox Text4 
      Height          =   270
      Left            =   4200
      TabIndex        =   13
      Top             =   5040
      Width           =   2655
   End
   Begin VB.CommandButton check 
      BackColor       =   &H00FFFFFF&
      Caption         =   "중복 확인"
      BeginProperty Font 
         Name            =   "맑은 고딕"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7080
      Style           =   1  '그래픽
      TabIndex        =   11
      Top             =   3600
      Width           =   855
   End
   Begin VB.TextBox Text3 
      Height          =   270
      Left            =   4200
      TabIndex        =   10
      Top             =   5760
      Width           =   2655
   End
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   4200
      TabIndex        =   9
      Top             =   4320
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   4200
      TabIndex        =   8
      Top             =   3600
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "회원가입 완료"
      BeginProperty Font 
         Name            =   "맑은 고딕"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4080
      Style           =   1  '그래픽
      TabIndex        =   7
      Top             =   7080
      Width           =   2775
   End
   Begin VB.Label Label5 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00FFFFFF&
      Caption         =   "설정 비밀번호를 다시 입력해주세요"
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
      Left            =   7080
      TabIndex        =   14
      Top             =   5040
      Width           =   3615
   End
   Begin VB.Label Label4 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00FFFFFF&
      Caption         =   "비밀번호 재입력"
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
      Left            =   1440
      TabIndex        =   12
      Top             =   5040
      Width           =   2535
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "이름은 최대 5글자로 지정해주세요"
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
      Index           =   2
      Left            =   7440
      TabIndex        =   6
      Top             =   5760
      Width           =   3135
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "비밀번호는 최대 13글자로 지정해주세요"
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
      Index           =   1
      Left            =   7320
      TabIndex        =   5
      Top             =   4320
      Width           =   3615
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "최대 13글자로 지정해주세요"
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
      Index           =   0
      Left            =   8160
      TabIndex        =   4
      Top             =   3600
      Width           =   2655
   End
   Begin VB.Label Label2 
      Alignment       =   2  '가운데 맞춤
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
      Index           =   2
      Left            =   1560
      TabIndex        =   3
      Top             =   5760
      Width           =   2175
   End
   Begin VB.Label Label2 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00FFFFFF&
      Caption         =   "비밀번호를 입력해주세요"
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
      Index           =   1
      Left            =   1560
      TabIndex        =   2
      Top             =   4320
      Width           =   2415
   End
   Begin VB.Label Label2 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00FFFFFF&
      Caption         =   "ID를 입력해주세요"
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
      Index           =   0
      Left            =   1560
      TabIndex        =   1
      Top             =   3600
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00FFFFFF&
      Caption         =   "Create Account"
      BeginProperty Font 
         Name            =   "맑은 고딕"
         Size            =   36
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   11415
   End
End
Attribute VB_Name = "account"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub check_Click()      '중복확인 버튼 클릭시
    Dim i As Integer
    Dim id As String
    Dim pwd As String
    Dim user As String
    Dim strID As String
    Dim strPWD As String
    Dim myconnobj As New ADODB.Connection
    Dim myRectSet As New ADODB.Recordset
    Dim sqlStr As String
    Dim tmp As Integer
    Dim count As Integer
    
    If Text1.Text = "" Then     'ID입력란에 공백이 있을 시
        MsgBox "아이디를 입력해주세요", vbCritical, "ID를 입력!"
    Else
        myconnobj.Open "Provider=SQLOLEDB.1;Password=qorwhddnjs23;Persist Security Info=true;User ID=sa;Initial Catalog=Grade; Data Source=IT_09\SQLEXPRESS"
        
        id = Text1.Text
        
        sqlStr = "select *  from systemLogin WHERE ID = '" & Trim(Text1.Text) & "'"
        
        myRectSet.Open sqlStr, myconnobj, adOpenStatic, adLockReadOnly
        '================================================================================='--ID중복확인 IF문
         If myRectSet.RecordCount = 1 Then
            MsgBox id + "" + "은(는) 이미 있는 아이디입니다", vbCritical, "중복된 아이디"
            Text1.Text = ""     '--<중복을 방지하기 위해서 아이디를 다시적게 한다>
         ElseIf myRectSet.RecordCount < 1 Then
            MsgBox id + "" + "은(는) 사용가능한 아이디입니다", vbOKOnly, "사용 가능한 아이디"
        '================================================================================='
         End If '--중복확인 시, 공백 확인 여부 IF문
    End If
End Sub

Private Sub Command1_Click()        ' 회원가입 완료 클릭시
    Dim myconnobj As New ADODB.Connection
    Dim myRectSet As New ADODB.Recordset
    Dim sqlStr As String
    
    myconnobj.Open "Provider=SQLOLEDB.1;Password=qorwhddnjs23;Persist Security Info=true;User ID=sa;Initial Catalog=Grade; Data Source=IT_09\SQLEXPRESS"
    
    If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Or Text2.Text <> Text4.Text Then
        MsgBox "아이디, 비밀번호, 이름, 재확인을 해주세요", vbCritical, "Account Error"
    Else
        sqlStr = "insert into systemLogin(ID, Password, Username) values('" & Trim(Text1.Text) & "', '" & Trim(Text2.Text) & "', '" & Trim(Text3.Text) & "')"
        myRectSet.Open sqlStr, myconnobj    'lstConStr, adOpenStatic, adLockReadOnly
        MsgBox "가입 정상처리 되었습니다!", vbOKOnly, "Make Account"
    End If
    
    myconnobj.Close
    Set myRectSet = Nothing
    Set myconnobj = Nothing
    Me.Hide
End Sub

Private Sub Form_Load()
    Text2.PasswordChar = "*"   '비밀번호를 화면에 *로 표기해줌
    Text4.PasswordChar = "*"
    'Text4.ForeColor = vbb    '비밀번호 재입력 textbox에는 파란색 글씨
End Sub

