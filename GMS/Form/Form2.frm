VERSION 5.00
Begin VB.Form Login 
   BackColor       =   &H00FFFFFF&
   Caption         =   "GMS : �α���"
   ClientHeight    =   9840
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   17745
   BeginProperty Font 
      Name            =   "���� ���"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   9840
   ScaleWidth      =   17745
   StartUpPosition =   2  'ȭ�� ���
   Begin VB.CommandButton cmd_account 
      BackColor       =   &H00FFFFFF&
      Caption         =   "ȸ������"
      BeginProperty Font 
         Name            =   "���� ���"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      Style           =   1  '�׷���
      TabIndex        =   6
      Top             =   9240
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "���� ���"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   10200
      Style           =   1  '�׷���
      TabIndex        =   4
      Top             =   4680
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Height          =   375
      IMEMode         =   3  '��� ����
      Left            =   7800
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   5400
      Width           =   2055
   End
   Begin VB.TextBox text1 
      Height          =   375
      Left            =   7800
      TabIndex        =   2
      Top             =   4680
      Width           =   2055
   End
   Begin VB.Label lbl_design 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "����濵ȸ�����б� �л����� ���� �ý����Դϴ�."
      BeginProperty Font 
         Name            =   "���� ���"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      TabIndex        =   7
      Top             =   3720
      Width           =   9975
   End
   Begin VB.Image Image1 
      Height          =   1785
      Left            =   11880
      Picture         =   "Form2.frx":27872
      Top             =   7920
      Width           =   6000
   End
   Begin VB.Label Label3 
      Alignment       =   2  '��� ����
      BackColor       =   &H00FFFFFF&
      Caption         =   "GMMAH Management System"
      BeginProperty Font 
         Name            =   "���� ���"
         Size            =   24
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4080
      TabIndex        =   5
      Top             =   2640
      Width           =   9735
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "���� ���"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6240
      TabIndex        =   1
      Top             =   5400
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��� ����
      BackColor       =   &H00FFFFFF&
      Caption         =   "ID"
      BeginProperty Font 
         Name            =   "���� ���"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6240
      TabIndex        =   0
      Top             =   4680
      Width           =   1215
   End
End
Attribute VB_Name = "Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_account_Click()
    account.Show
End Sub

Private Sub Command1_Click()    '--Login��ư Ŭ������ �̺�Ʈ
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
    On Error GoTo LoginErrHandler    '���� �߻��� LoginErrorHandler
    
    '������ ���� DB Server�� �ִ� �����ͺ��̽� Grade�� ID = sa PWD = qorwhddnjs23 �� ����.
     myconnobj.Open "Provider=SQLOLEDB.1;Password=qorwhddnjs23;Persist Security Info=true;User ID=sa;Initial Catalog=Grade; Data Source=IT_09\SQLEXPRESS"
    
'    sqlStr = "select count(*) as cnt from systemLogin"

'    myRectSet.Open sqlStr, myconnobj

'    count = myRectSet.Fields("cnt")

'    myRectSet.Close
    
    sqlStr = "select *  from systemLogin WHERE ID = '" & Trim(Text1.Text) & "'"
    
    id = Text1.Text     '--String���� id���ٰ� �Է��� ID������ ��´�.
    pwd = Text2.Text    '--String���� pwd�� �Է��� ��й�ȣ�� �ִ´�.
    
'    If text1.Text = "" Or Text2.Text = "" Then
'        MsgBox "ID�� Password�� �Է����ּ���", vbCritical, "�α��� ����"
'    End If

    myRectSet.Open sqlStr, myconnobj    'lstConStr, adOpenStatic, adLockReadOnly
    
     'If myRectSet.RecordCount <> 0 Then
     
            '==�����ͺ��̽����� ���� �ҷ��ͼ� �α����� �Ѵ�==
                strID = myRectSet.Fields(0)             'ID�� �����ͼ� StrID ������ �ִ´�.
                strPWD = myRectSet.Fields(1).Value      'PASSWORD�� �����ͼ� StrID ������ �ִ´�.
                user = myRectSet.Fields(2).Value        '������̸��� user ������ �ִ´�.
                form1.NameLabel = "Have a Good Day! " & "" & user
                If id = strID And pwd = strPWD Then
                    MsgBox user + "" + "�� �α��� ����", vbOKOnly, "�α��� ����"
                    form1.Show
                    Me.Hide
                End If
                myRectSet.MoveNext      '--���� Row�� �Ѿ
        'End If
        myconnobj.Close
LoginErrHandler:
    '--���̵�, ��й�ȣ ��ġ ���� LoginError
    If id <> strID Or pwd <> strPWD Or Text1.Text = "" Or Text2.Text = "" Then
        MsgBox "���̵�, ��й�ȣ�� �ٽ� Ȯ�����ּ���", vbCritical, "Login Error"
    End If
End Sub

Private Sub return_KeyPress(KeyAscii As Integer)
    Call Command1_Click
End Sub

Private Sub Form_Unload(cancel As Integer)
    End
End Sub

Private Sub text1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Text2.SetFocus
    End If
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Call Command1_Click
    End If
End Sub
