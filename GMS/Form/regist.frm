VERSION 5.00
Object = "{B9411660-10E6-4A53-BE96-7FED334704FA}#7.0#0"; "fpSpru70.ocx"
Begin VB.Form regist 
   BackColor       =   &H00FFFFFF&
   Caption         =   "GMS : �л� ���� ��ȸ/����"
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
   StartUpPosition =   3  'Windows �⺻��
   Begin VB.Frame Frame_Select 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   10440
      TabIndex        =   23
      Top             =   3720
      Width           =   7095
      Begin VB.CommandButton cmd_Select 
         BackColor       =   &H00FFFFFF&
         Caption         =   "��ȸ"
         BeginProperty Font 
            Name            =   "���� ���"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6120
         Style           =   1  '�׷���
         TabIndex        =   24
         Top             =   210
         Width           =   855
      End
   End
   Begin VB.TextBox Text6 
      BeginProperty Font 
         Name            =   "���� ���"
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
         Name            =   "���� ���"
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
         Name            =   "���� ���"
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
         Name            =   "���� ���"
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
         Name            =   "���� ���"
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
         Name            =   "���� ���"
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
         Name            =   "���� ���"
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
         Name            =   "���� ���"
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
         Name            =   "���� ���"
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
      ToolTipText     =   "�г��� �������ּ���."
      Top             =   1200
      Width           =   2535
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "���� ��ȸ"
      BeginProperty Font 
         Name            =   "���� ���"
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
            Name            =   "���� ���"
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
      Caption         =   "�����ڵ� �� �а��ڵ�"
      BeginProperty Font 
         Name            =   "���� ���"
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
            Name            =   "���� ���"
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
      Alignment       =   1  '������ ����
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "���� ���"
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
      Caption         =   "�����ڵ带 �Է����ּ���"
      BeginProperty Font 
         Name            =   "���� ���"
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
      Caption         =   "�а��ڵ带 �Է����ּ���"
      BeginProperty Font 
         Name            =   "���� ���"
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
      Caption         =   "������ �Է����ּ���"
      BeginProperty Font 
         Name            =   "���� ���"
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
      Caption         =   "���ø� �������ּ���"
      BeginProperty Font 
         Name            =   "���� ���"
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
      Caption         =   "�б⸦ �������ּ���"
      BeginProperty Font 
         Name            =   "���� ���"
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
      Caption         =   "��ȣ�� �Է����ּ���"
      BeginProperty Font 
         Name            =   "���� ���"
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
      Caption         =   "�̸��� �Է����ּ���"
      BeginProperty Font 
         Name            =   "���� ���"
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
      Caption         =   "���� �������ּ���"
      BeginProperty Font 
         Name            =   "���� ���"
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
      Caption         =   "�г��� �������ּ���"
      BeginProperty Font 
         Name            =   "���� ���"
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

Private Sub Command1_Click()        '�ʱ�ȭ Ŭ����
    Combo1.Text = ""
    Combo2.Text = ""
    Combo3.Text = ""
    Combo4.Text = ""
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    Text5.Text = ""
    Text6.Text = ""
    MsgBox "�Է�â�� �ʱ�ȭ�߽��ϴ�", vbOKOnly, "�ʱ�ȭ"
End Sub

Private Sub delete_Click()
    Dim myconnobj As New ADODB.Connection
    Dim myRectSet As New ADODB.Recordset
    Dim sqlStr As String
    
    myconnobj.Open "Provider=SQLOLEDB.1;Password=qorwhddnjs23;Persist Security Info=true;User ID=sa;Initial Catalog=Grade; Data Source=IT_09\SQLEXPRESS"
    
    sqlStr = "delete from record where  �̸� = '" & Trim(Text2.Text) & "'"
    
    myRectSet.Open sqlStr, myconnobj    'lstConStr, adOpenStatic, adLockReadOnly
    
    MsgBox "���� �Ϸ�", vbOKOnly, "���� ����"
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
    
    ' ���� Ŭ���ϸ� �� ��ü�� ���õǴ� �ڵ�
    fpSpread1.OperationMode = OperationModeRow
    
    '�г��� �������ּ���
    Combo1.AddItem ("1�г�")
    Combo1.AddItem ("2�г�")
    Combo1.AddItem ("3�г�")
    
    '���� �������ּ���
    Combo2.AddItem ("1��")
    Combo2.AddItem ("2��")
    Combo2.AddItem ("3��")
    Combo2.AddItem ("4��")
    Combo2.AddItem ("5��")
    Combo2.AddItem ("6��")
    Combo2.AddItem ("7��")
    Combo2.AddItem ("8��")
    Combo2.AddItem ("9��")
    Combo2.AddItem ("10��")
    
    '�б⸦ �������ּ���
    Combo3.AddItem ("1�б�")
    Combo3.AddItem ("2�б�")
    
    '���ø� �������ּ���
    Combo4.AddItem ("1�� ������")
    Combo4.AddItem ("2�� ������")
    Combo4.AddItem ("���� ��")
    
    'fpspread2�� ��ü ��40, ��40���� ���δ�
    fpSpread2.MaxCols = 40
    fpSpread2.MaxRows = 40
    
    'fpspread2(�а��ڵ� �� �����ڵ�)�� ���� �� ǥ��
    fpSpread2.Row = 0
    fpSpread2.Col = 1:      fpSpread2.Value = "���� �ڵ�"
    fpSpread2.Col = 2:      fpSpread2.Value = "������ ��"
    fpSpread2.Col = 3:      fpSpread2.Value = "�а� �ڵ�"
    fpSpread2.Col = 4:      fpSpread2.Value = "�� ��  ��"
    
    '����� �ʵ�
    fpSpread2.Col = 2
    fpSpread2.Row = 1:      fpSpread2.Value = "����"
    fpSpread2.Row = 2:      fpSpread2.Value = "����"
    fpSpread2.Row = 3:      fpSpread2.Value = "��ȸ"
    fpSpread2.Row = 4:      fpSpread2.Value = "����"
    fpSpread2.Row = 5:      fpSpread2.Value = "�ǿ� ����"
    fpSpread2.Row = 6:      fpSpread2.Value = "�ѱ���"
    fpSpread2.Row = 7:      fpSpread2.Value = "�߱���"
    fpSpread2.Row = 8:      fpSpread2.Value = "â��"
    fpSpread2.Row = 9:      fpSpread2.Value = "ȭ���� �۹�"
    fpSpread2.Row = 10:     fpSpread2.Value = "ȸ�� ����"
    fpSpread2.Row = 11:     fpSpread2.Value = "��� ����"
    fpSpread2.Row = 12:     fpSpread2.Value = "��ǻ�� �Ϲ�"
    fpSpread2.Row = 13:     fpSpread2.Value = "�� ��"
    fpSpread2.Row = 14:     fpSpread2.Value = "���α׷��� �ǹ�"
    fpSpread2.Row = 15:     fpSpread2.Value = "ȸ�� �ǹ�"
    fpSpread2.Row = 16:     fpSpread2.Value = "�� ��"
    fpSpread2.Row = 17:     fpSpread2.Value = "�� ��"
    fpSpread2.Row = 18:     fpSpread2.Value = "���ڰ�� �ǹ�"
    fpSpread2.Row = 19:     fpSpread2.Value = "��ǻ�� �׷���"
    fpSpread2.Row = 20:     fpSpread2.Value = "���� �Ϲ�"
    fpSpread2.Row = 21:     fpSpread2.Value = "���� �ǹ�"
    
    ' �а��ڵ�
    fpSpread2.Row = 1
    fpSpread2.Col = 3:      fpSpread2.Value = "HG01"
    fpSpread2.ForeColor = vbRed
    fpSpread2.Row = 2:      fpSpread2.Value = "HG02"
    fpSpread2.ForeColor = vbRed
    fpSpread2.Row = 3:      fpSpread2.Value = "HG03"
    fpSpread2.ForeColor = vbRed
    
    ' �а� ��
    fpSpread2.Row = 1
    fpSpread2.Col = 4:      fpSpread2.Value = "�����濵��"
    fpSpread2.Row = 2:      fpSpread2.Value = "����ȸ���"
    fpSpread2.Row = 3:      fpSpread2.Value = "ȸ��������"
    
    '�� ũ�� ����
    fpSpread2.ColWidth(2) = 13      ' fpspread2�� "���� ��" ���� ũ�⸦ 13���� �ø�
    fpSpread1.ColWidth(7) = 13      ' fpspread1�� "����" ���� ũ�⸦ 13���� �ø�
    
    'fpspread1 (������ȸ) �� ���� �� ǥ��
    fpSpread1.Row = 0
    fpSpread1.Col = 1:      fpSpread1.Value = "�а��ڵ�"
    fpSpread1.Col = 2:      fpSpread1.Value = "�г�"
    fpSpread1.Col = 3:      fpSpread1.Value = "��"
    fpSpread1.Col = 4:      fpSpread1.Value = "��ȣ"
    fpSpread1.Col = 5:      fpSpread1.Value = "�̸�"
    fpSpread1.Col = 6:      fpSpread1.Value = "�б�"
    fpSpread1.Col = 7:      fpSpread1.Value = "����"
    fpSpread1.Col = 8:      fpSpread1.Value = "�����ڵ�"
    fpSpread1.Col = 9:      fpSpread1.Value = "����"
    
    ' �����ڵ� ���� ǥ��
    For i = 1 To 21 Step 1
        fpSpread2.Col = 1
        fpSpread2.Row = i
        If i < 10 Then                          ' AA01~AA09���� ǥ��
            fpSpread2.Value = "AA" & "0" & i
            fpSpread2.ForeColor = vbBlue
        Else                                    ' AA10~AA21���� ǥ��
            fpSpread2.Value = "AA" & i
            fpSpread2.ForeColor = vbBlue
        End If
    Next i
End Sub

Private Sub insert_Click()      '�߰� ��ư Ŭ����
    Dim myconnobj As New ADODB.Connection
    Dim myRectSet As New ADODB.Recordset
    Dim sqlStr As String
    
    myconnobj.Open "Provider=SQLOLEDB.1;Password=qorwhddnjs23;Persist Security Info=true;User ID=sa;Initial Catalog=Grade; Data Source=IT_09\SQLEXPRESS"
    
    If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Or Text5.Text = "" Or Text6.Text = "" Or Combo1.Text = "" Or Combo2.Text = "" Or Combo3.Text = "" Or Combo4.Text = "" Then
        MsgBox "��� �׸� ���� �־��ּ���", vbCritical, "�߰� ����"
    Else
            sqlStr = "insert into record(�а��ڵ�, �г�, ��, ��ȣ, �̸�, �б�, ����, �����ڵ�, ����) values('" & Trim(Text3.Text) & "','" & Trim(Combo1.Text) & "','" & Trim(Combo2.Text) & "','" & Trim(Text1.Text) & "','" & Trim(Text2.Text) & "','" & Trim(Combo3.Text) & "','" & Trim(Combo4.Text) & "', '" & Trim(Text6.Text) & "', '" & Trim(Text5.Text) & "')"
            myRectSet.Open sqlStr, myconnobj    'lstConStr, adOpenStatic, adLockReadOnly
             MsgBox "�߰��Ϸ�", vbOKOnly, "�����߰�"
    End If
    myconnobj.Close
End Sub
