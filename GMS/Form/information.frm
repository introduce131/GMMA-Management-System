VERSION 5.00
Object = "{B9411660-10E6-4A53-BE96-7FED334704FA}#7.0#0"; "fpSpru70.ocx"
Begin VB.Form information 
   BackColor       =   &H00FFFFFF&
   Caption         =   "GMS : ���� �л� ����"
   ClientHeight    =   10920
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   18750
   LinkTopic       =   "Form2"
   ScaleHeight     =   10920
   ScaleWidth      =   18750
   StartUpPosition =   3  'Windows �⺻��
   Begin FPUSpreadADO.fpSpread fpSpread1 
      Height          =   5775
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   17895
      _Version        =   458752
      _ExtentX        =   31565
      _ExtentY        =   10186
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
      SpreadDesigner  =   "information.frx":0000
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�˻�"
      Height          =   375
      Left            =   5520
      TabIndex        =   2
      Top             =   240
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   3000
      TabIndex        =   1
      ToolTipText     =   "ã���� �ϴ� ���� �л� �̸��� �����ּ���"
      Top             =   240
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�л� �̸��� �Է����ּ���"
      BeginProperty Font 
         Name            =   "���� ���"
         Size            =   11.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2775
   End
End
Attribute VB_Name = "information"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Dim myconnobj As New ADODB.Connection
    Dim myRectSet As New ADODB.Recordset
    Dim sqlStr As String
    Dim i As Integer
    Dim j As Integer
    myconnobj.Open "Provider=SQLOLEDB.1;Password=qorwhddnjs23;Persist Security Info=true;User ID=sa;Initial Catalog=Grade; Data Source=IT_09\SQLEXPRESS"
    
    sqlStr = "Select * from info_student where Name = '" & Trim(Text1.Text) & "'"
    
    myRectSet.Open sqlStr, myconnobj, adOpenStatic, adLockReadOnly
    
    If myRectSet.RecordCount = 0 Then
        MsgBox "" & Text1.Text & "�� ��ġ�ϴ� ���� �����ϴ�!", vbCritical, "Search Fail"
    Else
        For i = 0 To myRectSet.RecordCount - 1
            fpSpread1.Row = i + 1
            For j = 0 To 15
                fpSpread1.Col = j + 1
                fpSpread1.Value = myRectSet.Fields(j)
            Next j
            myRectSet.MoveNext
        Next i
    End If
End Sub

Private Sub Form_Unload(cancel As Integer)
    Me.Hide
End Sub
