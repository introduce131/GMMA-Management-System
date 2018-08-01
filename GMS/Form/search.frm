VERSION 5.00
Object = "{B9411660-10E6-4A53-BE96-7FED334704FA}#7.0#0"; "fpSpru70.ocx"
Begin VB.Form search 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form2"
   ClientHeight    =   6570
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9495
   LinkTopic       =   "Form2"
   ScaleHeight     =   6570
   ScaleWidth      =   9495
   StartUpPosition =   3  'Windows ±âº»°ª
   Begin FPUSpreadADO.fpSpread fpSpread1 
      Height          =   3255
      Left            =   840
      TabIndex        =   2
      Top             =   360
      Width           =   7455
      _Version        =   458752
      _ExtentX        =   13150
      _ExtentY        =   5741
      _StockProps     =   64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SpreadDesigner  =   "search.frx":0000
   End
   Begin VB.CommandButton Command2 
      Caption         =   "µÚ·Î"
      BeginProperty Font 
         Name            =   "¸¼Àº °íµñ"
         Size            =   11.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   5880
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Á¶È¸"
      BeginProperty Font 
         Name            =   "¸¼Àº °íµñ"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1320
      TabIndex        =   0
      Top             =   3840
      Width           =   6495
   End
End
Attribute VB_Name = "search"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
   regist.List1.ListIndex = 0
   Text1.Text = regist.List1.Text & vbCrLf
End Sub

Private Sub Command2_Click()
Me.Hide
form1.Show
End Sub
