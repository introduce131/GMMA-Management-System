Attribute VB_Name = "DBmod"
Option Explicit
    Public ConnectionString As String   '-- DB정보를 담아주는 String형 변수이다.
    Public NM_USER As String    '--현재 사용자의 이름을 저장하는 String변수

Public Sub OpenConnection(ByRef adoCon As ADODB.Connection)  '--데이터베이스와 연결을 해주는 프로시저
    Dim StrCon As String    '--DB정보를 받아서 저장하는 String변수 StrCon
    
    ConnectionString = ""
    ConnectionString = "Provider=SQLOLEDB.1;"
    ConnectionString = ConnectionString & "Password=qorwhddnjs23;"         '--SqlServer 접속 비밀번호
    ConnectionString = ConnectionString & "Persist Security Info=true;"    'True;
    ConnectionString = ConnectionString & "User ID=sa;"                    '--접속 ID
    ConnectionString = ConnectionString & "Initial Catalog=Grade; "         '--사용할 데이터베이스
    ConnectionString = ConnectionString & "Data Source=192.168.1.13,8080"          '--데이터베이스 주소 (기본값 = localhost)
    
    StrCon = ConnectionString
    
    If StrCon = "" Then
        MsgBox "서버 연결에 실패했습니다", vbCritical
        Exit Sub
    End If
    
    DBmod.CloseConnection adoCon
    
    Set adoCon = New ADODB.Connection
    adoCon.Open StrCon
    adoCon.CursorLocation = adUseClient
End Sub

Public Sub CloseConnection(ByRef adoCon As ADODB.Connection)
    If adoCon Is Nothing Then
    Else
        If adoCon.State = adStateOpen Then
            adoCon.Close
        End If
        Set adoCon = Nothing
    End If
End Sub

Public Function OpenRecordSet(ByRef adoRs As ADODB.Recordset, ByVal sql As String) As Long   '--OpenRecordSet함수 RecordCount를 반환한다.
    DBmod.CloseRecordSet adoRs
    Set adoRs = New ADODB.Recordset
    
    adoRs.Open sql, ConnectionString, adOpenStatic, adLockReadOnly   '--Record Open
    
    OpenRecordSet = adoRs.RecordCount   '--함수호출 후 RecordCount를 반환한다
End Function

Public Function CloseRecordSet(ByRef adoRs As ADODB.Recordset)
    If adoRs.State = adStateOpen Then
        adoRs.Close
    End If
    Set adoRs = Nothing
End Function

Public Function ExcuteQuery(ByVal adocn As ADODB.Connection, ByVal sql As String) As Long
    Dim RecordAffected As Long
    adocn.Execute sql, RecordAffected
    ExcuteQuery = RecordAffected
End Function

Public Sub SetRows(Spread As fpSpread, rs As ADODB.Recordset)
    Dim RowCnt As Integer
    Dim ColCnt As Integer
    
    Spread.MaxRows = 0
    
    For RowCnt = 0 To rs.RecordCount - 1
        Spread.MaxRows = RowCnt + 1
        Spread.Row = RowCnt + 1
        For ColCnt = 0 To rs.Fields.count - 1
            Spread.Col = ColCnt + 1
            Spread.Value = rs.Fields(ColCnt).Value
        Next ColCnt
        rs.MoveNext
    Next RowCnt
End Sub
