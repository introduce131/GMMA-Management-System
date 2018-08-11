Attribute VB_Name = "DBmod"
Option Explicit
    Public ConnectionString As String   '-- DB������ ����ִ� String�� �����̴�.
    Public NM_USER As String    '--���� ������� �̸��� �����ϴ� String����

Public Sub OpenConnection(ByRef adoCon As ADODB.Connection)  '--�����ͺ��̽��� ������ ���ִ� ���ν���
    Dim StrCon As String    '--DB������ �޾Ƽ� �����ϴ� String���� StrCon
    
    ConnectionString = ""
    ConnectionString = "Provider=SQLOLEDB.1;"
    ConnectionString = ConnectionString & "Password=qorwhddnjs23;"         '--SqlServer ���� ��й�ȣ
    ConnectionString = ConnectionString & "Persist Security Info=true;"    'True;
    ConnectionString = ConnectionString & "User ID=sa;"                    '--���� ID
    ConnectionString = ConnectionString & "Initial Catalog=Grade; "         '--����� �����ͺ��̽�
    ConnectionString = ConnectionString & "Data Source=192.168.1.13,8080"          '--�����ͺ��̽� �ּ� (�⺻�� = localhost)
    
    StrCon = ConnectionString
    
    If StrCon = "" Then
        MsgBox "���� ���ῡ �����߽��ϴ�", vbCritical
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

Public Function OpenRecordSet(ByRef adoRs As ADODB.Recordset, ByVal sql As String) As Long   '--OpenRecordSet�Լ� RecordCount�� ��ȯ�Ѵ�.
    DBmod.CloseRecordSet adoRs
    Set adoRs = New ADODB.Recordset
    
    adoRs.Open sql, ConnectionString, adOpenStatic, adLockReadOnly   '--Record Open
    
    OpenRecordSet = adoRs.RecordCount   '--�Լ�ȣ�� �� RecordCount�� ��ȯ�Ѵ�
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
