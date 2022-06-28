Attribute VB_Name = "M02_Syain"
Option Explicit

Sub Syain_Update()

    Dim cnW    As New ADODB.Connection
    Dim cnA    As New ADODB.Connection
    Dim rsW    As New ADODB.Recordset
    Dim rsA    As New ADODB.Recordset
    Dim strNT  As String
    Dim strSQL As String
    Dim lngC   As Long
    
    Set cnW = New ADODB.Connection
    strNT = "Initial Catalog=KYUYO;"
    cnW.ConnectionString = MYPROVIDERE & MYSERVER & strNT & USER & PSWD
    cnW.Open
    
    strNT = "Initial Catalog=process_os;"
    cnA.ConnectionString = MYPROVIDERE & MYSERVER9 & strNT & USER & PSWD9
    cnA.Open
    
    DoEvents
     
    'SQLServer�̎Ј��I�[�v��
    strSQL = ""
    strSQL = strSQL & "SELECT �x�X,"
    strSQL = strSQL & "       CODE,"
    strSQL = strSQL & "       �Ј���,"
    strSQL = strSQL & "       ��E,"
    strSQL = strSQL & "       ����,"
    strSQL = strSQL & "       ���喼"
    strSQL = strSQL & "  FROM �Ј�"
    rsA.Open strSQL, cnA, adOpenStatic, adLockOptimistic
    
    '���������̎Ј��}�X�^�[�ǂݍ���
    strSQL = ""
    strSQL = strSQL & "SELECT KBN,"
    strSQL = strSQL & "       SCODE,"
    strSQL = strSQL & "       SNAME,"
    strSQL = strSQL & "       MGR,"
    strSQL = strSQL & "       BMN3,"
    strSQL = strSQL & "       BMNNM"
    strSQL = strSQL & "  FROM KYUMTA"
    rsW.Open strSQL, cnW, adOpenStatic, adLockReadOnly
    If rsW.EOF = False Then
        rsW.MoveFirst
        Do Until rsW.EOF
            rsA.AddNew
            For lngC = 0 To 5
                rsA(lngC) = Trim(rsW(lngC))
            Next lngC
            rsA.Update
            rsW.MoveNext
        Loop
    End If

    DoEvents
    
Exit_DB:

    If Not rsA Is Nothing Then
        If rsA.State = adStateOpen Then rsA.Close
        Set rsA = Nothing
    End If
    If Not cnA Is Nothing Then
        If cnA.State = adStateOpen Then cnA.Close
        Set cnA = Nothing
    End If
    If Not rsW Is Nothing Then
        If rsW.State = adStateOpen Then rsW.Close
        Set rsW = Nothing
    End If
    If Not cnW Is Nothing Then
        If cnW.State = adStateOpen Then cnW.Close
        Set cnW = Nothing
    End If

End Sub

Sub CR_TBL_SYN()

Dim cnG    As New ADODB.Connection
Dim rsG    As New ADODB.Recordset
Dim strNT  As String
Dim strSQL As String

    strNT = "Initial Catalog=process_os;"
    cnG.ConnectionString = MYPROVIDERE & MYSERVER9 & strNT & USER & PSWD9 'SQL-�������o��
    cnG.Open
    
    '�Ј��e�[�u���폜
    strSQL = strSQL & "if exists (select * from sysobjects where id = "
    strSQL = strSQL & "object_id(N'[dbo].[�Ј�]') and "
    strSQL = strSQL & "OBJECTPROPERTY(id, N'IsUserTable') = 1) "
    strSQL = strSQL & "DROP TABLE [dbo].[�Ј�]"
    rsG.Open strSQL, cnG, adOpenStatic, adLockOptimistic
    
    'USTHA�e�[�u���쐬�i�������o���j
    strSQL = strSQL & "CREATE TABLE [dbo].[�Ј�]( "
    strSQL = strSQL & "    [CODE]   [nchar](5)  NOT NULL, "
    strSQL = strSQL & "    [�Ј���] [nchar](20) NULL DEFAULT '',"
    strSQL = strSQL & "    [��E]   [nchar](8)  NULL DEFAULT '',"
    strSQL = strSQL & "    [�x�X]   [nchar](2)  NULL DEFAULT '',"
    strSQL = strSQL & "    [����]   [nchar](2)  NULL DEFAULT '',"
    strSQL = strSQL & "    [���喼] [nchar](20) NULL DEFAULT '',"
    strSQL = strSQL & "CONSTRAINT [PK_SYN] PRIMARY KEY CLUSTERED"
    strSQL = strSQL & "( "
    strSQL = strSQL & "[CODE] ASC "
    strSQL = strSQL & ") WITH "
    strSQL = strSQL & "(PAD_INDEX = OFF, "
    strSQL = strSQL & " STATISTICS_NORECOMPUTE = OFF, "
    strSQL = strSQL & " IGNORE_DUP_KEY = OFF, "
    strSQL = strSQL & " ALLOW_ROW_LOCKS = ON, "
    strSQL = strSQL & " ALLOW_PAGE_LOCKS = ON, "
    strSQL = strSQL & " OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF "
    strSQL = strSQL & ") ON [PRIMARY]"
    strSQL = strSQL & ") ON [PRIMARY]"
    rsG.Open strSQL, cnG, adOpenStatic, adLockOptimistic
    
    If Not rsG Is Nothing Then
        If rsG.State = adStateOpen Then rsG.Close
        Set rsG = Nothing
    End If
    If Not cnG Is Nothing Then
        If cnG.State = adStateOpen Then cnG.Close
        Set cnG = Nothing
    End If
    
End Sub
