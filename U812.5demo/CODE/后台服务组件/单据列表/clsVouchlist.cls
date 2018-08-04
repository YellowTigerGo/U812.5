VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsVouchlist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim m_bCusAuth As Boolean
Dim m_bDepAuth As Boolean
Dim m_bVenAuth As Boolean
Dim m_bPerAuth As Boolean
Dim m_bInvAuth As Boolean
Dim m_bUseAuth As Boolean
Dim m_sTableName As String
Dim m_bRenew As Boolean
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Private Function ComputerName() As String
    Dim strBuffer   As String * 255
    Dim lngIndex    As Long
    ComputerName = ""
    lngIndex = GetComputerName(strBuffer, 255)
    If lngIndex > 0 Then
        For lngIndex = 1 To Len(strBuffer)
            If Mid(strBuffer, lngIndex, 1) = Chr(0) Then Exit For
        Next lngIndex
    End If
    ComputerName = Left(strBuffer, lngIndex - 1)
End Function

'**********************************************************************************
'sysID 系统号:sa pu  cKey 单据类型 ,Login :Longin, rs:列表结果集 ,rsSum 和结果集
'blistAll 全部取出, nPageSize :页内记录数, nCurrentPage:当前页, nPageCount:页数
' oColset 栏目对象  sWhere :Where条件 sErrRet:返回错误
'****************************************************************************
Public Function GetVchListData(ByVal sysID As String, ByVal cKey As String, login As Object, _
     ByVal bListAll As Boolean, ByVal nPageSize As Long, ByVal nCurrentPage As Long, ByRef nPageCount As Long, _
     ByRef rs As ADODB.Recordset, Optional ByRef rsSum As Variant, Optional ByRef ocolSet As Object, _
    Optional ByVal sWhere As String = "", Optional ByRef sErrRet As String, Optional sTableName As String, _
    Optional ByVal bRenew As Boolean = False, Optional nRecordCount As Long, Optional strVouchType As String) As Boolean
    On Error GoTo err_log
    Dim sSqlBody As String
    Dim sSqlSum As String
    Dim sFrom As String
    Dim m_rs As New ADODB.Recordset
    Dim m_rsSum As New ADODB.Recordset
    Dim m_conn As New ADODB.Connection
    Dim ColSet As New U8colset.clsColSet
    Dim sTableKey As String
    Dim sAuthWhere As String
    Dim strTmp As String
    Dim rst As New ADODB.Recordset
    Dim I As Long
    Dim nRsCount As Long
    rst.CursorLocation = adUseClient
    m_conn.CommandTimeout = 600
    m_conn.ConnectionTimeout = 600
    m_rs.CursorLocation = adUseClient
    m_rsSum.CursorLocation = adUseClient
    m_sTableName = "TempDB.." & "TMPUF_" & Replace(login.TaskId, ".", "") & "_" & sTableName
    m_bRenew = bRenew
    If ocolSet Is Nothing Then
        ColSet.Init login.UfDbName, login.cUserId
    Else
        Set ColSet = ocolSet
    End If
    '*************控制生成Sql*******************
    Call ColSet.setColMode(cKey, 1) '设置栏目返回模式
    '*******************************************
    m_conn.Open login.UfDbName
    sAuthWhere = GetAuthString(sysID, cKey, m_conn, login, cKey) '权限控制
    sWhere = sWhere & IIf(sAuthWhere = "", "", " And ") & sAuthWhere
    If Trim(sWhere) <> "" Then
        sWhere = " where 1=1  and " & sWhere
    Else
        sWhere = " where 1=1"
    End If
    sFrom = GetFrom(cKey, m_conn, sTableKey, strVouchType)   '数据源
    strTmp = ColSet.GetSqlString
    sSqlBody = "Select " & strTmp & " From " & sFrom & " Inner Join " & m_sTableName & " On " & sTableKey & "=A " & sWhere
    If (Not IsMissing(rsSum)) Then
        If ColSet.GetSumString = "" Then
            sSqlSum = "Select * From " & m_sTableName
        Else
            sSqlSum = "Select " & ColSet.GetSumString & " From " & m_sTableName
        End If
    End If
    If Not bListAll Then
        Dim nMinID As Long, nMaxID As Long
        If m_bRenew Then Call CreateTmpTable(strTmp, sFrom, sWhere & IIf(Len(ColSet.GetOrderString) = 0, "", " Order By " & ColSet.GetOrderString), m_conn, sTableKey)
        Dim sInArea As String
        Call GetArea(cKey, sFrom, sWhere & IIf(Len(ColSet.GetOrderString) = 0, "", " Order By " & ColSet.GetOrderString), m_conn, nPageSize, nCurrentPage, nPageCount, sTableKey, sInArea, nMinID, nMaxID, nRsCount) '分页
        If (Not IsMissing(nRecordCount)) Then nRecordCount = nRsCount
        sSqlBody = ""
        rst.Open "select * from " & m_sTableName & " where 1=2", m_conn, adOpenForwardOnly, adLockReadOnly, adCmdText
        For I = 0 To rst.Fields.Count - 1
            If LCase(rst.Fields(I).Name) <> LCase("IDOrder") Then
                sSqlBody = sSqlBody & IIf(sSqlBody = "", "", ",") & rst.Fields(I).Name
            End If
        Next
        rst.Close
        If sSqlBody = "" Then sSqlBody = " * "
        sSqlBody = " Select " & sSqlBody & " from " & m_sTableName & " Where IdOrder  >= " & nMinID & " And IDOrder <= " & nMaxID & " order by IdOrder"
    End If
    Set rst = Nothing
    m_rs.Open sSqlBody, m_conn, adOpenStatic, adLockReadOnly
    If (Not IsMissing(rsSum)) Then m_rsSum.Open sSqlSum, m_conn, adOpenForwardOnly, adLockReadOnly
    Set m_rs.ActiveConnection = Nothing
    Set m_rsSum.ActiveConnection = Nothing
    Set rs = m_rs.Clone(adLockReadOnly)
    If (Not IsMissing(rsSum)) Then Set rsSum = m_rsSum.Clone(adLockReadOnly)
    GetVchListData = True
    Exit Function
err_log:
        If Not IsMissing(sErrRet) Then sErrRet = Err.Description
        If Err.Number = -2147217913 Then
            sErrRet = "不能选择非数值型项目作为合计列"
        End If
End Function


'创建带自增ID的临时表
Private Function CreateTmpTable(ByVal sSqlString As String, ByVal sFrom As String, ByVal sWhere As String, ByVal Conn As ADODB.Connection, ByVal sTableKey As String) As Boolean
    On Error GoTo err_log
    Dim str As String
    Dim strEx As String
    Dim lngcol As Integer
    
    If sSqlString <> "" Then sSqlString = "," & sSqlString
    If m_bRenew Then
        Conn.Execute "if exists (select * from tempdb..sysobjects where name = '" & Mid(m_sTableName, 9) & "'" & _
            " and type='U') Drop Table " & m_sTableName
        str = LCase("Select IDENTITY(int,1,1) as IDOrder  " & sSqlString & " Into " & m_sTableName & " From " & sFrom & sWhere)
        str = Left(str, InStr(1, str, "1=1") - 1) & "1=2"
        sSqlString = Mid(sSqlString, 2)
        strEx = " Insert into " & m_sTableName & " Select " & sSqlString & " From " & sFrom & sWhere
        lngcol = InStr(1, str, "selcol")
        If lngcol > 0 Then
            str = Left(str, lngcol - 7) & " ' ' " & Right(str, Len(str) - lngcol + 5)
        End If
        Conn.Execute str
        Conn.Execute strEx
        Conn.Execute "create index IX_IDOrder on " & m_sTableName & "(idorder)"
    End If
    CreateTmpTable = True
    Exit Function
err_log:
    CreateTmpTable = False
    Err.Raise Err.Number, , Err.Description
End Function


Private Function GetArea(ByVal cKey As String, sFrom As String, ByVal sWhere As String, ByVal Conn As ADODB.Connection, ByVal nPageSize As Long, _
    ByVal nCurrentPage As Long, ByRef nPageCount As Long, ByVal sTableKey As String, ByRef sInsArea As String, ByRef nMinID As Long, ByRef nMaxID As Long, ByRef nRecordCount As Long) As Boolean
    
    Dim I As Long
    Dim rs As New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.PageSize = nPageSize
    
    rs.Open "Select IDOrder From " & m_sTableName & " order by IDOrder", Conn, adOpenStatic, adLockReadOnly
    If Not rs.EOF Then
        If nCurrentPage > rs.PageCount Then nCurrentPage = rs.PageCount
        If nCurrentPage <= 1 Then nCurrentPage = 1
        nPageCount = rs.PageCount
        rs.AbsolutePage = nCurrentPage
        If Not (rs.BOF And rs.EOF) Then nMinID = rs("IDOrder")
        For I = 1 To rs.PageSize
            If rs.EOF Then Exit For
            nMaxID = rs("IDOrder")
            rs.MoveNext
        Next
        nRecordCount = rs.RecordCount
    Else
        nRecordCount = rs.RecordCount
        nPageCount = 0
    End If
    rs.Close
End Function

'/////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'   从vouchers中取得from字符串
'单据列表对应的列表视图和关键字
'/by 客户化开发中心 2007/03/01
Private Function GetFrom(ByVal sKey As String, ByVal Conn As ADODB.Connection, ByRef sTableKey As String, ByVal VouchType As String) As String
 Dim rds As New Recordset
 Dim sql As String
 Select Case sKey
    '//////////////////////////////////////
    '可以根据业务需要进行个性化的设置
    '//////////////////////////////////////
    Case "FA1100"   '固定资产卡片列表
         GetFrom = "vw_wjbfa_cards_list"
         sTableKey = "id"
         
         Case "MT003"
            GetFrom = " V_MT_baseset03 inner join V_MT_basesets03 on V_MT_baseset03.id=V_MT_basesets03.id "
            sTableKey = "id"
            
         Case "MT004"
            GetFrom = "V_MT_baseset04 inner join V_MT_basesets04 on V_MT_baseset04.id=V_MT_basesets04.id"
            sTableKey = "id"
         
         Case "MT001"
            GetFrom = "V_MT_baseset01 inner join V_MT_basesets01 on V_MT_baseset01.id=V_MT_basesets01.id"
            sTableKey = "id"
            
    Case Else
        sql = "select BTTblName,BTQName,BWTblName,BWQName,VchListQName from Vouchers " & _
                "where [CardNumber] ='" & GetstrCardNum(VouchType) & "'"
        With rds
            If .State <> 0 Then .Close
            .CursorLocation = adUseClient
            .Open sql, Conn, 3, 4
            GetFrom = .Fields("VchListQName")
            sTableKey = "id"
        End With
 End Select
 Set rds = Nothing
End Function

Public Function DropTmpTable(ByVal sTableName As String, ByVal mLogin As Object) As Boolean
    On Error Resume Next
    Dim cnn As New ADODB.Connection
    cnn.Open mLogin.UfDbName
    cnn.Execute "if exists (select * from tempdb..sysobjects where name='" & Mid(sTableName, 9) & "' " & _
            " and type='U') Drop Table " & sTableName
    DropTmpTable = True
    cnn.Close
    Set cnn = Nothing
End Function


'************************
'获得头格式xml
'************************
'cKey, puLogin, oColSet
Public Function GetHeadXml(ByVal sKey As String, ByVal login As Object, Optional ByRef ocolSet As Object) As String
    'ByRef ocolSet As U8colset.clsColSet
    If ocolSet Is Nothing Then
        Dim ColSet As New U8colset.clsColSet
        ColSet.Init login.UfDbName, login.cUserId
        ColSet.setColMode sKey, 1        '设置栏目返回模式
        GetHeadXml = ColSet.getColInfo()
        Exit Function
    End If
    ocolSet.setColMode sKey, 1   '设置栏目返回模式
    GetHeadXml = ocolSet.getColInfo()
End Function

'权限控制
Private Function GetAuthString(ByVal sSysID As String, ByVal strBusObId As String, Conn As ADODB.Connection, login As Object, strVouchType As String) As String
    Dim objRowAuthsrv As New U8RowAuthsvr.clsRowAuth
    Dim strTmp As String
        objRowAuthsrv.Init login.UfDbName, login.cUserId '"UFSOFT"
            If bCusAuth And UCase(sSysID) = "SA" Then
                strTmp = objRowAuthsrv.GetAuthString("CUSTOMER", , "R")
                If strTmp <> "" Then
                    If strTmp = "1=2" Then
                        GetAuthString = strTmp
                        Exit Function
                    Else
                        GetAuthString = IIf(GetAuthString = "", "", GetAuthString & " and ") & "( cAuthid " & "In (" & strTmp & ") or ccuscode is null)"
                    End If
                End If
            End If
End Function

Public Property Get bCusAuth() As Boolean
    bCusAuth = m_bCusAuth
End Property

Public Property Let bCusAuth(ByVal vNewValue As Boolean)
    m_bCusAuth = vNewValue
End Property

Public Property Get bDepAuth() As Boolean
    bDepAuth = m_bDepAuth
End Property

Public Property Let bDepAuth(ByVal vNewValue As Boolean)
    m_bDepAuth = vNewValue
End Property

Public Property Get bVenAuth() As Boolean
    bVenAuth = m_bVenAuth
End Property

Public Property Let bVenAuth(ByVal vNewValue As Boolean)
    m_bVenAuth = vNewValue
End Property

Public Property Get bPerAuth() As Boolean
    bPerAuth = m_bPerAuth
End Property

Public Property Let bPerAuth(ByVal vNewValue As Boolean)
    m_bPerAuth = vNewValue
End Property

Public Property Get bInvAuth() As Boolean
    bInvAuth = m_bInvAuth
End Property

Public Property Let bInvAuth(ByVal vNewValue As Boolean)
    m_bInvAuth = vNewValue
End Property

Public Property Get bUseAuth() As Boolean
    bUseAuth = m_bUseAuth
End Property

Public Property Let bUseAuth(ByVal vNewValue As Boolean)
    m_bUseAuth = vNewValue
End Property

'/////////////////////////////////////////////////////////////////////////////////////
'
'根据自定义类型得到单据的 CardNumber 号
'by 客户化开发中心 2006/03/01
'//////////////////////////////////////////////////////////////////////////////////////
Public Function GetstrCardNum(strVouchType As String) As String
    '//xzq
    Select Case strVouchType
        Case "91"
            GetstrCardNum = "MT05"
            
        Case "92"
            GetstrCardNum = "MT06"
        
        Case "93"
            GetstrCardNum = "MT07"
        
        Case "94"
            GetstrCardNum = "MT08"
        
        Case "95"
            GetstrCardNum = "MT09"
        
        Case "87"
            GetstrCardNum = "MT01"
        Case "88"
            GetstrCardNum = "MT02"
        Case "89"
            GetstrCardNum = "MT03"
        Case "90"
            GetstrCardNum = "MT04"
            
        Case "97"
            GetstrCardNum = "FA01"  '资产增加
            
    End Select

End Function
