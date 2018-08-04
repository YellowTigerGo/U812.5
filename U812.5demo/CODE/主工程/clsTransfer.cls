VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTransfer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private Const g_cSubID = "EFPB"
Private Const g_cSubCap = "出版行业插件"

'cAccID 当前的帐套名  cLastYear 上一年度
'cCurrentYear  当前年度 cParameterInfo  用户设置的参数信息
Public Function Transfer(cAccId As String, cLastYear As String, cCurrentYear As String, _
                cParameterInfo As String, oLogin As U8Login.clsLogin, _
                cErrMsg As String, oView As Object) As Boolean

'    Dim cmd As New ADODB.Command
    Dim adoConn As New ADODB.Connection
    Dim i As Integer
    cErrMsg = ""
    adoConn.ConnectionString = oLogin.UfDbName
    adoConn.Open

    On Error GoTo rlb1
'    cmd.ActiveConnection = adoConn
'    cmd.CommandType = adCmdStoredProc
'    cmd.CommandText = "ZZ_p_YearTran"
'    cmd.CommandTimeout = 0
'    cmd.Prepared = False
'
'    cmd.Parameters.Append cmd.CreateParameter("cAcc", adVarChar, adParamInput, 20, CStr(oLogin.cAcc_Id))
'    cmd.Parameters.Append cmd.CreateParameter("cYear", adVarChar, adParamInput, 20, CStr(CInt(oLogin.cIYear) - 1))
'    cmd.Parameters.Append cmd.CreateParameter("ErrStr", adVarChar, adParamOutput, 4000)
'    cmd.Execute
'    cErrMsg = CStr(cmd.Parameters("ErrStr"))
    Transfer = TransferMain(adoConn, cAccId, cLastYear, cCurrentYear, cParameterInfo, oLogin, cErrMsg, oView)
    If cErrMsg <> "" Then MsgBox "[" & g_cSubCap & "]结转失败!原因：" & cErrMsg
    Exit Function
rlb1:
    cErrMsg = "[" & g_cSubCap & "]结转失败!"
    MsgBox Err.Description
    Exit Function
End Function

'年度结转主过程
Private Function TransferMain(adoConn As ADODB.Connection, cAccId As String, cLastYear As String, cCurrentYear As String, _
                cParameterInfo As String, oLogin As U8Login.clsLogin, _
                cErrMsg As String, oView As Object) As Boolean
    Dim rs As Object: Dim strsql As String
    Dim dom As DOMDocument: Dim eleMent As IXMLDOMElement
    Dim bTran As Boolean: bTran = False
    Dim sLastData As String:  Dim sCurrData As String
    Dim cSubID As String
    Dim sTableName As String
    Dim sWhere As String
    Dim strInsertField As String: Dim strSelectField As String
    Dim clsSys As New EFVoucherMo.clsSystem '
    
    On Error GoTo errExit
    strsql = "select iyear from ufsystem..UA_Period  where cAcc_id=" & reNSql(cAccId) & " and iyear=" & val(cLastYear) & " "
    Set rs = adoConn.Execute(strsql)
    If rs.EOF Then cErrMsg = "您尚未创建下年度帐套，请先创建！": GoTo noerrExit
    
    sLastData = "UFDATA_" & cAccId & "_" & cLastYear
    sCurrData = "UFDATA_" & cAccId & "_" & cCurrentYear
    cSubID = oLogin.cSub_Id
    cSubID = g_cSubID
    
'    adoConn.BeginTrans: bTran = True
    
    If CheckTransfer(adoConn, sCurrData, cLastYear, cErrMsg) = False Then
        GoTo noerrExit
    End If
    
    strsql = "select * from Long_YearTransfer where csubid=" & reNSql(cSubID) & " order by iorder"
    Set rs = adoConn.Execute(strsql)
    Set dom = New DOMDocument: rs.Save dom, 1: rs.Close
    For Each eleMent In dom.selectNodes("//z:row")
        sTableName = eleMent.getAttribute("sTableName") & ""
        sWhere = eleMent.getAttribute("sWhere") & ""
        sWhere = VBA.Replace(sWhere, "$&sLastData&$", sLastData)    '替换系统值
        sWhere = VBA.Replace(sWhere, "$&cSubID&$", cSubID)
        
        getTableInsertFields adoConn, sTableName, strInsertField, strSelectField
        If UCase(eleMent.getAttribute("bDelOld") & "") = "TRUE" Then
            strsql = "delete from " & sCurrData & ".." & sTableName & IIf(sWhere = "", "", " where " & sWhere)
            'adoConn.Execute strSql
            clsSys.ExecuteSQL adoConn, strsql
        End If
        
        strsql = "insert into " & sCurrData & ".." & sTableName & "(" & strInsertField & ") select " & strSelectField & " from " & sLastData & ".." & sTableName & IIf(sWhere = "", "", " where " & sWhere)
        
        'adoConn.Execute strSql
        clsSys.ExecuteSQL adoConn, strsql
    Next
    
trueExit:
    If bTran Then adoConn.CommitTrans: bTran = False
    TransferMain = True: GoTo oExit
errExit:
    cErrMsg = Err.Description
    If Err.Number <> 0 Then MsgBox Err.Description & vbCrLf & strsql
noerrExit:
    If cErrMsg = "" Then cErrMsg = "Error"
    TransferMain = False
oExit:
    If Not rs Is Nothing Then
        If rs.State <> 0 Then rs.Close
        Set rs = Nothing
    End If
    If bTran Then adoConn.RollbackTrans: bTran = False
End Function

Private Function reNSql(ByVal sql As String, Optional ByVal strBefore As String = "N") As String
    reNSql = strBefore & "'" & VBA.Replace(sql, "'", "''") & "'"
End Function

Private Sub replaceSysValue(ByRef Str As String)
    Str = Replace(Str, "$&sLastData&$", "")
End Sub

Private Sub getTableInsertFields(adoConn As ADODB.Connection, ByVal sTableName As String, ByRef strInsertField As String, ByRef strSelectField As String)
    Dim rs As Object: Dim strsql As String
    Dim i As Integer
    
    strsql = "select * from " & sTableName & " where 1=2"
    Set rs = adoConn.Execute(strsql)
    strInsertField = "": strSelectField = ""
    For i = 0 To rs.Fields.Count - 1
        If rs.Fields(i).Attributes <> 16 And rs.Fields(i).Type <> adBinary Then  '自动编号
            Select Case rs.Fields(i).Type
                Case adGUID
                    strInsertField = strInsertField & "[" & rs.Fields(i).Name & "]": strSelectField = strSelectField & "newid() as [" & rs.Fields(i).Name & "]"
                Case Else
                    strInsertField = strInsertField & "[" & rs.Fields(i).Name & "]": strSelectField = strSelectField & "[" & rs.Fields(i).Name & "]"
            End Select
            strInsertField = strInsertField & ",": strSelectField = strSelectField & ","
        End If
    Next
    If strInsertField <> "" Then strInsertField = Left(strInsertField, Len(strInsertField) - 1)
    If strSelectField <> "" Then strSelectField = Left(strSelectField, Len(strSelectField) - 1)
    rs.Close: Set rs = Nothing
End Sub

Private Function CheckTransfer(adoConn As ADODB.Connection, ByVal sCurrData As String, ByVal cLastYear As String, ByRef errDec As String) As Boolean
    Dim rs As Object: Dim strsql As String
    
    strsql = "select * from " & sCurrData & "..efbwgl_dbformat"
    Set rs = adoConn.Execute(strsql)
    If Not rs.EOF Then
        CheckTransfer = False
        errDec = "已经年转，不可再次年转"
    Else
        CheckTransfer = True
    End If
    rs.Close: Set rs = Nothing
End Function



Public Sub GetSQL(e As IFilterArgs)
    Dim Temptablename As String
    Dim sql As String
    Dim conn As New ADODB.Connection
    Dim mlogin As New U8Login.clsLogin
    e.DataSource.Type = CustomDataSourceTypeEnum_TemplateTable
    Temptablename = e.DataSource.sql
    Dim comm As New ADODB.Command
    Dim rs As New ADODB.Recordset
    Dim j As Integer
    Dim strSql1 As String
    Dim sql2 As String
    Dim sql3 As String
    Dim sql4 As String
    Dim sql5 As String
    Set mlogin = e.login
    Dim strTemp As String
    conn.ConnectionString = mlogin.UfDbName
    conn.CommandTimeout = 0
    conn.Open
    comm.ActiveConnection = conn.ConnectionString
    comm.CommandType = adCmdStoredProc
    comm.CommandTimeout = 0
    comm.Prepared = False
    Dim RsTemp As New ADODB.Recordset
    Dim lngCyc As Long
    Dim strsql As String
    Dim strFieldName As String
    Dim strItemFieldName As String
    Dim strGroupFieldName As String
    Dim strTotalFieldName As String
'    MsgBox "clsReportUAP" & e.ReportID
    Select Case UCase(e.ReportID)
        
       '图书成本卡
 
            
       Case UCase("1N[__]9edbb0bf-379a-496a-ba92-f096eaf3fce4")
            Call GettYSDiffZsDate(conn, e)
            
            
    End Select
    
End Sub
'
Private Function GettYSDiffZsDate(conn As ADODB.Connection, e As IFilterArgs) As String
e.DataSource.Type = CustomDataSourceTypeEnum_SQL

Dim Uniontable As String  '两列合并
Dim DiffSumTabel As String ''
Uniontable = GettYSDiffZsTable(conn, e)
DiffSumTabel = GetYSDiffZSSUM(conn, Uniontable)
e.DataSource.sql = "select * from " & OutFormatTable(conn, Uniontable, DiffSumTabel)
conn.Execute e.DataSource.sql
End Function

Private Function OutFormatTable(conn As ADODB.Connection, table As String, tablesum As String) As String
Dim sql As String
Dim tmptable As String
tmptable = CreateTempTable("Formate")
sql = "select IDENTITY(int, 1,1) autoid ,* into " & tmptable
sql = sql & "   from (select top 100 percent * from ( select * from " & table
sql = sql & "    union select * from  " & tablesum & " ) aa order by id) a"
conn.Execute sql
OutFormatTable = tmptable

End Function
'两列合计
Private Function GetYSDiffZSSUM(conn As ADODB.Connection, ResultTable As String) As String

tmptable = CreateTempTable("TmpUnionTable")
sql = " select 3 id, N'差异' class, cinvcode,cinvname,celementcode,celementname, sum(isnull(imoney,0)) imoney "
sql = sql & "into " & tmptable
sql = sql & " from ("
sql = sql & " select  * from " & ResultTable & " where class ='预算'"
sql = sql & " Union"
sql = sql & "  select id, class,cinvcode,cinvname,celementcode,celementname,-isnull(imoney,0) imoney "
sql = sql & " from " & ResultTable & " where class ='结算'"
sql = sql & " ) a "
sql = sql & " group by cinvcode,cinvname ,celementcode,celementname"
conn.Execute sql
GetYSDiffZSSUM = tmptable
End Function





Private Function GettYSDiffZsTable(conn As ADODB.Connection, e As IFilterArgs) As String
Dim Tmpfilter As String
Dim tmptable As String
Dim sql As String
    tmptable = CreateTempTable("TmpUnionTable")
    If e.DataSource.FilterString = "" Then
        Tmpfilter = ""
    Else
        Tmpfilter = "where " & e.DataSource.FilterString & vbCrLf
    End If
    sql = "select * into  " & tmptable & " from"
    sql = sql & "("
    sql = sql & " select 1 id, '预算' class,p.cinvcode, p.cinvname,p.celementcode,p.celementname , sum( isnull(p.imoney,0)) imoney "
    sql = sql & " from dbo.EFFYGL_V_PcostbudgetList P "
    sql = sql & Tmpfilter
    sql = sql & " group by p.cinvcode, p.cinvname,p.celementcode,p.celementname "
    sql = sql & " Union" & vbCrLf
    sql = sql & " select 2 id, '结算' class ,p.cinvcode, p.cinvname,p.celementcode,p.celementname , sum( isnull(p.imoney,0)) imoney "
    sql = sql & " from dbo.EFFYGL_V_SettleVouchList P "
    sql = sql & Tmpfilter
    sql = sql & " group by p.cinvcode, p.cinvname,p.celementcode,p.celementname "
    sql = sql & ") a"
    conn.Execute sql
    GettYSDiffZsTable = tmptable
End Function




'图书成本
Public Sub GetTSCBDate(conn As ADODB.Connection, e As IFilterArgs)
Dim ygTable As String
Dim sjTable As String
Dim result As String
Dim allbook As String
Dim YGZJUnionTable As String
    e.DataSource.Type = CustomDataSourceTypeEnum_SQL
  
   ygTable = YGResult(conn, e.DataSource.FilterString)
   sjTable = RealResult(conn, e.DataSource.FilterString)
    result = JoinTmp(conn, ygTable, sjTable)

  'allbook = GetAllBook(conn, e.DataSource.FilterString)
    ' YGZJUnionTable = YGZJUnion(conn)
    'result = JoinResult(conn, allbook, YGZJUnionTable)
    e.DataSource.sql = "select * from " & result
    conn.Execute e.DataSource.sql
End Sub

'估算与结算连接统计要素金额表
Private Function YGZJUnion(conn As ADODB.Connection) As String
Dim tmptable As String
    tmptable = CreateTempTable("YGZJUnion")
Dim sql As String
sql = ""
sql = sql & "select * into " & tmptable & " from ("
sql = sql & " select  p.celementcode, p.celementname,p.cinvcode,sum(p.imoney) imoney , p.cfree1,p.cfree2, "
sql = sql & "        s.cinvcode scinvcode,s.celementcode scelementcode, s.celementname scelementname,sum(s.imoney) simoney ,s.cfree1 scfree1 ,s.cfree2 scfree2   "
sql = sql & "     From"
sql = sql & "        dbo.EFFYGL_V_PcostbudgetList p"
sql = sql & "        full join EFFYGL_V_SettleVouchList s on p.cinvcode=s.cinvcode and p.celementcode=p.celementcode"
sql = sql & "    group by p.celementcode ,s.celementcode ,p.cinvcode,s.cinvcode ,p.celementname ,s.celementname ,p.cfree1,p.cfree2 ,s.cfree1,s.cfree2"
sql = sql & " ) a "
conn.Execute sql
YGZJUnion = tmptable
End Function

'得有结算和预估 中得到所有的图书
Private Function GetAllBook(conn As ADODB.Connection, strFilter As String) As String
Dim sql As String
Dim tmptable As String
Dim Tmpfilter As String
    If strFilter = "" Then
        Tmpfilter = ""
    Else
        Tmpfilter = " where " & strFilter & vbCrLf
    End If
    
    tmptable = CreateTempTable("SJCB")
    sql = ""
    sql = sql & "       select  *  into " & tmptable & vbCrLf
    sql = sql & "  From" & vbCrLf
    sql = sql & "   (select  p.cinvcode ,p.cinvname,p.cfree1,p.cfree2 "
    sql = sql & "       from dbo.EFFYGL_V_PcostbudgetList P "
    sql = sql & Tmpfilter & vbCrLf
    sql = sql & "   Union " & vbCrLf
    sql = sql & "   select p.cinvcode,p.cinvname,p.cfree1,p.cfree2 "
    sql = sql & "       from dbo.EFFYGL_V_SettleVouchList P " & vbCrLf
    sql = sql & Tmpfilter & " ) a "
conn.Execute (sql)
    GetAllBook = tmptable
End Function


'得到图书预估成本
Private Function YGResult(conn As ADODB.Connection, strFilter As String) As String
Dim sql As String
Dim tmptable As String
    tmptable = CreateTempTable("RGCB")
    sql = sql & " select IDENTITY(int, 1,1) id, * into " & tmptable
    sql = sql & " from ("
    sql = sql & "     select  p.cinvcode,p.cinvname,p.celementcode,p.celementname,p.cfree1,p.cfree2,p.imoney"
    sql = sql & "     from dbo.EFFYGL_V_PcostbudgetList P"
        If strFilter = "" Then
             
        Else
            sql = sql & " where " & strFilter & vbCrLf
        End If
    
    sql = sql & " ) result"
conn.Execute (sql)
YGResult = tmptable
End Function


'得到图书实际成本
Private Function RealResult(conn As ADODB.Connection, strFilter As String) As String
Dim sql As String
Dim tmptable As String
    tmptable = CreateTempTable("SJCB")
    sql = ""
    sql = sql & " select IDENTITY(int, 1,1) id, * into " & tmptable
    sql = sql & " from ("
    sql = sql & "     select  s.cinvcode Scinvcode, s.cinvname scinvname ,s.celementcode Scelementcode,s.celementname scelementname "
    sql = sql & " ,s.cfree1 scfree1 ,s.cfree2 scfree2,s.imoney simoney "
    sql = sql & "     from dbo.EFFYGL_V_SettleVouchList s"
        If strFilter = "" Then
             
        Else
            sql = sql & " where " & strFilter & vbCrLf
        End If
    
    sql = sql & " ) result"
conn.Execute (sql)
RealResult = tmptable
    
End Function

'估算与结算表
Private Function CreateTable(conn As ADODB.Connection) As String
Dim tmptable As String
      tmptable = CreateTempTable("resultTable")
Dim sql As String
sql = ""
sql = sql & "CREATE TABLE  " & tmptable & "("
sql = sql & "    [cinvcode] [nvarchar](150) ,"
sql = sql & "    [cfree1] [float] ,"
sql = sql & "    [cfree2] [float] ,"
sql = sql & "    [class] [nvarchar](150),"
sql = sql & "    [celementcode] [nvarchar](150),"
sql = sql & "    imoney [float] ,"
sql = sql & "    [scelementcode] [nvarchar](150), "
sql = sql & "    simoney [float] "
sql = sql & " ) ON [PRIMARY]"
conn.Execute sql
End Function


Private Function JoinTmp(conn As ADODB.Connection, ygTable As String, sjTable As String) As String
Dim sql As String
Dim tmptable As String
    tmptable = CreateTempTable("Result")
     sql = ""
     sql = "select  IDENTITY(int, 1,1) colno,"
     sql = sql & " p.cinvcode ,P.cinvname,p.celementcode,p.celementname,p.cfree1,p.cfree2,p.imoney, s.*"
     sql = sql & " into  " & tmptable & " from "
     sql = sql & ygTable & " p  full join  " & sjTable & " s on p.id=s.id "
     conn.Execute sql
JoinTmp = tmptable
End Function
Private Function JoinResult(conn As ADODB.Connection, AllTable As String, sjTable As String) As String
    Dim sql As String
Dim tmptable As String
    tmptable = CreateTempTable("Result")
     sql = ""
     sql = "select  IDENTITY(int, 1,1) colno , * into " & tmptable & " from ("
     sql = sql & " select a.cinvcode,a.cinvname,a.cfree1,a.cfree2 "
     sql = sql & " ,t.celementcode,t.celementname,t.imoney,t.scelementcode,t.Scelementname,t.simoney "
     sql = sql & " from " & AllTable & " a left join " & sjTable & " t on  a.cinvcode =t.cinvcode --and  a.cfree1 = t.cfree1 and a.cfree2=t.cfree2) and (a.cinvcode =t.cinvcode and  a.cfree1 = t.scfree1 and a.cfree2=t.scfree2)" & vbCrLf
     sql = sql & " ) result"
     conn.Execute sql
JoinResult = tmptable
End Function






