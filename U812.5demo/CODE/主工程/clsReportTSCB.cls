VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsReportTSCB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit





Public Sub GetSQL(e As IFilterArgs)
    Dim Temptablename As String
    Dim sql As String
    Dim conn As New ADODB.Connection
    Dim mlogin As New U8Login.clsLogin
    e.DataSource.Type = CustomDataSourceTypeEnum_SQL
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
    Dim tmptable As String
'    MsgBox "clsReportTSCB" & e.ReportID
    Select Case UCase(e.ReportID)
    
       '图书成本卡
       Case UCase("1N[__]3783581a-efe4-41c4-a666-134c3ce32b6b")
            tmptable = GettYSDiffZsTable(conn, e)
            e.DataSource.sql = " select * from  " & tmptable
            conn.Execute e.DataSource.sql
    End Select
End Sub



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
    sql = sql & " select '估算成本' class,p.cinvcode, p.cinvname,p.celementcode,p.celementname ,cfree1,cfree2 , sum( isnull(p.imoney,0)) imoney "
    sql = sql & " from dbo.EFFYGL_V_PcostbudgetList P "
    sql = sql & Tmpfilter
    sql = sql & " group by p.cinvcode, p.cinvname,p.celementcode,p.celementname ,cfree1,cfree2 "
    sql = sql & " Union" & vbCrLf
    sql = sql & " select '结算成本' class ,p.cinvcode, p.cinvname,p.celementcode,p.celementname ,cfree1,cfree2 , sum( isnull(p.imoney,0)) imoney "
    sql = sql & " from dbo.EFFYGL_V_SettleVouchList P "
    sql = sql & Tmpfilter
    sql = sql & " group by p.cinvcode, p.cinvname,p.celementcode,p.celementname ,cfree1,cfree2 "
    sql = sql & ") a"
    conn.Execute sql
    GettYSDiffZsTable = tmptable
End Function
