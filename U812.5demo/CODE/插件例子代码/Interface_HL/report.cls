VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "report"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True



Public Sub GetSql(e As IFilterArgs)

Dim ado As ADODB.Connection
 Dim wheres As String
Dim strsql As String
Dim temptable As String
Dim fit As UFIDA_U8_UAP_Services_ReportFilterService.FilterSrv

Dim ff As UFGeneralFilter.FilterSrv
'    e.DataSource.sql = ?select * from AA_Bank?
    e.DataSource.Type = 2
    temptable = e.DataSource.sql
    Set DBconn = New ADODB.Connection
    DBconn.Open e.Login.UfDbName
'    e.Login


Set ff = e.RawFilter
ff.GetSQLWhere

 
    wheres = e.FltSrv.GetValue1("str1")
    strsql = " select ccode as 单据编号,ddate as 单据日期 ,cmaker as 制单人  into " & temptable & " from EF_TABLE_PX where cvouchtype='PXEF0201' "
    
    
    
    
    DBconn.Execute strsql
 
    
     
End Sub


