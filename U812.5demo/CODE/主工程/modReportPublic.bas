Attribute VB_Name = "modReportPublic"
Option Explicit
Option Compare Text
'---------------------------------------------------------------------------------------------
'报表管理控件指针
'查询分析器初始化
Public Const cApplicationIsRuning = True
Public pObjDefineDb As ADODB.Connection         '报表定义数据库指针
Public pObjDataDb As ADODB.Connection           '报表数据指针
'---------------------------------------------------------------------------------------------
Public Report_cMenuId As String

Sub OpenNewReport(ByVal StrReportName As String, Optional strTrask As String = vbNullString)
      Set pRepSysinfo = New clsSysInterface
      pRepSysinfo.systemId = m_Login.cSub_Id
      pRepSysinfo.ServerRunmode = False
      Set pRepSysinfo.objU8login = m_Login
      pRepSysinfo.InitInterFace m_Login.cSub_Id, , , DBConn, , , , , , , , , m_Login
      Set pRepLst = pRepSysinfo.GetReportEngine()
      
      
      '打开报表
      pRepSysinfo.systemId = "KI"
      pRepSysinfo.ServerRunmode = True
      Set pRepSysinfo.objU8login = m_Login
      pRepSysinfo.HelpFile = App.HelpFile
      Select Case StrReportName
      End Select


End Sub

 
