Attribute VB_Name = "modReportPublic"
Option Explicit
Option Compare Text
'---------------------------------------------------------------------------------------------
'�������ؼ�ָ��
'��ѯ��������ʼ��
Public Const cApplicationIsRuning = True
Public pObjDefineDb As ADODB.Connection         '���������ݿ�ָ��
Public pObjDataDb As ADODB.Connection           '��������ָ��
'---------------------------------------------------------------------------------------------
Public Report_cMenuId As String

Sub OpenNewReport(ByVal StrReportName As String, Optional strTrask As String = vbNullString)
      Set pRepSysinfo = New clsSysInterface
      pRepSysinfo.systemId = m_Login.cSub_Id
      pRepSysinfo.ServerRunmode = False
      Set pRepSysinfo.objU8login = m_Login
      pRepSysinfo.InitInterFace m_Login.cSub_Id, , , DBConn, , , , , , , , , m_Login
      Set pRepLst = pRepSysinfo.GetReportEngine()
      
      
      '�򿪱���
      pRepSysinfo.systemId = "KI"
      pRepSysinfo.ServerRunmode = True
      Set pRepSysinfo.objU8login = m_Login
      pRepSysinfo.HelpFile = App.HelpFile
      Select Case StrReportName
      End Select


End Sub

 
