VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_Show"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'本插件主要是实现在用户界面上的接口插件
Option Explicit

 '初始化
Public Function init(ByVal g_busines As Object, ByVal mLogin As Object)
'MsgBox "Init"


End Function


Public Sub Show(m_login As Object)

    If DBconn Is Nothing Then
        Set DBconn = New ADODB.Connection
        DBconn.ConnectionTimeout = 600
        DBconn.CommandTimeout = 1200
    End If
    If DBconn.State = 1 Then DBconn.Close
    DBconn.Open m_login.UfDbName
    Set frmOptions.m_login = m_login
    frmOptions.Show 1


'    Call WriteWinIniParam(m_login.GetIstallPath & "\EF\YXZB\PB\UFDIASoft.dat", "MENU", "UserName", m_login.cUserName)
'    Call WriteWinIniParam(m_login.GetIstallPath & "\EF\YXZB\PB\UFDIASoft.dat", "Database", "DBParm", m_login.UfDbName)
'
'    Shell m_login.GetIstallPath & "\EF\YXZB\PB\StockFind.exe", vbNormalFocus
End Sub
