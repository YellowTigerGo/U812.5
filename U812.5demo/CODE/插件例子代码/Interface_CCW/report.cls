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


Public Sub Show(m_login As Object)

    Call WriteWinIniParam(m_login.GetIstallPath & "\EF\YXZB\PB\UFDIASoft.dat", "MENU", "UserName", m_login.cUserName)
    Call WriteWinIniParam(m_login.GetIstallPath & "\EF\YXZB\PB\UFDIASoft.dat", "Database", "DBParm", m_login.UfDbName)
    
    Shell m_login.GetIstallPath & "\EF\YXZB\PB\StockFind.exe", vbNormalFocus
End Sub
