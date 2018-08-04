VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsUserControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Function CallBack_FilterInit(flt As UFGeneralFilter.FilterSrv)
    Dim strSql As String
    strSql = "SELECT distinct bObjectCode ,bObjectName from v_HY_DZ_BorrowOut where " & sAuth_UnitR
    flt.SetItemFilterSQL "bObjectCode", strSql
End Function

