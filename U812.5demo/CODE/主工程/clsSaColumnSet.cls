VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSaColumnSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim m_ColSet As U8ColumnSet.clsColSet
Dim m_strColumnKey As String
Public Function Init(strColKey As String, strMode As String) As Boolean
    strColumnKey = strColKey
    Set m_ColSet = New U8ColumnSet.clsColSet
    m_ColSet.Init m_Login.UfDbName, m_Login.cUserId               'ColSet.setColMode (cKey) '设置栏目返回模式
    Call m_ColSet.setColMode(strColKey, val(strMode))
End Function
Public Function GetOrderString() As String
    GetOrderString = m_ColSet.GetOrderString
End Function

Private Sub Class_Initialize()
    Set m_ColSet = Nothing
End Sub
Public Function GetSqlString() As String
    GetSqlString = m_ColSet.GetSqlString
End Function
Public Function ColumnSet() As Boolean
    ColumnSet = True
    m_ColSet.isShowTitle = True
    m_ColSet.AllowMerge = False
    m_ColSet.ClearCache
End Function
Public Function GetColFormatStr() As String
    m_ColSet.Init m_Login.UfDbName, m_Login.cUserId               'ColSet.setColMode (cKey) '设置栏目返回模式
    m_ColSet.ClearCache
    m_ColSet.setColMode strColumnKey, 0
    GetColFormatStr = m_ColSet.getColInfo()
End Function

Public Property Get strColumnKey() As String
    strColumnKey = m_strColumnKey
End Property

Public Property Let strColumnKey(ByVal vNewValue As String)
    m_strColumnKey = vNewValue
End Property
