VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsReportCallBack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Option Compare Text
Public StrReportName As String
Dim domFilter As DOMDocument


Public Sub CallBack_ValueChange(sKey As String, flt As UFGeneralFilter.FilterSrv)
    Dim sValue As String
    Dim strSQL As String
    Dim strID As String
'    MsgBox StrReportName
    
End Sub
'当过滤初始化是回调此函数
Public Function CallBack_FilterInit(objFilter As UFGeneralFilter.FilterSrv, m_Login As U8Login.clsLogin)
If m_Login Is Nothing Then Exit Function
'    If UCase(m_Login.cSub_Id) <> "SA" Then Exit Function
    'Dim objflt As UfHeronReport.clsFilter
    Dim objflt As UFGeneralFilter.FilterItem
    Dim tmp As String
    Dim i As Integer
    Dim ele As IXMLDOMElement
    Dim lst As IXMLDOMNodeList
    StrReportName = Replace(objFilter.FilterID, "EF[__]", "")
    If Not domFilter Is Nothing Then
        Set lst = domFilter.documentElement.selectNodes("row[@type='refer']")
        For Each ele In lst
            objFilter.SetEditFilter ele.Attributes.getNamedItem("itemname").nodeValue, "", ele.Attributes.getNamedItem("expression").nodeValue
        Next
        Exit Function
    End If

    

End Function

'当点过滤时回调此函数

Public Function CallBack_FilterClick(objFilter As UFGeneralFilter.FilterSrv, Cancel As Boolean, m_Login As U8Login.clsLogin)
      '当过滤条件界面过滤按纽按下时激活该事件
      'StrReportName  报表名称
      'ObjFilter  过滤条件名称
      'cancel  true  取消操作 false 继续'
      '
      Dim i As Long
      If m_Login Is Nothing Then Exit Function

      If UCase(m_Login.cSub_Id) <> "SA" Then Exit Function
      StrReportName = Replace(objFilter.FilterID, "SA[__]", "")
      
End Function

Public Function CallBack_UserBrowse(ByVal sKey As String, objFilter As UFGeneralFilter.FilterSrv)

        If sKey = "设置账龄区间" Then
'            Set frmAge.mconn = New ADODB.Connection
'            frmAge.mconn.Open m_login.UfDbName
'            frmAge.Show vbModal
'            objFilter.setValue sKey, frmAge.StrResult
        End If
End Function

Public Sub On_FilterOk(ByVal StrReportName As String, objFilter As clsReportFilter, Cancel As Boolean, m_Login As U8Login.clsLogin)
      '当过滤条件界面过滤按纽按下时激活该事件
      'StrReportName  报表名称
      'ObjFilter  过滤条件名称
      'cancel  true  取消操作 false 继续'
      '
      Dim i As Long
      If m_Login Is Nothing Then Exit Sub
      
      If LCase(m_Login.cSub_Id) <> "SA" Then Exit Sub
      
End Sub


Public Function getAccinformation(strSysID As String, strName As String, DBConn As ADODB.Connection) As String
    Dim rst As New ADODB.Recordset

    rst.Open "Select cValue from accinformation where cSysID='" & strSysID & "' and cName='" & strName & "'", DBConn, adOpenForwardOnly, adLockReadOnly, adCmdText
    If rst.EOF Then
        getAccinformation = ""
    Else
        If IsNull(rst(0)) Then
            getAccinformation = ""
        Else
            getAccinformation = rst(0)
        End If
    End If
    rst.Close
    Set rst = Nothing
End Function


