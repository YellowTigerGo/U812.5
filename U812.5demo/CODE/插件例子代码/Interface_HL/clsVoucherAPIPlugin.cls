VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsVoucherAPIPlugin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim envCtxCache As New UFIDA_U8_MomServiceCommon.MomCallContextCache         '数据连接文件
Dim envCtx As UFIDA_U8_MomServiceCommon.MomCallContext                       '数据连接文件
Dim Login As U8Login.clsLogin                                                '数据连接文件
Dim conn As ADODB.Connection                                                 '数据连接文件/可以执行命令
Option Explicit

'请购单保存后事件处理
Public Function PurchaseRequisitionSave_After(domHead As MSXML2.IXMLDOMDocument2, domBody As MSXML2.IXMLDOMDocument2, ByRef errMsg As String) As Boolean
    Dim eleline   As IXMLDOMElement
    Dim rs As New ADODB.Recordset
    
    Dim iordertype As String
    Dim csoordercode As String
    Dim iorderseq As String
    Dim editprop As String
    Dim autoid As String
 On Error GoTo ErrHandler
    '从上下文获取U8Login对象及取帐套库连接对象
    Dim strsql As String
    Set envCtx = envCtxCache.CurrentMomCallContext
    Set Login = envCtx.U8Login
    Set conn = envCtx.BizDbConnection
    conn.CursorLocation = adUseClient
    For Each eleline In domBody.selectNodes("//z:row")
        editprop = GetNodeValue(eleline, "editprop")
        iordertype = GetNodeValue(eleline, "iordertype")
        If iordertype = "7" And editprop <> "D" Then
            csoordercode = GetNodeValue(eleline, "csoordercode")
            iorderseq = GetNodeValue(eleline, "iorderseq")
            autoid = GetNodeValue(eleline, "autoid")
            strsql = "select t.define28,e.cbdefine14,e.cbdefine15,e.cbdefine24,e.cbdefine25 from mom_order m left outer join mom_orderdetail t on m.moid=t.moid" & _
                    " left outer join mom_orderdetail_extradefine e on t.modid=e.modid where m.MoCode='" & csoordercode & "' and t.SortSeq='" & iorderseq & "'"
            If rs.State <> 0 Then rs.Close
            rs.Open strsql, conn
            If Not rs.EOF And Not rs.BOF Then
                '更新材料的设计清单、屏体和项目信息
                conn.Execute "update PU_AppVouchs set cdefine28='" & rs!Define28 & "' where autoid='" & autoid & "' and isnull(cdefine28,'')=''"
                conn.Execute "update PU_AppVouchs_extradefine set cbdefine14='" & rs!cbdefine14 & "',cbdefine15='" & rs!cbdefine15 & "',cbdefine24='" & rs!cbdefine24 & "',cbdefine25='" & rs!cbdefine25 & "' where autoid='" & autoid & "'"
            End If
        End If
    Next
    
    PurchaseRequisitionSave_After = True
    Exit Function
    
ErrHandler:
    errMsg = Err.Description
End Function

