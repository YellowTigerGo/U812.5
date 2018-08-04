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
    Dim cInvCode As String
    Dim fQuantity As Double
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
            cInvCode = GetNodeValue(eleline, "cinvcode")
            fQuantity = str2Dbl(GetNodeValue(eleline, "fquantity"))
            strsql = "select t.define28,e.cbdefine14,e.cbdefine15,e.cbdefine24,e.cbdefine25 from mom_order m left outer join mom_orderdetail t on m.moid=t.moid" & _
                    " left outer join mom_orderdetail_extradefine e on t.modid=e.modid where m.MoCode='" & csoordercode & "' and t.SortSeq='" & iorderseq & "'"
            If rs.State <> 0 Then rs.Close
            rs.Open strsql, conn
            If Not rs.EOF And Not rs.BOF Then
                '更新材料的设计清单、屏体和项目信息
                conn.Execute "update PU_AppVouchs set cdefine28='" & rs!Define28 & "' where autoid='" & autoid & "' and isnull(cdefine28,'')=''"
                conn.Execute "update PU_AppVouchs_extradefine set cbdefine14='" & rs!cbdefine14 & "',cbdefine15='" & rs!cbdefine15 & "',cbdefine24='" & rs!cbdefine24 & "',cbdefine25='" & rs!cbdefine25 & "' where autoid='" & autoid & "'"
            End If
            strsql = "select c.Remark from mom_order m left outer join mom_orderdetail t on m.moid=t.moid" & _
                    " left outer join mom_moallocate c on t.MoDId=c.MoDId" & _
                    " where m.MoCode='" & csoordercode & "' and t.SortSeq='" & iorderseq & "' and c.InvCode ='" & cInvCode & "' and c.Qty=" & fQuantity
            If rs.State <> 0 Then rs.Close
            rs.Open strsql, conn
            If Not rs.EOF And Not rs.BOF Then
                If rs!Remark & "" <> "" Then
                    conn.Execute "update PU_AppVouchs set cbMemo='" & rs!Remark & "' where autoid='" & autoid & "' and isnull(cbMemo,'')=''"
                End If
            End If
        End If
    Next
    
    PurchaseRequisitionSave_After = True
    Exit Function
    
ErrHandler:
    errMsg = Err.Description
End Function

'产成品入库单保存前插件
Public Function ProductInSave_After(domHead As MSXML2.IXMLDOMDocument2, domBody As MSXML2.IXMLDOMDocument2, ByRef errMsg As String) As Boolean
    Dim eleline   As IXMLDOMElement
    Dim rs As New ADODB.Recordset
    
    Dim impoids As String
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
        impoids = GetNodeValue(eleline, "impoids")
        If impoids <> "" And editprop <> "D" Then
            autoid = GetNodeValue(eleline, "autoid")
            strsql = "update rdrecords10 set cItem_class='97',citemcname='项目管理',cItemCode=s.cbdefine24,cName=s.cbdefine25 " & _
                    " from mom_orderdetail_extradefine s where s.modid=rdrecords10.iMPoIds and rdrecords10.autoid='" & autoid & "' and isnull(rdrecords10.cName,'')=''"
            conn.Execute strsql
            strsql = "update r set r.cbdefine14=s.cbdefine14,r.cbdefine15=s.cbdefine15,r.cbdefine24=s.cbdefine24,r.cbdefine25=s.cbdefine25 " & _
                    " from rdrecords10 d, rdrecords10_extradefine r,mom_orderdetail_extradefine s where s.modid=d.iMPoIds and d.autoid=r.autoid and r.autoid='" & autoid & "'"
            conn.Execute strsql
        End If
    Next
    
    ProductInSave_After = True
    Exit Function
    
ErrHandler:
    errMsg = Err.Description
End Function
