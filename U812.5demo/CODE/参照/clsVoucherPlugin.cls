VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsVoucherPlugin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim envCtxCache As New UFIDA_U8_MomServiceCommon.MomCallContextCache         '数据连接文件
Dim envCtx As UFIDA_U8_MomServiceCommon.MomCallContext                       '数据连接文件
Dim login As U8Login.clsLogin                                                '数据连接文件
Dim conn As ADODB.Connection                                                 '数据连接文件/可以执行命令
Option Explicit

'请购删除后事件
Public Function PurchaseRequisition_Delete_After(domHead As MSXML2.IXMLDOMDocument2, domBody As MSXML2.IXMLDOMDocument2, ByRef errMsg As String) As Boolean
    Dim eleline   As IXMLDOMElement
    Dim editprop As String
On Error GoTo ErrHandler
    Dim cdefine23 As String
    '从上下文获取U8Login对象及取帐套库连接对象
    Dim strSql As String
    Set envCtx = envCtxCache.CurrentMomCallContext
    Set login = envCtx.U8Login
    Set conn = envCtx.BizDbConnection
    conn.CursorLocation = adUseClient
    For Each eleline In domBody.selectNodes("//z:row")
        editprop = Null2Something(eleline.getAttribute("editprop"))
        cdefine23 = Null2Something(eleline.getAttribute("cdefine23"))
        If cdefine23 <> "" Then
            strSql = "update EF_ProcurementPlans set iqglqty=(select sum(fquantity) as iqty from PU_AppVouchs where cdefine23='" & cdefine23 & "') where autoid=" & cdefine23
            conn.Execute strSql
        End If
    Next
    
    PurchaseRequisition_Delete_After = True
    Exit Function
    
ErrHandler:
    errMsg = Err.Description
End Function


'材料出库单删除后事件
Public Function MaterialOut_Delete_After(domHead As MSXML2.IXMLDOMDocument2, domBody As MSXML2.IXMLDOMDocument2, ByRef errMsg As String) As Boolean
    Dim eleline   As IXMLDOMElement
    Dim editprop As String
On Error GoTo ErrHandler
    Dim cdefine23 As String
    '从上下文获取U8Login对象及取帐套库连接对象
    Dim strSql As String
    Set envCtx = envCtxCache.CurrentMomCallContext
    Set login = envCtx.U8Login
    Set conn = envCtx.BizDbConnection
    conn.CursorLocation = adUseClient
    For Each eleline In domBody.selectNodes("//z:row")
        editprop = Null2Something(eleline.getAttribute("editprop"))
        cdefine23 = Null2Something(eleline.getAttribute("cdefine23"))
        If cdefine23 <> "" Then
            strSql = "update EF_ProjectMRPs set illqty=(select sum(fquantity) as iqty from PU_AppVouchs where cdefine23='" & cdefine23 & "') where autoid=" & cdefine23
            conn.Execute strSql
        End If
    Next
    
    MaterialOut_Delete_After = True
    Exit Function
    
ErrHandler:
    errMsg = Err.Description
End Function

'请购单保存后事件
Public Function PurchaseRequisition_Save_After(domHead As MSXML2.IXMLDOMDocument2, domBody As MSXML2.IXMLDOMDocument2, ByRef errMsg As String) As Boolean
    Dim eleline   As IXMLDOMElement
    Dim editprop As String
    Dim rs As New ADODB.Recordset
    Dim fquantity As Double
    Dim isjqty As Double
    Dim cInvCode As String
    Dim iCEBL As Double
On Error GoTo ErrHandler
    '从上下文获取U8Login对象及取帐套库连接对象
    Dim strSql As String
    Dim cdefine23 As String
    Set envCtx = envCtxCache.CurrentMomCallContext
    Set login = envCtx.U8Login
    Set conn = envCtx.BizDbConnection
    conn.CursorLocation = adUseClient
    
    If getAccinformation("EP", "bqgcjh", conn) = "1" Then
        '校验
        For Each eleline In domBody.selectNodes("//z:row")
            editprop = Null2Something(eleline.getAttribute("editprop"))
            cdefine23 = Null2Something(eleline.getAttribute("cdefine23"))
            cInvCode = Null2Something(eleline.getAttribute("cinvcode"))
            If editprop <> "" And cdefine23 <> "" Then
                strSql = "select isjqty from  EF_ProcurementPlans where autoid=" & cdefine23
                If rs.State = adStateOpen Then Call rs.Close
                rs.Open strSql, conn
                If Not rs.EOF And Not rs.BOF Then
                    isjqty = Null2Something(rs!isjqty, 0)
                End If
                If rs.State = adStateOpen Then Call rs.Close
                strSql = "select sum(fquantity) as iqty from  PU_AppVouchs where cdefine23=" & cdefine23
                rs.Open strSql, conn
                If Not rs.EOF And Not rs.BOF Then
                    fquantity = Null2Something(rs!iqty, 0)
                End If
                If fquantity > isjqty Then
                    If rs.State = adStateOpen Then Call rs.Close
                        strSql = "SELECT cInvDefine13 FROM Inventory WHERE cInvCode='" & cInvCode & "'"
                        rs.Open strSql, conn
                        If Not rs.EOF And Not rs.BOF Then
                            iCEBL = Null2Something(rs!cInvDefine13, 0)
                            If iCEBL <> 0 Then
                                If (fquantity - isjqty) / fquantity > iCEBL Then
                                    errMsg = errMsg & "存货" & cInvCode & "采购计划量为" & isjqty & "，累计请购量为" & fquantity & ",超出额度范围。"
                                End If
                            Else
                                errMsg = errMsg & "存货" & cInvCode & "采购计划量为" & isjqty & "，累计请购量为" & fquantity & ",超出额度范围。"
                            End If
                        End If
                End If
            End If
        Next
        If Len(errMsg) > 0 Then
            PurchaseRequisition_Save_After = False
            GoTo Finish
        End If
    End If
    
    
    '回写
    For Each eleline In domBody.selectNodes("//z:row")
        editprop = Null2Something(eleline.getAttribute("editprop"))
        cdefine23 = Null2Something(eleline.getAttribute("cdefine23"))
        If editprop <> "" And cdefine23 <> "" Then
            strSql = "update EF_ProcurementPlans set iqglqty=(select sum(fquantity) as iqty from PU_AppVouchs where cdefine23='" & cdefine23 & "') where autoid=" & cdefine23
            conn.Execute strSql
        End If
    Next
    PurchaseRequisition_Save_After = True
    Exit Function
ErrHandler:
    errMsg = Err.Description
Finish:
    If rs.State = adStateOpen Then Call rs.Close
    Set rs = Nothing
End Function


'材料出库单保存后事件
Public Function MaterialOut_Save_After(domHead As MSXML2.IXMLDOMDocument2, domBody As MSXML2.IXMLDOMDocument2, ByRef errMsg As String) As Boolean
    Dim eleline   As IXMLDOMElement
    Dim editprop As String
    Dim rs As New ADODB.Recordset
    Dim fquantity As Double
    Dim iXQQty As Double
    Dim cInvCode As String
    Dim iCEBL As Double
On Error GoTo ErrHandler
    '从上下文获取U8Login对象及取帐套库连接对象
    Dim strSql As String
    Dim cdefine23 As String
    Set envCtx = envCtxCache.CurrentMomCallContext
    Set login = envCtx.U8Login
    Set conn = envCtx.BizDbConnection
    conn.CursorLocation = adUseClient
    
    If getAccinformation("EP", "bllcjh", conn) = "1" Then
        '校验
        For Each eleline In domBody.selectNodes("//z:row")
            editprop = Null2Something(eleline.getAttribute("editprop"))
            cdefine23 = Null2Something(eleline.getAttribute("cdefine23"))
            cInvCode = Null2Something(eleline.getAttribute("cinvcode"))
            If editprop <> "" And cdefine23 <> "" Then
                strSql = "select iqty from  EF_ProjectMRPs where autoid=" & cdefine23
                If rs.State = adStateOpen Then Call rs.Close
                rs.Open strSql, conn
                If Not rs.EOF And Not rs.BOF Then
                    iXQQty = Null2Something(rs!iqty, 0)
                End If
                If rs.State = adStateOpen Then Call rs.Close
                strSql = "select sum(iQuantity) as iqty from  rdrecords11 where cdefine23=" & cdefine23
                rs.Open strSql, conn
                If Not rs.EOF And Not rs.BOF Then
                    fquantity = Null2Something(rs!iqty, 0)
                End If
                If fquantity > iXQQty Then
                    errMsg = errMsg & "存货" & cInvCode & "计划需求量为" & iXQQty & "，累计出库量为" & fquantity & ",超计划。"
                End If
            End If
        Next
        If Len(errMsg) > 0 Then
            MaterialOut_Save_After = False
            GoTo Finish
        End If
    End If
    
    
    '回写
    For Each eleline In domBody.selectNodes("//z:row")
        editprop = Null2Something(eleline.getAttribute("editprop"))
        cdefine23 = Null2Something(eleline.getAttribute("cdefine23"))
        If editprop <> "" And cdefine23 <> "" Then
            strSql = "update EF_ProjectMRPs set illqty=(select sum(iQuantity) as iqty from rdrecords11 where cdefine23='" & cdefine23 & "') where autoid=" & cdefine23
            conn.Execute strSql
            strSql = "update EF_ProjectMRPs set cbcloser='" & g_oLogin.cUserName & "',cbclosedate='" & g_oLogin.CurDate & "' where illqty=iqty and autoid=" & cdefine23
            conn.Execute strSql
        End If
    Next
    MaterialOut_Save_After = True
    Exit Function
ErrHandler:
    errMsg = Err.Description
Finish:
    If rs.State = adStateOpen Then Call rs.Close
    Set rs = Nothing
End Function

