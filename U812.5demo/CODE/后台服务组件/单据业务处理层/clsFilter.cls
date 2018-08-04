VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim m_bshowSumType As Boolean
Dim m_bHideSum As Boolean
Public m_Login As U8Login.clsLogin
Public Property Get bShowSumType() As Boolean
    bShowSumType = m_bshowSumType
End Property
Public Function ShowFilter(StrFilterName As String, strSubSysID As String, strFilterStr As String, Optional strFltConfig As String) As Boolean
    On Error GoTo ErrorHandle
    Dim objFilter As New UFGeneralFilter.FilterSrv
    Dim objFltCb As New clsReportCallBack
    Dim domFilter As DOMDocument

    
    If strFltConfig <> "" Then
        Set domFilter = New DOMDocument
        domFilter.loadXML strFltConfig
    End If
    ShowFilter = False
    objFltCb.StrReportName = StrFilterName
    Set objFilter.BehaviorObject = objFltCb
    objFilter.HiddenTotalView = bHideSum
    If StrFilterName <> "" Then
        ShowFilter = objFilter.OpenFilter(m_Login, "", StrFilterName, "ZZ", "") ', strSubSysID, DBConn, , True, True, m_Login.cUserId, m_Login)
        If ShowFilter Then
            strFilterStr = ConvertFilter(StrFilterName, objFilter)
            m_bshowSumType = objFilter.ShowSumType
        End If
    Else
        ShowFilter = True
    End If
    Set domFilter = Nothing
    Set objFilter = Nothing
    Set objFltCb = Nothing
    Exit Function
ErrorHandle:
    MsgBox Err.Description, vbExclamation
End Function

Public Sub SetFilter(StrFilterName As String, strSubSysID As String, m_Login As U8Login.clsLogin)
'    Dim objfltint As Object
'
'    Set objfltint = CreateObject("ufheronreport.clsFilterInterface")
    Dim objFilter As New UFGeneralFilter.FilterSrv
    Call objFilter.OpenFilterConfig(m_Login, "", StrFilterName, strSubSysID, "")
'    objfltint.DeleteFilter
    Set objFilter = Nothing
End Sub
'翻译过滤条件
Public Function ConvertFilter(StrFilterName As String, objFilter As UFGeneralFilter.FilterSrv) As String
    Dim oldWhereSql As String
    Dim newWhereSql As String
    Dim strfhdforclose As String
    Dim i As Integer
    Select Case LCase(StrFilterName)
    Case "sa26"
        For i = 1 To objFilter.FilterList.Count
            If objFilter.FilterList(i).Name = "vouchstate" Then
                If objFilter.FilterList(i).varValue = "未审核" Then
                    newWhereSql = newWhereSql & " and isnull(cverifier,N'') =N''"
                ElseIf objFilter.FilterList(i).varValue = "审核" Then
                    newWhereSql = newWhereSql & " and isnull(cverifier,N'') <>N'' and isnull(cscloser,N'')=N''"
                ElseIf objFilter.FilterList(i).varValue = "关闭" Then
                    newWhereSql = newWhereSql & " and isnull(cscloser,N'')<>N''"
                End If
                objFilter.FilterList(i).LngGrouptoken = 1
            End If
            If objFilter.FilterList(i).Name = "vouchexec" Then
                If objFilter.FilterList(i).varValue = "0" Then
                    newWhereSql = newWhereSql & " and ((isnull(iquantity,0)-isnull(fdhquantity,0)>0) or (igrouptype=2 and (isnull(inum,0)-isnull(fdhnum,0))>0))"
                ElseIf objFilter.FilterList(i).varValue = "1" Then
                    newWhereSql = newWhereSql & " and ((isnull(iquantity,0)<=isnull(fdhquantity,0)) or (igrouptype=2 and (isnull(inum,0)<=isnull(fdhnum,0))))"
                End If
                objFilter.FilterList(i).LngGrouptoken = 1
            End If
        Next i
    Case "07fwsk"
        For i = 1 To objFilter.FilterList.Count
            If objFilter.FilterList(i).Name = "cinfotypecode" Then
                If objFilter.FilterList(i).varValue = "是" Then
                    newWhereSql = newWhereSql & " and isnull(cinfotypecode,N'') <>N''"
                ElseIf objFilter.FilterList(i).varValue = "否" Then
                    newWhereSql = newWhereSql & " and isnull(cinfotypecode,N'') =N''"
                End If
                objFilter.FilterList(i).LngGrouptoken = 1
            End If
        Next i
    Case "sa19", "sa18history", "sa19history", "sa18"
        For i = 1 To objFilter.FilterList.Count
            If objFilter.FilterList(i).Name = "isverified" Then
                If objFilter.FilterList(i).varValue = "审核" Or objFilter.FilterList(i).varValue = "是" Then
                    newWhereSql = newWhereSql & " and isnull(cverifier,N'') <>N''"
                ElseIf objFilter.FilterList(i).varValue = "未审核" Or objFilter.FilterList(i).varValue = "否" Then
                    newWhereSql = newWhereSql & " and isnull(cverifier,N'') =N''"
                End If
                objFilter.FilterList(i).LngGrouptoken = 1
            ElseIf objFilter.FilterList(i).Name = "Search" Then
                If objFilter.FilterList(i).varValue = "客户价格" Then
                    newWhereSql = newWhereSql & " and isnull(ccuscode,N'') <>N''"
                ElseIf objFilter.FilterList(i).varValue = "客户类价格" Then
                    newWhereSql = newWhereSql & " and isnull(ccuscode,N'')=N''"
                End If
                objFilter.FilterList(i).LngGrouptoken = 1
            End If
            If objFilter.FilterList(i).Name = "bsales" Then
                If objFilter.FilterList(i).varValue = "是" Then
                    newWhereSql = newWhereSql & " and bsales=1"
                ElseIf objFilter.FilterList(i).varValue = "否" Then
                    newWhereSql = newWhereSql & " and bsales=0"
                End If
                objFilter.FilterList(i).LngGrouptoken = 1
            End If
             If objFilter.FilterList(i).Name = "binvalid" Then
                If objFilter.FilterList(i).varValue = "是" Then
                    newWhereSql = newWhereSql & " and binvalid=1"
                ElseIf objFilter.FilterList(i).varValue = "否" Then
                    newWhereSql = newWhereSql & " and binvalid=0"
                End If
                objFilter.FilterList(i).LngGrouptoken = 1
            End If
            If LCase(objFilter.FilterList(i).Name) = "owneraudit" Then
                If objFilter.FilterList(i).varValue = "是" Or objFilter.FilterList(i).varValue = "1" Then
                    newWhereSql = newWhereSql & GetOwnerAudit(LCase(StrFilterName), objFilter.FilterList(i).FieldName, m_Login)
                End If
                objFilter.FilterList(i).LngGrouptoken = 1
            End If
        Next i
    Case "cusprice", "invprice"
        For i = 1 To objFilter.FilterList.Count
            If objFilter.FilterList(i).Name = "Search" Then
                If objFilter.FilterList(i).varValue = "客户价格" Then
                    newWhereSql = newWhereSql & " and isnull(ccuscode,N'') <>N''"
                ElseIf objFilter.FilterList(i).varValue = "客户类价格" Then
                    newWhereSql = newWhereSql & " and isnull(ccuscode,N'') =N''"
                End If
                objFilter.FilterList(i).LngGrouptoken = 1
            End If
            If objFilter.FilterList(i).Name = "bsales" Then
                If objFilter.FilterList(i).varValue = "是" Then
                    newWhereSql = newWhereSql & " and bsales=1"
                ElseIf objFilter.FilterList(i).varValue = "否" Then
                    newWhereSql = newWhereSql & " and bsales=0"
                End If
                objFilter.FilterList(i).LngGrouptoken = 1
            End If
            If objFilter.FilterList(i).Name = "binvalid" Then
                If objFilter.FilterList(i).varValue = "是" Then
                    newWhereSql = newWhereSql & " and binvalid=1"
                ElseIf objFilter.FilterList(i).varValue = "否" Then
                    newWhereSql = newWhereSql & " and binvalid=0"
                End If
                objFilter.FilterList(i).LngGrouptoken = 1
            End If
        Next i
'    Case "sa20"
'        For i = 1 To objFilter.FilterList.count
'            If objFilter.FilterList(i).Name = "issigninover" Then
'                If objFilter.FilterList(i).varValue = "是" Or objFilter.FilterList(i).varValue = "1" Then
'
'                ElseIf objFilter.FilterList(i).varValue = "否" Or objFilter.FilterList(i).varValue = "0" Then
'                    newWhereSql = newWhereSql & " and isnull(cverifier,N'') =N''"
'                End If
'                objFilter.FilterList(i).LngGrouptoken = 1
'            End If
'        Next i
    Case "01", "02", "03", "04", "05", "06", "07", "08", "09", "13", "14", "15", "16", "17"
        For i = 1 To objFilter.FilterList.Count
            Select Case LCase(StrFilterName)
                Case "13", "14", "15", "07"
                    If objFilter.FilterList(i).Name = "isverified" Then
                        If objFilter.FilterList(i).varValue = "是" Or objFilter.FilterList(i).varValue = "1" Then
                            newWhereSql = newWhereSql & " and isnull(cChecker,N'') <>N''"
                        ElseIf objFilter.FilterList(i).varValue = "否" Or objFilter.FilterList(i).varValue = "0" Then
                            newWhereSql = newWhereSql & " and isnull(cChecker,N'') =N''"
                        End If
                        objFilter.FilterList(i).LngGrouptoken = 1
                    End If
                    If LCase(objFilter.FilterList(i).Name) = "cinvalider" Then
                        objFilter.FilterList(i).LngGrouptoken = 1
                        If InStr(1, objFilter.FilterList(i).varValue, "所有未作废") > 0 Then
                            newWhereSql = newWhereSql & " AND isnull(cinvalider,N'') =N'' "
                        ElseIf InStr(1, objFilter.FilterList(i).varValue, "所有已作废") Then
                            newWhereSql = newWhereSql & " AND isnull(cinvalider,N'') <>N'' "
                        Else
                            objFilter.FilterList(i).LngGrouptoken = 0
                        End If
                        
                    End If
                Case Else
                    If objFilter.FilterList(i).Name = "isverified" Then
                        If objFilter.FilterList(i).varValue = "是" Or objFilter.FilterList(i).varValue = "1" Then
                            newWhereSql = newWhereSql & " and isnull(cverifier,N'') <>N''"
                        ElseIf objFilter.FilterList(i).varValue = "否" Or objFilter.FilterList(i).varValue = "0" Then
                            newWhereSql = newWhereSql & " and isnull(cverifier,N'') =N''"
                        End If
                        objFilter.FilterList(i).LngGrouptoken = 1
                    End If
            End Select
            If objFilter.FilterList(i).Name = "isclosed" Then
                If objFilter.FilterList(i).varValue = "是" Or objFilter.FilterList(i).varValue = "1" Then
                    newWhereSql = newWhereSql & " and isnull(ccloser,N'') <>N''"
                ElseIf objFilter.FilterList(i).varValue = "否" Or objFilter.FilterList(i).varValue = "0" Then
                    newWhereSql = newWhereSql & " and isnull(ccloser,N'') =N''"
                End If
                objFilter.FilterList(i).LngGrouptoken = 1
            End If
            If LCase(objFilter.FilterList(i).Name) = "owneraudit" Then
                If objFilter.FilterList(i).varValue = "是" Or objFilter.FilterList(i).varValue = "1" Then
                    newWhereSql = newWhereSql & GetOwnerAudit(LCase(StrFilterName), objFilter.FilterList(i).FieldName, m_Login)
                End If
                objFilter.FilterList(i).LngGrouptoken = 1
            End If
            If objFilter.FilterList(i).Name = "islocked" Then
                If objFilter.FilterList(i).varValue = "是" Or objFilter.FilterList(i).varValue = "1" Then
                    newWhereSql = newWhereSql & " and isnull(clocker,N'') <>N''"
                ElseIf objFilter.FilterList(i).varValue = "否" Or objFilter.FilterList(i).varValue = "0" Then
                    newWhereSql = newWhereSql & " and isnull(clocker,N'') =N''"
                End If
                objFilter.FilterList(i).LngGrouptoken = 1
            End If
            If LCase(StrFilterName) = "17" Then 'execcondition
                If objFilter.FilterList(i).Name = "execcondition" Then
                    If objFilter.FilterList(i).varValue = "发货已完成" Then
                        newWhereSql = newWhereSql & " AND  (CASE WHEN ISNULL(SaleOrderSQ.iQuantity,0)=0 THEN ISNULL(SaleOrderSQ.iFHMoney,0)-ISNULL(SaleOrderSQ.iSum,0) ELSE ISNULL(SaleOrderSQ.iFHQuantity,0)-ISNULL(SaleOrderSQ.iQuantity,0) END ) >=0"
                    ElseIf objFilter.FilterList(i).varValue = "发票已完成" Then
                        newWhereSql = newWhereSql & " AND  (CASE WHEN ISNULL(SaleOrderSQ.iQuantity,0)=0 THEN ISNULL(SaleOrderSQ.iKPMoney,0)-ISNULL(SaleOrderSQ.iSum,0) ELSE ISNULL(SaleOrderSQ.iKPQuantity,0)-ISNULL(SaleOrderSQ.iQuantity,0) END ) >=0"
                    ElseIf objFilter.FilterList(i).varValue = "发货已完成且开票已完成" Then
                        newWhereSql = newWhereSql & " AND  (CASE WHEN ISNULL(SaleOrderSQ.iQuantity,0)=0 THEN ISNULL(SaleOrderSQ.iFHMoney,0)-ISNULL(SaleOrderSQ.iSum,0) ELSE ISNULL(SaleOrderSQ.iFHQuantity,0)-ISNULL(SaleOrderSQ.iQuantity,0) END ) >=0"
                        newWhereSql = newWhereSql & " AND  (CASE WHEN ISNULL(SaleOrderSQ.iQuantity,0)=0 THEN ISNULL(SaleOrderSQ.iKPMoney,0)-ISNULL(SaleOrderSQ.iSum,0) ELSE ISNULL(SaleOrderSQ.iKPQuantity,0)-ISNULL(SaleOrderSQ.iQuantity,0) END ) >=0"
                    ElseIf objFilter.FilterList(i).varValue = "发货未完成" Then
                        newWhereSql = newWhereSql & " AND  (CASE WHEN ISNULL(SaleOrderSQ.iQuantity,0)=0 THEN ISNULL(SaleOrderSQ.iFHMoney,0)-ISNULL(SaleOrderSQ.iSum,0) ELSE ISNULL(SaleOrderSQ.iFHQuantity,0)-ISNULL(SaleOrderSQ.iQuantity,0) END ) <0"
                    ElseIf objFilter.FilterList(i).varValue = "发票未完成" Then
                        newWhereSql = newWhereSql & " AND  (CASE WHEN ISNULL(SaleOrderSQ.iQuantity,0)=0 THEN ISNULL(SaleOrderSQ.iKPMoney,0)-ISNULL(SaleOrderSQ.iSum,0) ELSE ISNULL(SaleOrderSQ.iKPQuantity,0)-ISNULL(SaleOrderSQ.iQuantity,0) END ) <0"
                    ElseIf objFilter.FilterList(i).varValue = "发货未完成且开票未完成" Then
                        newWhereSql = newWhereSql & " AND  (CASE WHEN ISNULL(SaleOrderSQ.iQuantity,0)=0 THEN ISNULL(SaleOrderSQ.iFHMoney,0)-ISNULL(SaleOrderSQ.iSum,0) ELSE ISNULL(SaleOrderSQ.iFHQuantity,0)-ISNULL(SaleOrderSQ.iQuantity,0) END ) <0"
                        newWhereSql = newWhereSql & " AND  (CASE WHEN ISNULL(SaleOrderSQ.iQuantity,0)=0 THEN ISNULL(SaleOrderSQ.iKPMoney,0)-ISNULL(SaleOrderSQ.iSum,0) ELSE ISNULL(SaleOrderSQ.iKPQuantity,0)-ISNULL(SaleOrderSQ.iQuantity,0) END ) <0"
                    
                    Else
                        newWhereSql = newWhereSql & ""
                    End If
                    objFilter.FilterList(i).LngGrouptoken = 1
                End If
                If objFilter.FilterList(i).Name = "mocondition" Then
                    If objFilter.FilterList(i).varValue = "生产已下达" Then
                        newWhereSql = newWhereSql & " AND  (ISNULL(SaleOrderSQ.iMoQuantity,0)-ISNULL(SaleOrderSQ.iQuantity,0)) >=0"
                    ElseIf objFilter.FilterList(i).varValue = "委外已下达" Then
                        newWhereSql = newWhereSql & " AND  (ISNULL(SaleOrderSQ.fomquantity,0)-ISNULL(SaleOrderSQ.iQuantity,0)) >=0"
                    Else
                        newWhereSql = newWhereSql & " "
                    End If
                    objFilter.FilterList(i).LngGrouptoken = 1
                End If
                If LCase(objFilter.FilterList(i).Name) = "csrpolicy" Then
                    If objFilter.FilterList(i).varValue = "PE" Then
                        newWhereSql = newWhereSql & " AND cSRPolicy=N'PE' "
                    ElseIf objFilter.FilterList(i).varValue = "LP" Then
                        newWhereSql = newWhereSql & " AND cSRPolicy=N'LP' "
                    Else
                        newWhereSql = newWhereSql & ""
                    End If
                    objFilter.FilterList(i).LngGrouptoken = 1
                End If
            End If
            If LCase(StrFilterName) = "01" And LCase(objFilter.FilterList(i).Name) = "fhdforclose" Then
                strfhdforclose = getValueFor01Close(objFilter.FilterList(i).varValue)
                If strfhdforclose <> "" Then
                    newWhereSql = newWhereSql & strfhdforclose
                End If
                objFilter.FilterList(i).LngGrouptoken = 1
            End If
        Next i
    End Select

    oldWhereSql = objFilter.GetWhereWithToken(0)
    ConvertFilter = IIf(oldWhereSql = "", " 1=1 ", "") & oldWhereSql & newWhereSql ' & IIf(objFilter.sp_strSQLString = "", "", " and " & objFilter.sp_strSQLString)
    
End Function

Public Property Get bHideSum() As Boolean
    bHideSum = m_bHideSum
End Property

Public Property Let bHideSum(ByVal vNewValue As Boolean)
    m_bHideSum = vNewValue
End Property

'871edit for 過濾发货单执行情况
Private Function getValueFor01Close(fltvalue As String) As String
    Dim value() As String
    Dim i As Integer
    Dim j As Integer
    If fltvalue = "" Then getValueFor01Close = "": Exit Function
    value = Split(fltvalue, vbTab)
    i = UBound(value)
    For j = 0 To i
        Select Case value(j)
        Case "出库已完成"
            getValueFor01Close = getValueFor01Close & " and (case when isnull(iquantity,0)=0 then 0 else isnull(foutquantity,0)/iquantity end)>=1"
        Case "开票已完成"
            getValueFor01Close = getValueFor01Close & " and (case when isnull(iQuantity,0)<>0  then (abs(isnull(iQuantity,0))-abs(isnull(iRetQuantity,0))-abs(isnull(iSettleQuantity,0))) else (abs(isnull(inatsum,0))-abs(isnull(fretsum,0))-abs(isnull(isettlenum,0))) end)<=0"
'        Case "收款已完成"
'            getValueFor01Close = getValueFor01Close & " and "
        Case "存货已记账"
            getValueFor01Close = getValueFor01Close & " and isnull(cbaccounter,N'')<>N''"
'        Case "签回已完成"
'            getValueFor01Close = getValueFor01Close & " and (case isnull(iquantity,0) when 1=1 else isnull(iquantity,0)>isnull(isumsignquantity,0) end)"
        Case "出库未完成"
            getValueFor01Close = getValueFor01Close & " and (case when isnull(iquantity,0)=0 then 1 else isnull(foutquantity,0)/iquantity end)<1"
        Case "开票未完成"
            getValueFor01Close = getValueFor01Close & " and (case when isnull(iQuantity,0)<>0  then (abs(isnull(iQuantity,0))-abs(isnull(iRetQuantity,0))-abs(isnull(iSettleQuantity,0))) else (abs(isnull(inatsum,0))-abs(isnull(fretsum,0))-abs(isnull(isettlenum,0))) end)>0"
'        Case "收款未完成"
'            getValueFor01Close = getValueFor01Close & " and "
        Case "存货未记账"
            getValueFor01Close = getValueFor01Close & " and isnull(cbaccounter,N'')=N''"
'        Case "签回未完成"
'            getValueFor01Close = getValueFor01Close & " and "
        End Select
    Next
End Function

Private Function GetOwnerAudit(strVouchType As String, strFieldName As String, m_Login As U8Login.clsLogin) As String
Dim strMainView As String
Select Case strVouchType
    Case "16"
        strMainView = ""
End Select
GetOwnerAudit = " and ((" + strFieldName + " in (select cvoucherid from U8WorkFlow..table_task where (cTaskType=1 OR cTaskType=4) and cuserid='" & m_Login.cEmployeeId & "'" _
    & " and caccountid='" & m_Login.cacc_id & "' and cvouchertype='" & strVouchType & "' and ctk_state=0) " _
    & " and isnull(iswfcontrolled,0)=1) or (isnull(cverifier,N'')=N''  and isnull(iswfcontrolled,0)=0))"

End Function


