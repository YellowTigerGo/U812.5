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
Public Property Get bShowSumType() As Boolean
    bShowSumType = m_bshowSumType
End Property
Public Function ShowFilter(StrFilterName As String, strSubSysID As String, strFilterStr As String, Optional strFltConfig As String) As Boolean
    On Error GoTo ErrorHandle
    Dim objFilter As New UFGeneralFilter.FilterSrv
    Dim objFltCb As New clsReportCallBack
    
    If strFltConfig <> "" Then
        Set domFilter = New DOMDocument
        domFilter.loadXML strFltConfig
    End If
    ShowFilter = False
    objFltCb.StrReportName = StrFilterName
    Set objFilter.BehaviorObject = objFltCb
    objFilter.HiddenTotalView = bHideSum
    ShowFilter = objFilter.OpenFilter(m_login, StrFilterName, "", "MT", "")  ', strSubSysID, DBConn, , True, True, m_Login.cUserId, m_Login)
    If ShowFilter Then
        strFilterStr = ConvertFilter(StrFilterName, objFilter)
        m_bshowSumType = objFilter.ShowSumType
        Set objFilterNew = objFilter
    End If
    Set domFilter = Nothing
    Set objFilter = Nothing
    Set objFltCb = Nothing
    Exit Function
ErrorHandle:
    MsgBox Err.Description, vbExclamation
End Function

Public Sub SetFilter(StrFilterName As String, strSubSysID As String)
    Dim objFilter As New UFGeneralFilter.FilterSrv
    Call objFilter.OpenFilterConfig(m_login, "", StrFilterName, strSubSysID, "")
    Set objFilter = Nothing
End Sub
'翻译过滤条件
Public Function ConvertFilter(StrFilterName As String, objFilter As UFGeneralFilter.FilterSrv) As String
    Dim oldWhereSql As String
    Dim newWhereSql As String
    Dim strfhdforclose As String
    Dim i As Integer
    Select Case LCase(StrFilterName)
    Case "01", "02", "03", "04"
        For i = 1 To objFilter.FilterList.Count
            If objFilter.FilterList(i).Name = "isclosed" Then
                If objFilter.FilterList(i).varValue = "是" Or objFilter.FilterList(i).varValue = "1" Then
                    newWhereSql = newWhereSql & " and isnull(ccloser,N'') <>N''"
                ElseIf objFilter.FilterList(i).varValue = "否" Or objFilter.FilterList(i).varValue = "0" Then
                    newWhereSql = newWhereSql & " and isnull(ccloser,N'') =N''"
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

Private Function GetOwnerAudit(strVouchtype As String, strFieldName As String) As String
Dim strMainView As String
Select Case strVouchtype
    Case "16"
        strMainView = ""
End Select
GetOwnerAudit = " and ((" + strFieldName + " in (select cvoucherid from U8WorkFlow..table_task where (cTaskType=1 OR cTaskType=4) and cuserid='" & m_login.cEmployeeId & "'" _
    & " and caccountid='" & m_login.cacc_id & "' and cvouchertype='" & strVouchtype & "' and ctk_state=0) " _
    & " and isnull(iswfcontrolled,0)=1) or (isnull(cverifier,N'')=N''  and isnull(iswfcontrolled,0)=0))"

End Function


