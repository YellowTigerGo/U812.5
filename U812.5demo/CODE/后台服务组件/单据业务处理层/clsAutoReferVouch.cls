VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAutoReferVouch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim domReferConfig As New DOMDocument
Dim domReferConfigBak As New DOMDocument
Dim m_strFilter As String
Dim clsHeadColumn As New clsAutoColumnSet
Dim clsBodyColumn As New clsAutoColumnSet
Dim domFillConfig As New DOMDocument
Dim m_blnSigleColumn As Boolean
Public domSourceHead As DOMDocument
Public domSourceBody As DOMDocument
Dim m_strHeadCellCheckFields As String
Dim m_strBodyCellCheckFields As String
Dim m_PageSize As Long
Dim m_strAuth As String
Dim m_strSelectType As String
Dim m_strFillType As String
Dim m_strReferName As String
Dim m_strMainKey As String
Dim m_strDetailKey As String
Dim m_strFilterName As String
Dim m_SelectedStr As String
Dim m_strCheckFlds As String
Dim m_blnReturnFlag As Boolean
Dim domHeadFormat As New DOMDocument
Dim domBodyFormat As New DOMDocument
Public m_Login As U8Login.clsLogin
Public DBConn As ADODB.Connection
Public Event AfterGetBodyDatas(lngRow As Long, domBodyDatas As DOMDocument, strError As String)
Public StrOutPutcode As String
Public strCardNum As String
Public strTaskFilter As String
'引用 ：UFGeneralFilter.dll
Public Function reInit(backFilter As UFGeneralFilter.FilterSrv) As Boolean
    Dim lst As IXMLDOMNodeList
    Dim ele As IXMLDOMElement
    Dim strValue As String
    Dim domValue As New DOMDocument
    Dim arr As IXMLDOMAttribute
    Dim i As Long
    
    On Error Resume Next
    Set domReferConfig = domReferConfigBak.cloneNode(True)
    Set lst = domReferConfig.selectNodes("//z:row")
    If lst.length > 1 Then
        For Each ele In lst
            If Not ele.Attributes.getNamedItem("condition") Is Nothing Then
                strValue = ele.Attributes.getNamedItem("condition").nodeValue
                domValue.loadXML strValue
                If Not domValue Is Nothing Then
                    For i = 1 To backFilter.FilterList.Count
                        If Not domValue.documentElement.Attributes.getNamedItem(LCase(backFilter.FilterList.Item(i).Name)) Is Nothing Then
                            If Not backFilter.FilterList.Item(i).varValue = domValue.documentElement.Attributes.getNamedItem(LCase(backFilter.FilterList.Item(i).Name)).nodeValue Then
                                ele.setAttribute "valid", "0"
                                Exit For
                            End If
                        End If
                    Next
                End If
            End If
        Next
        Set lst = domReferConfig.selectNodes("//z:row[@valid='0']")
        For Each ele In lst
            domReferConfig.selectSingleNode("//rs:data").removeChild ele
        Next
    End If
End Function
Public Function Init(strCardNum As String, strReferKey As String) As Boolean
    'Set clsReferInfo = clsInfo
    Dim rst As New ADODB.Recordset
    Dim strFillKey As String
    
    '' DbSales 数据库对象针对公用函数有效,可以防止多次打开连接
    Set DBConn = New ADODB.Connection
    If DBConn.State = 1 Then DBConn.Close
    DBConn.ConnectionTimeout = 600
    DBConn.CommandTimeout = 1200
    DBConn.Open m_Login.UfDbName
    rst.CursorLocation = adUseClient
    rst.Open "select * from sa_refervoucherconfig where cardnum=N'" + strCardNum + "' and referkey=N'" + strReferKey + "'", DBConn, adOpenForwardOnly, adLockReadOnly, adCmdText
    rst.Save domReferConfig, adPersistXML
    Set domReferConfigBak = domReferConfig.cloneNode(True)
    rst.Close
    Dim nod As IXMLDOMNode
    Set nod = domReferConfig.selectSingleNode("//z:row")
    If Not nod.Attributes.getNamedItem("fillkey") Is Nothing Then
        strFillKey = nod.Attributes.getNamedItem("fillkey").Text
    End If
    If Not nod.Attributes.getNamedItem("captionresid") Is Nothing Then
        strReferName = nod.Attributes.getNamedItem("captionresid").Text
    End If
    If Not nod.Attributes.getNamedItem("selecttype") Is Nothing Then
        strSelectType = nod.Attributes.getNamedItem("selecttype").Text
    End If
    If Not nod.Attributes.getNamedItem("filltype") Is Nothing Then
        strFillType = nod.Attributes.getNamedItem("filltype").Text
    End If
    If nod.Attributes.getNamedItem("pagesize") Is Nothing Then
        pageSize = 20
    Else
        pageSize = CLng(nod.Attributes.getNamedItem("pagesize").Text)
    End If
    If Not nod.Attributes.getNamedItem("mainuniquekey") Is Nothing Then
        strMainKey = nod.Attributes.getNamedItem("mainuniquekey").Text
    End If
    If Not nod.Attributes.getNamedItem("detailuniquekey") Is Nothing Then
        strDetailKey = nod.Attributes.getNamedItem("detailuniquekey").Text
    End If
    If Not nod.Attributes.getNamedItem("filtername") Is Nothing Then
        StrFilterName = nod.Attributes.getNamedItem("filtername").Text
    End If
    If Not nod.Attributes.getNamedItem("uniqueflds") Is Nothing Then
        strCheckFlds = nod.Attributes.getNamedItem("uniqueflds").Text
    End If
    rst.Open "select * from sa_fieldexch where ckey=N'" + strFillKey + "' ", DBConn, adOpenForwardOnly, adLockReadOnly, adCmdText
    rst.Save domFillConfig, adPersistXML
    Dim lst As IXMLDOMNodeList
    Set lst = domFillConfig.selectNodes("//z:row[@cellcheck='True' and @targetsection='T']")
    For Each nod In lst
        strHeadCellCheckFields = strHeadCellCheckFields + IIf(strHeadCellCheckFields = "", "", ",") + nod.Attributes.getNamedItem("targetfld").Text
    Next
    Set lst = domFillConfig.selectNodes("//z:row[@cellcheck='True' and @targetsection='B']")
    For Each nod In lst
        strBodyCellCheckFields = strBodyCellCheckFields + IIf(strBodyCellCheckFields = "", "", ",") + nod.Attributes.getNamedItem("targetfld").Text
    Next
    rst.Close
    Set rst = Nothing
    Dim clsAuth As New clsUserInterfaceInit
    Set clsAuth.clsAuth = m_Login.RowAuth
    clsAuth.Init DBConn, m_Login.cUserId, "refervoucher", strCardNum
    strAuth = clsAuth.GetAuthStringByOperation(strReferKey, "T")
    Set clsAuth = Nothing
End Function
Public Sub InitReferVoucher(ctlrefervoucher As Object)
    Dim strHeadKey As String
    Dim strBodyKey As String
    Dim nod As IXMLDOMNode
    Dim rst As New ADODB.Recordset
    Dim strButtons As String
    Dim domButtons As DOMDocument
    
    rst.CursorLocation = adUseClient
    Set nod = domReferConfig.selectSingleNode("//z:row")
    strHeadKey = nod.Attributes.getNamedItem("maincolumnkey").Text
    rst.Open "select * from sa_voucherlistfieldconfig where ckey=N'" + strHeadKey + "'", DBConn, adOpenForwardOnly, adLockReadOnly, adCmdText
    rst.Save domHeadFormat, adPersistXML
    rst.Close
    If Not nod.Attributes.getNamedItem("detailcolumnkey") Is Nothing Then
        strBodyKey = nod.Attributes.getNamedItem("detailcolumnkey").Text
        rst.Open "select * from sa_voucherlistfieldconfig where ckey=N'" + strBodyKey + "'", DBConn, adOpenForwardOnly, adLockReadOnly, adCmdText
        rst.Save domBodyFormat, adPersistXML
        rst.Close
    End If
    If strBodyKey = "" Then
        blnSigleColumn = True
    Else
        blnSigleColumn = False
    End If
    If Not nod.Attributes.getNamedItem("buttonspara") Is Nothing Then
        strButtons = nod.Attributes.getNamedItem("buttonspara").Text
        strButtons = ReplaceResId(strButtons)
        Set domButtons = New DOMDocument
        domButtons.loadXML strButtons
    End If
    If Not ctlrefervoucher Is Nothing Then
        ctlrefervoucher.Init blnSigleColumn, Nothing, domButtons, 1, 1, 1, Nothing, m_Login
        ctlrefervoucher.setKey strHeadKey, strBodyKey
        ctlrefervoucher.SetFilterOne
        ctlrefervoucher.UFShowBody.Visible = False
    End If
    Set clsHeadColumn.m_Login = m_Login
    Set clsBodyColumn.m_Login = m_Login
    
    clsHeadColumn.Init strHeadKey, "0"
    
    If strBodyKey <> "" Then
        clsBodyColumn.Init strBodyKey, "0"
    End If
    Set rst = Nothing
End Sub
Private Sub Class_Terminate()
    Set domReferConfig = Nothing
    Set clsHeadColumn = Nothing
    Set clsBodyColumn = Nothing
    Set domFillConfig = Nothing
    Set domSourceHead = Nothing
    Set domSourceBody = Nothing
    Set domHeadFormat = Nothing
    Set domBodyFormat = Nothing
End Sub
Public Function ShowFilter(m_Login As U8Login.clsLogin) As Boolean
    Dim clsFilter As New clsFilter
    Dim nod As IXMLDOMNode
    Dim StrFilterName As String
    Dim strSubSysID As String
    Dim strFilterStr As String
    Dim strReferKey As String
    
    Set nod = domReferConfig.selectSingleNode("//z:row")
    If Not nod.Attributes.getNamedItem("filtername") Is Nothing Then
        StrFilterName = nod.Attributes.getNamedItem("filtername").Text
    End If
    If Not nod.Attributes.getNamedItem("subsysid") Is Nothing Then
        strSubSysID = nod.Attributes.getNamedItem("subsysid").Text
    End If
    clsFilter.bHideSum = True
    ShowFilter = clsFilter.ShowFilter(StrFilterName, strSubSysID, strFilterStr)
    Set clsFilter = Nothing
    If ShowFilter Then
        strReferKey = nod.Attributes.getNamedItem("referkey").Text
        strFilter = ConvertFilter(strReferKey, strFilterStr)
    End If
End Function

Private Function ConvertFilter(strReferKey As String, strFilter) As String
    Dim strwhere As String

    'Select Case strReferKey
    '    Case "copyblue01"
    '        If Not CBool(clsSAWeb.GetSysDicOption("SA", "RedAlowBlue")) Then
    '            strWhere = " and (isnull(iquantity,0)>=0 or isnull(cwhcode,N'')=N'')"
    '        End If
    '    Case "11"
    '        If blnReturnFlag And Not CBool(clsSAWeb.GetSysDicOption("SA", "RedAlowBlue")) Then
    '            strWhere = " and (isnull(ipbvquantity,0)<=0)"
    '        End If
    'End Select
    ConvertFilter = strFilter & strwhere
End Function

Public Sub SetFilter()
    Dim clsFilter As New clsFilter
    Dim nod As IXMLDOMNode
    Dim StrFilterName As String
    Dim strSubSysID As String
    
    Set nod = domReferConfig.selectSingleNode("//z:row")
    If Not nod.Attributes.getNamedItem("filtername") Is Nothing Then
        StrFilterName = nod.Attributes.getNamedItem("filtername").Text
    End If
    If Not nod.Attributes.getNamedItem("subsysid") Is Nothing Then
        strSubSysID = nod.Attributes.getNamedItem("subsysid").Text
    End If
    clsFilter.SetFilter StrFilterName, strSubSysID, m_Login
    Set clsFilter = Nothing
End Sub

Public Property Get strFilter() As String
    strFilter = m_strFilter
End Property

Public Property Let strFilter(ByVal vNewValue As String)
    m_strFilter = vNewValue
End Property
Public Sub GetOneVoucher(domHead As DOMDocument, domBody As DOMDocument)
    Dim strUniqueKey As String
    Dim nodSet As IXMLDOMNode
    Dim domHeadData As New DOMDocument
    Dim domBodyData As New DOMDocument
    Dim nod As IXMLDOMNode
    Dim strKeyValue As String
    
    Set nodSet = domReferConfig.selectSingleNode("//z:row")
    If Not nodSet Is Nothing Then
        strUniqueKey = nodSet.Attributes.getNamedItem("mainuniquekey").Text
        GetOnePageDatas domHeadData
        Set nod = domHeadData.selectSingleNode("//z:row")
        If Not nod Is Nothing Then
            Set domHead = domHeadData
            strKeyValue = nod.Attributes.getNamedItem(strUniqueKey).Text
            GetBodyDatas strKeyValue, domHead, domBodyData
            Set domBody = domBodyData
        End If
    End If
    Set domHeadData = Nothing
    Set domBodyData = Nothing
End Sub

Private Sub GetOnePageDatas(domDatas As DOMDocument, Optional intRows As Long, Optional intCurrentPage As Long, Optional intTotalPages As Long, Optional intTotalRecords As Long)
    Dim strSql As String
    Dim nod As IXMLDOMNode
    Dim strSelect As String
'    Dim strFrom As String
    Dim strOrder As String
    Dim lngMin As Long
    Dim lngMax As Long
    Dim strUniqueKey As String
    Dim strTempTblName As String
    Dim strwhere As String
    Dim strMainDataSource As String
    Dim strDetailDataSource As String
    Dim strDefaultFilter As String
'    Dim StrRequireclasscode As String
    
    Set nod = domReferConfig.selectSingleNode("//z:row")
    strMainDataSource = nod.Attributes.getNamedItem("maindatasource").Text
    If Not nod.Attributes.getNamedItem("detaildatasource") Is Nothing Then
        strDetailDataSource = nod.Attributes.getNamedItem("detaildatasource").Text
    End If
    If Not nod.Attributes.getNamedItem("defaultfilter") Is Nothing Then
        strDefaultFilter = nod.Attributes.getNamedItem("defaultfilter").Text
    End If
    strOrder = clsHeadColumn.GetOrderString
    strUniqueKey = nod.Attributes.getNamedItem("mainuniquekey").Text
    strTempTblName = "sarefertemp" + CStr(Rnd())
    If strDefaultFilter <> "" Then strwhere = strDefaultFilter
    
    If strFilter <> "" Then
        If strwhere = "" Then
            strwhere = strFilter
        Else
            strwhere = "(" & strwhere & ") and (" & strFilter & ")"
        End If
    End If
    If strTaskFilter <> "" Then
        strwhere = "(" & strwhere & ") and (" & strTaskFilter & ")"
    End If
    If strAuth <> "" Then
        strwhere = strwhere & IIf(strwhere = "", "", " and (" & strAuth & ")")
    End If
    strSql = "if exists(select * from tempdb..sysobjects where name='" + strTempTblName + "' and xtype='U') drop table tempdb..[" + strTempTblName + "] "
    If strDetailDataSource = "" Then
        strSql = strSql & "select identity(int) as  identityautoid ," + strUniqueKey + " as mainkey into tempdb..[" + strTempTblName + "] from " + strMainDataSource + " WITH (NOLOCK) " + IIf(strwhere = "", "", " where " + strwhere) & IIf(strOrder = "", "", " order by " & strOrder)
'        strFrom = nod.Attributes.getNamedItem("maindatasource").Text
    Else
        strSql = strSql & "select identity(int) as  identityautoid ," + strUniqueKey + " as mainkey into tempdb..[" + strTempTblName + "] from " + strMainDataSource _
                + " where " + strUniqueKey + " in (select distinct " + strMainDataSource + "." + strUniqueKey + " from " & strMainDataSource & " WITH (NOLOCK) inner join " & strDetailDataSource & " WITH (NOLOCK) on " & strMainDataSource & "." & strUniqueKey & "=" & strDetailDataSource & "." & strUniqueKey _
                + IIf(strwhere = "", "", " where " + strwhere) & ")" & IIf(strOrder = "", "", " order by " & strOrder)
    End If
    DBConn.Execute ConvertSQLString(strSql)
    Dim rst As New ADODB.Recordset
    rst.CursorLocation = adUseClient
    strSql = "select * from tempdb..[" + strTempTblName + "] order by identityautoid"
    rst.Open ConvertSQLString(strSql), DBConn, adOpenForwardOnly, adLockReadOnly
    If Not rst.EOF Then
        intTotalRecords = rst.RecordCount
        If intRows <> 0 Then
            rst.pageSize = intRows
        Else
            rst.pageSize = rst.RecordCount
        End If
        If intCurrentPage > rst.PageCount Then intCurrentPage = rst.PageCount
        If intCurrentPage <= 1 Then intCurrentPage = 1
        intTotalPages = rst.PageCount
        
        rst.AbsolutePage = intCurrentPage
        lngMin = rst.Fields("identityautoid")
        If intCurrentPage < intTotalPages Then   '非最后一页
            rst.Move IIf(intRows > 0, intRows - 1, rst.pageSize - 1)
            lngMax = rst.Fields("identityautoid")
        Else
            rst.MoveLast
            lngMax = rst.Fields("identityautoid")
        End If
    End If
    rst.Close
    strSelect = clsHeadColumn.GetSqlString
    strSql = "select " + strSelect + " from " & strMainDataSource & " WITH (NOLOCK) inner join tempdb..[" + strTempTblName + "] on tempdb..[" + strTempTblName + "].mainkey=" & strMainDataSource & "." + strUniqueKey + " where tempdb..[" + strTempTblName + "].identityautoid>=" + CStr(lngMin) + " and tempdb..[" + strTempTblName + "].identityautoid<=" + CStr(lngMax) + " order by tempdb..[" + strTempTblName + "].identityautoid"
    rst.Open ConvertSQLString(strSql), DBConn, adOpenForwardOnly, adLockReadOnly, adCmdText
    rst.Save domDatas, adPersistXML
    rst.Close
    Set rst = Nothing
End Sub
Public Sub RemoveHeadLines(domHead As DOMDocument, domBody As DOMDocument)
    Dim lst As IXMLDOMNodeList
    Dim ele As IXMLDOMElement
    Dim strKeyName As String
    Dim strKeyValue As String
    Dim lstBody As IXMLDOMNodeList
    Dim nod As IXMLDOMElement
    Set nod = domReferConfig.selectSingleNode("//z:row")
    strKeyName = nod.Attributes.getNamedItem("mainuniquekey").Text
    Set lst = domHead.selectNodes("//z:row")
    For Each ele In lst
        strKeyValue = GetNodeAtrVal(ele, strKeyName)
        Set lstBody = domBody.selectNodes("//z:row[@" + strKeyName + "='" + strKeyValue + "']")
        If lstBody.length <= 0 Then
            domHead.selectSingleNode("//rs:data").removeChild ele
        End If
    Next
    
End Sub

Private Sub GetBodyDatas(strKeyValue As String, domHeadDatas As DOMDocument, domDatas As DOMDocument)
    Dim strSql As String
    Dim strSelect As String
    Dim strOrder As String
    Dim nod As IXMLDOMElement
    Dim strDetailDataSource As String
    Dim strDefaultFilter As String
    Dim strMainDataSource As String
    Dim strUniqueKey As String
    Dim strDetailKey As String
    Dim strwhere As String
    Dim rst As New ADODB.Recordset
    
    rst.CursorLocation = adUseClient
    strSelect = clsBodyColumn.GetSqlString
    strOrder = clsBodyColumn.GetOrderString
    Set nod = domReferConfig.selectSingleNode("//z:row")
    strMainDataSource = nod.Attributes.getNamedItem("maindatasource").Text
    If Not nod.Attributes.getNamedItem("detaildatasource") Is Nothing Then
        strDetailDataSource = nod.Attributes.getNamedItem("detaildatasource").Text
    End If
    If Not nod.Attributes.getNamedItem("defaultfilter") Is Nothing Then
        strDefaultFilter = nod.Attributes.getNamedItem("defaultfilter").Text
    End If
    strUniqueKey = nod.Attributes.getNamedItem("mainuniquekey").Text
    strDetailKey = nod.Attributes.getNamedItem("detailuniquekey").Text
    strwhere = strDetailDataSource & "." & strUniqueKey & "=" & strKeyValue
    If strDefaultFilter <> "" Then
        strwhere = "(" & strwhere & ") and (" & strDefaultFilter & ")"
    End If
    If strFilter <> "" Then
        strwhere = IIf(strwhere = "", "", "(" & strwhere & ") and (") & strFilter & ")"
    End If
    If strAuth <> "" Then
        strwhere = strwhere & IIf(strwhere = "", "", " and (" & strAuth & ")")
    End If
    
    '' CLin 2009-10-24 为任务关联指定时增加过滤条件
    Select Case UCase(m_strFilterName)
        Case "EFZZ060601"
            strwhere = strwhere & IIf(strwhere = "", "", " and cinvcode = '" & StrOutPutcode & "'")
        Case "EFZZ060602"
            strwhere = strwhere & IIf(strwhere = "", "", " and cinvcode in(select cinvcode from PU_AppVouchs where cinvcode = '" & StrOutPutcode & "')")
        Case "EFZZ060603"
            strwhere = strwhere & IIf(strwhere = "", "", " and Dinvcode in (select invcode from mom_orderdetail where invcode = '" & StrOutPutcode & "')")
    End Select
    
    strSql = "select " & strSelect & " from " & strDetailDataSource & " where " & strDetailDataSource & "." & strUniqueKey & "=" & strKeyValue & " and " & strDetailKey & " in (select " & strDetailKey & " from " & strMainDataSource & " inner join " & strDetailDataSource & " on " & strMainDataSource & "." & strUniqueKey & "=" & strDetailDataSource & "." & strUniqueKey _
            + IIf(strwhere = "", "", " where " + strwhere) & ")" & IIf(strOrder = "", "", " order by " & strOrder)
    rst.Open ConvertSQLString(strSql), DBConn, adOpenForwardOnly, adLockReadOnly, adCmdText
    rst.Save domDatas, adPersistXML
    rst.Close
    Set rst = Nothing
    If Not nod.Attributes.getNamedItem("setdetailvalues") Is Nothing Then
        Dim varOne As Variant
        Dim varFlds As Variant
        Dim i As Long
        
        Dim nodHead As IXMLDOMElement
        Set nodHead = domHeadDatas.selectSingleNode("//z:row")
        varOne = Split(nod.Attributes.getNamedItem("setdetailvalues").nodeValue, ";")
        For i = 0 To UBound(varOne)
            varFlds = Split(varOne(i), ",")
            For Each nod In domDatas.selectNodes("//z:row")
                nod.setAttribute varFlds(1), GetNodeAtrVal(nodHead, CStr(varFlds(0)))
            Next
        Next
    End If
End Sub
Public Sub SetHeadData(ctlrefervoucher As Object, Optional intRows As Long, Optional intCurrentPage As Long, Optional intTotalPages As Long)
    Dim lngRecordCount As Long
    Dim domHeadData As New DOMDocument
    GetOnePageDatas domHeadData, intRows, intCurrentPage, intTotalPages, lngRecordCount
    If Not domHeadData Is Nothing Then
        FormatVouchList ctlrefervoucher.HeadList, domHeadFormat
        ctlrefervoucher.SetHeadDom domHeadData
        ctlrefervoucher.HeadList.RecordCount = lngRecordCount
    End If
    Set domHeadData = Nothing
    SelectedStr = ""
End Sub

Public Sub FormatVouchList(Voulist As Object, domFieldConfig As DOMDocument)
    Dim lst As IXMLDOMNodeList
    Dim nod As IXMLDOMNode
    Set lst = domFieldConfig.selectNodes("//z:row")
    For Each nod In lst
        Voulist.SetFormatString nod.Attributes.getNamedItem("fieldname").Text, nod.Attributes.getNamedItem("formatstr").Text
    Next
End Sub

Private Sub RemoveBodyDatas(ctlrefervoucher As Object, strKeyName As String, strKeyValue As String)
    Dim i As Long
    Dim lngPos As Long
    lngPos = ctlrefervoucher.GetBodyColIndex(strKeyName)
    Dim tmpSumStyle As Variant
    tmpSumStyle = ctlrefervoucher.BodyList.SumStyle
    ctlrefervoucher.BodyList.SumStyle = vlSumNone
    For i = ctlrefervoucher.BodyList.Rows - 1 To 1 Step -1
         If ctlrefervoucher.BodyList.TextMatrix(i, lngPos) = strKeyValue Then
            ctlrefervoucher.BodyList.RemoveItem i
         End If
    Next i
    ctlrefervoucher.BodyList.SumStyle = tmpSumStyle
    ctlrefervoucher.RefreshSumLine ctlrefervoucher.BodyList.GetGridBody, True
'    ctlrefervoucher.BodyList.RecordCount = ctlrefervoucher.BodyList.GetGridBody().rows - 2
End Sub
Public Sub SelectBodyDatas(ctlrefervoucher As Object, strKeyName As String, strKeyValue As String)
    Dim i As Long
    Dim lngPos As Long
    Dim lngPosSel As Long
    
    lngPos = ctlrefervoucher.GetBodyColIndex(strKeyName)
    lngPosSel = ctlrefervoucher.GetBodyColIndex("selcol")
    For i = ctlrefervoucher.BodyList.Rows - 1 To 1 Step -1
         If ctlrefervoucher.BodyList.TextMatrix(i, lngPos) = strKeyValue Then
            ctlrefervoucher.BodyList.TextMatrix(i, lngPosSel) = "Y"
         End If
    Next i
End Sub
'ctlrefervoucher:U8VoucherList.ocx
Public Sub SetBodyData(ctlrefervoucher As Object, lngCurRow As Long, Optional SelectIt As Boolean = False)
'    Dim lngCurRow As Long
    Dim nod As IXMLDOMNode
    Dim strKeyName As String
    Dim strKeyValue As String
'    Dim tmpDom As DOMDocument
    Dim domLine As New DOMDocument
    Dim ele As IXMLDOMElement
    Dim domHead As DOMDocument
'    Dim I As Long
    Dim oND As IXMLDOMNode
'    lngCurRow = ctlReferVoucher.HeadList.row
    Set nod = domReferConfig.selectSingleNode("//z:row")
    strKeyName = nod.Attributes.getNamedItem("mainuniquekey").Text
    strKeyValue = ctlrefervoucher.HeadList.TextMatrix(lngCurRow, ctlrefervoucher.GetHeadColIndex(strKeyName))
    If strKeyValue = "" Then Exit Sub
    If ctlrefervoucher.HeadList.TextMatrix(lngCurRow, 0) = "" Then
        RemoveBodyDatas ctlrefervoucher, strKeyName, strKeyValue
        RemoveSelected ctlrefervoucher, lngCurRow
    Else
        Dim domBodyData As New DOMDocument
'        Dim Dombody As DOMDocument
        Set domHead = ctlrefervoucher.GetHeadLine(lngCurRow)

        GetBodyDatas strKeyValue, domHead, domBodyData
        Dim strError As String
        If Not domBodyData Is Nothing Then
            FormatVouchList ctlrefervoucher.BodyList, domBodyFormat
            If strSelectType = "uniquemain" Then
                ctlrefervoucher.RemoveBodyAll
            End If
            If ctlrefervoucher.BodyList.Rows <= 1 Then
                Set domLine = domBodyData.cloneNode(True)
                Set oND = domLine.selectSingleNode("//rs:data")
                For Each ele In domLine.selectNodes("//z:row")
                    oND.removeChild ele
                Next
                ctlrefervoucher.SetBodyDom domLine
                'ctlrefervoucher.RemoveBodyAll
            End If
            RaiseEvent AfterGetBodyDatas(lngCurRow, domBodyData, strError)
'            ctlrefervoucher.BodyList.FillMode = FillAppend
'            ctlrefervoucher.SetBodyDom domBodyData
'            Dim I As Long
            Set domLine = domBodyData.cloneNode(True)
'            Dim ele As IXMLDOMElement
            Dim eleline As IXMLDOMElement
            Dim TmpLinedom As DOMDocument
            
            Set oND = domLine.selectSingleNode("//rs:data")
            For Each eleline In domLine.selectNodes("//z:row")
                oND.removeChild eleline
            Next
            Set TmpLinedom = domLine.cloneNode(True)
'            ctlrefervoucher.bAppending = True
            Dim tmpSumStyle As Variant
            RemoveSumLine ctlrefervoucher.BodyList.GetGridBody()
            Dim i As Long
            i = ctlrefervoucher.BodyList.GetGridBody().Rows
            tmpSumStyle = ctlrefervoucher.BodyList.SumStyle
            ctlrefervoucher.BodyList.SumStyle = vlSumNone
            
            For Each ele In domBodyData.selectNodes("//z:row")
                Set domLine = TmpLinedom.cloneNode(True)
                Set oND = domLine.selectSingleNode("//rs:data")
                oND.appendChild ele
                'domLine.selectSingleNode("//rs:data").appendChild ele
                ctlrefervoucher.AddBodyLine i, domLine
                If SelectIt Then ctlrefervoucher.BodyList.TextMatrix(i, 0) = "Y"
                i = i + 1
'                oND.removeChild ele
            Next
            Set TmpLinedom = Nothing
'            RemoveSumLine ctlrefervoucher.BodyList
'            ctlrefervoucher.bAppending = False
            'ctlrefervoucher.RemoveSumLine ctlrefervoucher.BodyList.GetGridBody()
            ctlrefervoucher.BodyList.SumStyle = tmpSumStyle
'            ctlrefervoucher.AddSumLine ctlrefervoucher.BodyList.GetGridBody(), True
            ctlrefervoucher.RefreshSumLine ctlrefervoucher.BodyList.GetGridBody(), True
'            ctlrefervoucher.BodyList.RecordCount = ctlrefervoucher.BodyList.GetGridBody().rows - 2
'            ctlrefervoucher.BodyList.CalcSum
'            ctlrefervoucher.BodyList.rows = ctlrefervoucher.BodyList.GetGridBody().rows - 1
        End If
        Set domLine = Nothing
        Set domBodyData = Nothing
    End If
End Sub

Private Sub RemoveSumLine(GridBody As Object)
On Error Resume Next
    With GridBody
        If .Rows - 1 >= 0 Then
            If .TextMatrix(.Rows - 1, 0) = GetString("U8.ST.Default.00418") Then '"合计"
                .RemoveItem .Rows - 1
            End If
        End If
    End With
End Sub
'Public Property Get strHeadKey() As String
'    strHeadKey = m_strHeadKey
'End Property
'
'Public Property Let strHeadKey(ByVal vNewValue As String)
'    m_strHeadKey = vNewValue
'End Property
Public Sub FillVoucherItems(domDestHead As DOMDocument, domDestBody As DOMDocument, blnCellCheck As Boolean)
    Dim lst As IXMLDOMNodeList
    Dim nod As IXMLDOMNode
    Dim strSFieldName As String
    Dim strDFieldName As String
    Dim strSSection As String
    Dim nodS As IXMLDOMNode
    Dim nodD As IXMLDOMElement
    
    Set nodS = domSourceHead.selectSingleNode("//z:row")
    Set nodD = domDestHead.selectSingleNode("//z:row")
    Set lst = domFillConfig.selectNodes("//z:row[@targetsection='T']")      '" & IIf(blnCellCheck, " and @cellcheck='True'", " and @cellcheck!='True'") & "
    For Each nod In lst
        If Not nod.Attributes.getNamedItem("sourcefld") Is Nothing Then
            strSFieldName = nod.Attributes.getNamedItem("sourcefld").Text
        Else
            strSFieldName = ""
        End If
        strDFieldName = nod.Attributes.getNamedItem("targetfld").Text
        If strSFieldName = "" Then
            strSFieldName = nod.Attributes.getNamedItem("cmemo").Text
            strSFieldName = ReplaceSysPara(strSFieldName)
            nodD.setAttribute strDFieldName, strSFieldName
        Else
            If Not nodS.Attributes.getNamedItem(strSFieldName) Is Nothing Then
                nodD.setAttribute strDFieldName, nodS.Attributes.getNamedItem(strSFieldName).nodeValue
            End If
        End If
    Next
    Set lst = domFillConfig.selectNodes("//z:row[@targetsection='B']")  '" & IIf(blnCellCheck, " and @cellcheck='True'", " and @cellcheck!='True'") & "
    Dim i As Long
    Dim lstDesLst As IXMLDOMNodeList
    Dim lstSorLst As IXMLDOMNodeList
    Set lstSorLst = domSourceBody.selectNodes("//z:row")
    Set lstDesLst = domDestBody.selectNodes("//z:row[@editprop='A']")
    i = lstDesLst.length - lstSorLst.length
    Dim strOption As String
    Dim strMemo As String
    Dim clsComp As New UsSaCompStr.clsCompStr
    For Each nodS In lstSorLst
        Set nodD = lstDesLst(i)
        For Each nod In lst
            If Not nod.Attributes.getNamedItem("coption") Is Nothing Then
                strOption = nod.Attributes.getNamedItem("coption").Text
            Else
                strOption = "0"
            End If
            If strOption = "2" Then
                If Not nod.Attributes.getNamedItem("cmemo") Is Nothing Then
                    strMemo = nod.Attributes.getNamedItem("cmemo").Text
                    strMemo = ReplaceSysPara(strMemo)
                    strMemo = ReplaceItemValues(strMemo, domDestHead.selectSingleNode("//z:row"), nodS)
                    strMemo = clsComp.GetExp(strMemo)
                Else
                    strMemo = "True"
                End If
            Else
                strMemo = "True"
            End If
            If CBool(strMemo) Then
                If Not nod.Attributes.getNamedItem("sourcefld") Is Nothing Then
                    strSFieldName = nod.Attributes.getNamedItem("sourcefld").Text
                Else
                    strSFieldName = ""
                End If
                strDFieldName = nod.Attributes.getNamedItem("targetfld").Text
                If strSFieldName = "" Then
                    strSFieldName = nod.Attributes.getNamedItem("cmemo").Text
                    strSFieldName = ReplaceSysPara(strSFieldName)
                    nodD.setAttribute strDFieldName, strSFieldName
                Else
                    If strOption = "4" Then
'                        Dim strSSection As String
                        Dim eleS As IXMLDOMElement
                        strSSection = nod.Attributes.getNamedItem("sourcesection").Text
                        strSFieldName = nod.Attributes.getNamedItem("sourcefld").Text
                        If strSSection = "T" Then
                            Set eleS = domDestHead.selectSingleNode("//z:row")
                            If Not eleS.Attributes.getNamedItem(strSFieldName) Is Nothing Then
                                strMemo = eleS.Attributes.getNamedItem(strSFieldName).nodeValue
                                nodD.setAttribute strDFieldName, strMemo
                            End If
                        End If
                    End If
                    If strOption = "3" Then
                        strMemo = ReplaceItemValues(strSFieldName, domDestHead.selectSingleNode("//z:row"), nodS)
                        strMemo = clsComp.GetExp(strMemo)
                        nodD.setAttribute strDFieldName, strMemo
                    Else
                        If Not nodS.Attributes.getNamedItem(strSFieldName) Is Nothing Then
                            nodD.setAttribute strDFieldName, nodS.Attributes.getNamedItem(strSFieldName).nodeValue
                        End If
                    End If
                End If
            End If
        Next
        i = i + 1
    Next
End Sub

Private Function ReplaceItemValues(strSql As String, eleHead As IXMLDOMElement, eleBody As IXMLDOMElement) As String
    Dim lngPos1 As Integer
    Dim lngPos2 As Integer
    Dim strFieldName As String
    Dim varField As Variant
    Dim strValue As String
    
    lngPos1 = InStr(1, strSql, "[")
    Do While lngPos1 > 0
        lngPos2 = InStr(1, strSql, "]")
        strFieldName = Mid(strSql, lngPos1 + 1, lngPos2 - lngPos1 - 1)
        varField = Split(strFieldName, ",")
        strValue = ""
        If UBound(varField) = 1 Then
            If varField(0) = "T" Then
                If eleHead.Attributes.getNamedItem(varField(1)) Is Nothing Then
                    strSql = Replace(strSql, "[" + varField(0) + "," + varField(1) + "]", "")
                Else
                    strValue = eleHead.Attributes.getNamedItem(varField(1)).nodeValue
                    strSql = Replace(strSql, "[" + varField(0) + "," + varField(1) + "]", strValue)
                End If
            Else
                If eleBody.Attributes.getNamedItem(varField(1)) Is Nothing Then
                    strSql = Replace(strSql, "[" + varField(0) + "," + varField(1) + "]", "")
                Else
                    strValue = eleBody.Attributes.getNamedItem(varField(1)).nodeValue
                    strSql = Replace(strSql, "[" + varField(0) + "," + varField(1) + "]", strValue)
                End If
            End If
        Else
            strSql = Replace(strSql, "[" + varField(0) & "]", varField(0) + "")
        End If
        lngPos1 = InStr(lngPos1 + Len(strValue), strSql, "[")
    Loop
    ReplaceItemValues = strSql
End Function
'Public Sub FillCellCheckItems(domDestHead As DOMDocument, domDestBody As DOMDocument, Optional blnCellCheck As Boolean = False)
'    Dim lst As IXMLDOMNodeList
'    Dim nod As IXMLDOMNode
'    Dim strSFieldName As String
'    Dim strDFieldName As String
'    Dim strSSection As String
'    Dim nodS As IXMLDOMNode
'    Dim nodD As IXMLDOMElement
'
'    Set nodS = domSourceHead.selectSingleNode("//z:row")
'    Set nodD = domDestHead.selectSingleNode("//z:row")
'    Set lst = domFillConfig.selectNodes("//z:row[@targetsection='T'" & IIf(blnCellCheck, " and @cellcheck='True'", " and @cellcheck!='True'") & "]")
'    For Each nod In lst
'        strSFieldName = nod.Attributes.getNamedItem("sourcefld").Text
'        strDFieldName = nod.Attributes.getNamedItem("targetfld").Text
'        If Not nodS.Attributes.getNamedItem(strSFieldName) Is Nothing Then
'            nodD.setAttribute strDFieldName, nodS.Attributes.getNamedItem(strSFieldName).nodeValue
'        End If
'    Next
'    Set lst = domFillConfig.selectNodes("//z:row[@targetsection='B' and @cellcheck='True']")
'    Dim i As Long
'    Dim lstDesLst As IXMLDOMNodeList
'    Dim lstSorLst As IXMLDOMNodeList
'    Set lstSorLst = domSourceBody.selectNodes("//z:row")
'    Set lstDesLst = domDestBody.selectNodes("//z:row[@editprop='A']")
'    i = lstDesLst.length - lstSorLst.length
'    For Each nodS In lstSorLst
'        Set nodD = lstDesLst(i)
'        For Each nod In lst
'            If Not nod.Attributes.getNamedItem("sourcefld") Is Nothing Then
'                strSFieldName = nod.Attributes.getNamedItem("sourcefld").Text
'            Else
'                strSFieldName = ""
'            End If
'            strDFieldName = nod.Attributes.getNamedItem("targetfld").Text
'            If strSFieldName = "" Then
'                strSFieldName = nod.Attributes.getNamedItem("cmemo").Text
'                nodD.setAttribute strDFieldName, strSFieldName
'            Else
'                If Not nodS.Attributes.getNamedItem(strSFieldName) Is Nothing Then
'                    nodD.setAttribute strDFieldName, nodS.Attributes.getNamedItem(strSFieldName).nodeValue
'                End If
'            End If
'        Next
'        i = i + 1
'    Next
'End Sub

Public Property Get blnSigleColumn() As Boolean
    blnSigleColumn = m_blnSigleColumn
End Property

Public Property Let blnSigleColumn(ByVal vNewValue As Boolean)
    m_blnSigleColumn = vNewValue
End Property

Public Property Get strHeadCellCheckFields() As String
    strHeadCellCheckFields = m_strHeadCellCheckFields
End Property

Public Property Let strHeadCellCheckFields(ByVal vNewValue As String)
    m_strHeadCellCheckFields = vNewValue
End Property
Public Property Get strBodyCellCheckFields() As String
    strBodyCellCheckFields = m_strBodyCellCheckFields
End Property

Public Property Let strBodyCellCheckFields(ByVal vNewValue As String)
    m_strBodyCellCheckFields = vNewValue
End Property


Public Property Get pageSize() As Long
    pageSize = m_PageSize
End Property

Public Property Let pageSize(ByVal vNewValue As Long)
    m_PageSize = vNewValue
End Property
Public Sub SavePageSize()
    Dim nod As IXMLDOMNode
    Dim strCardNum As String
    Dim strReferKey As String
    Dim strSql As String
    
    Set nod = domReferConfig.selectSingleNode("//z:row")
    strCardNum = nod.Attributes.getNamedItem("cardnum").Text
    strReferKey = nod.Attributes.getNamedItem("referkey").Text
    strSql = "update sa_refervoucherconfig set pagesize=" & pageSize & " where cardnum='" & strCardNum & "' and referkey='" + strReferKey + "'"
    DBConn.Execute ConvertSQLString(strSql)
End Sub


Public Property Get strAuth() As String
    strAuth = m_strAuth
End Property

Public Property Let strAuth(ByVal vNewValue As String)
    m_strAuth = vNewValue
End Property

Public Property Get strSelectType() As String
    strSelectType = m_strSelectType
End Property

Public Property Let strSelectType(ByVal vNewValue As String)
    m_strSelectType = vNewValue
End Property

Public Property Get strFillType() As String
    strFillType = m_strFillType
End Property

Public Property Let strFillType(ByVal vNewValue As String)
    m_strFillType = vNewValue
End Property

Public Property Get strReferName() As String
    strReferName = m_strReferName
End Property

Public Property Let strReferName(ByVal vNewValue As String)
    m_strReferName = vNewValue
End Property

Public Property Get strMainKey() As String
    strMainKey = m_strMainKey
End Property

Public Property Let strMainKey(ByVal vNewValue As String)
    m_strMainKey = vNewValue
End Property
Public Property Get strDetailKey() As String
    strDetailKey = m_strDetailKey
End Property

Public Property Let strDetailKey(ByVal vNewValue As String)
    m_strDetailKey = vNewValue
End Property

Public Property Get StrFilterName() As String
    StrFilterName = m_strFilterName
End Property

Public Property Let StrFilterName(ByVal vNewValue As String)
    m_strFilterName = vNewValue
End Property
Public Function ColumnSet(blnHead As Boolean, ctlrefervoucher As Object) As Boolean
    Dim strColStr As String
    If blnHead Then
    
        strColStr = clsHeadColumn.GetColFormatStr()
        ctlrefervoucher.HeadColSetXml = strColStr
        ctlrefervoucher.HeadList.InitHead strColStr
    Else
'        strColStr = clsBodyColumn.ColumnSet()
        strColStr = clsBodyColumn.GetColFormatStr()
        ctlrefervoucher.BodyColSetXml = strColStr
        ctlrefervoucher.BodyList.InitHead strColStr
'        SelectedStr = ""
'        ctlrefervoucher.BodyList.RecordCount = 0
    End If
    ColumnSet = True
End Function

Public Function CheckBeforeSelected(ctlReferVouch As Object, lngRow As Long) As String
    Dim varRows As Variant
    Dim varFlds As Variant
    Dim i As Long
    Dim j As Long
    Dim lngCol As Long
    Dim strFvalue As String
    Dim strOvalue As String
    Dim lngORow As Long
    
    varRows = Split(SelectedStr, ",")
    If UBound(varRows) > 0 Then
        varFlds = Split(strCheckFlds, ",")
        If UBound(varFlds) >= 0 Then
            lngORow = varRows(0)
            For j = 0 To UBound(varFlds)
                lngCol = ctlReferVouch.HeadList.GetColIndex(varFlds(j))
                strFvalue = ctlReferVouch.HeadTextMatrix(lngORow, lngCol)
                strOvalue = ctlReferVouch.HeadTextMatrix(lngRow, lngCol)
                If strFvalue <> strOvalue Then
                    CheckBeforeSelected = varFlds(j)
                    Exit Function
                End If
            Next
        End If
    End If
End Function


Public Property Get SelectedStr() As String
    SelectedStr = m_SelectedStr
End Property

Public Property Let SelectedStr(ByVal vNewValue As String)
    m_SelectedStr = vNewValue
End Property
Public Sub ClickHeadList(ctlReferVouch As Object, ByVal Row As Long, bChange As Boolean, sErr As String)
'    Dim strStyle As String
    Dim i As Long
    Dim strFieldName As String
    
    bChange = True
    sErr = ""
'    strStyle = clsRefeVoucher.strSelectType
    If strSelectType = "uniquemain" Then
        For i = 1 To ctlReferVouch.HeadList.Rows - 1
            If i <> Row Then
                ctlReferVouch.HeadList.TextMatrix(i, 0) = ""
            End If
        Next
        If ctlReferVouch.HeadList.TextMatrix(Row, 0) <> "" Then
            SelectedStr = CStr(Row)
            If clsBodyColumn.strColumnKey <> "" Then
                SetBodyData ctlReferVouch, Row, True
            End If
        Else
            SelectedStr = ""
        End If
    Else
        If ctlReferVouch.HeadList.TextMatrix(Row, 0) <> "" Then
            strFieldName = CheckBeforeSelected(ctlReferVouch, Row)
            If strFieldName = "" Then
                SelectedStr = SelectedStr & " " & CStr(Row) & ","
                If clsBodyColumn.strColumnKey <> "" Then
                    SetBodyData ctlReferVouch, Row, True
                End If
            Else
                bChange = False
                sErr = GetStringPara("U8.SA.xsglsql_2.refervoucher.selectcheck", ctlReferVouch.HeadTextMatrix(0, ctlReferVouch.HeadList.GetColIndex(strFieldName)))
            End If
        Else
            RemoveSelected ctlReferVouch, Row
        End If
    End If
End Sub
Public Function RemoveSelected(ctlReferVouch As Object, Row As Long)
    Dim lngPos1 As Long
    Dim lngPos2 As Long
    Dim strValue As String
    
    strValue = Row
    lngPos1 = InStr(1, SelectedStr, strValue & ",")
    If lngPos1 > 0 Then
'        lngpos2 = InStr(lngPos1, SelectedStr, ",")
        SelectedStr = Replace(SelectedStr, " " & strValue & ",", "")
    End If
End Function
Public Sub HeadSelectAll(ctlReferVouch As Object, strError As String)
    Dim i As Long
    Dim strErr As String
    Dim nod As IXMLDOMNode
    Dim strKeyName As String
    Dim strKeyValue As String
    
    Set nod = domReferConfig.selectSingleNode("//z:row")
    strKeyName = nod.Attributes.getNamedItem("mainuniquekey").Text
    
    If ctlReferVouch.HeadList.Rows = 1 Then Exit Sub
    ctlReferVouch.AddSumLine ctlReferVouch.HeadList.GetGridBody(), False
    ctlReferVouch.AddSumLine ctlReferVouch.BodyList.GetGridBody(), False
    For i = 1 To ctlReferVouch.HeadList.Rows - 1
        If InStr(1, SelectedStr, " " & i & ",") <= 0 Then
            ClickHeadList ctlReferVouch, i, False, strErr
            If strErr <> "" Then
                ctlReferVouch.HeadList.TextMatrix(i, 0) = ""
                strError = strError & IIf(strError = "", "", Chr(13)) & GetStringPara("U8.SA.xsglsql_2.refervoucher.atline", CStr(i)) & strErr
            End If
        Else
            strKeyValue = ctlReferVouch.HeadList.TextMatrix(i, ctlReferVouch.GetHeadColIndex(strKeyName))
            SelectBodyDatas ctlReferVouch, strKeyName, strKeyValue
        End If
    Next
    ctlReferVouch.AddSumLine ctlReferVouch.HeadList.GetGridBody(), True
'    ctlReferVouch.BodyList.RecordCount = ctlReferVouch.BodyList.GetGridBody().rows - 1
'    ctlReferVouch.AddSumLine ctlReferVouch.BodyList.GetGridBody(), True
End Sub
Public Sub HeadUnSelectAll(ctlReferVouch As Object)
    ctlReferVouch.RemoveBodyAll
    SelectedStr = ""
    ctlReferVouch.AddSumLine ctlReferVouch.BodyList.GetGridBody(), False
    ctlReferVouch.BodyList.RecordCount = ctlReferVouch.BodyList.GetGridBody().Rows - 2
    ctlReferVouch.AddSumLine ctlReferVouch.BodyList.GetGridBody(), True
End Sub

Public Property Get strCheckFlds() As String
    strCheckFlds = m_strCheckFlds
End Property

Public Property Let strCheckFlds(ByVal vNewValue As String)
    m_strCheckFlds = vNewValue
End Property


Private Function ReplaceResId(strSource As String) As String
    Dim lngPos1 As Long
    Dim lngPos2 As Long
    Dim strValue As String
    Dim strNew As String
    
    Dim var As Variant
    lngPos1 = InStr(1, strSource, "{")
    If lngPos1 > 0 Then
        Do While lngPos1 > 0
            lngPos2 = InStr(lngPos1, strSource, "}")
            strValue = Mid(strSource, lngPos1 + 1, lngPos2 - lngPos1 - 1)
            var = Split(strValue, ":")
            If var(0) = "res" Then
                strNew = GetString(var(1))
            End If
            strSource = Replace(strSource, "{" & strValue & "}", strNew)
            lngPos1 = InStr(1, strSource, "{")
        Loop
        ReplaceResId = strSource
    Else
        ReplaceResId = strSource
    End If
End Function


Public Property Get blnReturnFlag() As Boolean
    blnReturnFlag = m_blnReturnFlag
End Property

Public Property Let blnReturnFlag(ByVal vNewValue As Boolean)
    m_blnReturnFlag = vNewValue
End Property



