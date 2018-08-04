VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsReferVoucher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim m_strFilter As String
Dim m_blnSigleColumn As Boolean

Dim m_PageSize As Long
Dim m_strAuth As String
Dim m_bSelectSingle As Boolean           '标记表头是否只支持单选
Dim m_strMainKey As String
Dim m_strDetailKey As String
Dim m_SelectedStr As String

Dim clsHeadColumn As New clsColumnSet
Dim clsBodyColumn As New clsColumnSet

Private strFilterKey As String, strFilterSubID As String   ''过滤器名称和ID
Attribute strFilterSubID.VB_VarUserMemId = 1073938442
Private strhelpID As String
Attribute strhelpID.VB_VarUserMemId = 1073938444

Private strMakerAuth As String    '制单人权限的条件
Attribute strMakerAuth.VB_VarUserMemId = 1073938445
Private strOtherFilter As String    '其他过滤条件
Attribute strOtherFilter.VB_VarUserMemId = 1073938446
Private strFrmCaption As String    '参照窗体的caption
Attribute strFrmCaption.VB_VarUserMemId = 1073938447
Private strHeadKey As String      '主列表的columnkey ,对应预置的栏目表ckey
Attribute strHeadKey.VB_VarUserMemId = 1073938448
Private strBodyKey As String      '子表的columnkey ,对应预置的栏目表ckey
Attribute strBodyKey.VB_VarUserMemId = 1073938449
Private strButtons As String      '添加的自定义按钮
Attribute strButtons.VB_VarUserMemId = 1073938450
Private strMainDataSource As String   ' 主表的数据源
Attribute strMainDataSource.VB_VarUserMemId = 1073938451
Private strDetailDataSource As String   '子表的数据源
Attribute strDetailDataSource.VB_VarUserMemId = 1073938452
Private strDefaultFilter As String      '默认的过滤条件
Attribute strDefaultFilter.VB_VarUserMemId = 1073938453
Private m_bHeadEnabled As Boolean       '表头可编辑属性
Attribute m_bHeadEnabled.VB_VarUserMemId = 1073938454
Private m_bBodyEnabled As Boolean       '表体可编辑属性
Attribute m_bBodyEnabled.VB_VarUserMemId = 1073938455

Dim m_strHeadCellCheckFields As String
Attribute m_strHeadCellCheckFields.VB_VarUserMemId = 1073938456
Dim m_strBodyCellCheckFields As String
Attribute m_strBodyCellCheckFields.VB_VarUserMemId = 1073938457
Dim m_strFillType As String
Attribute m_strFillType.VB_VarUserMemId = 1073938458
Dim m_strReferName As String
Attribute m_strReferName.VB_VarUserMemId = 1073938459
Dim m_strFilterName As String
Attribute m_strFilterName.VB_VarUserMemId = 1073938460
Dim m_blnReturnFlag As String
Attribute m_blnReturnFlag.VB_VarUserMemId = 1073938461
Dim m_strCheckFlds As String              '主表需要检查的字段
Attribute m_strCheckFlds.VB_VarUserMemId = 1073938462
Dim m_strCheckMsg As String              '主表需要检查的提示信息
Attribute m_strCheckMsg.VB_VarUserMemId = 1073938463

Public Property Get strCheckMsg() As String
    strCheckMsg = m_strCheckMsg
End Property

Public Property Let strCheckMsg(ByVal vNewValue As String)
    m_strCheckMsg = vNewValue
End Property

Public Property Get strCheckFlds() As String
    strCheckFlds = m_strCheckFlds
End Property

Public Property Let strCheckFlds(ByVal vNewValue As String)
    m_strCheckFlds = vNewValue
End Property

Public Property Get HeadEnabled() As Boolean
    HeadEnabled = m_bHeadEnabled
End Property

Public Property Let HeadEnabled(ByVal vNewValue As Boolean)
    m_bHeadEnabled = vNewValue
End Property

Public Property Get BodyEnabled() As Boolean
    BodyEnabled = m_bBodyEnabled
End Property

Public Property Let BodyEnabled(ByVal vNewValue As Boolean)
    m_bBodyEnabled = vNewValue
End Property

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

Public Property Get strAuth() As String
    strAuth = m_strAuth
End Property

Public Property Let strAuth(ByVal vNewValue As String)
    m_strAuth = vNewValue
End Property

Public Property Get bSelectSingle() As Boolean
    bSelectSingle = m_bSelectSingle
End Property

Public Property Let bSelectSingle(ByVal vNewValue As Boolean)
    m_bSelectSingle = vNewValue
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

Public Property Get strFilter() As String
    strFilter = m_strFilter
End Property

Public Property Let strFilter(ByVal vNewValue As String)
    m_strFilter = vNewValue
End Property

Public Property Get SelectedStr() As String
    SelectedStr = m_SelectedStr
End Property

Public Property Let SelectedStr(ByVal vNewValue As String)
    m_SelectedStr = vNewValue
End Property

Public Property Get blnReturnFlag() As Boolean
    blnReturnFlag = m_blnReturnFlag
End Property

Public Property Let blnReturnFlag(ByVal vNewValue As Boolean)
    m_blnReturnFlag = vNewValue
End Property

Public Property Get FilterKey() As String
    FilterKey = strFilterKey
End Property

Public Property Let FilterKey(ByVal vNewValue As String)
    strFilterKey = vNewValue
End Property

Public Property Get HelpID() As String
    HelpID = strhelpID
End Property

Public Property Let HelpID(ByVal vNewValue As String)
    strhelpID = vNewValue
End Property

Public Property Get MakerAuth() As String
    MakerAuth = strMakerAuth
End Property

Public Property Let MakerAuth(ByVal vNewValue As String)
    strMakerAuth = vNewValue
End Property

Public Property Get OtherFilter() As String
    OtherFilter = strOtherFilter
End Property

Public Property Let OtherFilter(ByVal vNewValue As String)
    strOtherFilter = vNewValue
End Property

Public Property Get FrmCaption() As String
    FrmCaption = strFrmCaption
End Property

Public Property Let FrmCaption(ByVal vNewValue As String)
    strFrmCaption = vNewValue
End Property

Public Property Get FilterSubID() As String
    FilterSubID = strFilterSubID
End Property

Public Property Let FilterSubID(ByVal vNewValue As String)
    strFilterSubID = vNewValue
End Property

Public Property Get HeadKey() As String
    HeadKey = strHeadKey
End Property

Public Property Let HeadKey(ByVal vNewValue As String)
    strHeadKey = vNewValue
End Property

Public Property Get BodyKey() As String
    BodyKey = strBodyKey
End Property

Public Property Let BodyKey(ByVal vNewValue As String)
    strBodyKey = vNewValue
End Property

Public Property Get Buttons() As String
    Buttons = strButtons
End Property

Public Property Let Buttons(ByVal vNewValue As String)
    strButtons = vNewValue
End Property

Public Property Get MainDataSource() As String
    MainDataSource = strMainDataSource
End Property

Public Property Let MainDataSource(ByVal vNewValue As String)
    strMainDataSource = vNewValue
End Property

Public Property Get DetailDataSource() As String
    DetailDataSource = strDetailDataSource
End Property

Public Property Let DetailDataSource(ByVal vNewValue As String)
    strDetailDataSource = vNewValue
End Property

Public Property Get DefaultFilter() As String
    DefaultFilter = strDefaultFilter
End Property

Public Property Let DefaultFilter(ByVal vNewValue As String)
    strDefaultFilter = vNewValue
End Property

'初始化参照生单控件
Public Sub InitReferVoucher(ctlrefervoucher As ctlReferMakeVouch)

    Dim domButtons As DOMDocument
    If strBodyKey = "" Then          '不配置子表列信息则只显示单独的主表信息
        blnSigleColumn = True
    Else
        blnSigleColumn = False
    End If

    Set domButtons = New DOMDocument
    domButtons.loadXML strButtons

    If Not ctlrefervoucher Is Nothing Then
        ctlrefervoucher.Init blnSigleColumn, Nothing, domButtons, 1, 1, 1, Nothing, g_oLogin
        ctlrefervoucher.SetKey strHeadKey, strBodyKey
        ctlrefervoucher.SetFilterOne
        ctlrefervoucher.UFShowBody.Visible = False
    End If
    clsHeadColumn.Init strHeadKey, "0"
    Call ColumnSet(True, ctlrefervoucher)

    If strBodyKey <> "" Then
        clsBodyColumn.Init strBodyKey, "0"
        Call ColumnSet(False, ctlrefervoucher)
    End If

End Sub

Private Sub Class_Terminate()
    Set clsHeadColumn = Nothing
    Set clsBodyColumn = Nothing
End Sub

Private Sub GetOnePageDatas(domDatas As DOMDocument, Optional intRows As Long, Optional intCurrentPage As Long, Optional intTotalPages As Long, Optional intTotalRecords As Long)
    Dim strSql As String
    Dim nod As IXMLDOMNode
    Dim strSelect As String
    '    Dim strFrom As String
    Dim strOrder As String
    Dim lngMin As Long
    Dim lngMax As Long

    Dim strTempTblName As String
    Dim strWhere As String

    strOrder = clsHeadColumn.GetOrderString
'    strOrder = ""
    strTempTblName = "sarefertemp" + CStr(Rnd())

    If strDefaultFilter <> "" Then strWhere = strDefaultFilter
    If strFilter <> "" Then
        If strWhere = "" Then
            strWhere = strFilter
        Else
            strWhere = "(" & strWhere & ") and (" & strFilter & ")"
        End If
    End If
    If strAuth <> "" Then
        strWhere = strWhere & IIf(strWhere = "", "", " and ( " & strAuth & ")")
    End If
    strSql = "if exists(select * from tempdb..sysobjects where name='" + strTempTblName + "' and xtype='U') drop table tempdb..[" + strTempTblName + "] "
    'modify by tiger 2017-08-06
    If strDetailDataSource = "" Then
        strSql = strSql & "select identity(int) as  identityautoid ," + strMainKey + " as mainkey into tempdb..[" + strTempTblName + "] from " + strMainDataSource + " WITH (NOLOCK) " + IIf(strWhere = "", "", " where " + strWhere) & IIf(strOrder = "", "", " order by " & strOrder)
    Else
        strSql = strSql & "select identity(int) as  identityautoid ," + strMainKey + " as mainkey into tempdb..[" + strTempTblName + "] from " + strMainDataSource _
               + " where " + strMainKey + " in (select distinct " + strMainDataSource + "." + strMainKey + " from " & strMainDataSource & " WITH (NOLOCK) inner join " & strDetailDataSource & " WITH (NOLOCK) on " & strMainDataSource & "." & strMainKey & "=" & strDetailDataSource & "." & strMainKey _
               + IIf(strWhere = "", "", " where " + strWhere) & ")" & IIf(strOrder = "", "", " order by " & strOrder)
    End If
    g_Conn.Execute ConvertSQLString(strSql)
    Dim rst As New ADODB.Recordset
    rst.CursorLocation = adUseClient
    strSql = "select * from tempdb..[" + strTempTblName + "] order by identityautoid"
    rst.Open ConvertSQLString(strSql), g_Conn, adOpenForwardOnly, adLockReadOnly
    If Not rst.EOF Then
        intTotalRecords = rst.RecordCount
        If intRows <> 0 Then
            rst.pageSize = intRows
        Else
            rst.pageSize = rst.RecordCount
        End If
        If intCurrentPage > rst.pageCount Then intCurrentPage = rst.pageCount
        If intCurrentPage <= 1 Then intCurrentPage = 1
        intTotalPages = rst.pageCount

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
    strSql = "select " + strSelect + " from " & strMainDataSource & " WITH (NOLOCK) inner join tempdb..[" + strTempTblName + "] on tempdb..[" + strTempTblName + "].mainkey=" & strMainDataSource & "." + strMainKey + " where tempdb..[" + strTempTblName + "].identityautoid>=" + CStr(lngMin) + " and tempdb..[" + strTempTblName + "].identityautoid<=" + CStr(lngMax) + " order by tempdb..[" + strTempTblName + "].identityautoid"
    rst.Open ConvertSQLString(strSql), g_Conn, adOpenForwardOnly, adLockReadOnly, adCmdText
    rst.Save domDatas, adPersistXML
    rst.Close
    Set rst = Nothing
End Sub

Public Sub RemoveHeadLines(domHead As DOMDocument, domBody As DOMDocument)
    Dim lst As IXMLDOMNodeList
    Dim ele As IXMLDOMElement

    Dim strKeyValue As String
    Dim lstBody As IXMLDOMNodeList
    Dim nod As IXMLDOMElement

    Set lst = domHead.selectNodes("//z:row")
    For Each ele In lst
        strKeyValue = GetNodeAtrVal(ele, strMainKey)
'        Set lstBody = domBody.selectNodes("//z:row[@" + strMainKey + "='" + strKeyValue + "']")
'        If lstBody.Length <= 0 Then
'            domHead.selectSingleNode("//rs:data").removeChild ele
'        End If
    Next

End Sub

Private Sub GetBodyDatas(strKeyValue As String, domHeadDatas As DOMDocument, domDatas As DOMDocument)
    Dim strSql As String
    Dim strSelect As String
    Dim strOrder As String
    Dim nod As IXMLDOMElement

    Dim strWhere As String
    Dim rst As New ADODB.Recordset

    rst.CursorLocation = adUseClient
    strSelect = clsBodyColumn.GetSqlString
    strOrder = clsBodyColumn.GetOrderString

       strWhere = "(1=1)"
'    If strDefaultFilter <> "" Then
'        strWhere = "(" & strWhere & ") and (" & strDefaultFilter & ")"
'    End If
'    If strFilter <> "" Then
'        strWhere = IIf(strWhere = "", "", "(" & strWhere & ") and (") & strFilter & ")"
'    End If
'    If strAuth <> "" Then
'        strWhere = strWhere & IIf(strWhere = "", "", " and (1=1 and " & strAuth & ")")
'    End If

  strWhere = strWhere & IIf(sAuth_PositionW = "", "", " and (isnull(cposition,'')='' or cposition in (" & sAuth_PositionW & "))")
     strWhere = strWhere & IIf(sAuth_invW = "", "", " and cinvcode in (select cinvcode from inventory where iid in (" & sAuth_invW & "))")
     strWhere = strWhere & IIf(sAuth_WareHouseW = "", "", " and (ISNULL(cwhcode,N'')='' OR cwhcode in (" & sAuth_WareHouseW & "))")

'    If strDetailDataSource = "" Then
'        strSql = strSql & "select identity(int) as  identityautoid ," + strMainKey + " as mainkey into tempdb..[" + strTempTblName + "] from " + strMainDataSource + " WITH (NOLOCK) " + IIf(strWhere = "", "", " where " + strWhere) & IIf(strOrder = "", "", " order by " & strOrder)
'    Else
'        strSql = strSql & "select identity(int) as  identityautoid ," + strMainKey + " as mainkey into tempdb..[" + strTempTblName + "] from " + strMainDataSource _
'               + " where " + strMainKey + " in (select distinct " + strMainDataSource + "." + strMainKey + " from " & strMainDataSource & " WITH (NOLOCK) inner join " & strDetailDataSource & " WITH (NOLOCK) on " & strMainDataSource & "." & strMainKey & "=" & strDetailDataSource & "." & strMainKey _
'               + IIf(strWhere = "", "", " where " + strWhere) & ")" & IIf(strOrder = "", "", " order by " & strOrder)
'    End If


    'add by tiger 2017-8-6
    strWhere = " " & strDetailKey & " in(select distinct " + strDetailDataSource + "." + strDetailKey + " from " & strMainDataSource & " WITH (NOLOCK) inner join " & strDetailDataSource & " WITH (NOLOCK) on " & strMainDataSource & "." & strMainKey & "=" & strDetailDataSource & "." & strMainKey & " where " & strFilter & ")"
    strSql = "select " & strSelect & " from " & strDetailDataSource & " where " & strDetailDataSource & "." & strMainKey & "=" & strKeyValue & " and " _
            + IIf(strWhere = "", "(1=1)", strWhere) & IIf(strOrder = "", "", " order by " & strOrder)
    rst.Open ConvertSQLString(strSql), g_Conn, adOpenForwardOnly, adLockReadOnly, adCmdText
    rst.Save domDatas, adPersistXML
    rst.Close
    Set rst = Nothing
End Sub

Public Sub SetHeadData(ctlrefervoucher As Object, Optional intRows As Long, Optional intCurrentPage As Long, Optional intTotalPages As Long)
    Dim lngRecordCount As Long
    Dim domHeadData As New DOMDocument
    GetOnePageDatas domHeadData, intRows, intCurrentPage, intTotalPages, lngRecordCount
    If Not domHeadData Is Nothing Then
        ctlrefervoucher.SetHeadDom domHeadData
        ctlrefervoucher.HeadList.RecordCount = lngRecordCount
    End If
    Set domHeadData = Nothing
    SelectedStr = ""
End Sub

Private Sub RemoveBodyDatas(ctlrefervoucher As Object, strKeyName As String, strKeyValue As String)
    Dim i As Long
    Dim lngPos As Long
    lngPos = ctlrefervoucher.GetBodyColIndex(strKeyName)
    Dim tmpSumStyle As Variant
    tmpSumStyle = ctlrefervoucher.BodyList.SumStyle
    ctlrefervoucher.BodyList.SumStyle = 0
    For i = ctlrefervoucher.BodyList.rows - 1 To 1 Step -1
        If ctlrefervoucher.BodyList.TextMatrix(i, lngPos) = strKeyValue Then
            ctlrefervoucher.BodyList.RemoveItem i
        End If
    Next i
    ctlrefervoucher.BodyList.SumStyle = tmpSumStyle
    ctlrefervoucher.RefreshSumLine ctlrefervoucher.BodyList.GetGridBody, True

End Sub

Public Sub SelectBodyDatas(ctlrefervoucher As Object, strKeyName As String, strKeyValue As String)
    Dim i As Long
    Dim lngPos As Long
    Dim lngPosSel As Long

    lngPos = ctlrefervoucher.GetBodyColIndex(strKeyName)
    lngPosSel = ctlrefervoucher.GetBodyColIndex("selcol")
    For i = ctlrefervoucher.BodyList.rows - 1 To 1 Step -1
        If ctlrefervoucher.BodyList.TextMatrix(i, lngPos) = strKeyValue Then
            ctlrefervoucher.BodyList.TextMatrix(i, lngPosSel) = "Y"
        End If
    Next i
End Sub

Public Sub SetBodyData(ctlrefervoucher As ctlReferMakeVouch, lngCurRow As Long, Optional SelectIt As Boolean = False)

    Dim strKeyValue As String
    Dim domline As New DOMDocument
    Dim ele As IXMLDOMElement
    Dim oND As IXMLDOMNode
    Dim domHead As DOMDocument

    strKeyValue = ctlrefervoucher.HeadList.TextMatrix(lngCurRow, ctlrefervoucher.GetHeadColIndex(strMainKey))
    If strKeyValue = "" Then Exit Sub
    If ctlrefervoucher.HeadList.TextMatrix(lngCurRow, 0) = "" Then
        RemoveBodyDatas ctlrefervoucher, strMainKey, strKeyValue
        RemoveSelected ctlrefervoucher, lngCurRow
    Else
        Dim domBodyData As New DOMDocument
        Set domHead = ctlrefervoucher.GetHeadLine(lngCurRow)

        GetBodyDatas strKeyValue, domHead, domBodyData
        Dim strError As String
        If Not domBodyData Is Nothing Then
            'FormatVouchList ctlrefervoucher.BodyList, domBodyFormat        ' 格式化列表栏目数据精度格式
            If m_bSelectSingle Then
                ctlrefervoucher.RemoveBodyAll
            End If
            If ctlrefervoucher.BodyList.rows <= 1 Then
                Set domline = domBodyData.cloneNode(True)
                Set oND = domline.selectSingleNode("//rs:data")
                For Each ele In domline.selectNodes("//z:row")
                    oND.removeChild ele
                Next
                ctlrefervoucher.SetBodyDom domline
                'ctlrefervoucher.RemoveBodyAll
            End If

            Set domline = domBodyData.cloneNode(True)
            '            Dim ele As IXMLDOMElement
            Dim eleline As IXMLDOMElement
            Dim TmpLinedom As DOMDocument

            Set oND = domline.selectSingleNode("//rs:data")
            For Each eleline In domline.selectNodes("//z:row")
                oND.removeChild eleline
            Next
            Set TmpLinedom = domline.cloneNode(True)
            '            ctlrefervoucher.bAppending = True
            Dim tmpSumStyle As Variant
            RemoveSumLine ctlrefervoucher.BodyList.GetGridBody()
            Dim i As Long
            i = ctlrefervoucher.BodyList.GetGridBody().rows
            tmpSumStyle = ctlrefervoucher.BodyList.SumStyle
            ctlrefervoucher.BodyList.SumStyle = 0                '不合计
            For Each ele In domBodyData.selectNodes("//z:row")
                Set domline = TmpLinedom.cloneNode(True)
                Set oND = domline.selectSingleNode("//rs:data")
                oND.appendChild ele
                'domLine.selectSingleNode("//rs:data").appendChild ele
                ctlrefervoucher.AddBodyLine i, domline
                If iSinvCZ = False Then
                If SelectIt Then ctlrefervoucher.BodyList.TextMatrix(i, 0) = "Y"
                i = i + 1
                End If
                '                oND.removeChild ele
            Next
            ctlrefervoucher.SelectBodyAll
            Set TmpLinedom = Nothing
            ctlrefervoucher.BodyList.SumStyle = tmpSumStyle
            ctlrefervoucher.RefreshSumLine ctlrefervoucher.BodyList.GetGridBody(), True
        End If
        Set domline = Nothing
        Set domBodyData = Nothing
    End If
End Sub

Private Sub RemoveSumLine(GridBody As Object)
    On Error Resume Next
    With GridBody
        If .rows - 1 >= 0 Then
            If .TextMatrix(.rows - 1, 0) = GetString("U8.ST.Default.00418") Then    '"合计"
                .RemoveItem .rows - 1
            End If
        End If
    End With
End Sub

Private Function ReplaceItemValues(strSql As String, eleHead As IXMLDOMElement, eleBody As IXMLDOMElement) As String
    Dim lngPos1 As Integer
    Dim lngpos2 As Integer
    Dim strFieldName As String
    Dim varField As Variant
    Dim strValue As String

    lngPos1 = InStr(1, strSql, "[")
    Do While lngPos1 > 0
        lngpos2 = InStr(1, strSql, "]")
        strFieldName = Mid(strSql, lngPos1 + 1, lngpos2 - lngPos1 - 1)
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
            strSql = Replace(strSql, "[" + varField(0) & "]", varField(0))
        End If
        lngPos1 = InStr(lngPos1 + Len(strValue), strSql, "[")
    Loop
    ReplaceItemValues = strSql
End Function

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

Public Sub ClickHeadList(ctlReferVouch As Object, ByVal row As Long, bChange As Boolean, sErr As String)
'    Dim strStyle As String
    Dim i As Long
    Dim strFieldName As String

    bChange = True
    sErr = ""
    '    strStyle = clsRefeVoucher.strSelectType
    If m_bSelectSingle Then
        For i = 1 To ctlReferVouch.HeadList.rows - 1
            If i <> row Then
                ctlReferVouch.HeadList.TextMatrix(i, 0) = ""
            End If
        Next
        If ctlReferVouch.HeadList.TextMatrix(row, 0) <> "" Then
            SelectedStr = CStr(row)
            If clsBodyColumn.strColumnKey <> "" Then
                SetBodyData ctlReferVouch, row, True
            End If
        Else
            SelectedStr = ""
        End If
    Else
        If ctlReferVouch.HeadList.TextMatrix(row, 0) <> "" Then
            strFieldName = CheckBeforeSelected(ctlReferVouch, row)
            If strFieldName = "" Then
                SelectedStr = SelectedStr & " " & CStr(row) & ","
                If clsBodyColumn.strColumnKey <> "" Then
                    SetBodyData ctlReferVouch, row, True
                End If
            Else
                bChange = False
                sErr = GetStringPara("U8.SA.xsglsql_2.refervoucher.selectcheck", ctlReferVouch.HeadTextMatrix(0, ctlReferVouch.HeadList.GetColIndex(strFieldName)))
            End If
        Else
            RemoveSelected ctlReferVouch, row
        End If
    End If
End Sub

Public Function RemoveSelected(ctlReferVouch As Object, row As Long)
    Dim lngPos1 As Long
    Dim lngpos2 As Long
    Dim strValue As String

    strValue = row
    lngPos1 = InStr(1, SelectedStr, strValue & ",")
    If lngPos1 > 0 Then
        '        lngpos2 = InStr(lngPos1, SelectedStr, ",")
        SelectedStr = Replace(SelectedStr, " " & strValue & ",", "")
    End If
End Function

Private Function ReplaceResId(strSource As String) As String
    Dim lngPos1 As Long
    Dim lngpos2 As Long
    Dim strValue As String
    Dim strNew As String

    Dim var As Variant
    lngPos1 = InStr(1, strSource, "{")
    If lngPos1 > 0 Then
        Do While lngPos1 > 0
            lngpos2 = InStr(lngPos1, strSource, "}")
            strValue = Mid(strSource, lngPos1 + 1, lngpos2 - lngPos1 - 1)
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

'对过滤获得的SQL串进行加工
Public Function convertWhere(backFilter As UFGeneralFilter.FilterSrv) As String
'    Dim i As Integer
'    Dim newWhere As String
'    Dim DispatchListType As String
    On Error GoTo err1:
'            For i = 1 To backFilter.FilterList.Count
'                Select Case LCase(backFilter.FilterList.Item(i).name)
'                Case LCase("DispatchListType")
'                    Select Case backFilter.FilterList.Item(i).varValue
'                    Case "蓝字记录"
'                        newWhere = newWhere & " and isnull(isum,0)>=0"
'                    Case "红字记录"
'                        newWhere = newWhere & " and isnull(isum,0)<=0"
'                    Case Else
'                        newWhere = newWhere & ""
'                    End Select
'                    backFilter.FilterList(i).LngGrouptoken = 1
'
'                Case LCase("cCusCode")
'                    If backFilter.FilterList.Item(i).varValue <> "" Then
'                        strClientCode = backFilter.FilterList.Item(i).varValue
'                    Else
'                        strClientCode = ""
'                    End If
'                Case LCase("cInvCode")
'                    If backFilter.FilterList.Item(i).varValue <> "" Then
'                        strInvCode = backFilter.FilterList.Item(i).varValue
'                    Else
'                        strInvCode = ""
'                    End If
'                Case LCase("isShowOverOrder") '是否显示已执行完的订单
'                    If backFilter.FilterList.Item(i).varValue = "是" Then '显示
'                        isShowOverOrder = 2
'                    ElseIf backFilter.FilterList.Item(i).varValue = "否" Then '不显示
'                        isShowOverOrder = 1
'                    Else
'                        isShowOverOrder = 0 '不选
'                    End If
'                    backFilter.FilterList(i).LngGrouptoken = 1
'                Case LCase("cinvoicecompany")
'                    If backFilter.FilterList.Item(i).varValue <> "" Then
'                        newWhere = newWhere & " and (ccuscode =N'" + backFilter.FilterList.Item(i).varValue + "' or ccuscode in (select ccuscode from sa_invoicecustomers where cinvoicecompany=N'" + backFilter.FilterList.Item(i).varValue + "'))"
'                        backFilter.FilterList(i).LngGrouptoken = 1
'                    End If
'                End Select
'            Next

    convertWhere = IIf(backFilter.GetWhereWithToken(0) = "", " 1=1 ", backFilter.GetWhereWithToken(0))

    Exit Function
err1:
    MsgBox Err.Description, vbInformation, GetString("U8.DZ.JA.Res030")
    '    newWhere = m_where
End Function

Public Function CheckHeadSelecting(ctlmakevoucher As ctlReferMakeVouch) As Boolean
    Dim strFieldName As String
    Dim strFiledValue As String
    Dim strcodedep As String
    Dim strcodeps As String
    Dim cexch_name As String
    Dim row As Long
    Dim varTmp As Variant
    Dim varTmpCheckFld As Variant
    Dim strTmpCheckVal() As String
    Dim i As Integer

    If m_strCheckFlds = "" Then
        CheckHeadSelecting = True
        Exit Function
    End If
    row = ctlmakevoucher.HeadList.row
    varTmpCheckFld = Split(m_strCheckFlds, ",")
    ReDim strTmpCheckVal(UBound(varTmpCheckFld)) As String

    For i = 0 To UBound(varTmpCheckFld)

        strFieldName = varTmpCheckFld(i)  '"cDepCode"
        strFiledValue = ctlmakevoucher.HeadList.TextMatrix(row, ctlmakevoucher.GetHeadColIndex(strFieldName))
        If SelectedStr = "" Then
            strTmpCheckVal(i) = strFiledValue
        Else
            varTmp = Split(SelectedStr, ",")

            If UBound(varTmp) = 1 Then
                strTmpCheckVal(i) = ctlmakevoucher.HeadList.TextMatrix(varTmp(0), ctlmakevoucher.GetHeadColIndex(strFieldName))
            End If
            If strFiledValue <> strTmpCheckVal(i) And strTmpCheckVal(i) <> "" Then
                CheckHeadSelecting = False
                MsgBox m_strCheckMsg, 64, GetString("U8.DZ.JA.Res030")    'zh-CN：选择了不同部门的单据
                Exit Function
            End If
        End If

    Next
    CheckHeadSelecting = True
End Function


Public Sub SetFilter()
    Dim objFilter As New UFGeneralFilter.FilterSrv
    Call objFilter.OpenFilterConfig(g_oLogin, "", strFilterKey, strFilterSubID, "")
    '    objfltint.DeleteFilter
    Set objFilter = Nothing
End Sub
