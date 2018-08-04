VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "InvokeApi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

' -----------------------------------------------------
'  销售单据导入导出类
'  数据来源: XML格式的销售订单,销售发货单，销售发票;
'  数据去向:  U8 销售订单,销售发货单，销售发票
'  外部接口: transact(ByVal sXML As String, login As U8Login.clsLogin)
'  参数说明:d
'  sXml (String)               '待处理的XML
'  login (U8Login.clsLogin)    'Login对象
'  作者: 郝勇
'  日期:  2002.5－2002.6
' -----------------------------------------------------
Dim eaiLogger As Object
Dim m_errDom As DOMDocument
Private m_Conn As New ADODB.Connection
Public SaveHandle As Object
'Dim clsVouchs As New USSAServer.clsVouchLoad
Dim clsSystem As Object
Private mLogin As New U8Login.clsLogin
Dim errMsg As String
Dim m_bRetail As Boolean    '门店标记
Dim m_Ids As String  'id集和
Dim gObjRef As Object
Dim domSN As DOMDocument
Option Explicit

'--------------------------------------------------------------------
' 功能：对外接口
' 参数说明:
' sXML As String  导入导出条件XML描述，同时作为导入导出结果返回
' login As U8Login.clsLogin
' EleLst As IXMLDOMNodeList    导入导出条件中XML节点
' ByRef sXML As String         同Transact中sXML定义
' login As U8Login.clsLogin    同Transact中sXML定义
' RSXMLRefile As String,       导入导出对照表文件名称
' ViewHead As String,          大小写转换主视图
' ViewBody As String,          大小写转换子视图
' SaveType As String,          保存类型，对导入而言
' BillType As String           单据类型，对导入而言
'--------------------------------------------------------------------
Public Function transact(ByVal sXML As String, login As U8Login.clsLogin) As String

    Dim str As String
    Dim RSXMLRefile As String
    Dim ViewHead As String
    Dim ViewBody As String
    Dim SaveType As String
    Dim BillType As String
    BillType = ""
    Dim TableName As String
    Dim SubTableName As String

    Dim dom As New DOMDocument
    Dim ele As IXMLDOMElement
    Dim EleLst As IXMLDOMNodeList
    Dim docid As String
    Dim proc As String
    Dim rs As New ADODB.Recordset
    Dim sproc As String
    Dim sRetXml As String
    Dim sRootTag As String

    Set mLogin = login
    Set g_oLogin = login
    gstrCardNumber = "HYJCGH001"
    Set m_errDom = New DOMDocument
    '  WriteLog 0, "进入销售EAI", "SA", "SalesEAI.dll", "SalesEAICls", "transact"
    If Not dom.loadXML(sXML) Then   '如果不能Load
        transact = ReturnXml("100", "解析XML时发生错误,请你检查XML参数格式")
        Exit Function
    End If

    Set ele = dom.documentElement
    If ele Is Nothing Then
        transact = ReturnXml("101", "ufinterface结点不存在")
        Exit Function
    End If

    sRootTag = GetAttrValue(ele, "roottag")
    docid = GetAttrValue(ele, "docid")
    proc = GetAttrValue(ele, "proc")
    m_errDom.loadXML (GetErrorTemplate(docid, proc))
    sproc = GetAttrValue(ele, "proc")   '取得操作类型
    '根据单据类型分别设置不同的导入导出参数
    Select Case UCase(sproc)
    Case "ADD", "EDIT"   '导入
        RSXMLRefile = "HY_JCGH001.Xml"
        ViewHead = "V_HY_DZ_BorrowOut"
        ViewBody = "V_HY_DZ_BorrowOuts"
        SaveType = "借出借用单"
        Set EleLst = ele.selectNodes("HY_JCGH001")



        transact = TransactGeneral(EleLst, sXML, login, RSXMLRefile, ViewHead, ViewBody, SaveType, BillType)
        Exit Function
    Case "DELETE"
        RSXMLRefile = "HY_JCGH001.Xml"
        ViewHead = "V_HY_DZ_BorrowOut"
        ViewBody = "V_HY_DZ_BorrowOuts"
        SaveType = "借出借用单"
        Set EleLst = ele.selectNodes("HY_JCGH001")
        transact = TransactGeneral(EleLst, sXML, login, RSXMLRefile, ViewHead, ViewBody, SaveType, BillType)
        Exit Function
        '        Case "QUERY"    '导出
        '
        '
        '            Select Case LCase(sRootTag)
        '                Case "saleorder"            '销售订单
        '                    RSXMLRefile = "SaleOrderXmlRs.Xml"
        '                    TableName = "so_somain"
        '                    SubTableName = "so_sodetails"
        '                    Set EleLst = ele.selectNodes("saleorder")
        '                Case "saleinvoice"          '销售发票
        '                    RSXMLRefile = "SaleInvoiceXmlRs.Xml"
        '                    TableName = "SaleBillVouchZT_All"
        '                    SubTableName = "SaleBillVouchZW"
        '                    Set EleLst = ele.selectNodes("saleinvoice")
        '                Case "consignment"          '销售发货单
        '                    RSXMLRefile = "consignmentXmlRs.Xml"
        '                    TableName = "dispatchlist"
        '                    SubTableName = "dispatchlists"
        '                    Set EleLst = ele.selectNodes("consignment")
        '                Case "sa_invpricejust"     '销售存货调价单
        '                    RSXMLRefile = "SA_InvPriceJustXmlRs.xml"
        '                    TableName = "SA_InvPriceJustMain"
        '                    SubTableName = "SA_InvPriceJustDetail"
        '                    Set EleLst = ele.selectNodes("sa_invpricejust")
        '                Case "sa_cuspricejust"     '销售客户调价单
        '                    RSXMLRefile = "SA_CusPriceJustXmlRs.xml"
        '                    TableName = "SA_CusPriceJustMain"
        '                    SubTableName = "SA_CusPriceJustDetail"
        '                    Set EleLst = ele.selectNodes("sa_cuspricejust")
        '                Case "salereport" '移动查询销售统计表
        '                    RSXMLRefile = DoQuerySaleReport(sXML)
        '                Case "currentstocksa" '移动查询销售现存量查询
        '                    RSXMLRefile = DoQuerySaleReport(sXML)
        '                Case "salereportcolligate" '移动查询销售综合统计表
        '                   transact = DoQuerySaleReportcolligate(sXML, GetAttrValue(ele, "sender"))
        '            End Select
        '             If LCase(sRootTag) <> "salereportcolligate" Then
        '                transact = TransactGeneral(EleLst, sXML, login, RSXMLRefile, TableName, SubTableName, " ", " ")
        '            End If

    Case Else
        transact = ReturnXml("710", "本模块暂不提供此功能")
        Exit Function
    End Select

End Function

Private Sub SplitVouch(ele As IXMLDOMElement)
    Dim elePay As IXMLDOMElement

    Set elePay = ele.selectSingleNode("serialno")
    If elePay Is Nothing Then Exit Sub

    Set domSN = New DOMDocument
    domSN.loadXML "<?xml version='1.0' encoding='gb2312'?> " + Chr(13) + elePay.xml

    Set elePay = Nothing


End Sub
'--------------------------------------------------------------------------
' 功能： 导入导出实现方法
' 参数说明：
' EleLst As IXMLDOMNodeList    导入导出条件中XML节点
' ByRef sXML As String         同Transact中sXML定义
' login As U8Login.clsLogin    同Transact中sXML定义
' RSXMLRefile As String,       导入导出对照表文件名称
' ViewHead As String,          大小写转换主视图
' ViewBody As String,          大小写转换子视图
' SaveType As String,          保存类型，对导入而言
' BillType As String           单据类型，对导入而言
'--------------------------------------------------------------------------
Public Function TransactGeneral(EleLst As IXMLDOMNodeList, ByRef sXML As String, login As U8Login.clsLogin, RSXMLRefile As String, ViewHead As String, ViewBody As String, SaveType As String, BillType As String) As String
    On Error GoTo Err_log
    Dim dom As New DOMDocument
    Dim ele As IXMLDOMElement
    Dim docid As String
    Dim proc As String
    Dim rs As New ADODB.Recordset
    Dim sproc As String
    Dim sRetXml As String
    Dim sRootTag As String
    Dim del As IXMLDOMElement
    Dim eleHead As IXMLDOMElement
    Dim sBillCode As String


    Set mLogin = login
    Set m_errDom = New DOMDocument
    ' WriteLog 0, "导入导出实现方法", "SA", "SalesEAI.dll", "SalesEAICls", "TransactGeneral"
    If Not m_bRetail Then    '门店处理外部启动连接
        m_Conn.ConnectionTimeout = 0
        m_Conn.Open login.UfDbName
        If m_Conn Is Nothing Then TransactGeneral = ReturnXml("529", "数据库连接错误!"): Exit Function
    End If

    If Not dom.loadXML(sXML) Then   '如果不能Load
        TransactGeneral = ReturnXml("100", "解析XML时发生错误,请你检查XML参数格式")
        Exit Function
    End If

    Set ele = dom.documentElement
    If ele Is Nothing Then
        TransactGeneral = ReturnXml("101", "ufinterface结点不存在")
        Exit Function
    End If

    sRootTag = GetAttrValue(ele, "roottag")
    docid = GetAttrValue(ele, "docid")
    proc = GetAttrValue(ele, "proc")
    m_errDom.loadXML (GetErrorTemplate(docid, proc))
    sproc = GetAttrValue(ele, "proc")   '取得操作类型

    Select Case UCase(sproc)
    Case "ADD"
        Call AddProc(EleLst, sRetXml, RSXMLRefile, ViewHead, ViewBody, SaveType, sRootTag)
    Case "EDIT"
        Call EditProc(EleLst, sRetXml, RSXMLRefile, ViewHead, ViewBody, SaveType, sRootTag)
    Case "QUERY"
        sRetXml = dom.xml
        Dim StrType As String
        If Not ele.selectSingleNode(sRootTag) Is Nothing Then
            If Not ele.selectSingleNode(sRootTag).Attributes.getNamedItem("type") Is Nothing Then
                StrType = ele.selectSingleNode(sRootTag).Attributes.getNamedItem("type").nodeValue
            End If
            If StrType = "custom" Then
                Set EleLst = ele.selectNodes(sRootTag & "/sql")
            Else
                Set EleLst = ele.selectNodes(sRootTag & "/field")
            End If
        End If
        Call QueryProc(EleLst, sRetXml, RSXMLRefile, ViewHead, ViewBody)
    Case "DELETE"
        For Each del In EleLst
            Set eleHead = del.selectSingleNode("header")
            sBillCode = GetElementValue(eleHead, "ccode")
            'enum by modify
            If GetElementValue(eleHead, "ccreatetype") = "转换单据" Then
                ReDim varArgs(0)
                varArgs(0) = sBillCode
                TransactGeneral = GetStringPara("U8.DZ.JA.Res150", varArgs(0))
'                TransactGeneral = "该单据[" + sBillCode + " ]由借用转换单生成，不能删除！"
                Exit Function
            End If
        Next
        For Each del In EleLst
            Set eleHead = del.selectSingleNode("header")
            sBillCode = GetElementValue(eleHead, "ccode")
            m_Conn.Execute "delete from HY_DZ_BorrowOuts where ID=(select ID from HY_DZ_BorrowOut where ccode='" & sBillCode & "')"
            '删除字表
            g_Conn.Execute "delete from HY_DZ_BorrowOut where ccode='" & sBillCode & "'"


        Next

    Case Else
        TransactGeneral = ReturnXml("710", "本模块暂不提供此功能")
        Exit Function
    End Select

    If m_bRetail = False Then
        Set m_Conn = Nothing
    End If
    Set dom = Nothing
    TransactGeneral = sRetXml
    Exit Function
Err_log:
    TransactGeneral = ReturnXml(CStr(Err.Number), Err.Description)
    Err.Clear

End Function

'----------------------------------------------------------------
' 功能： 添加EAI单据(单张或多张单据)
' 参数说明：
' 同名参数定义同TransactGeneral
'----------------------------------------------------------------

Private Function AddProc(ByVal EleLst As IXMLDOMNodeList, ByRef sRet As String, RSXMLRefile As String, ViewHead As String, ViewBody As String, SaveType As String, BillType As String) As Boolean
    On Error GoTo Err_log

    Dim TmpEle As IXMLDOMElement
    Dim bSuccee As Boolean
    WriteLog 0, "添加EAI单据", "SA", "SalesEAI.dll", "SalesEAICls", "AddProc"
    'TmpEle为一张EAI单据的内容
    For Each TmpEle In EleLst

        Call AddSingleSO(TmpEle, RSXMLRefile, ViewHead, ViewBody, SaveType, BillType)
    Next

    sRet = m_errDom.xml
    AddProc = True
    Exit Function
Err_log:
    Call ErrDomAppendChild(m_errDom, "Item", 999, Err.Description)
    sRet = m_errDom.xml

End Function
'----------------------------------------------------------------
' 功能： 添加EAI单据(单张或多张单据)
' 参数说明：
' 同名参数定义同TransactGeneral
'----------------------------------------------------------------

Private Function EditProc(ByVal EleLst As IXMLDOMNodeList, ByRef sRet As String, RSXMLRefile As String, ViewHead As String, ViewBody As String, SaveType As String, BillType As String) As Boolean
    On Error GoTo Err_log

    Dim TmpEle As IXMLDOMElement
    Dim bSuccee As Boolean
    WriteLog 0, "添加EAI单据", "SA", "SalesEAI.dll", "SalesEAICls", "AddProc"
    'TmpEle为一张EAI单据的内容
    For Each TmpEle In EleLst

        Call EditSingleSO(TmpEle, RSXMLRefile, ViewHead, ViewBody, SaveType, BillType)
    Next

    sRet = m_errDom.xml
    EditProc = True
    Exit Function
Err_log:
    Call ErrDomAppendChild(m_errDom, "Item", 999, Err.Description)
    sRet = m_errDom.xml

End Function

'----------------------------------------------------------------------
' 功能： 保存单张单据
' 参数说明：
' 同名参数定义同TransactGeneral
'------------------------------------------------------------------------
Private Function AddSingleSO(ByVal ele As IXMLDOMElement, RSXMLRefile As String, ViewHead As String, ViewBody As String, SaveType As String, BillType As String) As Boolean
    On Error GoTo errorhandle

    Dim rsHeader As New ADODB.Recordset
    Dim rsBody As New ADODB.Recordset
    Dim eleHead As IXMLDOMElement                      '表头节点
    Dim eleBody As IXMLDOMElement                '表体节点
    Dim entryLst As IXMLDOMNodeList              '分录集
    Dim sBillCode As String                      '当前单据号
    Dim vouchtype As VoucherCO_Sa.VoucherTypeSA  '单据类型
    Dim sVoucherType As String
    Dim sSource As String
    Dim sMode As String, cMaker As String, cVerify As String
    Dim elelocal As IXMLDOMElement
    Dim scdepcode As String
    Dim scfreightcompany As String
    Dim strSql As String
    Dim sddate As String
    Dim DMO As CDMO
    Dim nNewID As Long

    On Error GoTo DoErr
    '    WriteLog 0, "保存单张单据", "SA", "SalesEAI.dll", "SalesEAICls", "AddSingleSO"
    sMode = "bs"
    Set domSN = Nothing
    SplitVouch ele

    Set eleHead = ele.selectSingleNode("header")

    '获取单据编号
    '    Select Case UCase(ViewHead)
    '        Case "SALEBILLVOUCHZT"
    '            sBillCode = GetElementValue(eleHead, "invoiceno")
    '            sCusCode = GetElementValue(eleHead, "customercode")
    '        Case "SA_INVPRICEJUSTMAINVIEW", "SA_CUSPRICEJUSTMAINVIEW"
    '            sBillCode = GetElementValue(eleHead, "ccode")
    '        Case Else
    sBillCode = GetElementValue(eleHead, "ccode")
    scdepcode = GetElementValue(eleHead, "cdepcode")
    scfreightcompany = GetElementValue(eleHead, "cfreightcompany")

    '    End Select

    '870 设置零售eai默认值
    Dim bRetailEAI As Boolean
    '    bRetailEAI = CusBshop(sCusCode)
    '    If bRetailEAI = True Or m_bRetail Then
    '        Call setDefaultValue(ele, ViewHead)
    '    End If

    Set eleBody = ele.selectSingleNode("body")
    Set entryLst = eleBody.selectNodes("entry")

    rsHeader.CursorLocation = adUseClient
    rsBody.CursorLocation = adUseClient

    rsHeader.Open "Select * From " & ViewHead & "  Where 1=2", m_Conn, adOpenStatic, adLockOptimistic
    rsBody.Open "Select * From " & ViewBody & "  Where 1=2", m_Conn, adOpenStatic, adLockOptimistic

    Set rsHeader.ActiveConnection = Nothing
    Set rsBody.ActiveConnection = Nothing

    Dim DomHeader As New DOMDocument
    Dim domBody As New DOMDocument
    Dim EXmlToRs As New EAIXMLRSExch20.clsXmlRsExch

    '将RescordSet分别转换为对应的含有小写属性字符串的XML文件DomHeader，DomBody中
    rsHeader.Save DomHeader, 1
    rsBody.Save domBody, 1


    '将数据源XML文件中的数据分别写入上面得到的DomHeader，DomBody中
    Call EXmlToRs.XmlToDomHeader(ele.xml, BillType, DomHeader, rsHeader, RSXMLRefile)
    '    MsgBox ele.xml
    Call EXmlToRs.XmlToDomBody(ele.xml, BillType, domBody, rsBody, RSXMLRefile)

    '对于每一个Entry增加属性editprop，并且设置默认值"A"
    '因为该属性不存在于模板和对照表文件中，却为销售单据保存方法所必须


    '调用销售后台保存方法


    Dim ReturnFlag As String
    Dim FirstFlag As String
    Dim BusinessType As String
    Set elelocal = DomHeader.selectSingleNode("//z:row")
    elelocal.setAttribute "id", "0"
    elelocal.setAttribute "istatus", "新建"
    elelocal.setAttribute "ccreatetype", "新增单据"
    'elelocal.setAttribute "ctype", "客户"
    elelocal.setAttribute "vouchertype", "HY_JCGH001"
    elelocal.setAttribute "iswfcontrolled", "1"


    '清除表头中的不可导入的字段 2008-07-24

    If DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("ctype") Is Nothing Then
        DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("ctype").nodeValue = "客户"
    End If
    If Not DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("chandler") Is Nothing Then
        DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("chandler").nodeValue = ""
    End If

    If Not DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("dveridate") Is Nothing Then
        DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("dveridate").nodeValue = ""
    End If
    If Not DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("closeuser") Is Nothing Then
        DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("closeuser").nodeValue = ""
    End If
    If Not DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("dclosedate") Is Nothing Then
        DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("dclosedate").nodeValue = ""
    End If
    If Not DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("intouser") Is Nothing Then
        DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("intouser").nodeValue = ""
    End If
    If Not DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("dintodate") Is Nothing Then
        DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("dintodate").nodeValue = ""
    End If

    '检查
    Dim rs As New ADODB.Recordset
    '表头
    '单位


    If Not DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("bobjectcode") Is Nothing Then
        If DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("ctype").nodeValue = "客户" Then
            strSql = "select * from Customer where cCusCode='" & DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("bobjectcode").nodeValue & "'"
            rs.Open strSql, m_Conn, adOpenForwardOnly, adLockReadOnly
            If rs.EOF Then
                errMsg = "对应客户的单位[" + DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("bobjectcode").nodeValue + "]是错误的!"
                AddSingleSO = False
                Exit Function
            End If
        ElseIf DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("ctype").nodeValue = "供应商" Then
            strSql = "select * from Vendor where cVenCode='" & DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("bobjectcode").nodeValue & "'"
            rs.Open strSql, m_Conn, adOpenForwardOnly, adLockReadOnly
            If rs.EOF Then
                errMsg = "对应供应商的单位[" + DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("bobjectcode").nodeValue + "]是错误的!"
                AddSingleSO = False
                Exit Function
            End If
        ElseIf DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("ctype").nodeValue = "部门" Then
            strSql = "select * from Department where cDepCode='" & DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("bobjectcode").nodeValue & "'"
            rs.Open strSql, m_Conn, adOpenForwardOnly, adLockReadOnly
            If rs.EOF Then
                errMsg = "对应部门的单位[" + DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("bobjectcode").nodeValue + "]是错误的!"
                AddSingleSO = False
                Exit Function
            End If
        End If
    End If
    '经办人

    If Not DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("cpersoncode") Is Nothing Then
        strSql = "select * from Person where cpersoncode='" & DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("cpersoncode").nodeValue & "'"
        rs.Open strSql, m_Conn, adOpenForwardOnly, adLockReadOnly
        If rs.EOF Then
            errMsg = "对应经办人[" + DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("cpersoncode").nodeValue + "]是错误的!"
            AddSingleSO = False
            Exit Function
        End If
    End If
    '部门
    If Not DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("cdepcode") Is Nothing Then
        strSql = "select * from Department where cDepCode='" & DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("cdepcode").nodeValue & "'"
        rs.Open strSql, m_Conn, adOpenForwardOnly, adLockReadOnly
        If rs.EOF Then
            errMsg = "对应部门[" + DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("cdepcode").nodeValue + "]是错误的!"
            AddSingleSO = False
            Exit Function
        End If
    End If
    '发送单位
    If Not DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("cfreightcompany") Is Nothing Then
        strSql = "select * from Vendor where cVenCode='" & DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("cfreightcompany").nodeValue & "'"
        rs.Open strSql, m_Conn, adOpenForwardOnly, adLockReadOnly
        If rs.EOF Then
            errMsg = "对应发送单位[" + DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("cfreightcompany").nodeValue + "]是错误的!"
            AddSingleSO = False
            Exit Function
        End If
    End If
    If Not DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("cfreight") Is Nothing Then
        If DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("cfreight").nodeValue = "是" Then
            If DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("mycdefinet2") Is Nothing Or DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("cfreighttype") Is Nothing Or DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("cfreightcost") Is Nothing Then
                errMsg = GetString("U8.DZ.JA.Res1150")
                AddSingleSO = False
                Exit Function
            End If
        End If
    End If
    If DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("cmaker") Is Nothing Then
        errMsg = GetString("U8.DZ.JA.Res1130")
        AddSingleSO = False
        Exit Function
    End If
    If DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("ddate") Is Nothing Then
        errMsg = "借用日期不能为空"
        AddSingleSO = False
        Exit Function
    Else
        sddate = DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("ddate").nodeValue
    End If

    '表体
    Dim nd As IXMLDOMElement
    Dim iRow As Integer
    iRow = 0
    For Each nd In domBody.selectNodes("//z:row")

        iRow = iRow + 1
        sAutoId = sAutoId + iRow - 1
        nd.setAttribute "autoid", sAutoId
        nd.setAttribute "id", sID
        '            BodyData.Item(i).Item("AutoId").Value = sAutoId
        '            BodyData.Item(i).Item("ID").Value = sID
        '1
        '比填项,简单校验：存货编码，数量
        '其他的校验根据实际业务而定
        If nd.getAttribute("cinvcode") = "" Then
            errMsg = "第" & iRow & "行存货编码不能为空"
            AddSingleSO = False
            Exit Function
        End If

        If Val(nd.getAttribute("iquantity") & "") = 0 And Val(nd.getAttribute("inum") & "") = 0 Then
            '            MsgBox "第" & iRow & "行 数量件数不能同时为空或等于0", vbInformation, getstring("U8.DZ.JA.Res030")
            errMsg = "第" & iRow & "行 数量/件数不能同时为空或等于0"
            AddSingleSO = False
            Exit Function
        End If

        'dxb 仓库必须录入！
        If nd.getAttribute("cwhcode") = "" Then
            errMsg = "第" & iRow & "行仓库不能为空"
            AddSingleSO = False
            Exit Function
        End If

        '2
        '启用的自由项是否都输入了
        '结构性，非结构性自由项都必须输入
        '自由项组合是否合法
        If ExecFunFreeCheck(nd.getAttribute("cinvcode"), nd, iRow) = False Then
            '标识当前行,着蓝色

            AddSingleSO = False
            Exit Function
        End If


        '３
        '批次
        '出库跟踪入库
        If ExecFuncbatch(nd.getAttribute("cinvcode"), _
                         nd.getAttribute("cbatch") & "", _
                         nd.getAttribute("cinvouchcode") & "", _
                         iRow) = False Then
            '标识当前行,着蓝色

            AddSingleSO = False
            Exit Function
        End If



        '5 预计归还日期 dxb
        If nd.getAttribute("backdate") = "" Then
            errMsg = "第" & iRow & "行预计归还日期不能为空"
            AddSingleSO = False
            Exit Function
        Else
            If CDate(Mid(nd.getAttribute("backdate"), 1, 10)) < CDate(sddate) Then
                errMsg = "第" & iRow & "行预计归还日期不能早于借用日期"
                AddSingleSO = False
                Exit Function
            End If
        End If

    Next

    '发送方式
    If Not DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("cfreighttype") Is Nothing Then
        strSql = "select * from ShippingChoice where csccode='" & DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("cfreighttype").nodeValue & "'"
        rs.Open strSql, m_Conn, adOpenForwardOnly, adLockReadOnly
        If rs.EOF Then
            errMsg = "对应发送方式[" + DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("cfreighttype").nodeValue + "]是错误的!"
            AddSingleSO = False
            Exit Function
        End If
    End If
    m_Conn.Execute "update ufsystem..ua_identity set ichildid=(select isnull(max(autoid),1) from HY_DZ_BorrowOuts ) where cvouchtype='HY_JCGH001'"
    Call GetMaxID
    DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("id").nodeValue = sID
    DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("dmdate").nodeValue = mLogin.CurDate
    DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("istatus").nodeValue = 1

    '保存

    '表头
    Dim str1 As String
    Dim str2 As String
    str1 = " insert into HY_DZ_BorrowOut (ID,"
    str2 = " values ( " & sID & ","
    If Not DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("ccreatetype") Is Nothing Then
        str1 = str1 + "cCreateType,"
        str2 = str2 + " '" & DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("ccreatetype").nodeValue & "',"
    End If
    If Not DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("ccode") Is Nothing Then
        str1 = str1 + "ccode,"
        str2 = str2 + " '" & DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("ccode").nodeValue & "',"
    End If
    If Not DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("ctype") Is Nothing Then
        str1 = str1 + "ctype,"
        str2 = str2 + " '" & DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("ctype").nodeValue & "',"
    End If
    If Not DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("bobjectcode") Is Nothing Then
        str1 = str1 + "bobjectcode,"
        str2 = str2 + " '" & DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("bobjectcode").nodeValue & "',"
    End If
    If Not DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("cpersoncode") Is Nothing Then
        str1 = str1 + "cpersoncode,"
        str2 = str2 + " '" & DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("cpersoncode").nodeValue & "',"
    End If
    If Not DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("ccontactperson") Is Nothing Then
        str1 = str1 + "ccontactperson,"
        str2 = str2 + " '" & DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("ccontactperson").nodeValue & "',"
    End If
    If Not DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("ccontactway") Is Nothing Then
        str1 = str1 + "ccontactway,"
        str2 = str2 + " '" & DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("ccontactway").nodeValue & "',"
    End If
    If Not DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("cdepcode") Is Nothing Then
        str1 = str1 + "cdepcode,"
        str2 = str2 + " '" & DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("cdepcode").nodeValue & "',"
    End If
    If Not DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("ddate") Is Nothing Then
        str1 = str1 + "ddate,"
        str2 = str2 + " '" & DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("ddate").nodeValue & "',"
    End If
    If Not DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("cmemo") Is Nothing Then
        str1 = str1 + "cmemo,"
        str2 = str2 + " '" & DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("cmemo").nodeValue & "',"
    End If
    If Not DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("cfreight") Is Nothing Then
        str1 = str1 + "cfreight,"
        str2 = str2 + " '" & DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("cfreight").nodeValue & "',"
    End If
    If Not DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("cfreighttype") Is Nothing Then
        str1 = str1 + "cfreighttype,"
        str2 = str2 + " '" & DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("cfreighttype").nodeValue & "',"
    End If
    If Not DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("cfreightcompany") Is Nothing Then
        str1 = str1 + "cfreightcompany,"
        str2 = str2 + " '" & DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("cfreightcompany").nodeValue & "',"
    End If
    If Not DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("cfreightcost") Is Nothing Then
        str1 = str1 + "cfreightcost,"
        str2 = str2 + " '" & DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("cfreightcost").nodeValue & "',"
    End If
    If Not DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("cabouvoucher") Is Nothing Then
        str1 = str1 + "cabouvoucher,"
        str2 = str2 + " '" & DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("cabouvoucher").nodeValue & "',"
    End If
    If Not DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("ccodeaboutvoucher") Is Nothing Then
        str1 = str1 + "ccodeaboutvoucher,"
        str2 = str2 + " '" & DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("ccodeaboutvoucher").nodeValue & "',"
    End If
    If Not DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("cmaker") Is Nothing Then
        str1 = str1 + "cmaker,"
        str2 = str2 + " '" & DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("cmaker").nodeValue & "',"
    End If
    If Not DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("dmdate") Is Nothing Then
        str1 = str1 + "dmdate,"
        str2 = str2 + " '" & DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("dmdate").nodeValue & "',"
    End If
    If Not DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("chandler") Is Nothing Then
        str1 = str1 + "chandler,"
        str2 = str2 + " '" & DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("chandler").nodeValue & "',"
    End If
    If Not DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("dveridate") Is Nothing Then
        str1 = str1 + "dveridate,"
        str2 = str2 + " '" & DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("dveridate").nodeValue & "',"
    End If
    If Not DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("istatus") Is Nothing Then
        str1 = str1 + "istatus,"
        str2 = str2 + " '" & DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("istatus").nodeValue & "',"
    End If
    If Not DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("vouchertype") Is Nothing Then
        str1 = str1 + "vouchertype,"
        str2 = str2 + " '" & DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("vouchertype").nodeValue & "',"
    End If
    If Not DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("iverifystate") Is Nothing Then
        str1 = str1 + "iverifystate,"
        str2 = str2 + " '" & DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("iverifystate").nodeValue & "',"
    End If
    If Not DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("iswfcontrolled") Is Nothing Then
        str1 = str1 + "iswfcontrolled,"
        str2 = str2 + " '" & DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("iswfcontrolled").nodeValue & "',"
    End If
    If Not DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("ireturncount") Is Nothing Then
        str1 = str1 + "ireturncount,"
        str2 = str2 + " '" & DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("ireturncount").nodeValue & "',"
    End If
    If Not DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("cdefine1") Is Nothing Then
        str1 = str1 + "cdefine1,"
        str2 = str2 + " '" & DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("cdefine1").nodeValue & "',"
    End If
    If Not DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("cdefine2") Is Nothing Then
        str1 = str1 + "cdefine2,"
        str2 = str2 + " '" & DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("cdefine2").nodeValue & "',"
    End If
    If Not DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("cdefine3") Is Nothing Then
        str1 = str1 + "cdefine3,"
        str2 = str2 + " '" & DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("cdefine3").nodeValue & "',"
    End If
    If Not DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("cdefine4") Is Nothing Then
        str1 = str1 + "cdefine4,"
        str2 = str2 + " '" & DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("cdefine4").nodeValue & "',"
    End If
    If Not DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("cdefine5") Is Nothing Then
        str1 = str1 + "cdefine5,"
        str2 = str2 + " '" & DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("cdefine5").nodeValue & "',"
    End If
    If Not DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("cdefine6") Is Nothing Then
        str1 = str1 + "cdefine6,"
        str2 = str2 + " '" & DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("cdefine6").nodeValue & "',"
    End If
    If Not DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("cdefine7") Is Nothing Then
        str1 = str1 + "cdefine7,"
        str2 = str2 + " '" & DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("cdefine7").nodeValue & "',"
    End If
    If Not DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("cdefine8") Is Nothing Then
        str1 = str1 + "cdefine8,"
        str2 = str2 + " '" & DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("cdefine8").nodeValue & "',"
    End If
    If Not DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("cdefine9") Is Nothing Then
        str1 = str1 + "cdefine9,"
        str2 = str2 + " '" & DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("cdefine9").nodeValue & "',"
    End If
    If Not DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("cdefine10") Is Nothing Then
        str1 = str1 + "cdefine10,"
        str2 = str2 + " '" & DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("cdefine10").nodeValue & "',"
    End If
    If Not DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("cdefine11") Is Nothing Then
        str1 = str1 + "cdefine11,"
        str2 = str2 + " '" & DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("cdefine11").nodeValue & "',"
    End If
    If Not DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("cdefine12") Is Nothing Then
        str1 = str1 + "cdefine12,"
        str2 = str2 + " '" & DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("cdefine12").nodeValue & "',"
    End If
    If Not DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("cdefine13") Is Nothing Then
        str1 = str1 + "cdefine13,"
        str2 = str2 + " '" & DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("cdefine13").nodeValue & "',"
    End If
    If Not DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("cdefine14") Is Nothing Then
        str1 = str1 + "cdefine14,"
        str2 = str2 + " '" & DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("cdefine14").nodeValue & "',"
    End If
    If Not DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("cdefine15") Is Nothing Then
        str1 = str1 + "cdefine15,"
        str2 = str2 + " '" & DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("cdefine15").nodeValue & "',"
    End If
    If Not DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("cdefine16") Is Nothing Then
        str1 = str1 + "cdefine16,"
        str2 = str2 + " '" & DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("cdefine16").nodeValue & "',"
    End If
    If Not DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("mycdefinet1") Is Nothing Then
        str1 = str1 + "mycdefinet1,"
        str2 = str2 + " '" & DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("mycdefinet1").nodeValue & "',"
    End If
    If Not DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("mycdefinet2") Is Nothing Then
        str1 = str1 + "mycdefinet2,"
        str2 = str2 + " '" & DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("mycdefinet2").nodeValue & "',"
    End If
    If Not DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("mycdefinet3") Is Nothing Then
        str1 = str1 + "mycdefinet3,"
        str2 = str2 + " '" & DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("mycdefinet3").nodeValue & "',"
    End If
    If Not DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("mycdefinet4") Is Nothing Then
        str1 = str1 + "mycdefinet4,"
        str2 = str2 + " '" & DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("mycdefinet4").nodeValue & "',"
    End If
    If Not DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("mycdefinet5") Is Nothing Then
        str1 = str1 + "mycdefinet5,"
        str2 = str2 + " '" & DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("mycdefinet5").nodeValue & "',"
    End If
    If Not DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("mycdefinet6") Is Nothing Then
        str1 = str1 + "mycdefinet6,"
        str2 = str2 + " '" & DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("mycdefinet6").nodeValue & "',"
    End If
    If Not DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("mycdefinet7") Is Nothing Then
        str1 = str1 + "mycdefinet7,"
        str2 = str2 + " '" & DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("mycdefinet7").nodeValue & "',"
    End If
    If Not DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("mycdefinet8") Is Nothing Then
        str1 = str1 + "mycdefinet8,"
        str2 = str2 + " '" & DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("mycdefinet8").nodeValue & "',"
    End If
    If Not DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("mycdefinet9") Is Nothing Then
        str1 = str1 + "mycdefinet9,"
        str2 = str2 + " '" & DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("mycdefinet9").nodeValue & "',"
    End If
    If Not DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("mycdefinet10") Is Nothing Then
        str1 = str1 + "mycdefinet10,"
        str2 = str2 + " '" & DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("mycdefinet10").nodeValue & "',"
    End If
    If Not DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("downstreamcode") Is Nothing Then
        str1 = str1 + "downstreamcode,"
        str2 = str2 + " '" & DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("downstreamcode").nodeValue & "',"
    End If
    If Not DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("upstreamcode") Is Nothing Then
        str1 = str1 + "upstreamcode,"
        str2 = str2 + " '" & DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("upstreamcode").nodeValue & "',"
    End If
    str1 = str1 + "closeuser )  " + str2 + " null )"
    m_Conn.Execute str1
    '表体
    iRow = 0
    Dim str As String
    For Each nd In domBody.selectNodes("//z:row")

        iRow = iRow + 1

        str1 = " insert into HY_DZ_BorrowOuts(ID,AutoID,"
        str2 = " values(" & sID & "," & nd.getAttribute("autoid") & ","
        If nd.getAttribute("backdate") <> "" Then
            str1 = str1 + " backdate, "
            str2 = str2 + "'" & nd.getAttribute("backdate") & "',"

        End If

        If nd.getAttribute("backdate") = "" Then
            str1 = str1 + " backdate, "
            str2 = str2 + "'" & nd.getAttribute("backdate") & "',"
        End If
        If nd.getAttribute("cbatch") = "" Then
            str1 = str1 + " cbatch, "
            str2 = str2 + "'" & nd.getAttribute("cbatch") & "',"
        End If
        If nd.getAttribute("cdefine22") = "" Then
            str1 = str1 + " cdefine22, "
            str2 = str2 + "'" & nd.getAttribute("cdefine22") & "',"
        End If
        If nd.getAttribute("cdefine23") = "" Then
            str1 = str1 + " cdefine23, "
            str2 = str2 + "'" & nd.getAttribute("cdefine23") & "',"
        End If
        If nd.getAttribute("cdefine24") = "" Then
            str1 = str1 + " cdefine24, "
            str2 = str2 + "'" & nd.getAttribute("cdefine24") & "',"
        End If
        If nd.getAttribute("cdefine25") = "" Then
            str1 = str1 + " cdefine25, "
            str2 = str2 + "'" & nd.getAttribute("cdefine25") & "',"
        End If
        If nd.getAttribute("cdefine26") = "" Then
            str1 = str1 + " cdefine26, "
            str2 = str2 + "'" & nd.getAttribute("cdefine26") & "',"
        End If
        If nd.getAttribute("cdefine27") = "" Then
            str1 = str1 + " cdefine27, "
            str2 = str2 + "'" & nd.getAttribute("cdefine27") & "',"
        End If
        If nd.getAttribute("cdefine28") = "" Then
            str1 = str1 + " cdefine28, "
            str2 = str2 + "'" & nd.getAttribute("cdefine28") & "',"
        End If
        If nd.getAttribute("cdefine29") = "" Then
            str1 = str1 + " cdefine29, "
            str2 = str2 + "'" & nd.getAttribute("cdefine29") & "',"
        End If
        If nd.getAttribute("cdefine30") = "" Then
            str1 = str1 + " cdefine30, "
            str2 = str2 + "'" & nd.getAttribute("cdefine30") & "',"
        End If
        If nd.getAttribute("cdefine31") = "" Then
            str1 = str1 + " cdefine31, "
            str2 = str2 + "'" & nd.getAttribute("cdefine31") & "',"
        End If
        If nd.getAttribute("cdefine32") = "" Then
            str1 = str1 + " cdefine32, "
            str2 = str2 + "'" & nd.getAttribute("cdefine32") & "',"
        End If
        If nd.getAttribute("cdefine33") = "" Then
            str1 = str1 + " cdefine33, "
            str2 = str2 + "'" & nd.getAttribute("cdefine33") & "',"
        End If
        If nd.getAttribute("cdefine34") = "" Then
            str1 = str1 + " cdefine34, "
            str2 = str2 + "'" & nd.getAttribute("cdefine34") & "',"
        End If
        If nd.getAttribute("cdefine35") = "" Then
            str1 = str1 + " cdefine35, "
            str2 = str2 + "'" & nd.getAttribute("cdefine35") & "',"
        End If
        If nd.getAttribute("cdefine36") = "" Then
            str1 = str1 + " cdefine36, "
            str2 = str2 + "'" & nd.getAttribute("cdefine36") & "',"
        End If
        If nd.getAttribute("cdefine37") = "" Then
            str1 = str1 + " cdefine37, "
            str2 = str2 + "'" & nd.getAttribute("cdefine37") & "',"
        End If
        If nd.getAttribute("cfree1") = "" Then
            str1 = str1 + " cfree1, "
            str2 = str2 + "'" & nd.getAttribute("cfree1") & "',"
        End If
        If nd.getAttribute("cfree10") = "" Then
            str1 = str1 + " cfree10, "
            str2 = str2 + "'" & nd.getAttribute("cfree10") & "',"
        End If
        If nd.getAttribute("cfree2") = "" Then
            str1 = str1 + " cfree2, "
            str2 = str2 + "'" & nd.getAttribute("cfree2") & "',"
        End If
        If nd.getAttribute("cfree3") = "" Then
            str1 = str1 + " cfree3, "
            str2 = str2 + "'" & nd.getAttribute("cfree3") & "',"
        End If
        If nd.getAttribute("cfree4") = "" Then
            str1 = str1 + " cfree4, "
            str2 = str2 + "'" & nd.getAttribute("cfree4") & "',"
        End If
        If nd.getAttribute("cfree5") = "" Then
            str1 = str1 + " cfree5, "
            str2 = str2 + "'" & nd.getAttribute("cfree5") & "',"
        End If
        If nd.getAttribute("cfree6") = "" Then
            str1 = str1 + " cfree6, "
            str2 = str2 + "'" & nd.getAttribute("cfree6") & "',"
        End If
        If nd.getAttribute("cfree7") = "" Then
            str1 = str1 + " cfree7, "
            str2 = str2 + "'" & nd.getAttribute("cfree7") & "',"
        End If
        If nd.getAttribute("cfree8") = "" Then
            str1 = str1 + " cfree8, "
            str2 = str2 + "'" & nd.getAttribute("cfree8") & "',"
        End If
        If nd.getAttribute("cfree9") = "" Then
            str1 = str1 + " cfree9, "
            str2 = str2 + "'" & nd.getAttribute("cfree9") & "',"
        End If
        If nd.getAttribute("cinvcode") = "" Then
            str1 = str1 + " cinvcode, "
            str2 = str2 + "'" & nd.getAttribute("cinvcode") & "',"
        End If
        If nd.getAttribute("cmemo") = "" Then
            str1 = str1 + " cmemo, "
            str2 = str2 + "'" & nd.getAttribute("cmemo") & "',"
        End If
        If nd.getAttribute("cposition") = "" Then
            str1 = str1 + " cposition, "
            str2 = str2 + "'" & nd.getAttribute("cposition") & "',"
        End If
        If nd.getAttribute("cunitid") = "" Then
            str1 = str1 + " cunitid, "
            str2 = str2 + "'" & nd.getAttribute("cunitid") & "',"
        End If
        If nd.getAttribute("cwhcode") = "" Then
            str1 = str1 + " cwhcode, "
            str2 = str2 + "'" & nd.getAttribute("cwhcode") & "',"
        End If
        If nd.getAttribute("ID") = "" Then
            str1 = str1 + " ID, "
            str2 = str2 + "'" & nd.getAttribute("ID") & "',"
        End If
        If nd.getAttribute("iinvexchrate") = "" Then
            str1 = str1 + " iinvexchrate, "
            str2 = str2 + "'" & nd.getAttribute("iinvexchrate") & "',"
        End If
        If nd.getAttribute("inum") = "" Then
            str1 = str1 + " inum, "
            str2 = str2 + "'" & nd.getAttribute("inum") & "',"
        End If
        If nd.getAttribute("iquantity") = "" Then
            str1 = str1 + " iquantity, "
            str2 = str2 + "'" & nd.getAttribute("iquantity") & "',"
        End If
        If nd.getAttribute("MycdefineB1") = "" Then
            str1 = str1 + " MycdefineB1, "
            str2 = str2 + "'" & nd.getAttribute("MycdefineB1") & "',"
        End If
        If nd.getAttribute("MycdefineB10") = "" Then
            str1 = str1 + " MycdefineB10, "
            str2 = str2 + "'" & nd.getAttribute("MycdefineB10") & "',"
        End If
        If nd.getAttribute("MycdefineB2") = "" Then
            str1 = str1 + " MycdefineB2, "
            str2 = str2 + "'" & nd.getAttribute("MycdefineB2") & "',"
        End If
        If nd.getAttribute("MycdefineB3") = "" Then
            str1 = str1 + " MycdefineB3, "
            str2 = str2 + "'" & nd.getAttribute("MycdefineB3") & "',"
        End If
        If nd.getAttribute("MycdefineB4") = "" Then
            str1 = str1 + " MycdefineB4, "
            str2 = str2 + "'" & nd.getAttribute("MycdefineB4") & "',"
        End If
        If nd.getAttribute("MycdefineB5") = "" Then
            str1 = str1 + " MycdefineB5, "
            str2 = str2 + "'" & nd.getAttribute("MycdefineB5") & "',"
        End If
        If nd.getAttribute("MycdefineB6") = "" Then
            str1 = str1 + " MycdefineB6, "
            str2 = str2 + "'" & nd.getAttribute("MycdefineB6") & "',"
        End If
        If nd.getAttribute("MycdefineB7") = "" Then
            str1 = str1 + " MycdefineB7, "
            str2 = str2 + "'" & nd.getAttribute("MycdefineB7") & "',"
        End If
        If nd.getAttribute("MycdefineB8") = "" Then
            str1 = str1 + " MycdefineB8, "
            str2 = str2 + "'" & nd.getAttribute("MycdefineB8") & "',"
        End If
        If nd.getAttribute("MycdefineB9") = "" Then
            str1 = str1 + " MycdefineB9, "
            str2 = str2 + "'" & nd.getAttribute("MycdefineB9") & "',"
        End If
        str = str + vbCrLf + Left(str1, Len(str1) - 1) + ") " + Left(str2, Len(str2) - 1) + ")"

    Next
    m_Conn.Execute str1

    '结束

    ''    WriteLog 0, "开始保存单据", "SA", "SalesEAI.dll", "SalesEAICls", "AddSingleSO"
    '    If LCase(sMode) = "ar" Or (bRetailEAI And Not m_bRetail) Then m_Conn.BeginTrans
    '    errMsg = SaveHandle.Save(DomHeader, domBody, 0, nNewID)
    '    If LCase(sMode) = "ar" Then
    '        If Trim(errMsg) = "" Then
    '            m_Conn.CommitTrans
    '            WriteLog 0, "保存单据成功", "SA", "SalesEAI.dll", "SalesEAICls", "AddSingleSO"
    '        Else
    '            m_Conn.RollbackTrans
    '            WriteLog 1, "保存单据失败" + errMsg, "SA", "SalesEAI.dll", "SalesEAICls", "AddSingleSO"
    '        End If
    '    End If
    '
    '    If m_Ids = "" Then '获得保存后新的单据id值
    '        m_Ids = CStr(nNewID)
    '    Else
    '        m_Ids = m_Ids & "," & CStr(nNewID)
    '    End If

    '------------永川项目，自动审核发货单-----------------2005.12.15
    '销售EAI改进 , 在header节点增加属性bautoverify, 用来标志是否自动审核
    '    If errMsg = "" Then
    '        Dim sErrMsgVerify As String
    '        WriteLog 0, "开始审核单据", "SA", "SalesEAI.dll", "SalesEAICls", "AddSingleSO"
    '        If bRetailEAI Then
    '            eleHead.setAttribute "beairetailvouch", "1"
    '        End If
    '        Call AutoVerify(SaveHandle, ViewHead, ViewBody, nNewID, eleHead, sMode, sErrMsgVerify)
    '        If sErrMsgVerify = "" Then
    '            If LCase(sMode) <> "ar" And bRetailEAI And Not m_bRetail Then
    '                m_Conn.CommitTrans
    '            End If
    '            WriteLog 0, "审核单据成功", "SA", "SalesEAI.dll", "SalesEAICls", "AddSingleSO"
    '        Else
    '            If LCase(sMode) <> "ar" And bRetailEAI And Not m_bRetail Then
    '                m_Conn.RollbackTrans
    '            End If
    '            WriteLog 1, "审核单据失败" + sErrMsgVerify, "SA", "SalesEAI.dll", "SalesEAICls", "AddSingleSO"
    '        End If
    '        errMsg = errMsg & sErrMsgVerify
    '    Else
    '        If LCase(sMode) <> "ar" And bRetailEAI And Not m_bRetail Then
    '            m_Conn.RollbackTrans
    '        End If
    '    End If
    '    '---------------------------------------------------
    '
errorhandle:
    If (errMsg = "") And (Err.Description = "") Then
        Call ErrDomAppendChild(m_errDom, sBillCode, 0, "ok")
        Exit Function
    Else
        Call ErrDomAppendChild(m_errDom, sBillCode, 101, sBillCode & "导入失败" & errMsg & Err.Description)
        Exit Function
    End If
DoErr:
    Call ErrDomAppendChild(m_errDom, sBillCode, 101, "程序错误" & Err.Description)
    '    WriteLog 1, "异常错误:" + Err.Description, "SA", "SalesEAI.dll", "SalesEAICls", "AddSingleSO"
End Function

'----------------------------------------------------------------------
' 功能： 保存单张单据
' 参数说明：
' 同名参数定义同TransactGeneral
'------------------------------------------------------------------------
Private Function EditSingleSO(ByVal ele As IXMLDOMElement, RSXMLRefile As String, ViewHead As String, ViewBody As String, SaveType As String, BillType As String) As Boolean
    On Error GoTo errorhandle

    Dim rsHeader As New ADODB.Recordset
    Dim rsBody As New ADODB.Recordset
    Dim eleHead As IXMLDOMElement                      '表头节点
    Dim eleBody As IXMLDOMElement                '表体节点
    Dim entryLst As IXMLDOMNodeList              '分录集
    Dim sBillCode As String                      '当前单据号
    Dim vouchtype As VoucherCO_Sa.VoucherTypeSA  '单据类型
    Dim sVoucherType As String
    Dim sSource As String
    Dim sMode As String, cMaker As String, cVerify As String
    Dim elelocal As IXMLDOMElement
    Dim scdepcode As String
    Dim scfreightcompany As String
    Dim strSql As String
    Dim sddate As String
    Dim DMO As CDMO
    Dim nNewID As Long
    Dim ssdi As String

    On Error GoTo DoErr
    '    WriteLog 0, "保存单张单据", "SA", "SalesEAI.dll", "SalesEAICls", "AddSingleSO"
    sMode = "bs"
    Set domSN = Nothing
    SplitVouch ele

    Set eleHead = ele.selectSingleNode("header")

    sBillCode = GetElementValue(eleHead, "ccode")
    ssdi = GetElementValue(eleHead, "id")
    scdepcode = GetElementValue(eleHead, "cdepcode")
    scfreightcompany = GetElementValue(eleHead, "cfreightcompany")


    '870 设置零售eai默认值
    Dim bRetailEAI As Boolean

    Set eleBody = ele.selectSingleNode("body")
    Set entryLst = eleBody.selectNodes("entry")

    rsHeader.CursorLocation = adUseClient
    rsBody.CursorLocation = adUseClient

    rsHeader.Open "Select * From " & ViewHead & "  Where 1=2", m_Conn, adOpenStatic, adLockOptimistic
    rsBody.Open "Select * From " & ViewBody & "  Where 1=2", m_Conn, adOpenStatic, adLockOptimistic

    Set rsHeader.ActiveConnection = Nothing
    Set rsBody.ActiveConnection = Nothing

    Dim DomHeader As New DOMDocument
    Dim domBody As New DOMDocument
    Dim EXmlToRs As New EAIXMLRSExch20.clsXmlRsExch

    '将RescordSet分别转换为对应的含有小写属性字符串的XML文件DomHeader，DomBody中
    rsHeader.Save DomHeader, 1
    rsBody.Save domBody, 1


    '将数据源XML文件中的数据分别写入上面得到的DomHeader，DomBody中
    Call EXmlToRs.XmlToDomHeader(ele.xml, BillType, DomHeader, rsHeader, RSXMLRefile)
    '    MsgBox ele.xml
    Call EXmlToRs.XmlToDomBody(ele.xml, BillType, domBody, rsBody, RSXMLRefile)

    '对于每一个Entry增加属性editprop，并且设置默认值"A"
    '因为该属性不存在于模板和对照表文件中，却为销售单据保存方法所必须


    '调用销售后台保存方法


    Dim ReturnFlag As String
    Dim FirstFlag As String
    Dim BusinessType As String
    Set elelocal = DomHeader.selectSingleNode("//z:row")
    elelocal.setAttribute "vouchertype", "HY_JCGH001"



    '清除表头中的不可导入的字段 2008-07-24


    If Not DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("chandler") Is Nothing Then
        DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("chandler").nodeValue = ""
    End If

    If Not DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("dveridate") Is Nothing Then
        DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("dveridate").nodeValue = ""
    End If
    If Not DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("closeuser") Is Nothing Then
        DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("closeuser").nodeValue = ""
    End If
    If Not DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("dclosedate") Is Nothing Then
        DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("dclosedate").nodeValue = ""
    End If
    If Not DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("intouser") Is Nothing Then
        DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("intouser").nodeValue = ""
    End If
    If Not DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("dintodate") Is Nothing Then
        DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("dintodate").nodeValue = ""
    End If
    m_Conn.Execute "update ufsystem..ua_identity set ichildid=(select isnull(max(autoid),1) from HY_DZ_BorrowOuts ) where cvouchtype='HY_JCGH001'"
    Call GetMaxID


    '检查
    Dim rs As New ADODB.Recordset
    '表头
    '单位

    If Not DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("bobjectcode") Is Nothing Then
        If DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("ctype").nodeValue = "客户" Then
            strSql = "select * from Customer where cCusCode='" & DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("bobjectcode").nodeValue & "'"
            rs.Open strSql, m_Conn, adOpenForwardOnly, adLockReadOnly
            If rs.EOF Then
                errMsg = "对应客户的单位[" + DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("bobjectcode").nodeValue + "]是错误的!"
                EditSingleSO = False
                Exit Function
            End If
        ElseIf DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("ctype").nodeValue = "供应商" Then
            strSql = "select * from Vendor where cVenCode='" & DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("bobjectcode").nodeValue & "'"
            rs.Open strSql, m_Conn, adOpenForwardOnly, adLockReadOnly
            If rs.EOF Then
                errMsg = "对应供应商的单位[" + DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("bobjectcode").nodeValue + "]是错误的!"
                EditSingleSO = False
                Exit Function
            End If
        ElseIf DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("ctype").nodeValue = "部门" Then
            strSql = "select * from Department where cDepCode='" & DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("bobjectcode").nodeValue & "'"
            rs.Open strSql, m_Conn, adOpenForwardOnly, adLockReadOnly
            If rs.EOF Then
                errMsg = "对应部门的单位[" + DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("bobjectcode").nodeValue + "]是错误的!"
                EditSingleSO = False
                Exit Function
            End If
        End If
    Else
        errMsg = "没有输入单位类型!"
        EditSingleSO = False
        Exit Function
    End If
    If sBillCode = "" Then
        errMsg = "没有输入订单号!"
        EditSingleSO = False
        Exit Function
    End If
    If DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("id") Is Nothing Then
        errMsg = "没有输入订单ID!"
        EditSingleSO = False
        Exit Function
    End If
    '经办人

    If Not DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("cpersoncode") Is Nothing Then
        strSql = "select * from Person where cpersoncode='" & DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("cpersoncode").nodeValue & "'"
        rs.Open strSql, m_Conn, adOpenForwardOnly, adLockReadOnly
        If rs.EOF Then
            errMsg = "对应经办人[" + DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("cpersoncode").nodeValue + "]是错误的!"
            EditSingleSO = False
            Exit Function
        End If
    End If
    '部门
    If Not DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("cdepcode") Is Nothing Then
        strSql = "select * from Department where cDepCode='" & DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("cdepcode").nodeValue & "'"
        rs.Open strSql, m_Conn, adOpenForwardOnly, adLockReadOnly
        If rs.EOF Then
            errMsg = "对应部门[" + DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("cdepcode").nodeValue + "]是错误的!"
            EditSingleSO = False
            Exit Function
        End If
    End If
    '发送单位
    If Not DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("cfreightcompany") Is Nothing Then
        strSql = "select * from Vendor where cVenCode='" & DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("cfreightcompany").nodeValue & "'"
        rs.Open strSql, m_Conn, adOpenForwardOnly, adLockReadOnly
        If rs.EOF Then
            errMsg = "对应发送单位[" + DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("cfreightcompany").nodeValue + "]是错误的!"
            EditSingleSO = False
            Exit Function
        End If
    End If
    If Not DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("cfreight") Is Nothing Then
        If DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("cfreight").nodeValue = "是" Then
            If DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("mycdefinet2") Is Nothing Or DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("cfreighttype") Is Nothing Or DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("cfreightcost") Is Nothing Then
                errMsg = GetString("U8.DZ.JA.Res1150")
                EditSingleSO = False
                Exit Function
            End If
        End If
    End If
    If DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("cmaker") Is Nothing Then
        errMsg = GetString("U8.DZ.JA.Res1130")
        EditSingleSO = False
        Exit Function
    End If
    If DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("ddate") Is Nothing Then
        errMsg = "借用日期不能为空"
        EditSingleSO = False
        Exit Function
    Else
        sddate = DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("ddate").nodeValue
    End If

    '表体
    Dim nd As IXMLDOMElement
    Dim iRow As Integer
    iRow = 0
    Dim j As Integer: j = 1
    For Each nd In domBody.selectNodes("//z:row")

        iRow = iRow + 1
        nd.setAttribute "id", ssdi

        sAutoId = sAutoId + j - 1
        j = j + 1
        If Null2Something(nd.getAttribute("autoid")) = "" Then
            nd.setAttribute "autoid", sAutoId
        End If

        '1
        '比填项,简单校验：存货编码，数量
        '其他的校验根据实际业务而定
        If nd.getAttribute("cinvcode") = "" Then
            errMsg = "第" & iRow & "行存货编码不能为空"
            EditSingleSO = False
            Exit Function
        Else
            strSql = "SELECT i.cInvName,i.cInvStd,i.cInvAddCode," & _
                     "i.cGroupCode,g.cGroupName,i.cComUnitCode,u.cComunitName ,i.cSTComUnitCode as cAssUnit,u2.cComunitName as cAssUnitName,u2.iChangRate as iinvexchrate," & _
                     "i.cMassUnit ,i.imassdate,isub.iExpiratDateCalcu ," & _
                     "i.cInvDefine1,i.cInvDefine2,i.cInvDefine3,i.cInvDefine4,i.cInvDefine5," & _
                     "i.cInvDefine6,i.cInvDefine7,i.cInvDefine8,i.cInvDefine9,i.cInvDefine10," & _
                     "i.cInvDefine11,i.cInvDefine12,i.cInvDefine13,i.cInvDefine14,i.cInvDefine15,i.cInvDefine16, " & _
                     "i.iGroupType,i.bInvBatch,i.bInvQuality,i.bTrack," & _
                     "i.bFree1,i.bFree2,i.bFree3,i.bFree4,i.bFree5,i.bFree6,i.bFree7,i.bFree8,i.bFree9,i.bFree10," & _
                   " cast(isub.bSalePriceFree1 as int) bSalePriceFree1,cast(isub.bSalePriceFree2 as int) bSalePriceFree2,cast(isub.bSalePriceFree3 as int) bSalePriceFree3,cast(isub.bSalePriceFree4 as int) bSalePriceFree4,cast(isub.bSalePriceFree5 as int) bSalePriceFree5, " & _
                   " cast(isub.bSalePriceFree6 as int) bSalePriceFree6,cast(isub.bSalePriceFree7 as int) bSalePriceFree7,cast(isub.bSalePriceFree8 as int) bSalePriceFree8,cast(isub.bSalePriceFree9 as int) bSalePriceFree9,cast(isub.bSalePriceFree10 as int) bSalePriceFree10" & _
                   " from inventory i " & _
                   " left outer join Inventory_Sub isub on i.cinvcode=isub.cInvSubCode  " & _
                   " left outer join ComputationGroup g on g.cGroupCode=i.cGroupCode" & _
                   " left outer join ComputationUnit u on u.cComunitCode=i.cComunitCode" & _
                   " left outer join ComputationUnit u2 on i.cSTComUnitCode=u2.cComunitCode" & _
                   " where i.cinvcode='" & nd.getAttribute("cinvcode") & "'"
            rs.Open strSql, m_Conn, adOpenForwardOnly, adLockReadOnly
            If rs.EOF Then
                errMsg = "存货[" & nd.getAttribute("cinvcode") & "]不存在或者没有销售属性或者没有权限或者停用，请重新输入"
                EditSingleSO = False
                Exit Function
            End If
        End If

        If Val(nd.getAttribute("iquantity") & "") = 0 And Val(nd.getAttribute("inum") & "") = 0 Then
            '            MsgBox "第" & iRow & "行 数量件数不能同时为空或等于0", vbInformation, getstring("U8.DZ.JA.Res030")
            errMsg = "第" & iRow & "行 数量/件数不能同时为空或等于0"
            EditSingleSO = False
            Exit Function
        End If

        'dxb 仓库必须录入！
        If nd.getAttribute("cwhcode") = "" Then
            errMsg = "第" & iRow & "行仓库不能为空"
            EditSingleSO = False
            Exit Function
        End If

        '2
        '启用的自由项是否都输入了
        '结构性，非结构性自由项都必须输入
        '自由项组合是否合法
        If ExecFunFreeCheck(nd.getAttribute("cinvcode"), nd, iRow) = False Then
            '标识当前行,着蓝色

            EditSingleSO = False
            Exit Function
        End If


        '３
        '批次
        '出库跟踪入库
        If ExecFuncbatch(nd.getAttribute("cinvcode"), _
                         nd.getAttribute("cbatch") & "", _
                         nd.getAttribute("cinvouchcode") & "", _
                         iRow) = False Then
            '标识当前行,着蓝色

            EditSingleSO = False
            Exit Function
        End If



        '5 预计归还日期 dxb
        If nd.getAttribute("backdate") = "" Then
            errMsg = "第" & iRow & "行预计归还日期不能为空"
            EditSingleSO = False
            Exit Function
        Else
            If CDate(Mid(nd.getAttribute("backdate"), 1, 10)) < CDate(sddate) Then
                errMsg = "第" & iRow & "行预计归还日期不能早于借用日期"
                EditSingleSO = False
                Exit Function
            End If
        End If

    Next

    '发送方式
    If Not DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("cfreighttype") Is Nothing Then
        strSql = "select * from ShippingChoice where csccode='" & DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("cfreighttype").nodeValue & "'"
        rs.Open strSql, m_Conn, adOpenForwardOnly, adLockReadOnly
        If rs.EOF Then
            errMsg = "对应发送方式[" + DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("cfreighttype").nodeValue + "]是错误的!"
            EditSingleSO = False
            Exit Function
        End If
    End If


    '保存

    '表头
    Dim str1 As String
    Dim str2 As String
    str1 = " Update HY_DZ_BorrowOut set ccode='" & DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("ccode").nodeValue & "'"

    If Not DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("ccreatetype") Is Nothing Then
        str1 = str1 + ",ccreatetype='" & DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("ccreatetype").nodeValue & "' "
    End If
    If Not DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("bobjectcode") Is Nothing Then str1 = str1 + ",bObjectCode='" & DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("bobjectcode").nodeValue & "' "
    If Not DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("caboutvoucher") Is Nothing Then str1 = str1 + ",cAboutVoucher='" & DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("caboutvoucher").nodeValue & "' "
    If Not DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("ccodeaboutvoucher") Is Nothing Then str1 = str1 + ",cCodeAboutVoucher='" & DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("ccodeaboutvoucher").nodeValue & "' "
    If Not DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("ccontactperson") Is Nothing Then str1 = str1 + ",cContactperson='" & DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("ccontactperson").nodeValue & "' "
    If Not DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("ccontactway") Is Nothing Then str1 = str1 + ",cContactWay='" & DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("ccontactway").nodeValue & "' "
    If Not DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("ccreatetype") Is Nothing Then str1 = str1 + ",cCreateType='" & DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("ccreatetype").nodeValue & "' "
    If Not DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("cdefine1") Is Nothing Then str1 = str1 + ",cdefine1='" & DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("cdefine1").nodeValue & "' "
    If Not DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("cdefine10") Is Nothing Then str1 = str1 + ",cdefine10='" & DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("cdefine10").nodeValue & "' "
    If Not DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("cdefine11") Is Nothing Then str1 = str1 + ",cdefine11='" & DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("cdefine11").nodeValue & "' "
    If Not DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("cdefine12") Is Nothing Then str1 = str1 + ",cdefine12='" & DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("cdefine12").nodeValue & "' "
    If Not DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("cdefine13") Is Nothing Then str1 = str1 + ",cdefine13='" & DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("cdefine13").nodeValue & "' "
    If Not DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("cdefine14") Is Nothing Then str1 = str1 + ",cdefine14='" & DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("cdefine14").nodeValue & "' "
    If Not DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("cdefine15") Is Nothing Then str1 = str1 + ",cdefine15='" & DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("cdefine15").nodeValue & "' "
    If Not DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("cdefine16") Is Nothing Then str1 = str1 + ",cdefine16='" & DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("cdefine16").nodeValue & "' "
    If Not DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("cdefine2") Is Nothing Then str1 = str1 + ",cdefine2='" & DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("cdefine2").nodeValue & "' "
    If Not DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("cdefine3") Is Nothing Then str1 = str1 + ",cdefine3='" & DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("cdefine3").nodeValue & "' "
    If Not DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("cdefine4") Is Nothing Then str1 = str1 + ",cdefine4='" & DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("cdefine4").nodeValue & "' "
    If Not DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("cdefine5") Is Nothing Then str1 = str1 + ",cdefine5='" & DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("cdefine5").nodeValue & "' "
    If Not DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("cdefine6") Is Nothing Then str1 = str1 + ",cdefine6='" & DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("cdefine6").nodeValue & "' "
    If Not DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("cdefine7") Is Nothing Then str1 = str1 + ",cdefine7='" & DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("cdefine7").nodeValue & "' "
    If Not DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("cdefine8") Is Nothing Then str1 = str1 + ",cdefine8='" & DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("cdefine8").nodeValue & "' "
    If Not DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("cdefine9") Is Nothing Then str1 = str1 + ",cdefine9='" & DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("cdefine9").nodeValue & "' "
    If Not DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("cdepcode") Is Nothing Then str1 = str1 + ",cdepcode='" & DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("cdepcode").nodeValue & "' "
    If Not DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("cfreight") Is Nothing Then str1 = str1 + ",cfreight='" & DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("cfreight").nodeValue & "' "
    If Not DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("cfreightcompany") Is Nothing Then str1 = str1 + ",cfreightCompany='" & DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("cfreightcompany").nodeValue & "' "
    If Not DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("cfreightcost") Is Nothing Then str1 = str1 + ",cfreightCost='" & DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("cfreightcost").nodeValue & "' "
    If Not DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("cfreighttype") Is Nothing Then str1 = str1 + ",cfreightType='" & DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("cfreighttype").nodeValue & "' "
    If Not DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("chandler") Is Nothing Then str1 = str1 + ",cHandler='" & DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("chandler").nodeValue & "' "
    If Not DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("closeuser") Is Nothing Then str1 = str1 + ",CloseUser='" & DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("closeuser").nodeValue & "' "
    If Not DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("cmaker") Is Nothing Then str1 = str1 + ",cMaker='" & DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("cmaker").nodeValue & "' "
    If Not DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("cmemo") Is Nothing Then str1 = str1 + ",cmemo='" & DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("cmemo").nodeValue & "' "
    If Not DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("cpersoncode") Is Nothing Then str1 = str1 + ",cpersoncode='" & DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("cpersoncode").nodeValue & "' "
    If Not DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("ctype") Is Nothing Then str1 = str1 + ",cType='" & DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("ctype").nodeValue & "' "
    If Not DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("dclosedate") Is Nothing Then str1 = str1 + ",dCloseDate='" & DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("dclosedate").nodeValue & "' "
    If Not DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("ddate") Is Nothing Then str1 = str1 + ",ddate='" & DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("ddate").nodeValue & "' "
    If Not DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("dintodate") Is Nothing Then str1 = str1 + ",dIntoDate='" & DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("dintodate").nodeValue & "' "
    If Not DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("dmdate") Is Nothing Then str1 = str1 + ",dmDate='" & DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("dmdate").nodeValue & "' "
    If Not DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("downstreamcode") Is Nothing Then str1 = str1 + ",DownstreamCode='" & DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("downstreamcode").nodeValue & "' "
    If Not DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("dveridate") Is Nothing Then str1 = str1 + ",dVeriDate='" & DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("dveridate").nodeValue & "' "
    If Not DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("intouser") Is Nothing Then str1 = str1 + ",IntoUser='" & DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("intouser").nodeValue & "' "
    If Not DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("ireturncount") Is Nothing Then str1 = str1 + ",ireturncount='" & DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("ireturncount").nodeValue & "' "
    If Not DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("istatus") Is Nothing Then str1 = str1 + ",iStatus='" & DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("istatus").nodeValue & "' "
    If Not DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("iswfcontrolled") Is Nothing Then str1 = str1 + ",iswfcontrolled='" & DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("iswfcontrolled").nodeValue & "' "
    If Not DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("iverifystate") Is Nothing Then str1 = str1 + ",iverifystate='" & DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("iverifystate").nodeValue & "' "
    If Not DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("mycdefinet1") Is Nothing Then str1 = str1 + ",MycdefineT1='" & DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("mycdefinet1").nodeValue & "' "
    If Not DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("mycdefinet10") Is Nothing Then str1 = str1 + ",MycdefineT10='" & DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("mycdefinet10").nodeValue & "' "
    If Not DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("mycdefinet2") Is Nothing Then str1 = str1 + ",MycdefineT2='" & DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("mycdefinet2").nodeValue & "' "
    If Not DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("mycdefinet3") Is Nothing Then str1 = str1 + ",MycdefineT3='" & DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("mycdefinet3").nodeValue & "' "
    If Not DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("mycdefinet4") Is Nothing Then str1 = str1 + ",MycdefineT4='" & DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("mycdefinet4").nodeValue & "' "
    If Not DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("mycdefinet5") Is Nothing Then str1 = str1 + ",MycdefineT5='" & DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("mycdefinet5").nodeValue & "' "
    If Not DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("mycdefinet6") Is Nothing Then str1 = str1 + ",MycdefineT6='" & DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("mycdefinet6").nodeValue & "' "
    If Not DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("mycdefinet7") Is Nothing Then str1 = str1 + ",MycdefineT7='" & DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("mycdefinet7").nodeValue & "' "
    If Not DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("mycdefinet8") Is Nothing Then str1 = str1 + ",MycdefineT8='" & DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("mycdefinet8").nodeValue & "' "
    If Not DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("mycdefinet9") Is Nothing Then str1 = str1 + ",MycdefineT9='" & DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("mycdefinet9").nodeValue & "' "
    If Not DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("upstreamcode") Is Nothing Then str1 = str1 + ",UpStreamCode='" & DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("upstreamcode").nodeValue & "' "
    If Not DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("vouchertype") Is Nothing Then str1 = str1 + ",VoucherType='" & DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("vouchertype").nodeValue & "' "
    str1 = str1 + " where ID=" & DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("id").nodeValue & ""
    m_Conn.Execute str1
    '表体
    iRow = 0
    Dim str As String
    m_Conn.Execute "delete from HY_DZ_BorrowOuts where ID=" & DomHeader.selectSingleNode("//z:row").Attributes.getNamedItem("id").nodeValue & ""
    m_Conn.Execute "update ufsystem..ua_identity  set ichildid=" & sAutoId & " where cacc_id='" & mLogin.cAcc_Id & "' and cvouchtype='HY_JCGH001'"

    str = ""
    For Each nd In domBody.selectNodes("//z:row")

        iRow = iRow + 1
        str1 = " insert into HY_DZ_BorrowOuts(ID,AutoID,"
        str2 = " values(" & sID & "," & nd.getAttribute("autoid") & ","

        If nd.getAttribute("backdate") = "" Then
            str1 = str1 + " backdate, "
            str2 = str2 + "'" & nd.getAttribute("backdate") & "',"
        End If
        If nd.getAttribute("cbatch") = "" Then
            str1 = str1 + " cbatch, "
            str2 = str2 + "'" & nd.getAttribute("cbatch") & "',"
        End If
        If nd.getAttribute("cdefine22") = "" Then
            str1 = str1 + " cdefine22, "
            str2 = str2 + "'" & nd.getAttribute("cdefine22") & "',"
        End If
        If nd.getAttribute("cdefine23") = "" Then
            str1 = str1 + " cdefine23, "
            str2 = str2 + "'" & nd.getAttribute("cdefine23") & "',"
        End If
        If nd.getAttribute("cdefine24") = "" Then
            str1 = str1 + " cdefine24, "
            str2 = str2 + "'" & nd.getAttribute("cdefine24") & "',"
        End If
        If nd.getAttribute("cdefine25") = "" Then
            str1 = str1 + " cdefine25, "
            str2 = str2 + "'" & nd.getAttribute("cdefine25") & "',"
        End If
        If nd.getAttribute("cdefine26") = "" Then
            str1 = str1 + " cdefine26, "
            str2 = str2 + "'" & nd.getAttribute("cdefine26") & "',"
        End If
        If nd.getAttribute("cdefine27") = "" Then
            str1 = str1 + " cdefine27, "
            str2 = str2 + "'" & nd.getAttribute("cdefine27") & "',"
        End If
        If nd.getAttribute("cdefine28") = "" Then
            str1 = str1 + " cdefine28, "
            str2 = str2 + "'" & nd.getAttribute("cdefine28") & "',"
        End If
        If nd.getAttribute("cdefine29") = "" Then
            str1 = str1 + " cdefine29, "
            str2 = str2 + "'" & nd.getAttribute("cdefine29") & "',"
        End If
        If nd.getAttribute("cdefine30") = "" Then
            str1 = str1 + " cdefine30, "
            str2 = str2 + "'" & nd.getAttribute("cdefine30") & "',"
        End If
        If nd.getAttribute("cdefine31") = "" Then
            str1 = str1 + " cdefine31, "
            str2 = str2 + "'" & nd.getAttribute("cdefine31") & "',"
        End If
        If nd.getAttribute("cdefine32") = "" Then
            str1 = str1 + " cdefine32, "
            str2 = str2 + "'" & nd.getAttribute("cdefine32") & "',"
        End If
        If nd.getAttribute("cdefine33") = "" Then
            str1 = str1 + " cdefine33, "
            str2 = str2 + "'" & nd.getAttribute("cdefine33") & "',"
        End If
        If nd.getAttribute("cdefine34") = "" Then
            str1 = str1 + " cdefine34, "
            str2 = str2 + "'" & nd.getAttribute("cdefine34") & "',"
        End If
        If nd.getAttribute("cdefine35") = "" Then
            str1 = str1 + " cdefine35, "
            str2 = str2 + "'" & nd.getAttribute("cdefine35") & "',"
        End If
        If nd.getAttribute("cdefine36") = "" Then
            str1 = str1 + " cdefine36, "
            str2 = str2 + "'" & nd.getAttribute("cdefine36") & "',"
        End If
        If nd.getAttribute("cdefine37") = "" Then
            str1 = str1 + " cdefine37, "
            str2 = str2 + "'" & nd.getAttribute("cdefine37") & "',"
        End If
        If nd.getAttribute("cfree1") = "" Then
            str1 = str1 + " cfree1, "
            str2 = str2 + "'" & nd.getAttribute("cfree1") & "',"
        End If
        If nd.getAttribute("cfree10") = "" Then
            str1 = str1 + " cfree10, "
            str2 = str2 + "'" & nd.getAttribute("cfree10") & "',"
        End If
        If nd.getAttribute("cfree2") = "" Then
            str1 = str1 + " cfree2, "
            str2 = str2 + "'" & nd.getAttribute("cfree2") & "',"
        End If
        If nd.getAttribute("cfree3") = "" Then
            str1 = str1 + " cfree3, "
            str2 = str2 + "'" & nd.getAttribute("cfree3") & "',"
        End If
        If nd.getAttribute("cfree4") = "" Then
            str1 = str1 + " cfree4, "
            str2 = str2 + "'" & nd.getAttribute("cfree4") & "',"
        End If
        If nd.getAttribute("cfree5") = "" Then
            str1 = str1 + " cfree5, "
            str2 = str2 + "'" & nd.getAttribute("cfree5") & "',"
        End If
        If nd.getAttribute("cfree6") = "" Then
            str1 = str1 + " cfree6, "
            str2 = str2 + "'" & nd.getAttribute("cfree6") & "',"
        End If
        If nd.getAttribute("cfree7") = "" Then
            str1 = str1 + " cfree7, "
            str2 = str2 + "'" & nd.getAttribute("cfree7") & "',"
        End If
        If nd.getAttribute("cfree8") = "" Then
            str1 = str1 + " cfree8, "
            str2 = str2 + "'" & nd.getAttribute("cfree8") & "',"
        End If
        If nd.getAttribute("cfree9") = "" Then
            str1 = str1 + " cfree9, "
            str2 = str2 + "'" & nd.getAttribute("cfree9") & "',"
        End If
        If nd.getAttribute("cinvcode") = "" Then
            str1 = str1 + " cinvcode, "
            str2 = str2 + "'" & nd.getAttribute("cinvcode") & "',"
        End If
        If nd.getAttribute("cmemo") = "" Then
            str1 = str1 + " cmemo, "
            str2 = str2 + "'" & nd.getAttribute("cmemo") & "',"
        End If
        If nd.getAttribute("cposition") = "" Then
            str1 = str1 + " cposition, "
            str2 = str2 + "'" & nd.getAttribute("cposition") & "',"
        End If
        If nd.getAttribute("cunitid") = "" Then
            str1 = str1 + " cunitid, "
            str2 = str2 + "'" & nd.getAttribute("cunitid") & "',"
        End If
        If nd.getAttribute("cwhcode") = "" Then
            str1 = str1 + " cwhcode, "
            str2 = str2 + "'" & nd.getAttribute("cwhcode") & "',"
        End If
        If nd.getAttribute("iinvexchrate") = "" Then
            str1 = str1 + " iinvexchrate, "
            str2 = str2 + "'" & nd.getAttribute("iinvexchrate") & "',"
        End If
        If nd.getAttribute("inum") = "" Then
            str1 = str1 + " inum, "
            str2 = str2 + "'" & nd.getAttribute("inum") & "',"
        End If
        If nd.getAttribute("iquantity") = "" Then
            str1 = str1 + " iquantity, "
            str2 = str2 + "'" & nd.getAttribute("iquantity") & "',"
        End If
        If nd.getAttribute("MycdefineB1") = "" Then
            str1 = str1 + " MycdefineB1, "
            str2 = str2 + "'" & nd.getAttribute("MycdefineB1") & "',"
        End If
        If nd.getAttribute("MycdefineB10") = "" Then
            str1 = str1 + " MycdefineB10, "
            str2 = str2 + "'" & nd.getAttribute("MycdefineB10") & "',"
        End If
        If nd.getAttribute("MycdefineB2") = "" Then
            str1 = str1 + " MycdefineB2, "
            str2 = str2 + "'" & nd.getAttribute("MycdefineB2") & "',"
        End If
        If nd.getAttribute("MycdefineB3") = "" Then
            str1 = str1 + " MycdefineB3, "
            str2 = str2 + "'" & nd.getAttribute("MycdefineB3") & "',"
        End If
        If nd.getAttribute("MycdefineB4") = "" Then
            str1 = str1 + " MycdefineB4, "
            str2 = str2 + "'" & nd.getAttribute("MycdefineB4") & "',"
        End If
        If nd.getAttribute("MycdefineB5") = "" Then
            str1 = str1 + " MycdefineB5, "
            str2 = str2 + "'" & nd.getAttribute("MycdefineB5") & "',"
        End If
        If nd.getAttribute("MycdefineB6") = "" Then
            str1 = str1 + " MycdefineB6, "
            str2 = str2 + "'" & nd.getAttribute("MycdefineB6") & "',"
        End If
        If nd.getAttribute("MycdefineB7") = "" Then
            str1 = str1 + " MycdefineB7, "
            str2 = str2 + "'" & nd.getAttribute("MycdefineB7") & "',"
        End If
        If nd.getAttribute("MycdefineB8") = "" Then
            str1 = str1 + " MycdefineB8, "
            str2 = str2 + "'" & nd.getAttribute("MycdefineB8") & "',"
        End If
        If nd.getAttribute("MycdefineB9") = "" Then
            str1 = str1 + " MycdefineB9, "
            str2 = str2 + "'" & nd.getAttribute("MycdefineB9") & "',"
        End If
        str = str + vbCrLf + Left(str1, Len(str1) - 1) + ")" + Left(str2, Len(str2) - 1) + ")"

    Next
    m_Conn.Execute str1

    '结束

    ''    WriteLog 0, "开始保存单据", "SA", "SalesEAI.dll", "SalesEAICls", "AddSingleSO"
    '    If LCase(sMode) = "ar" Or (bRetailEAI And Not m_bRetail) Then m_Conn.BeginTrans
    '    errMsg = SaveHandle.Save(DomHeader, domBody, 0, nNewID)
    '    If LCase(sMode) = "ar" Then
    '        If Trim(errMsg) = "" Then
    '            m_Conn.CommitTrans
    '            WriteLog 0, "保存单据成功", "SA", "SalesEAI.dll", "SalesEAICls", "AddSingleSO"
    '        Else
    '            m_Conn.RollbackTrans
    '            WriteLog 1, "保存单据失败" + errMsg, "SA", "SalesEAI.dll", "SalesEAICls", "AddSingleSO"
    '        End If
    '    End If
    '
    '    If m_Ids = "" Then '获得保存后新的单据id值
    '        m_Ids = CStr(nNewID)
    '    Else
    '        m_Ids = m_Ids & "," & CStr(nNewID)
    '    End If

    '------------永川项目，自动审核发货单-----------------2005.12.15
    '销售EAI改进 , 在header节点增加属性bautoverify, 用来标志是否自动审核
    '    If errMsg = "" Then
    '        Dim sErrMsgVerify As String
    '        WriteLog 0, "开始审核单据", "SA", "SalesEAI.dll", "SalesEAICls", "AddSingleSO"
    '        If bRetailEAI Then
    '            eleHead.setAttribute "beairetailvouch", "1"
    '        End If
    '        Call AutoVerify(SaveHandle, ViewHead, ViewBody, nNewID, eleHead, sMode, sErrMsgVerify)
    '        If sErrMsgVerify = "" Then
    '            If LCase(sMode) <> "ar" And bRetailEAI And Not m_bRetail Then
    '                m_Conn.CommitTrans
    '            End If
    '            WriteLog 0, "审核单据成功", "SA", "SalesEAI.dll", "SalesEAICls", "AddSingleSO"
    '        Else
    '            If LCase(sMode) <> "ar" And bRetailEAI And Not m_bRetail Then
    '                m_Conn.RollbackTrans
    '            End If
    '            WriteLog 1, "审核单据失败" + sErrMsgVerify, "SA", "SalesEAI.dll", "SalesEAICls", "AddSingleSO"
    '        End If
    '        errMsg = errMsg & sErrMsgVerify
    '    Else
    '        If LCase(sMode) <> "ar" And bRetailEAI And Not m_bRetail Then
    '            m_Conn.RollbackTrans
    '        End If
    '    End If
    '    '---------------------------------------------------
    '
errorhandle:
    If (errMsg = "") And (Err.Description = "") Then
        Call ErrDomAppendChild(m_errDom, sBillCode, 0, "ok")
        Exit Function
    Else
        Call ErrDomAppendChild(m_errDom, sBillCode, 101, sBillCode & "导入失败" & errMsg & Err.Description)
        Exit Function
    End If
DoErr:
    Call ErrDomAppendChild(m_errDom, sBillCode, 101, "程序错误" & Err.Description)

End Function

'批次
'入库单号
Private Function ExecFuncbatch(cinvcode As String, cbatch As String, cTrackCode As String, iRow As Integer) As Boolean
    On Error GoTo Err_Handler

    Dim sql As String
    Dim rs As New ADODB.Recordset

    sql = "select bInvBatch,bTrack from inventory where cinvcode='" & cinvcode & "'"
    rs.Open sql, m_Conn, 1, 1
    If Not rs.EOF Then
        If CBool(rs("bInvBatch")) And cbatch = "" Then
            errMsg = "第" & iRow & "行存货启用了批次管理，批号不能为空"
            '            MsgBox "第" & iRow & "行存货启用了批次管理，批号不能为空", vbInformation, getstring("U8.DZ.JA.Res030")
            ExecFuncbatch = False
            rs.Close
            Set rs = Nothing
            Exit Function
        End If
        '        If CBool(Rs("bTrack")) And cTrackCode = "" Then
        '            MsgBox "第" & iRow & "行存货启用了出库跟踪入库，入库单号不能为空", vbInformation, getstring("U8.DZ.JA.Res030")
        '            ExecFuncbatch = False
        '            Rs.Close
        '            Set Rs = Nothing
        '            Exit Function
        '        End If
    Else
        errMsg = "第" & iRow & "行存货编码" & cinvcode & "不存在"
        '        MsgBox "第" & iRow & "行存货编码" & cinvcode & "不存在", vbInformation, getstring("U8.DZ.JA.Res030")
        ExecFuncbatch = False
        rs.Close
        Set rs = Nothing
        Exit Function
    End If

    rs.Close
    Set rs = Nothing
    ExecFuncbatch = True
    Exit Function

Err_Handler:
    rs.Close
    Set rs = Nothing
    ExecFuncbatch = False
    '    MsgBox Err.Description, vbInformation, getstring("U8.DZ.JA.Res030")
End Function

'结构性自由项

'启用的自由项是否都输入了
'结构性，非结构性自由项都必须输入
'自由项组合是否合法
Private Function ExecFunFreeCheck(cinvcode As String, _
                                  bodyele As IXMLDOMElement, _
                                  iRow As Integer _
                                ) As Boolean
    On Error GoTo Err_Handler

    'cinvcode :bodyele.getAttribute("cinvcode")
    'cFree:bodyele.getAttribute("cfree1")
    'irow:irow 行号
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim bRS As New ADODB.Recordset
    Dim sFind As String
    Dim i As Integer

    sFind = "select 1 as Free from bas_part where invcode='" & cinvcode & "'"

    sql = "select bFree1,bConfigFree1,bFree2,bConfigFree2,bFree3,bConfigFree3,bFree4,bConfigFree4," & _
          "bFree5,bConfigFree5,bFree6,bConfigFree6,bFree7,bConfigFree7,bFree8,bConfigFree8,bFree9," & _
          "bConfigFree9,bFree10,bConfigFree10 from inventory WHERE cInvCode='" & cinvcode & "'"
    rs.Open sql, m_Conn, 1, 1
    If Not rs.EOF Then

        For i = 1 To 10
            If rs("bFree" & i) = True And bodyele.getAttribute("cfree" & i) = "" Then
                bRS.Open "select cItemName from userdef where cClass = '存货' and cDicDbname='cFree" & i & "'", g_Conn, 1, 1
                errMsg = "第" & iRow & "行存货" & cinvcode & "启用了自由项" & bRS("cItemName") & ",必须输入"
                '                    MsgBox "第" & iRow & "行存货" & cinvcode & "启用了自由项" & bRS("cItemName") & ",必须输入", vbInformation, getstring("U8.DZ.JA.Res030")
                ExecFunFreeCheck = False
                bRS.Close
                Set bRS = Nothing
                rs.Close
                Set rs = Nothing
                Exit Function
            Else
                If rs("bConfigFree" & i) = True Then sFind = sFind & " and Free" & CStr(i) & "='" & bodyele.getAttribute("cfree" & CStr(i)) & "'"
            End If
        Next i

        bRS.Open sFind, m_Conn, 1, 1
        If bRS.EOF Then
            errMsg = "第" & iRow & "行存货结构性自由项组合不合法"
            '                MsgBox "第" & iRow & "行存货结构性自由项组合不合法", vbInformation, getstring("U8.DZ.JA.Res030")
            bRS.Close
            Set bRS = Nothing
            rs.Close
            Set rs = Nothing
            ExecFunFreeCheck = False
            Exit Function
        End If


    Else
        errMsg = "第" & iRow & "行存货编码" & cinvcode & "不存在"
        '            MsgBox "第" & iRow & "行存货编码" & cinvcode & "不存在", vbInformation, getstring("U8.DZ.JA.Res030")
        ExecFunFreeCheck = False
        rs.Close
        Set rs = Nothing
        Exit Function
    End If



    rs.Close
    Set rs = Nothing
    ExecFunFreeCheck = True
    Exit Function



Err_Handler:
    rs.Close
    Set rs = Nothing
    ExecFunFreeCheck = False
    ' MsgBox Err.Description, vbInformation, getstring("U8.DZ.JA.Res030")
End Function

'--------------------------------------------------------------------------
' 功能： 导出销售单据
' 参数说明：
' 同名参数定义同TransactGeneral
'--------------------------------------------------------------------------

Public Function QueryProc(ByVal eleList As IXMLDOMNodeList, ByRef sRetXml As String, RSXMLRefile As String, TableName As String, SubTableName As String) As Boolean
    Dim rs As New Recordset
    Dim rsDetail As New Recordset
    Dim dom As New DOMDocument
    Dim sql As String
    Dim el As IXMLDOMElement
    Dim EXmlToRs As New EAIXMLRSExch20.clsXmlRsExch

    dom.loadXML sRetXml
    dom.documentElement.Text = ""
    WriteLog 0, "导出销售单据", "SA", "SalesEAI.dll", "SalesEAICls", "QueryProc"
    Dim paginate As Long
    Dim counter As Long    '计数器
    Dim lIndex As Long  '导出第几页
    Dim sumCount As Long    '记录的总数

    Dim tempDom As New DOMDocument
    tempDom.loadXML sRetXml
    tempDom.documentElement.Text = ""

    If Not dom.documentElement.Attributes.getNamedItem("paginate") Is Nothing Then
        paginate = CLng(dom.documentElement.Attributes.getNamedItem("paginate").Text)
    Else
        paginate = 0
    End If
    If paginate > 0 Then
        '取得页的总数
        Dim sumSql As String
        Dim lPageCount As Long

        '        sql = "select * from " & TableName & "  " & GetQueryCondition(eleList, m_Conn, TableName)
        sumSql = "select count(*) from " + TableName & "  " & GetQueryCondition(eleList, m_Conn, TableName)
        rs.Open sumSql, m_Conn, adOpenStatic, adLockReadOnly

        If Not rs.EOF Then
            sumCount = rs.Fields(0)
        End If
        If sumCount Mod paginate = 0 Then
            lPageCount = sumCount / paginate
        Else
            lPageCount = Int(sumCount / paginate) + 1
        End If

        rs.Close
        '定义计数器
        counter = 1
        lIndex = 1
    End If
    If TableName = "dispatchlist" Then
        '        sql = "if not exists(select name from sysobjects where name='sa_dispatchlist' and xtype='V') exec ('CREATE  VIEW SA_Dispatchlist " _
                 '            & " AS " _
                 '            & " select convert(nchar,convert(money,dispatchlist.ufts),2) as ufts,cbustype,caccounter,ivtid," _
                 '            & " ccrechpname,dispatchlist.cdlcode, dispatchlist.cvouchtype, vouchtype.cvouchname, dispatchlist.cstcode," _
                 '            & " saletype.cstname, dispatchlist.ddate, dispatchlist.cdepcode, department.cdepname," _
                 '            & " dispatchlist.cpersoncode, person.cpersonname, dispatchlist.csbvcode, dispatchlist.csocode," _
                 '            & " dispatchlist.ccuscode, customer.ccusabbname, dispatchlist.cpaycode, dispatchlist.csccode," _
                 '            & " shippingchoice.cscname, dispatchlist.cshipaddress, dispatchlist.cexch_name, dispatchlist.iexchrate," _
                 '            & " dispatchlist.itaxrate, dispatchlist.cmemo, dispatchlist.cdefine1, dispatchlist.cdefine2," _
                 '            & " DispatchList.inetlock , DispatchList.breturnflag, PayCondition.cpayname, customer.ccusoaddress, customer.ccusotype, customer.ccuspaycond, dispatchlist.dlid, dispatchlist.cverifier, dispatchlist.cmaker," _
                 '            & " customer.icuscreline, customer.iarmoney, dispatchlist.bfirst, dispatchlist.sbvid," _
                 '            & " dispatchlist.cdefine3, dispatchlist.cdefine4, dispatchlist.cdefine5, dispatchlist.cdefine6," _
                 '            & " dispatchlist.cdefine7, dispatchlist.cdefine8, dispatchlist.cdefine9, dispatchlist.cdefine10," _
                 '            & " dispatchlist.cdefine11,dispatchlist.cdefine12,dispatchlist.cdefine13,dispatchlist.cdefine14," _
                 '            & " dispatchlist.cdefine15,dispatchlist.cdefine16,dispatchlist.isale, dispatchlist.ccusname, customer.ccusphone,customer.ccuspostcode," _
                 '            & " customer.ccusperson,customer.ccusdefine1,customer.ccusdefine2,customer.ccusdefine3,customer.ccusdefine4,customer.ccusdefine5,customer.ccusdefine6," _
                 '            & " customer.ccusdefine7,customer.ccusdefine8,customer.ccusdefine9,customer.ccusdefine10,customer.ccusdefine11,customer.ccusdefine12,customer.ccusdefine13," _
                 '            & " customer.ccusdefine14,customer.ccusdefine15,customer.ccusdefine16,dispatchlist.crdcode,ccloser,customer.iid as cauthid,customer.ccusaddress" _
                 '            & " ,cusdeliveradd.caddcode, cusdeliveradd.cdeliverunit, Crm_Contact.cmobilephone, Crm_Contact.cofficephone" _
                 '            & " FROM DispatchList inner JOIN Customer ON DispatchList.cCusCode = Customer.cCusCode LEFT OUTER JOIN Department ON DispatchList.cDepCode = Department.cDepCode LEFT OUTER JOIN PayCondition ON" _
                 '            & " DispatchList.cPayCode = PayCondition.cPayCode LEFT OUTER JOIN Person ON DispatchList.cPersonCode = Person.cPersonCode LEFT OUTER JOIN SaleType ON" _
                 '            & " DispatchList.cSTCode = SaleType.cSTCode LEFT OUTER JOIN ShippingChoice ON DispatchList.cSCCode = ShippingChoice.cSCCode LEFT OUTER JOIN VouchType ON" _
                 '            & " DispatchList.cVouchType = vouchtype.cVouchType LEFT join cusdeliveradd on DispatchList.ccuscode=CusDeliverAdd.ccuscode  and DispatchList.caddcode = CusDeliverAdd.caddcode" _
                 '            & " LEFT JOIN Crm_Contact ON CusDeliverAdd.cLinkPerson = Crm_Contact.cContactCode  " _
                 '            & " where (bFirst=1 or (bFirst=0 and dDate>=(select cvalue from accinformation where csysid=N''SA'' and cName=N''dStartDate''))) )')"
        '        m_Conn.Execute sql
        sql = "select * from sa_dispatchlist " & "  " & GetQueryCondition(eleList, m_Conn, TableName)
    Else
        sql = "select * from " & TableName & "  " & GetQueryCondition(eleList, m_Conn, TableName)
    End If
    rs.Open sql, m_Conn, adOpenStatic, adLockReadOnly
    While Not rs.EOF
        Set el = dom.createElement(dom.documentElement.getAttribute("roottag"))
        Call EXmlToRs.RStoXmlHead(el, RSXMLRefile, rs)
        Select Case LCase(TableName)
        Case "so_somain"
            sql = "select * from so_sodetails where ID=" & rs("ID")
        Case "dispatchlist"
            '                sql = "select * from dispatchlists where DLID=" & rs("DLID")
            sql = "select * from sales_fhd_w where DLID=" & rs("DLID")
        Case LCase("SaleBillVouchZT"), LCase("SaleBillVouchZT_All")
            'Case "SaleBillVouchZT_All"
            sql = "select * from SaleBillVouchZW where SBVID=" & rs("SBVID")
        Case "sa_cuspricejustmain"
            sql = " select * from SA_CusPriceJustDetailView where id=" & rs("id")
        Case "sa_invpricejustmain"
            sql = "select * from SA_InvPriceJustDetailView where id=" & rs("id")
            '                sql = "select * from SA_InvPriceJustDetail where id=" & rs("id")
        End Select

        rsDetail.Open sql, m_Conn, adOpenStatic, adLockReadOnly

        Call EXmlToRs.RStoXmlBody(el, RSXMLRefile, rsDetail)

        rsDetail.Close
        dom.documentElement.appendChild el
        If paginate > 0 Then
            If counter = paginate Then
                gObjRef.ProcessExport dom.xml, lIndex, lPageCount
                '清空Dom
                Set dom = Nothing
                Set dom = New DOMDocument
                dom.loadXML tempDom.xml

                lIndex = lIndex + 1
                counter = 0
            Else
                If (lIndex = lPageCount) And (counter = sumCount Mod paginate) Then
                    gObjRef.ProcessExport dom.xml, lIndex, lPageCount
                End If
            End If
            counter = counter + 1
        End If
        rs.MoveNext
    Wend
    rs.Close
    Set rs = Nothing
    Set rsDetail = Nothing

    sRetXml = dom.xml
End Function

'--------------------------------------------------------------------------
' 功能： 获得用户在界面上输入的查询条件
' 参数说明：
' eleList As IXMLDOMNodeList,   查询条件XML文件描述
' ByRef conn As Connection,     数据库连接
' TableName As String           待查目的数据库表名
'--------------------------------------------------------------------------
'2004.07.01 modified by jzq
Private Function GetQueryCondition(eleList As IXMLDOMNodeList, ByRef conn As Connection, TableName As String) As String
    Dim rs As New Recordset
    Dim ele As IXMLDOMElement
    Dim fld As ADODB.Field
    Dim where As String
    Dim uftsValue As String, uftsOperation As String
    Dim uftsWhere As String

    If eleList.Length > 0 Then
        For Each ele In eleList
            If ele.nodename = "sql" Then
                If Not ele.Attributes.getNamedItem("value") Is Nothing Then
                    GetQueryCondition = ele.Attributes.getNamedItem("value").nodeValue
                End If
                If Trim(GetQueryCondition) <> "" Then
                    GetQueryCondition = " where " & GetQueryCondition
                End If
                Exit Function
            End If
        Next
    End If

    rs.Open "select * from " & TableName & " where 1>1", conn, adOpenForwardOnly, adLockReadOnly

    where = ""
    Dim i As Integer
    i = 1
    For Each ele In eleList
        Set fld = rs(GetAttrValue(ele, "name"))
        If Not IsNull(fld) Then
            If fld.Name = "ufts" Then
                '----------拼写时间戳条件-------
                uftsValue = GetAttrValue(ele, "value")
                uftsOperation = GetAttrValue(ele, "operation")
                '---2004.07.05将传入的ufts转换为money后再转换为char ---- by jzq
                '                uftsWhere = uftsWhere & IIf(uftsWhere = "", "", " AND ") _
                                 '                            & "( convert(char,convert(money," & TableName & ".ufts),2) " _
                                 '                            & uftsOperation & "convert(char,convert(money," & uftsValue & "),2))"
                uftsWhere = uftsWhere & IIf(uftsWhere = "", "", " AND ") _
                          & "( convert(char,convert(money," & IIf(LCase(TableName) = "dispatchlist", "", TableName & ".") & "ufts),2) " _
                          & uftsOperation & "convert(char,convert(money," & uftsValue & "),2))"

            Else
                '---------拼写普通字段条件----
                If i = 1 Then
                    where = ""
                End If
                where = where & fld.Name & GetAttrValue(ele, "operation")
                Select Case fld.Type
                Case adBSTR, adDate, adDBDate, adDBTime, adDBTimeStamp, _
                     adLongVarChar, adLongVarBinary, adBSTR, adLongVarChar, _
                     adVarChar, adVarWChar, adWChar
                    where = where & "'" & GetAttrValue(ele, "value") & "'"
                Case Else
                    where = where & GetAttrValue(ele, "value")
                End Select
                If i < eleList.Length Then
                    where = where & " " & GetAttrValue(ele, "logic") & " "
                End If
            End If
            i = i + 1
        End If
    Next

    Set rs = Nothing

    where = Trim(where)
    If Len(where) > 0 Then
        '截去条件尾部的连接词and或者or---2004.07.14 by jzq---
        If Trim(LCase(Right(where, 4))) = "and" Then
            where = Left(where, Len(where) - 3)
        ElseIf Trim(LCase(Right(where, 3))) = "or" Then
            where = Left(where, Len(where) - 2)
        End If
        where = " ( " & where & " ) "
        '过滤条件与时间戳是“ and ”的关系
        where = IIf(Trim(uftsWhere) = "", where, where & " and ( " & uftsWhere & " ) ")
    Else
        where = uftsWhere
    End If

    '    where = IIf(where = "", uftsWhere, where & " and ( " & uftsWhere & " )")
    GetQueryCondition = IIf(Len(where) > 0, " where " + where, where)
    Exit Function
End Function

'为门店使用，初始化Conn

Public Property Let conn(ByVal NewData As ADODB.Connection)
    Set m_Conn = NewData
End Property

Public Property Let bRetail(ByVal NewData As Boolean)
    m_bRetail = NewData
End Property

'ID集和 1,2,3
Public Property Get ids()
    ids = m_Ids
End Property

Private Sub Class_Initialize()
    m_Ids = ""
    If eaiLogger Is Nothing Then
        Set eaiLogger = CreateObject("EAIDebug.EAITraceLogger")
    End If
    Set SaveHandle = CreateObject("VoucherCO_Sa.ClsVoucherCO_SA")
    Set clsSystem = CreateObject("USSAServer.clsSystem")

End Sub




'------------永川项目，重取单价金额-----------------2005.12.01
'销售EAI改进 , 在header节点增加属性brecountprice, 用来标志是否重新取单价
Private Function ReGetValues(DomHeader As DOMDocument, domBody As DOMDocument, eleHead As IXMLDOMNode) As Boolean
    Dim clsDatacheck As Object   '需要后台较验

'判断是否重取
    Dim eleTemp As IXMLDOMNode
    Set eleTemp = eleHead.Attributes.getNamedItem("brecountprice")
    If Not eleTemp Is Nothing Then
        If Not (eleTemp.Text = "1") Then
            Set eleTemp = Nothing
            Exit Function
        End If
    Else
        Exit Function
    End If
    Set clsDatacheck = CreateObject("USSAServer.clsCommCheck")
    Set clsDatacheck.clsSys = clsSystem
    Set clsDatacheck.m_Conn = m_Conn
    '    Set clsSystem.dbsale = m_Conn
    Dim ele As IXMLDOMElement
    Dim DOMBodyRow As DOMDocument
    Dim DOMBodyCheck As DOMDocument
    Dim bFree1 As Boolean, bFree2 As Boolean
    Dim cfree1 As String, cfree2 As String
    Dim sKey As String

    Dim sK, sV As String
    Dim eleR As IXMLDOMElement
    Dim eleRList As IXMLDOMNodeList
    Dim dicRegetItems As New Dictionary
    With dicRegetItems
        .Add "iquotedprice", ""
        .Add "isum", ""
        .Add "imoney", ""
        .Add "itax", ""
        .Add "idiscount", ""
        .Add "inatsum", ""
        .Add "inatmoney", ""
        .Add "inattax", ""
        .Add "inatdiscount", ""
        .Add "itaxunitprice", ""
        .Add "iunitprice", ""
        .Add "kl", ""
        .Add "kl2", ""
        .Add "dkl1", ""
        .Add "dkl2", ""
        .Add "inattaxunitprice", ""
        .Add "inatunitprice", ""
        .Add "fsalecost", ""
        .Add "fsaleprice", ""
    End With

    '    If SaveHandle.clsSystemWeb.CurrentPriceType = "存货价格" Then
    '        Dim rst As New ADODB.Recordset
    '        rst.CursorLocation = adUseClient
    '        rst.Open ""
    ''        If Mid(SaveHandle.clsSystemWeb.FreePriceType, 1, 1) = "1" Then bFree1 = True
    ''        If Mid(SaveHandle.clsSystemWeb.FreePriceType, 2, 1) = "1" Then bFree2 = True
    '    End If

    Set DOMBodyRow = domBody.cloneNode(True)
    For Each ele In DOMBodyRow.selectNodes("//z:row")
        DOMBodyRow.selectSingleNode("//rs:data").removeChild ele
    Next

    For Each ele In domBody.selectNodes("//z:row")
        '        cfree1 = ele.getAttribute("cfree1")
        '        cfree2 = ele.getAttribute("cfree2")
        sKey = "cinvcode"
        '        If bFree1 Or bFree2 Then
        '            If cfree1 <> "" And bFree1 Then
        '                sKey = "cfree1"
        '            ElseIf cfree2 <> "" And bFree2 Then
        '                sKey = "cfree2"
        '            End If
        '        End If

        Set eleTemp = ele.cloneNode(True)
        DOMBodyRow.selectSingleNode("//rs:data").appendChild eleTemp
        Set DOMBodyCheck = DOMBodyRow.cloneNode(True)
        'call cellcheck '获取单价 金额等信息（需要临时构造行dom ）
        Call clsDatacheck.GetPrice(DomHeader, DOMBodyCheck, True)
        '回写单价金额信息
        Dim attr As IXMLDOMAttribute

        For Each attr In DOMBodyCheck.selectSingleNode("//z:row").Attributes
            ele.setAttribute attr.Name, attr.Value
        Next
        '        Set eleRList = DOMBodyCheck.selectNodes("//R")
        '        For Each eleR In eleRList
        '            sK = LCase(GetAttrValue(eleR, "K"))
        '            If dicRegetItems.Exists(sK) Then
        '                sV = GetAttrValue(eleR, "V")
        '                ele.setAttribute sK, sV
        '            End If
        '        Next

        DOMBodyRow.selectSingleNode("//rs:data").removeChild eleTemp

    Next
    Set clsDatacheck = Nothing
    Set DOMBodyRow = Nothing
    Set DOMBodyCheck = Nothing
End Function




'------------永川项目，自动审核单据-----------------
'销售EAI改进 , 在header节点增加属性bautoverify, 用来标志是否自动审核
Private Function AutoVerify(ByVal SaveHandle As Object, ViewHead As String, ViewBody As String, nNewID As Long, eleHead As IXMLDOMNode, sMode As String, errMsg As String)
    On Error GoTo ErrHandle
    '判断是否自动审核
    Dim bAutoVerify As Boolean
    Dim bAutoCreate As Boolean
    Dim eleTemp As IXMLDOMNode
    bAutoVerify = False
    Set eleTemp = eleHead.Attributes.getNamedItem("bautoverify")
    If eleTemp Is Nothing Then
        bAutoVerify = False
    Else
        If (eleTemp.Text = "1") Then
            bAutoVerify = True
        Else
            bAutoVerify = False
        End If
        Set eleTemp = Nothing
    End If
    Set eleTemp = eleHead.Attributes.getNamedItem("beairetailvouch")
    If eleTemp Is Nothing Then
        bAutoCreate = False
    Else
        If (eleTemp.Text = "1") Then
            bAutoCreate = True
        Else
            bAutoCreate = False
        End If
        Set eleTemp = Nothing
    End If
    If bAutoVerify = True Then
        Dim DOMHeadNew As New DOMDocument
        Dim rsHeadNew As New ADODB.Recordset
        Dim elelocal As IXMLDOMElement
        Dim strSql As String
        Dim blnSN As Boolean

        Select Case LCase(ViewHead)
        Case "sa_invpricejustmain", "sa_invpricejustmainview", "sa_cuspricejustmain", "sa_cuspricejustmainview"
            strSql = "Select * From " & ViewHead & "  Where ID=" & nNewID
        Case Else
            strSql = "Select * From " & ViewHead & "  Where DLID=" & nNewID
        End Select

        rsHeadNew.CursorLocation = adUseClient
        rsHeadNew.Open strSql, m_Conn, adOpenStatic, adLockOptimistic
        rsHeadNew.Save DOMHeadNew, adPersistXML
        rsHeadNew.Close
        Set elelocal = DOMHeadNew.selectSingleNode("//rs:data/z:row")
        Select Case LCase(ViewHead)
        Case "sa_cuspricejustmainview", "sa_cuspricejustmain"
            elelocal.setAttribute "cvouchtype", "sa19"
        Case "sa_invpricejustmainview", "sa_invpricejustmain"
            elelocal.setAttribute "cvouchtype", "sa18"
        End Select
        blnSN = False
        If Not domSN Is Nothing Then
            If domSN.documentElement.childNodes.Length >= 1 Then
                SaveHandle.clsSystemWeb.bAutoVerSaleOut = False
                blnSN = True
            End If
        End If
        If bAutoCreate Then
            elelocal.setAttribute "beairetailvouch", "1"
        End If
        If LCase(sMode) = "ar" Then m_Conn.BeginTrans
        errMsg = SaveHandle.VerifyVouch(DOMHeadNew, True)
        If LCase(sMode) = "ar" Then
            If Trim(errMsg) = "" Then
                m_Conn.CommitTrans
            Else
                m_Conn.RollbackTrans
            End If
        End If
        If blnSN Then
            Select Case LCase(ViewHead)
            Case "sa_invpricejustmain", "sa_invpricejustmainview", "sa_cuspricejustmain", "sa_cuspricejustmainview"
                strSql = "Select * From " & ViewBody & "  Where ID=" & nNewID
            Case Else
                strSql = "Select * From " & ViewBody & "  Where DLID=" & nNewID
            End Select
            rsHeadNew.Open strSql, m_Conn, adOpenStatic, adLockOptimistic
            rsHeadNew.Save DOMHeadNew, adPersistXML
            errMsg = SaveSN(CStr(nNewID), DOMHeadNew)
        End If
        Set rsHeadNew = Nothing
        Set DOMHeadNew = Nothing
    End If
    If Trim(errMsg) <> "" Then
        errMsg = "审核发货单 " & nNewID & " 失败：" & errMsg
    End If
    Exit Function
ErrHandle:
    errMsg = "审核发货单 " & nNewID & " 失败。"
End Function
Private Function SaveSN(strMainID As String, domBody As DOMDocument) As String
    Dim eleBody As IXMLDOMElement
    Dim lstBody As IXMLDOMNodeList
    Dim eleSN As IXMLDOMElement
    Dim lstSN As IXMLDOMNodeList
    Dim objSN As Object
    Dim strErr As String

    If Not domSN Is Nothing Then
        If domSN.documentElement.childNodes.Length >= 1 Then
            domSN.documentElement.setAttribute "vouchtype", "dispatchlist"
            domSN.documentElement.setAttribute "mainid", strMainID
            Set lstBody = domBody.selectNodes("//z:row")
            For Each eleBody In lstBody
                Set lstSN = domSN.documentElement.selectNodes("row[@irowno='" + eleBody.Attributes.getNamedItem("irowno").nodeValue + "']")
                For Each eleSN In lstSN
                    eleSN.setAttribute "saleautoid", eleBody.Attributes.getNamedItem("autoid").nodeValue
                Next
            Next
            Set objSN = CreateObject("U8SCMSTImportSN.clsImportFromEAI")
            If Not objSN Is Nothing Then
                objSN.SaveSNForSaleEAI mLogin, m_Conn, domSN, strErr
                If strErr <> "" Then
                    SaveSN = "序列号处理失败"
                End If
            End If
        End If
    End If
End Function


Public Function AddObjectReference(Caller As Object) As Boolean
    On Error GoTo AddObjectReferenceError
    Set gObjRef = Caller
    AddObjectReference = True
    Exit Function
AddObjectReferenceError:
    AddObjectReference = False
    Exit Function
End Function
Public Function DropObjectReference(Caller As Object) As Boolean
    On Error GoTo DropObjReferenceError
    If gObjRef Is Caller Then
        DropObjectReference = True
    Else
        DropObjectReference = False
    End If
    Exit Function
DropObjReferenceError:
    DropObjectReference = False
    Exit Function
End Function

''如果是门店数据，设置默认的销售类型
Private Function setDefaultValue(ByRef eleVouch As IXMLDOMElement, ViewHead As String) As Boolean
    On Error GoTo ErrHandle
    setDefaultValue = False
    Dim strSql As String
    Dim sCusCode As String
    Dim sCusCodeFld As String
    Dim sDepCodeFld As String
    Dim sWHCodeFld As String
    Dim sAddressFld As String
    Dim eleHead As IXMLDOMElement
    Dim ele As IXMLDOMElement
    Dim ndLstEntry As IXMLDOMNodeList
    Dim Vou As String

    Set eleHead = eleVouch.selectSingleNode("header")
    '获得销售类型在导入模版中的字段名
    Select Case UCase(ViewHead)
    Case "SALEBILLVOUCHZT"
        sCusCodeFld = "customercode"
        sDepCodeFld = "departmentcode"
        sWHCodeFld = "warehousecode"
        sAddressFld = "cdeliveradd"
        Vou = "saleinvoice"
    Case "SALEORDERQ"
        sCusCodeFld = "custcode"
        sDepCodeFld = "deptcode"
        sAddressFld = "sendaddress"
        Vou = "saleorder"
    Case "SALES_FHD_T", UCase("Sales_DXFH_T")
        sCusCodeFld = "custcode"
        sDepCodeFld = "deptcode"
        sWHCodeFld = "warehouse_code"
        sAddressFld = "address"    '"cdeliveradd"
        Vou = "consignment"
    End Select

    '取单据头客户编码
    sCusCode = GetElementValue(eleHead, sCusCodeFld)
    If sCusCode <> "" Then
        '取客户档案的：客户编码、是否直营专柜、主管部门、仓库
        strSql = "select ccuscode,isnull(bshop,0) as bshop,isnull(cCusDepart,N'') as ccusdepart,isnull(cCusWhCode,N'') as ccuswhcode,isnull(ccuspperson,N'') as ccuspperson from customer where ccuscode = N'" & sCusCode & "'"
        Dim rsCus As New ADODB.Recordset
        rsCus.Open strSql, m_Conn, adOpenStatic, adLockReadOnly, adCmdText
        If Not (rsCus.BOF And rsCus.EOF) Then
            If CBool(rsCus.Fields("bshop")) Or m_bRetail = True Then
                '业务员
                If eleVouch.selectSingleNode("//header/personcode").Text = "" Then
                    eleVouch.selectSingleNode("//header/personcode").Text = rsCus.Fields("ccuspperson")
                End If
                '设置客户部门
                If eleVouch.selectSingleNode("//header/" & sDepCodeFld).Text = "" Then
                    eleVouch.selectSingleNode("//header/" & sDepCodeFld).Text = rsCus.Fields("ccusdepart")
                End If
                '设置客户仓库
                If UCase(ViewHead) <> "SALEORDERQ" Then
                    Set ndLstEntry = eleVouch.selectNodes("//body/entry")
                    For Each ele In ndLstEntry
                        If rsCus.Fields("ccuswhcode") <> "" Then
                            ele.selectSingleNode(sWHCodeFld).Text = rsCus.Fields("ccuswhcode")
                        End If
                    Next
                End If
            End If
        End If
        rsCus.Close
        Set rsCus = Nothing

        '--门店上传的是caddcode，在此取出对应的地址
        If Not eleVouch.selectSingleNode("//header/" & sAddressFld) Is Nothing Then
            Dim sAddCode As String
            Dim sAddress As String
            Dim eleAddress As IXMLDOMElement

            sAddCode = GetElementValue(eleHead, "caddcode")
            strSql = "select cdeliveradd from cusdeliveradd " _
                   & "     where caddcode = N'" & sAddCode & "'" _
                   & "       and ccuscode = N'" & sCusCode & "'"
            Set rsCus = New ADODB.Recordset
            rsCus.Open strSql, m_Conn, adOpenStatic, adLockReadOnly, adCmdText
            If Not (rsCus.BOF And rsCus.EOF) Then
                Set eleAddress = eleHead.selectSingleNode("//header/" & sAddressFld)
                If Not eleAddress Is Nothing Then
                    If Trim(eleAddress.Text) = "" Then
                        eleAddress.Text = rsCus.Fields("cdeliveradd")
                    End If
                End If
            End If
            rsCus.Close
            Set rsCus = Nothing
        End If

    End If

    '--将单据日期强制转换成短日期格式
    Dim dDate As String
    If Not (eleVouch.selectSingleNode("//header/date") Is Nothing) Then
        dDate = eleVouch.selectSingleNode("//header/date").Text
        dDate = Left(dDate, Len("yyyy-mm-dd"))
        eleVouch.selectSingleNode("//header/date").Text = dDate
    End If

    '调用默认参数设置的服务
    Dim obj As Object
    Dim DOMVouch As DOMDocument
    Set DOMVouch = New DOMDocument
    DOMVouch.appendChild eleVouch
    'DOMVouch.loadXML "<?xml version='1.0' encoding='gb2312'?> " + Chr(13) + eleVouch.selectSingleNode("//" & Vou).xml
    Set obj = CreateObject("U8SCMImpDefValSet.clsSCMImpDefValSet")
    Call obj.getConnection(m_Conn)
    Call obj.setDefValtoDOM(DOMVouch, m_Conn, "SA", , Vou)    'login
    Call obj.reComputebody(DOMVouch, m_Conn, Vou)

    Set eleVouch = DOMVouch.documentElement.cloneNode(True)

    setDefaultValue = True
    Exit Function

ErrHandle:
    setDefaultValue = False
    Err.Clear
End Function



Private Function CusBshop(sCusCode As String) As Boolean
    On Error GoTo ErrHandle
    '    sCusCode = GetElementValue(eleHead, sCusCodeFld)
    Dim bRetailEAI As Boolean
    Dim strSql As String
    strSql = "select isnull(bshop,0) as bshop from customer where ccuscode = N'" & sCusCode & "'"

    Dim rsCus As New ADODB.Recordset
    rsCus.Open strSql, m_Conn, adOpenStatic, adLockReadOnly, adCmdText
    If Not (rsCus.BOF And rsCus.EOF) Then
        If rsCus.Fields("bshop") = 1 Or rsCus.Fields("bshop") = True Then
            bRetailEAI = True
        Else
            bRetailEAI = False
        End If
    Else
        bRetailEAI = False
    End If
    rsCus.Close
    Set rsCus = Nothing
    CusBshop = bRetailEAI
    Exit Function
ErrHandle:
    CusBshop = False
End Function
Private Sub Class_Terminate()
    Set eaiLogger = Nothing
    Set SaveHandle = Nothing
    Set clsSystem = Nothing
End Sub
Private Sub WriteLog(logLevel As Integer, content As String, productName As String, dllName As String, classOrModuleName As String, methodName As String)
    If Not eaiLogger Is Nothing Then
        eaiLogger.WriteLog Val(logLevel), content, productName, dllName, classOrModuleName, methodName
    End If
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''
'增加移动取销售统计表
Private Function DoQuerySaleReport(sXML As String) As String
    Dim sSel As String, sWhere As String, sGroup As String, sorder As String
    Dim strsaledate As String, strkeepdate As String, cDepcode As String, cWhCode As String, tempdbname As String
    Dim rst As New ADODB.Recordset
    Dim strSql As String
    Dim oDom As New DOMDocument
    If GetMobileCondition(sXML, sSel, sGroup, sorder, "") = False Then Exit Function
    Randomize
    tempdbname = "tempdbeai001" & CStr(CInt(Rnd(10) * 10))
    If GetMobileConditionForSaleReport(sXML, sWhere, strsaledate, strkeepdate, cDepcode, cWhCode, "") = False Then Exit Function
    If ExecuteProcSaleReport(m_Conn, tempdbname, sWhere, strsaledate, strkeepdate, cDepcode, cWhCode, "") = False Then Exit Function
    rst.CursorLocation = adUseClient
    strSql = "select " + sSel + " from tempdb.." + tempdbname + " group by " + sGroup + " " + IIf(sorder <> "", "order by " + sorder, "")
    rst.Open strSql, m_Conn, adOpenForwardOnly, adLockReadOnly
    rst.Save oDom, adPersistXML
    rst.Close
    DoQuerySaleReport = oDom.xml
    oDom.Save "c:\resultsalereport.xml"
    Set oDom = Nothing
End Function
'''''''''''''''''''''''''''''''''''''''''''''''''
'增加移动取销售现存量统计表
Private Function DoQuerySaCurStock(sXML As String) As String
    Dim sSel As String, sWhere As String, sGroup As String, sorder As String
    Dim sStr1 As String, sStr2 As String
    Dim rst As New ADODB.Recordset
    Dim strSql As String
    Dim oDom As New DOMDocument
    If GetMobileConditionCommon(sXML, sSel, sWhere, sGroup, sorder, "") = False Then Exit Function
    rst.CursorLocation = adUseClient
    sSel = Replace(sSel, "'", "''")
    sWhere = Replace(sWhere, "'", "''")
    sGroup = Replace(sGroup, "'", "''")
    sorder = Replace(sorder, "'", "''")


    If Len(sWhere) > 3800 Then
        sStr1 = Left(sWhere, 3800)
        sStr2 = Mid(sWhere, 3801)
    Else
        sStr1 = sWhere
    End If
    '调用库存存储过程 处理 现存量查询 使用库存的对照表
    strSql = "prc_SCM_GetCurrentStockForVouch_All N'SA',N'" + sSel + "',N'" + sGroup + "',N'" + sStr1 + "',N'" + sStr2 + "',N'" + sorder + "'"
    rst.Open strSql, m_Conn, adOpenForwardOnly, adLockReadOnly
    rst.Save oDom, adPersistXML
    rst.Close
    DoQuerySaCurStock = oDom.xml
    oDom.Save "c:\result.xml"
    Set oDom = Nothing
End Function

'销售统计表的选择列 分组列 派序列处理
Private Function GetMobileCondition(sXML As String, sSel As String, sGroup As String, sorder As String, errMsg As String) As Boolean
    Dim oDom As New DOMDocument
    Dim oEle As IXMLDOMElement
    Dim oDic As New Scripting.Dictionary
    Dim sNormalSel As String, sSourceSel As String, sGroupSel As String
    Dim sValue As String, sOperate As String, sLogic As String, sName As String
    Dim bSpecilWhere As Boolean
    sSel = ""
    sGroup = ""
    sorder = ""
    If oDom.loadXML(sXML) = False Then Exit Function
    '先处理group
    For Each oEle In oDom.selectNodes("//group/field")
        If Not oDic.Exists("G:" + LCase(oEle.getAttribute("name"))) Then
            oDic.Add "G:" + LCase(oEle.getAttribute("name")), LCase(oEle.getAttribute("name"))
            sGroup = sGroup + oEle.getAttribute("name") + ","
        End If
    Next
    If sGroup <> "" Then sGroup = Left(sGroup, Len(sGroup) - 1)
    '处理select 要特殊处理group中的字段以及group中没有的 select 中的重复字段由于sqlserver可以处理 先不予处理
    For Each oEle In oDom.selectNodes("//select/field")
        sNormalSel = oEle.getAttribute("name")
        sSourceSel = oEle.getAttribute("name")
        If IsNull(oEle.getAttribute("grouptype")) Then
            sGroupSel = "max"
        Else
            sGroupSel = oEle.getAttribute("grouptype")
        End If
        If sGroupSel <> "" Then
            sGroupSel = sGroupSel + "(" + sSourceSel + ") as " + sNormalSel
        Else
            sGroupSel = " max (" + sSourceSel + ") as " + sNormalSel
        End If
        If oDic.Exists("G:" + LCase(oEle.getAttribute("name"))) Then
            sSel = sSel + sNormalSel + ","
        Else
            sSel = sSel + sGroupSel + ","
        End If
    Next
    If sSel <> "" Then sSel = Left(sSel, Len(sSel) - 1)

    '处理order by
    For Each oEle In oDom.selectNodes("//order/field")
        sorder = sorder + oEle.getAttribute("name") + ","
    Next
    If sorder <> "" Then sorder = Left(sorder, Len(sorder) - 1)
    GetMobileCondition = True
End Function

'销售统计表特殊条件处理
Private Function GetMobileConditionForSaleReport(sXML As String, sWhere As String, strsaledate As String, strkeepdate As String, cDepcode As String, cWhCode As String, errMsg As String) As Boolean
    Dim oDom As New DOMDocument
    Dim oEle As IXMLDOMElement
    Dim oDic As New Scripting.Dictionary
    Dim sNormalSel As String, sGroupSel As String
    Dim sValue As String, sOperate As String, sLogic As String, sName As String
    Dim bSpecilWhere As Boolean
    sWhere = ""
    If oDom.loadXML(sXML) = False Then Exit Function
    '处理where
    For Each oEle In oDom.selectNodes("//where/field")
        sName = oEle.getAttribute("name")
        sValue = oEle.getAttribute("value")
        sOperate = oEle.getAttribute("operation")
        sLogic = oEle.getAttribute("logic")
        Select Case LCase(sName)
        Case "ddate"
            '需要处理 is null 和 is not null
            If Trim(LCase(sOperate)) = "is null" Or Trim(LCase(sOperate)) = "is not null" Then
                strsaledate = strsaledate + " " + sLogic + " " + sName + " " + sOperate
            Else
                strsaledate = strsaledate + " " + sLogic + " " + sName + " " + sOperate + "  N'" + sValue + "' "
            End If
        Case "dkeepdate"
            If Trim(LCase(sOperate)) = "is null" Or Trim(LCase(sOperate)) = "is not null" Then
                strkeepdate = strkeepdate + " " + sLogic + " " + sName + " " + sOperate
            Else
                strkeepdate = strkeepdate + " " + sLogic + " " + sName + " " + sOperate + "  N'" + sValue + "' "
            End If
        Case "cdepcode"
            cDepcode = " Like N'" + sValue + "%' "
        Case "cwhcode"
            cWhCode = sValue
        Case Else
            If sName = "(" Or sName = ")" Then
                sValue = ""
                sOperate = ""
                sLogic = ""
            End If
            If sWhere = "" Then
                sWhere = " and "
                sLogic = ""
            End If
            '需要处理 is null 和 is not null
            If Trim(LCase(sOperate)) = "is null" Or Trim(LCase(sOperate)) = "is not null" Then
                sWhere = sWhere + " " + sLogic + " " + sName + " " + sOperate
            ElseIf Trim(LCase(sOperate)) = "in" Then
                If Left(Trim(sValue), 1) = "(" Then
                    sWhere = sWhere + " " + sLogic + " " + sName + " " + sOperate + " " + sValue + " "
                Else
                    sWhere = sWhere + " " + sLogic + " " + sName + " " + sOperate + " (" + sValue + ") "
                End If
            Else
                sWhere = sWhere + " " + sLogic + " " + sName + " " + sOperate + "  N'" + sValue + "' "
            End If
        End Select

    Next
    GetMobileConditionForSaleReport = True
End Function
'执行销售统计表存储过程
Private Function ExecuteProcSaleReport(dbconn As ADODB.Connection, tempdbname As String, sWhere As String, strsaledate As String, strkeepdate As String, cDepcode As String, cWhCode As String, errMsg As String) As Boolean
    Dim adocomm As ADODB.Command
    Dim adopara As ADODB.Parameter
    Set adocomm = New ADODB.Command
    If Len(sWhere) - 5 > 0 Then sWhere = Right(sWhere, Len(sWhere) - 5)
    If Len(strsaledate) - 5 > 0 Then strsaledate = Right(strsaledate, Len(strsaledate) - 5)
    If Len(strkeepdate) - 5 > 0 Then strkeepdate = Right(strkeepdate, Len(strkeepdate) - 5)
    strkeepdate = IIf(strkeepdate = "", " ", strkeepdate)
    cDepcode = IIf(cDepcode = "", " ", cDepcode)
    cWhCode = IIf(cWhCode = "", " ", cWhCode)

    Set adopara = adocomm.CreateParameter("chrtable", adVarWChar, adParamInput _
                , LenB(tempdbname), tempdbname)
    adocomm.Parameters.Append adopara

    '一般条件参数
    Set adopara = adocomm.CreateParameter("chrWhere1", adLongVarWChar, adParamInput, LenB(sWhere), sWhere)
    adocomm.Parameters.Append adopara
    Set adopara = Nothing
    '开票日期参数
    Set adopara = adocomm.CreateParameter("chrsaledate", adLongVarWChar, adParamInput, LenB(strsaledate), strsaledate)
    adocomm.Parameters.Append adopara
    Set adopara = Nothing
    '结算日期参数
    Set adopara = adocomm.CreateParameter("chrkeepdate", adLongVarWChar, adParamInput, LenB(strkeepdate), strkeepdate)
    adocomm.Parameters.Append adopara
    Set adopara = Nothing
    '部门条件
    Set adopara = adocomm.CreateParameter("chrdep", adLongVarWChar, adParamInput, LenB(cDepcode), cDepcode)
    adocomm.Parameters.Append adopara
    Set adopara = Nothing
    '复合条件
    Set adopara = adocomm.CreateParameter("chrchecker", adLongVarWChar, adParamInput, LenB(" "), " ")
    adocomm.Parameters.Append adopara
    Set adopara = Nothing
    '仓库条件
    Set adopara = adocomm.CreateParameter("chrwarehouse", adLongVarWChar, adParamInput, LenB(cWhCode), cWhCode)
    adocomm.Parameters.Append adopara
    Set adopara = Nothing
    '增加操作员权限控制
    Set adopara = adocomm.CreateParameter("chrAuth", adVarWChar, adParamInput, LenB(" "), " ")
    adocomm.Parameters.Append adopara
    Set adopara = Nothing
    '增加业务项目的权限控制

    Set adopara = adocomm.CreateParameter("chrRecordAuth", adVarWChar, adParamInput, LenB(" "), " ")
    adocomm.Parameters.Append adopara
    Set adopara = Nothing

    Set adopara = adocomm.CreateParameter("intExpress", adInteger, adParamInput, 3, 0)
    adocomm.Parameters.Append adopara
    Set adopara = Nothing

    '执行存储过程
    Set adocomm.ActiveConnection = dbconn
    adocomm.CommandText = "Sa_MoveSaleDetail"
    adocomm.CommandType = adCmdStoredProc
    adocomm.CommandTimeout = 0
    adocomm.Execute
    Set adocomm = Nothing
    Set adopara = Nothing
    Set adopara = Nothing
    ExecuteProcSaleReport = True
End Function

'一般报表的选择分组 条件 排序处理
Private Function GetMobileConditionCommon(sXML As String, sSel As String, sWhere As String, sGroup As String, sorder As String, errMsg As String) As Boolean
    Dim oDom As New DOMDocument
    Dim oEle As IXMLDOMElement
    Dim oDic As New Scripting.Dictionary
    Dim sNormalSel As String, sGroupSel As String
    Dim sValue As String, sOperate As String, sLogic As String, sName As String
    Dim bSpecilWhere As Boolean
    sSel = ""
    sWhere = ""
    sGroup = ""
    sorder = ""
    If oDom.loadXML(sXML) = False Then Exit Function
    '先处理group
    For Each oEle In oDom.selectNodes("//group/field")
        If Not oDic.Exists("G:" + LCase(oEle.getAttribute("name"))) Then
            oDic.Add "G:" + LCase(oEle.getAttribute("name")), LCase(oEle.getAttribute("name"))
            sGroup = sGroup + oEle.getAttribute("name") + ","
        End If
    Next
    If sGroup <> "" Then sGroup = Left(sGroup, Len(sGroup) - 1)
    '处理select 要特殊处理group中的字段以及group中没有的 select 中的重复字段由于sqlserver可以处理 先不予处理
    For Each oEle In oDom.selectNodes("//select/field")
        sNormalSel = oEle.getAttribute("name")
        If IsNull(oEle.getAttribute("grouptype")) Then
            sGroupSel = "max"
        Else
            sGroupSel = oEle.getAttribute("grouptype")
        End If
        If sGroupSel <> "" Then
            sGroupSel = sGroupSel + "(" + sNormalSel + ") as " + sNormalSel
        Else
            sGroupSel = " max (" + sNormalSel + ") as " + sNormalSel
        End If
        If oDic.Exists("G:" + LCase(oEle.getAttribute("name"))) Then
            sSel = sSel + sNormalSel + ","
        Else
            sSel = sSel + sGroupSel + ","
        End If
    Next
    If sSel <> "" Then sSel = Left(sSel, Len(sSel) - 1)
    '处理where
    For Each oEle In oDom.selectNodes("//where/field")
        sName = oEle.getAttribute("name")
        sValue = oEle.getAttribute("value")
        sOperate = oEle.getAttribute("operation")
        sLogic = oEle.getAttribute("logic")
        bSpecilWhere = False
        If Not IsNull(oEle.getAttribute("bspecial")) Then
            bSpecilWhere = CBool(oEle.getAttribute("bspecial"))
        End If
        If sName = "(" Or sName = ")" Then
            sValue = ""
            sOperate = ""
            sLogic = ""
        End If
        If sWhere = "" Then
            sWhere = " and "
            sLogic = ""
        End If
        '需要处理 is null 和 is not null
        If Trim(LCase(sOperate)) = "is null" Or Trim(LCase(sOperate)) = "is not null" Or bSpecilWhere Then
            sWhere = sWhere + " " + sLogic + " " + sName + " " + sOperate
        ElseIf Trim(LCase(sOperate)) = "in" Then
            If Left(Trim(sValue), 1) = "(" Then
                sWhere = sWhere + " " + sLogic + " " + sName + " " + sOperate + " " + sValue + " "
            Else
                sWhere = sWhere + " " + sLogic + " " + sName + " " + sOperate + " (" + sValue + ") "
            End If
        Else
            sWhere = sWhere + " " + sLogic + " " + sName + " " + sOperate + "  N'" + sValue + "' "
        End If
    Next
    '处理order by
    For Each oEle In oDom.selectNodes("//order/field")
        sorder = sorder + oEle.getAttribute("name") + ","
    Next
    If sorder <> "" Then sorder = Left(sorder, Len(sorder) - 1)
    GetMobileConditionCommon = True
End Function
'''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''
'增加移动取销售综合统计表
Private Function DoQuerySaleReportcolligate(sXML As String, strsender As String) As String
    Dim sSel As String, sWhere As String, sWhere1 As String, sGroup As String, sorder As String
    Dim strStartDate As String, StrEndDate As String, strVouchType As String, strBack As String, strChecker As String, strWare As String, strbz As String, strItemCode As String, tempdbname As String
    Dim rst As New ADODB.Recordset
    Dim strSql As String
    Dim oDom As New DOMDocument
    If m_Conn.State = 0 Then m_Conn.Open mLogin.UfDbName
    If GetMobileCondition(sXML, sSel, sGroup, sorder, "") = False Then Exit Function
    Randomize
    tempdbname = "tempdbeai001" & CStr(CInt(Rnd(10) * 10))
    If GetMobileConditionForSaleReportcoll(sXML, sWhere, sWhere1, strStartDate, StrEndDate, strVouchType, strBack, strChecker, strWare, strbz, strItemCode, "") = False Then Exit Function
    If ExecuteProcSaleReportcoll(m_Conn, tempdbname, sWhere, sWhere1, strStartDate, StrEndDate, strVouchType, strBack, strChecker, strWare, strbz, strItemCode, "") = False Then Exit Function
    rst.CursorLocation = adUseClient
    strSql = "select " + sSel + " from tempdb.." + tempdbname + " group by " + sGroup + " " + IIf(sorder <> "", "order by " + sorder, "")
    rst.Open strSql, m_Conn, adOpenForwardOnly, adLockReadOnly
    rst.Save oDom, adPersistXML
    rst.Close

    DoQuerySaleReportcolligate = "<?xml version='1.0' encoding='UTF-8' ?> " + vbCrLf
    DoQuerySaleReportcolligate = DoQuerySaleReportcolligate + "<ufinterface codeexchanged='n' display='' docid='' exportneedexch='n'" + vbCrLf
    DoQuerySaleReportcolligate = DoQuerySaleReportcolligate + " family='' proc='Query' receiver='u8' roottag='salereportcolligate' " + vbCrLf
    DoQuerySaleReportcolligate = DoQuerySaleReportcolligate + "sender='" + strsender + "' paginate='0'> " + vbCrLf
    DoQuerySaleReportcolligate = DoQuerySaleReportcolligate + "<salereportcolligate>" + vbCrLf
    DoQuerySaleReportcolligate = DoQuerySaleReportcolligate + oDom.xml + vbCrLf
    DoQuerySaleReportcolligate = DoQuerySaleReportcolligate + "  </salereportcolligate></ufinterface>"
    '    oDom.Save "c:\resultsalereportcolligate.xml"
    Set oDom = Nothing
End Function


'销售综合统计表特殊条件处理
Private Function GetMobileConditionForSaleReportcoll(sXML As String, sWhere As String, sWhere1 As String, strStartDate As String, StrEndDate As String, _
                                                     strVouchType As String, strBack As String, strChecker As String, strWare As String, strbz As String, strItemCode As String, errMsg As String) As Boolean
    Dim oDom As New DOMDocument
    Dim oEle As IXMLDOMElement
    Dim oDic As New Scripting.Dictionary
    Dim sNormalSel As String, sGroupSel As String
    Dim sValue As String, sOperate As String, sLogic As String, sName As String
    Dim bSpecilWhere As Boolean
    sWhere = ""
    If oDom.loadXML(sXML) = False Then Exit Function
    '处理where
    For Each oEle In oDom.selectNodes("//where/field")
        sName = oEle.getAttribute("name")
        sValue = oEle.getAttribute("value")
        sOperate = oEle.getAttribute("operation")
        sLogic = oEle.getAttribute("logic")
        Select Case LCase(sName)
        Case "ddate1"
            If Trim(sValue) <> vbNullString Then
                strStartDate = sValue
            End If
        Case "ddate2"
            If Trim(sValue) <> vbNullString Then
                StrEndDate = sValue
            End If
        Case "单据类型"
            If Trim(sValue) <> vbNullString Then
                strVouchType = "N'" + sValue + "'"
            End If
        Case "是否退货"
            If Trim(sValue) <> vbNullString Then
                strBack = Trim(sValue)
            End If
        Case "审核"
            If Trim(sValue) <> vbNullString Then
                strChecker = Trim(sValue)
            End If
        Case "cwhcode"
            If Trim(sValue) <> vbNullString Then
                'strWare = "(" + sName + "=N'" + sValue + "')"

                If Trim(LCase(sOperate)) = "in" Or Trim(LCase(sOperate)) = "not in" Then
                    If Left(Trim(sValue), 1) = "(" Then
                        strWare = sName + " " + sOperate + " " + sValue + " "
                    Else
                        strWare = sName + " " + sOperate + " (" + sValue + ") "
                    End If
                Else
                    strWare = "(" + sName + " " + sOperate + " N'" + sValue + "')"
                End If
            End If
        Case "cexch_name"       '币种
            If Trim(sValue) <> vbNullString Then
                strbz = "(" + sName + "=N'" + sValue + "')"
            End If
        Case "citemcode"
            If Trim(sValue) <> vbNullString Then
                strItemCode = "(" + sName + "=N'" + sValue + "')"
            End If
        Case Else
            If sName = "(" Or sName = ")" Then
                sValue = ""
                sOperate = ""
                sLogic = ""
            End If
            If sWhere = "" Then
                sWhere = " and "
                sLogic = ""
            End If

            '需要处理 is null 和 is not null
            If Trim(LCase(sOperate)) = "is null" Or Trim(LCase(sOperate)) = "is not null" Then
                sWhere = sWhere + " " + sLogic + " " + sName + " " + sOperate
            ElseIf Trim(LCase(sOperate)) = "in" Or Trim(LCase(sOperate)) = "not in" Then
                If Left(Trim(sValue), 1) = "(" Then
                    sWhere = sWhere + " " + sLogic + " " + sName + " " + sOperate + " " + sValue + " "
                Else
                    sWhere = sWhere + " " + sLogic + " " + sName + " " + sOperate + " (" + sValue + ") "
                End If
            Else
                sWhere = sWhere + " " + sLogic + " " + sName + " " + sOperate + "  N'" + sValue + "' "
            End If
            Select Case LCase(sName)
            Case "cfree1", "cfree2", "cfree3", "cfree4", "cfree5", "cfree6", "cfree7", "cfree8", "cfree9", "cfree10", _
                 "cdefine22", "cdefine23", "cdefine24", "cdefine25", "cdefine26", "cdefine27", "cstcode", "cexch_name", "cbustype"
                'AP_DETAIL表不支持的查询字段
            Case Else
                If Trim(LCase(sOperate)) = "is null" Or Trim(LCase(sOperate)) = "is not null" Then
                    sWhere1 = sWhere1 + " " + sLogic + " " + sName + " " + sOperate
                ElseIf Trim(LCase(sOperate)) = "in" Or Trim(LCase(sOperate)) = "not in" Then
                    If Left(Trim(sValue), 1) = "(" Then
                        sWhere1 = sWhere1 + " " + sLogic + " " + sName + " " + sOperate + " " + sValue + " "
                    Else
                        sWhere1 = sWhere1 + " " + sLogic + " " + sName + " " + sOperate + " (" + sValue + ") "
                    End If
                Else
                    sWhere1 = sWhere1 + " " + sLogic + " " + sName + " " + sOperate + "  N'" + sValue + "' "
                End If
            End Select
        End Select

    Next
    If Trim(sWhere1) <> vbNullString Then

        '替换业务员字段名
        sWhere1 = Replace(sWhere1, "cPersonCode", "cPerson")
        '替换部门字段名称
        sWhere1 = Replace(sWhere1, "cDepCode", "cDeptCode")
    End If
    GetMobileConditionForSaleReportcoll = True
End Function
'执行销售综合统计表存储过程
Private Function ExecuteProcSaleReportcoll(dbconn As ADODB.Connection, tempdbname As String, sWhere As String, sWhere1 As String, strStartDate As String, StrEndDate As String, _
                                           strVouchType As String, strBack As String, strChecker As String, strWare As String, strbz As String, strItemCode As String, errMsg As String) As Boolean
    Dim adocomm As ADODB.Command
    Dim adopara1 As ADODB.Parameter
    Set adocomm = New ADODB.Command
    If Len(sWhere) - 5 > 0 Then sWhere = Right(sWhere, Len(sWhere) - 5)
    If sWhere = "" Then sWhere = " "
    If sWhere1 = "" Then sWhere1 = " "
    If strStartDate = "" Then strStartDate = " "
    If StrEndDate = "" Then StrEndDate = " "
    If strVouchType = "" Then strVouchType = " "
    If strBack = "" Then strBack = " "
    If strChecker = "" Then strChecker = " "
    If strWare = "" Then strWare = " "
    If strbz = "" Then strbz = " "
    If strItemCode = "" Then strItemCode = " "
    '确认临时表参数
    Set adopara1 = adocomm.CreateParameter("chrTable", adVarWChar, adParamInput _
                 , LenB(tempdbname), tempdbname)
    adocomm.Parameters.Append adopara1
    '追加参数

    '开始日期参数
    Set adopara1 = adocomm.CreateParameter("chrStartDate", adVarWChar, adParamInput, LenB(strStartDate), strStartDate)
    adocomm.Parameters.Append adopara1
    Set adopara1 = Nothing
    '结束日期参数
    Set adopara1 = adocomm.CreateParameter("chrEndDate", adVarWChar, adParamInput, LenB(StrEndDate), StrEndDate)
    adocomm.Parameters.Append adopara1
    Set adopara1 = Nothing
    '其他条件
    Set adopara1 = adocomm.CreateParameter("chrWhere", adVarWChar, adParamInput, LenB(sWhere), sWhere)
    adocomm.Parameters.Append adopara1
    Set adopara1 = Nothing
    '退货条件
    Set adopara1 = adocomm.CreateParameter("chrBack", adVarWChar, adParamInput, LenB(strBack), strBack)
    adocomm.Parameters.Append adopara1
    Set adopara1 = Nothing
    '审核条件
    Set adopara1 = adocomm.CreateParameter("chrChecker", adVarWChar, adParamInput, LenB(strChecker), strChecker)
    adocomm.Parameters.Append adopara1
    Set adopara1 = Nothing
    '单据条件
    Set adopara1 = adocomm.CreateParameter("chrVouchType", adVarWChar, adParamInput, LenB(strVouchType), strVouchType)
    adocomm.Parameters.Append adopara1
    Set adopara1 = Nothing
    '追加附加条件
    Set adopara1 = adocomm.CreateParameter("chrWhere1", adVarWChar, adParamInput, LenB(sWhere1), sWhere1)
    adocomm.Parameters.Append adopara1
    Set adopara1 = Nothing
    '追加仓库条件
    Set adopara1 = adocomm.CreateParameter("chrWare", adVarWChar, adParamInput, LenB(strWare), strWare)
    adocomm.Parameters.Append adopara1
    Set adopara1 = Nothing
    '增加操作员权限控制

    Set adopara1 = adocomm.CreateParameter("chrAuth", adVarWChar, adParamInput, LenB(" "), " ")
    adocomm.Parameters.Append adopara1

    Set adopara1 = Nothing
    Set adopara1 = adocomm.CreateParameter("chrBz", adVarWChar, adParamInput, LenB(strbz), strbz)
    adocomm.Parameters.Append adopara1
    Set adopara1 = Nothing
    '增加业务项目的权限控制
    Set adopara1 = adocomm.CreateParameter("chrRecordAuth", adVarWChar, adParamInput, LenB(" "), " ")
    adocomm.Parameters.Append adopara1
    Set adopara1 = Nothing

    '增加项目条件控制
    Set adopara1 = adocomm.CreateParameter("chrItemCode", adVarWChar, adParamInput, LenB(strItemCode), strItemCode)
    adocomm.Parameters.Append adopara1
    Set adopara1 = Nothing

    '执行存储过程
    Set adocomm.ActiveConnection = dbconn
    adocomm.CommandText = "Sa_MoveSaleColligate"
    adocomm.CommandType = adCmdStoredProc
    adocomm.CommandTimeout = 0
    adocomm.Execute
    Set adocomm = Nothing
    Set adopara1 = Nothing
    ExecuteProcSaleReportcoll = True
End Function





