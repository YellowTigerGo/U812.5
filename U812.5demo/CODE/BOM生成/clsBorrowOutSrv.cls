VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSrv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit

Private Const AuthBrowse = "ST02JC020101"                  '浏览
Private Const AuthAdd = "ST02JC020102"                     '新增
Private Const AuthModify = "ST02JC020103"                  '修改
Private Const AuthDelete = "ST02JC020104"                  '删除
Private Const AuthVerify = "ST02JC020105"                  '审核
Private Const AuthUnVerify = "ST02JC020106"                '弃审
Private Const AuthOpen = "ST02JC020107"                    '打开
Private Const AuthClose = "ST02JC020108"                   '关闭
Private Const AuthPrint = "ST02JC020109"                   '打印
Private Const AuthOut = "ST02JC020110"                     '输出
Private Const AuthModifyWF = "ST02JC020111"                '审批流中修改单据 20110817 by zhangwchb



' 审批流
Dim iverifystate As Integer
Dim ufts As String
Dim IsWFControlled As Boolean                              '是否审批流控制
Dim vouchercode As String                                  '单据号
Dim ireturncount As Integer                                '单据被退回次数
Dim bststarted As Boolean                                  '库存是否启用



Public Function Init(ologin As U8Login.clsLogin) As Boolean
        
     On Error GoTo ErrHandle
        If Not g_oLogin Is Nothing Then Set g_oLogin = Nothing
    
        Set g_oLogin = ologin
        g_oLogin.login "ST"
        If g_Conn Is Nothing Then
           Set g_Conn = New ADODB.Connection
           g_Conn.CursorLocation = adUseClient
        End If
        
        If g_Conn.State = 1 Then g_Conn.Close
        
        g_Conn.ConnectionString = g_oLogin.UfDbName
        g_Conn.CursorLocation = adUseClient
        g_Conn.CommandTimeout = 0
        g_Conn.Open

        g_bLogined = True
        
        GlobalInit ologin
        
        bststarted = GetSTStartDate("")
        
        Init = True
        Exit Function
        
ErrHandle:
        Init = False
    
End Function


Public Function headCellCheck(ByVal sFieldName As String, ByRef oDomHead As DOMDocument, ByRef oDomBody As DOMDocument, ByRef sRetValue As String, _
                              Optional ByRef conn As ADODB.Connection, Optional ByRef sErrMsg As String = "") As Boolean
    
    If Not conn Is Nothing Then
       Set g_Conn = conn
    End If
    
    sRetValue = CCrmStr(sRetValue)
    
    headCellCheck = HandRecord_T(sFieldName, oDomHead, oDomBody, sRetValue, sErrMsg)
    
End Function

'手工录入表头字段校验
'部门、仓库、业务员、客户
Private Function HandRecord_T(ByVal sFieldName As String, ByRef oDomHead As DOMDocument, ByRef oDomBody As DOMDocument, ByRef retvalue As String, _
                              Optional ByRef sErrMsg As String = "") As Boolean
    On Error GoTo ErrHandler

    Dim strValue As Boolean
    Dim sql As String
    Dim sMetaItemName As String
    Dim cvencode As String
    Dim cvenname As String
    
    HandRecord_T = False

    sMetaItemName = LCase(sFieldName)
    '手工输入时,需要校验是否存在

    '/*B*/ 根据单据模板表头设置确定是否需要以下栏目
    Select Case sMetaItemName
        Case LCase("cfreightCost")
            If Trim(GetDomValue(oDomHead, "cfreightCost")) <> "" Then
                If IsNumeric(GetDomValue(oDomHead, "cfreightCost")) Then
                    If CDbl(GetDomValue(oDomHead, "cfreightCost")) <= 0 Then
                        sErrMsg = GetString("U8.DZ.JA.Res600")
                        Exit Function
                    End If
                End If
            End If
            '运费否
        Case LCase("cfreight")
            If retvalue = "否" Then
                SetDomValue oDomHead, "cfreightCompany", ""
                SetDomValue oDomHead, "cfreightType", ""
                SetDomValue oDomHead, "cfreightCost", ""
                SetDomValue oDomHead, "MycdefineT2", ""
'
'                Voucher.EnableHead "cfreightCompany", False
'                Voucher.EnableHead "cfreightType", False
'                Voucher.EnableHead "cfreightCost", False
'                Voucher.EnableHead "MycdefineT2", False
'                Voucher.EnableHead "MycdefineT6", False
'                Voucher.SetCurrentRow ("@AutoID=" & bodyele.getAttribute("AutoID") & "")
            ElseIf retvalue = "是" Then
'                Voucher.EnableHead "cfreightCompany", True
'                Voucher.EnableHead "cfreightType", True
'                Voucher.EnableHead "cfreightCost", True
'                Voucher.EnableHead "MycdefineT2", True
'                Voucher.EnableHead "MycdefineT6", True
            Else
                sErrMsg = "输入数据不合法"
                Exit Function
            End If
            '关联单据类型
        Case LCase("cAboutVoucher")
            If retvalue = "销售出库单" Or retvalue = "服务单" Or retvalue = "委外出库单" Then
            Else
                sErrMsg = "输入数据不合法"
                Exit Function
            End If
            '类型
        Case LCase("cType")
            If retvalue = "客户" Or retvalue = "供应商" Or retvalue = "部门" Then    ' Or retvalue = "人员"
                SetDomValue oDomHead, "bObjectCode", ""
                SetDomValue oDomHead, "bObjectName", ""
            Else
                sErrMsg = "输入数据不合法"
                Exit Function
            End If
            '单位
        Case LCase("bObjectCode"), LCase("bObjectName")
            If HandRecord_TbObjectCode(sFieldName, oDomHead, oDomBody, retvalue, sErrMsg) = False Then
                Exit Function
            End If
            '部门
        Case "cdepcode", "cdepname"

            If retvalue = "" Then
                SetDomValue oDomHead, "cdepcode", ""
                SetDomValue oDomHead, "cdepname", ""
                SetDomValue oDomHead, "cpersoncode", ""
                SetDomValue oDomHead, "cpersonname", ""
            Else
                sql = "select cdepcode code,cdepname name from department where (cdepcode='" & retvalue & "' or cdepname='" & retvalue & "') and isnull(dDepEndDate,'9999-12-31')>N'" & g_oLogin.CurDate & "' "
                '                     '权限
                '                    Authstr = GetRowAuthAlls(g_Conn, g_oLogin.cUserId, "R", "bAuth_Dep")
                '                    If Authstr <> "" Then sql = sql & Authstr
                sql = sql & IIf(sAuth_depW = "", "", " and cdepcode in (" & sAuth_depW & ")")

                strValue = CheckCellValue(sql)

                If strValue = False Then
                    ReDim varArgs(0)
                    varArgs(0) = retvalue
                    sErrMsg = GetStringPara("U8.DZ.JA.Res560", varArgs(0))

                    'MsgBox "部门" & RetValue & "不存在或者没有权限,请重新输入!", vbInformation, GetString("U8.DZ.JA.Res030")
                 
                    Exit Function
                Else
                    SetDomValue oDomHead, "cdepcode", strCellCode
                    SetDomValue oDomHead, "cdepname", strCellName
                    If sMetaItemName Like "*code" Then
                        retvalue = strCellCode
                    Else
                        retvalue = strCellName
                    End If
                End If

            End If



            '仓库
        Case "cwhcode", "cwhname"

            If retvalue = "" Then
                SetDomValue oDomHead, "cwhcode", ""
                SetDomValue oDomHead, "cwhname", ""
            Else

                sql = "select cwhcode code,cWhName name from warehouse where (cwhcode='" & retvalue & "' or cWhname='" & retvalue & "') and (datediff(d, '" & CDate(Mid(GetDomValue(oDomHead, StrdDate), 1, 10)) & "',isnull(dWhEndDate ,'') ) >0 or isnull(dWhEndDate ,'')='')"
                sql = sql & IIf(sAuth_WareHouseW = "", "", " and cwhcode in (" & sAuth_WareHouseW & ")")

                '                     '权限
                '                    Authstr = GetRowAuthAlls(g_Conn, g_oLogin.cUserId, "R", "bAuth_Wh")
                '                    If Authstr <> "" Then sql = sql & Authstr

                strValue = CheckCellValue(sql)

                If strValue = False Then
                    ReDim varArgs(0)
                    varArgs(0) = retvalue
                    sErrMsg = GetStringPara("U8.DZ.JA.Res610", varArgs(0))

                    'MsgBox "仓库" & RetValue & "不存在或者没有权限,请重新输入!", vbInformation, GetString("U8.DZ.JA.Res030")
                   
                    Exit Function
                Else
                    SetDomValue oDomHead, "cwhcode", strCellCode
                    SetDomValue oDomHead, "cwhname", strCellName
                    If sMetaItemName Like "*code" Then
                        retvalue = strCellCode
                    Else
                        retvalue = strCellName
                    End If
                End If


            End If

            '发运方式 select cSCCode,cSCName from ShippingChoice
        Case LCase("cfreightType"), LCase("MycdefineT6")
            '            Voucher.headerText("MycdefineT6") = referpara.rstGrid.Fields("cSCName")
            '            retvalue = referpara.rstGrid.Fields("cSCCode")
            If retvalue = "" Then
                SetDomValue oDomHead, "cfreightType", ""
                SetDomValue oDomHead, "MycdefineT6", ""
            Else
                sql = "select cSCCode AS code,cSCName AS Name from ShippingChoice where (cSCCode='" & retvalue & "' or cSCName='" & retvalue & "')"
                strValue = CheckCellValue(sql)

                If strValue = False Then
                    ReDim varArgs(0)
                    varArgs(0) = retvalue
                    sErrMsg = GetStringPara("U8.DZ.JA.Res620", varArgs(0))

                    'MsgBox "发运方式" & RetValue & "不存在,请重新输入!", vbInformation, GetString("U8.DZ.JA.Res030")
                    
                    Exit Function
                Else
                    SetDomValue oDomHead, "cfreightType", strCellCode
                    SetDomValue oDomHead, "MycdefineT6", strCellName
                    If sMetaItemName Like "*code" Then
                        retvalue = strCellCode
                    Else
                        retvalue = strCellName
                    End If
                End If
            End If

            '业务员
        Case "cpersoncode", "cpersonname"

            If retvalue = "" Then
                SetDomValue oDomHead, "cpersoncode", ""
                SetDomValue oDomHead, "cpersonname", ""
            Else

                sql = "select cPersonCode code,cPersonName name from person where (cPersonCode='" & retvalue & "' or cPersonName='" & retvalue & "') and  '" & g_oLogin.CurDate & "' between dPValidDate  and IsNull(dPInValidDate,'2099-12-31')  "
                sql = sql & " and (" & IIf(GetDomValue(oDomHead, "cdepcode") = "", "1=1", " cdepcode='" & GetDomValue(oDomHead, "cdepcode") & "' ") & ") and dPValidDate <='" & g_oLogin.CurDate & "' and  isnull(dPInValidDate ,'2099-12-31') >='" & g_oLogin.CurDate & "' "

                '权限
                '                    Authstr = GetRowAuthAlls(g_Conn, g_oLogin.cUserId, "R", "bAuth_Per")
                '                    If Authstr <> "" Then sql = sql & Authstr

                strValue = CheckCellValue(sql)

                If strValue = False Then
                    ReDim varArgs(0)
                    varArgs(0) = retvalue
                    sErrMsg = GetStringPara("U8.DZ.JA.Res570", varArgs(0))

                    'MsgBox "业务员" & RetValue & "不存在或者没有权限,请重新输入!", vbInformation, GetString("U8.DZ.JA.Res030")
                  
                    Exit Function
                Else
                    SetDomValue oDomHead, "cpersoncode", strCellCode
                    SetDomValue oDomHead, "cpersonname", strCellName
                    If sMetaItemName Like "*code" Then
                        retvalue = strCellCode
                    Else
                        retvalue = strCellName
                    End If
                End If
            End If
            '经办人
        Case "cpersoncode2"
            If retvalue = "" Then
                SetDomValue oDomHead, "cpersoncode2", ""
                SetDomValue oDomHead, "cpersonname2", ""
            Else

                sql = "select cPersonCode code,cPersonName name from person where (cPersonCode='" & retvalue & "' or cPersonName='" & retvalue & "') and  '" & g_oLogin.CurDate & "' between dPValidDate  and IsNull(dPInValidDate,'2099-12-31') "

                '权限
                '                    Authstr = GetRowAuthAlls(g_Conn, g_oLogin.cUserId, "R", "bAuth_Per")
                '                    If Authstr <> "" Then sql = sql & Authstr

                strValue = CheckCellValue(sql)

                If strValue = False Then
                    ReDim varArgs(0)
                    varArgs(0) = retvalue
                    sErrMsg = GetStringPara("U8.DZ.JA.Res570", varArgs(0))

                    'MsgBox "业务员" & RetValue & "不存在或者没有权限,请重新输入!", vbInformation, GetString("U8.DZ.JA.Res030")
                    'bChanged = Cancel
                    Exit Function
                Else
                    SetDomValue oDomHead, "cpersoncode2", strCellCode
                    SetDomValue oDomHead, "cpersonname2", strCellName
                    If sMetaItemName Like "*code" Then
                        retvalue = strCellCode
                    Else
                        retvalue = strCellName
                    End If
                End If
            End If

            '客户
        Case "ccuscode", "ccusabbname"
            Dim cName As String
            Dim Address As String

            If retvalue = "" Then
                SetDomValue oDomHead, "ccuscode", ""
                SetDomValue oDomHead, "ccusname", ""
                SetDomValue oDomHead, "ccusabbname", ""
                SetDomValue oDomHead, "ccusaddress", ""
            Else

                sql = "select cCusCode ,cCusAbbName,cCusName,cCusAddress  from Customer where (cCusCode='" & retvalue & "' or cCusAbbName='" & retvalue & "' or cCusMnemCode='" & retvalue & "' or cCusAbbName ='" & retvalue & "' )"
                sql = sql & IIf(sAuth_CusW = "", "", " and cCusCode in (" & sAuth_CusW & ")")

                '权限
                '                    Authstr = GetRowAuthAlls(g_Conn, g_oLogin.cUserId, "R", "bAuth_Cus")
                '                    If Authstr <> "" Then sql = sql & Authstr

                strValue = CheckCustomer(sql, cName, Address)

                If strValue = False Then
                    ReDim varArgs(0)
                    varArgs(0) = retvalue
                    sErrMsg = GetStringPara("U8.DZ.JA.Res580", varArgs(0))

                    'MsgBox "客户" & RetValue & "不存在或者没有权限,请重新输入!", vbInformation, GetString("U8.DZ.JA.Res030")
                    'bChanged = Cancel
                    Exit Function
                Else
                    SetDomValue oDomHead, "ccuscode", strCellCode
                    SetDomValue oDomHead, "ccusname", cName
                    SetDomValue oDomHead, "ccusabbname", strCellName
                    SetDomValue oDomHead, "ccusaddress", Address
                    If sMetaItemName Like "*code" Then
                        retvalue = strCellCode
                    Else
                        retvalue = strCellName
                    End If

                End If
            End If
        Case "cvencode", "cvenname"
            '            Dim cvencode As String
            '            Dim cvenname As String
            '
            If retvalue = "" Then
                SetDomValue oDomHead, "cvencode", ""
                SetDomValue oDomHead, "cvenname", ""
            Else
                sql = "select cvencode ,cvenneme,cvenabbname  from vendor where (cvencode='" & retvalue & "' or cvenname='" & retvalue & "' or cVenMnemCode ='" & retvalue & "' or cVenAbbName ='" & retvalue & "')"
                sql = sql & IIf(sAuth_vendorW = "", "", " and iid in (" & sAuth_vendorW & ")")

                strValue = CheckVendor(sql, cvencode, cvenname)

                If strValue = False Then
                    ReDim varArgs(0)
                    varArgs(0) = retvalue
                    sErrMsg = GetStringPara("U8.DZ.JA.Res590", varArgs(0))

                    'MsgBox "供应商" & RetValue & "不存在或者没有权限,请重新输入!", vbInformation, GetString("U8.DZ.JA.Res030")
                    'bChanged = Cancel
                    Exit Function
                Else
                    SetDomValue oDomHead, "cvencode", cvencode
                    SetDomValue oDomHead, "cvenname", cvenname

                    If sMetaItemName Like "*code" Then
                        retvalue = strCellCode
                    Else
                        retvalue = strCellName
                    End If

                End If
            End If

            '发运单位
        Case LCase("cfreightCompany"), LCase("MycdefineT2")


            If retvalue = "" Then
                SetDomValue oDomHead, "cfreightCompany", ""
                SetDomValue oDomHead, "MycdefineT2", ""
            Else
                sql = "select cvencode ,cvenname,cvenabbname  from vendor where (cvencode='" & retvalue & "' or cvenname='" & retvalue & "')"
                sql = sql & IIf(sAuth_vendorW = "", "", " and iid in (" & sAuth_vendorW & ")")

                strValue = CheckVendor(sql, cvencode, cvenname)

                If strValue = False Then
                    ReDim varArgs(0)
                    varArgs(0) = retvalue
                    sErrMsg = GetStringPara("U8.DZ.JA.Res590", varArgs(0))

                    'MsgBox "供应商" & RetValue & "不存在或者没有权限,请重新输入!", vbInformation, GetString("U8.DZ.JA.Res030")
                    'bChanged = Cancel
                    Exit Function
                Else
                    SetDomValue oDomHead, "cfreightCompany", cvencode
                    SetDomValue oDomHead, "MycdefineT2", cvenname

                    If sMetaItemName = LCase("cfreightCompany") Then
                        retvalue = strCellCode
                    Else
                        retvalue = strCellName
                    End If

                End If
            End If
            '汇率,币种
        Case "iexchrate", "cexch_name"
            If retvalue = "" Or retvalue = "0" Then
                sErrMsg = "数值输入不合法!"
                Exit Function
            Else
                '更新表体
                Call GetPrice("allprice", "97", oDomHead, oDomBody)
            End If

    End Select
 
    HandRecord_T = True
    
    Exit Function

ErrHandler:

   sErrMsg = Err.Description
   
End Function



Private Sub GlobalInit(login As U8Login.clsLogin)

    '变量初始化
    gstrCardNumberlist = "PD010301"
    gstrCardNumber = "PD010301"

    '数据表定义
    MainTable = "HY_DZ_BorrowOut"                          '单据主表
    DetailsTable = "HY_DZ_BorrowOuts"                      '单据字表
    HeadPKFld = "ID"                                       '主表主键字段
    MainView = "V_HY_DZ_BorrowOut"                         '表头视图
    DetailsView = "V_HY_DZ_BorrowOuts"                     '表体视图
    VoucherList = "V_HY_DZ_BorrowOutList"                  '列表视图

    strcCode = "cCode"                                     '单据编号
    StrcMaker = "cMaker"                                   '制单人
    StrdDate = "dDate"                                     '单据日期 借用日期
    StrcHandler = "cHandler"                               '审核人
    StrdVeriDate = "dVeriDate"                             '审核日期
    StrCloseUser = "CloseUser"                             '关闭人
    StrdCloseDate = "dCloseDate"                           '关闭日期
    StrIntoUser = "IntoUser"                               '生单人
    StrdIntoDate = "dIntoDate"                             '生单日期
    StriStatus = "iStatus"                                 '状态

    Set clsInfor = CreateObject("Info_PU.ClsS_Infor")      'New Info_PU.ClsS_Infor
    Call clsInfor.Init(login)

    Set m_SysInfor = clsInfor.Information

    ' 数量小数位
    m_sQuantityFmt = "#,##0" & IIf(m_SysInfor.iQuantityBit = 0, "", ".") & GetPrecision(m_SysInfor.iQuantityBit)

    ' 件数小数位数
    m_sNumFmt = "#,##0" & IIf(m_SysInfor.iNumBit = 0, "", ".") & GetPrecision(m_SysInfor.iNumBit)

    ' 换算率小数位数
    m_iExchRateFmt = "#,##0" & IIf(m_SysInfor.iExchRateBit = 0, "", ".") & GetPrecision(m_SysInfor.iExchRateBit)

    ' 税率小数位数
    m_iRateFmt = "#,##0" & IIf(m_SysInfor.iRateBit = 0, "", ".") & GetPrecision(m_SysInfor.iRateBit)

    ' 存货单价小数位(采购用)（库存用）
    m_sPriceFmt = "#,##0" & IIf(m_SysInfor.iCostBit = 0, "", ".") & GetPrecision(m_SysInfor.iCostBit)

    ' 开票单价小数位(销售用)
    m_sPriceFmtSA = "#,##0" & IIf(m_SysInfor.iBillCostBit = 0, "", ".") & GetPrecision(m_SysInfor.iBillCostBit)
    
    Call getAuthString(g_Conn)

    '控制过滤、参照等不会因为数据权限而无法显示数据 11-7-12
    g_oLogin.AuthString = "warehouse=True,department=True,inventory=True,person=False,user=True,vendor=True,purchasetype=False,customer=True,position=True"

End Sub





'以下提供校验、保存、审核，弃审、删除服务




Private Function GetDomValue(ByVal oDocument As DOMDocument, ByVal nCol As Variant, Optional ByVal nRow As Long) As Variant
    '================================================================================
    '模块: PublicSub   Function: GetDomValue
    '作者:
    '日期: 2002-04-24
    '--------------------------------------------------------------------------------
    '说明: 获取DOCUMENT单元结点值
    '
    '--------------------------------------------------------------------------------
    '参数:
    '   输入
    '   输出
    '依赖:
    '--------------------------------------------------------------------------------
    '修改:
    '
    '================================================================================
    On Error GoTo Error_General_Handler:
    
    'Call TraceIn(PROC_SIG)
    Dim i As Long
    Dim sItemName As String
    Dim element As IXMLDOMElement
    
    If LCase(TypeName(nCol)) <> "string" Then
        Set element = oDocument.selectSingleNode("//s:AttributeType[@rs:number='" & nCol & "']") ' and @rs:hidden='false']")
        If element Is Nothing Then
            sItemName = ""
        Else
            sItemName = element.getAttribute("name")
        End If
    Else
        sItemName = nCol
    End If
    
    GetDomValue = ""
    '    If oDocument.documentElement.selectSingleNode("//z:row") Is Nothing Then
    '
    '    End If
    Dim NodeList As IXMLDOMNodeList
    Set NodeList = oDocument.selectNodes("//z:row")
    If nRow > 0 Then
        i = 1
    End If
    For Each element In NodeList
        If i = nRow And GetAttrValue(element, sItemName) <> "" Then
            GetDomValue = element.getAttribute(sItemName)
            '            Exit For
            '        Else
            '            GetDomValue = ""
        End If
        i = i + 1
    Next
    'Call TraceOut(PROC_SIG)
ExitFunction:
    If IsNull(GetDomValue) Then
       GetDomValue = ""
    End If
    Exit Function
Error_General_Handler:
    
    GoTo ExitFunction
End Function

Private Sub SetDomValue(ByRef oDocument As DOMDocument, ByVal nCol As Variant, ByVal vNewValue As Variant, Optional ByVal nRow As Long)
    '================================================================================
    '模块: PublicSub   Function: SetDomValue
    '作者:
    '日期: 2002-04-24
    '--------------------------------------------------------------------------------
    '说明: 赋值DOCUMENT单元结点值
    '
    '--------------------------------------------------------------------------------
    '参数:
    '   输入
    '   输出
    '依赖:
    '--------------------------------------------------------------------------------
    '修改:
    '
    '================================================================================
    On Error GoTo Error_General_Handler:
    
    'Call TraceIn(PROC_SIG)
    Dim i As Long
    Dim sItemName As String
    Dim element As IXMLDOMElement
    
    If LCase(TypeName(nCol)) <> "string" Then
        Set element = oDocument.selectSingleNode("//s:AttributeType[@rs:number='" & nCol & "']") ' and @rs:hidden='false']")
        If element Is Nothing Then
            sItemName = ""
        Else
            sItemName = element.getAttribute("name")
        End If
    Else
        sItemName = nCol
    End If
    
    Dim NodeList As IXMLDOMNodeList
    Set NodeList = oDocument.selectNodes("//z:row")
    If nRow > 0 Then
        i = 1
    End If
    For Each element In NodeList
        'If i = nRow And GetAttrValue(element, sItemName) Then
        If i = nRow Then
            'GetDomValue = element.getAttribute(sItemName)
            element.setAttribute sItemName, vNewValue
            '            Exit For
            '        Else
            '            'element.getAttribute(sItemName) = ""
            '            element.setAttribute sItemName, ""
        End If
        i = i + 1
    Next
    'Call TraceOut(PROC_SIG)
ExitFunction:
    Exit Sub
Error_General_Handler:
    GoTo ExitFunction
End Sub


Private Function HandRecord_TbObjectCode(ByVal sFieldName As String, ByRef oDomHead As DOMDocument, ByRef oDomBody As DOMDocument, _
                                         ByRef retvalue As String, Optional ByRef sErrMsg As String = "") As Boolean
    On Error GoTo ErrHandler

    Dim strValue As Boolean
    Dim sql As String
    Dim sMetaItemName As String
    sMetaItemName = LCase(sFieldName)
    
    
    HandRecord_TbObjectCode = False

    '手工输入时,需要校验是否存在
    '/*B*/ 根据单据模板表头设置确定是否需要以下栏目
    Select Case GetDomValue(oDomHead, "cType")
            '部门
            'enum by modify
        Case "部门"
            If retvalue = "" Then
                SetDomValue oDomHead, "bObjectCode", ""
                SetDomValue oDomHead, "bObjectName", ""
            Else
                sql = "select cdepcode code,cdepname name from department where (cdepcode='" & retvalue & "' or cdepname='" & retvalue & "') and isnull(dDepEndDate,'9999-12-31')>N'" & g_oLogin.CurDate & "' "
                sql = sql & IIf(sAuth_depW = "", "", " and cdepcode in (" & sAuth_depW & ")")

                strValue = CheckCellValue(sql)

                If strValue = False Then
                    ReDim varArgs(0)
                    varArgs(0) = retvalue
                    sErrMsg = GetStringPara("U8.DZ.JA.Res560", varArgs(0))

                    '                MsgBox "部门" & RetValue & "不存在或者没有权限,请重新输入!", vbInformation, GetString("U8.DZ.JA.Res030")
                    
                    Exit Function
                Else
                    SetDomValue oDomHead, "bObjectCode", strCellCode
                    SetDomValue oDomHead, "bObjectName", strCellName
                    If sMetaItemName Like "*code" Then
                        retvalue = strCellCode
                    Else
                        retvalue = strCellName
                    End If
                End If
            End If
            '业务员
        Case "人员"
            If retvalue = "" Then
                SetDomValue oDomHead, "bObjectCode", ""
                SetDomValue oDomHead, "bObjectName", ""
            Else
                sql = "select cPersonCode code,cPersonName name from person where (cPersonCode='" & retvalue & "' or cPersonName='" & retvalue & "') and  '" & g_oLogin.CurDate & "' between dPValidDate  and IsNull(dPInValidDate,'2099-12-31') "

                strValue = CheckCellValue(sql)

                If strValue = False Then
                    ReDim varArgs(0)
                    varArgs(0) = retvalue
                    sErrMsg = GetStringPara("U8.DZ.JA.Res570", varArgs(0))

                    '                MsgBox "业务员" & RetValue & "不存在或者没有权限,请重新输入!", vbInformation, GetString("U8.DZ.JA.Res030")
                    'bChanged = Cancel
                    Exit Function
                Else
                    SetDomValue oDomHead, "bObjectCode", strCellCode
                    SetDomValue oDomHead, "bObjectName", strCellName
                    If sMetaItemName Like "*code" Then
                        retvalue = strCellCode
                    Else
                        retvalue = strCellName
                    End If
                End If
            End If
            '客户
        Case "客户"
            Dim cName As String
            Dim Address As String

            If retvalue = "" Then
                SetDomValue oDomHead, "bObjectCode", ""
                SetDomValue oDomHead, "bObjectName", ""
            Else
                sql = "select cCusCode ,cCusAbbName,cCusName,cCusAddress  from Customer where (cCusCode='" & retvalue & "' or cCusAbbName='" & retvalue & "' or cCusMnemCode ='" & retvalue & "' or cCusName ='" & retvalue & "') and  isnull(dEndDate,'9999-12-31')>N'" & g_oLogin.CurDate & "' "
                sql = sql & IIf(sAuth_CusW = "", "", " and cCusCode in (" & sAuth_CusW & ")")

                strValue = CheckCustomer(sql, cName, Address)

                If strValue = False Then
                    ReDim varArgs(0)
                    varArgs(0) = retvalue
                    sErrMsg = GetStringPara("U8.DZ.JA.Res580", varArgs(0))

                    '                MsgBox "客户" & RetValue & "不存在或者没有权限,请重新输入!", vbInformation, GetString("U8.DZ.JA.Res030")
                    'bChanged = Cancel
                    Exit Function
                Else
                    SetDomValue oDomHead, "bObjectCode", strCellCode
                    SetDomValue oDomHead, "bObjectName", strCellName

                    If sMetaItemName Like "*code" Then
                        retvalue = strCellCode
                    Else
                        retvalue = strCellName
                    End If
                End If
            End If
        Case "供应商"
            Dim cvencode As String
            Dim cvenname As String

            If retvalue = "" Then
                SetDomValue oDomHead, "bObjectCode", ""
                SetDomValue oDomHead, "bObjectName", ""
            Else
                sql = "select cvencode ,cvenname,cvenabbname  from vendor where (cvencode='" & retvalue & "' or cvenname='" & retvalue & "' or cVenMnemCode='" & retvalue & "' or cVenAbbName ='" & retvalue & "') and  isnull(dEndDate,'9999-12-31')>N'" & g_oLogin.CurDate & "' "
                sql = sql & IIf(sAuth_vendorW = "", "", " and iid in (" & sAuth_vendorW & ")")

                strValue = CheckVendor(sql, cvencode, cvenname)

                If strValue = False Then
                    ReDim varArgs(0)
                    varArgs(0) = retvalue
                    sErrMsg = GetStringPara("U8.DZ.JA.Res590", varArgs(0))

                    '                MsgBox "供应商" & RetValue & "不存在或者没有权限,请重新输入!", vbInformation, GetString("U8.DZ.JA.Res030")
                    'bChanged = Cancel
                    Exit Function
                Else
                    SetDomValue oDomHead, "bObjectCode", cvencode
                    SetDomValue oDomHead, "bObjectName", cvenname

                    If sMetaItemName Like "*code" Then
                        retvalue = strCellCode
                    Else
                        retvalue = strCellName
                    End If

                End If
            End If
    End Select

    HandRecord_TbObjectCode = True
    Exit Function
ErrHandler:
    sErrMsg = Err.Description
End Function


'取价（单行、整张）
Public Function GetPrice(ByVal strKey As String, ByVal strVouchType As String, ByRef oDomHead As DOMDocument, ByRef oDomBody As DOMDocument, _
                          Optional ByVal iRow As Long = 0) As Boolean
    Dim domHead As New DOMDocument
    Dim domBody As New DOMDocument
    Dim errMsg As String
    Dim i As Long
    
    GetPrice = True

    On Error GoTo Err_Handler
    If oDomBody.selectNodes("//z:row").Length < 1 Then
        'MsgBox GetString("U8.DZ.JA.Res770"), vbInformation, GetString("U8.DZ.JA.Res030")
        GetPrice = False
        Exit Function
    End If

    If strKey = "rowprice" Then
        If GetDomValue(oDomBody, "cinvcode", iRow) = "" Then
            'MsgBox GetString("U8.DZ.JA.Res770"), vbInformation, GetString("U8.DZ.JA.Res030")
            GetPrice = False
            Exit Function
        End If
    End If

    '销售选项信息 myinfo
    Dim clsSAWeb As Object
    Set clsSAWeb = Nothing
    Set clsSAWeb = New USSAServer.clsSystem
    clsSAWeb.Init g_oLogin
    clsSAWeb.INIMyInfor

    Dim clsVoucherCO As New VoucherCO_Sa.ClsVoucherCO_SA
    clsVoucherCO.Init strVouchType, g_oLogin, g_Conn, "CS", clsSAWeb

    
    Dim oEle As IXMLDOMElement
    
    Dim oEleRow As IXMLDOMElement
    
    Dim oEleBody As IXMLDOMNode
    
    If strKey = "allprice" Then
        If strVouchType <> "95" And strVouchType <> "92" And strVouchType <> "98" And strVouchType <> "99" Then
            Set oEleBody = oDomBody.selectSingleNode("//rs:data")
            For Each oEle In oEleBody.selectNodes("//z:row")
                If GetEleVal(oEle, "cinvcode") = "" Then
                    oEleBody.removeChild oEle
                End If
            Next
        End If
       
    Else
        If GetDomValue(oDomBody, "cscloser", iRow) <> "" Then
            '      MsgBox GetString("U8.SA.clsOpSA.clscommcheck.00018")
            GetPrice = False
            Exit Function
        End If
        
        Set domHead = oDomHead.cloneNode(True)
        Set domBody = GetLineDom(oDomBody, iRow)
    End If
    errMsg = clsVoucherCO.VoucherGetPrice(g_Conn, domHead, domBody)

    If errMsg <> "" Then
        GetPrice = False
        Set domHead = Nothing
        Set domBody = Nothing
        Exit Function
    End If
    If strKey = "allprice" Then
        Set oDomHead = domHead.cloneNode(True)
        Set oDomBody = domBody.cloneNode(True)
    Else
        'Voucher.UpdateLineData domBody, Voucher.row
        
       
        Set oEle = oDomBody.selectNodes("//z:row").Item(iRow - 1)
        Set oEleRow = domBody.selectSingleNode("//z:row")
        
        Dim iAtrr As IXMLDOMAttribute
        If Not oEle Is Nothing Then
             For i = 0 To oEleRow.Attributes.Length - 1
                  oEle.setAttribute oEleRow.Attributes.Item(i).nodename, oEleRow.getAttribute(oEleRow.Attributes.Item(i).nodename)
             Next
        End If
        
    End If
    Set domHead = Nothing
    Set domBody = Nothing
    Set clsSAWeb = Nothing

    Exit Function



Err_Handler:
    Set domHead = Nothing
    Set domBody = Nothing
    Set clsSAWeb = Nothing
    GetPrice = False
    VBA.Err.Clear

End Function

Private Function GetLineDom(domFrom As DOMDocument, Optional ByVal nRow As Long = 0) As DOMDocument
    On Error GoTo Err_log
    Dim dom     As New DOMDocument
    Dim domBody As New DOMDocument
    
    Dim nd      As IXMLDOMNode
    Dim ndrs    As IXMLDOMNode
    
    Dim n       As Long
    Dim sEdit   As String
    n = nRow
    Set GetLineDom = New DOMDocument
    Set domBody = CloneDomDocument(domFrom)
    If Not dom.loadXML(CloneDomDocument(domFrom, True).xml) Then Exit Function
    Set nd = GetValidNode(n, domBody)
    
    Set ndrs = dom.selectSingleNode("//rs:data")
    If ndrs Is Nothing Then
        Set ndrs = dom.createElement("rs:data")
        dom.documentElement.appendChild ndrs
    End If
    If Not ndrs Is Nothing Then ndrs.appendChild nd
    
    Set GetLineDom = dom
    Exit Function
Err_log:
    'RaiseEvent controlError(Err.Number, Err.Description)
End Function


Private Function GetEleVal(oEle As IXMLDOMElement, ByVal sName As String) As String
    
    If oEle Is Nothing Then
       GetEleVal = ""
       Exit Function
    End If
    
    Dim v As Variant
    v = oEle.getAttribute(sName)
    
    If IsNull(v) Then
       GetEleVal = ""
    Else
       GetEleVal = CStr(v)
    End If
End Function

'表体校验（自定义项、自由项 之外的栏目）
'存货编码
'需要赋值的字段：存货名称、存货代码、规则型号
'               计量单位组编码、名称、主计量单位编码、名称、辅计量单位名称、编码、换算率
'               保质期单位、保质期天数
'               1-16存货自定义项

Public Function BodyCellCheck(ByVal sFieldName As String, ByRef oDomHead As DOMDocument, ByRef oDomBody As DOMDocument, _
                                ByRef retvalue As Variant, ByRef sErrMsg As String) As Boolean
    On Error GoTo Err_Handler

    Dim sMetaItemXML As String
    sMetaItemXML = sFieldName



    BodyCellCheck = False
    Dim sSql As String
    Dim oInventoryPst As USERPDMO.InventoryPst
    Dim moInventory As USERPVO.Inventory
    Dim moStockPst As USERPDMO.StockPst
    'Dim cWhcode As String

    Dim rs As New ADODB.Recordset
    Dim sql As String
    Dim aGetFloatRateRule As Boolean
    aGetFloatRateRule = GetFloatRateRule(g_Conn)
    Dim rst As New ADODB.Recordset


    Dim sWhCode As String                                  '仓库编码
    Dim sInvCode As String
    Dim sFreeName As String
    Dim sDate As String
    '存货编码
    '存货编码
    sInvCode = GetDomValue(oDomBody, "cinvcode")
    sWhCode = GetDomValue(oDomBody, "cwhcode")
    sDate = GetDomValue(oDomHead, StrdDate)

    If LCase(sMetaItemXML) = "cbatch" Or LCase(sMetaItemXML) = "cinvouchcode" Then    '批号和入库单号参照所需处理

        '********************************************
        '2008-11-17
        '为匹配872中LP件多种销售跟踪方式的处理
        Dim sSosId As String                               '销售订单行ID
        Dim sDemandType As String                          '销售订单类型
        Dim sDemandCode As String                          '销售订单分类号
        Dim lDemandCode As Long                            '整数型订单行号

        '********************************************
        '销售订单行ID

        

        '/************批号参照相关参数初始化*******************/ chenliangc
        sSosId = GetDomValue(oDomBody, "isosid")

        Call GetSoDemandType(sSosId, sDemandType, sDemandCode, g_Conn)
        If IsNumeric(sDemandCode) Then
            lDemandCode = CLng(sDemandCode)
        Else
            lDemandCode = 0
        End If
    End If
    '/****************************************************/


    Select Case sMetaItemXML
            '仓库
        Case "cwhcode", "cwhname"
            
            retvalue = CCrmStr(Null2Something(retvalue, ""))
    
            If retvalue = "" Then
                SetDomValue oDomBody, "cwhname", ""
                SetDomValue oDomBody, "cwhcode", ""
                SetDomValue oDomBody, "cPosition", ""
                SetDomValue oDomBody, "cPosition2", ""

            Else
                Set rs = cWhCodeRefer(CStr(retvalue), sDate)
                If rs Is Nothing Or rs.State = 0 Then
                    ReDim varArgs(0)
                    varArgs(0) = retvalue
                    sErrMsg = GetStringPara("U8.DZ.JA.Res670", varArgs(0))
                    'MsgBox "仓库" & RetValue & "不存在，请重新输入", vbInformation, GetString("U8.DZ.JA.Res030")
                    BodyCellCheck = False
                    Exit Function
                Else
                    If rs.RecordCount > 0 Then

                        '检查存货仓库对照关系

'                        sql = "select  cvalue from accinformation where cname='bCheckWareInv' and csysid='ST'"
'                        Set rst = New ADODB.Recordset
'                        rst.Open sql, g_Conn
'                        If rst.Fields("cvalue") = "True" Then
'                            sql = "select cinvcode from  WhInvContrapose  where cinvcode='" & GetDomValue(oDomBody, "cinvcode") & "' and  cwhcode='" & rs.Fields("cwhcode").Value & "'"
'                            Set rst = New ADODB.Recordset
'                            rst.Open sql, g_Conn
'                            If rst.EOF Then
'                                If MsgBox(GetStringPara("U8.DZ.JA.Res2180", rs.Fields("cwhcode").Value), vbInformation + vbYesNo, GetString("U8.DZ.JA.Res030")) = vbNo Then
'                                    BodyCellCheck = False
'                                    Exit Function
'                                End If
'                            End If
'                        End If

                        SetDomValue oDomBody, "cwhcode", rs.Fields("cwhcode").Value
                        SetDomValue oDomBody, "cwhname", rs.Fields("cwhname").Value
                        SetDomValue oDomBody, "cPosition", ""
                        SetDomValue oDomBody, "cPosition2", ""


                        If sMetaItemXML = "cwhname" Then retvalue = rs.Fields("cwhname").Value
                    End If
                End If
            End If


            '货位
        Case "cPosition", "cPosition2"
            
            retvalue = CCrmStr(Null2Something(retvalue, ""))
    
            If retvalue = "" Then
                SetDomValue oDomBody, "cPosition", ""
                SetDomValue oDomBody, "cPosition2", ""

            Else
                ' cwchcode = Voucher.bodyText(r, "cwhcode") & ""
                If GetDomValue(oDomBody, "cwhcode") = "" Then
                    sErrMsg = GetString("U8.DZ.JA.Res1840")

                    Exit Function
                End If

                Set rs = cPosionRefer(CStr(retvalue), GetDomValue(oDomBody, "cwhcode"))
                If rs Is Nothing Or rs.State = 0 Then
                    ReDim varArgs(0)
                    varArgs(0) = retvalue
                    sErrMsg = GetStringPara("U8.DZ.JA.Res680", varArgs(0))

                    'MsgBox "货位" & RetValue & "不存在，请重新输入", vbInformation, GetString("U8.DZ.JA.Res030")

                    Exit Function
                Else
                    If rs.RecordCount > 0 Then

                        '检查存货货位对照关系

'                        sql = "select  cvalue from accinformation where cname='bCheckInvPos' and csysid='ST'"
'                        Set rst = New ADODB.Recordset
'                        rst.Open sql, g_Conn
'                        If rst.Fields("cvalue") = "True" Then
'                            sql = "select cinvcode from  InvPosContrapose  where cinvcode='" & Voucher.bodyText(r, "cinvcode") & "' and  cposcode='" & rs.Fields("cposcode").Value & "'"
'                            Set rst = New ADODB.Recordset
'                            rst.Open sql, g_Conn
'                            If rst.EOF Then
'                                If MsgBox(GetStringPara("U8.DZ.JA.Res2190", rs.Fields("cposcode").Value), vbInformation + vbYesNo, GetString("U8.DZ.JA.Res030")) = vbNo Then
'                                    bChanged = Cancel
'                                    Exit Function
'                                End If
'                            End If
'                        End If

                        SetDomValue oDomBody, "cPosition", rs.Fields("cposcode").Value
                        SetDomValue oDomBody, "cPosition2", rs.Fields("cposname").Value


                        If sMetaItemXML = "cPosition2" Then retvalue = rs.Fields("cposname").Value
                    End If
                End If
            End If
        Case "cinvcode"
            Dim i As Integer
            Dim iRow As Long
          
            retvalue = CCrmStr(Null2Something(retvalue, ""))
    
            If retvalue = "" Then
                SetDomValue oDomBody, "cinvcode", ""
                SetDomValue oDomBody, "cinvname", ""
                BodyCellCheck = True
                Exit Function
            Else
                Set rs = cInvCodeRefer(CStr(retvalue))
                If rs Is Nothing Or rs.State = 0 Then
                    ReDim varArgs(0)
                    varArgs(0) = retvalue
                    sErrMsg = GetStringPara("U8.DZ.JA.Res690", varArgs(0))

                    '                 MsgBox GetString("U8.DZ.JA.Res780") & RetValue & "不存在或者没有销售属性或者没有权限或者停用，请重新输入", vbInformation, GetString("U8.DZ.JA.Res030")

                    Exit Function

                Else

                    '检查存货仓库货位对照关系
'                    If Voucher.bodyText(iRow, "cwhcode") & "" <> "" Then
'                        sql = "select  cvalue from accinformation where cname='bCheckWareInv' and csysid='ST'"
'                        Set rst = New ADODB.Recordset
'                        rst.Open sql, g_Conn
'                        If rst.Fields("cvalue") = "True" Then
'                            sql = "select cinvcode from  WhInvContrapose  where cinvcode='" & rs.Fields("cinvcode").Value & "' and  cwhcode='" & Voucher.bodyText(iRow, "cwhcode") & "'"
'                            Set rst = New ADODB.Recordset
'                            rst.Open sql, g_Conn
'                            If rst.EOF Then
'                                If MsgBox(GetStringPara("U8.DZ.JA.Res2200", rs.Fields("cinvcode").Value), vbInformation + vbYesNo, GetString("U8.DZ.JA.Res030")) = vbNo Then
'                                    bChanged = Cancel
'                                    Exit Function
'                                End If
'                            End If
'                        End If
'                    End If

'                    If Voucher.bodyText(iRow, "cPosition") & "" <> "" Then
'                        sql = "select  cvalue from accinformation where cname='bCheckInvPos' and csysid='ST'"
'                        Set rst = New ADODB.Recordset
'                        rst.Open sql, g_Conn
'                        If rst.Fields("cvalue") = "True" Then
'                            sql = "select cinvcode from  InvPosContrapose  where cinvcode='" & rs.Fields("cinvcode").Value & "' and  cposcode='" & Voucher.bodyText(iRow, "cPosition") & "'"
'                            Set rst = New ADODB.Recordset
'                            rst.Open sql, g_Conn
'                            If rst.EOF Then
'                                If MsgBox(GetStringPara("U8.DZ.JA.Res2210", rs.Fields("cinvcode").Value), vbInformation + vbYesNo, GetString("U8.DZ.JA.Res030")) = vbNo Then
'                                    bChanged = Cancel
'                                    Exit Function
'                                End If
'                            End If
'                        End If
'                    End If
            End If
        End If



        '赋值
        '给存货编码等赋值
setvalue:
        If Not rs Is Nothing And rs.State <> 0 Then
            If Not rs.EOF Then
                SetBodyCellValue oDomHead, oDomBody, rs
                retvalue = rs!cinvcode
            End If
            

           
           '                    Voucher.bodyText(iRow, sMetaItemXML) = Rs!cinvcode
            rs.Close
            Set rs = Nothing
        End If


    Case "iquantity"
        If Not myCheckNumValue(retvalue, GetString("U8.DZ.JA.Res710"), sErrMsg, 1) Then
            
            Exit Function
        End If

        '1无换算率
        If GetDomValue(oDomBody, "igrouptype") = 0 Then
            GoTo lsuccess
        End If

        '2固定换算率
        If GetDomValue(oDomBody, "igrouptype") = 1 Then
            SetDomValue oDomBody, "inum", CDbl(GetDomValue(oDomBody, "iquantity")) / IIf(ConvertStrToDbl(GetDomValue(oDomBody, "iinvexchrate")) = 0, 1, ConvertStrToDbl(GetDomValue(oDomBody, "iinvexchrate")))

            GoTo lsuccess
        End If
        '3浮动换算率
        If GetDomValue(oDomBody, "igrouptype") = 2 Then
            If aGetFloatRateRule Then
                If GetDomValue(oDomBody, "iinvexchrate") <> "" Then
                    If GetDomValue(oDomBody, "iinvexchrate") = 0 Then
                        SetDomValue oDomBody, "inum", 0
                    Else
                        SetDomValue oDomBody, "inum", CDbl(retvalue) / CDbl(GetDomValue(oDomBody, "iinvexchrate"))
                    End If
                Else
                    If GetDomValue(oDomBody, "inum") <> "" Then
                        If CDbl(GetDomValue(oDomBody, "inum")) = 0 Then
                            SetDomValue oDomBody, "iinvexchrate", 0
                        Else
                            'Voucher.bodyText(r, "iinvexchrate") = retvalue / Voucher.bodyText(r, "inum")
                            SetDomValue oDomBody, "iinvexchrate", CDbl(retvalue) / CDbl(GetDomValue(oDomBody, "inum"))
                        End If
                    End If
                End If
            Else
                If GetDomValue(oDomBody, "inum") <> "" Then
                    If CDbl(GetDomValue(oDomBody, "inum")) = 0 Then
                        SetDomValue oDomBody, "iinvexchrate", 0
                    Else
                        SetDomValue oDomBody, "iinvexchrate", CDbl(retvalue) / CDbl(GetDomValue(oDomBody, "inum"))
                    End If
                Else
                    If GetDomValue(oDomBody, "iinvexchrate") <> "" Then
                        If CDbl(GetDomValue(oDomBody, "iinvexchrate")) = 0 Then
                            SetDomValue oDomBody, "inum", 0
                        Else
                             SetDomValue oDomBody, "inum", CDbl(retvalue) / CDbl(GetDomValue(oDomBody, "iinvexchrate"))
                        End If
                    End If

                End If
            End If


        End If

        '件数
    Case "inum"
        If Not myCheckNumValue2(retvalue, GetString("U8.DZ.JA.Res720"), sErrMsg, 0) Then
            
            Exit Function
        End If

        '1无换算率
        If CDbl(GetDomValue(oDomBody, "igrouptype")) = 0 Then
            GoTo lsuccess
        End If

        '2固定换算率
        If GetDomValue(oDomBody, "igrouptype") = 1 Then
            If CDbl(GetDomValue(oDomBody, "inum")) <> 0 Then
                SetDomValue oDomBody, "iquantity", CDbl(GetDomValue(oDomBody, "inum")) * CDbl(GetDomValue(oDomBody, "iinvexchrate"))
            Else
                SetDomValue oDomBody, "iquantity", ""
            End If

            GoTo lsuccess
        End If

        '3浮动换算率
        If GetDomValue(oDomBody, "igrouptype") = 2 Then
            If aGetFloatRateRule Then
                If retvalue = "" Then
                    SetDomValue oDomBody, "iquantity", ""
                    BodyCellCheck = True
                    Exit Function
                End If

                If GetDomValue(oDomBody, "iquantity") <> "" Then
                    If CDbl(retvalue) = 0 Then
                         SetDomValue oDomBody, "iinvexchrate", 0
                    Else
                        SetDomValue oDomBody, "iinvexchrate", CDbl(GetDomValue(oDomBody, "iquantity")) / CDbl(retvalue)
                    End If
                ElseIf GetDomValue(oDomBody, "iinvexchrate") <> "" Then
                    If GetDomValue(oDomBody, "iquantity") = "" Then
                        SetDomValue oDomBody, "iquantity", CDbl(retvalue) * CDbl(GetDomValue(oDomBody, "iinvexchrate"))
                    End If
                End If
            Else
                If GetDomValue(oDomBody, "iinvexchrate") <> "" Then
                    If retvalue = "" Then
                        SetDomValue oDomBody, "iquantity", ""
                    Else
                        SetDomValue oDomBody, "iquantity", CDbl(retvalue) * CDbl(GetDomValue(oDomBody, "iinvexchrate"))
                    End If
                End If
                If GetDomValue(oDomBody, "iquantity") <> "" Then
                    If GetDomValue(oDomBody, "iinvexchrate") = "" Then
                        If CDbl(retvalue) > 0 Then
                            SetDomValue oDomBody, "iinvexchrate", CDbl(GetDomValue(oDomBody, "iquantity")) / CDbl(retvalue)
                        End If
                    End If
                End If
            End If
        End If
        '换算率
    Case "iinvexchrate"
        '           '浮动换算率
        '           If Voucher.bodyText(r, "igrouptype") <> 2 Then
        '                bChanged = Cancel
        '                Exit Function
        '           End If

        '空就清空项目
        If Trim(retvalue) = "" Then
            If aGetFloatRateRule Then
                SetDomValue oDomBody, "inum", ""
            Else
                SetDomValue oDomBody, "iquantity", ""
            End If

            GoTo lsuccess
        End If

        If Not myCheckNumValue(retvalue, GetString("U8.DZ.JA.Res730"), sErrMsg, 0) Then
            
            Exit Function
        End If

        '变化关系
        If aGetFloatRateRule Then
            If GetDomValue(oDomBody, "iquantity") <> "" Then
                If retvalue = 0 Then
                    SetDomValue oDomBody, "inum", 0
                Else
                    SetDomValue oDomBody, "inum", CDbl(GetDomValue(oDomBody, "iquantity")) / IIf(ConvertStrToDbl(GetDomValue(oDomBody, "iinvexchrate")) = 0, 1, ConvertStrToDbl(GetDomValue(oDomBody, "iinvexchrate")))
                End If
            End If
            If GetDomValue(oDomBody, "iquantity") = "" Then
                If GetDomValue(oDomBody, "inum") <> "" Then SetDomValue oDomBody, "iquantity", CDbl(GetDomValue(oDomBody, "inum")) * CDbl(GetDomValue(oDomBody, "iinvexchrate"))
            End If

        Else
            If GetDomValue(oDomBody, "inum") <> "" Then _
                    SetDomValue oDomBody, "iquantity", CDbl(GetDomValue(oDomBody, "inum")) * CDbl(GetDomValue(oDomBody, "iinvexchrate"))
            If GetDomValue(oDomBody, "iquantity") <> "" Then
                If GetDomValue(oDomBody, "inum") = "" Then
                    If CDbl(retvalue) = 0 Then
                        SetDomValue oDomBody, "inum", 0
                    Else
                        SetDomValue oDomBody, "inum", CDbl(GetDomValue(oDomBody, "iquantity")) / IIf(ConvertStrToDbl(GetDomValue(oDomBody, "iinvexchrate")) = 0, 1, ConvertStrToDbl(GetDomValue(oDomBody, "iinvexchrate")))
                    End If
                End If
            End If
        End If

        '修改件数,辅计量单位、数量,换算率,项目大类,项目编码
        '报价、本币无税单价,原币含税单价,原币无税单价
        '       Case "inum", "cinva_unit", "iquantity", "iinvexchrate", "citem_class", "citemcode", _
                '            "iquotedprice", "inatunitprice", "inatmoney", "inattax", "inatsum", "itaxrate", "inatdiscount", _
                '            "itaxunitprice", "iunitprice", "imoney", "itax", "isum", "idiscount", "kl", "kl2", "dkl1", "dkl2", "fsalecost", "fsaleprice", "fcusminprice"
        '       Case "iquantity", "inum"
        '                Dim A As USERPCO.VoucherCO
        '                Dim StLogin As New USCOMMON.login
        '                A.IniLogin g_oLogin, errmsg
        '                Set StLogin = A.login
        '                A.CheckBody "0301", nOther, r, "", dombody, errmsg, domHead

        '有效期校验 chenliangc
    Case "dmadedate", "dvdate"

        If (sMetaItemXML = "dmadedate" And GetDomValue(oDomBody, "dmadedate") = "") Or (sMetaItemXML = "dvdate" And GetDomValue(oDomBody, "dvdate") = "") Then
            SetDomValue oDomBody, "dmadedate", ""
            SetDomValue oDomBody, "dvdate", ""
            SetDomValue oDomBody, "dexpirationdate", ""
            SetDomValue oDomBody, "cexpirationdate", ""
        Else
            If GetDomValue(oDomBody, "cMassUnit") <> "" And GetDomValue(oDomBody, "imassdate") <> "" Then
                Select Case GetDomValue(oDomBody, "cMassUnit")

                    Case "1":
                        If sMetaItemXML = "dmadedate" Then SetDomValue oDomBody, "dvdate", DateAdd("yyyy", CInt(GetDomValue(oDomBody, "imassdate")), CDate(GetDomValue(oDomBody, "dmadedate")))
                        SetDomValue oDomBody, "dMadeDate", DateAdd("yyyy", 0 - CInt(GetDomValue(oDomBody, "imassdate")), CDate(GetDomValue(oDomBody, "dvdate")))
                    Case "2":
                        If sMetaItemXML = "dmadedate" Then SetDomValue oDomBody, "dvdate", DateAdd("m", CInt(GetDomValue(oDomBody, "imassdate")), CDate(GetDomValue(oDomBody, "dmadedate")))
                        SetDomValue oDomBody, "dMadeDate", DateAdd("m", 0 - CInt(GetDomValue(oDomBody, "imassdate")), CDate(GetDomValue(oDomBody, "dvdate")))
                    Case "3":
                        If sMetaItemXML = "dmadedate" Then SetDomValue oDomBody, "dvdate", DateAdd("d", CInt(GetDomValue(oDomBody, "imassdate")), CDate(GetDomValue(oDomBody, "dmadedate")))
                        SetDomValue oDomBody, "dMadeDate", DateAdd("d", 0 - CInt(GetDomValue(oDomBody, "imassdate")), CDate(GetDomValue(oDomBody, "dvdate")))
                End Select

                Select Case GetDomValue(oDomBody, "iExpiratDateCalcu")

                    Case "1":                              '"按月":
                        SetDomValue oDomBody, "dexpirationdate", CDate(GetDomValue(oDomBody, "dvdate")) - DatePart("d", CDate(GetDomValue(oDomBody, "dvdate")))
                        SetDomValue oDomBody, "cexpirationdate", Format(GetDomValue(oDomBody, "dexpirationdate"), "yyyy-mm")
                    Case "2":                              '"按日":
                        SetDomValue oDomBody, "dexpirationdate", DateAdd("d", -1, CDate(GetDomValue(oDomBody, "dvdate")))
                        SetDomValue oDomBody, "cexpirationdate", GetDomValue(oDomBody, "dexpirationdate")
                End Select

                If DateDiff("d", CDate(GetDomValue(oDomBody, "dvdate")), g_oLogin.CurDate) >= 0 Then
                    'If MsgBox(GetString("U8.DZ.JA.Res740"), vbInformation + vbYesNo, GetString("U8.DZ.JA.Res030")) = vbNo Then
                        retvalue = ""
                        SetDomValue oDomBody, "dmadedate", ""
                        SetDomValue oDomBody, "dvdate", ""
                        SetDomValue oDomBody, "dexpirationdate", ""
                        SetDomValue oDomBody, "cexpirationdate", ""
                    'End If
                End If
            End If
        End If


        '批号的校验 -chenliangc
    Case "cbatch"
'        If sWhCode = "" Then
'            MsgBox GetString("U8.DZ.JA.Res640"), vbInformation, GetString("U8.DZ.JA.Res030")
'            retvalue = ""
'            Exit Function
'        End If
            
        retvalue = CCrmStr(Null2Something(retvalue, ""))
    

        If retvalue = "" Then
            SetDomValue oDomBody, "cinvouchcode", ""
            SetDomValue oDomBody, "cvouchcode", ""
            SetDomValue oDomBody, "dmadedate", ""
            SetDomValue oDomBody, "dvdate", ""
            SetDomValue oDomBody, "dexpirationdate", ""
            SetDomValue oDomBody, "cexpirationdate", ""

            '清空批次属性
            For i = 0 To 10
                SetDomValue oDomBody, "cbatchproperty" & CStr(i), ""
            Next i
            BodyCellCheck = True
            
            Exit Function
        End If

        Set oInventoryPst = New USERPDMO.InventoryPst
        Set moInventory = New USERPVO.Inventory
        oInventoryPst.login = mologin
        oInventoryPst.Load sInvCode, moInventory

        If moInventory.IsBatch Then
            Set moStockPst = New StockPst
            moStockPst.login = mologin

            '871和872调用库存函数时，传递的参数类型发生变化
            '871的订单行参数要求是整数型，872的改成字符型，因此需要单独处理
            If gU8Version = "872" Then
                moStockPst.BatchList sSql, sWhCode, sInvCode, _
                        GetDomValue(oDomBody, "cfree1"), _
                        GetDomValue(oDomBody, "cfree2"), _
                        GetDomValue(oDomBody, "cfree3"), _
                        GetDomValue(oDomBody, "cfree4"), _
                        GetDomValue(oDomBody, "cfree5"), _
                        GetDomValue(oDomBody, "cfree6"), _
                        GetDomValue(oDomBody, "cfree7"), _
                        GetDomValue(oDomBody, "cfree8"), _
                        GetDomValue(oDomBody, "cfree9"), _
                        GetDomValue(oDomBody, "cfree10"), _
                        CStr(GetDomValue(oDomBody, sMetaItemXML)), , CLng(sDemandType), sDemandCode, ""
            Else
                moStockPst.BatchList sSql, sWhCode, sInvCode, _
                        GetDomValue(oDomBody, "cfree1"), _
                        GetDomValue(oDomBody, "cfree2"), _
                        GetDomValue(oDomBody, "cfree3"), _
                        GetDomValue(oDomBody, "cfree4"), _
                        GetDomValue(oDomBody, "cfree5"), _
                        GetDomValue(oDomBody, "cfree6"), _
                        GetDomValue(oDomBody, "cfree7"), _
                        GetDomValue(oDomBody, "cfree8"), _
                        GetDomValue(oDomBody, "cfree9"), _
                        GetDomValue(oDomBody, "cfree10"), _
                        CStr(GetDomValue(oDomBody, sMetaItemXML)), , CLng(sDemandType), lDemandCode, ""
            End If

            '                 Dim iQ As Long
            ''                Dim SQL As String
            '                Dim iSTConMode As Long
            '                Dim cvalue As String

            Set rs = New ADODB.Recordset

            '                 iQ = 0
            '                Set Rs = Nothing
            '                SQL = " select iSTConMode from Warehouse where cWhCode ='" & Voucher.bodyText(r, "cwhcode") & "'"
            '                Rs.Open SQL, g_Conn, 1, 1
            '                iSTConMode = Rs!iSTConMode
            '                Set Rs = Nothing
            '                 SQL = " select *  From V_CurrentStock left join vendor v on v.cvencode=V_CurrentStock.cvmivencode  left join v_aa_enum v1 on v1.enumcode=v_currentstock.iexpiratdatecalcu and v1.enumtype=N'SCM.ExpiratDateCalcu' left join AA_BatchProperty batch on Batch.cinvcode=V_CurrentStock.cinvcode and isnull(Batch.cbatch,N'')=isnull(V_CurrentStock.cbatch,N'') and isnull(Batch.cfree1,N'')=isnull(V_CurrentStock.cfree1,N'') and isnull(Batch.cfree2,N'')=isnull(V_CurrentStock.cfree2,N'') and isnull(Batch.cfree3,N'')=isnull(V_CurrentStock.cfree3,N'') and isnull(Batch.cfree4,N'')=isnull(V_CurrentStock.cfree4,N'') and isnull(Batch.cfree5,N'')=isnull(V_CurrentStock.cfree5,"
            '                              SQL = SQL + "N'') and isnull(Batch.cfree6,N'')=isnull(V_CurrentStock.cfree6,N'') and isnull(Batch.cfree7,N'')=isnull(V_CurrentStock.cfree7,N'') and isnull(Batch.cfree8,N'')=isnull(V_CurrentStock.cfree8,N'') and isnull(Batch.cfree9,N'')=isnull(V_CurrentStock.cfree9,N'') and isnull(Batch.cfree10,N'')=isnull(V_CurrentStock.cfree10,N'') Where V_CurrentStock.cWhcode=N'" & Voucher.bodyText(r, "cwhcode") & "' And V_CurrentStock.cInvCode =N'" & Voucher.bodyText(r, "cinvcode") & "' And V_CurrentStock.cBatch= N'" & Voucher.bodyText(r, "cbatch") & "' And IsNull(V_CurrentStock.cBatch,N'')<>N''  And isnull( bstopflag,0)=0  And (ISNULL(isotype,0)= 0 And ISNULL(isodid,N'')= N'') and (iQuantity+IsNull(fInQuantity,0)-IsNull(fOutQuantity,0)-IsNull(fStopQuantity,0)-" & Voucher.bodyText(r, "iquantity") & ") >0 order by  V_CurrentStock.dvdate,V_CurrentStock.cbatch"
            '                 Rs.Open SQL, g_Conn, 1, 1
            '                 If Rs.EOF Then
            '                    iQ = -1
            '                 End If
            '                 Set Rs = Nothing
            '                 SQL = "select cValue from accinformation where cname=N'bAllowZero' and csysid=N'ST'"
            '                Rs.Open SQL, g_Conn, 1, 1
            '                 cvalue = Rs!cvalue



            Set rs = Nothing
            rs.CursorLocation = adUseClient
            rs.Open sSql, g_Conn, adOpenDynamic, adLockBatchOptimistic

            
            
            If rs.RecordCount = 0 Then                     '没有找到批次

                If Not mologin.Account.BatchAllowZeroOut Then

                    '                                   MsgBox GetString("U8.DZ.JA.Res780") & Voucher.bodyText(r, "cinvname") & "没有找到批次结存，请参照录入" & vbCrLf, vbInformation, getstring("U8.DZ.JA.Res030")
                    '                                    retvalue = ""
                End If

                '                    '没有找到就清空 by liwq
                '                    For i = 0 To rs.Fields.Count - 1
                '                        sFreeName = IIf(IsNull(rs.Fields(i).Properties("BASECOLUMNNAME")), "", rs.Fields(i).Properties("BASECOLUMNNAME"))
                '                        If LCase(Mid(sFreeName, 1, 5)) = "cfree" Then
                '                            Voucher.bodyText(r, sFreeName) = ""
                '                        End If
                '                        'by liwqa 带出批次属性
                '                        If LCase(Mid(sFreeName, 1, 14)) = "cbatchproperty" Then
                '                            Voucher.bodyText(r, sFreeName) = ""
                '                        End If
                '                    Next

                '                             Voucher.bodyText(r, "iquantity") = Null2Something(Rs.Fields("出库数量"))
                '                             Voucher.bodyText(r, "inum") = Null2Something(Rs.Fields("出库件数"))

                SetDomValue oDomBody, "dmadedate", ""
                SetDomValue oDomBody, "dvdate", ""
                SetDomValue oDomBody, "dexpirationdate", ""
                SetDomValue oDomBody, "cexpirationdate", ""

            Else
                
                If rs.RecordCount >= 1 Then

                    '带出自由项
                    For i = 0 To rs.Fields.Count - 1
                        sFreeName = IIf(IsNull(rs.Fields(i).Properties("BASECOLUMNNAME")), "", rs.Fields(i).Properties("BASECOLUMNNAME"))
                        If LCase(Mid(sFreeName, 1, 5)) = "cfree" Then
                            SetDomValue oDomBody, LCase(sFreeName), Null2Something(rs.Fields(i).Value)
                        End If
                        'by liwqa 带出批次属性
                        If LCase(Mid(sFreeName, 1, 14)) = "cbatchproperty" Then
                            SetDomValue oDomBody, LCase(sFreeName), Null2Something(rs.Fields(i).Value)
                        End If
                    Next

                    '                             Voucher.bodyText(r, "iquantity") = Null2Something(Rs.Fields("出库数量"))
                    '                             Voucher.bodyText(r, "inum") = Null2Something(Rs.Fields("出库件数"))

                    SetDomValue oDomBody, "dmadedate", Null2Something(rs.Fields("生产日期"))
                    SetDomValue oDomBody, "dvdate", Null2Something(rs.Fields("失效日期"))
                    SetDomValue oDomBody, "dexpirationdate", Null2Something(rs.Fields("有效期计算项"))
                    SetDomValue oDomBody, "cexpirationdate", Null2Something(rs.Fields("有效期至"))
                    
                    
                    SetDomValue oDomBody, "imassdate", Null2Something(rs.Fields("保质期"))
                    SetDomValue oDomBody, "cmassunit", Null2Something(rs.Fields("保质期单位"))
                    SetDomValue oDomBody, "iexpiratdatecalcu", Null2Something(rs.Fields("有效期推算方式"))


                    '如果跟踪型存货有多条记录，显示参照窗口。
                Else
                   ' Call VoucherbodyBrowUser(Voucher, r, c, retvalue, referpara)
                End If

            End If
            rs.Close
            Set rs = Nothing
        End If

        '出库跟踪入库校验 chenliangc
    Case "cinvouchcode"

        If sWhCode = "" Then
            sErrMsg = GetString("U8.DZ.JA.Res640")
            retvalue = ""
            Exit Function
        End If
            
        retvalue = CCrmStr(Null2Something(retvalue, ""))

        If retvalue = "" Then
            SetDomValue oDomBody, "cbatch", ""
            SetDomValue oDomBody, "cvouchcode", ""
            SetDomValue oDomBody, "dmadedate", ""
            SetDomValue oDomBody, "dvdate", ""
            SetDomValue oDomBody, "dexpirationdate", ""
            SetDomValue oDomBody, "cexpirationdate", ""
            BodyCellCheck = True
            Exit Function
        End If

        Dim sInVouchCode As String                         '入库单号
        Dim sRdsID As String                               '入库单行ID


        '入库单号
        sInVouchCode = GetDomValue(oDomBody, "cinvouchcode")
        

        '本次调拨数量
        'iquantity = ConvertStrToDbl(GetDomValue(oDomBody, "iquantity"))

        If 1 = 1 Then
            If sInVouchCode <> "" Then
                Dim moBatchPst As New BatchPst
                moBatchPst.login = mologin

                '871和872调用库存函数时，传递的参数类型发生变化
                '871的订单行参数要求是整数型，872的改成字符型，因此需要单独处理
                If gU8Version = "872" Then
                    moBatchPst.List sSql, sWhCode, sInvCode, _
                            GetDomValue(oDomBody, "cfree1"), _
                            GetDomValue(oDomBody, "cfree2"), _
                            GetDomValue(oDomBody, "cfree3"), _
                            GetDomValue(oDomBody, "cfree4"), _
                            GetDomValue(oDomBody, "cfree5"), _
                            GetDomValue(oDomBody, "cfree6"), _
                            GetDomValue(oDomBody, "cfree7"), _
                            GetDomValue(oDomBody, "cfree8"), _
                            GetDomValue(oDomBody, "cfree9"), _
                            GetDomValue(oDomBody, "cfree10"), _
                            sInVouchCode, sRdsID, , GetDomValue(oDomBody, "cbatch"), CLng(sDemandType), sDemandCode, ""
                Else
                    moBatchPst.List sSql, sWhCode, sInvCode, _
                            GetDomValue(oDomBody, "cfree1"), _
                            GetDomValue(oDomBody, "cfree2"), _
                            GetDomValue(oDomBody, "cfree3"), _
                            GetDomValue(oDomBody, "cfree4"), _
                            GetDomValue(oDomBody, "cfree5"), _
                            GetDomValue(oDomBody, "cfree6"), _
                            GetDomValue(oDomBody, "cfree7"), _
                            GetDomValue(oDomBody, "cfree8"), _
                            GetDomValue(oDomBody, "cfree9"), _
                            GetDomValue(oDomBody, "cfree10"), _
                            sInVouchCode, sRdsID, , GetDomValue(oDomBody, "cbatch"), CLng(sDemandType), lDemandCode, ""
                End If
                Set rs = New ADODB.Recordset
                rs.Open sSql, g_Conn, 1, 1
                If rs.RecordCount = 0 Then
                    ReDim varArgs(0)
                    varArgs(0) = GetDomValue(oDomBody, "cinvcode")
                    sErrMsg = GetStringPara("U8.DZ.JA.Res750", varArgs(0))
                    '                    MsgBox "跟踪型存货" & Voucher.bodyText(r, "cinvcode") & "指定的入库单号" & sInVouchCode & "不存在，请参照输入入库单号。", vbInformation, GetString("U8.DZ.JA.Res030")
                    retvalue = ""
                    SetDomValue oDomBody, "rdsid", ""
                Else
                    '如果跟踪型存货只有一条记录
                    If rs.RecordCount >= 1 Then
                        SetDomValue oDomBody, "rdsid", rs("入库系统编号")
                        '带出自由项
                        For i = 0 To rs.Fields.Count - 1
                            sFreeName = IIf(IsNull(rs.Fields(i).Properties("BASECOLUMNNAME")), "", rs.Fields(i).Properties("BASECOLUMNNAME"))
                            If LCase(Mid(sFreeName, 1, 5)) = "cfree" Then
                                SetDomValue oDomBody, LCase(sFreeName), Null2Something(rs.Fields(i).Value)
                            End If
                        Next

                        '带出批次信息
                        If Not IsNull(GetDomValue(oDomBody, "cbatch")) Then
                            SetDomValue oDomBody, "cbatch", Null2Something(rs.Fields("批号"))
                        End If


                        SetDomValue oDomBody, "iquantity", Null2Something(rs.Fields("出库数量"))
                        SetDomValue oDomBody, "inum", Null2Something(rs.Fields("出库件数"))

                        SetDomValue oDomBody, "dmadedate", Null2Something(rs.Fields("生产日期"))
                        SetDomValue oDomBody, "dvdate", Null2Something(rs.Fields("失效日期"))
                        SetDomValue oDomBody, "dexpirationdate", Null2Something(rs.Fields("有效期计算项"))
                        SetDomValue oDomBody, "cexpirationdate", Null2Something(rs.Fields("有效期至"))
                        
                        SetDomValue oDomBody, "imassdate", Null2Something(rs.Fields("保质期"))
                        SetDomValue oDomBody, "cmassunit", Null2Something(rs.Fields("保质期单位"))
                        SetDomValue oDomBody, "iexpiratdatecalcu", Null2Something(rs.Fields("有效期推算方式"))

                        SetDomValue oDomBody, "cinvouchcode", Null2Something(rs.Fields("入库单号"))

                        '如果跟踪型存货有多条记录，显示参照窗口。
                    Else
                        'Call VoucherbodyBrowUser(Voucher, r, c, retvalue, referpara)

                    End If
                End If
                rs.Close
                Set rs = Nothing
            End If
        End If
    Case "cinva_unit"
        '
        Dim strGrid As String

        strGrid = " Select  cComUnitcode,ccomUnitName,ComputationUnit.cGroupCode,iChangRate,case when bMainUnit=1 then '是' else '否' end bMainUnit " & _
                " From computationUnit " & _
                " Left Join ComputationGroup on ComputationUnit.cGroupCode=ComputationGroup.cGroupCode " & _
                " Where ComputationUnit.cGroupCode = N'" & GetDomValue(oDomBody, "cgroupcode") & "'and (computationUnit.ccomUnitName like '" & retvalue & "%' or computationUnit.cComUnitcode like '" & retvalue & "%') " & _
                " Order by cComUnitcode ,iNumber ASC "
        Set rst = New ADODB.Recordset
        rst.Open strGrid, g_Conn

        '        If Voucher.bodyText(Row, "cincotermcode") = "" Then
        If rst.RecordCount > 0 Then
            retvalue = rst.Fields("ccomUnitName").Value
            SetDomValue oDomBody, "cunitid", rst.Fields("cComUnitcode").Value
            SetDomValue oDomBody, "cinva_unit", rst.Fields("ccomUnitName").Value
            SetDomValue oDomBody, "iinvexchrate", rst.Fields("iChangRate").Value
            If GetDomValue(oDomBody, "iquantity") <> "" And CDbl(GetDomValue(oDomBody, "iquantity")) <> 0 Then
                SetDomValue oDomBody, "inum", CDbl(GetDomValue(oDomBody, "iquantity")) / rst.Fields("iChangRate").Value
            ElseIf GetDomValue(oDomBody, "inum") <> "" And CDbl(GetDomValue(oDomBody, "inum")) <> 0 Then
                SetDomValue oDomBody, "iquantity", CDbl(GetDomValue(oDomBody, "inum")) * rst.Fields("iChangRate").Value
            End If
        Else
            SetDomValue oDomBody, "cinva_unit", ""
            retvalue = ""
        End If
        rst.Close

End Select

lsuccess:
'bChanged = success
 BodyCellCheck = True

Exit Function

Err_Handler:
'MsgBox Err.Description, vbInformation, GetString("U8.DZ.JA.Res030")
BodyCellCheck = False
sErrMsg = Err.Description
End Function

Private Sub SetBodyCellValue(ByRef oDomHead As DOMDocument, oDomBody As DOMDocument, rs As ADODB.Recordset)
    Dim i As Integer
    Dim iele As IXMLDOMElement
   
    '/*B*/ 根据单据模板需要的字段赋值

    For Each iele In oDomBody.selectNodes("//z:row")
        iele.setAttribute "cinvcode", rs("cinvcode") & ""
        iele.setAttribute "cinvname", rs("cInvName")
        iele.setAttribute "cinvstd", rs("cInvStd") & ""
        iele.setAttribute "cinvaddcode", rs("cInvAddCode") & ""
        iele.setAttribute "igrouptype", rs("iGroupType") & ""
        iele.setAttribute "cgroupcode", rs("cGroupCode") & ""
        iele.setAttribute "cgroupname", rs("cGroupName") & ""
        iele.setAttribute "ccomunitcode", rs("cComUnitCode") & ""    '主计量单位
        iele.setAttribute "ccomunitname", rs("cComunitName") & ""
        iele.setAttribute "cunitid", rs("cAssUnit") & ""   '辅计量单位编码
        iele.setAttribute "cinva_unit", rs("cAssUnitName") & ""    '辅计量单位名称
        iele.setAttribute "iinvexchrate", rs("iinvexchrate") & ""    '换算率
        iele.setAttribute "cmassunit", rs("cMassUnit") & ""
        iele.setAttribute "imassdate", rs("imassdate") & ""
        iele.setAttribute "itaxrate", 17                   '税率,默认17
        iele.setAttribute "kl", 100                        '扣率
        iele.setAttribute "kl2", 100                       '扣率2
        iele.setAttribute "iExpiratDateCalcu", Null2Something(rs("iExpiratDateCalcu"))    ' 有效期推算方式

        If iele.getAttribute("iquantity") <> "" Then iele.setAttribute "iquantity", ""
        If iele.getAttribute("inum") <> "" Then iele.setAttribute "inum", ""
        If iele.getAttribute("backdate") = "" Then
            iele.setAttribute "backdate", g_oLogin.CurDate
            If GetDomValue(oDomHead, "ddate") <> "" Then
                If CDate(GetDomValue(oDomHead, "ddate")) > g_oLogin.CurDate Then
                    iele.setAttribute "backdate", CDate(GetDomValue(oDomHead, "ddate"))
                End If
            End If
        End If
        '        If Voucher.bodyText(iRow, "iquantity") <> "" Then iele.setAttribute "iquantity", ""
        '        If Voucher.bodyText(iRow, "inum") <> "" Then iele.setAttribute "inum", ""

'        If Voucher.VoucherStatus = VSeAddMode Then
'            Voucher.bodyText(iRow, sAutoId) = iRow
'        Else
'            Call GetMaxIDs
'            Voucher.bodyText(iRow, sID) = lngVoucherID
'            Voucher.bodyText(iRow, sAutoId) = sAutoId
'            '退回主表id,避免浪费id号
'            '            g_Conn.Execute "update ufsystem..ua_identity  set Ifatherid=" & sID - 1 & " where cacc_id='" & g_oLogin.cAcc_Id & "' and cvouchtype='" & gstrCardNumber & "'"
'        End If

        '自由项
        For i = 1 To 10
            iele.setAttribute "cfree" & i, ""
        Next i

        '是否销售定价
        For i = 1 To 10
            iele.setAttribute "bsalepricefree" & i, Null2Something(rs("bSalePriceFree" & i))
        Next i

        '存货自定义项
        For i = 1 To 16
            iele.setAttribute "cInvDefine" & i, Null2Something(rs("cInvDefine" & i)) & ""
        Next i

       
    Next

    Set iele = Nothing
   
End Sub

Private Function myCheckNumValue(ByVal retvalue As Variant, ByVal cCaption As String, ByRef sErrMsg As String, Optional ByVal ctype As Byte = 0) As Boolean
    On Error GoTo lerr

    myCheckNumValue = False
    If retvalue = "" Then
        If ctype = 0 Then
            myCheckNumValue = True
            Exit Function
        Else
            sErrMsg = cCaption & GetString("U8.DZ.JA.Res650")
            Exit Function
        End If
    End If

    If IsNumeric(retvalue) Then
        If CDbl(retvalue) < 0 Then
            sErrMsg = cCaption & GetString("U8.DZ.JA.Res650")
            Exit Function
        End If
    Else
        sErrMsg = cCaption & "数据不合法"
        Exit Function
    End If

lTrue:
    myCheckNumValue = True
    Exit Function
lerr:
End Function

'检查合法性 dxb
Private Function myCheckNumValue2(ByVal retvalue As Variant, ByVal cCaption As String, ByRef sErrMsg As String, Optional ByVal ctype As Byte = 0) As Boolean
    On Error GoTo lerr

    myCheckNumValue2 = False
    
    If retvalue = "" Then
        If ctype = 0 Then
            myCheckNumValue2 = True
            Exit Function
        Else
            sErrMsg = cCaption & GetString("U8.DZ.JA.Res650")
            Exit Function
        End If
    End If

    If IsNumeric(retvalue) Then
        If CDbl(retvalue) < 0 Then
            sErrMsg = cCaption & GetString("U8.DZ.JA.Res660")
            Exit Function
        End If
    Else
        sErrMsg = cCaption & "数据不合法"
        Exit Function
    End If

lTrue:
    myCheckNumValue2 = True
    Exit Function
lerr:
End Function


Public Function Verify(ByVal sVouchId As String, ByRef sErrMsg As String, _
                              Optional ByVal sTimeStamp As String = "", _
                              Optional ByRef cnnFrom As ADODB.Connection) As Boolean
                              
                              
                              
                              
                              
    On Error GoTo Err_Handler
    
    
    Verify = False
    
    
    
    If cnnFrom Is Nothing Then
        
    Else
        Set g_Conn = cnnFrom
    End If
    
    Dim strSql As String
    Dim m_bTrans As Boolean
    m_bTrans = False
    
    'CRM有自己的权限控制
'    If ZwTaskExec(g_oLogin, AuthVerify, 1, True) = False Then
'       sErrMsg = "无审核权限！"
'       Exit Function
'    End If
    
    
    
    If ExecFunCompareUfts(sVouchId, sTimeStamp, sErrMsg) = False Then
        Exit Function
    End If
    
     
    
'    Call ZwTaskExec(g_oLogin, AuthVerify, 0)
    
    
    
    Dim lngAct As Long
    
    If cnnFrom Is Nothing Then
      g_Conn.BeginTrans
      m_bTrans = True
    End If
    
    
    g_Conn.Execute "update " & MainTable & " set " & StrcHandler & "='" & g_oLogin.cUserName & "' , " & StrdVeriDate & "='" & g_oLogin.CurDate & "' , " & StriStatus & "=2 " & vbCrLf & _
            " where convert(nchar,convert(money,ufts),2)  =" & IIf(sTimeStamp = "", "convert(nchar,convert(money,ufts),2)", sTimeStamp) & " and " & HeadPKFld & "=" & sVouchId, lngAct

    If lngAct = 0 Then

        If m_bTrans Then
            g_Conn.RollbackTrans
        End If
        sErrMsg = GetString("U8.DZ.JA.Res290")
        GoTo Exit_Label
    End If
    
    '库存未启用，借出借用单审核即写入累计出库数量
    If Not bststarted Then
        strSql = "update a set a.iQtyOut = isnull(a.iquantity,0), a.iQtyOut2 = isnull(a.inum,0)  from HY_DZ_Borrowouts a with (updlock) where a.id=" & sVouchId
        
        g_Conn.Execute strSql, lngAct
        If lngAct = 0 Then
            If m_bTrans Then
                g_Conn.RollbackTrans
                m_bTrans = False
            End If
            sErrMsg = GetString("U8.DZ.JA.Res290")
            GoTo Exit_Label
        End If
    End If
    
    If m_bTrans Then
    
        '业务通知
        NotifySrvSend "HYJCGH001", "HYJCGH001.Audit", CStr(sVouchId), g_oLogin
        g_Conn.CommitTrans
        
    End If
    
    Verify = True

Exit_Label:
    On Error GoTo 0
    Exit Function
Err_Handler:

    If m_bTrans Then
        g_Conn.RollbackTrans
        sErrMsg = VBA.Err.Description
    End If
    
End Function



Public Function UnVerify(ByVal sVouchId As String, ByRef sErrMsg As String, _
                              Optional ByVal sTimeStamp As String = "", _
                              Optional ByRef cnnFrom As ADODB.Connection) As Boolean
                              
                              
                              
                              
                              
    On Error GoTo Err_Handler
    
    
    UnVerify = False
    
    Dim strSql As String
    
    If cnnFrom Is Nothing Then
        
    Else
        Set g_Conn = cnnFrom
    End If
    
    Dim m_bTrans As Boolean
    m_bTrans = False
    
    
    Dim sVouchType As String
    Dim oRd As New ADODB.Recordset
    oRd.CursorLocation = adUseClient
    
    oRd.Open "select ccreatetype from  HY_DZ_BorrowOut where id =" & sVouchId, g_Conn, adOpenDynamic, adLockReadOnly
    
    If oRd.RecordCount > 0 Then
       sVouchType = vFieldVal(oRd.Fields("ccreatetype"))
    Else
       sVouchType = ""
    End If
    
    Set oRd = Nothing
    
    '库存没启用时，也当成期初处理
    If sVouchType = "期初单据" Or Not bststarted Then
        If VoucherIsCreate2(sVouchId) Then
            sErrMsg = GetString("U8.DZ.JA.Res250")
            Exit Function
        End If

    Else

        If VoucherIsCreate(sVouchId) Then
            sErrMsg = GetString("U8.DZ.JA.Res250")
            Exit Function
        End If
    End If
                
'    If VoucherIsCreate(sVouchId) Then
'        sErrMsg = GetString("U8.DZ.JA.Res250")
'        Exit Function
'    End If
           
    
    'CRM有自己的权限控制
'    If ZwTaskExec(g_oLogin, AuthUnVerify, 1, True) = False Then
'       sErrMsg = "无弃审权限！"
'       Exit Function
'    End If
    
    
    
    If ExecFunCompareUfts(sVouchId, sTimeStamp, sErrMsg) = False Then
        Exit Function
    End If
    
     
    
'    Call ZwTaskExec(g_oLogin, AuthVerify, 0)
    
    
    
    Dim lngAct As Long
    
    If cnnFrom Is Nothing Then
      g_Conn.BeginTrans
      m_bTrans = True
    End If
    
   
    g_Conn.Execute "update " & MainTable & " set " & StrcHandler & "=null , " & StrdVeriDate & "=null , " & StriStatus & "=1 " & vbCrLf & _
            " where convert(nchar,convert(money,ufts),2)  =" & IIf(sTimeStamp = "", "convert(nchar,convert(money,ufts),2)", "N'" & sTimeStamp & "'") & " and " & HeadPKFld & "=" & sVouchId, lngAct

    If lngAct = 0 Then

        If m_bTrans Then
            g_Conn.RollbackTrans
        End If
        sErrMsg = GetString("U8.DZ.JA.Res290")
        GoTo Exit_Label
    End If

    '库存未启用，借出借用单弃审即清空累计出库数量
    If Not bststarted Then
        strSql = "update a set a.iQtyOut = 0, a.iQtyOut2 = 0  from HY_DZ_Borrowouts a with (updlock) where a.id=" & sVouchId
        
        g_Conn.Execute strSql, lngAct
        If lngAct = 0 Then
            If m_bTrans Then
                g_Conn.RollbackTrans
                m_bTrans = False
            End If
            sErrMsg = GetString("U8.DZ.JA.Res290")
            GoTo Exit_Label
        End If
    End If
    
    If m_bTrans Then
        
        '业务通知
        NotifySrvSend "HYJCGH001", "HYJCGH001.UnAudit", CStr(sVouchId), g_oLogin
        g_Conn.CommitTrans
        
    End If
    UnVerify = True
    
Exit_Label:
    On Error GoTo 0
    Exit Function
Err_Handler:

    If m_bTrans Then
        g_Conn.RollbackTrans
        sErrMsg = VBA.Err.Description
    End If
    
End Function


'读取时间戳，并与旧时间戳比较
Private Function ExecFunCompareUfts(ByVal sVoucherID As String, ByVal sOlderTimeStamp As String, ByRef sErrMsg As String) As Boolean

'读取时间戳
    Dim sTimeStamp As String
    
    If sOlderTimeStamp = "" Then
        ExecFunCompareUfts = True
        Exit Function
    End If
    
    sTimeStamp = GetTimeStamp(g_Conn, MainTable, CLng(sVoucherID))

    If sTimeStamp = RecordDeleted Then
        sErrMsg = "单据已被其他用户删除,不可修改"
        '        MsgBox "单据(" & strcode & ")已被其他用户删除,不可修改", vbInformation, getstring("U8.DZ.JA.Res030")
        ExecFunCompareUfts = False
        Exit Function
    ElseIf sTimeStamp = RecordError Then
        sErrMsg = "单据数据出现错误,请刷新"
        '         MsgBox "单据(" & strcode & ")数据出现错误,请刷新", vbInformation, getstring("U8.DZ.JA.Res030")
        ExecFunCompareUfts = False
        Exit Function
    ElseIf sOlderTimeStamp <> sTimeStamp Then
        sErrMsg = "单据已被其他用户修改,请刷新"
        '        MsgBox "单据(" & strcode & ")已被其他用户修改,请刷新", vbInformation, getstring("U8.DZ.JA.Res030")
        ExecFunCompareUfts = False
        Exit Function
    Else
        
        ExecFunCompareUfts = True
    End If
End Function



Public Function Delete(ByVal sVouchId As String, ByRef sErrMsg As String, _
                              Optional ByVal sTimeStamp As String = "", _
                              Optional ByRef cnnFrom As ADODB.Connection) As Boolean
                              
                              
                              
                              
                              
    On Error GoTo Err_Handler
    
    
    Delete = False
    
    
    
    
    If cnnFrom Is Nothing Then
        
    Else
        Set g_Conn = cnnFrom
    End If
    
    Dim m_bTrans As Boolean
    m_bTrans = False
    
   
    'CRM有自己的权限控制
'    If ZwTaskExec(g_oLogin, AuthDelete, 1, True) = False Then
'       sErrMsg = "无删除权限！"
'       Exit Function
'    End If
    
    '12.0支持扩展自定义项
    Dim skeyfld As String
    Dim skeySubfld As String
    Dim objExtend As Object
    Set objExtend = CreateObject("VoucherExtendService.ClsExtendServer")
    
    If ExecFunCompareUfts(sVouchId, sTimeStamp, sErrMsg) = False Then
        Exit Function
    End If
    
    
    
'    Call ZwTaskExec(g_oLogin, AuthVerify, 0)
    
    
    
    Dim lngAct As Long
    
    If cnnFrom Is Nothing Then
      g_Conn.BeginTrans
      m_bTrans = True
    End If
    
    
    Dim sMessage, sSource As String
    
    Dim strXML As String
    Dim objDoc As New MSXML2.DOMDocument
    Dim ErrDesc As String
   
    '12.0扩展自定义项-删除-先删表体，再删表头
    skeyfld = objExtend.getKeyid(g_Conn, MainTable & "_ExtraDefine")
    skeySubfld = objExtend.getKeyid(g_Conn, DetailsTable & "_ExtraDefine")

    '删除字表之前删除表体的扩展自定义项
    objExtend.SavebyDeleteById sVouchId, DetailsTable & "_ExtraDefine", skeyfld, g_Conn, , DetailsTable, skeySubfld
    '删除字表
    g_Conn.Execute "delete from " & DetailsTable & " where " & HeadPKFld & "=" & sVouchId

    '删除主表之前删除表头的扩展自定义项
    objExtend.SavebyDeleteById sVouchId, MainTable & "_ExtraDefine", skeyfld, g_Conn
    '删除主表
    g_Conn.Execute "delete from " & MainTable & " where convert(nchar,convert(money,ufts),2)  = " & IIf(sTimeStamp = "", "convert(nchar,convert(money,ufts),2)", "N'" & sTimeStamp & "'") & "  and  " & HeadPKFld & "=" & sVouchId, lngAct

    If lngAct = 0 Then

        If m_bTrans Then
            g_Conn.RollbackTrans
           
        End If
        sErrMsg = GetString("U8.DZ.JA.Res290")
        GoTo Exit_Label
    End If
    
      '附件删除
    '*******************************
    Dim m_oServer As New UFVoucherServer85.clsVoucherTemplate
    Dim strAcc As String
    strAcc = "<Data AccessoriesChanged='1' Deleted='1' VoucherTypeID='" & gstrCardNumber & "' VoucherID='" & sVouchId & "'/>"
    m_oServer.SaveAccessories strAcc, g_Conn, sErrMsg
    Set m_oServer = Nothing
    
    If sErrMsg <> "" Then
        If m_bTrans Then
            g_Conn.RollbackTrans
            Exit Function
        End If
    Else
        If m_bTrans Then
            g_Conn.CommitTrans
            m_bTrans = False
        End If
    End If

 
  '*******************************
   Delete = True
Exit_Label:
    On Error GoTo 0
    Exit Function
Err_Handler:

    If m_bTrans Then
        g_Conn.RollbackTrans
        sErrMsg = VBA.Err.Description
    End If
  
    GoTo Exit_Label
End Function



Public Function SaveVouch(ByRef oDomHead As DOMDocument, ByRef oDomBody As DOMDocument, ByVal bInsert As Boolean, ByRef sErrMsg As String, _
                         ByRef iVouchID As Long, Optional ByRef cnnFrom As ADODB.Connection) As Boolean
    On Error GoTo Err_Handler

    Dim HeadData As New CDataVO
    Dim BodyData As New CDataVO
    Dim strSql As String
    Dim DMO As CDMO
    Dim Rult As CResult
    Dim rs As New ADODB.Recordset
    Dim i As Integer
    Dim sID2 As Long
    Dim strVoucherNo As String
    Dim bRepeatRedo As Boolean
    
    Dim bReGetNO As Boolean
    
    'by liwqa 并发
    Dim m_bTrans As Boolean
    Dim lngAct As Long
    
    '12.0支持扩展自定义项
    Dim skeyfld As String
    Dim skeySubfld As String
    Dim objExtend As Object
    Dim fldextends As ADODB.Fields
    Set objExtend = CreateObject("VoucherExtendService.ClsExtendServer")
    Dim oDomLine As IXMLDOMNodeList
    Dim oHeadElement As IXMLDOMElement
    Dim oBodyElement() As IXMLDOMElement

    Dim isSavedOK As Boolean

    '    Me.Voucher.RemoveEmptyRow   '清除空行

    isSavedOK = False                                      'by zhangwchb 20110829 提交保存

CHECK:
    
    Dim oEleNode As IXMLDOMNode
    
    Set oEleNode = oDomBody.selectSingleNode("//rs:data")
    
    Dim oEle As IXMLDOMElement
    
    For Each oEle In oEleNode.selectNodes("//z:row")
        If GetEleVal(oEle, "cinvcode") = "" And (UCase(GetEleVal(oEle, "editprop")) = "A" Or GetEleVal(oEle, "editprop") = "") Then
            If val(GetDomValue(oDomHead, "cborrowouttype")) = 0 Then
                oEleNode.removeChild oEle
            End If
        End If
    Next
    
    
    If cnnFrom Is Nothing Then
        
    Else
        Set g_Conn = cnnFrom
    End If

    '有效性校验
    If ExecFunSaveCheck(oDomHead, oDomBody, sErrMsg) = False Then
        SaveVouch = isSavedOK
        Exit Function
    End If

    '修改保存，需要比较时间戳，避免出现并发
    If bInsert = False Then
        If ExecFunCompareUfts(GetDomValue(oDomHead, "id"), GetDomValue(oDomHead, "ufts"), sErrMsg) = False Then
            SaveVouch = isSavedOK
            Exit Function
        End If
    End If


    '读取单据字段和数据
    ' 读取字段
    Set HeadData = GetHeadVouchData(g_Conn, oDomHead, oDomBody, MainTable)
    Set BodyData = GetBodyVouchData(g_Conn, oDomHead, oDomBody, DetailsTable)


    If cnnFrom Is Nothing Then
        g_Conn.BeginTrans
        m_bTrans = True
    End If
    
    '新增保存,获得最大id,autoid
    '修改保存时,不需要更新id
    'dxb  2009 6 15


    '    g_Conn.Execute "update ufsystem..ua_identity set ifatherid=(select isnull(max(ID),1) from " & DetailsTable & _
         '                   ") , ichildid=(select isnull(max(autoid),1) from " & DetailsTable & _
         '                   ") where cvouchtype='" & gstrCardNumber & "' and cAcc_Id ='" & g_oLogin.cAcc_Id & "'"

    '    If Voucher.VoucherStatus = VSeAddMode Then
    '      Call GetMaxID
    '    End If
    '    strSql = "select * from ufsystem..ua_identity where cvouchtype='" & gstrCardNumber & "' and cAcc_Id ='" & g_oLogin.cAcc_Id & "'"
    '    Set rs = g_Conn.Execute(strSql)
    '    sID = rs.Fields("iFatherId").Value
    '    sAutoId = rs.Fields("iChildID").Value

    If bInsert = False Then
         lngVoucherID = GetDomValue(oDomHead, "id")
    End If
    
    If bInsert Then '新增
        Call GetMaxID
        '        Call GetMaxID
        HeadData.Item(1).Item("ID").Value = sID
    End If

    gcCreateType = GetDomValue(oDomHead, "ccreatetype")
    
    '更新单据号流水号
    'Dim oDomHead As New DOMDocument
    Dim oDomFormat As DOMDocument
    Dim sError As String
   ' Set oDomHead = Voucher.GetHeadDom
    If Not BOGetVoucherNO(g_Conn, oDomHead, gstrCardNumber, sError, GetDomValue(oDomHead, strcCode), oDomFormat, False, , , True) Then
        If m_bTrans Then
            g_Conn.RollbackTrans
            m_bTrans = False
        End If
        Err.Raise 0, GetString("U8.DZ.JA.Res100")
        GoTo Err_Handler
    End If
    
    '12.0表头扩展自定义项
    Set oHeadElement = oDomHead.selectSingleNode("//z:row")
    skeyfld = objExtend.getKeyid(g_Conn, MainTable & "_ExtraDefine")
    skeySubfld = objExtend.getKeyid(g_Conn, DetailsTable & "_ExtraDefine")

    '保存表头
    HeadData.Item(1).Item("dmDate").Value = g_oLogin.CurDate    'Now() '制单时间 Format(Now(), "YYYY-MM-DD HH:MM:SS")
    HeadData.Item(1).Item("iStatus").Value = 1             '状态
    Set DMO = New CDMO
    '新增
    If bInsert Then
        

        Set Rult = DMO.Insert(g_Conn, HeadData)
        
        '12.0表头扩展自定义项-新增
        oHeadElement.setAttribute skeyfld, sID
        Set fldextends = objExtend.getVoucherExtendSaveInfo(g_Conn, MainTable & "_ExtraDefine")
        objExtend.SavebyInsert oHeadElement, MainTable & "_ExtraDefine", g_Conn, fldextends, , skeyfld
        
        '修改
    Else
        '修改前做并发
        g_Conn.Execute "update " & MainTable & " set " & HeadPKFld & " = " & HeadPKFld & vbCrLf & _
                " where convert(nchar,convert(money,ufts),2)  = " & GetDomValue(oDomHead, "ufts") & " and " & HeadPKFld & "=" & lngVoucherID, lngAct

        If lngAct = 0 Then

            If m_bTrans Then
                g_Conn.RollbackTrans
                m_bTrans = False
            End If
            sErrMsg = GetString("U8.DZ.JA.Res290")
            SaveVouch = isSavedOK
            Exit Function
        End If
        
'        If rs.State = adStateOpen Then rs.Close
'        rs.Open "select " & strcCode & " from " & MainTable & " where " & HeadPKFld & " = " & lngVoucherID, g_Conn, adOpenForwardOnly, adLockReadOnly
'        If Not rs.EOF Then
'            If Voucher.headerText(strcCode) <> rs.Fields(strcCode).Value Then
'                cSysBarCode = objBillBarCode.GeneralBarCode(g_Conn, g_oLogin.cAcc_Id, gstrCardNumber, Voucher.headerText(strcCode), "jc", "st1", cSplit)
'                HeadData.Item(1).Item("csysbarcode").Value = cSysBarCode
'            Else
'                cSysBarCode = objBillBarCode.GeneralBarCode(g_Conn, g_oLogin.cAcc_Id, gstrCardNumber, Voucher.headerText(strcCode), "jc", "st1", cSplit)
'                cSysBarCode = Voucher.headerText("csysbarcode")
'            End If
'
'        End If
        
        Set Rult = DMO.Update(g_Conn, HeadData)
        
        '12.0表头扩展自定义项-update
        Set fldextends = objExtend.getVoucherExtendSaveInfo(g_Conn, MainTable & "_ExtraDefine")
        objExtend.SavebyUpdate oHeadElement, MainTable & "_ExtraDefine", skeyfld, g_Conn, fldextends
    End If

    If Rult.Succeed = False Then

        If m_bTrans Then
            g_Conn.RollbackTrans
            m_bTrans = False
        End If
        sErrMsg = Rult.MsgCode & "," & GetString("U8.DZ.JA.Res110")
        SaveVouch = isSavedOK
        Exit Function
    End If

    'debug
    'oDomBody.Save "e:\sdsds.xml"
    '新增保存,获得最大id,autoid
    '修改保存时,不需要更新id
    If bInsert Then
        'g_Conn.Execute " select 'BodyData.Count = " & BodyData.Count & "'"
        ReDim oBodyElement(BodyData.Count)
        For i = 1 To BodyData.Count
            Call GetMaxIDs
            BodyData.Item(i).Item("AutoId").Value = sAutoId
            BodyData.Item(i).Item("ID").Value = sID
            If GetDomValue(oDomHead, "cborrowouttype") = "0" Or GetDomValue(oDomHead, "cborrowouttype") = "2" Then
               BodyData.Item(i).Item("bborrowoutback").Value = 1
            End If
            '            sAutoId = sAutoId + 1
            '12.0表体扩展自定义项-新增
            Set oDomLine = oDomBody.selectNodes("//z:row")
            'g_Conn.Execute "select  'oDomLine.length" & oDomLine.Length & "'"
            If oDomLine.Length > 0 Then
                Set oBodyElement(i - 1) = oDomLine(i - 1)
                oBodyElement(i - 1).setAttribute skeySubfld, sAutoId
                oBodyElement(i - 1).setAttribute skeyfld, sID
            End If
        Next i
        
        '更新ua_identity表中的字表标识,确保与单据字表autoid一致
        '        g_Conn.Execute "update ufsystem..ua_identity  set ifatherid=" & sID & ",ichildid=" & sAutoId - 1 & " where cacc_id='" & g_oLogin.cAcc_Id & "' and cvouchtype='" & gstrCardNumber & "'"

        Set Rult = DMO.Insert(g_Conn, BodyData)
        '12.0表体扩展自定义项-新增
        Set fldextends = objExtend.getVoucherExtendSaveInfo(g_Conn, DetailsTable & "_ExtraDefine")
        
        For i = 0 To UBound(oBodyElement) - 1
            objExtend.SavebyInsert oBodyElement(i), DetailsTable & "_ExtraDefine", g_Conn, fldextends, , skeySubfld
        Next i
        
        '修改保存,不更新id
    Else
        Dim j As Integer: j = 2
        '        sID2 = sID
        sID = HeadData.Item(1).Item("ID").Value
        
        ReDim oBodyElement(BodyData.Count)

        For i = 1 To BodyData.Count
            Set oDomLine = oDomBody.selectNodes("//z:row")
            'g_Conn.Execute "select  'oDomLine.length" & oDomLine.Length & "'"
            If oDomLine.Length > 0 Then
                Set oBodyElement(i - 1) = oDomLine(i - 1)
            End If
            
            If Null2Something(BodyData.Item(i).Item("AutoID").Value) = "" Then
                Call GetMaxIDs
                BodyData.Item(i).Item("AutoId").Value = sAutoId
                
                '新增的行才需更新ID和AutoID
                oBodyElement(i - 1).setAttribute skeySubfld, sAutoId
                oBodyElement(i - 1).setAttribute skeyfld, HeadData.Item(1).Item("ID").Value
            End If
            BodyData.Item(i).Item("ID").Value = HeadData.Item(1).Item("ID").Value
            
            If GetDomValue(oDomHead, "cborrowouttype") = "0" Or GetDomValue(oDomHead, "cborrowouttype") = "2" Then
               BodyData.Item(i).Item("bborrowoutback").Value = 1
            End If
                
        Next i

        '更新ua_identity表中的字表标识,确保与单据字表autoid一致
        '        g_Conn.Execute "update ufsystem..ua_identity  set ifatherid=" & sID2 & ", ichildid=" & sAutoId & " where cacc_id='" & g_oLogin.cAcc_Id & "' and cvouchtype='" & gstrCardNumber & "'"

        Set Rult = DMO.Insert(g_Conn, BodyData)
        
        '删除字表之前先删除表体的扩展自定义项
        objExtend.SavebyDelete oHeadElement, DetailsTable & "_ExtraDefine", skeyfld, g_Conn, , DetailsTable, skeySubfld
        Set Rult = DMO.DeleteByWhere(g_Conn, DetailsTable, "id=" & lngVoucherID)
        
        Set Rult = DMO.Insert(g_Conn, BodyData)
        
        '12.0表体扩展自定义项-update
        Set fldextends = objExtend.getVoucherExtendSaveInfo(g_Conn, DetailsTable & "_ExtraDefine")
        
        For i = 0 To UBound(oBodyElement) - 1
            '先删后查的方式
            'objExtend.SavebyUpdate oBodyElement(i), DetailsTable & "_ExtraDefine", skeySubfld, g_Conn, fldextends
            objExtend.SavebyInsert oBodyElement(i), DetailsTable & "_ExtraDefine", g_Conn, fldextends, , skeySubfld
        Next i

    End If



    If Rult.Succeed = False Then

        If m_bTrans Then
            g_Conn.RollbackTrans
            m_bTrans = False
        End If
        sErrMsg = Rult.MsgCode & "," & GetString("U8.DZ.JA.Res110")
        SaveVouch = isSavedOK
        Exit Function
    End If


 '表头附件保存
 
     
 

    Dim m_oServer As New UFVoucherServer85.clsVoucherTemplate
    Dim strAccessory As String
    Dim ErrDesc As String
    Dim ele As IXMLDOMElement
    Dim domAccessory As New DOMDocument
    Set ele = oDomHead.selectSingleNode("//accessories")
    If Not ele Is Nothing Then
        If domAccessory.loadXML(ele.Text) = True Then
            '处理存货的调整入，调整出
            '发出商品调整单修改
            domAccessory.documentElement.setAttribute "VoucherID", GetDomValue(oDomHead, "id")
           
            strAccessory = domAccessory.xml
            Call m_oServer.SaveAccessories(strAccessory, g_Conn, sErrMsg)
             If sErrMsg <> "" Then

                    If m_bTrans Then
                        g_Conn.RollbackTrans
                        m_bTrans = False
                    End If
                    
                    SaveVouch = isSavedOK
                    Exit Function
             End If
        End If
    End If
 

    'update HY_DZ_Borrowins set iqtyout=iquantity,iqtyout2=inum where id=11
    'enum by modify
    If gcCreateType = "期初单据" Then
        strSql = "update " & DetailsTable & " set iqtyout=iquantity,iqtyout2=inum where id = " & sID    '" and cCreateType = '期初单据'"
        g_Conn.Execute strSql
    End If


    '检验单据编号
    strVoucherNo = GetDomValue(oDomHead, "cCODE")
    bReGetNO = False
    
retry:
    If rs.State = adStateOpen Then rs.Close
    If bInsert Then
        rs.Open "select cCode from " & MainTable & " where cCode='" & strVoucherNo & "' and id<>" & sID, g_Conn, 1, 1
    Else
        rs.Open "select cCode from " & MainTable & " where cCode='" & strVoucherNo & "' and ID<>" & GetDomValue(oDomHead, "ID"), g_Conn, 1, 1
    End If
    
    
    If Not rs.EOF Then
        If Not BOGetVoucherNO(g_Conn, oDomHead, gstrCardNumber, sError, strVoucherNo, oDomFormat, False, False, bRepeatRedo, False) Then
            If m_bTrans Then
                g_Conn.RollbackTrans
                m_bTrans = False
            End If
            sErrMsg = sError
            SaveVouch = isSavedOK
            Exit Function
        End If

        If bRepeatRedo Then
            
            g_Conn.Execute "Update " & MainTable & " set cCode='" & strVoucherNo & "' where ID='" & sID & "'"
           
            GoTo retry:
        Else
            If m_bTrans Then
                g_Conn.RollbackTrans
                m_bTrans = False
            End If
            sErrMsg = GetString("U8.DZ.JA.Res1080")
            SaveVouch = isSavedOK
            Exit Function
        End If

        '        Voucher.headerText("cCODE") = strVoucherNo
        '        g_Conn.Execute "Update " & MainTable & " set cCode='" & strVoucherNo & "' where ID='" & sID & "'"
        '        Call BOGetVoucherNO(g_Conn, oDomHead, gstrCardNumber, sError, strVoucherNo, oDomFormat, False, , , True)

        bReGetNO = True
    Else
        If Not BOGetVoucherNO(g_Conn, oDomHead, gstrCardNumber, sError, strVoucherNo, oDomFormat, False, False, True, True) Then
            If m_bTrans Then
                g_Conn.RollbackTrans
                m_bTrans = False
            End If
            sErrMsg = sError
            SaveVouch = isSavedOK
            Exit Function
        Else
            '            If Voucher.headerText("cCODE") <> strVoucherNo Then
            '                Voucher.headerText("cCODE") = strVoucherNo
            '                g_Conn.Execute "Update " & MainTable & " set cCode='" & strVoucherNo & "' where ID='" & sID & "'"
            '            End If
        End If
    End If
    rs.Close

    Dim tmpDomHead As New DOMDocument
    Dim tmpDomBody As New DOMDocument
    Dim MaxRowNO As Long
    Dim t As Integer
    
    
    Dim cSysBarCode As String
    Dim cSplit As String
    Dim objBillBarCode As Object
    Set objBillBarCode = CreateObject("UFBillComponent.clsVoucherBarCodeSrv")
    If bInsert Or (bInsert = False And bReGetNO) Then
        cSysBarCode = objBillBarCode.GeneralBarCode(g_Conn, g_oLogin.cAcc_Id, gstrCardNumber, strVoucherNo, "jc", "st1", cSplit)
        g_Conn.Execute "update " & MainTable & " set csysbarcode=N'" & cSysBarCode & "' where ID=N'" & sID & "'"
        For i = 1 To BodyData.Count
            g_Conn.Execute "update " & DetailsTable & " set cbsysbarcode=N'" & cSysBarCode & cSplit & CStr(i) & "' where ID=" & sID & " and autoid= " & BodyData.Item(i).Item("autoid").Value
        Next i
    Else
        cSysBarCode = objBillBarCode.GeneralBarCode(g_Conn, g_oLogin.cAcc_Id, gstrCardNumber, strVoucherNo, "jc", "st1", cSplit)
        cSysBarCode = HeadData.Item(1).Item("csysbarcode").Value
       
        MaxRowNO = objBillBarCode.GetMaxRowBarcodeNum(oDomBody, "cbsysbarcode")
        t = 1
        For i = 1 To BodyData.Count
            If Null2Something(BodyData.Item(i).Item("cbsysbarcode").Value) = "" Then
                g_Conn.Execute "update " & DetailsTable & " set cbsysbarcode=N'" & cSysBarCode & cSplit & CStr(MaxRowNO + t) & "' where ID=" & sID & " and autoid= " & BodyData.Item(i).Item("autoid").Value
                t = t + 1
            End If
        Next i
    End If
    
    
    
    
    '20110905
'    strSql = "update " & MainTable & " set iswfcontrolled = " & IIf(bInWorkFlow, 1, 0) & " where ID='" & sID & "'"
'    g_Conn.Execute strSql
    
    iVouchID = sID

    If m_bTrans Then
        g_Conn.CommitTrans
        m_bTrans = False
    End If

    '
    isSavedOK = True                                       'by zhangwchb 20110829 提交保存
    SaveVouch = isSavedOK
    Exit Function

Err_Handler:
    If m_bTrans Then
        g_Conn.RollbackTrans
        m_bTrans = False
    End If
    sErrMsg = Err.Description & "," & GetString("U8.DZ.JA.Res110")

    SaveVouch = isSavedOK
End Function
Private Function bInWorkFlow() As Boolean
    Dim AuditServiceProxy As Object
    Dim objCalledContext As Object
    Dim strErr As String
    Dim blnSubmit As Boolean

    Set objCalledContext = CreateObject("UFSoft.U8.Framework.LoginContext.CalledContext")
    objCalledContext.SubId = g_oLogin.cSub_Id
    objCalledContext.TaskId = g_oLogin.TaskId
    objCalledContext.token = g_oLogin.userToken
    Set AuditServiceProxy = CreateObject("UFIDA.U8.Audit.ServiceProxy.AuditServiceProxy")
    blnSubmit = AuditServiceProxy.isflowenabled(gstrCardNumber, gstrCardNumber & ".Submit", objCalledContext, strErr)

    If strErr = "" Then
        bInWorkFlow = blnSubmit
    End If

    Set objCalledContext = Nothing
    Set AuditServiceProxy = Nothing
End Function


'以下为检验函数
'字段有效性检查
 Private Function ExecFunSaveCheck(oDomHead As DOMDocument, oDomBody As DOMDocument, ByRef sErrMsg As String) As Boolean
    Dim strMsgHead As String
    Dim strMsgBody As String

    On Error GoTo Err_Handler
    
    
    ExecFunSaveCheck = False


   

    
    '检查表体数据的有效性,主要检查:
    '必填项是否空
    '自由项组合是否合法
    '批次交验
    '入库单号校验
    '项目编码校验
    '有效期校验

    If ExecFunEffectiveCheck(oDomHead, oDomBody, sErrMsg) = False Then
        ExecFunSaveCheck = False
        Exit Function
    End If

    ExecFunSaveCheck = True
    Exit Function

Err_Handler:
    sErrMsg = Err.Description

End Function


'检查表体数据的有效性,主要检查:
'必填项是否空
'自由项组合是否合法
'批次交验
'入库单号校验
'项目编码校验
'有效期校验
Private Function ExecFunEffectiveCheck(ByRef oDomHead As DOMDocument, ByRef oDomBody As DOMDocument, ByRef sErrMsg As String) As Boolean
    On Error GoTo Err_Handler
    Dim sql  As String
    Dim rs As New ADODB.Recordset

    '表头校验:单据编号,制单人,单据日期为必填项,
    '仓库,部门等其他项目根据实际业务单据而定
    If GetDomValue(oDomHead, strcCode) = "" Then
        sErrMsg = GetString("U8.DZ.JA.Res1120")
        ExecFunEffectiveCheck = False
        Exit Function
    End If
    If GetDomValue(oDomHead, StrcMaker) = "" Then
        sErrMsg = GetString("U8.DZ.JA.Res1130")
        ExecFunEffectiveCheck = False
        Exit Function
    End If

    If GetDomValue(oDomHead, StrdDate) = "" Then
        sErrMsg = GetString("U8.DZ.JA.Res1140")
        ExecFunEffectiveCheck = False
        Exit Function
        '    Else
        '        '借用日期 不做大于制单日期控制，因为有补单需求
        '        If Voucher.VoucherStatus = VSeAddMode Then
        '        End If
    End If
    
    '检查部门有效性
    If GetDomValue(oDomHead, "cdepcode") <> "" Then
        sql = "select dDepEndDate from department where cdepcode='" + GetDomValue(oDomHead, "cdepcode") + "'"
        rs.Open sql, g_Conn
        If Not rs.EOF Or rs.BOF Then
            If Not IsNull(rs("dDepEndDate")) Then
                If DateDiff("d", CDate(rs("dDepEndDate")), CDate(GetDomValue(oDomHead, "ddate"))) >= 0 Then
                    sErrMsg = GetString("U8.ST.V870.00290")
                    ExecFunEffectiveCheck = False
                    Exit Function
                End If
            End If
        End If
        rs.Close
    End If
    
    gcCreateType = GetDomValue(oDomHead, "ccreatetype")
    
    '不启用库存，不校验（crm也需要做借出单据 for U8dp202841115）
    If gcCreateType = "期初单据" And bststarted Then
        sql = "select CVALUE  from AccInformation where csysid='ST' AND CNAME='dSTFirstDate'"
        rs.Open sql, g_Conn
        If Not rs.BOF Or Not rs.EOF Then
            If DateDiff("d", CDate(rs.Fields("CVALUE")), CDate(GetDomValue(oDomHead, "ddate"))) >= 0 Then
                sErrMsg = GetString("U8.DZ.JA.Res1860")
                ExecFunEffectiveCheck = False
                Exit Function
            End If
        End If
    End If


    If GetDomValue(oDomHead, "cfreight") = "是" Then
        If GetDomValue(oDomHead, "MycdefineT2") = "" Or GetDomValue(oDomHead, "cfreightType") = "" _
                Or GetDomValue(oDomHead, "cfreightCost") = "" Then
            sErrMsg = GetString("U8.DZ.JA.Res1150")
            ExecFunEffectiveCheck = False
            Exit Function
        End If
    End If

    '单据表体校验
    If ExecFunBodyCheck(oDomHead, oDomBody, sErrMsg) = False Then
        ExecFunEffectiveCheck = False
        Exit Function
    End If

    ExecFunEffectiveCheck = True
    Exit Function

Err_Handler:
    ExecFunEffectiveCheck = False
    sErrMsg = Err.Description

End Function

'表体有效性校验
'必填项是否空
'启用的自由项是否都输入了
'自由项组合是否合法
'批次交验
'入库单号校验
'有效期校验
Private Function ExecFunBodyCheck(ByRef oDomHead As DOMDocument, ByRef oDomBody As DOMDocument, ByRef sErrMsg As String) As Boolean
    On Error GoTo Err_Handler
    '表体有效性校验,由于表体是多行需要循环
    
    Dim bodyele As IXMLDOMElement
    Dim iRow As Integer
    Dim rs As New ADODB.Recordset
    Dim tRs As New ADODB.Recordset
    Dim sql As String
    Dim iSTConMode As Integer
    Dim iQ As Integer

    'Voucher.getVoucherDataXML oDomHead, oDomBody

    iRow = 0
    
    ExecFunBodyCheck = True

    For Each bodyele In oDomBody.selectNodes("//z:row[@editprop != 'D']")
        iRow = iRow + 1
        '临时赋值,作为行标
        '此处赋值不管用
        '    bodyele.setAttribute "AutoID", irow

        '1
        '比填项,简单校验：存货编码，数量
        '其他的校验根据实际业务而定
        If GetEleVal(bodyele, "cinvcode") = "" Then
            If val(GetDomValue(oDomHead, "cborrowouttype")) = 1 Then
                GoTo nexthandle
            End If
            ReDim varArgs(0)
            varArgs(0) = iRow
            sErrMsg = GetStringPara("U8.DZ.JA.Res1160", varArgs(0))
            '            MsgBox "第" & iRow & "行存货编码不能为空", vbInformation, GetString("U8.DZ.JA.Res030")
            ExecFunBodyCheck = False
            Exit Function
        End If

        If val(GetEleVal(bodyele, "iquantity")) = 0 And val(GetEleVal(bodyele, "inum")) = 0 Then
            '            MsgBox "第" & iRow & "行 数量件数不能同时为空或等于0", vbInformation, getstring("U8.DZ.JA.Res030")
            ReDim varArgs(0)
            varArgs(0) = iRow
            sErrMsg = GetStringPara("U8.DZ.JA.Res1170", varArgs(0))
            '            MsgBox "第" & iRow & "行 数量/件数不能同时为空或等于0", vbInformation, GetString("U8.DZ.JA.Res030")
           
            ExecFunBodyCheck = False
            Exit Function
        End If

          If gcCreateType = "期初单据" Then
                'dxb 仓库必须录入！
                If GetEleVal(bodyele, "cwhcode") = "" Then
                    ReDim varArgs(0)
                    varArgs(0) = iRow
                    MsgBox GetStringPara("U8.DZ.JA.Res1180", varArgs(0)), vbInformation, GetString("U8.DZ.JA.Res030")
                    'MsgBox "第" & iRow & "行仓库不能为空", vbInformation, GetString("U8.DZ.JA.Res030")
                    ExecFunBodyCheck = False
                    Exit Function
                End If
                sql = " select bwhpos from Warehouse where cwhcode='" & bodyele.getAttribute("cwhcode") & "' and bwhpos=1 "
                Set tRs = New ADODB.Recordset
                tRs.Open sql, g_Conn
                If Not tRs.BOF Or Not tRs.EOF Then
                    '检查仓库启用货位管理必须输入货位
                    If bodyele.getAttribute("cPosition") & "" = "" Then
                        ReDim varArgs(0)
                        varArgs(0) = iRow
                        sErrMsg = GetStringPara("U8.DZ.JA.Res1920", varArgs(0))
                        'MsgBox "第" & iRow & "行仓库不能为空", vbInformation, GetString("U8.DZ.JA.Res030")
                        ExecFunBodyCheck = False
                        Exit Function
                    End If
                End If
        End If

        '2
        '启用的自由项是否都输入了
        '结构性，非结构性自由项都必须输入
        '自由项组合是否合法
        If ExecFunFreeCheck(GetEleVal(bodyele, "cinvcode"), bodyele, iRow, sErrMsg) = False Then
            '标识当前行,着蓝色
           
            'Voucher.SetCurrentRow ("@AutoID=" & bodyele.getAttribute("AutoID") & "")
            ExecFunBodyCheck = False
            Exit Function
        End If


        '３
        '批次
        '出库跟踪入库

        If ExecFuncbatch(GetEleVal(bodyele, "cinvcode"), _
                GetEleVal(bodyele, "cbatch") & "", _
                GetEleVal(bodyele, "cinvouchcode") & "", _
                iRow, sErrMsg) = False Then
            '标识当前行,着蓝色
            'Voucher.row = iRow
            'Voucher.SetCurrentRow ("@AutoID=" & bodyele.getAttribute("AutoID") & "")
            ExecFunBodyCheck = False
            Exit Function
        End If
        '        If Voucher.headerText("dmDate") <> "" Then
        If gcCreateType = "期初单据" Then
            sql = " select * from warehouse where cwhcode='" & bodyele.getAttribute("cwhcode") & "' and (isnull(dWhEndDate,'')='' or datediff(d,'" & g_oLogin.CurDate & "', dWhEndDate) >0 )"
            Set rs = Nothing
            rs.Open sql, g_Conn, 1, 1
            If rs.EOF Then
                ReDim varArgs(0)
                varArgs(0) = iRow
                sErrMsg = GetStringPara("U8.DZ.JA.Res1190", varArgs(0))
                'MsgBox "第" & iRow & "单据日期大于等于仓库的失效日期", vbInformation, GetString("U8.DZ.JA.Res030")
                ExecFunBodyCheck = False
                Exit Function
            End If
        End If

        '        End If

        '        '4 保质期校验
        '         If ExecFunMassDate(bodyele.getAttribute("cinvcode"), _
                  '                            iRow, bodyele.getAttribute("cMassDateUnit") & "", _
                  '                            bodyele.getAttribute("imassdate") & "", _
                  '                            bodyele.getAttribute("dmadedate") & "", _
                  '                            bodyele.getAttribute("dvdate") & "") = False Then
        '             '标识当前行,着蓝色
        '            Voucher.row = iRow
        '            'Voucher.SetCurrentRow ("@AutoID=" & bodyele.getAttribute("AutoID") & "")
        '            ExecFunBodyCheck = False
        '            Exit Function
        '        End If
        '          iQ = 0
        '                Set Rs = Nothing
        '                SQL = " select iSTConMode from Warehouse where cWhCode ='" & bodyele.getAttribute("cwhcode") & "'"
        '                Rs.Open SQL, g_Conn, 1, 1
        '                iSTConMode = Rs!iSTConMode
        '                Set Rs = Nothing
        '                 SQL = " select *  From V_CurrentStock left join vendor v on v.cvencode=V_CurrentStock.cvmivencode  left join v_aa_enum v1 on v1.enumcode=v_currentstock.iexpiratdatecalcu and v1.enumtype=N'SCM.ExpiratDateCalcu' left join AA_BatchProperty batch on Batch.cinvcode=V_CurrentStock.cinvcode and isnull(Batch.cbatch,N'')=isnull(V_CurrentStock.cbatch,N'') and isnull(Batch.cfree1,N'')=isnull(V_CurrentStock.cfree1,N'') and isnull(Batch.cfree2,N'')=isnull(V_CurrentStock.cfree2,N'') and isnull(Batch.cfree3,N'')=isnull(V_CurrentStock.cfree3,N'') and isnull(Batch.cfree4,N'')=isnull(V_CurrentStock.cfree4,N'') and isnull(Batch.cfree5,N'')=isnull(V_CurrentStock.cfree5,"
        '                              SQL = SQL + "N'') and isnull(Batch.cfree6,N'')=isnull(V_CurrentStock.cfree6,N'') and isnull(Batch.cfree7,N'')=isnull(V_CurrentStock.cfree7,N'') and isnull(Batch.cfree8,N'')=isnull(V_CurrentStock.cfree8,N'') and isnull(Batch.cfree9,N'')=isnull(V_CurrentStock.cfree9,N'') and isnull(Batch.cfree10,N'')=isnull(V_CurrentStock.cfree10,N'') Where V_CurrentStock.cWhcode=N'" & bodyele.getAttribute("cwhcode") & "' And V_CurrentStock.cInvCode =N'" & bodyele.getAttribute("cinvcode") & "' And V_CurrentStock.cBatch= N'" & bodyele.getAttribute("cbatch") & "' And IsNull(V_CurrentStock.cBatch,N'')<>N''  And isnull( bstopflag,0)=0  And (ISNULL(isotype,0)= 0 And ISNULL(isodid,N'')= N'') and (iQuantity+IsNull(fInQuantity,0)-IsNull(fOutQuantity,0)-IsNull(fStopQuantity,0)-" & bodyele.getAttribute("iquantity") & ") >0 order by  V_CurrentStock.dvdate,V_CurrentStock.cbatch"
        '                 Rs.Open SQL, g_Conn, 1, 1
        '                 If Rs.EOF Then
        '                    iQ = -1
        '                 End If
        '                Select Case iSTConMode
        '                  Case 0
        '                        Set Rs = Nothing
        '                         SQL = "select cValue from accinformation where cname=N'bAllowZero' and csysid=N'ST'"
        '                        Rs.Open SQL, g_Conn, 1, 1
        '                        If Not Rs.EOF Then
        '                            If Rs!cvalue = "False" Then
        '                              If iQ = -1 Then
        '                                 MsgBox "产品[" + bodyele.getAttribute("cinvcode") + "]仓库[" + bodyele.getAttribute("cwhcode") + "]批次[" + bodyele.getAttribute("cbatch") + "]的数量不够借出", vbInformation, GetString("U8.DZ.JA.Res030")
        '                                 ExecFunBodyCheck = False
        '                                 Exit Function
        '                              End If
        '
        '                            End If
        '                        End If
        '                  Case 2
        '                       If iQ = -1 Then
        '                             MsgBox "产品[" + bodyele.getAttribute("cinvcode") + "]仓库[" + bodyele.getAttribute("cwhcode") + "]批次[" + bodyele.getAttribute("cbatch") + "]的数量不够借出", vbInformation, GetString("U8.DZ.JA.Res030")
        ''                                 ExecFunBodyCheck = False
        ''                                 Exit Function
        '                          End If
        '
        '                End Select

        '5 预计归还日期 dxb
        If GetEleVal(bodyele, "backdate") = "" Then
            ReDim varArgs(0)
            varArgs(0) = iRow
            sErrMsg = GetStringPara("U8.DZ.JA.Res1200", varArgs(0))
            'MsgBox "第" & iRow & "行预计归还日期不能为空", vbInformation, GetString("U8.DZ.JA.Res030")
            ExecFunBodyCheck = False
            Exit Function
        Else
            If CDate(Mid(GetEleVal(bodyele, "backdate"), 1, 10)) < CDate(Mid(GetDomValue(oDomHead, StrdDate), 1, 10)) Then
                ReDim varArgs(0)
                varArgs(0) = iRow
                sErrMsg = GetStringPara("U8.DZ.JA.Res1210", varArgs(0))
                'MsgBox "第" & iRow & "行预计归还日期不能早于借用日期", vbInformation, GetString("U8.DZ.JA.Res030")
                ExecFunBodyCheck = False
                Exit Function
            End If
        End If
        
nexthandle:

    Next

    If CheckNumOutLimit(oDomHead, oDomBody) = "" Then
        sErrMsg = GetString("U8.DZ.JA.Res1220")
        ExecFunBodyCheck = False
        Exit Function
    End If

    ExecFunBodyCheck = True
    Exit Function

Err_Handler:
    sErrMsg = Err.Description
End Function

Private Function CheckNumOutLimit(oDomHead As DOMDocument, oDomBody As DOMDocument) As String
    On Error GoTo lerr:
    CheckNumOutLimit = ""

    Dim i As Long
    
    Dim oEle As IXMLDOMElement
    
    For Each oEle In oDomBody.selectNodes("//z:row")
        '('iquantity','inum')
        If GetEleVal(oEle, "iquantity") <> "" Then
            If Len(GetEleVal(oEle, "iquantity")) > 15 Or CDbl(GetEleVal(oEle, "iquantity")) > 999999999999999# Then
                Exit Function
            End If
        End If

        If GetEleVal(oEle, "inum") <> "" Then
            If Len(GetEleVal(oEle, "inum")) > 15 Or CDbl(GetEleVal(oEle, "inum")) > 999999999999999# Then
                Exit Function
            End If
        End If
    Next

    CheckNumOutLimit = "OK"
lerr:

End Function





'****************************
'结构性自由项

'启用的自由项是否都输入了
'结构性，非结构性自由项都必须输入
'自由项组合是否合法
Private Function ExecFunFreeCheck(cinvcode As String, _
                                  bodyele As IXMLDOMElement, _
                                  iRow As Integer, ByRef sErrMsg As String _
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
    rs.Open sql, g_Conn, 1, 1
    If Not rs.EOF Then

        For i = 1 To 10
            If rs("bFree" & i) = True And bodyele.getAttribute("cfree" & i) = "" Then
                bRS.Open "select cItemName from userdef where cClass = '存货' and cDicDbname='cFree" & i & "'", g_Conn, 1, 1
                ReDim varArgs(2)
                varArgs(0) = iRow
                varArgs(1) = cinvcode
                varArgs(2) = bRS("cItemName")
                sErrMsg = GetStringPara("U8.DZ.JA.Res1280", varArgs(0), varArgs(1), varArgs(2))
                'MsgBox "第" & iRow & "行存货" & cinvcode & "启用了自由项" & bRS("cItemName") & ",必须输入", vbInformation, GetString("U8.DZ.JA.Res030")
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

        bRS.Open sFind, g_Conn, 1, 1
        If bRS.EOF Then
            ReDim varArgs(0)
            varArgs(0) = iRow
            sErrMsg = GetStringPara("U8.DZ.JA.Res1290", varArgs(0))
            'MsgBox "第" & iRow & "行存货结构性自由项组合不合法", vbInformation, GetString("U8.DZ.JA.Res030")
            bRS.Close
            Set bRS = Nothing
            rs.Close
            Set rs = Nothing
            ExecFunFreeCheck = False
            Exit Function
        End If


    Else
        ReDim varArgs(1)
        varArgs(0) = iRow
        varArgs(1) = cinvcode
        sErrMsg = GetStringPara("U8.DZ.JA.Res1300", varArgs(0), varArgs(1))
        'MsgBox "第" & iRow & "行存货编码" & cinvcode & "不存在", vbInformation, GetString("U8.DZ.JA.Res030")
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
    sErrMsg = Err.Description
End Function

'批次
'入库单号
Private Function ExecFuncbatch(cinvcode As String, cbatch As String, cTrackCode As String, iRow As Integer, ByRef sErrMsg As String) As Boolean
    On Error GoTo Err_Handler

    Dim sql As String
    Dim rs As New ADODB.Recordset

    sql = "select bInvBatch,bTrack from inventory where cinvcode='" & cinvcode & "'"
    rs.Open sql, g_Conn, 1, 1
    If Not rs.EOF Then
        If gcCreateType = "期初单据" And CBool(rs("bInvBatch")) And cbatch = "" Then

            ReDim varArgs(0)
            varArgs(0) = iRow
            sErrMsg = GetStringPara("U8.DZ.JA.Res1310", varArgs(0))
            ' MsgBox "第" & iRow & "行存货启用了批次管理，批号不能为空", vbInformation, GetString("U8.DZ.JA.Res030")
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
        ReDim varArgs(1)
        varArgs(0) = iRow
        varArgs(1) = cinvcode
        sErrMsg = GetStringPara("U8.DZ.JA.Res1300", varArgs(0), varArgs(1))
        'MsgBox "第" & iRow & "行存货编码" & cinvcode & "不存在", vbInformation, GetString("U8.DZ.JA.Res030")
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
    sErrMsg = Err.Description
End Function


'从单据控件获得表体数据对象
Private Function GetBodyVouchData(conn As Object, oDomHead As DOMDocument, oDomBody As DOMDocument, BTable As String) As CDataVO
    Dim i As Integer, j As Integer
    Dim DMO As CDMO
    Set DMO = New CDMO
    Dim DVO As CDataVO
    Set DVO = DMO.NewData(conn, BTable)
    Dim DataItem As CDataItem
    
    Dim oEle As IXMLDOMElement
    
    Dim bFirst As Boolean
    bFirst = True
    For Each oEle In oDomBody.selectNodes("//z:row[@editprop != 'D']")
        If bFirst Then
            For i = 1 To DVO.Item(1).Count
                DVO.Item(1).Item(i).Value = GetEleVal(oEle, DVO.Item(1).Item(i).FieldCode)
            Next i
            bFirst = False
        Else
            Set DataItem = DMO.GetCloneDataItem(DVO.Item(1))
            For i = 1 To DataItem.Count
                DataItem.Item(i).Value = GetEleVal(oEle, DVO.Item(1).Item(i).FieldCode)
            Next i
            DVO.Add DataItem
        End If
    Next
    Set GetBodyVouchData = DVO

End Function
'从单据控件获得表头数据对象
Private Function GetHeadVouchData(conn As Object, oDomHead As DOMDocument, oDomBody As DOMDocument, HTable As String) As CDataVO
    Dim i As Integer
    Dim DMO As CDMO
    Set DMO = New CDMO
    Dim DVO As CDataVO
    Set DVO = DMO.NewData(conn, HTable)
    For i = 1 To DVO.Item(1).Count
        DVO.Item(1).Item(i).Value = GetDomValue(oDomHead, DVO.Item(1).Item(i).FieldCode)
    Next i
    Set GetHeadVouchData = DVO
End Function


Private Sub GetExtendedInfo(ByRef conn As Connection, ByVal CardNum As String, _
                            ByRef sExtendHeadField As String, ByRef sExtendHeadJoinSQL As String, _
                            ByRef sExtendBodyField As String, ByRef sExtendBodyJoinSQL As String)

    Dim oExtend  As Object
    Set oExtend = CreateObject("VoucherExtendService.ClsExtendServer")
    Call oExtend.GetExtendInfo(conn, CardNum, "T", sExtendHeadField, sExtendHeadJoinSQL)
    Call oExtend.GetExtendInfo(conn, CardNum, "B", sExtendBodyField, sExtendBodyJoinSQL)
    Set oExtend = Nothing
End Sub




Public Function LoadVoucher(ByVal sWhere As String, ByRef oDomHead As DOMDocument, ByRef oDomBody As DOMDocument, ByRef sErrMsg As String) As Boolean

    On Error GoTo Err_Handler
    
    
    LoadVoucher = True
    
    
    Set oDomHead = New DOMDocument
    Set oDomBody = New DOMDocument
    Dim oelement As IXMLDOMElement
    Dim tSql As String

    Dim sql As String
    Dim oRecordset As New Recordset
    oRecordset.CursorLocation = adUseClient
    
    
    Dim sMessage As String

    ' * 错误源
    
    
    
    Dim sSource As String
    

   

    

    'by zhangwchb 20110718 扩展字段
    
  
    Dim sExtendHeadField As String
    Dim sExtendHeadJoinSQL As String
    Dim sExtendBodyField As String
    Dim sExtendBodyJoinSQL As String
    Call GetExtendedInfo(g_Conn, gstrCardNumber, sExtendHeadField, sExtendHeadJoinSQL, sExtendBodyField, sExtendBodyJoinSQL)

    sql = "SELECT *,'' AS editprop " & sExtendHeadField & vbCrLf _
            & " FROM  " & MainView & " " & sExtendHeadJoinSQL & vbCrLf _
            & " WHERE  (1=1)  " & IIf(sWhere = "", "", " and " & sWhere)

    If oRecordset Is Nothing Then _
            Set oRecordset = CreateObject("ADODB.Recordset")

    If oRecordset.State = adStateOpen Then _
            Call oRecordset.Close

    Call oRecordset.Open( _
            sql, _
            g_Conn, _
            adOpenStatic, _
            adLockReadOnly, _
            adCmdText)

    
    ' * 转换成 XML 数据格式
    oRecordset.Save oDomHead, adPersistXML

    'by zhangwchb 20110718 扩展字段
    sql = "SELECT *,'' AS editprop " & sExtendBodyField & vbCrLf _
            & "FROM " & DetailsView & " " & sExtendBodyJoinSQL & vbCrLf _
            & "WHERE (1=1) " & IIf(sWhere = "", "", " and " & sWhere)

    If oRecordset Is Nothing Then _
            Set oRecordset = CreateObject("ADODB.Recordset")

    If oRecordset.State = adStateOpen Then _
            Call oRecordset.Close

    Call oRecordset.Open( _
            sql, _
            g_Conn, _
            adOpenStatic, _
            adLockReadOnly, _
            adCmdText)


    ' * 转换成 XML 数据格式
    oRecordset.Save oDomBody, adPersistXML
    
Exit_Label:
  
    If Not oRecordset Is Nothing Then
        If oRecordset.State = adStateOpen Then _
                Call oRecordset.Close
    End If
    Set oRecordset = Nothing

    Exit Function
Err_Handler:
    

    If Not oRecordset Is Nothing Then
        If oRecordset.State = adStateOpen Then _
                Call oRecordset.Close
    End If
    Set oRecordset = Nothing
    LoadVoucher = False
    sErrMsg = "数据加载错误!"
   

End Function


'审批流提交与撤销
Public Function ExecSubmit(DoOrUndo As Boolean, table As String, pk As String, id As Long, ByRef errMsg As String) As Boolean
    Dim retDoUndoSubmit As Boolean
    Dim strErrorResId As String                            '审批流的错误信息870 added
    
    ExecSubmit = True
    
    Screen.MousePointer = vbHourglass
    Call CheckSubmit(table, pk, CStr(id))

    If CBool(IsWFControlled) And ((DoOrUndo And (iverifystate = 0 Or (iverifystate = 1 And ireturncount > 0))) Or (DoOrUndo = False And iverifystate <> 0)) Then

        retDoUndoSubmit = DoUndoSubmit(DoOrUndo, gstrCardNumber, CStr(id), table, ufts, CBool(IsWFControlled), strErrorResId, vouchercode)
        If retDoUndoSubmit = False Then
            'MsgBox strErrorResId, vbInformation, GetString("U8.DZ.JA.Res030")
            ExecSubmit = False
            errMsg = strErrorResId
        Else
            If DoOrUndo Then
                'MsgBox GetString("U8.SA.xsglsql_2.saworkflowsrv.011"), vbInformation, GetString("U8.DZ.JA.Res030")    '"单据提交成功！"
                ExecSubmit = True
                errMsg = GetString("U8.SA.xsglsql_2.saworkflowsrv.011")
            Else
                'MsgBox GetString("U8.SA.xsglsql_2.saworkflowsrv.012"), vbInformation, GetString("U8.DZ.JA.Res030")    '撤销成功！
                ExecSubmit = True
                errMsg = GetString("U8.SA.xsglsql_2.saworkflowsrv.012")
            End If
        End If
    Else
        If DoOrUndo Then
            'MsgBox GetString("U8.SA.xsglsql_2.saworkflowsrv.001"), vbInformation, GetString("U8.DZ.JA.Res030")    '"该单据已经提交或者未启用审批流！"
            ExecSubmit = False
            errMsg = GetString("U8.SA.xsglsql_2.saworkflowsrv.001")
        Else
            'MsgBox GetString("U8.SA.xsglsql_2.saworkflowsrv.002"), vbInformation, GetString("U8.DZ.JA.Res030")    '"该单据已经撤销或者未启用审批流！"
            ExecSubmit = False
            errMsg = GetString("U8.SA.xsglsql_2.saworkflowsrv.002")
        End If
    End If
    Screen.MousePointer = vbDefault
End Function

' 审批流赋值，返回当前记录是否进入工作流
Public Function CheckSubmit(MainTable As String, pk As String, voucherID As String) As Boolean
    Dim strSql As String, rs As ADODB.Recordset
    strSql = "select isnull(id,0) VoucherId,isnull(iverifystate,0) as iverifystate,CONVERT(nchar,CONVERT(money,ufts),2) as ufts,isnull(iswfcontrolled,0) as iswfcontrolled,isnull(ccode,0) vouchercode,isnull(ireturncount,0) as ireturncount from " & MainTable & " where " & pk & "= " & voucherID & ""
    Set rs = New ADODB.Recordset
    rs.Open strSql, g_Conn
    If Not rs.EOF Then
        iverifystate = rs!iverifystate
        '        vstate = Rs!vstate
        ufts = rs!ufts
        IsWFControlled = rs!IsWFControlled
        CheckSubmit = CBool(IsWFControlled)
        vouchercode = rs!vouchercode
        ireturncount = rs!ireturncount
    End If
    rs.Close
End Function

'重新提交以及启用工作流审批时的操作
'HY_DZ_BorrowOut_FrmVoucher"
Public Function ExecRequestAudit(guid As String)
    SendShowViewMessage guid, "UFIDA.U8.Audit.AuditViews.TreatTaskViewPart"
    SendMessgeToPortal "DocRequestAudit", guid
End Function
'关闭
Public Function ExecSubClose(ByVal sVouchId As String, ByRef sErrMsg As String, _
                              Optional ByVal sTimeStamp As String = "", Optional ByRef cnnFrom As ADODB.Connection) As Boolean
    On Error GoTo Err_Handler
    'by liwqa 并发
    Dim m_bTrans As Boolean
    Dim lngAct As Long
    
    ExecSubClose = False
    
    If cnnFrom Is Nothing Then
        
    Else
        Set g_Conn = cnnFrom
    End If
    
    If cnnFrom Is Nothing Then
        g_Conn.BeginTrans
        m_bTrans = True
    End If
    
    g_Conn.Execute "update " & MainTable & " set " & StrCloseUser & "='" & g_oLogin.cUserName & "' , " & StrdCloseDate & "='" & g_oLogin.CurDate & "' , " & StriStatus & "=4 " & vbCrLf & _
            " where convert(nchar,convert(money,ufts),2)  =" & IIf(sTimeStamp = "", "convert(nchar,convert(money,ufts),2)", sTimeStamp) & " and " & HeadPKFld & "=" & sVouchId, lngAct
 
    If lngAct = 0 Then
 
        If m_bTrans Then
            g_Conn.RollbackTrans
            m_bTrans = False
        End If
        sErrMsg = GetString("U8.DZ.JA.Res290")
         ExecSubClose = False
        GoTo Exit_Label
    End If
 
    If m_bTrans Then
        g_Conn.CommitTrans
        m_bTrans = False
    End If
 
ExecSubClose = True
 
Exit_Label:
    On Error GoTo 0
    
    Exit Function
Err_Handler:
 
    If m_bTrans Then
        g_Conn.RollbackTrans
        m_bTrans = False
    End If
 
End Function
'打开
Public Function ExecSubOpen(ByVal sVouchId As String, ByRef sErrMsg As String, _
                              Optional ByVal sTimeStamp As String = "", Optional ByRef cnnFrom As ADODB.Connection) As Boolean
 
    '状态:如果生单人不为空,置为3,即已生单;如果生单人为空:
    '                                                   如果审核人不为空,置为2,即审核;如果审核人为空,则置为1,即新建
 
    On Error GoTo Err_Handler
    'by liwqa 并发
    Dim m_bTrans As Boolean
    Dim lngAct As Long
    
    ExecSubOpen = False
    
    If cnnFrom Is Nothing Then
        
    Else
        Set g_Conn = cnnFrom
    End If
    
    If cnnFrom Is Nothing Then
        g_Conn.BeginTrans
        m_bTrans = True
    End If
    g_Conn.Execute "update " & MainTable & " set " & StrCloseUser & "=null , " & StrdCloseDate & "=null, " & StriStatus & "=case when isnull(" & StrIntoUser & ",N'') <>N'' then 3 else case when isnull(" & StrcHandler & ",N'')<>N'' then 2 else 1 end end " & vbCrLf & _
            " where convert(nchar,convert(money,ufts),2)  = " & IIf(sTimeStamp = "", "convert(nchar,convert(money,ufts),2)", sTimeStamp) & " and " & HeadPKFld & "=" & sVouchId, lngAct
 
    If lngAct = 0 Then
 
        If m_bTrans Then
            g_Conn.RollbackTrans
            m_bTrans = False
        End If
        sErrMsg = GetString("U8.DZ.JA.Res290")
         ExecSubOpen = False
        GoTo Exit_Label
    End If
 
    If m_bTrans Then
        g_Conn.CommitTrans
        m_bTrans = False
    End If
 
'    Call ExecSubRefresh
ExecSubOpen = True
Exit_Label:
    On Error GoTo 0
    Exit Function
Err_Handler:
 
    If m_bTrans Then
        g_Conn.RollbackTrans
        m_bTrans = False
    End If
End Function


Public Function HasUserAuth(ologin As Object, user As String, auth As String) As Boolean
      
    
    On Error GoTo ErrHandle
    
    HasUserAuth = False
    
    Dim userStr As String
    Dim arr As Variant
   
    Dim i As Long
 
    user = CCrmStr(user)
   
    If ologin.isAdmin Then
        HasUserAuth = True
        Exit Function
    End If
    
    Dim oRow As Object
      
    Set oRow = CreateObject("U8RowAuthsvr.clsRowAuth")
   

    If oRow.Init(ologin.UfDbName, ologin.cUserId, False) = False Then
        
        Exit Function
    End If

   
    userStr = oRow.getAuthString("user", "", UCase(auth))

    Set oRow = Nothing
    
    If userStr = "" Then
       HasUserAuth = True
    ElseIf LCase(userStr) = "null" Then     'cann't remove it
       HasUserAuth = False
    Else
        arr = Split(userStr, ",")
        For i = LBound(arr) To UBound(arr)
            If LCase(Trim(arr(i))) = "'" & LCase(user) & "'" Or LCase(Trim(arr(i))) = "n'" & LCase(user) & "'" Then
                HasUserAuth = True
                Exit For
            End If
        Next i
    End If
    
    Exit Function
      
ErrHandle:
     VBA.Err.Clear
     HasUserAuth = False
End Function

Public Function GetBillNumberChecksucceed(ByRef oDomHead As DOMDocument, ByRef strVoucherNo As String, ByRef sErrMsg As String) As Boolean
   ' Dim oDomHead As DOMDocument
    Dim oDomFormat As DOMDocument
    Dim oelement As IXMLDOMElement
    Dim bManualCode As Boolean
    Dim bCanModyCode As Boolean
 
    Dim sKey As String
 
    Dim sError As String
   GetBillNumberChecksucceed = False
   
 
    If GetVoucherNO(g_Conn, oDomHead, gstrCardNumber, sError, strVoucherNo, oDomFormat, True, , False) = True Then
 
        Set oelement = oDomFormat.selectSingleNode("//单据编号")
 
        '支持完全手工编号
        '允许手工修改得含义为 完全手工编号， 重号自动重取的含义为 手工修改，重号自动重取
        bManualCode = oelement.getAttribute("允许手工修改")
        bCanModyCode = oelement.getAttribute("允许手工修改") Or oelement.getAttribute("重号自动重取")
    Else
        sErrMsg = sError
        GetBillNumberChecksucceed = False
        Exit Function
    End If
 
    '支持完全手工编号，此时不取单据号 2003-07-16 黄朝阳
    If Not bManualCode Then
 
            sKey = strcCode
 
            If GetVoucherNO(g_Conn, oDomHead, gstrCardNumber, sError, strVoucherNo, , , , False) = False Then
                sErrMsg = sError
               GetBillNumberChecksucceed = False
               Exit Function
            End If
 
    End If
    GetBillNumberChecksucceed = True
End Function

Function GetSTStartDate(ByRef strDate As String) As Boolean
    Dim strSql As String
    Dim rst As New ADODB.Recordset
    
    strDate = ""
    GetSTStartDate = True
    strSql = "select isnull(cvalue,'') as cvalue  from AccInformation where csysid='ST' AND CNAME='dSTFirstDate'"
    If rst Is Nothing Then Set rst = CreateObject("ADODB.Recordset")
    If rst.State = adStateOpen Then Call rst.Close
    Call rst.Open(strSql, g_Conn, adOpenStatic, adLockReadOnly, adCmdText)
    
    If Not rst.BOF Or Not rst.EOF Then
        strDate = rst.Fields("cvalue")
        If Null2Something(strDate, "") = "" Then
            GetSTStartDate = False
        End If
    Else
        GetSTStartDate = False
    End If
End Function
