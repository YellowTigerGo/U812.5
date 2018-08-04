Attribute VB_Name = "Fun"

' 审批流
Dim iverifystate As Integer
Dim ufts As String
Dim IsWFControlled As Boolean                              '是否审批流控制
Dim vstate As Integer                                      '是否终审标志
Dim vouchercode As String                                  '单据号
Dim ireturncount As Integer                                '单据被退回次数
Dim flag As Boolean                                        '存货参照多选

Public strwhereVou  As String
Dim NotifySend As Object '业务通知服务



'获取主表标志id
Public Function GetMaxID()
    On Error GoTo Err_Handler
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim cmd As New ADODB.Command

    With cmd
        .ActiveConnection = g_Conn
        .CommandType = adCmdStoredProc
        .CommandText = "sp_GetID"
        .CommandTimeout = 0
        .Prepared = True

        .Parameters.Append .CreateParameter("RemoteId", adVarChar, adParamInput, 2, "00")
        .Parameters.Append .CreateParameter("cAcc_Id", adVarChar, adParamInput, 100, g_oLogin.cAcc_Id)
        .Parameters.Append .CreateParameter("cVouchType", adVarChar, adParamInput, 50, gstrCardNumber)
        .Parameters.Append .CreateParameter("iAmount", adBigInt, adParamInput, 100, 1)
        .Parameters.Append .CreateParameter("iFatherId", adBigInt, adParamOutput, 100, sID)
        .Parameters.Append .CreateParameter("iChildId", adBigInt, adParamOutput, 100, sAutoId)
        .Execute
    End With

    sID = cmd.Parameters("iFatherId")
    '    sAutoId = cmd.Parameters("iChildId")

    Set cmd = Nothing

    Exit Function
Err_Handler:
    Set cmd = Nothing
    MsgBox Err.Description, vbInformation, GetString("U8.DZ.JA.Res030")
End Function

'获取主表标志id
Public Function GetMaxIDs()
    On Error GoTo Err_Handler
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim cmd As New ADODB.Command

    With cmd
        .ActiveConnection = g_Conn
        .CommandType = adCmdStoredProc
        .CommandText = "sp_GetID"
        .CommandTimeout = 0
        .Prepared = True

        .Parameters.Append .CreateParameter("RemoteId", adVarChar, adParamInput, 2, "00")
        .Parameters.Append .CreateParameter("cAcc_Id", adVarChar, adParamInput, 100, g_oLogin.cAcc_Id)
        .Parameters.Append .CreateParameter("cVouchType", adVarChar, adParamInput, 50, "hy_DZ_BorrowOuts")
        .Parameters.Append .CreateParameter("iAmount", adBigInt, adParamInput, 100, 1)
        .Parameters.Append .CreateParameter("iFatherId", adBigInt, adParamOutput, 100, sID)
        .Parameters.Append .CreateParameter("iChildId", adBigInt, adParamOutput, 100, sAutoId)
        .Execute
    End With

    sAutoId = cmd.Parameters("iFatherId")
    '    sAutoId = cmd.Parameters("iChildId")

    Set cmd = Nothing

    Exit Function
Err_Handler:
    Set cmd = Nothing
    MsgBox Err.Description, vbInformation, GetString("U8.DZ.JA.Res030")
End Function

'单据编号是否可编辑
Public Function GetbCanModifyVCode(manual As Boolean) As Boolean
    Dim oEle As IXMLDOMElement
    Dim myobj As New UFBillComponent.clsBillComponent
    Dim xmlDOMObj As New DOMDocument
    Dim strTemp As String

    '初始化单据编号规则
    myobj.InitBill g_Conn, gstrCardNumber
    Set xmlDOMObj = New DOMDocument
    strTemp = myobj.GetBillFormat
    '    m_sVouchRuler = strTemp
    xmlDOMObj.loadXML strTemp

    Set oEle = xmlDOMObj.selectSingleNode("//单据编号")

    If LCase(oEle.getAttribute("允许手工修改")) = "true" Then manual = True

    If LCase(oEle.getAttribute("允许手工修改")) = "true" Or LCase(oEle.getAttribute("重号自动重取")) = "true" Then
        GetbCanModifyVCode = True
    End If

End Function

'手工输入时,校验是否存在
Public Function CheckCellValue(sql As String) As Boolean
    On Error GoTo Err_Handler

    Dim rs As New ADODB.Recordset
    rs.Open sql, g_Conn, 1, 1
    If Not rs.EOF Then
        '当前输入的是编码,返回名称
        '当前输入的是名称,返回编码
        strCellCode = rs("code")
        strCellName = rs("name")
        CheckCellValue = True
    Else
        strCellCode = ""
        strCellName = ""
        CheckCellValue = False
    End If


    rs.Close
    Set rs = Nothing
    Exit Function

Err_Handler:
    rs.Close
    Set rs = Nothing
    CheckCellValue = False
    MsgBox Err.Description, vbInformation, GetString("U8.DZ.JA.Res030")
End Function

'客户档案校验,校验是否存在
Public Function CheckCustomer(sql As String, cName As String, Address As String) As Boolean
    On Error GoTo Err_Handler

    Dim rs As New ADODB.Recordset
    rs.Open sql, g_Conn, 1, 1
    If Not rs.EOF Then
        '当前输入的是编码,返回名称
        '当前输入的是名称,返回编码
        strCellCode = rs("cCusCode") & ""
        strCellName = rs("cCusAbbName") & ""
        cName = rs("cCusName") & ""                        '客户名称
        Address = rs("cCusAddress") & ""                   '客户地址

        CheckCustomer = True
    Else
        strCellCode = ""
        strCellName = ""
        cName = ""                                         '客户名称
        Address = ""                                       '客户地址
        CheckCustomer = False
    End If


    rs.Close
    Set rs = Nothing
    Exit Function

Err_Handler:
    rs.Close
    Set rs = Nothing
    CheckCustomer = False
    MsgBox Err.Description, vbInformation, GetString("U8.DZ.JA.Res030")
End Function

'客户档案校验,校验是否存在
Public Function CheckVendor(sql As String, vencode As String, venname As String) As Boolean
    On Error GoTo Err_Handler

    Dim rs As New ADODB.Recordset
    rs.Open sql, g_Conn, 1, 1
    If Not rs.EOF Then
        '当前输入的是编码,返回名称
        '当前输入的是名称,返回编码
        vencode = rs("cvencode") & ""
        venname = rs("cvenabbname") & ""
        strCellCode = rs("cvencode") & ""
        strCellName = rs("cvenabbname") & ""

        CheckVendor = True
    Else
        vencode = ""
        venname = ""
        CheckVendor = False
    End If


    rs.Close
    Set rs = Nothing
    Exit Function

Err_Handler:
    rs.Close
    Set rs = Nothing
    CheckVendor = False
    MsgBox Err.Description, vbInformation, GetString("U8.DZ.JA.Res030")
End Function
'运行过滤条件
Public Function GetFilterList(ologin As U8Login.clsLogin, Optional o_Filter As Object = Nothing, Optional sMenuPubFilter As String) As Boolean

    On Error GoTo Err_Handler
    
    Dim lngReturn As Long
    Dim objfltint As New UFGeneralFilter.FilterSrv
    Dim filtername, filtersa As String                     '过滤条件

    filtername = "PD010301"
    filtersa = "EF"
    
    '11.0列表发布 wangfb 2012-06-11
    Dim filterItf As New UFGeneralFilter.FilterSrv
    Dim sError As Variant
    Dim iRet As Boolean
    If o_Filter Is Nothing Then
        '原来只是简单的赋值1=2，发布菜单后要根据方案取值
        'strWhere = " (1=2) "
        '不是菜单发布调用的直接退出 11.0 进入列表不谈过滤
        If sMenuPubFilter = "" Then
'            strWhere = " (1=2) "
            GetFilterList = True
            Exit Function
        Else
            '11.0菜单发布直接传入解决方案id：sMenuPubFilter,然后过滤条件自动隐藏。wangfb
            filterItf.InitSolutionID = sMenuPubFilter
            '11.0菜单发布 bHiddenFilter作为参数(默认false)传入，wangfb
            iRet = filterItf.OpenFilter(ologin, "", filtername, filtersa, sError)
            If iRet Then
                Dim flt As UFGeneralFilter.FilterItem
                Dim strCondition As String
                strWhere = filterItf.GetSQLWhere
                GetFilterList = True
            Else
'                strWhere = " (1=2) "
                GetFilterList = True
            End If
        End If
        Exit Function
    Else
       Set objfltint = o_Filter
    End If
    
   
'    Dim FilterSrv As New clsUserControl
'    Set objfltint.BehaviorObject = FilterSrv
'
'    If o_Filter Is Nothing Then
'        lngReturn = objfltint.OpenFilter(ologin, "", filtername, filtersa)
'
'        If lngReturn = False Then
'            GetFilterList = False
'            '        strWhere = ""
'            Exit Function
'        End If
'    Else
'        Set objfltint = o_Filter
'    End If
    
    strWhere = objfltint.GetSQLWhere

    'by zhangwchb 20110720 增加过滤条件“是否提交”
 

'    If strWhere = "" Then
'        strWhere = sAuth_AllList                           'Replace(sAuth_AllList, "HY_DZ_BorrowOut.id", "V_HY_DZ_BorrowOutList.id")
'    Else
'        strWhere = strWhere & " and " & sAuth_AllList      'Replace(sAuth_ALL, "HY_DZ_BorrowOut.id", "V_HY_DZ_BorrowOutList.id")
'    End If

    '处理过滤条件
    '状态
    Call FilteriStatus(strWhere)


    GetFilterList = True

    Exit Function

Err_Handler:
    MsgBox Err.Description, vbInformation, GetString("U8.DZ.JA.Res030")

End Function

'运行过滤条件
Public Function GetFilter(ologin As U8Login.clsLogin) As Boolean

    On Error GoTo Err_Handler

    Dim lngReturn As Long
    Dim objfltint As New UFGeneralFilter.FilterSrv
    Dim filtername, filtersa As String                     '过滤条件

    'filtername = "ST[__]借出借用单"
    filtername = "借出借用单参照"
    filtersa = "ST"
    
    lngReturn = objfltint.OpenFilter(ologin, "", filtername, filtersa)

    If lngReturn = False Then
        GetFilter = False
        '        strwhereVou = ""
        Exit Function
    End If

    strwhereVou = objfltint.GetSQLWhere

    '处理过滤条件
    '状态
    Call FilteriStatus(strwhereVou)


    GetFilter = True

    Exit Function

Err_Handler:
    MsgBox Err.Description, vbInformation, GetString("U8.DZ.JA.Res030")

End Function
Private Sub FilteriStatus(ByRef str As String)
    '全部

    'by zhangwchb 20110720 增加过滤条件“是否提交”
    str = Replace(str, "And (bSubmitted = N'1')", "")
    str = Replace(str, "And (bSubmitted = N'0')", "")

    str = Replace(str, "iPrintCount", "isNull(iPrintCount,0)")

    'enum by modify
    If InStr(1, str, "全部") > 0 Then
        str = Replace(str, "And (iStatus = N'全部')", "")

        '    ElseIf InStr(1, str, "新建") > 0 Then
        '        str = Replace(str, "新建", "1")
        '
        '    ElseIf InStr(1, str, "审核") > 0 Then
        '        str = Replace(str, "审核", "2")
        '
        ''    ElseIf InStr(1, str, "生单") > 0 Then
        ''        str = Replace(str, "生单", "3")
        '
        '    ElseIf InStr(1, str, "关闭") > 0 Then
        '        str = Replace(str, "关闭", "4")

    End If
End Sub

'更新当前页变量
Public Sub UpdatePageCurrent(iID As Long)
    Dim sql As String
    Dim rs As New ADODB.Recordset

    '    If tmpLinkTbl <> "" Then '单据联查 时 按钮状态控制 by zhangwchb 20110809
    '        sql = "select count(1) orderid from " & MainTable & _
             '                " inner join " & tmpLinkTbl & " on " & MainTable & "." & HeadPKFld & " = " & tmpLinkTbl & ".id " & _
             '                " where " & MainTable & "." & HeadPKFld & "<=" & iID & " and " & sAuth_ALL
    '
    '    Else
    sql = "select count(1) orderid from " & MainTable & " where " & HeadPKFld & "<=" & iID & " and " & sAuth_ALL
    '    End If
    rs.Open sql, g_Conn, 1, 1
    If Not rs.EOF Then
        PageCurrent = rs("orderid")
    Else
        PageCurrent = 1
    End If

    rs.Close
    Set rs = Nothing

End Sub



'获取数据权限
'constring 连接字符
'user 操作员
'cBusObId 业务对象 Department 部门 Customer 客户
'cClassCode 项目大类（项目专用）
'cFuncId 读写审核等权限

Public Function GetRowAuth(constring As String, user As String, _
                           cBusObId As String, Optional ByVal cClassCode As String, Optional ByVal cFuncId As String) As String
    '判断数据权限
    Dim oRow As New clsRowAuth
    Dim Ret As String

    Ret = ""


    On Error Resume Next


    If oRow.Init(constring, user, False) = False Then
        GetRowAuth = ""
        Exit Function
    End If

    '部门"Department""R"
    Ret = oRow.getAuthString(cBusObId, "", cFuncId)

    GetRowAuth = Ret


    Set oRow = Nothing

End Function

'conn 连接字符
'user 操作员
'cFuncId 读写审核等权限

'此处以销售订单为例
Public Function GetRowAuthAlls(conn As Connection, user As String, _
                               Optional ByVal cFuncId As String, Optional ByVal AuthID As String = "") As String


    Dim sRet As String
    Dim sql As String
    Dim strSql As String
    Dim rs As New ADODB.Recordset

    strSql = ""

    On Error Resume Next

    '判断销售管理需要控制哪些数据权限
    sql = "select * from  accinformation where csysid=N'sa' and  ((cname in(N'bAuth_Dep',N'bAuth_Per',N'bAuth_Inv',N'bAuth_Cus','bAuth_Wh') and cvalue='true') or " & _
            " (cname='bMaker' and cvalue='false')) "
    If AuthID <> "" Then
        sql = sql & " and cname ='" & AuthID & "'"
    End If

    If rs.State = adStateOpen Then Set rs = Nothing
    rs.Open sql, conn, 1, 1

    Do While Not rs.EOF

        Select Case rs("cName")
                '客户
            Case "bAuth_Cus"
                sRet = GetRowAuth(conn.ConnectionString, user, "Customer", "", cFuncId)
                If sRet <> "" And Trim(sRet) <> "1=2" Then
                    strSql = strSql & " AND (isnull(cCusCode,N'')=N'' or cCusCode in (select cCusCode from customer where iId in (" & sRet & ")))"
                End If

                '部门
            Case "bAuth_Dep"
                sRet = GetRowAuth(conn.ConnectionString, user, "Department", "", cFuncId)
                If sRet <> "" Then
                    strSql = strSql & " AND (isnull(cDepCode,N'')=N'' or cDepCode in (" & sRet & "))"
                End If

                '存货
                '            Case "bAuth_Inv"
                '                sret = GetRowAuth(conn.ConnectionString, user, "Inventory", "", cFuncId)
                '                If sret <> "" And Trim(sret) <> "1=2"  Then
                '                    strSql = strSql & " AND cinvcode in (" & sret & ")"
                '                End If


                '业务员
            Case "bAuth_Per"
                sRet = GetRowAuth(conn.ConnectionString, user, "Person", "", cFuncId)
                If sRet <> "" Then
                    strSql = strSql & " AND (isnull(cPersonCode,N'')=N'' or  cPersonCode in (" & sRet & "))"
                End If

                '仓库
            Case "bAuth_Wh"
                sRet = GetRowAuth(conn.ConnectionString, user, "Warehouse", "", cFuncId)
                If sRet <> "" Then
                    strSql = strSql & " AND (isnull(cwhcode,N'')=N'' or cwhcode in (" & sRet & "))"
                End If

                '操作员
            Case "bMaker"
                sRet = GetRowAuth(conn.ConnectionString, user, "user", "", cFuncId)
                If sRet <> "" Then
                    strSql = strSql & " AND (isnull(cMaker,N'')=N'' or cMaker in (" & sRet & "))"
                End If


        End Select


        rs.MoveNext
    Loop


    GetRowAuthAlls = strSql


    rs.Close
    Set rs = Nothing

End Function

Public Function VoucherheadBrowUser(Voucher As Object, ByVal Index As Variant, sRet As Variant, referpara As UAPVoucherControl85.ReferParameter)

    On Error GoTo errhander

    Dim sMetaXML As String
    Dim sHeadItemName As String
    Dim Authstr As String
    sMetaXML = "<Ref><RefSet bAuth='0'/></Ref>"
    referpara.ReferMetaXML = sMetaXML

    '获取当前编辑的字段名
    sHeadItemName = Voucher.ItemState(Index, siHeader).sFieldName

    If LCase(sHeadItemName) Like "cdefine*" Or LCase(sHeadItemName) Like "chdefine*" Then
        '表头自定义项参照
        Dim oDefPro As Object
        referpara.Cancel = True
        Set oDefPro = New U8DefPro.clsDefPro
        If oDefPro.Init(False, g_oLogin.UfDbName, g_oLogin.cUserId) Then
            '0:'手工输入;1:'系统档案;2:'单据
            Dim arr As Variant
            arr = Split(Voucher.ItemState(Index, 0).sDataRule, ",")
            '(1)当表体自定义项来源于基础档案时，arr(0) 是基础档案的表名；(2)当表体自定义项来源于单据时，arr(0) 是单据的类型（如：采购入库单(24)）
            '而接口：GetRefVal 在(1)时参数sCardNumber 是没有实际意义的；在(2)时参数sTableName 是没有实际意义的！
            If UBound(arr) > 0 Then
                sRet = oDefPro.GetRefVal(Voucher.ItemState(Index, 0).nDataSource, siHeader, Voucher.ItemState(Index, 0).sFieldName, arr(0), arr(1), arr(0), sRet, False, 5, 1)
            Else
                sRet = oDefPro.GetRefVal(Voucher.ItemState(Index, 0).nDataSource, siHeader, Voucher.ItemState(Index, 0).sFieldName, Voucher.ItemState(Index, 0).sTableName, Voucher.ItemState(Index, 0).sFieldName, gstrCardNumber, sRet, False, 5, 1)
            End If
        End If

    Else
        Call Refer_T(Voucher, Index, sRet, referpara)

    End If


    Exit Function

errhander:
    MsgBox Err.Description, vbInformation, GetString("U8.DZ.JA.Res030")

End Function

'表头 部门、客户、仓库、业务员、币种 参照
Private Function Refer_T(Voucher As Object, _
                         ByVal Index As Variant, _
                         sRet As Variant, _
                         referpara As UAPVoucherControl85.ReferParameter)

    On Error GoTo ErrHandler

    '获取当前编辑的字段名
    Dim sHeadItemName As String

    Dim btype         As Long

    Dim rst           As New ADODB.Recordset

    Dim sqlstr        As String

    sHeadItemName = LCase(Voucher.ItemState(Index, siHeader).sFieldName)

    referpara.Cancel = False

    '单据本身自带的参照服务
    '/*B*/ 根据单据表头模板设置确定是否需要以下栏目

    Select Case sHeadItemName

            '单位
        Case LCase("ecustcode"), LCase("cCusAbbName")
     
            referpara.ReferMetaXML = "<Ref><RefSet bAuth='" & IIf(bCus_ControlAuth, "1", "0") & "' authFunID='W' /></Ref>"
            referpara.id = "Customer_AA"
            referpara.RetField = "ccuscode"        '= " cCusCode   like '%" & sRet & "%' or cCusname like '%" & sRet & "%' or cCusAbbName like '%" & sRet & "%' or cCusMnemCode like '%" & sRet & "%' "
            referpara.sSql = " isnull(#FN[dEndDate],'9999-12-31')>N'" & g_oLogin.CurDate & "'"

            If sAuth_CusW <> "" Then
                referpara.sSql = referpara.sSql & "  and  #FN[cCusCode] in (" & sAuth_CusW & ")"
            End If
         '设计单位
            Case LCase("designunits"), LCase("decabbname")
            
           
            referpara.ReferMetaXML = "<Ref><RefSet bAuth='" & IIf(bCus_ControlAuth, "1", "0") & "' authFunID='W' /></Ref>"
            referpara.id = "Customer_AA"
            referpara.RetField = "ccuscode"        '= " cCusCode   like '%" & sRet & "%' or cCusname like '%" & sRet & "%' or cCusAbbName like '%" & sRet & "%' or cCusMnemCode like '%" & sRet & "%' "
            referpara.sSql = " isnull(#FN[dEndDate],'9999-12-31')>N'" & g_oLogin.CurDate & "'"

            If sAuth_CusW <> "" Then
                referpara.sSql = referpara.sSql & "  and  #FN[cCusCode] in (" & sAuth_CusW & ")"
            End If
   
    Case LCase("chdepartcode")
            referpara.id = "Department_AA"
            referpara.ReferMetaXML = "<Ref><RefSet bAuth='" & IIf(bDep_ControlAuth, "1", "0") & "' authFunID='W' /></Ref>"
            '过滤条件
            referpara.RetField = "cdepcode"
            referpara.sSql = " isnull(#FN[dDepEndDate],'9999-12-31')>N'" & g_oLogin.CurDate & "'"

            If sAuth_depW <> "" Then
                referpara.sSql = referpara.sSql & " and #FN[cDepCode] IN (" & sAuth_depW & ") "
            End If
            '部门
        Case LCase("edepmentcode"), LCase("cDepName")
            referpara.id = "Department_AA"
            referpara.ReferMetaXML = "<Ref><RefSet bAuth='" & IIf(bDep_ControlAuth, "1", "0") & "' authFunID='W' /></Ref>"
            '过滤条件
            referpara.RetField = "cdepcode"
            referpara.sSql = " isnull(#FN[dDepEndDate],'9999-12-31')>N'" & g_oLogin.CurDate & "'"

            If sAuth_depW <> "" Then
                referpara.sSql = referpara.sSql & " and #FN[cDepCode] IN (" & sAuth_depW & ") "
            End If
            
            '部门
        Case LCase("consubject"), LCase("consubname")
            referpara.id = "Department_AA"
            referpara.ReferMetaXML = "<Ref><RefSet bAuth='" & IIf(bDep_ControlAuth, "1", "0") & "' authFunID='W' /></Ref>"
            '过滤条件
            referpara.RetField = "cdepcode"
            referpara.sSql = " isnull(#FN[dDepEndDate],'9999-12-31')>N'" & g_oLogin.CurDate & "'"

            If sAuth_depW <> "" Then
                referpara.sSql = referpara.sSql & " and #FN[cDepCode] IN (" & sAuth_depW & ") "
            End If
            
        Case "checkproscode", "checkpername"
            referpara.id = "Person_AA"
            referpara.RetField = "cpersoncode"
            referpara.sSql = " dPValidDate <='" & g_oLogin.CurDate & "' and  isnull(dPInValidDate ,'2099-12-31') >='" & g_oLogin.CurDate & "' "
          
            '部门
        Case LCase("enfdepcode"), LCase("enfdepname")
            referpara.id = "Department_AA"
            referpara.ReferMetaXML = "<Ref><RefSet bAuth='" & IIf(bDep_ControlAuth, "1", "0") & "' authFunID='W' /></Ref>"
            '过滤条件
            referpara.RetField = "cdepcode"
            referpara.sSql = " isnull(#FN[dDepEndDate],'9999-12-31')>N'" & g_oLogin.CurDate & "'"

            If sAuth_depW <> "" Then
                referpara.sSql = referpara.sSql & " and #FN[cDepCode] IN (" & sAuth_depW & ") "
            End If
            
        Case "conproscode", "conpername"
            referpara.id = "Person_AA"
            referpara.RetField = "cpersoncode"
            referpara.sSql = " dPValidDate <='" & g_oLogin.CurDate & "' and  isnull(dPInValidDate ,'2099-12-31') >='" & g_oLogin.CurDate & "' "
             
            ' 统计分类conproscode
        Case LCase("buscode"), LCase("busname")   '"cfreightType", "MycdefineT6"
            referpara.Cancel = True
            Set Ref = New UFReferC.UFReferClient
            Ref.SetLogin g_oLogin

            Ref.SetRWAuth "", "", False
            
            sqlstr = "select isnull(btype,0) as btype  from HY_FYSL_Accounting where  cCode ='" & Voucher.headerText("acccode") & " '"
            Set rst = New ADODB.Recordset
            rst.Open sqlstr, g_Conn, 1, 1

            If Not rst.EOF Then
                btype = rst.Fields("btype")
                 
            Else
                btype = 0
            End If

            If btype = 1 Then
                sqlstr = "select  distinct ccode,cname from HY_FYSL_Business where   isnull(islevel,1)=1 and isnull(btype,1)=1 "
            Else
                sqlstr = "select  distinct ccode,cname from HY_FYSL_Business where   isnull(islevel,1)=1 and isnull(stype,1)=1 "
            End If
            
              If Voucher.headerText("acccode") = "" Then
                     sqlstr = "select  distinct ccode,cname from HY_FYSL_Business where   isnull(islevel,1)=1   "
                End If

            If Ref.StrRefInit_SetColWidth(g_oLogin, False, "", sqlstr, "业务类型编码,业务类型名称", "2000,4000") = True Then
                Ref.Show
            End If

            If Not (Ref.recmx Is Nothing) Then
                sRet = Ref.recmx("ccode")
            End If

            Ref.SetRWAuth "", "", True
        
            ' 统计分类conproscode
        Case LCase("statcode"), LCase("stcname")   '"cfreightType", "MycdefineT6"
            referpara.Cancel = True
            Set Ref = New UFReferC.UFReferClient
            Ref.SetLogin g_oLogin

            Ref.SetRWAuth "", "", False

            sqlstr = "select  distinct ccode,cname from HY_FYSL_Statclass where   isnull(islevel,1)=1 "

            If Ref.StrRefInit_SetColWidth(g_oLogin, False, "", sqlstr, "统计分类编码,统计分类名称", "2000,4000") = True Then
                Ref.Show
            End If

            If Not (Ref.recmx Is Nothing) Then
                sRet = Ref.recmx("ccode")
            End If

            Ref.SetRWAuth "", "", True

            '核算分类
        Case LCase("acccode"), LCase("accname")   '"cfreightType", "MycdefineT6"
            referpara.Cancel = True
            Set Ref = New UFReferC.UFReferClient
            Ref.SetLogin g_oLogin

            Ref.SetRWAuth "", "", False

            sqlstr = "select  distinct ccode,cname from HY_FYSL_Accounting where   isnull(islevel,1)=1 "

            If Ref.StrRefInit_SetColWidth(g_oLogin, False, "", sqlstr, "核算分类编码,核算分类名称", "2000,4000") = True Then
                Ref.Show
            End If

            If Not (Ref.recmx Is Nothing) Then
                sRet = Ref.recmx("ccode")
            End If

            Ref.SetRWAuth "", "", True

            '工程属性
        Case LCase("engproperties"), LCase("procname")   '"cfreightType", "MycdefineT6"
            referpara.Cancel = True
            Set Ref = New UFReferC.UFReferClient
            Ref.SetLogin g_oLogin

            Ref.SetRWAuth "", "", False

            sqlstr = "select  distinct ccode,cname from HY_FYSL_Properties where   isnull(islevel,1)=1"

            If Ref.StrRefInit_SetColWidth(g_oLogin, False, "", sqlstr, "工程属性编码,工程属性名称", "2000,4000") = True Then
                Ref.Show
            End If

            If Not (Ref.recmx Is Nothing) Then
                sRet = Ref.recmx("ccode")
            End If

            Ref.SetRWAuth "", "", True
            
            '上级发布单
        Case LCase("engcode"), LCase("engname")   '"cfreightType", "MycdefineT6"
            referpara.Cancel = True
            Set Ref = New UFReferC.UFReferClient
            Ref.SetLogin g_oLogin

            Ref.SetRWAuth "", "", False

            sqlstr = "select  distinct ccode,cname from V_HY_FYSL_Contract_refer2 where  ccode not in (select engcode from HY_FYSL_Contract) and id<>'" & lngVoucherID & "'"

            If Ref.StrRefInit_SetColWidth(g_oLogin, False, "", sqlstr, "工 程 编 号,工 程 名 称", "2000,4000") = True Then
                Ref.Show
            End If

            If Not (Ref.recmx Is Nothing) Then
                sRet = Ref.recmx("ccode")
            End If

            Ref.SetRWAuth "", "", True
      Case LCase("proccode"), LCase("proname")
            referpara.Cancel = True
            Set Ref = New UFReferC.UFReferClient
            Ref.SetLogin g_oLogin

            Ref.SetRWAuth "", "", False

            sqlstr = "select  distinct ccode,cname from V_HY_FYSL_Contract_refer3 where  ccode not in (select proccode from HY_FYSL_Contract) and id<>'" & lngVoucherID & "'"

            If Ref.StrRefInit_SetColWidth(g_oLogin, False, "", sqlstr, "工 程 编 号,工 程 名 称", "2000,4000") = True Then
                Ref.Show
            End If

            If Not (Ref.recmx Is Nothing) Then
                sRet = Ref.recmx("ccode")
            End If

            Ref.SetRWAuth "", "", True
            
     
            
     Case LCase("icode"), LCase("iname")   '"cfreightType", "MycdefineT6"
            referpara.Cancel = True
            Set Ref = New UFReferC.UFReferClient
            Ref.SetLogin g_oLogin

            Ref.SetRWAuth "", "", False
            
 
                sqlstr = "select  distinct ccode,cname from HY_FYSL_Investor where   isnull(islevel,1)=1 "

            If Ref.StrRefInit_SetColWidth(g_oLogin, False, "", sqlstr, "出资人分类编码,出资人分类名称", "2000,4000") = True Then
                Ref.Show
            End If

            If Not (Ref.recmx Is Nothing) Then
                sRet = Ref.recmx("ccode")
            End If

            Ref.SetRWAuth "", "", True
            
     Case LCase("procode")
            referpara.Cancel = True
            Set Ref = New UFReferC.UFReferClient
            Ref.SetLogin g_oLogin

            Ref.SetRWAuth "", "", False
             
'            If Voucher.headerText("proccode") <> "" Then
                 sqlstr = "select  distinct ccode,ccname,proccode,proname from V_HY_FYSL_Measurement where   isnull(cHandler,'')<>'' and  proccode='" & Voucher.headerText("proccode") & "'"
'             Else
'                sqlStr = "select  distinct ccode,ccname,proccode,proname from V_HY_FYSL_Measurement where   isnull(cHandler,'')<>''"
'            End If
            If Ref.StrRefInit_SetColWidth(g_oLogin, False, "", sqlstr, "计量单号,计量名称,项目单号,项目名称", "2000,2000,2000,2000") = True Then
                Ref.Show
            End If

            If Not (Ref.recmx Is Nothing) Then
                sRet = Ref.recmx("ccode")
            End If

            Ref.SetRWAuth "", "", True
            
    End Select

    Exit Function

ErrHandler:
    MsgBox Err.Description, vbInformation, GetString("U8.DZ.JA.Res030")

End Function

Public Function VoucherheadCellCheckFun(Voucher As Object, Index As Variant, retvalue As String, bChanged As UAPVoucherControl85.CheckRet, referpara As UAPVoucherControl85.ReferParameter)
    On Error GoTo Err_Handler

    Dim sMetaItemName As String
    sMetaItemName = Voucher.ItemState(Index, siHeader).sFieldName


    '表头自定义项校验

    If LCase(sMetaItemName) Like "cdefine*" Or LCase(sMetaItemName) Like "chdefine*" Then
        Call DefineCheck_T(Voucher, Index, retvalue, bChanged, referpara)

    Else

        '处理编码,名称参照赋值,非手工输入
        If Not referpara.rstGrid Is Nothing Then

            Call ReferCheck_T(Voucher, Index, retvalue, bChanged, referpara)

            referpara.rstGrid.Close
            Set referpara.rstGrid = Nothing

        Else
            '手工输入编码、名称校验
            Call HandRecord_T(Voucher, Index, retvalue, bChanged, referpara)

        End If


    End If


    Exit Function

Err_Handler:
    MsgBox Err.Description, vbInformation, GetString("U8.DZ.JA.Res030")

End Function


'表头自定义项校验

Private Function DefineCheck_T(Voucher As Object, Index As Variant, retvalue As String, bChanged As UAPVoucherControl85.CheckRet, referpara As UAPVoucherControl85.ReferParameter)

    On Error GoTo ErrHandler

    Dim sMetaItemName As String
    sMetaItemName = Voucher.ItemState(Index, siHeader).sFieldName

    Dim bFixLen As Boolean
    Dim lFixLen As Long
    Dim arr As Variant
    Dim iRet As Integer
    Dim cDefValue As String
    Dim oDefPro As Object
    Dim RecCurRow As DOMDocument
    Dim STMsgTitle As String
    Set RecCurRow = New DOMDocument

    STMsgTitle = GetString("U8.DZ.JA.Res030")
    cDefValue = retvalue
    arr = Split(Voucher.ItemState(Index, 0).sDefaultValue, ",")
    If UBound(arr) > 0 Then
        bFixLen = CBool(arr(0))
        lFixLen = val(arr(1))
        If bFixLen And Len(retvalue) > lFixLen Then
            'Result:Row=4792        Col=76  Content="]超出了定长！" ID=88797142-25e8-4611-be75-a8586a97d0c2
            MsgBox GetResString("U8.ST.USKCGLSQL.frmqc.01806", Array("[" & Voucher.ItemState(Index, 0).sCardFormula1)), vbOKOnly + vbInformation, STMsgTitle
            bChanged = Cancel
            Exit Function
        End If
    End If

    If Voucher.ItemState(Index, 0).bValidityCheck Then
        '0:'手工输入;1:'系统档案;2:'单据
        arr = Split(Voucher.ItemState(Index, 0).sDataRule, ",")
        Set oDefPro = New U8DefPro.clsDefPro
        If oDefPro.Init(False, g_oLogin.UfDbName, g_oLogin.cUserId) Then
            cDefValue = retvalue
            If UBound(arr) > 0 Then
                iRet = oDefPro.ValidateAr(Voucher.ItemState(Index, 0).nDataSource, 0, Voucher.ItemState(Index, 0).sFieldName, arr(0), arr(1), cDefValue, "FC03FZ121", "", Voucher.ItemState(Index, 0).bBuildArchives)
            Else
                iRet = oDefPro.ValidateAr(Voucher.ItemState(Index, 0).nDataSource, 0, Voucher.ItemState(Index, 0).sFieldName, Voucher.ItemState(Index, 0).sTableName, Voucher.ItemState(Index, 0).sFieldName, cDefValue, "FC03FZ121", "", Voucher.ItemState(Index, 0).bBuildArchives)
            End If
            If iRet < 0 Then
                If Voucher.ItemState(Index, 0).bValidityCheck Then
                    'Result:Row=4814        Col=78  Content="不合法,请重新录入！"   ID=6d0a4805-7f50-4a25-a795-b499fe42d6b1
                    MsgBox GetResString("U8.ST.USKCGLSQL.clsbasebillhandler.01453", Array(Voucher.ItemState(Index, 0).sCardFormula1)), vbOKOnly + vbExclamation, STMsgTitle
                    bChanged = Cancel
                    Exit Function
                End If
            Else
                retvalue = cDefValue
                Voucher.headerText(Index) = cDefValue
            End If
        End If
    End If


    Exit Function

ErrHandler:
    MsgBox Err.Description, vbInformation, GetString("U8.DZ.JA.Res030")
End Function

'表头参照校验
'部门、仓库、业务员、客户、币种、汇率
Private Function ReferCheck_T(Voucher As Object, Index As Variant, retvalue As String, bChanged As UAPVoucherControl85.CheckRet, referpara As UAPVoucherControl85.ReferParameter)
    On Error GoTo ErrHandler

    Dim rs As ADODB.Recordset
    Dim sql As String

    Dim sMetaItemName As String
    sMetaItemName = LCase(Voucher.ItemState(Index, siHeader).sFieldName)

    If Not referpara.rstGrid.EOF Then

        '/*B*/ 根据单据模板表头设置确定是否需要以下栏目,以及各栏目的名称、大小写 bObjectCode bObjectName
        Select Case sMetaItemName
           
                '单位
            Case LCase("ecustcode"), LCase("cCusAbbName")
                
                Voucher.headerText("ecustcode") = referpara.rstGrid.Fields("ccuscode")
                Voucher.headerText("cCusAbbName") = referpara.rstGrid.Fields("ccusabbname")
                 Voucher.headerText("cCusName") = referpara.rstGrid.Fields("ccusname")
                If sMetaItemName = LCase("ecustcode") Then
                    retvalue = referpara.rstGrid.Fields("ccuscode")
                Else
                    retvalue = referpara.rstGrid.Fields("ccusabbname")
                End If
             '设计单位
            Case LCase("designunits"), LCase("decabbname")
                
                Voucher.headerText("designunits") = referpara.rstGrid.Fields("ccuscode")
                Voucher.headerText("decabbname") = referpara.rstGrid.Fields("ccusabbname")
                Voucher.headerText("descuname") = referpara.rstGrid.Fields("ccusname")
                If sMetaItemName = LCase("designunits") Then
                    retvalue = referpara.rstGrid.Fields("ccuscode")
                Else
                    retvalue = referpara.rstGrid.Fields("ccusabbname")
                End If
                     
                  Case LCase("chdepartcode")
                   Voucher.headerText("chdepname") = referpara.rstGrid.Fields("cdepname")
                retvalue = referpara.rstGrid.Fields("cdepcode")
'
                '部门
            Case "edepmentcode"
                Voucher.headerText("cDepName") = referpara.rstGrid.Fields("cdepname")
                retvalue = referpara.rstGrid.Fields("cdepcode")
            Case "cDepName"
                Voucher.headerText("edepmentcode") = referpara.rstGrid.Fields("cdepcode")
                retvalue = referpara.rstGrid.Fields("cdepname")
                
             Case "consubject"
                Voucher.headerText("consubname") = referpara.rstGrid.Fields("cdepname")
                retvalue = referpara.rstGrid.Fields("cdepcode")
            Case "consubname"
                Voucher.headerText("consubject") = referpara.rstGrid.Fields("cdepcode")
                retvalue = referpara.rstGrid.Fields("cdepname")
                
           Case "conproscode"
                Voucher.headerText("conpername") = referpara.rstGrid.Fields("cpersonname")
 
                retvalue = referpara.rstGrid.Fields("cpersoncode")
                
                
             Case "verdeptcode"
                Voucher.headerText("verdeptname") = referpara.rstGrid.Fields("cdepname")
                retvalue = referpara.rstGrid.Fields("cdepcode")
            Case "verdeptname"
                Voucher.headerText("verdeptcode") = referpara.rstGrid.Fields("cdepcode")
                retvalue = referpara.rstGrid.Fields("cdepname")
                
             Case "enfdepcode"
                Voucher.headerText("enfdepname") = referpara.rstGrid.Fields("cdepname")
                retvalue = referpara.rstGrid.Fields("cdepcode")
            Case "enfdepname"
                Voucher.headerText("enfdepcode") = referpara.rstGrid.Fields("cdepcode")
                retvalue = referpara.rstGrid.Fields("cdepname")
                
           Case "checkproscode"
                Voucher.headerText("checkpername") = referpara.rstGrid.Fields("cpersonname")
 
                retvalue = referpara.rstGrid.Fields("cpersoncode")
                
                
                
           Case LCase("buscode"), LCase("busname")
            
             Voucher.headerText("buscode") = referpara.rstGrid.Fields("ccode")
                Voucher.headerText("busname") = referpara.rstGrid.Fields("cname")
                If sMetaItemName = LCase("buscode") Then
                    retvalue = referpara.rstGrid.Fields("ccode")
                Else
                    retvalue = referpara.rstGrid.Fields("cname")
                End If
                
            Case LCase("statcode"), LCase("stcname")
            
             Voucher.headerText("statcode") = referpara.rstGrid.Fields("ccode")
                Voucher.headerText("stcname") = referpara.rstGrid.Fields("cname")
                If sMetaItemName = LCase("statcode") Then
                    retvalue = referpara.rstGrid.Fields("ccode")
                Else
                    retvalue = referpara.rstGrid.Fields("cname")
                End If
            
              
              
            Case LCase("acccode"), LCase("accname")
               Voucher.headerText("acccode") = referpara.rstGrid.Fields("ccode")
                Voucher.headerText("accname") = referpara.rstGrid.Fields("cname")
                If sMetaItemName = LCase("acccode") Then
                    retvalue = referpara.rstGrid.Fields("ccode")
                Else
                    retvalue = referpara.rstGrid.Fields("cname")
                End If
           Case LCase("engproperties"), LCase("procname")
                Voucher.headerText("engproperties") = referpara.rstGrid.Fields("ccode")
                Voucher.headerText("procname") = referpara.rstGrid.Fields("cname")
                If sMetaItemName = LCase("engproperties") Then
                    retvalue = referpara.rstGrid.Fields("ccode")
                Else
                    retvalue = referpara.rstGrid.Fields("cname")
                End If
                
            Case LCase("engcode"), LCase("engname")
                Voucher.headerText("engcode") = referpara.rstGrid.Fields("ccode")
                Voucher.headerText("engname") = referpara.rstGrid.Fields("cname")
                If sMetaItemName = LCase("engcode") Then
                    retvalue = referpara.rstGrid.Fields("ccode")
                Else
                    retvalue = referpara.rstGrid.Fields("cname")
                End If
                
           Case LCase("proccode"), LCase("proname")
                Voucher.headerText("proccode") = referpara.rstGrid.Fields("ccode")
                Voucher.headerText("proname") = referpara.rstGrid.Fields("cname")
                If sMetaItemName = LCase("proccode") Then
                    retvalue = referpara.rstGrid.Fields("ccode")
                Else
                    retvalue = referpara.rstGrid.Fields("cname")
                End If
                    
                
            Case LCase("icode"), LCase("iname")
                Voucher.headerText("icode") = referpara.rstGrid.Fields("ccode")
                Voucher.headerText("iname") = referpara.rstGrid.Fields("cname")
                If sMetaItemName = LCase("icode") Then
                    retvalue = referpara.rstGrid.Fields("ccode")
                Else
                    retvalue = referpara.rstGrid.Fields("cname")
                    
                End If
                    
             Case LCase("procode")
              Voucher.headerText("procode") = referpara.rstGrid.Fields("ccode")
               retvalue = referpara.rstGrid.Fields("ccode")
                 
                
              Case LCase("isengdec")
                If retvalue = "否" Then
                    Voucher.headerText("supengcode") = ""
                    Voucher.headerText("supcname") = ""

                    Voucher.EnableHead "supengcode", False
                    Voucher.EnableHead "supcname", False
                    '                  Voucher.SetCurrentRow ("@AutoID=" & bodyele.getAttribute("AutoID") & "")
                ElseIf retvalue = "是" Then
                    Voucher.EnableHead "supengcode", True
                    bChanged = Cancel
                    Exit Function
                End If
           
        End Select

        bChanged = success

    End If


    Exit Function

ErrHandler:
    MsgBox Err.Description, vbInformation, GetString("U8.DZ.JA.Res030")

End Function

'                Select Case Voucher.headerText("cType")
'                    Case "客户"
'                        referpara.id = "Customer_AA"
'                        referpara.sSql = sHeadItemName & "  like '%" & sRet & "%'"
'                    Case "供应商"
'                        referpara.id = "Vendor_AA"
'                        referpara.sSql = sHeadItemName & "  like '%" & sRet & "%'"
'                    Case "部门"
'                         referpara.id = "Department_AA"
'                         '过滤条件
'                         referpara.sSql = sHeadItemName & "  like '%" & sRet & "%'"
'                    Case "人员"
'                         referpara.id = "Person_AA"
'                         referpara.sSql = sHeadItemName & "  like '%" & sRet & "%' and personentity_person.cdepcode like '%" & Voucher.headerText("cdepcode") & "%'"
'                End Select
Private Function HandRecord_TbObjectCode(Voucher As Object, Index As Variant, retvalue As String, _
                                         bChanged As UAPVoucherControl85.CheckRet, _
                                         referpara As UAPVoucherControl85.ReferParameter)
    On Error GoTo ErrHandler

    Dim strValue As Boolean
    Dim sql As String
    Dim sMetaItemName As String
    sMetaItemName = LCase(Voucher.ItemState(Index, siHeader).sFieldName)

    '手工输入时,需要校验是否存在
    '/*B*/ 根据单据模板表头设置确定是否需要以下栏目
    Select Case Voucher.headerText("cType")
            '部门
            'enum by modify
        Case "部门"
            If retvalue = "" Then
                Voucher.headerText("bObjectCode") = ""
                Voucher.headerText("bObjectName") = ""
            Else
                sql = "select cdepcode code,cdepname name from department where (cdepcode='" & retvalue & "' or cdepname='" & retvalue & "') and isnull(dDepEndDate,'9999-12-31')>N'" & g_oLogin.CurDate & "' "
                sql = sql & IIf(sAuth_depW = "", "", " and cdepcode in (" & sAuth_depW & ")")

                strValue = CheckCellValue(sql)

                If strValue = False Then
                    ReDim varArgs(0)
                    varArgs(0) = retvalue
                    MsgBox GetStringPara("U8.DZ.JA.Res560", varArgs(0)), vbInformation, GetString("U8.DZ.JA.Res030")

                    '                MsgBox "部门" & RetValue & "不存在或者没有权限,请重新输入!", vbInformation, GetString("U8.DZ.JA.Res030")
                    bChanged = Cancel
                    Exit Function
                Else
                    Voucher.headerText("bObjectCode") = strCellCode
                    Voucher.headerText("bObjectName") = strCellName
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
                Voucher.headerText("bObjectCode") = ""
                Voucher.headerText("bObjectName") = ""
            Else
                sql = "select cPersonCode code,cPersonName name from person where (cPersonCode='" & retvalue & "' or cPersonName='" & retvalue & "') and  '" & g_oLogin.CurDate & "' between dPValidDate  and IsNull(dPInValidDate,'2099-12-31') "

                strValue = CheckCellValue(sql)

                If strValue = False Then
                    ReDim varArgs(0)
                    varArgs(0) = retvalue
                    MsgBox GetStringPara("U8.DZ.JA.Res570", varArgs(0)), vbInformation, GetString("U8.DZ.JA.Res030")

                    '                MsgBox "业务员" & RetValue & "不存在或者没有权限,请重新输入!", vbInformation, GetString("U8.DZ.JA.Res030")
                    bChanged = Cancel
                    Exit Function
                Else
                    Voucher.headerText("bObjectCode") = strCellCode
                    Voucher.headerText("bObjectName") = strCellName
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
                Voucher.headerText("bObjectCode") = ""
                Voucher.headerText("bObjectName") = ""
            Else
                sql = "select cCusCode ,cCusAbbName,cCusName,cCusAddress  from Customer where (cCusCode='" & retvalue & "' or cCusAbbName='" & retvalue & "' or cCusMnemCode ='" & retvalue & "' or cCusName ='" & retvalue & "') and  isnull(dEndDate,'9999-12-31')>N'" & g_oLogin.CurDate & "' "
                sql = sql & IIf(sAuth_CusW = "", "", " and cCusCode in (" & sAuth_CusW & ")")

                strValue = CheckCustomer(sql, cName, Address)

                If strValue = False Then
                    ReDim varArgs(0)
                    varArgs(0) = retvalue
                    MsgBox GetStringPara("U8.DZ.JA.Res580", varArgs(0)), vbInformation, GetString("U8.DZ.JA.Res030")

                    '                MsgBox "客户" & RetValue & "不存在或者没有权限,请重新输入!", vbInformation, GetString("U8.DZ.JA.Res030")
                    bChanged = Cancel
                    Exit Function
                Else
                    Voucher.headerText("bObjectCode") = strCellCode
                    Voucher.headerText("bObjectName") = strCellName

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
                Voucher.headerText("bObjectCode") = ""
                Voucher.headerText("bObjectName") = ""
            Else
                sql = "select cvencode ,cvenname,cvenabbname  from vendor where (cvencode='" & retvalue & "' or cvenname='" & retvalue & "' or cVenMnemCode='" & retvalue & "' or cVenAbbName ='" & retvalue & "') and  isnull(dEndDate,'9999-12-31')>N'" & g_oLogin.CurDate & "' "
                sql = sql & IIf(sAuth_vendorW = "", "", " and iid in (" & sAuth_vendorW & ")")

                strValue = CheckVendor(sql, cvencode, cvenname)

                If strValue = False Then
                    ReDim varArgs(0)
                    varArgs(0) = retvalue
                    MsgBox GetStringPara("U8.DZ.JA.Res590", varArgs(0)), vbInformation, GetString("U8.DZ.JA.Res030")

                    '                MsgBox "供应商" & RetValue & "不存在或者没有权限,请重新输入!", vbInformation, GetString("U8.DZ.JA.Res030")
                    bChanged = Cancel
                    Exit Function
                Else
                    Voucher.headerText("bObjectCode") = cvencode
                    Voucher.headerText("bObjectName") = cvenname

                    If sMetaItemName Like "*code" Then
                        retvalue = strCellCode
                    Else
                        retvalue = strCellName
                    End If

                End If
            End If
    End Select

    bChanged = success
    Exit Function
ErrHandler:
    MsgBox Err.Description, vbInformation, GetString("U8.DZ.JA.Res030")
End Function

'手工录入表头字段校验
'部门、仓库、业务员、客户
Private Function HandRecord_T(Voucher As Object, _
                              Index As Variant, _
                              retvalue As String, _
                              bChanged As UAPVoucherControl85.CheckRet, _
                              referpara As UAPVoucherControl85.ReferParameter)

    On Error GoTo ErrHandler

    Dim strValue      As Boolean

    Dim rs            As New ADODB.Recordset

    Dim sql           As String

    Dim sMetaItemName As String

    Dim cvencode      As String

    Dim cvenname      As String

    Dim rst           As New ADODB.Recordset

    Dim stype         As Long

    Dim btype         As Long

    Dim cName         As String

    Dim Address       As String

    sMetaItemName = LCase(Voucher.ItemState(Index, siHeader).sFieldName)

    '手工输入时,需要校验是否存在

    '/*B*/ 根据单据模板表头设置确定是否需要以下栏目
    Select Case sMetaItemName
            '单位
        
        Case "ecustcode", "cCusAbbName"

            If retvalue = "" Then
                Voucher.headerText("ecustcode") = ""
                Voucher.headerText("cCusAbbName") = ""
                Voucher.headerText("cCusName") = ""
 
            Else

                sql = "select cCusCode ,cCusAbbName,cCusName,cCusAddress  from Customer where (cCusCode='" & retvalue & "' or cCusAbbName='" & retvalue & "' or cCusMnemCode='" & retvalue & "' or cCusAbbName ='" & retvalue & "' )"
                ' sql = sql & IIf(sAuth_CusW = "", "", " and cCusCode in (" & sAuth_CusW & ")")

                '权限
                '                    Authstr = GetRowAuthAlls(g_Conn, g_oLogin.cUserId, "R", "bAuth_Cus")
                '                    If Authstr <> "" Then sql = sql & Authstr

                strValue = CheckCustomer(sql, cName, Address)

                If strValue = False Then
                    ReDim varArgs(0)
                    varArgs(0) = retvalue
                    MsgBox GetStringPara("U8.DZ.JA.Res580", varArgs(0)), vbInformation, GetString("U8.DZ.JA.Res030")

                    'MsgBox "客户" & RetValue & "不存在或者没有权限,请重新输入!", vbInformation, GetString("U8.DZ.JA.Res030")
                    bChanged = Cancel

                    Exit Function

                Else
                    Voucher.headerText("ecustcode") = strCellCode
                     
                    Voucher.headerText("cCusAbbName") = strCellName
                    Voucher.headerText("cCusName") = cName
                    
                    If sMetaItemName Like "*code" Then
                        retvalue = strCellCode
                    Else
                        retvalue = strCellName
                    End If

                End If
            End If
            
        Case LCase("designunits"), LCase("decabbname")

            If retvalue = "" Then
                Voucher.headerText("designunits") = ""
                Voucher.headerText("decabbname") = ""
                Voucher.headerText("descuname") = ""
 
            Else

                sql = "select cCusCode ,cCusAbbName,cCusName,cCusAddress  from Customer where (cCusCode='" & retvalue & "' or cCusAbbName='" & retvalue & "' or cCusMnemCode='" & retvalue & "' or cCusAbbName ='" & retvalue & "' )"
                ' sql = sql & IIf(sAuth_CusW = "", "", " and cCusCode in (" & sAuth_CusW & ")")

                '权限
                '                    Authstr = GetRowAuthAlls(g_Conn, g_oLogin.cUserId, "R", "bAuth_Cus")
                '                    If Authstr <> "" Then sql = sql & Authstr

                strValue = CheckCustomer(sql, cName, Address)

                If strValue = False Then
                    ReDim varArgs(0)
                    varArgs(0) = retvalue
                    MsgBox GetStringPara("U8.DZ.JA.Res580", varArgs(0)), vbInformation, GetString("U8.DZ.JA.Res030")

                    'MsgBox "客户" & RetValue & "不存在或者没有权限,请重新输入!", vbInformation, GetString("U8.DZ.JA.Res030")
                    bChanged = Cancel

                    Exit Function

                Else
                    Voucher.headerText("designunits") = strCellCode
                     
                    Voucher.headerText("decabbname") = strCellName
                    Voucher.headerText("descuname") = cName
                    
                    If sMetaItemName Like "*code" Then
                        retvalue = strCellCode
                    Else
                        retvalue = strCellName
                    End If

                End If
            End If
                
        Case LCase("chdepartcode")
                 
            If retvalue = "" Then
                Voucher.headerText("chdepartcode") = ""
                Voucher.headerText("chdepname") = ""
            
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
                    MsgBox GetStringPara("U8.DZ.JA.Res560", varArgs(0)), vbInformation, GetString("U8.DZ.JA.Res030")

                    'MsgBox "部门" & RetValue & "不存在或者没有权限,请重新输入!", vbInformation, GetString("U8.DZ.JA.Res030")
                    bChanged = Cancel

                    Exit Function

                Else
                    Voucher.headerText("chdepartcode") = strCellCode
                    Voucher.headerText("chdepname") = strCellName

                    If sMetaItemName Like "*code" Then
                        retvalue = strCellCode
                    Else
                        retvalue = strCellName
                    End If
                End If

            End If
                 
        Case "consubject", "consubname"

            If retvalue = "" Then
                Voucher.headerText("consubject") = ""
                Voucher.headerText("consubname") = ""
            
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
                    MsgBox GetStringPara("U8.DZ.JA.Res560", varArgs(0)), vbInformation, GetString("U8.DZ.JA.Res030")

                    'MsgBox "部门" & RetValue & "不存在或者没有权限,请重新输入!", vbInformation, GetString("U8.DZ.JA.Res030")
                    bChanged = Cancel

                    Exit Function

                Else
                    Voucher.headerText("consubject") = strCellCode
                    Voucher.headerText("consubname") = strCellName

                    If sMetaItemName Like "*code" Then
                        retvalue = strCellCode
                    Else
                        retvalue = strCellName
                    End If
                End If

            End If
            
        Case "enfdepcode", "enfdepname"

            If retvalue = "" Then
                Voucher.headerText("enfdepcode") = ""
                Voucher.headerText("enfdepname") = ""
                Voucher.headerText("conproscode") = ""
                Voucher.headerText("conpername") = ""
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
                    MsgBox GetStringPara("U8.DZ.JA.Res560", varArgs(0)), vbInformation, GetString("U8.DZ.JA.Res030")

                    'MsgBox "部门" & RetValue & "不存在或者没有权限,请重新输入!", vbInformation, GetString("U8.DZ.JA.Res030")
                    bChanged = Cancel

                    Exit Function

                Else
                    Voucher.headerText("enfdepcode") = strCellCode
                    Voucher.headerText("enfdepname") = strCellName

                    If sMetaItemName Like "*code" Then
                        retvalue = strCellCode
                    Else
                        retvalue = strCellName
                    End If
                End If

            End If
                
        Case "conproscode", "conpername"

            If retvalue = "" Then
                Voucher.headerText("conproscode") = ""
                Voucher.headerText("conpername") = ""
            Else

                sql = "select cPersonCode code,cPersonName name from person where (cPersonCode='" & retvalue & "' or cPersonName='" & retvalue & "') and  '" & g_oLogin.CurDate & "' between dPValidDate  and IsNull(dPInValidDate,'2099-12-31')  "
                sql = sql & "  and dPValidDate <='" & g_oLogin.CurDate & "' and  isnull(dPInValidDate ,'2099-12-31') >='" & g_oLogin.CurDate & "' "

                '权限
                '                    Authstr = GetRowAuthAlls(g_Conn, g_oLogin.cUserId, "R", "bAuth_Per")
                '                    If Authstr <> "" Then sql = sql & Authstr

                strValue = CheckCellValue(sql)

                If strValue = False Then
                    ReDim varArgs(0)
                    varArgs(0) = retvalue
                    MsgBox GetStringPara("U8.DZ.JA.Res570", varArgs(0)), vbInformation, GetString("U8.DZ.JA.Res030")

                    'MsgBox "业务员" & RetValue & "不存在或者没有权限,请重新输入!", vbInformation, GetString("U8.DZ.JA.Res030")
                    bChanged = Cancel

                    Exit Function

                Else
                    Voucher.headerText("conproscode") = strCellCode
                    Voucher.headerText("conpername") = strCellName

                    If sMetaItemName Like "*code" Then
                        retvalue = strCellCode
                    Else
                        retvalue = strCellName
                    End If
                End If
            End If
                
        Case "verdeptcode", "verdeptname"

            If retvalue = "" Then
                Voucher.headerText("verdeptcode") = ""
                Voucher.headerText("verdeptname") = ""
                 
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
                    MsgBox GetStringPara("U8.DZ.JA.Res560", varArgs(0)), vbInformation, GetString("U8.DZ.JA.Res030")

                    'MsgBox "部门" & RetValue & "不存在或者没有权限,请重新输入!", vbInformation, GetString("U8.DZ.JA.Res030")
                    bChanged = Cancel

                    Exit Function

                Else
                    Voucher.headerText("verdeptcode") = strCellCode
                    Voucher.headerText("verdeptname") = strCellName

                    If sMetaItemName Like "*code" Then
                        retvalue = strCellCode
                    Else
                        retvalue = strCellName
                    End If
                End If

            End If
                
        Case "checkproscode", "checkpername"

            If retvalue = "" Then
                Voucher.headerText("checkproscode") = ""
                Voucher.headerText("checkpername") = ""
            Else

                sql = "select cPersonCode code,cPersonName name from person where (cPersonCode='" & retvalue & "' or cPersonName='" & retvalue & "') and  '" & g_oLogin.CurDate & "' between dPValidDate  and IsNull(dPInValidDate,'2099-12-31')  "
                sql = sql & "  and dPValidDate <='" & g_oLogin.CurDate & "' and  isnull(dPInValidDate ,'2099-12-31') >='" & g_oLogin.CurDate & "' "

                '权限
                '                    Authstr = GetRowAuthAlls(g_Conn, g_oLogin.cUserId, "R", "bAuth_Per")
                '                    If Authstr <> "" Then sql = sql & Authstr

                strValue = CheckCellValue(sql)

                If strValue = False Then
                    ReDim varArgs(0)
                    varArgs(0) = retvalue
                    MsgBox GetStringPara("U8.DZ.JA.Res570", varArgs(0)), vbInformation, GetString("U8.DZ.JA.Res030")

                    'MsgBox "业务员" & RetValue & "不存在或者没有权限,请重新输入!", vbInformation, GetString("U8.DZ.JA.Res030")
                    bChanged = Cancel

                    Exit Function

                Else
                    Voucher.headerText("checkproscode") = strCellCode
                    Voucher.headerText("checkpername") = strCellName

                    If sMetaItemName Like "*code" Then
                        retvalue = strCellCode
                    Else
                        retvalue = strCellName
                    End If
                End If
            End If
    
        Case LCase("buscode"), LCase("busname")

            If retvalue = "" Then
                Voucher.headerText("buscode") = ""
                Voucher.headerText("busname") = ""

                Exit Function

            End If
              
            sql = "select isnull(btype,0) as btype  from HY_FYSL_Accounting where  cCode ='" & Voucher.headerText("acccode") & " '"
            Set rst = New ADODB.Recordset
            rst.Open sql, g_Conn, 1, 1

            If Not rst.EOF Then
                btype = rst.Fields("btype")
            Else
                btype = 0
            End If
              
            If btype = 1 Then
                sql = "SELECT cCode,cName from HY_FYSL_Business where (cName='" & retvalue & "' or cCode='" & retvalue & "') and isnull(islevel,1)=1  and isnull(btype,1)=1  "
            Else
                sql = "SELECT cCode,cName from HY_FYSL_Business where (cName='" & retvalue & "' or cCode='" & retvalue & "') and isnull(islevel,1)=1  and isnull(stype,1)=1  "
            End If
                
            If Voucher.headerText("acccode") = "" Then
                sql = "SELECT cCode,cName from HY_FYSL_Business where (cName='" & retvalue & "' or cCode='" & retvalue & "') and isnull(islevel,1)=1     "
            End If
                
            Set rs = New ADODB.Recordset
            rs.Open sql, g_Conn, 1, 1

            If rs.EOF Then
                MsgBox "不存在此分类信息", vbInformation, "提示"
                bChanged = Cancel

                Exit Function

            Else
                Voucher.headerText("buscode") = rs.Fields("cCode") & ""
                Voucher.headerText("busname") = rs.Fields("cName") & ""

                If sMetaItemName Like "*code" Then
                    retvalue = rs.Fields("cCode") & ""
                Else
                    retvalue = rs.Fields("cName") & ""
                End If
            End If
                
        Case LCase("statcode"), LCase("stcname")

            If retvalue = "" Then
                Voucher.headerText("statcode") = ""
                Voucher.headerText("stcname") = ""

                Exit Function

            End If
              
            sql = "SELECT cCode,cName from HY_FYSL_Statclass where (cName='" & retvalue & "' or cCode='" & retvalue & "') and isnull(islevel,1)=1"
            Set rs = New ADODB.Recordset
            rs.Open sql, g_Conn, 1, 1

            If rs.EOF Then
                MsgBox "不存在此分类信息", vbInformation, "提示"
                bChanged = Cancel

                Exit Function

            Else
                Voucher.headerText("statcode") = rs.Fields("cCode") & ""
                Voucher.headerText("stcname") = rs.Fields("cName") & ""

                If sMetaItemName Like "*code" Then
                    retvalue = rs.Fields("cCode") & ""
                Else
                    retvalue = rs.Fields("cName") & ""
                End If
            End If

        Case LCase("icode"), LCase("iname")

            If retvalue = "" Then
                Voucher.headerText("icode") = ""
                Voucher.headerText("iname") = ""

                Exit Function

            End If
              
            sql = "SELECT cCode,cName from HY_FYSL_Investor where (cName='" & retvalue & "' or cCode='" & retvalue & "') and isnull(islevel,1)=1"
            Set rs = New ADODB.Recordset
            rs.Open sql, g_Conn, 1, 1

            If rs.EOF Then
                MsgBox "不存在此分类信息", vbInformation, "提示"
                bChanged = Cancel

                Exit Function

            Else
                Voucher.headerText("icode") = rs.Fields("cCode") & ""
                Voucher.headerText("iname") = rs.Fields("cName") & ""

                If sMetaItemName Like "*code" Then
                    retvalue = rs.Fields("cCode") & ""
                Else
                    retvalue = rs.Fields("cName") & ""
                End If
            End If
              
        Case LCase("acccode"), LCase("accname")

            If retvalue = "" Then
                Voucher.headerText("acccode") = ""
                Voucher.headerText("accname") = ""

                Exit Function

            End If
             
            sql = "SELECT cCode,cName from HY_FYSL_Accounting where  (cName='" & retvalue & "' or cCode='" & retvalue & "') and isnull(islevel,1)=1"
            Set rs = New ADODB.Recordset
            rs.Open sql, g_Conn, 1, 1

            If rs.EOF Then
                MsgBox "不存在此分类信息", vbInformation, "提示"
                bChanged = Cancel

                Exit Function

            Else
                Voucher.headerText("acccode") = rs.Fields("cCode") & ""
                Voucher.headerText("accname") = rs.Fields("cName") & ""

                If sMetaItemName Like "*code" Then
                    retvalue = rs.Fields("cCode") & ""
                Else
                    retvalue = rs.Fields("cName") & ""
                End If
            End If
             
        Case LCase("engproperties"), LCase("procname")

            If retvalue = "" Then
                Voucher.headerText("engproperties") = ""
                Voucher.headerText("procname") = ""

                Exit Function

            End If
             
            sql = "SELECT cCode,cName from HY_FYSL_Properties  where  (cName='" & retvalue & "' or cCode='" & retvalue & "') and isnull(islevel,1)=1"
            Set rs = New ADODB.Recordset
            rs.Open sql, g_Conn, 1, 1

            If rs.EOF Then
                MsgBox "不存在此分类信息", vbInformation, "提示"
                bChanged = Cancel

                Exit Function

            Else
                Voucher.headerText("engproperties") = rs.Fields("cCode") & ""
                Voucher.headerText("procname") = rs.Fields("cName") & ""

                If sMetaItemName Like "engproperties" Then
                    retvalue = rs.Fields("cCode") & ""
                Else
                    retvalue = rs.Fields("cName") & ""
                End If
            End If
                
        Case LCase("engcode"), LCase("engname")

            If retvalue = "" Then
                Voucher.headerText("engcode") = ""
                Voucher.headerText("engname") = ""

                Exit Function

            End If
             
            sql = "SELECT cCode,cName from V_HY_FYSL_Contract_refer2  where  (cName='" & retvalue & "' or cCode='" & retvalue & "')  and  id<>'" & lngVoucherID & "'"
            Set rs = New ADODB.Recordset
            rs.Open sql, g_Conn, 1, 1

            If rs.EOF Then
                MsgBox "工程编号不存在", vbInformation, "提示"
                bChanged = Cancel

                Exit Function

            Else
                Voucher.headerText("engcode") = rs.Fields("cCode") & ""
                Voucher.headerText("engname") = rs.Fields("cName") & ""

                If sMetaItemName Like "*code" Then
                    retvalue = rs.Fields("cCode") & ""
                Else
                    retvalue = rs.Fields("cName") & ""
                End If
            End If
                
        Case LCase("proccode"), LCase("procname")

            If retvalue = "" Then
                Voucher.headerText("proccode") = ""
                Voucher.headerText("procname") = ""

                Exit Function

            End If
             
            sql = "SELECT cCode,cName from V_HY_FYSL_Contract_refer3  where  (cName='" & retvalue & "' or cCode='" & retvalue & "')  and  id<>'" & lngVoucherID & "'"
            Set rs = New ADODB.Recordset
            rs.Open sql, g_Conn, 1, 1

            If rs.EOF Then
                MsgBox "项目编号不存在", vbInformation, "提示"
                bChanged = Cancel

                Exit Function

            Else
                Voucher.headerText("proccode") = rs.Fields("cCode") & ""
                Voucher.headerText("procname") = rs.Fields("cName") & ""

                If sMetaItemName Like "*code" Then
                    retvalue = rs.Fields("cCode") & ""
                Else
                    retvalue = rs.Fields("cName") & ""
                End If
            End If
                
        Case LCase("procode")
          
            If retvalue = "" Then
                Voucher.headerText("procode") = ""
                Voucher.headerText("engproperties") = ""
                Voucher.headerText("procname") = ""
                
                Voucher.headerText("progressmoney") = ""
                Voucher.headerText("ccname") = ""
                 Voucher.headerText("mverdata") = ""
                Voucher.headerText("progressdesc") = ""
                 Voucher.headerText("cname") = ""
                
                Exit Function

            End If
             
            sql = "SELECT cCode,engproperties,procname,progressmoney,ccname,dVeriDate,progressdesc from V_HY_FYSL_Measurement  where    cCode='" & retvalue & "'   "
            Set rs = New ADODB.Recordset
            rs.Open sql, g_Conn, 1, 1

            If rs.EOF Then
                MsgBox "项目计量进度单号不存在", vbInformation, "提示"
                bChanged = Cancel

                Exit Function

            Else
                Voucher.headerText("procode") = rs.Fields("cCode") & ""
                Voucher.headerText("engproperties") = rs.Fields("engproperties") & ""
                Voucher.headerText("procname") = rs.Fields("procname") & ""
                retvalue = rs.Fields("cCode") & ""
                Voucher.headerText("progressmoney") = rs.Fields("progressmoney") & ""
                Voucher.headerText("ccname") = rs.Fields("ccname") & ""
                 Voucher.headerText("mverdata") = rs.Fields("dVeriDate") & ""
                  Voucher.headerText("progressdesc") = rs.Fields("progressdesc") & ""
                  Voucher.headerText("cname") = rs.Fields("progressdesc") & ""
                 
                    
            End If
                
        Case LCase$("appprice")

            If retvalue = "" Then
                Voucher.headerText("appprice") = ""
                
                Exit Function

            End If
           If Voucher.headerText("sourcetype") = "参照合同" Then
            If retvalue <= 0 Then
                MsgBox "申请金额不能小于等于0,请修改"
                Voucher.headerText("appprice") = ""

                Exit Function

            End If
          End If
            If Voucher.headerText("contype") = "普通合同" And Voucher.headerText("sourcetype") = "参照合同" Then
                If Null2Something(Voucher.headerText("conpaymoney")) <> "" Then
                    'If Null2Something(Voucher.headerText("totalappmoney"), 0) > Null2Something(Voucher.headerText("totalpaymoney"), 0) Then
                        If val(Null2Something(Voucher.headerText("appprice"), 0)) > val(Null2Something(Voucher.headerText("conpaymoney"), 0)) + val(Null2Something(Voucher.headerText("designmoney"), 0)) - val(Null2Something(Voucher.headerText("addesignmoney"), 0)) - val(Null2Something(Voucher.headerText("totalappmoney"), 0)) + val(numappprice) Then
                            MsgBox "申请金额不能大于合同核定合同总额（如合同未结算则为合同金额）- 累计申请金额-预收设计费 ", vbInformation, "提示"
                             bChanged = Cancel
                            Exit Function

                        End If

'                    Else
'
'                        If Null2Something(Voucher.headerText("appprice"), 0) > Val(Null2Something(Voucher.headerText("conpaymoney"), 0)) + Val(Null2Something(Voucher.headerText("designmoney"), 0)) - Val(Null2Something(Voucher.headerText("totalpaymoney"), 0)) - Val(Null2Something(Voucher.headerText("addesignmoney"), 0)) + Val(numappprice) Then
'                            MsgBox "申请金额不能大于合同核定合同总额（如合同未结算则为合同金额）- 累计申请金额-预收设计费", vbInformation, "提示"
'                             bChanged = Cancel
'                            Exit Function
'
'                        End If
'                    End If
                    
                Else
               
                   ' If Null2Something(Voucher.headerText("totalappmoney"), 0) > Null2Something(Voucher.headerText("totalpaymoney"), 0) Then
                        If val(Null2Something(Voucher.headerText("appprice"), 0)) > val(Null2Something(Voucher.headerText("conmoney"), 0)) + val(Null2Something(Voucher.headerText("designmoney"), 0)) - val(Null2Something(Voucher.headerText("totalappmoney"), 0)) - val(Null2Something(Voucher.headerText("addesignmoney"), 0)) + val(numappprice) Then
                            MsgBox "申请金额不能大于合同核定合同总额（如合同未结算则为合同金额）- 累计申请金额-预收设计费 ", vbInformation, "提示"
                             bChanged = Cancel
                            Exit Function

                        End If

'                    Else
'
'                        If Null2Something(Voucher.headerText("appprice"), 0) > Val(Null2Something(Voucher.headerText("conmoney"), 0)) + Val(Null2Something(Voucher.headerText("designmoney"), 0)) - Val(Null2Something(Voucher.headerText("totalpaymoney"), 0)) - Val(Null2Something(Voucher.headerText("addesignmoney"), 0)) + Val(numappprice) Then
'                            MsgBox "申请金额不能大于合同核定合同总额（如合同未结算则为合同金额）- 累计申请金额-预收设计费 ", vbInformation, "提示"
'                             bChanged = Cancel
'                            Exit Function
'
'                        End If
'                    End If
               
                End If
    
      
        End If
           
End Select

bChanged = success

Exit Function

ErrHandler:

MsgBox Err.Description, vbInformation, GetString("U8.DZ.JA.Res030")
End Function

'表体参照
Public Function VoucherbodyBrowUser(Voucher As Object, ByVal row As Long, ByVal Col As Long, sRet As Variant, referpara As UAPVoucherControl85.ReferParameter)

    On Error GoTo Err_Handler
    Dim oDefPro As Object
    Dim sMetaItemXML As String
    Dim sMetaXML As String
    Dim Ref As UFReferC.UFReferClient
    Dim sql As String
    Dim oRefSelect As Object                               '批号，入库单号参照
    Dim rstTmp As New ADODB.Recordset
    Dim sWhCode As String                                  '仓库编码
    Dim sInvCode As String                                 '存货编码

    sMetaItemXML = Voucher.ItemState(Col, sibody).sFieldName

    If sMetaItemXML <> "cinvcode" And Voucher.bodyText(row, "cinvcode") & "" = "" Then
        MsgBox GetString("U8.DZ.JA.Res630"), vbInformation, GetString("U8.DZ.JA.Res030")
        sRet = ""
        referpara.Cancel = True
        Exit Function
    End If

    '发货仓库
    sWhCode = Voucher.bodyText(row, "cwhcode")
    '存货编码
    sInvCode = Voucher.bodyText(row, "cinvcode")

    If LCase(sMetaItemXML) = "cbatch" Or LCase(sMetaItemXML) = "cinvouchcode" Then    '批号和入库单号参照所需处理
        '批次与入库单号参照相关参数  -chenliangc
        Dim i As Integer
        Dim sFree As Collection                            '自由项集合
        Dim errStr As String
        Dim sSql As String
        Dim strFilter As String                            '过滤条件串
        Dim iquantity As Double                            '数量
        Dim iNum As Double                                 '件数
        Dim iExchange As Double                            '换算率
        Dim sFreeName As String                            '自由项字段名
        Dim sBatch As String                               '批次

        '********************************************
        '2008-11-17
        '为匹配872中LP件多种销售跟踪方式的处理
        Dim sSosId As String                               '销售订单行ID
        Dim sDemandType As String                          '销售订单类型
        Dim sDemandCode As String                          '销售订单分类号
        Dim lDemandCode As Long                            '整数型订单行号
        Dim j As Long
        'Dim domline As DOMDocument

        '********************************************
        '销售订单行ID

        Set oRefSelect = CreateObject("USCONTROL.RefSelect")    '批号参照组件

        '/************批号参照相关参数初始化*******************/ chenliangc
        sSosId = Voucher.bodyText(row, "isosid")

        Call GetSoDemandType(sSosId, sDemandType, sDemandCode, g_Conn)
        If IsNumeric(sDemandCode) Then
            lDemandCode = CLng(sDemandCode)
        Else
            lDemandCode = 0
        End If

        '本次调拨数量
        iquantity = ConvertStrToDbl(Voucher.bodyText(row, "iquantity"))
        '换算率
        iExchange = ConvertStrToDbl(Voucher.bodyText(row, "iinvexchrate"))
        '件数
        If iExchange = 0 Then
            iNum = 0
        Else
            iNum = iquantity / iExchange
        End If

        If Col > 0 Then
            sBatch = Voucher.bodyText(row, "cbatch")
        End If

        '自由项集合
        Set sFree = New Collection
        For i = 1 To 10
            sFree.Add Null2Something(Voucher.bodyText(row, "cfree" & i))
        Next
        '/*****************************************************/
    End If


    '表体自定义项参照
    If LCase(sMetaItemXML) Like "cdefine*" Or LCase(sMetaItemXML) Like "cbdefine*" Then

        referpara.Cancel = True
        Set oDefPro = New U8DefPro.clsDefPro
        If oDefPro.Init(False, g_oLogin.UfDbName, g_oLogin.cUserId) Then
            '0:'手工输入;1:'系统档案;2:'单据
            Dim arr As Variant
            arr = Split(Voucher.ItemState(Col, 1).sDataRule, ",")
            '(1)当表体自定义项来源于基础档案时，arr(0) 是基础档案的表名；(2)当表体自定义项来源于单据时，arr(0) 是单据的类型（如：采购入库单(24)）
            '而接口：GetRefVal 在(1)时参数sCardNumber 是没有实际意义的；在(2)时参数sTableName 是没有实际意义的！
            If UBound(arr) > 0 Then
                sRet = oDefPro.GetRefVal(Voucher.ItemState(Col, 1).nDataSource, sibody, Voucher.ItemState(Col, 1).sFieldName, arr(0), arr(1), arr(0), sRet, False, 5, 1)
            Else
                sRet = oDefPro.GetRefVal(Voucher.ItemState(Col, 1).nDataSource, sibody, Voucher.ItemState(Col, 1).sFieldName, Voucher.ItemState(Col, 1).sTableName, Voucher.ItemState(Col, 1).sFieldName, gstrCardNumber, sRet, False, 5, 1)
            End If
        End If



        '表体自由项参照
    ElseIf LCase(sMetaItemXML) Like "cfree*" Then
        referpara.Cancel = True
        Set oDefPro = New U8DefPro.clsDefPro
        If oDefPro.InitNew(g_oLogin, False) Then
            sRet = oDefPro.GetStruFreeRefVal(sInvCode, Voucher.ItemState(Col, 1).sFieldName, sRet, False, 5, 1)
        End If



        '项目大类
    ElseIf LCase(sMetaItemXML) = "citem_class" Then
        referpara.Cancel = True

        Set Ref = New UFReferC.UFReferClient
        Ref.SetLogin g_oLogin

        Ref.SetRWAuth "", "", False
        clsbill.RefItemClass Voucher.bodyText(row, Col), sql
        If Ref.StrRefInit(g_oLogin, False, "", sql, GetResString("U8.ST.USKCGLSQL.clsbasebillhandler.01485")) = True Then
            Ref.Show
        End If
        If Not (Ref.recmx Is Nothing) Then
            sRet = Ref.recmx("项目大类")
            Voucher.bodyText(row, "citem_cname") = Ref.recmx("大类名称")
            Voucher.bodyText(row, "cItemCode") = ""
            Voucher.bodyText(row, "cName") = ""
        End If
        Ref.SetRWAuth "", "", True


        '项目参照
    ElseIf LCase(sMetaItemXML) = "citemcode" Then
        referpara.Cancel = True
        Dim sFilters As String

        sFilters = sRet
        If sFilters <> "" Then
            sFilters = "(citemcode like N'%" & sFilters & "%' or citemname like N'%" & sFilters & "%') "
            If Voucher.bodyText(row, "cItem_class") = "ch" Then
                sFilters = "(cinvcode like '%" & sRet & "%' or cinvname like '%" & sRet & "%') "
            End If
            sFilters = sFilters & " and (isnull(bclose,0)=0) "
        Else
            sFilters = " isnull(bclose,0)=0 "
        End If
        '**************************
        Set Ref = New UFReferC.UFReferClient
        Ref.SetLogin g_oLogin

        Ref.SetRWAuth "", "", False
        ' 如果先参照项目编码,则公共参照控件不提供模糊参照,所以要先参照
        ' 项目大类才行
        If Voucher.bodyText(row, "citem_class") = "" Then
            Ref.EnumRefInit g_oLogin, enuTreeViewAndGrid, False, enuItem, sFilters
        Else
            Ref.ItemRefInit g_oLogin, False, Voucher.bodyText(row, "citem_class"), sFilters
        End If
        '*********************************
        Ref.Show
        If Voucher.bodyText(row, "citem_class") = "" Then
            If Not Ref.RstSelClass Is Nothing Then
                If Not Ref.RstSelClass.EOF Then
                    Voucher.bodyText(row, "citem_class") = Ref.RstSelClass("citem_class")
                    Voucher.bodyText(row, "citem_cname") = Ref.RstSelClass("cItem_Name")
                End If
            End If
        End If
        If Not Ref.recmx Is Nothing Then
            If Voucher.bodyText(row, "citem_class") = "ch" Then
                sRet = Ref.recmx.Fields("cinvcode")
                Voucher.bodyText(row, "cName") = Ref.recmx.Fields("cinvname")
            Else
                sRet = Ref.recmx.Fields("citemcode")
                Voucher.bodyText(row, "cName") = Ref.recmx.Fields("citemname")
            End If
            ' 如果强行给单据控件的行列重新赋值后,导致不走BODYCELLCHECK事件.
            Voucher.bodyText(row, sMetaItemXML) = sRet
            Voucher.ProtectUnload2
        End If
        Ref.SetRWAuth "", , True





        '存货编码
    ElseIf LCase(sMetaItemXML) = "cinvcode" Then

        sMetaXML = "<Ref><RefSet bAuth='" & IIf(bInv_ControlAuth, "1", "0") & "' authFunID='W' bMultiSel= '1' /></Ref>"

        '            referpara.Cancel = False
        referpara.id = "Inventory_AA"
        referpara.RetField = "cinvcode"

        'referpara.sSql = " ('" & g_oLogin.CurDate & "' <=isnull(dEDate,'2099-12-31')  and bTrack<>1 and bInvQuality <> 1 and bSerial <> 1 )"
        referpara.sSql = " ('" & g_oLogin.CurDate & "' <=isnull(dEDate,'2099-12-31')  "
'        If isQualityInv = False Then
'            referpara.sSql = referpara.sSql + " and bInvQuality <> 1 "
'        End If
        
        If sAuth_invW <> "" Then
            referpara.sSql = referpara.sSql + " and #FN[iid] in (" & sAuth_invW & ")"
        End If
        
        referpara.sSql = referpara.sSql + " ) "
        referpara.ReferMetaXML = sMetaXML
        flag = False
        '仓库
    ElseIf LCase(sMetaItemXML) = "cwhcode" Or LCase(sMetaItemXML) = "cwhname" Then

        sMetaXML = "<Ref><RefSet bAuth='" & IIf(bWareHouse_ControlAuth, "1", "0") & "' authFunID='W' bMultiSel= '0' /></Ref>"
        referpara.id = "Warehouse_AA"                      ' "Warehouse_AA"
        referpara.RetField = "cwhcode"
        referpara.sSql = " ('" & CDate(Mid(IIf(IsBlank(Null2Something(Voucher.headerText(StrdDate))), g_oLogin.CurDate, Voucher.headerText(StrdDate)), 1, 10)) & "' < isnull(dWhEndDate,'2099-12-31'))  and bProxyWh=0"
        If sAuth_WareHouseW <> "" Then
            referpara.sSql = referpara.sSql & " and (#FN[cwhcode] in (" & sAuth_WareHouseW & "))"
        End If
       
        referpara.ReferMetaXML = sMetaXML

        '货位-
    ElseIf LCase(sMetaItemXML) = LCase("cPosition") Or LCase(sMetaItemXML) = LCase("cPosition2") Then
        '           If Voucher.bodyText(row, "cwhcode") = "" Then
        '                MsgBox GetString("U8.DZ.JA.Res1840"), vbInformation, GetString("U8.DZ.JA.Res030")
        '               sRet = ""
        '               referpara.Cancel = True
        '             End If

        sMetaXML = "<Ref><RefSet bAuth='" & IIf(bPosition_ControlAuth, "1", "0") & "' authFunID='' bMultiSel= '0' /></Ref>"
        referpara.id = "Position_AA_JA"
        referpara.ReferMetaXML = sMetaXML
        referpara.RetField = "cposcode"
        referpara.sSql = "cwhcode = '" & Voucher.bodyText(row, "cwhcode") & "'"
        referpara.sSql = referpara.sSql & IIf(sAuth_PositionW = "", "", " and cPosCode in (" & sAuth_PositionW & ")")

        '   referpara.sSql = "cwhcode = '" & Voucher.bodyText(row, "cwhcode") & "'"




        '辅计量单位,默认取库存计量单位
        '只有固定换算率可以修改辅助计量单位
    ElseIf LCase(sMetaItemXML) = "cinva_unit" Then
        sMetaXML = "<Ref><RefSet bAuth='0' bMultiSel= '0' /></Ref>"
        referpara.id = "ComputationUnit_AA"
        referpara.ReferMetaXML = sMetaXML
        referpara.sSql = " cgroupcode='" & Voucher.bodyText(row, "cGroupCode") & "' and (cComunitCode like '%" & sRet & "%' or cComunitName like '%" & sRet & "%')"

        '价格参照
    ElseIf LCase(sMetaItemXML) = "iquotedprice" Or LCase(sMetaItemXML) = "itaxunitprice" Or LCase(sMetaItemXML) = "iunitprice" Or LCase(sMetaItemXML) = "kl" Then

        referpara.RetField = sMetaItemXML
        BrowsePrice referpara, sMetaItemXML, row, Voucher, "97"    '97销售订单


        '批次参照 chenliangc
    ElseIf LCase(sMetaItemXML) = "cbatch" Then

        referpara.Cancel = True
'        If sWhCode = "" Then
'            MsgBox GetString("U8.DZ.JA.Res640"), vbInformation, GetString("U8.DZ.JA.Res030")
'            Exit Function
'        End If

        referpara.Cancel = True

        Set Ref = New UFReferC.UFReferClient
        Ref.SetLogin g_oLogin

        Ref.SetRWAuth "", "", False


        '871和872调用库存函数时，传递的参数类型发生变化
        '871的订单行参数要求是整数型，872的改成字符型，因此需要单独处理
        If gU8Version = "872" Then
            clsbill.RefBatchList sSql, errStr, sWhCode, sInvCode, sFree, sBatch, , CLng(sDemandType), sDemandCode
            '弹出通用的批次参照界面
            oRefSelect.Refer g_oLogin, sWhCode, sInvCode, sFree, iquantity, iNum, iExchange, , sSql & strFilter, , True, "", CLng(sDemandType), sDemandCode, ""
        Else
            clsbill.RefBatchList sSql, errStr, sWhCode, sInvCode, sFree, sBatch, , CLng(sDemandType), lDemandCode
            '弹出通用的批次参照界面
            oRefSelect.Refer g_oLogin, sWhCode, sInvCode, sFree, iquantity, iNum, iExchange, , sSql & strFilter, , True, "", CLng(sDemandType), lDemandCode, ""
        End If

        If Not oRefSelect.ReturnData Is Nothing Then
            Set rstTmp = oRefSelect.ReturnData
            If rstTmp.RecordCount = 1 Then
                sRet = rstTmp.Fields("批号")

                '带出自由项
                For i = 0 To rstTmp.Fields.Count - 1
                    sFreeName = IIf(IsNull(rstTmp.Fields(i).Properties("BASECOLUMNNAME")), "", rstTmp.Fields(i).Properties("BASECOLUMNNAME"))
                    If LCase(Mid(sFreeName, 1, 5)) = "cfree" Then
                        Voucher.bodyText(row, sFreeName) = Null2Something(rstTmp.Fields(i).Value)
                    End If
                    'by liwqa 带出批次属性
                    If LCase(Mid(sFreeName, 1, 14)) = "cbatchproperty" Then
                        Voucher.bodyText(row, sFreeName) = Null2Something(rstTmp.Fields(i).Value)
                    End If
                Next
                
                '这两行代码不知为何注掉了，20120702放开
                Voucher.bodyText(row, "iquantity") = Null2Something(rstTmp.Fields("出库数量"))
                Voucher.bodyText(row, "inum") = Null2Something(rstTmp.Fields("出库件数"))

                Voucher.bodyText(row, "dmadedate") = Null2Something(rstTmp.Fields("生产日期"))
                Voucher.bodyText(row, "dvdate") = Null2Something(rstTmp.Fields("失效日期"))
                Voucher.bodyText(row, "dexpirationdate") = Null2Something(rstTmp.Fields("有效期计算项"))
                Voucher.bodyText(row, "cexpirationdate") = Null2Something(rstTmp.Fields("有效期至"))
                
                Voucher.bodyText(row, "imassdate") = Null2Something(rstTmp.Fields("保质期"))
                Voucher.bodyText(row, "cmassunit") = Null2Something(rstTmp.Fields("保质期单位"))
                Voucher.bodyText(row, "iexpiratdatecalcu") = Null2Something(rstTmp.Fields("有效期推算方式"))



            ElseIf rstTmp.RecordCount > 1 Then
                For j = 1 To rstTmp.RecordCount
                    Voucher.DuplicatedLine row

                    '                Set domline = Voucher.GetLineDom(row)
                    '                Voucher.AddLine Voucher.BodyRows + 1
                    '                Voucher.UpdateLineData domline, Voucher.BodyRows
                    '带出自由项
                    For i = 0 To rstTmp.Fields.Count - 1
                        sFreeName = IIf(IsNull(rstTmp.Fields(i).Properties("BASECOLUMNNAME")), "", rstTmp.Fields(i).Properties("BASECOLUMNNAME"))
                        If LCase(Mid(sFreeName, 1, 5)) = "cfree" Then
                            Voucher.bodyText(Voucher.BodyRows, sFreeName) = Null2Something(rstTmp.Fields(i).Value)
                        End If
                        'by liwqa 带出批次属性
                        If LCase(Mid(sFreeName, 1, 14)) = "cbatchproperty" Then
                            Voucher.bodyText(Voucher.BodyRows, sFreeName) = Null2Something(rstTmp.Fields(i).Value)
                        End If
                    Next

                    '带出批次信息
                    If Not IsNull(Voucher.bodyText(Voucher.BodyRows, "cbatch")) Then
                        Voucher.bodyText(Voucher.BodyRows, "cbatch") = Null2Something(rstTmp.Fields("批号"))
                    End If
                    Voucher.bodyText(Voucher.BodyRows, "iquantity") = Null2Something(rstTmp.Fields("出库数量"))
                    Voucher.bodyText(Voucher.BodyRows, "inum") = Null2Something(rstTmp.Fields("出库件数"))

                    Voucher.bodyText(Voucher.BodyRows, "dmadedate") = Null2Something(rstTmp.Fields("生产日期"))
                    Voucher.bodyText(Voucher.BodyRows, "dvdate") = Null2Something(rstTmp.Fields("失效日期"))
                    Voucher.bodyText(Voucher.BodyRows, "dexpirationdate") = Null2Something(rstTmp.Fields("有效期计算项"))
                    Voucher.bodyText(Voucher.BodyRows, "cexpirationdate") = Null2Something(rstTmp.Fields("有效期至"))
                    
                    Voucher.bodyText(Voucher.BodyRows, "imassdate") = Null2Something(rstTmp.Fields("保质期"))
                    Voucher.bodyText(Voucher.BodyRows, "cmassunit") = Null2Something(rstTmp.Fields("保质期单位"))
                    Voucher.bodyText(Voucher.BodyRows, "iexpiratdatecalcu") = Null2Something(rstTmp.Fields("有效期推算方式"))
                    rstTmp.MoveNext
                Next
                Voucher.DelLine row
            End If
        End If

        '入库单号参照-chenliangc
    ElseIf LCase(sMetaItemXML) = "cinvouchcode" Then
        referpara.Cancel = True
        If sWhCode = "" Then
            MsgBox GetString("U8.DZ.JA.Res640"), vbInformation, GetString("U8.DZ.JA.Res030")
            retvalue = ""
            Exit Function
        End If

        Dim sInVouchCode As String                         '入库单号


        sInVouchCode = Voucher.bodyText(row, "cinvouchcode")


        '871和872调用库存函数时，传递的参数类型发生变化
        '871的订单行参数要求是整数型，872的改成字符型，因此需要单独处理
        If gU8Version = "872" Then
            clsbill.RefInVouchList sSql, "", sWhCode, sInvCode, sFree, sInVouchCode, True, sBatch, CLng(sDemandType), sDemandCode, ""
            oRefSelect.Refer g_oLogin, sWhCode, sInvCode, sFree, iquantity, iNum, iExchange, RefInVouch, sSql, False, True, "12", CLng(sDemandType), sDemandCode, ""
        Else
            clsbill.RefInVouchList sSql, "", sWhCode, sInvCode, sFree, sInVouchCode, True, sBatch, CLng(sDemandType), lDemandCode, ""
            oRefSelect.Refer g_oLogin, sWhCode, sInvCode, sFree, iquantity, iNum, iExchange, RefInVouch, sSql, False, True, "12", CLng(sDemandType), lDemandCode, ""
        End If

        If Not oRefSelect.ReturnData Is Nothing Then
            Set rstTmp = oRefSelect.ReturnData
            If rstTmp.RecordCount = 1 Then
                Voucher.bodyText(row, "cinvouchcode") = rstTmp.Fields("入库单号")
                Voucher.bodyText(row, "rdsid") = rstTmp.Fields("入库系统编号")

                '带出自由项
                For i = 0 To rstTmp.Fields.Count - 1
                    sFreeName = IIf(IsNull(rstTmp.Fields(i).Properties("BASECOLUMNNAME")), "", rstTmp.Fields(i).Properties("BASECOLUMNNAME"))
                    If LCase(Mid(sFreeName, 1, 5)) = "cfree" Then
                        Voucher.bodyText(row, sFreeName) = Null2Something(rstTmp.Fields(i).Value)
                    End If
                Next

                '带出批次信息
                If Not IsNull(Voucher.bodyText(row, "cbatch")) Then
                    Voucher.bodyText(row, "cbatch") = Null2Something(rstTmp.Fields("批号"))
                End If

                Voucher.bodyText(row, "iquantity") = Null2Something(rstTmp.Fields("结存数量"))
                Voucher.bodyText(row, "inum") = Null2Something(rstTmp.Fields("结存件数"))
                Voucher.bodyText(row, "dmadedate") = Null2Something(rstTmp.Fields("生产日期"))
                Voucher.bodyText(row, "dvdate") = Null2Something(rstTmp.Fields("失效日期"))
                Voucher.bodyText(row, "dexpirationdate") = Null2Something(rstTmp.Fields("有效期计算项"))
                Voucher.bodyText(row, "cexpirationdate") = Null2Something(rstTmp.Fields("有效期至"))
                
                Voucher.bodyText(row, "imassdate") = Null2Something(rstTmp.Fields("保质期"))
                Voucher.bodyText(row, "cmassunit") = Null2Something(rstTmp.Fields("保质期单位"))
                Voucher.bodyText(row, "iexpiratdatecalcu") = Null2Something(rstTmp.Fields("有效期推算方式"))

            ElseIf rstTmp.RecordCount > 1 Then
                For j = 1 To rstTmp.RecordCount
                    Voucher.DuplicatedLine row
                    '                 Set domline = Voucher.GetLineDom(row)
                    '                 Voucher.AddLine Voucher.BodyRows + 1
                    '                 Voucher.UpdateLineData domline, Voucher.BodyRows
                    '带出自由项
                    For i = 0 To rstTmp.Fields.Count - 1
                        sFreeName = IIf(IsNull(rstTmp.Fields(i).Properties("BASECOLUMNNAME")), "", rstTmp.Fields(i).Properties("BASECOLUMNNAME"))
                        If LCase(Mid(sFreeName, 1, 5)) = "cfree" Then
                            Voucher.bodyText(Voucher.BodyRows, sFreeName) = Null2Something(rstTmp.Fields(i).Value)
                        End If
                    Next

                    '带出批次信息
                    If Not IsNull(Voucher.bodyText(Voucher.BodyRows, "cbatch")) Then
                        Voucher.bodyText(Voucher.BodyRows, "cbatch") = Null2Something(rstTmp.Fields("批号"))
                    End If
                    Voucher.bodyText(Voucher.BodyRows, "iquantity") = Null2Something(rstTmp.Fields("出库数量"))
                    Voucher.bodyText(Voucher.BodyRows, "inum") = Null2Something(rstTmp.Fields("出库件数"))
                    Voucher.bodyText(Voucher.BodyRows, "dmadedate") = Null2Something(rstTmp.Fields("生产日期"))
                    Voucher.bodyText(Voucher.BodyRows, "dvdate") = Null2Something(rstTmp.Fields("失效日期"))
                    Voucher.bodyText(Voucher.BodyRows, "dexpirationdate") = Null2Something(rstTmp.Fields("有效期计算项"))
                    Voucher.bodyText(Voucher.BodyRows, "cexpirationdate") = Null2Something(rstTmp.Fields("有效期至"))
                    
                    Voucher.bodyText(Voucher.BodyRows, "imassdate") = Null2Something(rstTmp.Fields("保质期"))
                    Voucher.bodyText(Voucher.BodyRows, "cmassunit") = Null2Something(rstTmp.Fields("保质期单位"))
                    Voucher.bodyText(Voucher.BodyRows, "iexpiratdatecalcu") = Null2Something(rstTmp.Fields("有效期推算方式"))
                    rstTmp.MoveNext
                Next
                Voucher.DelLine row
            End If
        End If
    End If

    Exit Function

Err_Handler:
    MsgBox Err.Description, vbInformation, GetString("U8.DZ.JA.Res030")

End Function

'价格按钮
'报价参照列表

Public Sub PriceList(Voucher As Object)
    On Error GoTo Err_Handler

    Dim clsRefSrv As New U8RefService.IService
    Dim strError As String
    Dim referpara As UAPVoucherControl85.ReferParameter

    '价格参照
    '97 销售价格
    BrowsePrice referpara, "iquotedprice", Voucher.row, Voucher, "97"

    If referpara.id = "" Then Exit Sub
    Dim rstClass As ADODB.Recordset
    Dim rstGrid As ADODB.Recordset

    clsRefSrv.RefID = referpara.id
    clsRefSrv.FilterSQL = referpara.sSql
    clsRefSrv.MetaXML = referpara.ReferMetaXML

    If clsRefSrv.ShowRef(g_oLogin, rstClass, rstGrid, strError) Then
    End If

    If strError = "" And Not rstGrid Is Nothing Then
        Voucher.bodyText(Voucher.row, "iunitprice") = rstGrid("iquotedprice")
    End If

    Set clsRefSrv = Nothing
    Set rstClass = Nothing
    Set rstGrid = Nothing

    Exit Sub


Err_Handler:
    Set clsRefSrv = Nothing
    Set rstClass = Nothing
    Set rstGrid = Nothing
    MsgBox Err.Description, vbInformation, GetString("U8.DZ.JA.Res030")

End Sub


'表体校验
Public Function VoucherbodyCellCheck(Voucher As Object, retvalue As Variant, bChanged As Long, ByVal r As Long, ByVal c As Long, referpara As UAPVoucherControl85.ReferParameter)
    Dim sInvCode As String
    Dim sError As String
    Dim tmpstr As String
    Dim nRow As Long
    nRow = Voucher.row
    
    '记住最初行 for U8dp202764834
    Dim mdOldSelRow As Long
    mdOldSelRow = -1
    If mdOldSelRow = -1 Then mdOldSelRow = nRow
    
    Dim sMetaItemXML As String
    sMetaItemXML = Voucher.ItemState(c, sibody).sFieldName

    If sMetaItemXML <> "cinvcode" And Voucher.bodyText(r, "cinvcode") = "" Then
        MsgBox GetString("U8.DZ.JA.Res630"), vbInformation, GetString("U8.DZ.JA.Res030")
        bChanged = Cancel
        Exit Function
    End If

    '自定义项,自由项,项目大类,项目编码校验
    If LCase(sMetaItemXML) Like "cdefine*" Or LCase(sMetaItemXML) Like "cbdefine*" Or LCase(sMetaItemXML) Like "cfree*" Then

        Call VoucherbodyCellCheckDefine(Voucher, retvalue, bChanged, r, c, referpara)

    Else
        '表体自定义项、自由项 除外的其他栏目检验
        Call VoucherbodyCellCheckOther(Voucher, retvalue, bChanged, r, c, referpara)

    End If
    
    Select Case LCase(sMetaItemXML)
        'rowchange
        Case "cinvcode", "cinvname", "cwhcode", "cwhname", "cfree1", "cfree2", "cfree3", "cfree4", "cfree5", "cfree6", "cfree7", "cfree8", "cfree9", "cfree10", "ccvbatch", "cbatch", "cposition", "cposition2", "cinva_unit"
            sInvCode = Voucher.bodyText(nRow, "cInvCode")
            If sInvCode <> "" And bChanged = success Then
                Call ShowStock(Voucher, sInvCode, nRow)
            End If
    End Select
    Voucher.row = mdOldSelRow
End Function

'表体校验（自定义项、自由项 ）
Private Function VoucherbodyCellCheckDefine(Voucher As Object, retvalue As Variant, bChanged As Long, ByVal r As Long, ByVal c As Long, referpara As UAPVoucherControl85.ReferParameter)
    On Error GoTo Err_Handler

    Dim sMetaItemXML As String
    sMetaItemXML = Voucher.ItemState(c, sibody).sFieldName

    Dim bFixLen As Boolean
    Dim lFixLen As Long
    Dim arr As Variant
    Dim iRet As Integer
    Dim cDefValue As String
    Dim oDefPro As Object
    Dim RecCurRow As DOMDocument
    Dim STMsgTitle As String
    Set RecCurRow = New DOMDocument
    STMsgTitle = GetString("U8.DZ.JA.Res030")
    cDefValue = retvalue

    '自定义项
    If LCase(sMetaItemXML) Like "cdefine*" Or LCase(sMetaItemXML) Like "cbdefine*" Then
        arr = Split(Voucher.ItemState(c, 1).sDefaultValue, ",")
        If UBound(arr) > 0 Then
            bFixLen = CBool(arr(0))
            lFixLen = val(arr(1))
            Voucher.BodyRedraw = True
            If bFixLen And Len(retvalue) > lFixLen Then
                MsgBox GetResString("U8.ST.USKCGLSQL.frmqc.01806", Array("[" & Voucher.ItemState(c, 1).sCardFormula1)), vbOKOnly + vbInformation, STMsgTitle
                bChanged = Cancel
                Exit Function
            End If
        End If
        If Voucher.ItemState(c, 1).bValidityCheck Then
            '0:'手工输入;1:'系统档案;2:'单据
            arr = Split(Voucher.ItemState(c, 1).sDataRule, ",")

            Set oDefPro = New U8DefPro.clsDefPro
            If oDefPro.Init(False, g_oLogin.UfDbName, g_oLogin.cUserId) Then
                If UBound(arr) > 0 Then
                    iRet = oDefPro.ValidateAr(Voucher.ItemState(c, 1).nDataSource, 1, Voucher.ItemState(c, 1).sFieldName, arr(0), arr(1), cDefValue, gstrCardNumber, "", Voucher.ItemState(c, 1).bBuildArchives)
                Else
                    iRet = oDefPro.ValidateAr(Voucher.ItemState(c, 1).nDataSource, 1, Voucher.ItemState(c, 1).sFieldName, Voucher.ItemState(c, 1).sTableName, Voucher.ItemState(c, 1).sFieldName, cDefValue, gstrCardNumber, "", Voucher.ItemState(c, 1).bBuildArchives)
                End If
            End If
        End If

        '自由项
    Else
        Set oDefPro = New U8DefPro.clsDefPro
        If oDefPro.Init(False, g_oLogin.UfDbName, g_oLogin.cUserId) Then
            If cDefValue <> "" Then iRet = oDefPro.ValidateFreeAr(Voucher.ItemState(c, 1).sFieldName, cDefValue, Voucher.ItemState(c, 1).bBuildArchives)
        End If
        If bChanged = 2 Then                               '清楚批号
            '            Voucher.bodyText(r, "cbatch") = ""
            Voucher.bodyText(r, "cinvouchcode") = ""
            Voucher.bodyText(r, "cvouchcode") = ""
            Voucher.bodyText(r, "dmadedate") = ""
            Voucher.bodyText(r, "dvdate") = ""
            Voucher.bodyText(r, "dexpirationdate") = ""
            Voucher.bodyText(r, "cexpirationdate") = ""
        End If
    End If

    'iRet :0 校验成功；1 建档成功；-1 校验不成功；-2 建档不成功(只能返这四个值)
    If iRet < 0 Then
        If Voucher.ItemState(c, 1).bValidityCheck Or LCase(sMetaItemXML) Like "cfree*" Then
            Voucher.BodyRedraw = True
            MsgBox GetResString("U8.ST.USKCGLSQL.clsbasebillhandler.01453", Array(Voucher.ItemState(c, 1).sCardFormula1)), vbOKOnly + vbExclamation, STMsgTitle
            bChanged = Cancel
            Exit Function
        End If
    Else
        retvalue = cDefValue
        If LCase(sMetaItemXML) Like "cdefine*" Or LCase(sMetaItemXML) Like "cbdefine*" Then Exit Function
        Voucher.bodyText(r, LCase(sMetaItemXML)) = cDefValue
        Set RecCurRow = Voucher.GetLineDom(r)
    End If


    Exit Function

Err_Handler:
    MsgBox Err.Description, vbInformation, GetString("U8.DZ.JA.Res030")

End Function

'检查合法性 dxb
Private Function myCheckNumValue(ByVal retvalue As Variant, ByVal cCaption As String, Optional ByVal ctype As Byte = 0) As Boolean
    On Error GoTo lerr

    If retvalue = "" Then
        If ctype = 0 Then
            myCheckNumValue = True
            Exit Function
        Else
            MsgBox cCaption & GetString("U8.DZ.JA.Res650"), vbInformation, GetString("U8.DZ.JA.Res030")
            Exit Function
        End If
    End If

    If IsNumeric(retvalue) Then
        If CDbl(retvalue) < 0 Then
            MsgBox cCaption & GetString("U8.DZ.JA.Res650"), vbInformation, GetString("U8.DZ.JA.Res030")
            Exit Function
        End If
    Else
        Exit Function
    End If

lTrue:
    myCheckNumValue = True
    Exit Function
lerr:
End Function

'检查合法性 dxb
Private Function myCheckNumValue2(ByVal retvalue As Variant, ByVal cCaption As String, Optional ByVal ctype As Byte = 0) As Boolean
    On Error GoTo lerr

    If retvalue = "" Then
        If ctype = 0 Then
            myCheckNumValue2 = True
            Exit Function
        Else
            MsgBox cCaption & GetString("U8.DZ.JA.Res650"), vbInformation, GetString("U8.DZ.JA.Res030")
            Exit Function
        End If
    End If

    If IsNumeric(retvalue) Then
        If CDbl(retvalue) < 0 Then
            MsgBox cCaption & GetString("U8.DZ.JA.Res660"), vbInformation, GetString("U8.DZ.JA.Res030")
            Exit Function
        End If
    Else
        Exit Function
    End If

lTrue:
    myCheckNumValue2 = True
    Exit Function
lerr:
End Function

'表体校验（自定义项、自由项 之外的栏目）
'存货编码
'需要赋值的字段：存货名称、存货代码、规则型号
'               计量单位组编码、名称、主计量单位编码、名称、辅计量单位名称、编码、换算率
'               保质期单位、保质期天数
'               1-16存货自定义项

Private Function VoucherbodyCellCheckOther(Voucher As Object, retvalue As Variant, bChanged As Long, ByVal r As Long, ByVal c As Long, referpara As UAPVoucherControl85.ReferParameter)
    On Error GoTo Err_Handler

    Dim sMetaItemXML As String
    sMetaItemXML = Voucher.ItemState(c, sibody).sFieldName

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
    Dim sInvCode As String                                 '存货编码
    '存货编码
    sInvCode = Voucher.bodyText(r, "cinvcode")
    sWhCode = Voucher.bodyText(r, "cwhcode")

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

        Set oRefSelect = CreateObject("USCONTROL.RefSelect")    '批号参照组件

        '/************批号参照相关参数初始化*******************/ chenliangc
        sSosId = Voucher.bodyText(r, "isosid")

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
            If retvalue = "" Then
                Voucher.bodyText(r, "cwhname") = ""
                Voucher.bodyText(r, "cwhcode") = ""
                Voucher.bodyText(r, "cPosition") = ""
                Voucher.bodyText(r, "cPosition2") = ""
            Else
                Set rs = cWhCodeRefer(CStr(retvalue), IIf(IsBlank(Null2Something(Voucher.headerText(StrdDate))), g_oLogin.CurDate, Voucher.headerText(StrdDate)))
                If rs Is Nothing Or rs.State = 0 Then
                    ReDim varArgs(0)
                    varArgs(0) = retvalue
                    MsgBox GetStringPara("U8.DZ.JA.Res670", varArgs(0)), vbInformation, GetString("U8.DZ.JA.Res030")

                    '                MsgBox "仓库" & RetValue & "不存在，请重新输入", vbInformation, GetString("U8.DZ.JA.Res030")
                    bChanged = Cancel
                    Exit Function
                Else
                    If rs.RecordCount > 0 Then

                        '检查存货仓库对照关系

                        sql = "select  cvalue from accinformation where cname='bCheckWareInv' and csysid='ST'"
                        Set rst = New ADODB.Recordset
                        rst.Open sql, g_Conn
                        If rst.Fields("cvalue") = "True" Then
                            sql = "select cinvcode from  WhInvContrapose  where cinvcode='" & Voucher.bodyText(r, "cinvcode") & "' and  cwhcode='" & rs.Fields("cwhcode").Value & "'"
                            Set rst = New ADODB.Recordset
                            rst.Open sql, g_Conn
                            If rst.EOF Then
                                If MsgBox(GetStringPara("U8.DZ.JA.Res2180", rs.Fields("cwhcode").Value), vbInformation + vbYesNo, GetString("U8.DZ.JA.Res030")) = vbNo Then
                                    bChanged = Cancel
                                    Exit Function
                                End If
                            End If
                        End If

                        Voucher.bodyText(r, "cwhcode") = rs.Fields("cwhcode").Value
                        Voucher.bodyText(r, "cwhname") = rs.Fields("cwhname").Value
                        Voucher.bodyText(r, "cPosition") = ""
                        Voucher.bodyText(r, "cPosition2") = ""

                        'name触发时返回name,code触发时也要返回code
                        'If sMetaItemXML = "cwhname" Then retvalue = rs.Fields("cwhname").Value
                        retvalue = rs.Fields(sMetaItemXML).Value
                    End If
                End If
            End If


            '货位
        Case "cPosition", "cPosition2"
            If retvalue = "" Then
                Voucher.bodyText(r, "cPosition") = ""
                Voucher.bodyText(r, "cPosition2") = ""
            Else
                ' cwchcode = Voucher.bodyText(r, "cwhcode") & ""
                If Voucher.bodyText(r, "cwhcode") = "" Then
                    MsgBox GetString("U8.DZ.JA.Res1840"), vbInformation, GetString("U8.DZ.JA.Res030")
                    bChanged = Cancel
                    Exit Function
                End If

                Set rs = cPosionRefer(CStr(retvalue), Voucher.bodyText(r, "cwhcode"))
                If rs Is Nothing Or rs.State = 0 Then
                    ReDim varArgs(0)
                    varArgs(0) = retvalue
                    MsgBox GetStringPara("U8.DZ.JA.Res680", varArgs(0)), vbInformation, GetString("U8.DZ.JA.Res030")

                    '                MsgBox "货位" & RetValue & "不存在，请重新输入", vbInformation, GetString("U8.DZ.JA.Res030")
                    bChanged = Cancel
                    Exit Function
                Else
                    If rs.RecordCount > 0 Then

                        '检查存货货位对照关系

                        sql = "select  cvalue from accinformation where cname='bCheckInvPos' and csysid='ST'"
                        Set rst = New ADODB.Recordset
                        rst.Open sql, g_Conn
                        If rst.Fields("cvalue") = "True" Then
                            sql = "select cinvcode from  InvPosContrapose  where cinvcode='" & Voucher.bodyText(r, "cinvcode") & "' and  cposcode='" & rs.Fields("cposcode").Value & "'"
                            Set rst = New ADODB.Recordset
                            rst.Open sql, g_Conn
                            If rst.EOF Then
                                If MsgBox(GetStringPara("U8.DZ.JA.Res2190", rs.Fields("cposcode").Value), vbInformation + vbYesNo, GetString("U8.DZ.JA.Res030")) = vbNo Then
                                    bChanged = Cancel
                                    Exit Function
                                End If
                            End If
                        End If

                        Voucher.bodyText(r, "cPosition") = rs.Fields("cposcode").Value
                        Voucher.bodyText(r, "cPosition2") = rs.Fields("cposname").Value

                        'name触发时返回name,code触发时也要返回code
                        If sMetaItemXML = "cPosition2" Then retvalue = rs.Fields("cposname").Value
                        If sMetaItemXML = "cPosition" Then retvalue = rs.Fields("cposcode").Value
                    End If
                End If
            End If
        Case "cinvcode"
            Dim i As Integer
            Dim iRow As Long
            iRow = r                                       '当前行

            If retvalue = "" Then
                Voucher.DelLine iRow
                retvalue = Voucher.bodyText(iRow, sMetaItemXML)
            Else
                Set rs = cInvCodeRefer(CStr(retvalue))
                If rs Is Nothing Or rs.State = 0 Then
                    ReDim varArgs(0)
                    varArgs(0) = retvalue
                    MsgBox GetStringPara("U8.DZ.JA.Res690", varArgs(0)), vbInformation, GetString("U8.DZ.JA.Res030")

                    '                 MsgBox GetString("U8.DZ.JA.Res780") & RetValue & "不存在或者没有销售属性或者没有权限或者停用，请重新输入", vbInformation, GetString("U8.DZ.JA.Res030")
                    bChanged = Cancel
                    Exit Function

                Else

                    '检查存货仓库货位对照关系
                    If Voucher.bodyText(iRow, "cwhcode") & "" <> "" Then
                        sql = "select  cvalue from accinformation where cname='bCheckWareInv' and csysid='ST'"
                        Set rst = New ADODB.Recordset
                        rst.Open sql, g_Conn
                        If rst.Fields("cvalue") = "True" Then
                            sql = "select cinvcode from  WhInvContrapose  where cinvcode='" & rs.Fields("cinvcode").Value & "' and  cwhcode='" & Voucher.bodyText(iRow, "cwhcode") & "'"
                            Set rst = New ADODB.Recordset
                            rst.Open sql, g_Conn
                            If rst.EOF Then
                                If MsgBox(GetStringPara("U8.DZ.JA.Res2200", rs.Fields("cinvcode").Value), vbInformation + vbYesNo, GetString("U8.DZ.JA.Res030")) = vbNo Then
                                    bChanged = Cancel
                                    Exit Function
                                End If
                            End If
                        End If
                    End If

                    If Voucher.bodyText(iRow, "cPosition") & "" <> "" Then
                        sql = "select  cvalue from accinformation where cname='bCheckInvPos' and csysid='ST'"
                        Set rst = New ADODB.Recordset
                        rst.Open sql, g_Conn
                        If rst.Fields("cvalue") = "True" Then
                            sql = "select cinvcode from  InvPosContrapose  where cinvcode='" & rs.Fields("cinvcode").Value & "' and  cposcode='" & Voucher.bodyText(iRow, "cPosition") & "'"
                            Set rst = New ADODB.Recordset
                            rst.Open sql, g_Conn
                            If rst.EOF Then
                                If MsgBox(GetStringPara("U8.DZ.JA.Res2210", rs.Fields("cinvcode").Value), vbInformation + vbYesNo, GetString("U8.DZ.JA.Res030")) = vbNo Then
                                    bChanged = Cancel
                                    Exit Function
                                End If
                            End If
                        End If
                    End If
            End If
        End If



        '赋值
        '给存货编码等赋值
setvalue:
        If Not rs Is Nothing And rs.State <> 0 Then
            If Not rs.EOF Then
                SetBodyCellValue Voucher, rs, iRow
            End If
            If referpara.rstGrid Is Nothing Then
                retvalue = rs!cinvcode

            End If
            '                    Voucher.bodyText(iRow, sMetaItemXML) = Rs!cinvcode
            rs.Close
            Set rs = Nothing
        End If

selectmult:
        If Not referpara.rstGrid Is Nothing Then
            '参照存货编码时,多选
            If referpara.rstGrid.RecordCount > 1 Then

                If Voucher.BodyRowIsEmpty(iRow + 1) = False And flag = False Then flag = (MsgBox(GetString("U8.DZ.JA.Res700"), vbYesNo, GetString("U8.DZ.JA.Res030")) = vbYes)
                referpara.rstGrid.MoveNext

                Do While Not referpara.rstGrid.EOF
                    Set rs = cInvCodeRefer(CStr(referpara.rstGrid.Fields(sMetaItemXML)))
                    If flag Then
                        iRow = iRow + 1
                    Else
                        iRow = Voucher.BodyRows + 1
                    End If
                    If iRow > Voucher.BodyRows Then
                        Voucher.AddLine iRow
                    End If
                    Voucher.bodyText(iRow, sMetaItemXML) = CStr(referpara.rstGrid.Fields(sMetaItemXML))
                    GoTo setvalue

                    referpara.rstGrid.MoveNext
                Loop
            End If
        End If

    Case "iquantity"
        If Not myCheckNumValue(retvalue, GetString("U8.DZ.JA.Res710"), 1) Then
            bChanged = Cancel
            Exit Function
        End If

        '1无换算率
        If Voucher.bodyText(r, "igrouptype") = 0 Then
            GoTo lsuccess
        End If

        '2固定换算率
        If Voucher.bodyText(r, "igrouptype") = 1 Then
            Voucher.bodyText(r, "inum") = Voucher.bodyText(r, "iquantity") / IIf(ConvertStrToDbl(Voucher.bodyText(r, "iinvexchrate")) = 0, 1, ConvertStrToDbl(Voucher.bodyText(r, "iinvexchrate")))

            GoTo lsuccess
        End If
        '3浮动换算率
        If Voucher.bodyText(r, "igrouptype") = 2 Then
            If aGetFloatRateRule Then
                If Voucher.bodyText(r, "iinvexchrate") <> "" Then
                    If Voucher.bodyText(r, "iinvexchrate") = 0 Then
                        Voucher.bodyText(r, "inum") = 0
                    Else
                        Voucher.bodyText(r, "inum") = retvalue / Voucher.bodyText(r, "iinvexchrate")
                    End If
                Else
                    If Voucher.bodyText(r, "inum") <> "" Then
                        If Voucher.bodyText(r, "inum") = 0 Then
                            Voucher.bodyText(r, "iinvexchrate") = 0
                        Else
                            Voucher.bodyText(r, "iinvexchrate") = retvalue / Voucher.bodyText(r, "inum")
                        End If
                    End If
                End If
            Else
                If Voucher.bodyText(r, "inum") <> "" Then
                    If Voucher.bodyText(r, "inum") = 0 Then
                        Voucher.bodyText(r, "iinvexchrate") = 0
                    Else
                        Voucher.bodyText(r, "iinvexchrate") = retvalue / Voucher.bodyText(r, "inum")
                    End If
                Else
                    If Voucher.bodyText(r, "iinvexchrate") <> "" Then
                        If Voucher.bodyText(r, "iinvexchrate") = 0 Then
                            Voucher.bodyText(r, "inum") = 0
                        Else
                            Voucher.bodyText(r, "inum") = retvalue / Voucher.bodyText(r, "iinvexchrate")
                        End If
                    End If

                End If
            End If


        End If

        '件数
    Case "inum"
        If Not myCheckNumValue2(retvalue, GetString("U8.DZ.JA.Res720"), 0) Then
            bChanged = Cancel
            Exit Function
        End If

        '1无换算率
        If Voucher.bodyText(r, "igrouptype") = 0 Then
            GoTo lsuccess
        End If

        '2固定换算率
        If Voucher.bodyText(r, "igrouptype") = 1 Then
            If Voucher.bodyText(r, "inum") <> "" Then
                Voucher.bodyText(r, "iquantity") = Voucher.bodyText(r, "inum") * Voucher.bodyText(r, "iinvexchrate")
            Else
                Voucher.bodyText(r, "iquantity") = ""
            End If

            GoTo lsuccess
        End If

        '3浮动换算率
        If Voucher.bodyText(r, "igrouptype") = 2 Then
            If aGetFloatRateRule Then
                If retvalue = "" Then
                    Voucher.bodyText(r, "iquantity") = ""
                    Exit Function
                End If

                If Voucher.bodyText(r, "iquantity") <> "" Then
                    If retvalue = 0 Then
                        Voucher.bodyText(r, "iinvexchrate") = 0
                    Else
                        Voucher.bodyText(r, "iinvexchrate") = Voucher.bodyText(r, "iquantity") / retvalue
                    End If
                ElseIf Voucher.bodyText(r, "iinvexchrate") <> "" Then
                    If Voucher.bodyText(r, "iquantity") = "" Then
                        Voucher.bodyText(r, "iquantity") = retvalue * Voucher.bodyText(r, "iinvexchrate")
                    End If
                End If
            Else
                If Voucher.bodyText(r, "iinvexchrate") <> "" Then
                    If retvalue = "" Then
                        Voucher.bodyText(r, "iquantity") = ""
                    Else
                        Voucher.bodyText(r, "iquantity") = retvalue * Voucher.bodyText(r, "iinvexchrate")
                    End If
                End If
                If Voucher.bodyText(r, "iquantity") <> "" Then
                    If Voucher.bodyText(r, "iinvexchrate") = "" Then
                        If retvalue > 0 Then
                            Voucher.bodyText(r, "iinvexchrate") = Voucher.bodyText(r, "iquantity") / retvalue
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
                Voucher.bodyText(r, "inum") = ""
            Else
                Voucher.bodyText(r, "iquantity") = ""
            End If

            GoTo lsuccess
        End If

        If Not myCheckNumValue(retvalue, GetString("U8.DZ.JA.Res730"), 0) Then
            bChanged = Cancel
            Exit Function
        End If

        '变化关系
        If aGetFloatRateRule Then
            If Voucher.bodyText(r, "iquantity") <> "" Then
                If retvalue = 0 Then
                    Voucher.bodyText(r, "inum") = 0
                Else
                    Voucher.bodyText(r, "inum") = Voucher.bodyText(r, "iquantity") / IIf(ConvertStrToDbl(Voucher.bodyText(r, "iinvexchrate")) = 0, 1, ConvertStrToDbl(Voucher.bodyText(r, "iinvexchrate")))
                End If
            End If
            If Voucher.bodyText(r, "iquantity") = "" Then
                If Voucher.bodyText(r, "inum") <> "" Then Voucher.bodyText(r, "iquantity") = Voucher.bodyText(r, "inum") * Voucher.bodyText(r, "iinvexchrate")
            End If

        Else
            If Voucher.bodyText(r, "inum") <> "" Then _
                    Voucher.bodyText(r, "iquantity") = Voucher.bodyText(r, "inum") * Voucher.bodyText(r, "iinvexchrate")
            If Voucher.bodyText(r, "iquantity") <> "" Then
                If Voucher.bodyText(r, "inum") = "" Then
                    If retvalue = 0 Then
                        Voucher.bodyText(r, "inum") = 0
                    Else
                        Voucher.bodyText(r, "inum") = Voucher.bodyText(r, "iquantity") / IIf(ConvertStrToDbl(Voucher.bodyText(r, "iinvexchrate")) = 0, 1, ConvertStrToDbl(Voucher.bodyText(r, "iinvexchrate")))
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

        If (sMetaItemXML = "dmadedate" And Voucher.bodyText(r, "dmadedate") = "") Or (sMetaItemXML = "dvdate" And Voucher.bodyText(r, "dvdate") = "") Then
            Voucher.bodyText(r, "dmadedate") = ""
            Voucher.bodyText(r, "dvdate") = ""
            Voucher.bodyText(r, "dexpirationdate") = ""
            Voucher.bodyText(r, "cexpirationdate") = ""
        Else
            If Voucher.bodyText(r, "cMassUnit") <> "" And Voucher.bodyText(r, "imassdate") <> "" Then
                Select Case Voucher.bodyText(r, "cMassUnit")

                    Case "1":
                        If sMetaItemXML = "dmadedate" Then Voucher.bodyText(r, "dvdate") = DateAdd("yyyy", CInt(Voucher.bodyText(r, "imassdate")), CDate(Voucher.bodyText(r, "dmadedate")))
                        Voucher.bodyText(r, "dMadeDate") = DateAdd("yyyy", 0 - CInt(Voucher.bodyText(r, "imassdate")), CDate(Voucher.bodyText(r, "dvdate")))
                    Case "2":
                        If sMetaItemXML = "dmadedate" Then Voucher.bodyText(r, "dvdate") = DateAdd("m", CInt(Voucher.bodyText(r, "imassdate")), CDate(Voucher.bodyText(r, "dmadedate")))
                        Voucher.bodyText(r, "dMadeDate") = DateAdd("m", 0 - CInt(Voucher.bodyText(r, "imassdate")), CDate(Voucher.bodyText(r, "dvdate")))
                    Case "3":
                        If sMetaItemXML = "dmadedate" Then Voucher.bodyText(r, "dvdate") = DateAdd("d", CInt(Voucher.bodyText(r, "imassdate")), CDate(Voucher.bodyText(r, "dmadedate")))
                        Voucher.bodyText(r, "dMadeDate") = DateAdd("d", 0 - CInt(Voucher.bodyText(r, "imassdate")), CDate(Voucher.bodyText(r, "dvdate")))
                End Select

                Select Case Voucher.bodyText(r, "iExpiratDateCalcu")

                    Case "1":                              '"按月":
                        Voucher.bodyText(r, "dexpirationdate") = CDate(Voucher.bodyText(r, "dvdate")) - DatePart("d", CDate(Voucher.bodyText(r, "dvdate")))
                        Voucher.bodyText(r, "cexpirationdate") = Format(Voucher.bodyText(r, "dexpirationdate"), "yyyy-mm")
                    Case "2":                              '"按日":
                        Voucher.bodyText(r, "dexpirationdate") = DateAdd("d", -1, CDate(Voucher.bodyText(r, "dvdate")))
                        Voucher.bodyText(r, "cexpirationdate") = Voucher.bodyText(r, "dexpirationdate")
                End Select

                If DateDiff("d", CDate(Voucher.bodyText(r, "dvdate")), g_oLogin.CurDate) >= 0 Then
                    If MsgBox(GetString("U8.DZ.JA.Res740"), vbInformation + vbYesNo, GetString("U8.DZ.JA.Res030")) = vbNo Then
                        retvalue = ""
                        Voucher.bodyText(r, "dmadedate") = ""
                        Voucher.bodyText(r, "dvdate") = ""
                        Voucher.bodyText(r, "dexpirationdate") = ""
                        Voucher.bodyText(r, "cexpirationdate") = ""
                    End If
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

        If retvalue = "" Then
            Voucher.bodyText(r, "cinvouchcode") = ""
            Voucher.bodyText(r, "cvouchcode") = ""
            Voucher.bodyText(r, "dmadedate") = ""
            Voucher.bodyText(r, "dvdate") = ""
            Voucher.bodyText(r, "dexpirationdate") = ""
            Voucher.bodyText(r, "cexpirationdate") = ""

            '清空批次属性
            For i = 0 To 10
                Voucher.bodyText(r, "cbatchproperty" & CStr(i)) = ""
            Next i

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
                        Voucher.bodyText(r, "cfree1"), _
                        Voucher.bodyText(r, "cfree2"), _
                        Voucher.bodyText(r, "cfree3"), _
                        Voucher.bodyText(r, "cfree4"), _
                        Voucher.bodyText(r, "cfree5"), _
                        Voucher.bodyText(r, "cfree6"), _
                        Voucher.bodyText(r, "cfree7"), _
                        Voucher.bodyText(r, "cfree8"), _
                        Voucher.bodyText(r, "cfree9"), _
                        Voucher.bodyText(r, "cfree10"), _
                        CStr(Voucher.bodyText(r, c)), , CLng(sDemandType), sDemandCode, ""
            Else
                moStockPst.BatchList sSql, sWhCode, sInvCode, _
                        Voucher.bodyText(r, "cfree1"), _
                        Voucher.bodyText(r, "cfree2"), _
                        Voucher.bodyText(r, "cfree3"), _
                        Voucher.bodyText(r, "cfree4"), _
                        Voucher.bodyText(r, "cfree5"), _
                        Voucher.bodyText(r, "cfree6"), _
                        Voucher.bodyText(r, "cfree7"), _
                        Voucher.bodyText(r, "cfree8"), _
                        Voucher.bodyText(r, "cfree9"), _
                        Voucher.bodyText(r, "cfree10"), _
                        CStr(Voucher.bodyText(r, c)), , CLng(sDemandType), lDemandCode, ""
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

                Voucher.bodyText(r, "dmadedate") = ""
                Voucher.bodyText(r, "dvdate") = ""
                Voucher.bodyText(r, "dexpirationdate") = ""
                Voucher.bodyText(r, "cexpirationdate") = ""

            Else

                If rs.RecordCount = 1 Then

                    '带出自由项
                    For i = 0 To rs.Fields.Count - 1
                        sFreeName = IIf(IsNull(rs.Fields(i).Properties("BASECOLUMNNAME")), "", rs.Fields(i).Properties("BASECOLUMNNAME"))
                        If LCase(Mid(sFreeName, 1, 5)) = "cfree" Then
                            Voucher.bodyText(r, sFreeName) = Null2Something(rs.Fields(i).Value)
                        End If
                        'by liwqa 带出批次属性
                        If LCase(Mid(sFreeName, 1, 14)) = "cbatchproperty" Then
                            Voucher.bodyText(r, sFreeName) = Null2Something(rs.Fields(i).Value)
                        End If
                    Next

                    '                             Voucher.bodyText(r, "iquantity") = Null2Something(Rs.Fields("出库数量"))
                    '                             Voucher.bodyText(r, "inum") = Null2Something(Rs.Fields("出库件数"))

                    Voucher.bodyText(r, "dmadedate") = Null2Something(rs.Fields("生产日期"))
                    Voucher.bodyText(r, "dvdate") = Null2Something(rs.Fields("失效日期"))
                    Voucher.bodyText(r, "dexpirationdate") = Null2Something(rs.Fields("有效期计算项"))
                    Voucher.bodyText(r, "cexpirationdate") = Null2Something(rs.Fields("有效期至"))
                    
                    Voucher.bodyText(r, "imassdate") = Null2Something(rs.Fields("保质期"))
                    Voucher.bodyText(r, "cmassunit") = Null2Something(rs.Fields("保质期单位"))
                    Voucher.bodyText(r, "iexpiratdatecalcu") = Null2Something(rs.Fields("有效期推算方式"))

                    '如果跟踪型存货有多条记录，显示参照窗口。
                Else
                    Call VoucherbodyBrowUser(Voucher, r, c, retvalue, referpara)
                End If

            End If
            rs.Close
            Set rs = Nothing
        End If

        '出库跟踪入库校验 chenliangc
    Case "cinvouchcode"

        If sWhCode = "" Then
            MsgBox GetString("U8.DZ.JA.Res640"), vbInformation, GetString("U8.DZ.JA.Res030")
            retvalue = ""
            Exit Function
        End If

        If retvalue = "" Then
            Voucher.bodyText(r, "cbatch") = ""
            Voucher.bodyText(r, "cvouchcode") = ""
            Voucher.bodyText(r, "dmadedate") = ""
            Voucher.bodyText(r, "dvdate") = ""
            Voucher.bodyText(r, "dexpirationdate") = ""
            Voucher.bodyText(r, "cexpirationdate") = ""
            Exit Function
        End If

        Dim sInVouchCode As String                         '入库单号
        Dim sRdsID As String                               '入库单行ID


        '入库单号
        If c > 0 Then
            sInVouchCode = Voucher.bodyText(r, "cinvouchcode")
        End If

        '本次调拨数量
        iquantity = ConvertStrToDbl(Voucher.bodyText(r, "iquantity"))

        If c > 0 Then
            If sInVouchCode <> "" Then
                Dim moBatchPst As New BatchPst
                moBatchPst.login = mologin

                '871和872调用库存函数时，传递的参数类型发生变化
                '871的订单行参数要求是整数型，872的改成字符型，因此需要单独处理
                If gU8Version = "872" Then
                    moBatchPst.List sSql, sWhCode, sInvCode, _
                            Voucher.bodyText(r, "cfree1"), _
                            Voucher.bodyText(r, "cfree2"), _
                            Voucher.bodyText(r, "cfree3"), _
                            Voucher.bodyText(r, "cfree4"), _
                            Voucher.bodyText(r, "cfree5"), _
                            Voucher.bodyText(r, "cfree6"), _
                            Voucher.bodyText(r, "cfree7"), _
                            Voucher.bodyText(r, "cfree8"), _
                            Voucher.bodyText(r, "cfree9"), _
                            Voucher.bodyText(r, "cfree10"), _
                            sInVouchCode, sRdsID, , Voucher.bodyText(r, "cbatch"), CLng(sDemandType), sDemandCode, ""
                Else
                    moBatchPst.List sSql, sWhCode, sInvCode, _
                            Voucher.bodyText(r, "cfree1"), _
                            Voucher.bodyText(r, "cfree2"), _
                            Voucher.bodyText(r, "cfree3"), _
                            Voucher.bodyText(r, "cfree4"), _
                            Voucher.bodyText(r, "cfree5"), _
                            Voucher.bodyText(r, "cfree6"), _
                            Voucher.bodyText(r, "cfree7"), _
                            Voucher.bodyText(r, "cfree8"), _
                            Voucher.bodyText(r, "cfree9"), _
                            Voucher.bodyText(r, "cfree10"), _
                            sInVouchCode, sRdsID, , Voucher.bodyText(r, "cbatch"), CLng(sDemandType), lDemandCode, ""
                End If
                Set rs = New ADODB.Recordset
                rs.Open sSql, g_Conn, 1, 1
                If rs.RecordCount = 0 Then
                    ReDim varArgs(0)
                    varArgs(0) = Voucher.bodyText(r, "cinvcode")
                    MsgBox GetStringPara("U8.DZ.JA.Res750", varArgs(0)), vbInformation, GetString("U8.DZ.JA.Res030")
                    '                    MsgBox "跟踪型存货" & Voucher.bodyText(r, "cinvcode") & "指定的入库单号" & sInVouchCode & "不存在，请参照输入入库单号。", vbInformation, GetString("U8.DZ.JA.Res030")
                    retvalue = ""
                    Voucher.bodyText(r, "rdsid") = ""
                Else
                    '如果跟踪型存货只有一条记录
                    If rs.RecordCount = 1 Then
                        Voucher.bodyText(r, "rdsid") = rs("入库系统编号")

                        '带出自由项
                        For i = 0 To rs.Fields.Count - 1
                            sFreeName = IIf(IsNull(rs.Fields(i).Properties("BASECOLUMNNAME")), "", rs.Fields(i).Properties("BASECOLUMNNAME"))
                            If LCase(Mid(sFreeName, 1, 5)) = "cfree" Then
                                Voucher.bodyText(r, sFreeName) = Null2Something(rs.Fields(i).Value)
                            End If
                        Next

                        '带出批次信息
                        If Not IsNull(Voucher.bodyText(r, "cbatch")) Then
                            Voucher.bodyText(r, "cbatch") = Null2Something(rs.Fields("批号"))
                        End If


                        Voucher.bodyText(r, "iquantity") = Null2Something(rs.Fields("出库数量"))
                        Voucher.bodyText(r, "inum") = Null2Something(rs.Fields("出库件数"))

                        Voucher.bodyText(r, "dmadedate") = Null2Something(rs.Fields("生产日期"))
                        Voucher.bodyText(r, "dvdate") = Null2Something(rs.Fields("失效日期"))
                        Voucher.bodyText(r, "dexpirationdate") = Null2Something(rs.Fields("有效期计算项"))
                        Voucher.bodyText(r, "cexpirationdate") = Null2Something(rs.Fields("有效期至"))
                        Voucher.bodyText(r, "cinvouchcode") = Null2Something(rs.Fields("入库单号"))
                        
                        Voucher.bodyText(r, "imassdate") = Null2Something(rs.Fields("保质期"))
                        Voucher.bodyText(r, "cmassunit") = Null2Something(rs.Fields("保质期单位"))
                        Voucher.bodyText(r, "iexpiratdatecalcu") = Null2Something(rs.Fields("有效期推算方式"))

                        '如果跟踪型存货有多条记录，显示参照窗口。
                    Else
                        Call VoucherbodyBrowUser(Voucher, r, c, retvalue, referpara)

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
                " Where ComputationUnit.cGroupCode = N'" & Voucher.bodyText(r, "cgroupcode") & "'and (computationUnit.ccomUnitName like '" & retvalue & "%' or computationUnit.cComUnitcode like '" & retvalue & "%') " & _
                " Order by cComUnitcode ,iNumber ASC "
        Set rst = New ADODB.Recordset
        rst.Open strGrid, g_Conn

        '        If Voucher.bodyText(Row, "cincotermcode") = "" Then
        If rst.RecordCount > 0 Then
            retvalue = rst.Fields("ccomUnitName").Value
            Voucher.bodyText(r, "cunitid") = rst.Fields("cComUnitcode").Value
            Voucher.bodyText(r, "cinva_unit") = rst.Fields("ccomUnitName").Value
            Voucher.bodyText(r, "iinvexchrate") = rst.Fields("iChangRate").Value
            If Voucher.bodyText(r, "iquantity") <> "" And Voucher.bodyText(r, "iquantity") <> 0 Then
                Voucher.bodyText(r, "inum") = Voucher.bodyText(r, "iquantity") / rst.Fields("iChangRate").Value
            ElseIf Voucher.bodyText(r, "inum") <> "" And Voucher.bodyText(r, "inum") <> 0 Then
                Voucher.bodyText(r, "iquantity") = Voucher.bodyText(r, "inum") * rst.Fields("iChangRate").Value
            End If
        Else
            Voucher.bodyText(r, "cinva_unit") = ""
            retvalue = ""
        End If
        rst.Close

End Select

lsuccess:
bChanged = success

Exit Function

Err_Handler:
MsgBox Err.Description, vbInformation, GetString("U8.DZ.JA.Res030")

End Function


'价格参照
'遵循销售管理[销售选项]的价格策略
'skey 关键字 iquotedprice， "iunitprice", "itaxunitprice"

Public Function BrowsePrice(referpara As UAPVoucherControl85.ReferParameter, _
                            sKey As String, _
                            row As Long, _
                            Voucher As Object, _
                            strVouchType As String)    '单据类别

    Dim cCusCode As String
    Dim cinvcode As String
    Dim strExchName As String
    Dim strColumnKey As String
    Dim strWhere As String
    Dim strAuthtmp As String
    Dim strSql As String
    Dim i As Long
    Dim lngRow As Long
    Dim strOrder As String
    Dim strAuthKey As String

    Dim clsSAWeb As Object
    Dim myinfo As USSAServer.MyInformation                 '销售选项

    On Error GoTo Err_Handler

    '销售选项信息 myinfo
    Set clsSAWeb = Nothing
    Set clsSAWeb = New USSAServer.clsSystem
    clsSAWeb.Init g_oLogin
    clsSAWeb.INIMyInfor
    myinfo = clsSAWeb.SysInformation



    With Voucher

        '存货编码
        cCusCode = Voucher.headerText("ccuscode")
        cinvcode = Voucher.bodyText(.row, "cinvcode")
        If cinvcode = "" Then
            MsgBox GetString("U8.DZ.JA.Res630"), vbInformation, GetString("U8.DZ.JA.Res030")
            referpara.Cancel = True
            Set clsSAWeb = Nothing
            Exit Function
        End If

        '币种
        strExchName = .headerText("cexch_name")
        If strExchName = "" Then
            MsgBox GetString("U8.DZ.JA.Res760"), vbInformation, GetString("U8.DZ.JA.Res030")
            referpara.Cancel = True
            Set clsSAWeb = Nothing
            Exit Function
        End If


        Select Case LCase(sKey)

                '报价
            Case "iquotedprice"
                If myinfo.CostRefType = 0 Then             ''历次售价
                    Select Case myinfo.CostReferVouch
                        Case 0
                            strColumnKey = "SA_REF_SaleOrder_SA"
                            strAuthKey = "17"              '销售订单

                        Case 1
                            strColumnKey = "SA_REF_Dispatchlist_SA"

                            strAuthKey = "01"              '发货单
                        Case 2
                            strColumnKey = "SA_REF_SaleBillVouch_SA"

                            strAuthKey = "07"              '销售发票
                        Case 3
                            strColumnKey = "SA_REF_Quo_SA"

                            strAuthKey = "16"              '报价单
                    End Select

                    strWhere = "cinvcode='" & cinvcode & IIf(myinfo.CostRefCustomer = True, "' and ccuscode='" & cCusCode & "'", "'") & " and cexch_name='" & strExchName & "'"

                Else                                       ''各种报价
                    strColumnKey = "SA_REF_InvPrice_SA"
                    strWhere = " binvalid=0 and cinvcode='" & cinvcode & "'"
                    strAuthKey = "invprice"

                End If


                '无税单价、含税单价
            Case "iunitprice", "itaxunitprice"
                Select Case strVouchType
                    Case "97"
                        strColumnKey = "SA_REF_SaleOrder_SA"
                        strAuthKey = "17"                  '销售订单

                    Case "05", "06"
                        strColumnKey = "SA_REF_Dispatchlist_SA"

                        strAuthKey = "05"                  '委托代销发货单

                    Case "26", "27", "28", "29"
                        strColumnKey = "SA_REF_SaleBillVouch_SA"
                        strAuthKey = "07"

                    Case "16"
                        strColumnKey = "SA_REF_Quo_SA"
                        strAuthKey = "16"

                End Select
                strWhere = "cinvcode='" & cinvcode & IIf(myinfo.CostRefCustomer = True, "' and ccuscode='" & cCusCode & "'", "'") & " and cexch_name='" & strExchName & "'"

                '扣率
            Case "kl"
                strColumnKey = "SA_REF_QtyDiscount_SA"
                strWhere = "cinvcode='" & cinvcode & "'"
                strAuthKey = "quantitydisrate"

        End Select

        referpara.id = strColumnKey

        referpara.AutoDisplayText = False
        referpara.bValid = True
        referpara.sSql = strWhere & IIf(strAuthtmp = "", "", " and " & strAuthtmp)
        referpara.ReferMetaXML = "<Ref><RefSet bMultiSel= '0' /></Ref>"

    End With


    Set clsSAWeb = Nothing

    Exit Function

Err_Handler:
    Set clsSAWeb = Nothing
    MsgBox Err.Description, vbInformation, GetString("U8.DZ.JA.Res030")

End Function



'取价（单行、整张）
Public Function GetPrice(strKey As String, strVouchType As String, Voucher As Object)
    Dim domHead As New DOMDocument
    Dim domBody As New DOMDocument
    Dim errMsg As String
    Dim i As Long

    On Error GoTo Err_Handler
    If Voucher.BodyRows < 1 Then
        MsgBox GetString("U8.DZ.JA.Res770"), vbInformation, GetString("U8.DZ.JA.Res030")
        Exit Function
    End If

    If Voucher.BodyRowIsEmpty(Voucher.row) And strKey = "rowprice" Then
        MsgBox GetString("U8.DZ.JA.Res770"), vbInformation, GetString("U8.DZ.JA.Res030")
        Exit Function
    End If

    '销售选项信息 myinfo
    Dim clsSAWeb As Object
    Set clsSAWeb = Nothing
    Set clsSAWeb = New USSAServer.clsSystem
    clsSAWeb.Init g_oLogin
    clsSAWeb.INIMyInfor

    Dim clsVoucherCO As New VoucherCO_Sa.ClsVoucherCO_SA
    clsVoucherCO.Init strVouchType, g_oLogin, g_Conn, "CS", clsSAWeb


    If strKey = "allprice" Then
        If strVouchType <> "95" And strVouchType <> "92" And strVouchType <> "98" And strVouchType <> "99" Then
            For i = Voucher.BodyRows To 1 Step -1
                If Voucher.bodyText(i, "cinvcode") = "" Then
                    Voucher.DelLine i
                End If
            Next i
        End If
        Voucher.getVoucherDataXML domHead, domBody
    Else
        If Voucher.bodyText(Voucher.row, "cscloser") <> "" Then
            '            MsgBox GetString("U8.SA.clsOpSA.clscommcheck.00018")
            Exit Function
        End If
        Set domHead = Voucher.GetHeadDom
        Set domBody = Voucher.GetLineDom(Voucher.row)
    End If
    errMsg = clsVoucherCO.VoucherGetPrice(g_Conn, domHead, domBody)

    If errMsg <> "" Then
        MsgBox errMsg, , vbExclamation, GetString("U8.DZ.JA.Res030")
        Set domHead = Nothing
        Set domBody = Nothing
        Exit Function
    End If
    If strKey = "allprice" Then
        Voucher.SkipLoadAccessories = True
        Voucher.StopSetDefaultValue = True
        Voucher.setVoucherDataXML domHead, domBody
        Voucher.SkipLoadAccessories = False
        Voucher.StopSetDefaultValue = False
    Else
        Voucher.UpdateLineData domBody, Voucher.row
    End If
    Set domHead = Nothing
    Set domBody = Nothing
    Set clsSAWeb = Nothing

    Exit Function



Err_Handler:
    Set domHead = Nothing
    Set domBody = Nothing
    Set clsSAWeb = Nothing
    MsgBox Err.Description, vbInformation, GetString("U8.DZ.JA.Res030")

End Function

'快捷键录入批次与对应入库单号
Public Sub GetBatchInfoFun(Voucher As ctlVoucher, KeyCode As Integer, Shift As Integer)
    Dim sFree As Collection                                '存货自由项集合
    Dim sWhCode As String                                  '仓库
    Dim sInvCode As String                                 '存货编码
    Dim iSosID As String                                   '销售订单行ID
    Dim i, j As Long
    Dim curRow As Integer
    Dim sSql As String
    Dim strSql As String
    Dim sRet As String
    'Dim oRefSelect As RefSelect
    Dim oRefSelect As Object
    Dim errStr As String
    Dim oInventoryPst As InventoryPst
    Dim moInventory As USERPVO.Inventory
    Dim recRef As New ADODB.Recordset
    Dim iNum As Double                                     '调拨件数
    Dim iExchRate As Double                                '换算率
    Dim sBatch As String                                   '批次
    Dim sFreeName As String                                '自由项名称
    Dim sDemandType As String                              '销售订单类型
    Dim sDemandCode As String                              '销售订单分类号
    Dim lDemandCode As Long                                '整数型的销售订单行号
    'Dim domline As DOMDocument
    Dim r As Long                                          '记录单据表体行数
    Dim Quantity As Double                                 ' 数量
    Dim row As Long


    '取得空行的DOM
    row = Voucher.row                                      '记录当前行
    Dim domEmpty As DOMDocument
    Voucher.AddLine Voucher.BodyRows + 1
    Set domEmpty = Voucher.GetLineDom
    Voucher.DelLine Voucher.BodyRows
    Voucher.row = row                                      '恢复到当前行

    If Voucher.headerText("cwhcode") = "" Then
        MsgBox GetString("U8.DZ.JA.Res640"), vbInformation, GetString("U8.DZ.JA.Res030")
        Exit Sub
    End If

    '快捷键Ctrl+E或者Ctrl+B，自动指定批号
    If KeyCode = vbKeyE Or KeyCode = vbKeyB Then
        If Shift = vbCtrlMask Then

            If Voucher.rows > 1 Then
                'Set oRefSelect = New RefSelect
                Set oRefSelect = CreateObject("USCONTROL.RefSelect")

                oRefSelect.CreateAndDropTmpCurrentStock g_oLogin, True

                r = Voucher.BodyRows
                For i = 1 To r
                    'ctrl+B指定单行
                    If KeyCode = vbKeyB And i <> Voucher.row Then GoTo SearchNextBatch

                    '发货仓库
                    sWhCode = Voucher.headerText("cwhcode")
                    '存货编码
                    sInvCode = Voucher.bodyText(i, "cinvcode")
                    '销售订单表体行ID
                    iSosID = Voucher.bodyText(i, "isosid")

                    '件数
                    If Voucher.bodyText(i, "inum") = "" Then
                        iNum = 0
                    Else
                        iNum = CDbl(Voucher.bodyText(i, "inum"))
                    End If

                    If Voucher.bodyText(i, "iquantity") = "" Then
                        Quantity = 0
                    Else
                        Quantity = CDbl(Voucher.bodyText(i, "iquantity"))
                    End If

                    '换算率
                    If Voucher.bodyText(i, "iinvexchrate") <> "" Then
                        iExchRate = CDbl(Voucher.bodyText(i, "iinvexchrate"))
                    Else
                        iExchRate = 0
                    End If


                    '得到存货属性对象
                    Set oInventoryPst = New InventoryPst
                    oInventoryPst.login = mologin
                    oInventoryPst.Load sInvCode, moInventory

                    '对于是批次管理的存货,自动指定批号
                    If moInventory.IsBatch = True Then
                        '********************************************
                        '2008-11-17
                        '为匹配872中LP件多种销售跟踪方式的处理
                        Call GetSoDemandType(iSosID, sDemandType, sDemandCode, g_Conn)
                        If IsNumeric(sDemandCode) Then
                            lDemandCode = CLng(sDemandCode)
                        Else
                            lDemandCode = 0
                        End If
                        '********************************************

                        '自由项集合
                        Set sFree = New Collection
                        For j = 1 To 10
                            sFree.Add Null2Something(Voucher.bodyText(i, "cfree" & j))
                        Next j

                        '871和872调用库存函数时，传递的参数类型发生变化
                        '871的订单行参数要求是整数型，872的改成字符型，因此需要单独处理
                        If gU8Version = "872" Then
                            clsbill.RefBatchList sSql, "", sWhCode, sInvCode, sFree, sRet, False, CLng(sDemandType), sDemandCode, ""
                        Else
                            clsbill.RefBatchList sSql, "", sWhCode, sInvCode, sFree, sRet, False, CLng(sDemandType), lDemandCode, ""
                        End If

                        sSql = oRefSelect.GetAllBSQL(sSql)

                        errStr = "批量指定批次"

                        '871和872调用库存函数时，传递的参数类型发生变化
                        '871的订单行参数要求是整数型，872的改成字符型，因此需要单独处理
                        If gU8Version = "872" Then
                            oRefSelect.AutoRefer g_oLogin, sWhCode, sInvCode, sFree, Quantity, iNum, iExchRate, RefBatch, sSql, False, True, "12", errStr, CLng(sDemandType), sDemandCode, ""
                        Else
                            oRefSelect.AutoRefer g_oLogin, sWhCode, sInvCode, sFree, Quantity, iNum, iExchRate, RefBatch, sSql, False, True, "12", errStr, CLng(sDemandType), lDemandCode, ""
                        End If

                        If Not oRefSelect.ReturnData Is Nothing Then
                            Set recRef = oRefSelect.ReturnData
                            If recRef.RecordCount = 1 Then
                                Voucher.bodyText(i, "iquantity") = recRef.Fields("出库数量")
                                Voucher.bodyText(i, "inum") = recRef.Fields("出库件数")
                                Voucher.bodyText(i, "cbatch") = recRef.Fields("批号")
                                Voucher.bodyText(i, "dmadedate") = recRef.Fields("生产日期")
                                Voucher.bodyText(i, "dvdate") = recRef.Fields("失效日期")
                                Voucher.bodyText(i, "dexpirationdate") = recRef.Fields("有效期计算项")
                                Voucher.bodyText(i, "cexpirationdate") = recRef.Fields("有效期至")
                                
                                Voucher.bodyText(r, "imassdate") = recRef.Fields("保质期")
                                Voucher.bodyText(r, "cmassunit") = recRef.Fields("保质期单位")
                                Voucher.bodyText(r, "iexpiratdatecalcu") = recRef.Fields("有效期推算方式")
                                Voucher.bodyText(i, "cfree1") = recRef.Fields("cfree1")
                                Voucher.bodyText(i, "cfree2") = recRef.Fields("cfree2")
                                Voucher.bodyText(i, "cfree3") = recRef.Fields("cfree3")
                                Voucher.bodyText(i, "cfree4") = recRef.Fields("cfree4")
                                Voucher.bodyText(i, "cfree5") = recRef.Fields("cfree5")
                                Voucher.bodyText(i, "cfree6") = recRef.Fields("cfree6")
                                Voucher.bodyText(i, "cfree7") = recRef.Fields("cfree7")
                                Voucher.bodyText(i, "cfree8") = recRef.Fields("cfree8")
                                Voucher.bodyText(i, "cfree9") = recRef.Fields("cfree9")
                                Voucher.bodyText(i, "cfree10") = recRef.Fields("cfree10")

                                '871和872调用库存函数时，传递的参数类型发生变化
                                '871的订单行参数要求是整数型，872的改成字符型，因此需要单独处理
                                If gU8Version = "872" Then
                                    oRefSelect.UpdateTmpCurrentStock g_oLogin, IIf(IsNull(recRef.Fields("出库数量")), 0, recRef.Fields("出库数量")), IIf(IsNull(recRef.Fields("出库件数")), 0, recRef.Fields("出库件数")), oRefSelect.GetAutoID(g_oLogin, sWhCode, sInvCode, sFree(1), sFree(2), sFree(3), sFree(4), sFree(5), sFree(6), sFree(7), sFree(8), sFree(9), sFree(10), recRef.Fields("批号"), CLng(sDemandType), sDemandCode)
                                Else
                                    oRefSelect.UpdateTmpCurrentStock g_oLogin, IIf(IsNull(recRef.Fields("出库数量")), 0, recRef.Fields("出库数量")), IIf(IsNull(recRef.Fields("出库件数")), 0, recRef.Fields("出库件数")), oRefSelect.GetAutoID(g_oLogin, sWhCode, sInvCode, sFree(1), sFree(2), sFree(3), sFree(4), sFree(5), sFree(6), sFree(7), sFree(8), sFree(9), sFree(10), recRef.Fields("批号"), CLng(sDemandType), lDemandCode)
                                End If

                            ElseIf recRef.RecordCount > 1 Then    '如果批次存量不足，进行拆分行

                                While Not recRef.EOF

                                    '复制被拆行数据，拆分批次数量
                                    '                                            Voucher.AddLine Voucher.BodyRows + 1
                                    '                                            '复制当前行
                                    '                                            Set domline = Voucher.GetLineDom(i)
                                    '                                            Voucher.UpdateLineData domline, Voucher.BodyRows
                                    Voucher.DuplicatedLine i

                                    Voucher.bodyText(Voucher.BodyRows, "iquantity") = recRef.Fields("出库数量")
                                    Voucher.bodyText(Voucher.BodyRows, "inum") = recRef.Fields("出库件数")
                                    Voucher.bodyText(Voucher.BodyRows, "cbatch") = recRef.Fields("批号")
                                    Voucher.bodyText(Voucher.BodyRows, "dmadedate") = recRef.Fields("生产日期")
                                    Voucher.bodyText(Voucher.BodyRows, "dvdate") = recRef.Fields("失效日期")
                                    Voucher.bodyText(Voucher.BodyRows, "dexpirationdate") = recRef.Fields("有效期计算项")
                                    Voucher.bodyText(Voucher.BodyRows, "cexpirationdate") = recRef.Fields("有效期至")
                                    
                                    Voucher.bodyText(Voucher.BodyRows, "imassdate") = recRef.Fields("保质期")
                                    Voucher.bodyText(Voucher.BodyRows, "cmassunit") = recRef.Fields("保质期单位")
                                    Voucher.bodyText(Voucher.BodyRows, "iexpiratdatecalcu") = recRef.Fields("有效期推算方式")
                                    Voucher.bodyText(Voucher.BodyRows, "cfree1") = recRef.Fields("cfree1")
                                    Voucher.bodyText(Voucher.BodyRows, "cfree2") = recRef.Fields("cfree2")
                                    Voucher.bodyText(Voucher.BodyRows, "cfree3") = recRef.Fields("cfree3")
                                    Voucher.bodyText(Voucher.BodyRows, "cfree4") = recRef.Fields("cfree4")
                                    Voucher.bodyText(Voucher.BodyRows, "cfree5") = recRef.Fields("cfree5")
                                    Voucher.bodyText(Voucher.BodyRows, "cfree6") = recRef.Fields("cfree6")
                                    Voucher.bodyText(Voucher.BodyRows, "cfree7") = recRef.Fields("cfree7")
                                    Voucher.bodyText(Voucher.BodyRows, "cfree8") = recRef.Fields("cfree8")
                                    Voucher.bodyText(Voucher.BodyRows, "cfree9") = recRef.Fields("cfree9")
                                    Voucher.bodyText(Voucher.BodyRows, "cfree10") = recRef.Fields("cfree10")


                                    '871和872调用库存函数时，传递的参数类型发生变化
                                    '871的订单行参数要求是整数型，872的改成字符型，因此需要单独处理
                                    If gU8Version = "872" Then
                                        oRefSelect.UpdateTmpCurrentStock g_oLogin, IIf(IsNull(recRef.Fields("出库数量")), 0, recRef.Fields("出库数量")), IIf(IsNull(recRef.Fields("出库件数")), 0, recRef.Fields("出库件数")), oRefSelect.GetAutoID(g_oLogin, sWhCode, sInvCode, sFree(1), sFree(2), sFree(3), sFree(4), sFree(5), sFree(6), sFree(7), sFree(8), sFree(9), sFree(10), recRef.Fields("批号"), CLng(sDemandType), sDemandCode)
                                    Else
                                        oRefSelect.UpdateTmpCurrentStock g_oLogin, IIf(IsNull(recRef.Fields("出库数量")), 0, recRef.Fields("出库数量")), IIf(IsNull(recRef.Fields("出库件数")), 0, recRef.Fields("出库件数")), oRefSelect.GetAutoID(g_oLogin, sWhCode, sInvCode, sFree(1), sFree(2), sFree(3), sFree(4), sFree(5), sFree(6), sFree(7), sFree(8), sFree(9), sFree(10), recRef.Fields("批号"), CLng(sDemandType), lDemandCode)
                                    End If

                                    recRef.MoveNext
                                Wend

                                '删除被拆行记录，否则会出现重复记录
                                Voucher.UpdateLineData domEmpty, CLng(i)

                            End If
                        End If


                        If errStr <> "" Then
                            errStr = GetString("U8.DZ.JA.Res780") & sInvCode & errStr & vbCrLf
                        End If

                    End If                                 '匹配moInventory.IsBatch = True的End If
SearchNextBatch:
                Next i

                Voucher.RemoveEmptyRow
                If errStr <> "" Then
                    MsgBox errStr, vbCritical, GetString("U8.DZ.JA.Res030")
                End If

                '删除临时表
                oRefSelect.CreateAndDropTmpCurrentStock g_oLogin, False
                Set oRefSelect = Nothing
            End If
        End If                                             '匹配 Shift = vbCtrlMask 的End If

        '自动指定跟踪型存货入库单号
    ElseIf KeyCode = vbKeyQ Or KeyCode = vbKeyO Then

        If Shift = vbCtrlMask And Voucher.rows > 1 Then

            Set oRefSelect = CreateObject("USCONTROL.RefSelect")

            For i = 1 To Voucher.rows - 1
                'ctrl+Q指定单行
                If KeyCode = vbKeyQ And i <> Voucher.row Then GoTo SearchNextInVouchCode:
                '发货仓库
                sWhCode = Voucher.headerText("cwhcode")
                '存货编码
                sInvCode = Voucher.bodyText(i, "cinvcode")
                '销售订单表体行ID
                iSosID = Voucher.bodyText(i, "isosid")

                '调拨件数
                If Voucher.bodyText(i, "inum") = "" Then
                    iNum = 0
                Else
                    iNum = CDbl(Voucher.bodyText(i, "inum"))
                End If

                If Voucher.bodyText(i, "iquantity") = "" Then
                    Quantity = 0
                Else
                    Quantity = CDbl(Voucher.bodyText(i, "iquantity"))
                End If
                '换算率
                If Voucher.bodyText(i, "iinvexchrate") <> "" Then
                    iExchRate = CDbl(Voucher.bodyText(i, "iinvexchrate"))
                Else
                    iExchRate = 0
                End If

                '得到存货属性对象
                Set oInventoryPst = New InventoryPst
                oInventoryPst.login = mologin
                oInventoryPst.Load sInvCode, moInventory

                '对于是跟踪型存货,自动指定入库单号
                If moInventory.IsTrack = True Then
                    '********************************************
                    '2008-11-17
                    '为匹配872中LP件多种销售跟踪方式的处理
                    Call GetSoDemandType(iSosID, sDemandType, sDemandCode, g_Conn)
                    If IsNumeric(sDemandCode) Then
                        lDemandCode = CLng(sDemandCode)
                    Else
                        lDemandCode = 0
                    End If
                    '********************************************


                    '自由项集合
                    Set sFree = New Collection
                    For j = 1 To 10
                        sFree.Add Null2Something(Voucher.bodyText(i, "cfree" & j))
                    Next j

                    '                        ClsBill.RefInVouchList sSql, "", sWhCode, sInvCode, sFree, sRet, True, voucher.bodytext(I, "cbatch")), 0, 0, ""
                    '                        oRefSelect.AutoRefer g_oLogin, sWhCode, sInvCode, sFree, voucher.bodytext(I, "iquantity")), inum, iExchRate, RefInVouch, sSql, False, True, "12", errStr, 0, 0, ""

                    '871和872调用库存函数时，传递的参数类型发生变化
                    '871的订单行参数要求是整数型，872的改成字符型，因此需要单独处理
                    If gU8Version = "872" Then
                        clsbill.RefInVouchList sSql, "", sWhCode, sInvCode, sFree, sRet, True, Voucher.bodyText(i, "cbatch"), CLng(sDemandType), sDemandCode, ""
                        oRefSelect.AutoRefer g_oLogin, sWhCode, sInvCode, sFree, Voucher.bodyText(i, "iquantity"), iNum, iExchRate, RefInVouch, sSql, False, True, "12", errStr, CLng(sDemandType), sDemandCode, ""
                    Else
                        clsbill.RefInVouchList sSql, "", sWhCode, sInvCode, sFree, sRet, True, Voucher.bodyText(i, "cbatch"), CLng(sDemandType), lDemandCode, ""
                        oRefSelect.AutoRefer g_oLogin, sWhCode, sInvCode, sFree, Voucher.bodyText(i, "iquantity"), iNum, iExchRate, RefInVouch, sSql, False, True, "12", errStr, CLng(sDemandType), lDemandCode, ""
                    End If

                    If Not oRefSelect.ReturnData Is Nothing Then
                        Set recRef = oRefSelect.ReturnData
                        If Not IsNull(Voucher.bodyText(i, "cinvouchcode")) Then
                            If recRef.RecordCount = 1 Then
                                Voucher.bodyText(i, "iquantity") = Null2Something(recRef.Fields("出库数量"))
                                Voucher.bodyText(i, "inum") = Null2Something(recRef.Fields("出库件数"))
                                If Not IsNull(Voucher.bodyText(i, "cbatch")) Then
                                    Voucher.bodyText(i, "cbatch") = Null2Something(recRef.Fields("批号"))
                                End If

                                '带出自由项
                                For j = 0 To recRef.Fields.Count - 1
                                    sFreeName = IIf(IsNull(recRef.Fields(j).Properties("BASECOLUMNNAME")), "", recRef.Fields(j).Properties("BASECOLUMNNAME"))
                                    If LCase(Mid(sFreeName, 1, 5)) = "cfree" Then
                                        If Not IsNull(Voucher.bodyText(i, sFreeName)) > 0 Then
                                            Voucher.bodyText(i, sFreeName) = Null2Something(recRef.Fields(i).Value)
                                        End If
                                    End If
                                Next

                                If Not IsNull(Voucher.bodyText(i, "cinvouchcode")) > 0 Then
                                    Voucher.bodyText(i, "cinvouchcode") = Null2Something(recRef.Fields("入库单号"))
                                End If

                                '返回入库Id
                                If Not IsNull(Voucher.bodyText(i, "rdsid")) Then
                                    Voucher.bodyText(i, "rdsid") = Null2Something(recRef.Fields("入库系统编号"))
                                End If


                                '更新行记录的入库单号、入库系统编号、批次、自由项信息、选中标记、生单标记

                                Voucher.bodyText(i, "iquantity") = recRef.Fields("出库数量")
                                Voucher.bodyText(i, "inum") = recRef.Fields("出库件数")
                                Voucher.bodyText(i, "cbatch") = recRef.Fields("批号")
                                Voucher.bodyText(i, "cinvouchcode") = recRef.Fields("入库单号")
                                Voucher.bodyText(i, "rdsid") = recRef.Fields("入库系统编号")
                                Voucher.bodyText(i, "cfree1") = recRef.Fields("cfree1")
                                Voucher.bodyText(i, "cfree2") = recRef.Fields("cfree2")
                                Voucher.bodyText(i, "cfree3") = recRef.Fields("cfree3")
                                Voucher.bodyText(i, "cfree4") = recRef.Fields("cfree4")
                                Voucher.bodyText(i, "cfree5") = recRef.Fields("cfree5")
                                Voucher.bodyText(i, "cfree6") = recRef.Fields("cfree6")
                                Voucher.bodyText(i, "cfree7") = recRef.Fields("cfree7")
                                Voucher.bodyText(i, "cfree8") = recRef.Fields("cfree8")
                                Voucher.bodyText(i, "cfree9") = recRef.Fields("cfree9")
                                Voucher.bodyText(i, "cfree10") = recRef.Fields("cfree10")
                                Voucher.bodyText(i, "dmadedate") = recRef.Fields("生产日期")
                                Voucher.bodyText(i, "dvdate") = recRef.Fields("失效日期")
                                Voucher.bodyText(i, "dexpirationdate") = recRef.Fields("有效期计算项")
                                Voucher.bodyText(i, "cexpirationdate") = recRef.Fields("有效期至")
                                
                                Voucher.bodyText(i, "imassdate") = recRef.Fields("保质期")
                                Voucher.bodyText(i, "cmassunit") = recRef.Fields("保质期单位")
                                Voucher.bodyText(i, "iexpiratdatecalcu") = recRef.Fields("有效期推算方式")



                            ElseIf recRef.RecordCount > 1 Then    '如果批次存量不足，进行拆分行

                                While Not recRef.EOF

                                    '复制被拆行数据，拆分入库数量
                                    '                                        Voucher.AddLine Voucher.BodyRows + 1
                                    '                                            '复制当前行
                                    '                                        Set domline = Voucher.GetLineDom(i)
                                    '                                        Voucher.UpdateLineData domline, Voucher.BodyRows
                                    Voucher.DuplicatedLine i

                                    Voucher.bodyText(Voucher.BodyRows, "iquantity") = recRef.Fields("出库数量")
                                    Voucher.bodyText(Voucher.BodyRows, "inum") = recRef.Fields("出库件数")
                                    Voucher.bodyText(Voucher.BodyRows, "cbatch") = recRef.Fields("批号")
                                    Voucher.bodyText(Voucher.BodyRows, "cinvouchcode") = recRef.Fields("入库单号")
                                    Voucher.bodyText(Voucher.BodyRows, "rdsid") = recRef.Fields("入库系统编号")
                                    Voucher.bodyText(Voucher.BodyRows, "cfree1") = recRef.Fields("cfree1")
                                    Voucher.bodyText(Voucher.BodyRows, "cfree2") = recRef.Fields("cfree2")
                                    Voucher.bodyText(Voucher.BodyRows, "cfree3") = recRef.Fields("cfree3")
                                    Voucher.bodyText(Voucher.BodyRows, "cfree4") = recRef.Fields("cfree4")
                                    Voucher.bodyText(Voucher.BodyRows, "cfree5") = recRef.Fields("cfree5")
                                    Voucher.bodyText(Voucher.BodyRows, "cfree6") = recRef.Fields("cfree6")
                                    Voucher.bodyText(Voucher.BodyRows, "cfree7") = recRef.Fields("cfree7")
                                    Voucher.bodyText(Voucher.BodyRows, "cfree8") = recRef.Fields("cfree8")
                                    Voucher.bodyText(Voucher.BodyRows, "cfree9") = recRef.Fields("cfree9")
                                    Voucher.bodyText(Voucher.BodyRows, "cfree10") = recRef.Fields("cfree10")
                                    Voucher.bodyText(Voucher.BodyRows, "dmadedate") = recRef.Fields("生产日期")
                                    Voucher.bodyText(Voucher.BodyRows, "dvdate") = recRef.Fields("失效日期")
                                    Voucher.bodyText(Voucher.BodyRows, "dexpirationdate") = recRef.Fields("有效期计算项")
                                    Voucher.bodyText(Voucher.BodyRows, "cexpirationdate") = recRef.Fields("有效期至")
                                    
                                    Voucher.bodyText(Voucher.BodyRows, "imassdate") = recRef.Fields("保质期")
                                    Voucher.bodyText(Voucher.BodyRows, "cmassunit") = recRef.Fields("保质期单位")
                                    Voucher.bodyText(Voucher.BodyRows, "iexpiratdatecalcu") = recRef.Fields("有效期推算方式")

                                    recRef.MoveNext
                                Wend

                                '删除被拆行记录，否则会出现重复记录
                                Voucher.UpdateLineData domEmpty, CLng(i)

                            End If
                        End If
                    End If

                    If errStr <> "" Then
                        errStr = GetString("U8.DZ.JA.Res780") & sInvCode & errStr & vbCrLf
                    End If

                End If
SearchNextInVouchCode:
            Next i

            Voucher.RemoveEmptyRow                         '清除空行
            If errStr <> "" Then
                MsgBox errStr, vbInformation, GetString("U8.DZ.JA.Res030")
            End If

            Set oRefSelect = Nothing

        End If                                             '匹配If Shift = vbCtrlMask And voucher.Rows > 1 Then
    End If                                                 '匹配 KeyCode = vbKeyE 的End If

End Sub


'查审
Public Sub ExecViewVerify(guid As String)
    SendShowViewMessage guid, "UFIDA.U8.Audit.AuditHistoryView"
    SendMessgeToPortal "DocQueryAuditHistory", guid
End Sub

'重新提交以及启用工作流审批时的操作
Public Sub ExecRequestAudit(guid As String)
    SendShowViewMessage guid, "UFIDA.U8.Audit.AuditViews.TreatTaskViewPart"
    SendMessgeToPortal "DocRequestAudit", guid
End Sub

'启用工作流撤销时的操作
Public Sub ExecCancelAudit(guid As String)
    SendShowViewMessage guid, "UFIDA.U8.Audit.AuditViews.TreatTaskViewPart"
    SendMessgeToPortal "DocRequestCancelAudit", guid
End Sub


''向平台发消息
'Public Sub SendMessgeToPortal(strMessageType As String, guid As String)
'On Error GoTo errhandle
'
'
'    Call CheckSubmit(MainTable, "ID", CStr(lngVoucherID))
'    Dim strMaker As String
'
'    strMaker = FrmVoucher.Voucher.headerText("cmaker")
'    If strMaker = "" Then strMaker = g_oLogin.cUserName
'
'    SendPortalMessage guid, gstrCardNumber, CStr(lngVoucherID), strMessageType, strMaker, ufts, vouchercode, "SAM030204", "SAM030205" 'AuthVerify, AuthUnVerify
'
'    Exit Sub
'
'errhandle:
'
'End Sub

'向平台发消息
Public Sub SendMessgeToPortal(strMessageType As String, guid As String)
    On Error GoTo ErrHandle


    Call CheckSubmit(MainTable, "ID", CStr(lngVoucherID))
    Dim strMaker As String

    strMaker = FrmVoucher.Voucher.headerText("cmaker")
    If strMaker = "" Then strMaker = g_oLogin.cUserName

    SendPortalMessage guid, gstrCardNumber, CStr(lngVoucherID), strMessageType, strMaker, ufts, _
            vouchercode, "ST02JC020105", "ST02JC020106"    'AuthVerify, AuthUnVerify

    Exit Sub

ErrHandle:

End Sub

'leix begin
'单据助手
Public Sub SendPortalMessage(strFormGuid As String, strCardNumber As String, strID As String, _
                             Optional strMessageType As String = "CurrentDocChanged", _
                             Optional strMaker As String = "", Optional ufts As String = "", Optional vouchercode As String = "", _
                             Optional strAuditAuthId As String = "", Optional strAbandonAuthId As String = "")
    Dim tsb As Object
    Dim strXML As String

    If Not (g_oBusiness Is Nothing) Then
        Set tsb = g_oBusiness.GetToolbarSubjectEx(strFormGuid)
    End If
    strXML = "<?xml version='1.0' encoding='UTF-8'?>"
    strXML = strXML & "<Message type='" & strMessageType & "'>"
    strXML = strXML & "<Selection context='BO:" + UCase(strCardNumber) + "'>"
    strXML = strXML & "<Element typeName='Voucher' cVoucherId='" & strID & "' cMaker='" & strMaker & "' cCardNum='" & UCase(strCardNumber) & "' cVoucherCode='" & vouchercode & "' Ufts='" & ufts & "' AuditAuthId='" & strAuditAuthId & "' AbandonAuthId='" & strAbandonAuthId & "'/>"
    strXML = strXML & "</Selection>"
    strXML = strXML & "</Message>"
    If Not (tsb Is Nothing) Then
        Call tsb.TransMessage(strFormGuid, strXML)
    End If

    Set tsb = Nothing

End Sub

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

Public Sub SendShowViewMessage(guid As String, sViewID As String, Optional ByVal strMessageType As String = "SHOWVIEW")
    'sViewID:="UFIDA.U8.Audit.AuditViews.TreatTaskViewPart",审批视图,
    'sViewID:="UFIDA.U8.Audit.AuditHistoryView",审批进程表查,审时调用
    'SHOWVIEW显示视图，HIDEVIEW隐藏视图
    Dim tsb As Object
    Dim strXML As String
    If Not (g_oBusiness Is Nothing) Then
        Set tsb = g_oBusiness.GetToolbarSubjectEx(guid)
    End If
    strXML = ""
    strXML = strXML & "<Message type='" & strMessageType & "'>"
    strXML = strXML & "   <Selection context='BO:" + UCase(gstrCardNumber) + "'>"
    strXML = strXML & "      <Element typeName = 'ViewPart' viewID = '" & sViewID & "'  isFirstElement = 'true'/> "
    strXML = strXML & "   </Selection>"
    strXML = strXML & "</Message>"
    If Not (tsb Is Nothing) Then
        Call tsb.TransMessage(guid, strXML)
    End If

    If Not tsb Is Nothing Then Set tsb = Nothing

End Sub

'870 added 判断是否启用工作流
Public Function getIsWfControl(login As clsLogin, myConn As ADODB.Connection, ByRef errMsg As String, cardnumber As String) As Boolean
    Dim clsisWfCtl As Object
    Set clsisWfCtl = CreateObject("SCMWorkFlowCommon.clsWFController")
    Dim isWfCtl As Boolean
    Call clsisWfCtl.GetIsWFControlled(myConn, cardnumber, cardnumber & ".Submit", login.cIYear, login.cAcc_Id, isWfCtl, errMsg)
    getIsWfControl = isWfCtl
End Function

'12.0 added 判断是否激活过工作流
Public Function getIsWFHasActivated(login As clsLogin, myConn As ADODB.Connection, ByRef errMsg As String, cardnumber As String) As Boolean
    Dim clsisWfCtl As Object
    Set clsisWfCtl = CreateObject("SCMWorkFlowCommon.clsWFController")
    Dim isWfCtl As Boolean
    'Call clsisWfCtl.GetIsWFControlled(myConn, cardnumber, cardnumber & ".Submit", login.cIYear, login.cAcc_Id, isWfCtl, errMsg)
    Call clsisWfCtl.getIsWFHasActivated(myConn, cardnumber, cardnumber & ".Submit", isWfCtl, errMsg)
    getIsWFHasActivated = isWfCtl
End Function

'设置工作流相关按钮
Public Sub SetWFControlBrns(login As clsLogin, myConn As ADODB.Connection, Toolbar As Object, UFToolbar As Object, cardnumber As String)
    Dim rstfilter As String
    If getIsWFHasActivated(login, g_Conn, rstfilter, cardnumber) Then
        Toolbar.Buttons(sKey_Submit).Visible = True
        Toolbar.Buttons(sKey_Resubmit).Visible = True
        Toolbar.Buttons(sKey_Unsubmit).Visible = True
        Toolbar.Buttons(sKey_ViewVerify).Visible = True
    Else
        Toolbar.Buttons(sKey_Submit).Enabled = False
        Toolbar.Buttons(sKey_Resubmit).Enabled = False
        Toolbar.Buttons(sKey_Unsubmit).Enabled = False
        Toolbar.Buttons(sKey_ViewVerify).Enabled = False
    End If

    Toolbar.Buttons(sKey_CreateVoucher).Visible = False
    Toolbar.Buttons(sKey_Fetchprice).Visible = False
    Toolbar.Buttons(sKey_ReferVoucher).Visible = False
   ' Toolbar.Buttons(sKey_Acc).Visible = False
    '    Toolbar.Buttons(sKey_Open).Visible = False
    '    Toolbar.Buttons(sKey_Close).Visible = False
    Toolbar.Buttons(sKey_Addrecord).Visible = False
    '    Toolbar.Buttons(sKey_Copy).Visible = False

    'UFToolbar.RefreshVisible
    'UFToolbar.RefreshEnable
End Sub

'审批流提交与撤销
Public Sub ExecSubmit(DoOrUndo As Boolean, table As String, pk As String, id As Long)
    Dim retDoUndoSubmit As Boolean
    Dim strErrorResId As String                            '审批流的错误信息870 added

    Screen.MousePointer = vbHourglass
    Call CheckSubmit(table, pk, CStr(id))

    If CBool(IsWFControlled) And ((DoOrUndo And (iverifystate = 0 Or (iverifystate = 1 And ireturncount > 0))) Or (DoOrUndo = False And iverifystate <> 0)) Then

        retDoUndoSubmit = DoUndoSubmit(DoOrUndo, gstrCardNumber, CStr(id), table, ufts, CBool(IsWFControlled), strErrorResId, vouchercode)
        If retDoUndoSubmit = False Then
            MsgBox strErrorResId, vbInformation, GetString("U8.DZ.JA.Res030")
        Else
            If DoOrUndo Then
                MsgBox GetString("U8.SA.xsglsql_2.saworkflowsrv.011"), vbInformation, GetString("U8.DZ.JA.Res030")    '"单据提交成功！"
            Else
                MsgBox GetString("U8.SA.xsglsql_2.saworkflowsrv.012"), vbInformation, GetString("U8.DZ.JA.Res030")    '撤销成功！
            End If
        End If
    Else
        If DoOrUndo Then
            MsgBox GetString("U8.SA.xsglsql_2.saworkflowsrv.001"), vbInformation, GetString("U8.DZ.JA.Res030")    '"该单据已经提交或者未启用审批流！"
        Else
            MsgBox GetString("U8.SA.xsglsql_2.saworkflowsrv.002"), vbInformation, GetString("U8.DZ.JA.Res030")    '"该单据已经撤销或者未启用审批流！"
        End If
    End If
    Screen.MousePointer = vbDefault
End Sub

'提交和撤销的代码
Public Function DoUndoSubmit(m_Handle As Boolean, m_CardNumber As String, m_Mid As String, m_TablName As String, m_ufts As String, IsWFControlled As Boolean, strErr As String, Optional cVoucherCode As String, Optional CbillType As String = "") As Boolean
    On Error GoTo ErrHandler
    Dim objCalledContext As Object
    Set objCalledContext = CreateObject("UFSoft.U8.Framework.LoginContext.CalledContext")
    objCalledContext.SubId = g_oLogin.cSub_Id
    objCalledContext.TaskId = g_oLogin.TaskId
    objCalledContext.token = g_oLogin.userToken

    '    Dim clsSub As SZDZ_dxb_WorkFlowSrv.clsWorkFlowSrv
    '    Set clsSub = New SZDZ_dxb_WorkFlowSrv.clsWorkFlowSrv
    Dim clsSub As Object
    Set clsSub = CreateObject("SZDZ_dxb_WorkFlowSrv.clsWorkFlowSrv")

    Dim context As String
    Dim obj As Object
    Set obj = CreateObject("UFLTMService.clsService")
    obj.Start g_Conn.ConnectionString
    obj.BeginTransaction



    If m_Handle Then
        DoUndoSubmit = clsSub.DoSubmit(m_CardNumber, m_CardNumber & ".Submit", m_Mid, context, objCalledContext, m_ufts, IsWFControlled, strErr, g_oLogin, CbillType)
    Else
        DoUndoSubmit = clsSub.UndoSubmit(m_CardNumber, m_CardNumber & ".Submit", m_Mid, m_CardNumber, objCalledContext, m_ufts, IsWFControlled, strErr, cVoucherCode, g_oLogin)
    End If
    If DoUndoSubmit Then
        obj.Commit
    Else
        obj.Rollback
    End If
    obj.Finish
    Set obj = Nothing
    Exit Function
ErrHandler:
    strErr = VBA.Err.Description
    DoUndoSubmit = False
End Function

Private Function GetAuditSrvObj() As Object
    Dim obj As Object
    On Error GoTo ErrHandle
    Set obj = CreateObject("UFLTMService.clsService")
    Set GetAuditSrvObj = obj
    Exit Function
ErrHandle:
    Set GetAuditSrvObj = Nothing
End Function

''参照单据
Public Function ReferVouch() As Boolean

    Dim frm As New frmVouchRefers
    Dim clsReferVoucher As New clsReferVoucher

    ' 设置参照生单控件的属性
    clsReferVoucher.HelpID = "0"                           '帮助 10151180
    clsReferVoucher.pageSize = 20                          '默认分页大小
    clsReferVoucher.strMainKey = "ID"                      '主表唯一主键，作为和子表关联的依据
    clsReferVoucher.strDetailKey = " "                '子表唯一主键
    clsReferVoucher.FrmCaption = "计划参照承包合同"
    clsReferVoucher.FilterKey = "计划参照承包合同"                '"借出借用单参照"                  '过滤器名称 SA26
    clsReferVoucher.FilterSubID = "ST"
isfyflg = False

    clsReferVoucher.HeadKey = "FYSL0035"               '主表的列信息 AA_ColumnDic
    clsReferVoucher.BodyKey = ""               '子表的列信息 ,若设置只有表头时该属性置空
    '添加自定义按钮
    'clsReferVoucher.Buttons = "<root><row type='toolbar' name='tlbTest' caption='自动匹配' index='26' /></root>"
    '多语时
    'strButtons ="<root><row type='toolbar' name='tlbTest' caption='{res:U8.SA.xsglsql_2.frmvouchrefers.automatching}' index='26' /></root>"
    'strButtons =ReplaceResId("<root><row type='toolbar' name='tlbTest' caption='{res:U8.SA.xsglsql_2.frmvouchrefers.automatching}' index='26' /></root>")
    clsReferVoucher.MainDataSource = "V_HY_FYSL_Payment_refer"    '主表数据源视图
    clsReferVoucher.DetailDataSource = " "    '子表数据源试图
    clsReferVoucher.DefaultFilter = ""                     '"IsNULL(cVerifier,'')<>'' and IsNULL(cCloser,'')=''And IsNull(cSCloser,'')=''"           '默认过滤条件

    clsReferVoucher.strAuth = Replace(sMakeAuth_ALL, "HY_FYSL_Payment.id", "V_HY_FYSL_Payment_refer.id")    '数据权限SQL串, 这里不支持仓库权限控制

    clsReferVoucher.OtherFilter = ""                       '其他过滤条件

    clsReferVoucher.HeadEnabled = False                    '主表是否可编辑
    clsReferVoucher.BodyEnabled = False                    '子表是否可编辑

    'clsReferVoucher.bSelectSingle = True                                           '表头是否只能取唯一记录

    clsReferVoucher.bSelectSingle = False                  '表头是否只能取唯一记录
    clsReferVoucher.strCheckFlds = "ccode"                 '"cType,bObjectCode"    '
    clsReferVoucher.strCheckMsg = "选择了不同的合同编号"
    ' clsReferVoucher.strBodyCellCheckFields = "cwhcode"





    Set frm.clsReferVoucher = clsReferVoucher
    If frm.OpenFilter Then
        frm.Show vbModal
    Else
        frm.bcancel = True
    End If
    Set clsReferVoucher = Nothing


    If Not frm.bcancel Then
        ReferVouch = True
        'Set Domhead = frmVouchRef.Domhead
        'Set Dombody = frmVouchRef.Dombody
        Set gDomReferHead = frm.domHead
'        Set gDomReferBody = Frm.domBody
    End If

    Unload frm
    Set frm = Nothing

End Function
''参照单据
Public Function ReferVouchpro() As Boolean

    Dim frm As New frmVouchRefers
    Dim clsReferVoucher As New clsReferVoucher

    ' 设置参照生单控件的属性
    clsReferVoucher.HelpID = "0"                           '帮助 10151180
    clsReferVoucher.pageSize = 20                          '默认分页大小
    clsReferVoucher.strMainKey = "ID"                      '主表唯一主键，作为和子表关联的依据
    clsReferVoucher.strDetailKey = " "             '子表唯一主键
    clsReferVoucher.FrmCaption = "合同参照项目发布 "
    clsReferVoucher.FilterKey = "合同参照项目发布"                '"借出借用单参照"                  '过滤器名称 SA26
    clsReferVoucher.FilterSubID = "ST"


    clsReferVoucher.HeadKey = "FYSL0009"               '主表的列信息 AA_ColumnDic
    clsReferVoucher.BodyKey = ""               '子表的列信息 ,若设置只有表头时该属性置空
    '添加自定义按钮
    'clsReferVoucher.Buttons = "<root><row type='toolbar' name='tlbTest' caption='自动匹配' index='26' /></root>"
    '多语时
    'strButtons ="<root><row type='toolbar' name='tlbTest' caption='{res:U8.SA.xsglsql_2.frmvouchrefers.automatching}' index='26' /></root>"
    'strButtons =ReplaceResId("<root><row type='toolbar' name='tlbTest' caption='{res:U8.SA.xsglsql_2.frmvouchrefers.automatching}' index='26' /></root>")
    clsReferVoucher.MainDataSource = "V_HY_FYSL_Contract_refer3"    '主表数据源视图
    clsReferVoucher.DetailDataSource = " "    '子表数据源试图
    clsReferVoucher.DefaultFilter = ""                     '"IsNULL(cVerifier,'')<>'' and IsNULL(cCloser,'')=''And IsNull(cSCloser,'')=''"           '默认过滤条件

    clsReferVoucher.strAuth = Replace(sMakeAuth_ALL, "HY_FYSL_Contract.id", "V_HY_FYSL_Contract_refer3.id")    '数据权限SQL串, 这里不支持仓库权限控制

    clsReferVoucher.OtherFilter = ""                       '其他过滤条件

    clsReferVoucher.HeadEnabled = False                    '主表是否可编辑
    clsReferVoucher.BodyEnabled = False                    '子表是否可编辑

    'clsReferVoucher.bSelectSingle = True                                           '表头是否只能取唯一记录

    clsReferVoucher.bSelectSingle = False                  '表头是否只能取唯一记录
    clsReferVoucher.strCheckFlds = "ccode"                 '"cType,bObjectCode"    '
    clsReferVoucher.strCheckMsg = "选择了不同的项目编号"
    ' clsReferVoucher.strBodyCellCheckFields = "cwhcode"





    Set frm.clsReferVoucher = clsReferVoucher
    If frm.OpenFilter Then
        frm.Show vbModal
    Else
        frm.bcancel = True
    End If
    Set clsReferVoucher = Nothing


    If Not frm.bcancel Then
        ReferVouchpro = True
        'Set Domhead = frmVouchRef.Domhead
        'Set Dombody = frmVouchRef.Dombody
        Set gDomReferHead = frm.domHead
'        Set gDomReferBody = Frm.domBody
    End If

    Unload frm
    Set frm = Nothing

End Function

''参照单据
Public Function ReferVoucheng() As Boolean

    Dim frm As New frmVouchRefers
    Dim clsReferVoucher As New clsReferVoucher

    ' 设置参照生单控件的属性
    clsReferVoucher.HelpID = "0"                           '帮助 10151180
    clsReferVoucher.pageSize = 20                          '默认分页大小
    clsReferVoucher.strMainKey = "ID"                      '主表唯一主键，作为和子表关联的依据
    clsReferVoucher.strDetailKey = " "                '子表唯一主键
    clsReferVoucher.FrmCaption = "承包合同参照"
    clsReferVoucher.FilterKey = "承包合同参照"                '"借出借用单参照"                  '过滤器名称 SA26
    clsReferVoucher.FilterSubID = "ST"
    
    isfyflg = True
    
    clsReferVoucher.HeadKey = "FYSL0035"               '主表的列信息 AA_ColumnDic
    clsReferVoucher.BodyKey = ""               '子表的列信息 ,若设置只有表头时该属性置空
    '添加自定义按钮
    'clsReferVoucher.Buttons = "<root><row type='toolbar' name='tlbTest' caption='自动匹配' index='26' /></root>"
    '多语时
    'strButtons ="<root><row type='toolbar' name='tlbTest' caption='{res:U8.SA.xsglsql_2.frmvouchrefers.automatching}' index='26' /></root>"
    'strButtons =ReplaceResId("<root><row type='toolbar' name='tlbTest' caption='{res:U8.SA.xsglsql_2.frmvouchrefers.automatching}' index='26' /></root>")
    clsReferVoucher.MainDataSource = "V_HY_FYSL_Payment_refer1"    '主表数据源视图
    clsReferVoucher.DetailDataSource = " "    '子表数据源试图
    clsReferVoucher.DefaultFilter = ""                     '"IsNULL(cVerifier,'')<>'' and IsNULL(cCloser,'')=''And IsNull(cSCloser,'')=''"           '默认过滤条件

    clsReferVoucher.strAuth = Replace(sMakeAuth_ALL, "HY_FYSL_Payment.id", "V_HY_FYSL_Payment_refer1.id")    '数据权限SQL串, 这里不支持仓库权限控制

    clsReferVoucher.OtherFilter = ""                       '其他过滤条件

    clsReferVoucher.HeadEnabled = False                    '主表是否可编辑
    clsReferVoucher.BodyEnabled = False                    '子表是否可编辑

    'clsReferVoucher.bSelectSingle = True                                           '表头是否只能取唯一记录

    clsReferVoucher.bSelectSingle = False                  '表头是否只能取唯一记录
    clsReferVoucher.strCheckFlds = "ccode"                 '"cType,bObjectCode"    '
    clsReferVoucher.strCheckMsg = "选择了不同的合同编号"
    ' clsReferVoucher.strBodyCellCheckFields = "cwhcode"





    Set frm.clsReferVoucher = clsReferVoucher
    If frm.OpenFilter Then
        frm.Show vbModal
    Else
        frm.bcancel = True
    End If
    Set clsReferVoucher = Nothing


    If Not frm.bcancel Then
        ReferVoucheng = True
        'Set Domhead = frmVouchRef.Domhead
        'Set Dombody = frmVouchRef.Dombody
        Set gDomReferHead = frm.domHead
'        Set gDomReferBody = Frm.domBody
    End If

    Unload frm
    Set frm = Nothing

End Function

'生单操作-组织Dom
Public Function ExecmakeDom(oDomHead As DOMDocument, oDomBody As DOMDocument, conn As Object) As Boolean

    On Error GoTo ErrHandler:

    Dim strSel As String, strSql As String
    Dim eleList As IXMLDOMNodeList
    Dim ele As IXMLDOMElement
    Dim view As String
    Dim rs As New ADODB.Recordset
    Dim errMsg As String

    view = GetViewBody(conn, gstrCardNumberlist)

    Set eleList = oDomHead.selectNodes("//z:row")
    'enum by modify
    For Each ele In eleList

        If GetNodeAtrVal(ele, "iStatus") = "审核" Then
            strSel = strSel + "," + GetNodeAtrVal(ele, "ID")
        Else
            ReDim varArgs(1)
            varArgs(0) = GetNodeAtrVal(ele, "cCODE")
            varArgs(1) = GetNodeAtrVal(ele, "iStatus")
            errMsg = errMsg & GetStringPara("U8.DZ.JA.Res800", varArgs(0), varArgs(1)) & vbCrLf
            '            errMsg = errMsg & "单据" & GetNodeAtrVal(ele, "cCODE") & "当前状态为" & GetNodeAtrVal(ele, "iStatus") & ",不能推单！" & vbCrLf
        End If
    Next

    If errMsg <> "" Then
        MsgBox errMsg, vbInformation, GetString("U8.DZ.JA.Res030")
    End If

    If strSel = "" And errMsg = "" Then
        MsgBox GetString("U8.DZ.JA.Res810"), vbInformation, GetString("U8.DZ.JA.Res030")
        ExecmakeDom = False
        Exit Function
    End If

    strSql = "select *,'' as editprop from " & view & " where id in (-1" & strSel & ")"
    rs.Open strSql, conn, adOpenDynamic, adLockOptimistic
    rs.Save oDomBody, adPersistXML                         '得到入库单表头DOM结构对象
    rs.Close
    ExecmakeDom = True
    Set rs = Nothing
    Exit Function
ErrHandler:
    ExecmakeDom = False
    MsgBox Err.Description, vbInformation, GetString("U8.DZ.JA.Res030")
    Set rs = Nothing
End Function

'判断已生单标志
Public Function SdFlg(ByVal idStr As String, ByVal uftsStr As String) As String
    Dim strSql As String
    Dim oRs As New ADODB.Recordset
    'by lg081106 修改
    strSql = "select * from " & MainTable & " where " & HeadPKFld & " = '" & idStr & "' and isnull(DownstreamCode,'') ='' and convert(nchar, convert(money,ufts), 2) = '" & uftsStr & "'"
    oRs.Open strSql, g_Conn, adOpenForwardOnly, adLockReadOnly
    If Not oRs.EOF Then
        SdFlg = ""
    Else
        SdFlg = GetString("U8.DZ.JA.Res820")
    End If
    If oRs.State = 1 Then oRs.Close
    Set oRs = Nothing
End Function

'推单  组织Dom,调用对应的生单方法
Public Function ExecCreateVoucher(ByRef oDomHead As DOMDocument, ByRef oDomBody As DOMDocument, conn As Object, login As clsLogin, bill As BillType, Optional VoucherType As String) As Boolean
    Select Case bill
            'enum by modify
        Case 销售
            ExecCreateVoucher = WriteSABill(oDomHead, oDomBody, conn, login, VoucherType)
        Case 采购
            ExecCreateVoucher = WritePUBill(oDomHead, oDomBody, conn, login, VoucherType, "88")
        Case 库存
            ExecCreateVoucher = WriteSCBill(oDomHead, oDomBody, conn, login, VoucherType, "0301")
        Case 应付
            ExecCreateVoucher = WriteAPBill(oDomHead, oDomBody, conn, login, "AP", "AP04")    '应收对应AR，AR04
    End Select
End Function

'参照生单处理
Public Function ProcessData(Voucher As Object, Optional Form As Object = Null)

    On Error GoTo ErrHandler:

    Dim retvalue  As Variant

    Dim referpara As UAPVoucherControl85.ReferParameter

    Dim eleline   As IXMLDOMElement

    Dim echeck    As Long

    Dim i         As Integer

    Dim rstTmp    As New Recordset

    '    Dim oVoucherBo As New USERPBO.VoucherBO
    
    'Set oVoucherBo.login = g_oLogin

    '提供两种处理模式，表头使用recordset处理，表体使用xml解析处理
    Dim rshead    As New ADODB.Recordset

    rshead.Open gDomReferHead

    If Not rshead.EOF And Not rshead.BOF Then
        If rshead("selcol") = "Y" Then

            
            Voucher.headerText("mconcode") = Null2Something(rshead("mconcode"))    '
            Voucher.headerText("mconname") = Null2Something(rshead("mconname"))
             Voucher.headerText("concode") = Null2Something(rshead("ccode"))
              Voucher.headerText("conname") = Null2Something(rshead("cname"))
            
            Voucher.headerText("cCusAbbName") = Null2Something(rshead("cCusAbbName"))
            Voucher.headerText("icode") = Null2Something(rshead("icode"))    '
            Voucher.headerText("iname") = Null2Something(rshead("iname"))    '
            Voucher.headerText("acccode") = Null2Something(rshead("acccode"))    '
            Voucher.headerText("accname") = Null2Something(rshead("accname"))
            
            Voucher.headerText("ecustcode") = Null2Something(rshead("ecustcode"))    '
            Voucher.headerText("consubname") = Null2Something(rshead("consubname"))    '
            Voucher.headerText("consubject") = Null2Something(rshead("consubject"))    '
            Voucher.headerText("custcontacta") = Null2Something(rshead("custcontacta"))    '
            Voucher.headerText("contacta") = Null2Something(rshead("contacta"))    '
            Voucher.headerText("custcontactb") = Null2Something(rshead("custcontactb"))    '
            Voucher.headerText("contactb") = Null2Something(rshead("contactb"))
            Voucher.headerText("condescriptiona") = Null2Something(rshead("condescriptiona"))
            
            Voucher.headerText("condescriptionb") = Null2Something(rshead("condescriptionb"))    '
            Voucher.headerText("condescriptionc") = Null2Something(rshead("condescriptionc"))
            
             Voucher.headerText("virtualCon") = Null2Something(rshead("virtualCon"))    '
            Voucher.headerText("smtype") = Null2Something(rshead("smtype"))
               Voucher.headerText("smapptype") = Null2Something(rshead("smapptype"))    '
            Voucher.headerText("proccode") = Null2Something(rshead("proccode"))
               Voucher.headerText("proname") = Null2Something(rshead("proname"))    '
            Voucher.headerText("engcode") = Null2Something(rshead("engcode"))
               Voucher.headerText("engname") = Null2Something(rshead("engname"))    '
            Voucher.headerText("consdate") = Null2Something(rshead("consdate"))
            
                Voucher.headerText("conedate") = Null2Something(rshead("conedate"))    '
            Voucher.headerText("coneffedate") = Null2Something(rshead("coneffedate"))   '
            Voucher.headerText("cmemo") = Null2Something(rshead("cmemo"))
            
           '
            Voucher.headerText("paytype") = Null2Something(rshead("paytype"))    '
            Voucher.headerText("contolprice") = Null2Something(rshead("contolprice"))
            Voucher.headerText("conmoney") = Null2Something(rshead("conmoney"))    '
            Voucher.headerText("designmoney") = Null2Something(rshead("designmoney"))
            Voucher.headerText("designunits") = Null2Something(rshead("designunits"))
           
            
            Voucher.headerText("conpaytolmoney") = Null2Something(rshead("conpaytolmoney"))    '
            Voucher.headerText("conpaymoney") = Null2Something(rshead("conpaymoney"))
            Voucher.headerText("accdesignmoney") = Null2Something(rshead("accdesignmoney"))    '
            Voucher.headerText("concompleted") = Null2Something(rshead("concompleted"))
            Voucher.headerText("totalappmoney") = Null2Something(rshead("totalappmoney"))    '
            Voucher.headerText("totalpaymoney") = Null2Something(rshead("totalpaymoney"))
            
            Voucher.headerText("cCusName") = Null2Something(rshead("cCusName"))    '
            Voucher.headerText("decabbname") = Null2Something(rshead("decabbname"))
            Voucher.headerText("descuname") = Null2Something(rshead("descuname"))    '
            Voucher.headerText("prodescriptiona") = Null2Something(rshead("prodescriptiona"))
            Voucher.headerText("prodescriptionb") = Null2Something(rshead("prodescriptionb"))    '
            Voucher.headerText("prodescriptionc") = Null2Something(rshead("prodescriptionc"))
            
             Voucher.headerText("engdescripta") = Null2Something(rshead("prodescriptiona"))
            Voucher.headerText("engdescriptb") = Null2Something(rshead("engdescriptb"))    '
            Voucher.headerText("engdescriptc") = Null2Something(rshead("engdescriptc"))
             Voucher.headerText("sourcetype") = "FYSL0004"
             
            Voucher.headerText("engproperties") = Null2Something(rshead("engproperties"))
            Voucher.headerText("procname") = Null2Something(rshead("procname"))
            
             Voucher.headerText("chdepartcode") = Null2Something(rshead("consubject"))
            Voucher.headerText("chdepname") = Null2Something(rshead("consubname"))
 
            '
            Voucher.headerText("iStatus") = 1   '
            Voucher.headerText("cMaker") = g_oLogin.cUserName   '
            Voucher.headerText("dmDate") = Format(g_oLogin.CurDate, "yyyy-mm-dd")    '
            Voucher.headerText("ddate") = Format(g_oLogin.CurDate, "yyyy-mm-dd")
            Voucher.headerText("sourcetype") = "参照合同"
             Voucher.headerText("addesignmoney") = Null2Something(rshead("addesignmoney"))
              Voucher.headerText("contype") = Null2Something(rshead("contype"))
            
              
              
             If val(Null2Something(rshead("conpaymoney"))) <> 0 Then
             
             Voucher.headerText("appprice") = val(Null2Something(rshead("conpaymoney"))) + val(Null2Something(rshead("accdesignmoney"))) - val(Null2Something(rshead("addesignmoney"))) - val(Null2Something(rshead("totalappmoney")))
             
             
             Else
             
             Voucher.headerText("appprice") = val(Null2Something(rshead("conmoney"))) + val(Null2Something(rshead("designmoney"))) - val(Null2Something(rshead("addesignmoney"))) - val(Null2Something(rshead("totalappmoney")))
             
             End If
            
          

            For i = 1 To 16

                Voucher.headerText("cdefine" & i) = Null2Something(rshead("cdefine" & i))    '表头自定义项

            Next

        End If

    End If

     
    Voucher.RemoveEmptyRow
    Voucher.row = 1

    If Not Form Is Nothing Then Form.IsSimulateInput = False

    Set rshead = Nothing

    Exit Function

ErrHandler:
    Set rshead = Nothing
    MsgBox Err.Description, vbInformation, GetString("U8.DZ.JA.Res030")

End Function
'参照生单处理
Public Function ProcessDataeng(Voucher As Object, Optional Form As Object = Null)

    On Error GoTo ErrHandler:

    Dim retvalue  As Variant

    Dim referpara As UAPVoucherControl85.ReferParameter

    Dim eleline   As IXMLDOMElement

    Dim echeck    As Long

    Dim i         As Integer

    Dim rstTmp    As New Recordset

    '    Dim oVoucherBo As New USERPBO.VoucherBO
    
    'Set oVoucherBo.login = g_oLogin

    '提供两种处理模式，表头使用recordset处理，表体使用xml解析处理
    Dim rshead    As New ADODB.Recordset

    rshead.Open gDomReferHead

    If Not rshead.EOF And Not rshead.BOF Then
        If rshead("selcol") = "Y" Then

   '
            Voucher.headerText("mconcode") = Null2Something(rshead("mconcode"))    '
            Voucher.headerText("mconname") = Null2Something(rshead("mconname"))
             Voucher.headerText("concode") = Null2Something(rshead("ccode"))
              Voucher.headerText("conname") = Null2Something(rshead("cname"))
            
            Voucher.headerText("cCusAbbName") = Null2Something(rshead("cCusAbbName"))
            Voucher.headerText("icode") = Null2Something(rshead("icode"))    '
            Voucher.headerText("iname") = Null2Something(rshead("iname"))    '
            Voucher.headerText("acccode") = Null2Something(rshead("acccode"))    '
            Voucher.headerText("accname") = Null2Something(rshead("accname"))
            
            Voucher.headerText("ecustcode") = Null2Something(rshead("ecustcode"))    '
            Voucher.headerText("consubname") = Null2Something(rshead("consubname"))    '
            Voucher.headerText("consubject") = Null2Something(rshead("consubject"))    '
            Voucher.headerText("custcontacta") = Null2Something(rshead("custcontacta"))    '
            Voucher.headerText("contacta") = Null2Something(rshead("contacta"))    '
            Voucher.headerText("custcontactb") = Null2Something(rshead("custcontactb"))    '
            Voucher.headerText("contactb") = Null2Something(rshead("contactb"))
            Voucher.headerText("condescriptiona") = Null2Something(rshead("condescriptiona"))
            
            Voucher.headerText("condescriptionb") = Null2Something(rshead("condescriptionb"))    '
            Voucher.headerText("condescriptionc") = Null2Something(rshead("condescriptionc"))
            
             Voucher.headerText("virtualCon") = Null2Something(rshead("virtualCon"))    '
            Voucher.headerText("smtype") = Null2Something(rshead("smtype"))
               Voucher.headerText("smapptype") = Null2Something(rshead("smapptype"))    '
            Voucher.headerText("proccode") = Null2Something(rshead("proccode"))
               Voucher.headerText("proname") = Null2Something(rshead("proname"))    '
            Voucher.headerText("engcode") = Null2Something(rshead("engcode"))
               Voucher.headerText("engname") = Null2Something(rshead("engname"))    '
            Voucher.headerText("consdate") = Null2Something(rshead("consdate"))
            
                Voucher.headerText("conedate") = Null2Something(rshead("conedate"))    '
            Voucher.headerText("coneffedate") = Null2Something(rshead("coneffedate"))   '
            Voucher.headerText("cmemo") = Null2Something(rshead("cmemo"))
            
           '
            Voucher.headerText("paytype") = Null2Something(rshead("paytype"))    '
            Voucher.headerText("contolprice") = Null2Something(rshead("contolprice"))
            Voucher.headerText("conmoney") = Null2Something(rshead("conmoney"))    '
            Voucher.headerText("designmoney") = Null2Something(rshead("designmoney"))
            Voucher.headerText("designunits") = Null2Something(rshead("designunits"))
           
            
            Voucher.headerText("conpaytolmoney") = Null2Something(rshead("conpaytolmoney"))    '
            Voucher.headerText("conpaymoney") = Null2Something(rshead("conpaymoney"))
            Voucher.headerText("accdesignmoney") = Null2Something(rshead("accdesignmoney"))    '
            Voucher.headerText("concompleted") = Null2Something(rshead("concompleted"))
            Voucher.headerText("totalappmoney") = Null2Something(rshead("totalappmoney"))    '
            Voucher.headerText("totalpaymoney") = Null2Something(rshead("totalpaymoney"))
            
            Voucher.headerText("cCusName") = Null2Something(rshead("cCusName"))    '
            Voucher.headerText("decabbname") = Null2Something(rshead("decabbname"))
            Voucher.headerText("descuname") = Null2Something(rshead("descuname"))    '
            Voucher.headerText("prodescriptiona") = Null2Something(rshead("prodescriptiona"))
            Voucher.headerText("prodescriptionb") = Null2Something(rshead("prodescriptionb"))    '
            Voucher.headerText("prodescriptionc") = Null2Something(rshead("prodescriptionc"))
            
             Voucher.headerText("engdescripta") = Null2Something(rshead("prodescriptiona"))
            Voucher.headerText("engdescriptb") = Null2Something(rshead("engdescriptb"))    '
            Voucher.headerText("engdescriptc") = Null2Something(rshead("engdescriptc"))
             Voucher.headerText("sourcetype") = "FYSL0004"
             
            Voucher.headerText("engproperties") = Null2Something(rshead("engproperties"))
            Voucher.headerText("procname") = Null2Something(rshead("procname"))
            
                Voucher.headerText("chdepartcode") = Null2Something(rshead("consubject"))
            Voucher.headerText("chdepname") = Null2Something(rshead("consubname"))
            '
            Voucher.headerText("iStatus") = 1   '
            Voucher.headerText("cMaker") = g_oLogin.cUserName   '
            Voucher.headerText("dmDate") = Format(g_oLogin.CurDate, "yyyy-mm-dd")    '
             Voucher.headerText("ddate") = Format(g_oLogin.CurDate, "yyyy-mm-dd")
             Voucher.headerText("sourcetype") = "参照设计费"
             Voucher.headerText("addesignmoney") = Null2Something(rshead("addesignmoney"))
             Voucher.headerText("contype") = Null2Something(rshead("contype"))
             
            If val(Null2Something(rshead("accdesignmoney"))) <> 0 Then
             Voucher.headerText("appprice") = val(Null2Something(rshead("addesignmoney"))) - val(Null2Something(rshead("designmoney")))
             
             Else
             Voucher.headerText("appprice") = val(Null2Something(rshead("addesignmoney"))) - val(Null2Something(rshead("accdesignmoney")))
             
             
             End If
          

            For i = 1 To 16

                Voucher.headerText("cdefine" & i) = Null2Something(rshead("cdefine" & i))    '表头自定义项

            Next
      
        End If

    End If
 
    Voucher.RemoveEmptyRow
    Voucher.row = 1

    If Not Form Is Nothing Then Form.IsSimulateInput = False

    Set rshead = Nothing

    Exit Function

ErrHandler:
    Set rshead = Nothing
    MsgBox Err.Description, vbInformation, GetString("U8.DZ.JA.Res030")

End Function

'初始化全局对象变量
Public Sub GlobalInit(login As clsLogin)

    'TODO:
    '变量初始化
    gstrCardNumberlist = "PD010301"
    gstrCardNumber = ""

    '数据表定义
    MainTable = ""                          '单据主表
    DetailsTable = ""                      '单据字表
    HeadPKFld = "ID"                                       '主表主键字段
    MainView = ""                         '表头视图
    DetailsView = ""                     '表体视图
'    VoucherList = "EF_v_InScanMaintmp"                  '列表视图
    
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

End Sub

'单据精度控制
Public Sub FormatVouchList(rs As ADODB.Recordset)

    On Error GoTo ErrHandler:


    Dim sQuantityFmt As String                             ' 数量小数位
    Dim sNumFmt As String                                  ' 件数小数位数
    Dim iExchRateFmt As String                             ' 换算率小数位数
    Dim m_iRateFmt As String                               ' 税率小数位数
    Dim sPriceFmt As String                                ' 存货单价小数位(采购用)（库存用）
    Dim sPriceFmtSA As String                              ' 开票单价小数位(销售用)

    sQuantityFmt = m_SysInfor.iQuantityBit
    sNumFmt = m_SysInfor.iNumBit
    iExchRateFmt = m_SysInfor.iExchRateBit
    m_iRateFmt = m_SysInfor.iRateBit
    sPriceFmt = m_SysInfor.iCostBit
    sPriceFmtSA = m_SysInfor.iBillCostBit

    Dim DomFormat As New DOMDocument
    rs.Save DomFormat, adPersistXML
    rs.Close

    '    SetFormat DomFormat, "cfreightCost", sQuantityFmt  '运费 表头
    '    SetFormat DomFormat, "iquantity", sQuantityFmt     '数量
    '    SetFormat DomFormat, "inum", sQuantityFmt          '辅数量
    '    SetFormat DomFormat, "iinvexchrate", iExchRateFmt  '换算率
    '    SetFormat DomFormat, "iQtyOutSum", sQuantityFmt       '累计出库数量
    '    SetFormat DomFormat, "iQtyOut2Sum", sQuantityFmt      '
    '    SetFormat DomFormat, "iQtyBackSum", sQuantityFmt      '累计归还数量
    '    SetFormat DomFormat, "iQtyBack2Sum", sQuantityFmt
    '    SetFormat DomFormat, "iQtyCOutSum", sQuantityFmt       '累计转借出数量
    '    SetFormat DomFormat, "iQtyCOut2Sum", sQuantityFmt      '
    '    SetFormat DomFormat, "iQtyCSaleSum", sQuantityFmt      '累计转销售数量
    '    SetFormat DomFormat, "iQtyCSale2Sum", sQuantityFmt
    '    SetFormat DomFormat, "iQtyCFreeSum", sQuantityFmt       '累计转赠品数量
    '    SetFormat DomFormat, "iQtyCFree2Sum", sQuantityFmt      '
    '    SetFormat DomFormat, "iQtyCOverSum", sQuantityFmt       '累计转耗用辅数量
    '    SetFormat DomFormat, "iQtyCOver2Sum", sQuantityFmt

    '    Dim sQuantityFmt As String  ' 数量小数位
    '    Dim sNumFmt As String ' 件数小数位数
    '    Dim iExchRateFmt As String ' 换算率小数位数
    '    Dim m_iRateFmt As String  ' 税率小数位数
    '    Dim sPriceFmt As String  ' 存货单价小数位(采购用)（库存用）
    '    Dim sPriceFmtSA  As String  ' 开票单价小数位(销售用)

    '    MsgBox "11 "

    SetFormat DomFormat, "cfreightCost", sPriceFmt
    SetFormat DomFormat, "iinvexchrate", iExchRateFmt

    SetFormat DomFormat, "iquantity", sQuantityFmt
    SetFormat DomFormat, "inum", sNumFmt
    SetFormat DomFormat, "iQtyBackSum", sQuantityFmt
    SetFormat DomFormat, "iQtyBack2Sum", sNumFmt
    SetFormat DomFormat, "iQtyCFreeSum", sQuantityFmt
    SetFormat DomFormat, "iQtyCFree2Sum", sNumFmt
    SetFormat DomFormat, "iQtyCOutSum", sQuantityFmt
    SetFormat DomFormat, "iQtyCOut2Sum", sNumFmt
    SetFormat DomFormat, "iQtyCOverSum", sQuantityFmt
    SetFormat DomFormat, "iQtyCOver2Sum", sNumFmt
    SetFormat DomFormat, "iQtyCSaleSum", sQuantityFmt
    SetFormat DomFormat, "iQtyCSale2Sum", sNumFmt
    SetFormat DomFormat, "iQtyOutSum", sQuantityFmt
    SetFormat DomFormat, "iQtyOut2Sum", sNumFmt

    SetFormat DomFormat, "cdefine7", sQuantityFmt
    SetFormat DomFormat, "cdefine16", sQuantityFmt
    SetFormat DomFormat, "cdefine26", sQuantityFmt
    SetFormat DomFormat, "cdefine27", sQuantityFmt

    '    MsgBox "22 "

    'dxb
    'enum by modify
    Dim strShowFormat As String
    If gcCreateType = "期初单据" Then
        strShowFormat = "False"
    Else
        strShowFormat = "True"
    End If
    SetFormat2 DomFormat, "iQtyOutSum", strShowFormat
    SetFormat2 DomFormat, "iQtyOut2Sum", strShowFormat

    If gcCreateType = "期初单据" Then
        SetFormat3 DomFormat, "iquantity", "出库数量"
        SetFormat3 DomFormat, "inum", "出库件数"
    End If

    rs.Open DomFormat
    Set DomFormat = Nothing
    Exit Sub
ErrHandler:
    MsgBox Err.Description, vbInformation, GetString("U8.DZ.JA.Res030")
    Set DomFormat = Nothing
End Sub

'设置单据模板的表头表体项目的可见性 dixingben 2009/5/21
Public Sub SetFormat3(DomFormat As DOMDocument, FieldName As String, FormatStr As String)
    Dim ele As IXMLDOMElement
    Set ele = DomFormat.selectSingleNode("//z:row[@FieldName='" & FieldName & "']")
    ele.setAttribute "cardformula1", FormatStr
    Set ele = Nothing
End Sub

'设置单据模板的表头表体项目的可见性 dixingben 2009/5/21
Public Sub SetFormat2(DomFormat As DOMDocument, FieldName As String, FormatStr As String)
    Dim ele As IXMLDOMElement
    Set ele = DomFormat.selectSingleNode("//z:row[@FieldName='" & FieldName & "']")
    ele.setAttribute "ShowIt", FormatStr
    Set ele = Nothing
End Sub

'设置单据模板的小数位数
Public Sub SetFormat(DomFormat As DOMDocument, FieldName As String, FormatStr As String)
    Dim ele As IXMLDOMElement
    Set ele = DomFormat.selectSingleNode("//z:row[@FieldName='" & FieldName & "']")
    ele.setAttribute "NumPoint", FormatStr
    Set ele = Nothing
End Sub

'
'写销售类单据
Public Function WriteSABill(ByRef oDomHead As DOMDocument, ByRef oDomBody As DOMDocument, conn As Object, login As clsLogin, VoucherType As String) As Boolean
    On Error GoTo ErrHandler

    Dim SAHead As DOMDocument, SABody As DOMDocument
    Dim r As New ADODB.Recordset
    Dim rsize As Integer, rows As Long, dateStr As String
    '    Dim uprows As Integer
    '
    Dim retDom As New DOMDocument
    Dim rs As New ADODB.Recordset, fRs As New ADODB.Recordset
    Dim ViewHead As String, ViewBody As String
    Dim ele As IXMLDOMElement

    Dim oNode As IXMLDOMElement, oNodes As IXMLDOMElement
    Dim lstBody As IXMLDOMNodeList                         '表体节点列表


    Dim sVoucherID As String                               '单据ID

    Dim errMsg As String

    Dim strSql As String
    Dim txtSQL As String
    Dim i As Integer


    Dim lrows As Long
    '
    Dim cSoCodeStr As String, sqlstr As String, vidstr As String
    Dim RsTemp As ADODB.Recordset, vouchID As String, oneVouchID As String

    Dim voucherSuccSize As Integer                         '记录生成功的单子数量  2008-01-30
    Dim voucherErrMsg As String                            '记录生成失败的消息　　2008-01-30
    Dim exchNameStr As String
    Dim rsKL As ADODB.Recordset
    Dim CurrentId As String


    '/*********************************************************************************/'
    Dim pco As Object                                      'New VoucherCO_Sa.ClsVoucherCO_SA
    Dim clsSysSa As Object                                 'USSAServer.clsSystem

    Set pco = CreateObject("VoucherCO_Sa.ClsVoucherCO_SA")
    Set clsSysSa = CreateObject("USSAServer.clsSystem")
    '初始化对象
    'Set clsSysSa = New USSAServer.clsSystem
    clsSysSa.Init login
    clsSysSa.INIMyInfor
    clsSysSa.bManualTrans = True                           '2008-01-10
    '初始化销售生单接口,"97"表示销售订单
    'Pco.Init VoucherTypeSA.SODetails, login, conn, "CS", clsSysSa
    pco.Init VoucherType, login, conn, "CS", clsSysSa
    '/*********************************************************************************/'
    voucherSuccSize = 0
    voucherErrMsg = ""
    vouchID = ""

    '根据不同的审单类型组织不同的dom

    '表头用recordset处理,表体用xml解析处理
    r.Open oDomHead
    rsize = r.RecordCount
    For i = 1 To rsize

        rows = 0
        dateStr = Format(Now(), "YYYY-MM-DD HH:MM:SS")

        Set SAHead = New DOMDocument
        Set SABody = New DOMDocument
        ViewHead = GetViewHead(conn, "17")
        ViewBody = GetViewBody(conn, "17")

        '相同存货+预发货日期+预完工日期生成相同采购订单表体记录，数量汇总生成
        '写表头----------------------------------------------------------------------------------
        strSql = "select *,'' as editprop from " & ViewHead & " where 1>2"
        rs.Open strSql, conn, adOpenDynamic, adLockOptimistic
        rs.Save SAHead, adPersistXML                       '得到入库单表头DOM结构对象
        rs.Close
        Set oNodes = SAHead.selectSingleNode("//rs:data")
        Set oNode = SAHead.createElement("z:row")

        CurrentId = r!id                                   '记录当前的主表标示

        oNode.setAttribute "ccuscode", Null2Something(r!cCusCode)    '客户编码
        '2008-01-23 客户名称，发运方式编码，发运方式,付款条件编码，付款条件,发货地址,变更人
        Dim ftmpData As String
        If fRs.State = adStateOpen Then fRs.Close
        Set fRs = conn.Execute("Select cCusSAProtocol,ccusname,cCusOType,cCusPayCond,cCusOAddress,cModifyPerson From Customer Where ccuscode='" & r!cCusCode & "'")
        If Not fRs.EOF Then
            oNode.setAttribute "ccusname", Null2Something(fRs!ccusname)    '客户名称
            oNode.setAttribute "csccode", Null2Something(fRs!cCusOType)    '发运方式编码
            oNode.setAttribute "cpaycode", Null2Something(fRs!cCusPayCond)    '付款条件编码
            oNode.setAttribute "ccusoaddress", Null2Something(fRs!ccusoaddress)    '发货地址
            oNode.setAttribute "cgatheringplan", Null2Something(fRs!cCusSAProtocol)    '付款协议编码
            'oNode.setAttribute "cchanger", null2something(fRs!cModifyPerson)                     '变更人
            ftmpData = Null2Something(fRs!cCusPayCond)
            '发运方式
            sqlstr = "Select cscName From ShippingChoice Where cscCode='" & fRs!cCusOType & "'"
            If fRs.State = adStateOpen Then fRs.Close
            Set fRs = conn.Execute(sqlstr)
            If Not fRs.EOF Then oNode.setAttribute "cscname", Null2Something(fRs!cscname)
            '付款条件
            sqlstr = "Select cPayName From PayCondition Where cPayCode='" & ftmpData & "'"
            If fRs.State = adStateOpen Then fRs.Close
            Set fRs = conn.Execute(sqlstr)
            If Not fRs.EOF Then oNode.setAttribute "cpayname", Null2Something(fRs!cPayName)
        End If
        If fRs.State = adStateOpen Then fRs.Close

        oNode.setAttribute "cbustype", "普通销售"              '销售类型 "普通销售"

        '取默认销售类型,其次取第一个销售类型
        If fRs.State = adStateOpen Then fRs.Close
        Set fRs = conn.Execute("select * from saleType where bdefault=1")
        If fRs.EOF Then
            fRs.Close
            Set fRs = conn.Execute("select top1 * from saleType")
            If fRs.EOF Then
                MsgBox GetString("U8.DZ.JA.Res830"), vbInformation, GetString("U8.DZ.JA.Res030")
                Exit Function
            Else
                oNode.setAttribute "cstcode", Null2Something(fRs!cstcode)    '销售类型
                fRs.Close
            End If
        Else
            oNode.setAttribute "cstcode", Null2Something(fRs!cstcode)    '销售类型
            fRs.Close
        End If

        oNode.setAttribute "cpersoncode", Null2Something(r!cpersoncode)    '业务员
        oNode.setAttribute "cdepcode", Null2Something(r!cDepcode)    '部门编码

        oNode.setAttribute "cdefine1", Null2Something(r!cDefine1)    '表头自定义项
        oNode.setAttribute "cdefine2", Null2Something(r!cDefine2)    '表头自定义项
        oNode.setAttribute "cdefine3", Null2Something(r!cDefine3)    '表头自定义项
        oNode.setAttribute "cdefine4", Null2Something(r!cDefine4)    '表头自定义项
        oNode.setAttribute "cdefine5", Null2Something(r!cDefine5)    '表头自定义项
        oNode.setAttribute "cdefine6", Null2Something(r!cDefine6)    '表头自定义项
        oNode.setAttribute "cdefine7", Null2Something(r!cdefine7)    '表头自定义项
        oNode.setAttribute "cdefine8", Null2Something(r!cDefine8)    '表头自定义项
        oNode.setAttribute "cdefine9", Null2Something(r!cDefine9)    '表头自定义项
        oNode.setAttribute "cdefine10", Null2Something(r!cDefine10)    '表头自定义项
        oNode.setAttribute "cdefine11", Null2Something(r!cDefine11)    '表头自定义项
        oNode.setAttribute "cdefine12", Null2Something(r!cDefine12)    '表头自定义项
        oNode.setAttribute "cdefine13", Null2Something(r!cDefine13)    '表头自定义项
        oNode.setAttribute "cdefine14", Null2Something(r!cDefine14)    '表头自定义项
        oNode.setAttribute "cdefine15", Null2Something(r!cDefine15)    '表头自定义项
        oNode.setAttribute "cdefine16", Null2Something(r!cdefine16)    '表头自定义项


        oNode.setAttribute "ivtid", GetVoucherID(conn, "17")
        oNode.setAttribute "cmaker", login.cUserName
        oNode.setAttribute "cmemo", GetString("U8.DZ.JA.Res840") & Format(Time, "HH:MM:SS")    ' & "班次:"
        oNode.setAttribute "cvouchtype", "97"              '单据类型

        oNode.setAttribute "cexch_name", Null2Something(r!cexch_name)    '币种
        oNode.setAttribute "iexchrate", Null2Something(r!iExchRate)    '汇率

        oNode.setAttribute "dpredatebt", ""                '预发货日期
        oNode.setAttribute "dpremodatebt", ""              '预完工日期
        oNode.setAttribute "ddate", dateStr

        'oNode.setAttribute "cchanger", ""           '变更人
        oNode.setAttribute "ccrechpname", ""               '信用审核人
        'oNode.setAttribute "csccode", ""            '运输方式编码
        'oNode.setAttribute "cpaycode", ""           '付款条件编码
        oNode.setAttribute "istatus", ""                   '状态
        oNode.setAttribute "cverifier", ""                 '审核人

        oNode.setAttribute "itaxrate", "17"                '税率
        oNode.setAttribute "imoney", ""                    '定金
        oNode.setAttribute "ccloser", ""                   '关闭人
        oNode.setAttribute "cstname", "普通销售"               '销售类型名称
        oNode.setAttribute "iarmoney", "0"                 '应收余额
        oNode.setAttribute "bdisflag", "0"                 '是否整单打折
        oNode.setAttribute "clocker", ""                   '锁定人
        oNode.setAttribute "breturnflag", "0"              '退货标志
        oNode.setAttribute "icuscreline", "0"              '信用额度
        oNode.setAttribute "coppcode", ""                  '商机编码
        oNode.setAttribute "caddcode", ""                  '收货地址编码
        oNode.setAttribute "iverifystate", ""              'iVerifyState 工作流用
        oNode.setAttribute "ireturncount", ""              'iReturnCount 工作流用
        oNode.setAttribute "icreditstate", ""              '信用审批状态
        oNode.setAttribute "iswfcontrolled", ""            'IsWFControlled 工作流用
        oNode.setAttribute "editprop", "A"


        oNodes.appendChild oNode



        '根据currentid处理表体----------------------------------------

        strSql = "select *,'' as editprop from " & ViewBody & " where 1>2"
        rs.Open strSql, conn, adOpenDynamic, adLockOptimistic
        rs.Save SABody, adPersistXML
        rs.Close


        Set lstBody = oDomBody.selectNodes("//z:row[@" + HeadPKFld + "='" + CurrentId + "']")

        rows = lstBody.Length
        If rows > 0 Then
            For Each ele In lstBody

                Set oNodes = SABody.selectSingleNode("//rs:data")
                Set oNode = SABody.createElement("z:row")

                If Trim(GetNodeAtrVal(ele, "iquantity")) <> "0" Then    '汇总数量如果是0的时候，不写到入库单上
                    oNode.setAttribute "cconfigstatus", "" '描述
                    oNode.setAttribute "ikpquantity", ""   '已开票数量
                    oNode.setAttribute "iprekeepquantity", ""    '预留数量
                    oNode.setAttribute "ccontractrowguid", ""    '合同cGuid
                    oNode.setAttribute "ccontracttagcode", ""    '合同标的
                    oNode.setAttribute "bproductbill", ""  '允许生产订单
                    oNode.setAttribute "ccontractid", ""   '合同号
                    oNode.setAttribute "dEDate", ""        '企业负责人签字日期
                    oNode.setAttribute "cquocode", ""      '报价单号
                    oNode.setAttribute "iprekeeptotquantity", ""    '预留总数量
                    oNode.setAttribute "bTrack", ""        '是否出库跟踪入库
                    oNode.setAttribute "bInvBatch", ""     '是否批次管理
                    oNode.setAttribute "bspecialorder", "" '客户订单专用
                    oNode.setAttribute "ballpurchase", ""  '已经齐套采购标志
                    oNode.setAttribute "iprekeeptotnum", "" '预留总件数
                    oNode.setAttribute "bProxyForeign", "" '是否委外
                    oNode.setAttribute "fomquantity", ""   '委外数量
                    oNode.setAttribute "fimquantity", ""   '进口数量
                    oNode.setAttribute "dreleasedate", ""  '释放日期
                    oNode.setAttribute "iadvancedate", ""  '累计提前期
                    oNode.setAttribute "iprekeepnum", ""   '预留件数
                    oNode.setAttribute "fpurquan", ""      '采购数量
                    oNode.setAttribute "iquoid", ""        '报价单id
                    oNode.setAttribute "csrpolicy", "PE"   '供需政策
                    oNode.setAttribute "binvmodel", "否"    '模型
                    oNode.setAttribute "ippartqty", ""     '母件数量
                    oNode.setAttribute "ippartid", ""      '母件物料ID
                    oNode.setAttribute "ippartseqid", ""   'PTO母件顺序号
                    oNode.setAttribute "citem_class", GetNodeAtrVal(ele, "citem_class")    '项目大类编码
                    oNode.setAttribute "citemcode", GetNodeAtrVal(ele, "citemcode")    '项目编码
                    oNode.setAttribute "citemname", GetNodeAtrVal(ele, "citemname")    '项目名称
                    oNode.setAttribute "icusbomid", ""     '客户bomid
                    oNode.setAttribute "imoquantity", ""   '下达生产数量
                    oNode.setAttribute "irowno", " 1"      '订单行号
                    oNode.setAttribute "binvtype", "0"     '是否折扣
                    'oNode.setAttribute "iquotedprice", "0" '"1000.00"   '报价
                    oNode.setAttribute "cscloser", ""      '关闭标记
                    oNode.setAttribute "bservice", "0"     '是否应税劳务
                    'oNode.setAttribute "inum", "0"                      '件数
                    oNode.setAttribute "itax", "0"         '"12815.38"          '原币税额
                    oNode.setAttribute "isum", "0"         '"88200.00"          '原币价税合计
                    oNode.setAttribute "inatsum", "0"      ' "88200.00"      '本币价税合计
                    oNode.setAttribute "inatdiscount", "0" ' "11800.00" '本币折扣额
                    oNode.setAttribute "fsaleprice", "0"   '零售金额
                    oNode.setAttribute "ikpnum", ""        '累计开票辅计量数量
                    oNode.setAttribute "ikpmoney", ""      '累计原币开票金额
                    oNode.setAttribute "iinvexchrate", "0" ' "50.00"    '换算率
                    oNode.setAttribute "idiscount", "0"    ' "11800.00"    '原币折扣额
                    oNode.setAttribute "fsalecost", "0"    '零售单价
                    oNode.setAttribute "ifhquantity", ""   '累计发货数量
                    oNode.setAttribute "ifhnum", ""        '累计发货辅计量数量
                    oNode.setAttribute "itaxunitprice", "0" ' "882.00"  '原币含税单价
                    oNode.setAttribute "iunitprice", "0"   '"753.85"      '原币无税单价
                    oNode.setAttribute "inatunitprice", "0" '"753.85"   '本币无税单价
                    oNode.setAttribute "inatmoney", "0"    ' "75384.62"    '本币无税金额
                    oNode.setAttribute "inattax", "0"      ' "12815.38"      '本币税额
                    oNode.setAttribute "imoney", "0"       ' "75384.62"       '本币金额
                    oNode.setAttribute "itaxrate", "0"     '表头税率
                    oNode.setAttribute "ifhmoney", ""      '累计原币发货金额
                    oNode.setAttribute "idemantype", GetNodeAtrVal(ele, "sotype")
                    If GetNodeAtrVal(ele, "sotype") = "4" Then
                        oNode.setAttribute "cdemandcode", GetNodeAtrVal(ele, "socode")
                        oNode.setAttribute "cdemandmemo", GetNodeAtrVal(ele, "cdemandmemo")
                    End If

                    If gU8Version = "872" Then
                        sqlstr = "Select bFree1,bFree2,bFree3,bFree4,bFree5,bFree6,bFree7,bFree8,bFree9,bFree10 From Inventory Where cinvcode='" & GetNodeAtrVal(ele, "cinvcode") & "'"
                        If fRs.State = adStateOpen Then fRs.Close
                        Set fRs = conn.Execute(sqlstr)
                        If Not fRs.EOF Then
                            oNode.setAttribute "bsalepricefree1", fRs.Fields("bFree1")
                            oNode.setAttribute "bsalepricefree2", fRs.Fields("bFree2")
                            oNode.setAttribute "bsalepricefree3", fRs.Fields("bFree3")
                            oNode.setAttribute "bsalepricefree4", fRs.Fields("bFree4")
                            oNode.setAttribute "bsalepricefree5", fRs.Fields("bFree5")
                            oNode.setAttribute "bsalepricefree6", fRs.Fields("bFree6")
                            oNode.setAttribute "bsalepricefree7", fRs.Fields("bFree7")
                            oNode.setAttribute "bsalepricefree8", fRs.Fields("bFree8")
                            oNode.setAttribute "bsalepricefree9", fRs.Fields("bFree9")
                            oNode.setAttribute "bsalepricefree10", fRs.Fields("bFree10")
                        Else
                            oNode.setAttribute "bsalepricefree1", "0"
                            oNode.setAttribute "bsalepricefree2", "0"
                            oNode.setAttribute "bsalepricefree3", "0"
                            oNode.setAttribute "bsalepricefree4", "0"
                            oNode.setAttribute "bsalepricefree5", "0"
                            oNode.setAttribute "bsalepricefree6", "0"
                            oNode.setAttribute "bsalepricefree7", "0"
                            oNode.setAttribute "bsalepricefree8", "0"
                            oNode.setAttribute "bsalepricefree9", "0"
                            oNode.setAttribute "bsalepricefree10", "0"
                        End If
                    End If


                    '最低售价
                    Dim fiInvLSCost As Double
                    sqlstr = "Select iInvLSCost From Inventory Where cinvcode='" & Null2Something(GetNodeAtrVal(ele, "cinvcode")) & "'"
                    If fRs.State = adStateOpen Then fRs.Close
                    Set fRs = conn.Execute(sqlstr)
                    '最低售价
                    If Not fRs.EOF Then
                        oNode.setAttribute "iinvlscost", Null2Something(fRs!iInvLSCost)
                        fiInvLSCost = val(Null2Something(fRs!iInvLSCost))
                    Else
                        oNode.setAttribute "iinvlscost", ""
                    End If
                    '客户最低售价
                    sqlstr = "Select fCusminPrice From SA_CusPriceJustdetail Where ccuscode='" & Null2Something(r!cCusCode) & "' And cinvcode='" & Null2Something(GetNodeAtrVal(ele, "cinvcode")) & "'"
                    If fRs.State = adStateOpen Then fRs.Close
                    Set fRs = conn.Execute(sqlstr)
                    '客户最低售价
                    If Not fRs.EOF Then
                        If IsNull(fRs!fCusMinPrice) Then
                            oNode.setAttribute "fcusminprice", str(fiInvLSCost)
                        Else
                            oNode.setAttribute "fcusminprice", Null2Something(fRs!fCusMinPrice)    '客户最低售价
                        End If
                    Else
                        oNode.setAttribute "fcusminprice", ""
                    End If


                    '察看价格选贤是否启用价格政策             xin            2008-10-22
                    Dim klValue As Double, kl2Value As Double
                    Set rsKL = conn.Execute("select cvalue from accinformation where cname='bquantitydisrate'")
                    If Not (rsKL Is Nothing) Then
                        If rsKL.Fields(0) Then
                            '取扣率2
                            getKL2 conn, Null2Something(r!cCusCode), Null2Something(GetNodeAtrVal(ele, "cinvcode")), GetNodeAtrVal(ele, "iquantity"), kl2Value, ""
                        Else
                            kl2Value = 100
                        End If
                        rsKL.Close
                    Else
                        kl2Value = 100
                    End If

                    '2004-01-24
                    Select Case clsSysSa.CostGetType
                        Case 0, 1                          '0:最新价格  1：最新成本价格
                            oNode.setAttribute "kl", "100" '扣率
                            oNode.setAttribute "kl2", kl2Value    '二次扣率
                            oNode.setAttribute "dkl1", "0" '倒扣率1
                            oNode.setAttribute "dkl2", "0" '倒扣率2

                        Case 2                             '价格政策

                            '取扣率
                            If gU8Version = "870" Then
                                getKL conn, Null2Something(r!cCusCode), Null2Something(GetNodeAtrVal(ele, "cinvcode")), Null2Something(GetNodeAtrVal(ele, "cfree1")), Null2Something(GetNodeAtrVal(ele, "cfree2")), Format(dateStr, "YYYY-MM-DD"), GetNodeAtrVal(ele, "iquantity"), clsSysSa, klValue
                            ElseIf gU8Version = "871" Then
                                getKL871 conn, Null2Something(r!cCusCode), Null2Something(GetNodeAtrVal(ele, "cinvcode")), Null2Something(GetNodeAtrVal(ele, "cfree1")), Null2Something(GetNodeAtrVal(ele, "cfree2")), Format(dateStr, "YYYY-MM-DD"), GetNodeAtrVal(ele, "iquantity"), exchNameStr, clsSysSa, klValue
                                '872扣率   xin 2008-10-22
                            ElseIf gU8Version = "872" Then
                                getKL872 conn, Null2Something(r!cCusCode), Null2Something(GetNodeAtrVal(ele, "cinvcode")), Null2Something(GetNodeAtrVal(ele, "cfree1")), Null2Something(GetNodeAtrVal(ele, "cfree2")), GetNodeAtrVal(ele, "cfree3") & "", GetNodeAtrVal(ele, "cfree4") & "", GetNodeAtrVal(ele, "cfree5") & "", GetNodeAtrVal(ele, "cfree6") & "", GetNodeAtrVal(ele, "cfree7") & "", GetNodeAtrVal(ele, "cfree8") & "", GetNodeAtrVal(ele, "cfree9") & "", GetNodeAtrVal(ele, "cfree10") & "", Format(dateStr, "YYYY-MM-DD"), GetNodeAtrVal(ele, "iquantity"), exchNameStr, clsSysSa, klValue
                            End If


                            oNode.setAttribute "kl", klValue    '扣率
                            oNode.setAttribute "kl2", kl2Value    '二次扣率
                            oNode.setAttribute "dkl1", str(100 - klValue)    '倒扣率1
                            oNode.setAttribute "dkl2", str(100 - kl2Value)    '倒扣率2
                    End Select


                    oNode.setAttribute "iinvid", ""
                    oNode.setAttribute "cdefine22", GetNodeAtrVal(ele, "cdefine22")    '表体自定义项
                    oNode.setAttribute "cdefine23", GetNodeAtrVal(ele, "cdefine23")    '表体自定义项
                    oNode.setAttribute "cdefine24", GetNodeAtrVal(ele, "cdefine24")    '表体自定义项
                    oNode.setAttribute "cdefine25", GetNodeAtrVal(ele, "cdefine25")    '表体自定义项
                    oNode.setAttribute "cdefine26", GetNodeAtrVal(ele, "cdefine26")    '表体自定义项
                    oNode.setAttribute "cdefine27", GetNodeAtrVal(ele, "cdefine27")    '表体自定义项
                    oNode.setAttribute "cdefine28", GetNodeAtrVal(ele, "cdefine28")    '表体自定义项
                    oNode.setAttribute "cdefine29", GetNodeAtrVal(ele, "cdefine29")    '表体自定义项
                    oNode.setAttribute "cdefine30", GetNodeAtrVal(ele, "cdefine30")    '表体自定义项
                    oNode.setAttribute "cdefine31", GetNodeAtrVal(ele, "cdefine31")    '表体自定义项
                    oNode.setAttribute "cdefine32", GetNodeAtrVal(ele, "cdefine32")    '表体自定义项
                    oNode.setAttribute "cdefine33", GetNodeAtrVal(ele, "cdefine33")    '表体自定义项
                    oNode.setAttribute "cdefine34", GetNodeAtrVal(ele, "cdefine34")    '表体自定义项
                    oNode.setAttribute "cdefine35", GetNodeAtrVal(ele, "cdefine35")    '表体自定义项
                    oNode.setAttribute "cdefine36", GetNodeAtrVal(ele, "cdefine36")    '表体自定义项
                    oNode.setAttribute "cdefine37", GetNodeAtrVal(ele, "cdefine37")    '表体自定义项

                    '2008-01-23 换算率,件数
                    sqlstr = "Select iChangRate From ComputationUnit Where cComunitCode In (Select cSAComUnitCode From Inventory Where cinvcode='" & GetNodeAtrVal(ele, "cinvcode") & "')"
                    If fRs.State = adStateOpen Then fRs.Close
                    Set fRs = conn.Execute(sqlstr)
                    If Not fRs.EOF Then
                        '无固定换算率                 2008-11-06            xin
                        If Not IsNull(fRs.Fields(0)) Then
                            oNode.setAttribute "iinvexchrate", fRs.Fields("iChangRate")    '换算率
                            If fRs.Fields("iChangRate") <> 0 Then oNode.setAttribute "inum", GetNodeAtrVal(ele, "iquantity") / fRs.Fields("iChangRate")    '件数
                        End If
                    End If
                    '2008-01-15 税率
                    If fRs.State = adStateOpen Then fRs.Close
                    Dim taxRateNode As IXMLDOMElement, taxRateFlag As Boolean

                    sqlstr = "Select iTaxRate From Inventory Where cinvcode='" & GetNodeAtrVal(ele, "cinvcode") & "'"
                    Set fRs = conn.Execute(sqlstr)
                    taxRateFlag = True
                    If Not fRs.EOF Then
                        If Not IsNull(fRs!itaxrate) Then
                            oNode.setAttribute "itaxrate", Null2Something(fRs!itaxrate)


                            Set taxRateNode = SAHead.selectSingleNode("//z:row")
                            taxRateNode.setAttribute "itaxrate", Null2Something(fRs!itaxrate)

                        Else
                            taxRateFlag = False
                        End If
                    Else
                        taxRateFlag = False
                    End If

                    If Not taxRateFlag Then
                        If Null2Something(SAHead.selectSingleNode("//z:row").Attributes.getNamedItem("itaxrate").nodeValue) <> "" Then
                            '表头带到表体
                            oNode.setAttribute "itaxrate", Null2Something(SAHead.selectSingleNode("//z:row").Attributes.getNamedItem("itaxrate").nodeValue)
                        Else
                            '表头取默认值
                            Set taxRateNode = SAHead.selectSingleNode("//z:row")
                            taxRateNode.setAttribute "itaxrate", "0"
                        End If
                    End If
                    fRs.Close


                    oNode.setAttribute "dpredate", dateStr '预发货日期
                    oNode.setAttribute "dpremodate", dateStr    '预完工日期
                    '由PartID 查找存货编码及自由项-------------------------
                    sqlstr = "select cComUnitCode,cAssComUnitCode,Free1,Free2,Free3,Free4,Free5,Free6, " & _
                            " Free7,Free8,Free9,Free10 ,cSTComUnitCode,iGroupType,cGroupCode " & _
                            " from v_bas_part  " & _
                            " where cinvcode ='" & GetNodeAtrVal(ele, "cinvcode") & "' And Free1='" & GetNodeAtrVal(ele, "cfree1") & "' And Free2='" & GetNodeAtrVal(ele, "cfree2") & "' And Free3='" & GetNodeAtrVal(ele, "cfree3") & "'And " & _
                            " Free4='" & GetNodeAtrVal(ele, "cfree4") & "' And Free5='" & GetNodeAtrVal(ele, "cfree5") & "' And Free6='" & GetNodeAtrVal(ele, "cfree6") & "' And Free7='" & GetNodeAtrVal(ele, "cfree7") & "' And Free8='" & GetNodeAtrVal(ele, "cfree8") & "' And Free9='" & GetNodeAtrVal(ele, "cfree9") & "' And Free10='" & GetNodeAtrVal(ele, "cfree10") & "'"
                    Set fRs = conn.Execute(sqlstr)

                    If fRs.EOF = False Then
                        oNode.setAttribute "cbarcode", ""  '对应条形码编码
                        oNode.setAttribute "cinvcode", Null2Something(GetNodeAtrVal(ele, "cinvcode"))    '存货编码
                        oNode.setAttribute "iquantity", Null2Something(GetNodeAtrVal(ele, "iquantity"))    '数量
                        oNode.setAttribute "igrouptype", Null2Something(fRs!iGroupType)    '计量单位组类别
                        oNode.setAttribute "cfree1", Null2Something(fRs!Free1)
                        oNode.setAttribute "cfree2", Null2Something(fRs!Free2)
                        oNode.setAttribute "cfree3", Null2Something(fRs!Free3)
                        oNode.setAttribute "cfree4", Null2Something(fRs!Free4)
                        oNode.setAttribute "cfree5", Null2Something(fRs!Free5)
                        oNode.setAttribute "cfree6", Null2Something(fRs!Free6)
                        oNode.setAttribute "cfree7", Null2Something(fRs!Free7)
                        oNode.setAttribute "cfree8", Null2Something(fRs!Free8)
                        oNode.setAttribute "cfree9", Null2Something(fRs!Free9)
                        oNode.setAttribute "cfree10", Null2Something(fRs!Free10)
                        oNode.setAttribute "ccomunitcode", Null2Something(fRs!cComUnitCode)    '主计量单位编码
                        oNode.setAttribute "cgroupcode", Null2Something(fRs!cGroupCode)    '"05"           '计量单位组编码

                        fRs.Close

                        sqlstr = "Select cSAComUnitCode From Inventory Where cinvcode='" & Trim(GetNodeAtrVal(ele, "cinvcode")) & "'"
                        Set fRs = conn.Execute(sqlstr)
                        oNode.setAttribute "cunitid", Null2Something(fRs!cSAComUnitCode)    '计量单位

                    End If
                    fRs.Close

                    oNode.setAttribute "cmemo", ""

                    oNode.setAttribute "editprop", "A"

                    oNodes.appendChild oNode
                End If
            Next


            Dim strErr As String
            Dim domHead As New DOMDocument
            Dim domBody As New DOMDocument



            If SABody.selectNodes("//z:row").Length > 0 Then
                '对表体行存货进行取价
                '***************取价接口需要修改         -2008.10.13 -王昕
                strErr = pco.VoucherGetPrice(conn, SAHead, SABody)
                '计算零售单价
                GetPrice2 SABody

                Set domHead = SAHead.cloneNode(True)
                Set domBody = SABody.cloneNode(True)

ToSave:
                strErr = pco.Save(SAHead, SABody, 0, sVoucherID, retDom)
                If strErr <> "" Then
                    '                    MsgBox "审核时生成销售单出错!" & strErr, vbInformation, pustrMsgTitle
                    '                    WriteSABill = False
                    Set frmCheckCredit.myinfo = clsSysSa

                    If SAHead.selectNodes("//信用检查不通过").Length > 0 Then

                        If frmCheckCredit.CheckShow(SAHead, errMsg) = False Then
                            'MsgBox errMsg, vbExclamation, GetString("U8.SA.xsglsql.01.frmbillvouch.00402")  'zh-CN：信用检查
                        End If

                        If frmCheckCredit.bCanceled = False Then
                            'Me.Voucher.headerText("ccrechpname") = frmCheckCredit.cCheckerName
                            'Me.Voucher.headerText("ccrechppass") = frmCheckCredit.cCheckerPass
                            'bCreditCheck = True
                            'AfterCheckCredit = True
                            'Call ButtonClick(sKey, "", bCloseFHSingle)
                        Else
                            'AfterCheckCredit = False
                            'bCreditCheck = True
                        End If


                        Screen.MousePointer = vbDefault

                    Else
                        If SAHead.selectNodes("//最低售价").Length > 0 Then
                            Screen.MousePointer = vbDefault
                            If frmCheckCredit.CheckShow(SAHead, strErr, 3) = False Then
                                MsgBox strErr, vbExclamation, GetString("U8.DZ.JA.Res030")
                                Screen.MousePointer = vbDefault

                            Else
                                If frmCheckCredit.bCanceled = False Then

                                    'If SaveAfterOk Then
                                    Set ele = domHead.selectSingleNode("//z:row")
                                    ele.setAttribute "clowpricepass", clsSysSa.cLowPricePwd
                                    ele.setAttribute "saveafterok", "1"
                                    'End If
                                    'Screen.MousePointer = vbHourglass

                                    Set SAHead = domHead.cloneNode(True)
                                    Set SABody = domBody.cloneNode(True)
                                    GoTo ToSave
                                    Screen.MousePointer = vbDefault
                                Else
                                    Screen.MousePointer = vbDefault
                                    strErr = GetString("U8.DZ.JA.Res940")
                                    GoTo missPass
                                    'conn.RollbackTrans
                                    'Exit Function
                                End If
                            End If
                            'MsgBox strError                            ''最低售价返回处理
                        Else
                            If SAHead.selectNodes("//可用量检查不过").Length > 0 Then
                                If frmCheckCredit.CheckShow(SAHead, errMsg, 1) = False Then
                                    'MsgBox errMsg, vbExclamation, GetString("U8.SA.xsglsql.01.frmbillvouch.00403")  'zh-CN：可用量检查
                                End If

                                If frmCheckCredit.bCanceled = False Then
                                    'Me.Voucher.headerText("bcontinue") = "1"
                                    'Call ButtonClick(sKey, "")
                                Else
                                    'Me.Voucher.headerText("bcontinue") = "0"

                                End If
                            Else
                                '                                If InStr(1, strError, "<", vbTextCompare) <> 0 Then
                                '                                    'ShowErrDom strError, SAHead
                                '                                    If SaveAfterOk Then
                                '                                        .getVoucherDataXML Nothing, domBody
                                '                                        Set ele = SAHead.selectSingleNode("//z:row")
                                '                                        ele.setAttribute "saveafterok", "1"
                                '                                        'GoTo ToSave ' 可以保存
                                '                                    End If
                                '                                Else
                                '                                    MsgBox strError, vbExclamation
                                '                                    If SAHead.selectNodes("//z:row").length = 1 Then
                                '                                        If .headerText(getVoucherCodeName) <> GetHeadItemValue(SAHead, getVoucherCodeName) And strVouchType <> "92" Then
                                '                                            .headerText(getVoucherCodeName) = GetHeadItemValue(SAHead, getVoucherCodeName)
                                '                                        End If
                                '                                    End If
                                '                                End If
                            End If
                        End If
                    End If

                    'MsgBox "审核时生成销售单出错!" & strErr, vbInformation, pustrMsgTitle
missPass:

                    WriteSABill = False
                Else

                    '销售单生成成功后，回写对应的信息-----------------------------------------------------------------------------------------------------------------------------

                    '取出表体id
                    txtSQL = " select csocode " & _
                            " from SO_SOMain " & _
                            " where ID =" & sVoucherID & " "
                    Set RsTemp = conn.Execute(txtSQL)

                    lrows = 0
                    If Not RsTemp.EOF Then
                        oneVouchID = RsTemp("cSOCode")
                        txtSQL = "update " & MainTable & " set " & StriStatus & "=3,downstreamcode='" & RsTemp("cSOCode") & _
                                "'," & StrIntoUser & "='" & login.cUserId & "'," & StrdIntoDate & "='" & login.CurDate & "' where id= '" & r.Fields("ID") & "' and CONVERT(nchar,CONVERT(money,ufts),2)='" & r!ufts & "'"

                        conn.Execute txtSQL, lrows


                    End If
                    '自动审核
                    RsTemp.Close

                    If True Then                           '自动审核
                        '自动审核,判断是否受工作流控制
                        If SAHead.selectSingleNode("//z:row").Attributes.getNamedItem("iswfcontrolled").nodeValue <> "1" Then
                            If pco.VerifyVouch(SAHead, True) <> "" Then
                                MsgBox GetString("U8.DZ.JA.Res950") & errMsg, vbInformation, GetString("U8.DZ.JA.Res030")
                            End If
                        End If
                    End If

                End If
            End If

            If lrows = 0 Then
                If Trim(strErr) = "" Then strErr = GetString("U8.DZ.JA.Res960")
                '                 voucherErrMsg = voucherErrMsg & IIf(voucherErrMsg = "", "生单失败，原因：" & strErr, vbCrLf & "生单失败，原因：" & strErr)
                ReDim varArgs(0)
                varArgs(0) = strErr
                voucherErrMsg = voucherErrMsg & IIf(voucherErrMsg = "", GetStringPara("U8.DZ.JA.Res970", varArgs(0)), vbCrLf & GetStringPara("U8.DZ.JA.Res970", varArgs(0)))

            Else
                vouchID = vouchID & " " & oneVouchID
                oneVouchID = ""
                voucherSuccSize = voucherSuccSize + 1
            End If
        End If
        r.MoveNext
    Next i
    'If Trim(vouchID) <> "" Then
    If rsize = 0 Then voucherErrMsg = GetString("U8.DZ.JA.Res960")
    Screen.MousePointer = vbDefault
    Load FrmMsgBox
    ReDim varArgs(1)
    varArgs(0) = voucherSuccSize
    varArgs(1) = vouchID
    FrmMsgBox.Text1 = GetStringPara("U8.DZ.JA.Res980", varArgs(0), varArgs(1))
    '    FrmMsgBox.Text1 = "成功生成 " & voucherSuccSize & " 张销售订单" & IIf(voucherSuccSize > 0, "，单号 " & vouchID, "") & vbCrLf & voucherErrMsg
    FrmMsgBox.Show 1
    'MsgBox "生成 " & vouchID & " 号销售订单成功!", vbInformation, pustrMsgTitle
    'End If
    Set r = Nothing
    WriteSABill = True
    Exit Function
ErrHandler:
    MsgBox Err.Description, vbInformation, GetString("U8.DZ.JA.Res030")
    WriteSABill = False

End Function

'写采购类单据
Public Function WritePUBill(ByRef oDomHead As DOMDocument, _
                            ByRef oDomBody As DOMDocument, _
                            conn As Object, _
                            login As clsLogin, _
                            VoucherType As String, _
                            sBillType As String) As Boolean

    On Error GoTo ErrHandler

    Dim ele           As IXMLDOMElement, eleList As IXMLDOMNodeList

    Dim sBusType      As String                                 '业务类型:普通采购,直运采购,受托代销

    Dim mVouchCO      As Object

    Dim strSql        As String, rs As New ADODB.Recordset

    Dim POID          As String

    Dim errMsg        As String, strError As String

    Dim i             As Long, j As Long, curID As String

    Dim CGHead        As DOMDocument, CGBody As DOMDocument

    Dim ViewHead      As String, ViewBody As String

    Dim r             As New ADODB.Recordset, rsize As Long

    Dim rst           As New ADODB.Recordset

    Dim rows          As Long

    Dim sXML          As String

    Dim fUnitCost     As Variant, fTaxCost As Variant, bTaxCost As Boolean, fTaxRate As Variant

    Dim iGroupType    As String

    Dim RSTOP         As New Recordset

    Dim lrows         As Long

    Dim voucherErrMsg As String

    Dim vouchID       As String, voucherSuccSize As Integer

    ViewHead = GetViewHead(conn, sBillType)
    ViewBody = GetViewBody(conn, sBillType)
   
    strSql = "select distinct cdeptcode,cDepName from V_HY_LSDG_InputpuAppdata  where  isnull(istats,0) ='未导入' and  id in (" & idtmp & ")"
    Set r = New ADODB.Recordset
    r.Open strSql, conn, 1, 1

    If r.EOF Then
        
        MsgBox "数据不存在或有误，请检查", vbInformation, "提示"
        WritePUBill = False

        Exit Function

    Else

        While Not r.EOF

            sBusType = "普通采购"
            '2008-01-31 初始化采购生单接口
            Set mVouchCO = CreateObject("VoucherCO_PU.clsVoucherCO_PU")
            '   Sub Init(enmVoucherType As VoucherType, [Login As clsLogin], [conn As Connection], [clsInfor As ClsS_Infor], [bPositive As Boolean = True], [sBillType As String], [sBusType As String], [emnUseMode As UseMode])
            mVouchCO.Init VoucherType, login, conn, , True, sBillType, sBusType    'sbilltype=88为单据标示 代表采购订单
            mVouchCO.bOutTrans = True

            '   组织好odomhead后

            Set CGHead = New DOMDocument
            Set CGBody = New DOMDocument

            '写表头----------------------------------------------------------------------------------
            strSql = "select *,'' as editprop from " & ViewHead & " where 1>2"
            rs.Open strSql, conn, adOpenDynamic, adLockOptimistic
            rs.Save CGHead, adPersistXML                       '得到入库单表头DOM结构对象
            rs.Close
            Set oNodes = CGHead.selectSingleNode("//rs:data")
            Set oNode = CGHead.createElement("z:row")
 
            oNode.setAttribute "cdepcode", Null2Something(r!cdeptcode)    '  cDepCode 部门编码  varchar 12  True
        
            oNode.setAttribute "cbustype", sBusType          '  cBusType 业务类型  varchar 8  True
       
            oNode.setAttribute "ivtid", GetVoucherID(conn, sBillType)

            oNode.setAttribute "cdefine1", ""    '表头自定义项
            oNode.setAttribute "cdefine2", ""   '表头自定义项
            oNode.setAttribute "cdefine3", ""    '表头自定义项
            oNode.setAttribute "cdefine4", ""   '表头自定义项
            oNode.setAttribute "cdefine5", ""    '表头自定义项
            oNode.setAttribute "cdefine6", ""   '表头自定义项
            oNode.setAttribute "cdefine7", ""    '表头自定义项
            oNode.setAttribute "cdefine8", ""    '表头自定义项
            oNode.setAttribute "cdefine9", ""    '表头自定义项
            oNode.setAttribute "cdefine10", ""    '表头自定义项
            oNode.setAttribute "cdefine11", ""   '表头自定义项
            oNode.setAttribute "cdefine12", ""    '表头自定义项
            oNode.setAttribute "cdefine13", ""    '表头自定义项
            oNode.setAttribute "cdefine14", ""    '表头自定义项
            oNode.setAttribute "cdefine15", ""    '表头自定义项
            oNode.setAttribute "cdefine16", ""    '表头自定义项

            oNodes.appendChild oNode

            If mVouchCO.GetVoucherNO(CGHead, sBillType, errMsg, POID) = False Then
                WritePUBill = False

                Exit Function

            End If

            '填充表头表体单据编号
            Set ele = CGHead.selectSingleNode("//z:row")
            ele.setAttribute "ccode", POID
            ele.setAttribute "ufts", ""
            ele.setAttribute "ddate", login.CurDate
            ele.setAttribute "cbustype", sBusType

            '根据R!id处理表体----------------------------------------

            strSql = "select *,'' as editprop from " & ViewBody & " where 1>2"
            rs.Open strSql, conn, adOpenDynamic, adLockOptimistic
            rs.Save CGBody, adPersistXML
            rs.Close

            strSql = "select * from V_HY_LSDG_InputpuAppdata  where  isnull(istats,0) ='未导入' and  id in (" & idtmp & ") and cdeptcode='" & r.Fields("cdeptcode") & "'"
            Set rst = New ADODB.Recordset
            rst.Open strSql, conn, 1, 1
            rows = 1

            While Not rst.EOF
                Set oNodes = CGBody.selectSingleNode("//rs:data")
                Set oNode = CGBody.createElement("z:row")
                oNode.setAttribute "cinvcode", Null2Something(rst.Fields("cinvcode"))
                oNode.setAttribute "drequirdate", login.CurDate
                oNode.setAttribute "darrivedate", login.CurDate
               
                oNode.setAttribute "ipertaxrate", 17
                oNode.setAttribute "bTaxCost", 1
                oNode.setAttribute "cexch_name", "人民币"
                oNode.setAttribute "iexchrate", 1
                oNode.setAttribute "ivouchrowno", rows
                oNode.setAttribute "fquantity", CDbl(Format(Null2Something(rst.Fields("iqty")), m_sQuantityFmt))

                oNode.setAttribute "cdefine22", ""    '表体自定义项1
                oNode.setAttribute "cdefine23", ""    '表体自定义项2
                oNode.setAttribute "cdefine24", ""     '表体自定义项3
                oNode.setAttribute "cdefine25", ""    '表体自定义项4
                oNode.setAttribute "cdefine26", ""   '表体自定义项5
                oNode.setAttribute "cdefine27", ""     '表体自定义项6
                oNode.setAttribute "cdefine28", ""     '表体自定义项7
                oNode.setAttribute "cdefine29", ""     '表体自定义项8
                oNode.setAttribute "cdefine30", ""    '表体自定义项9
                oNode.setAttribute "cdefine31", ""     '表体自定义项10
                oNode.setAttribute "cdefine32", ""    '表体自定义项11
                oNode.setAttribute "cdefine33", Null2Something(rst.Fields("id"))     '表体自定义项12
                oNode.setAttribute "cdefine34", ""     '表体自定义项13
                oNode.setAttribute "cdefine35", ""    '表体自定义项14
                oNode.setAttribute "cdefine36", ""    '表体自定义项15
                oNode.setAttribute "cdefine37", ""    '表体自定义项16

                oNode.setAttribute "editprop", "A"
                oNodes.appendChild oNode
                rst.MoveNext
                rows = rows + 1

            Wend
            rst.Close

            '        '非代管采购的业务类型进行最高限价控制
            '        If sBusType <> "代管采购" Then
            '            If Not bGetMPService(sBillType, CGHead, CGBody, conn, login) Then
            '                strError = GetString("U8.DZ.JA.Res1000")
            '                WritePUBill = False
            '                Exit Function
            '            End If
            '        End If
            '2008-01-31 调用采购接口生单
            'Function VoucherSave2(DomHead As DOMDocument, domBody As DOMDocument, VoucherState As Integer, curID) As String
            strError = mVouchCO.VoucherSave2(CGHead, CGBody, 2, curID)

            If strError <> "" Then
                MsgBox strError, vbInformation, GetString("U8.DZ.JA.Res030")
                WritePUBill = False

                Exit Function
             
            End If

            If Trim(strError) = "" Then
          
                vouchID = vouchID & " " & POID
                POID = ""
                voucherSuccSize = voucherSuccSize + 1
                
                 strSql = "delete HY_LSDG_InputpuAppdata  where  id in (select distinct isnull(cDefine33,0)  from PU_AppVouchs where id='" & curID & "') "
                 conn.Execute strSql, lrows
                 
                     strSql = "update HY_LSDG_InputpuAppdatalist set istats =1 where  id in (select distinct isnull(cDefine33,0)  from PU_AppVouchs where id='" & curID & "') "
                 conn.Execute strSql, lrows
                 
            End If
        
            r.MoveNext

        Wend

    End If

    r.Close
    Set r = Nothing

     
    Screen.MousePointer = vbDefault
    Load FrmMsgBox
    ReDim varArgs(1)
    varArgs(0) = voucherSuccSize
    varArgs(1) = vouchID
    ' FrmMsgBox.Text1 = GetStringPara("U8.DZ.JA.Res1020", varArgs(0), varArgs(1))
    FrmMsgBox.Text1 = "成功生成 " & voucherSuccSize & " 张请购单" & IIf(voucherSuccSize > 0, "，单号 " & vouchID, "") & vbCrLf & voucherErrMsg
    FrmMsgBox.Show 1
    'MsgBox "生成 " & vouchID & " 号销售订单成功!", vbInformation, pustrMsgTitle
    'End If
    Set r = Nothing
    WritePUBill = True

    Exit Function

ErrHandler:
    MsgBox Err.Description, vbInformation, GetString("U8.DZ.JA.Res030")
    WritePUBill = False

End Function

'写库存类单据
Public Function WriteSCBill(ByRef oDomHead As DOMDocument, ByRef oDomBody As DOMDocument, conn As Object, login As clsLogin, VoucherType As String, sBillType As String) As Boolean
    On Error GoTo ErrHandler

    Dim pco As Object
    Dim errMsg As String
    Dim domMsg As DOMDocument
    Dim Postion As New DOMDocument                         '货位信息
    Dim sReturnStr As String
    Dim ele As IXMLDOMElement, eleList As IXMLDOMNodeList

    Dim strSql As String, rs As New ADODB.Recordset
    Dim strError As String
    Dim i As Long, j As Long, curID As String
    Dim SCHead As DOMDocument, SCBody As DOMDocument
    Dim ViewHead As String, ViewBody As String
    Dim r As New ADODB.Recordset, rsize As Long
    Dim rows As Long
    Dim lrows As Long
    Dim voucherErrMsg As String
    Dim vouchID As String, voucherSuccSize As Integer
    Dim rdID As String



    Set pco = CreateObject("USERPCO.VoucherCO")
    pco.IniLogin login, errMsg
    strSql = "select distinct moid,cmocode from  " & ViewDetailName
    Set r = New ADODB.Recordset
    r.Open strSql, conn, 1, 1

    If r.EOF Then
        
        MsgBox "数据不存在或有误，请检查", vbInformation, "提示"
        WriteSCBill = False

        Exit Function

    Else

        While Not r.EOF

        '   组织好odomhead
        Set SCHead = New DOMDocument
        Set SCBody = New DOMDocument
        ViewHead = GetViewHead(conn, sBillType)
        ViewBody = GetViewBody(conn, sBillType)

        '写表头----------------------------------------------------------------------------------
        strSql = "select *,'' as editprop from " & ViewHead & " where 1>2"
          Set rs = New ADODB.Recordset
        rs.Open strSql, conn, adOpenDynamic, adLockOptimistic
        rs.Save SCHead, adPersistXML                       '得到入库单表头DOM结构对象
        rs.Close
        Set oNodes = SCHead.selectSingleNode("//rs:data")
        Set oNode = SCHead.createElement("z:row")

        oNode.setAttribute "cwhcode", getAccinformation("QR", "cWhCode", conn) '"01" 'IIf(r!cwhcode & "" = "", "001", r!cwhcode) '"001"    '仓库编号 (必需字段)
        oNode.setAttribute "ddate", login.CurDate                    '单据日期
        oNode.setAttribute "crdcode", getAccinformation("QR", "cRdCode", conn) '"102"                   '收发类别
        oNode.setAttribute "btransflag", "0"               '是否传递
        oNode.setAttribute "cmaker", login.cUserName       '制单人
        oNode.setAttribute "cbustype", "成品入库"              '业务类型
'        oNode.setAttribute "inetlock", "0"                 '现无用
        oNode.setAttribute "brdflag", "1"                  '收发标识
        oNode.setAttribute "cvouchtype", "10"              '单据类型(其他入库单)
        oNode.setAttribute "csource", "生产订单"    '来原来据
        oNode.setAttribute "bpufirst", "0"                 '采购期初标志
        oNode.setAttribute "biafirst", "0"                 '存货期初标志
        oNode.setAttribute "bisstqc", "0"                  '库存期初标志
        oNode.setAttribute "bomfirst", "0"                 '委外起初标志

        oNode.setAttribute "cmemo", ""    '备注
        oNode.setAttribute "iexchrate", 1   '汇率
        oNode.setAttribute "cexch_name", "人民币"    '币种
        oNode.setAttribute "ccode", "0000000001"           '收发单据号
        oNode.setAttribute "iproorderid", r!moid '生产订单主表标识
        oNode.setAttribute "cmpocode", r!cmocode '生产订单编号
        oNode.setAttribute "cdepcode", GetMDeptCode(r!moid) 'Null2Something(r!cdeptcode)    '部门编码
'        oNode.setAttribute "cCusCode", Null2Something(r!cdeptcode)
        

        oNode.setAttribute "cdefine1", ""    '表头自定义项
            oNode.setAttribute "cdefine2", ""   '表头自定义项
            oNode.setAttribute "cdefine3", ""    '表头自定义项
            oNode.setAttribute "cdefine4", ""   '表头自定义项
            oNode.setAttribute "cdefine5", ""    '表头自定义项
            oNode.setAttribute "cdefine6", ""   '表头自定义项
            oNode.setAttribute "cdefine7", ""    '表头自定义项
            oNode.setAttribute "cdefine8", ""    '表头自定义项
            oNode.setAttribute "cdefine9", ""    '表头自定义项
            oNode.setAttribute "cdefine10", ""    '表头自定义项
            oNode.setAttribute "cdefine11", ""   '表头自定义项
            oNode.setAttribute "cdefine12", ""    '表头自定义项
            oNode.setAttribute "cdefine13", ""    '表头自定义项
            oNode.setAttribute "cdefine14", ""    '表头自定义项
            oNode.setAttribute "cdefine15", ""    '表头自定义项
            oNode.setAttribute "cdefine16", ""    '表头自定义项

        oNode.setAttribute "vt_id", GetVoucherID(conn, sBillType)    '单据显示模版号
        oNodes.appendChild oNode
        
         Dim oDomFormat As DOMDocument
     Dim sError As String
    Dim strVoucherNo As String
        
         If GetVoucherNO(conn, SCHead, sBillType, sError, strVoucherNo, , , , False) = True Then
             Set ele = SCHead.selectSingleNode("//z:row")
            ele.setAttribute "ccode", strVoucherNo
         End If
         
        
        '根据R!id处理表体----------------------------------------

        strSql = "select *,'' as editprop from " & ViewBody & " where 1>2"
         Set rs = New ADODB.Recordset
        rs.Open strSql, conn, adOpenDynamic, adLockOptimistic
        rs.Save SCBody, adPersistXML
        rs.Close

         strSql = "select cmocode,modid,imoseq,cinvcode,COUNT(cPN) AS iqty from " & ViewDetailName & " group by cmocode,modid,imoseq,cinvcode "
            Set rst = New ADODB.Recordset
            rst.Open strSql, conn, 1, 1
            rows = 1

            While Not rst.EOF
                  Set oNodes = SCBody.selectSingleNode("//rs:data")
                Set oNode = SCBody.createElement("z:row")
                oNode.setAttribute "cinvcode", Null2Something(rst.Fields("cinvcode"))   '存货编码
'
'                oNode.setAttribute "cinvm_unit", GetNodeAtrVal(ele, "ccomunitcode")    '            主计量
'
'                oNode.setAttribute "cassunit", GetNodeAtrVal(ele, "cunitid")    '            辅计量

                oNode.setAttribute "iquantity", CDbl(Format(Null2Something(rst.Fields("iqty")), m_sQuantityFmt))   '数量

'                If GetNodeAtrVal(ele, "inum") <> "" Then _
'                        oNode.setAttribute "inum", Format(CDbl(GetNodeAtrVal(ele, "inum")), m_sNumFmt)    '    件数
'                If GetNodeAtrVal(ele, "iinvexchrate") <> "" Then _
'                        oNode.setAttribute "iinvexchrate", Format(CDbl(GetNodeAtrVal(ele, "iinvexchrate")), m_iExchRateFmt)    '换算率
'                oNode.setAttribute "cbatch", GetNodeAtrVal(ele, "cBatch")    '批号

                oNode.setAttribute "bcosting", "1"         '是否核算
'                If Val(Format(Null2Something(rst.Fields("price")), m_sPriceFmt)) <> 0 Then _
'                        oNode.setAttribute "iunitcost", Val(Format(Null2Something(rst.Fields("price")), m_sPriceFmt))    '本币单价
               'If GetNodeAtrVal(ele, "inatmoney") <> "" Then
'                        oNode.setAttribute "iprice", Val(Format(Null2Something(rst.Fields("price")), m_sPriceFmt)) * Val(Format(Null2Something(rst.Fields("iqty")), m_sQuantityFmt))    '本币金额
               ' If GetNodeAtrVal(ele, "itaxrate") <> "" _
                        Then
                        oNode.setAttribute "itaxrate", 17    '税率
'
'                oNode.setAttribute "isotype", GetNodeAtrVal(ele, "sotype")    ' 需求跟踪方式
'                oNode.setAttribute "csocode", GetNodeAtrVal(ele, "socode")    ' 需求跟踪耗
'                oNode.setAttribute "cdemandmemo", GetNodeAtrVal(ele, "cdemandmemo")    '需求分类代号说明

'
'                oNode.setAttribute "iexpiratdatecalcu", GetNodeAtrVal(ele, "iexpiratdatecalcu")    '有效期推算方式
'                oNode.setAttribute "cexpirationdate", GetNodeAtrVal(ele, "cexpirationdate")    '有效期至
'                oNode.setAttribute "dexpirationdate", GetNodeAtrVal(ele, "dexpirationdate")    '有效期计算项
'                oNode.setAttribute "dmadedate", GetNodeAtrVal(ele, "dmadedate")    '生产日期
'                oNode.setAttribute "imassdate", GetNodeAtrVal(ele, "imassdate")    '保质期
'                oNode.setAttribute "cmassunit", GetNodeAtrVal(ele, "cmassunit")    '保质期单位
                    
                oNode.setAttribute "imoseq", rst!imoseq    '生产订单行号
                oNode.setAttribute "impoids", rst!modid    '生产订单子表标识
                oNode.setAttribute "cmocode", r!cmocode    '生产订单号
                 oNode.setAttribute "cdefine22", ""    '表体自定义项1
                oNode.setAttribute "cdefine23", ""    '表体自定义项2
                oNode.setAttribute "cdefine24", ""     '表体自定义项3
                oNode.setAttribute "cdefine25", ""    '表体自定义项4
                oNode.setAttribute "cdefine26", ""   '表体自定义项5
                oNode.setAttribute "cdefine27", ""     '表体自定义项6
                oNode.setAttribute "cdefine28", ""     '表体自定义项7
                oNode.setAttribute "cdefine29", ""     '表体自定义项8
                oNode.setAttribute "cdefine30", ""    '表体自定义项9
                oNode.setAttribute "cdefine31", ""     '表体自定义项10
                oNode.setAttribute "cdefine32", ""    '表体自定义项11
'                oNode.setAttribute "cdefine33", Null2Something(rst.Fields("id"))     '表体自定义项12
                oNode.setAttribute "cdefine34", ""     '表体自定义项13
                oNode.setAttribute "cdefine35", ""    '表体自定义项14
                oNode.setAttribute "cdefine36", ""    '表体自定义项15
                oNode.setAttribute "cdefine37", ""    '表体自定义项16

                oNode.setAttribute "ufts", ""
                oNode.setAttribute "editprop", "A"
                oNodes.appendChild oNode
                rst.MoveNext
                rows = rows + 1
                  
            Wend
            rst.Close

        '2008-01-31 调用库存接口生单
        'Insert(sVouchType As String, DomHead, domBody, domPosition, errMsg As String, [cnnFrom As Connection], [VouchId As String], [domMsg As DOMDocument], [bCheck As Boolean = True], [bBeforCheckStock As Boolean = True], [bIsRedVouch As Boolean = False], [sAddedState As String], [bReMote As Boolean = False]) As Boolean
        If pco.Insert(VoucherType, SCHead, SCBody, Postion, errMsg, conn, rdID, domMsg, True, True) = False Then
            If Not (domMsg.selectSingleNode("//z:row") Is Nothing) Then
                frmStockMsg.Message = domMsg
                frmStockMsg.vouchtype = VoucherType
                frmStockMsg.Show vbModal
                If frmStockMsg.Result <> vbYes Then
'                    sReturnStr = GetString("U8.DZ.JA.Res1030")
'                    MsgBox sReturnStr, vbInformation, GetString("U8.DZ.JA.Res030")
                    WriteSCBill = False
                    Exit Function
                End If
                
            Else
                MsgBox errMsg, vbInformation, GetString("U8.DZ.JA.Res030")
                WriteSCBill = False
                Exit Function
            End If
        End If

        If strError <> "" Then
            WriteSCBill = False
            Exit Function
        Else
               Set rs = New ADODB.Recordset
                Set rs = conn.Execute("select ccode from rdrecord10 where id='" & rdID & "'")    '"HY99"
            If Not rs Is Nothing Then
                If Not rs.EOF Then curID = Null2Something(rs(0))
            End If
          
                vouchID = vouchID & " " & curID
                
                voucherSuccSize = voucherSuccSize + 1
                
'                 strSql = "delete HY_LSDG_InputpuAppdata  where  id in (select distinct isnull(cDefine33,0)  from rdrecords32 where id='" & rdID & "') "
'                 conn.Execute strSql, lrows
                 
                     strSql = "update " & ViewDetailName & " set ccode ='" & curID & "' "
                 conn.Execute strSql, lrows
                curID = ""
        End If
       
        r.MoveNext
    
        Wend

    End If

    r.Close
    Set r = Nothing

   
    Screen.MousePointer = vbDefault
    Load FrmMsgBox
    ReDim varArgs(1)
    varArgs(0) = voucherSuccSize
    varArgs(1) = vouchID
'    FrmMsgBox.Text1 = GetStringPara("U8.DZ.JA.Res1050", varArgs(0), varArgs(1))
      FrmMsgBox.Text1 = "成功生成 " & voucherSuccSize & " 张库存单据" & IIf(voucherSuccSize > 0, "，单号 " & vouchID, "") & vbCrLf & voucherErrMsg
    FrmMsgBox.Show 1
    'MsgBox "生成 " & vouchID & " 号销售订单成功!", vbInformation, pustrMsgTitle
    'End If
    Set r = Nothing
    WriteSCBill = True

    Exit Function

ErrHandler:
    If Err.Description <> "" Then MsgBox Err.Description, vbInformation, GetString("U8.DZ.JA.Res030")
    WriteSCBill = False

End Function



'写应付单据
Public Function WriteAPBill(ByRef oDomHead As DOMDocument, ByRef oDomBody As DOMDocument, conn As Object, login As clsLogin, Optional VoucherType As String, Optional sBillType As String) As Boolean
    On Error GoTo ErrHandler

    Dim APBO As Object
    Dim errMsg As String
    Dim ele As IXMLDOMElement, eleList As IXMLDOMNodeList
    Dim strSql As String, rs As New ADODB.Recordset
    Dim strError As String
    Dim i As Long, j As Long, curID As String
    Dim APhead As DOMDocument, APBody As DOMDocument
    Dim ViewHead As String, ViewBody As String
    Dim r As New ADODB.Recordset, rsize As Long
    Dim rows As Long
    Dim lrows As Long
    Dim voucherErrMsg As String
    Dim vouchID As String, voucherSuccSize As Integer
    Dim ApID As String
    Dim sumQ As Double, sumM As Double, sumM_f As Double


    Set APBO = CreateObject("UFAPBO.clsApvouch")
    APBO.Init login, conn, VoucherType

    r.Open oDomHead
    rsize = r.RecordCount

    For i = 1 To rsize

        '   组织好odomhead
        Set APhead = New DOMDocument
        Set APBody = New DOMDocument
        ViewHead = GetViewHead(conn, sBillType)
        ViewBody = GetViewBody(conn, sBillType)

        '写表头----------------------------------------------------------------------------------
        strSql = "select *,'' as editprop from " & ViewHead & " where 1>2"
        rs.Open strSql, conn, adOpenDynamic, adLockOptimistic
        rs.Save APhead, adPersistXML                       '得到入库单表头DOM结构对象
        rs.Close
        Set oNodes = APhead.selectSingleNode("//rs:data")
        Set oNode = APhead.createElement("z:row")

        oNode.setAttribute "cVouchType", "P0"              '单据类型 -Ap_VouchType表   0为数字零
        '            oNode.setAttribute "cVouchID", "0000000001"                  '对应单据号
        oNode.setAttribute "cVouchID1", Null2Something(r!cCode)    '对应单据类型
        oNode.setAttribute "cCoVouchType", gstrCardNumber  '"HY99"                  '对应单据号
        oNode.setAttribute "dVouchDate", Date              '单据日期
        oNode.setAttribute "cDeptCode", Null2Something(r!cDepcode)    '部门编码
        oNode.setAttribute "cPerson", Null2Something(r!cpersoncode)    '业务员编码
        oNode.setAttribute "cCode", ""                     '科目编码
        oNode.setAttribute "cexch_name", Null2Something(r!cexch_name)    '币种
        oNode.setAttribute "iExchRate", Null2Something(r!iExchRate)    '汇率
        oNode.setAttribute "cDigest", GetString("U8.DZ.JA.Res990") & Time    '摘要
        oNode.setAttribute "cPayCode", ""                  '付款条件
        oNode.setAttribute "cOperator", login.cUserName    '录入人
        oNode.setAttribute "bStartFlag", "0"               '期初标志


        oNode.setAttribute "cFlag", VoucherType            '收发标识
        oNode.setAttribute "bd_c", IIf(VoucherType = "AP", "0", "1")    '借贷方向
        oNode.setAttribute "cDwCode", IIf(VoucherType = "AP", Null2Something(r!cvencode), Null2Something(r!cCusCode))    '单位

        sumQ = 0: sumM = 0: sumM_f = 0
        Set eleList = oDomBody.selectNodes("//z:row[@" + HeadPKFld + "='" & r!id & "']")
        For Each ele In eleList
            If GetNodeAtrVal(ele, "iquantity") <> "" Then sumQ = sumQ + CDbl(GetNodeAtrVal(ele, "iquantity"))
            If GetNodeAtrVal(ele, "inatmoney") <> "" Then sumM = sumM + CDbl(GetNodeAtrVal(ele, "inatsum"))
            If GetNodeAtrVal(ele, "isum") <> "" Then sumM_f = sumM_f + CDbl(GetNodeAtrVal(ele, "isum"))
        Next
        oNode.setAttribute "iAmount_s", Format(CDbl(sumQ), m_sQuantityFmt)    '数量
        oNode.setAttribute "iAmount", Format(CDbl(sumM), m_sPriceFmt)    '本币金额
        oNode.setAttribute "iAmount_f", Format(CDbl(sumM_f), m_sPriceFmt)    '原币金额

        oNode.setAttribute "cDefine1", Null2Something(r!cDefine1)    '表头自定义项
        oNode.setAttribute "cDefine2", Null2Something(r!cDefine2)    '表头自定义项
        oNode.setAttribute "cDefine3", Null2Something(r!cDefine3)    '表头自定义项
        oNode.setAttribute "cDefine4", Null2Something(r!cDefine4)    '表头自定义项
        oNode.setAttribute "cDefine5", Null2Something(r!cDefine5)    '表头自定义项
        oNode.setAttribute "cDefine6", Null2Something(r!cDefine6)    '表头自定义项
        oNode.setAttribute "cDefine7", Null2Something(r!cdefine7)    '表头自定义项
        oNode.setAttribute "cDefine8", Null2Something(r!cDefine8)    '表头自定义项
        oNode.setAttribute "cDefine9", Null2Something(r!cDefine9)    '表头自定义项
        oNode.setAttribute "cDefine10", Null2Something(r!cDefine10)    '表头自定义项
        oNode.setAttribute "cDefine11", Null2Something(r!cDefine11)    '表头自定义项
        oNode.setAttribute "cDefine12", Null2Something(r!cDefine12)    '表头自定义项
        oNode.setAttribute "cDefine13", Null2Something(r!cDefine13)    '表头自定义项
        oNode.setAttribute "cDefine14", Null2Something(r!cDefine14)    '表头自定义项
        oNode.setAttribute "cDefine15", Null2Something(r!cDefine15)    '表头自定义项
        oNode.setAttribute "cDefine16", Null2Something(r!cdefine16)    '表头自定义项

        oNode.setAttribute "vt_id", GetVoucherID(conn, sBillType)    '单据显示模版号
        oNodes.appendChild oNode


        'Function GetVouchID(cType As String, oDom As DOMDocument, xmlErrMsg As String) As String
        '根据R!id处理表体----------------------------------------

        strSql = "select *,'' as editprop from " & ViewBody & " where 1>2"
        rs.Open strSql, conn, adOpenDynamic, adLockOptimistic
        rs.Save APBody, adPersistXML
        rs.Close

        Set eleList = oDomBody.selectNodes("//z:row[@" + HeadPKFld + "='" & r!id & "']")
        rows = eleList.Length
        If rows > 0 Then
            Set oNodes = APBody.selectSingleNode("//rs:data")
            For Each ele In eleList
                Set oNode = APBody.createElement("z:row")

                oNode.setAttribute "cinvm_unit", GetNodeAtrVal(ele, "citemcode")    '项目编码
                oNode.setAttribute "cItemCode", GetNodeAtrVal(ele, "citem_class")    '   项目大类编码
                oNode.setAttribute "cItemName", GetNodeAtrVal(ele, "citemname")    '项目名称
                oNode.setAttribute "cPerson", Null2Something(r!cpersoncode)    '业务员编码
                oNode.setAttribute "cDeptCode", Null2Something((r!cDepcode))    '部门编码
                oNode.setAttribute "cDwCode", IIf(VoucherType = "AP", Null2Something(r!cvencode), Null2Something(r!cCusCode))    '单位
                oNode.setAttribute "iAmt_s", Format(CDbl(GetNodeAtrVal(ele, "iquantity")), m_sQuantityFmt)    '数量
                oNode.setAttribute "iTaxRate", Format(CDbl(GetNodeAtrVal(ele, "itaxrate")), m_iRateFmt)    '税率
                oNode.setAttribute "iTax", Format(CDbl(GetNodeAtrVal(ele, "itax")), m_sPriceFmt)    ' 税额
                oNode.setAttribute "iNatTax", Format(CDbl(GetNodeAtrVal(ele, "inattax")), m_sPriceFmt)    ' 本币税额
                oNode.setAttribute "cexch_name", Null2Something(r!cexch_name)    '币种
                oNode.setAttribute "iExchRate", Format(CDbl(Null2Something(r!iExchRate, 0)), m_iRateFmt)    '汇率
                oNode.setAttribute "bd_c", IIf(VoucherType = "AP", "1", "0")    '借贷方向
                oNode.setAttribute "iAmount", Format(CDbl(GetNodeAtrVal(ele, "inatsum")), m_sPriceFmt)    '本币金额
                oNode.setAttribute "iAmount_f", Format(CDbl(GetNodeAtrVal(ele, "isum")), m_sPriceFmt)    '原币金额

                oNode.setAttribute "cDefine22", GetNodeAtrVal(ele, "cDefine22")    '表体自定义项
                oNode.setAttribute "cDefine23", GetNodeAtrVal(ele, "cDefine23")    '表体自定义项
                oNode.setAttribute "cDefine24", GetNodeAtrVal(ele, "cDefine24")    '表体自定义项
                oNode.setAttribute "cDefine25", GetNodeAtrVal(ele, "cDefine25")    '表体自定义项
                oNode.setAttribute "cDefine26", GetNodeAtrVal(ele, "cDefine26")    '表体自定义项
                oNode.setAttribute "cDefine27", GetNodeAtrVal(ele, "cDefine27")    '表体自定义项
                oNode.setAttribute "cDefine28", GetNodeAtrVal(ele, "cDefine28")    '表体自定义项
                oNode.setAttribute "cDefine29", GetNodeAtrVal(ele, "cDefine29")    '表体自定义项
                oNode.setAttribute "cDefine30", GetNodeAtrVal(ele, "cDefine30")    '表体自定义项
                oNode.setAttribute "cDefine31", GetNodeAtrVal(ele, "cDefine31")    '表体自定义项
                oNode.setAttribute "cDefine32", GetNodeAtrVal(ele, "cDefine32")    '表体自定义项
                oNode.setAttribute "cDefine33", GetNodeAtrVal(ele, "cDefine33")    '表体自定义项
                oNode.setAttribute "cDefine34", GetNodeAtrVal(ele, "cDefine34")    '表体自定义项
                oNode.setAttribute "cDefine35", GetNodeAtrVal(ele, "cDefine35")    '表体自定义项
                oNode.setAttribute "cDefine36", GetNodeAtrVal(ele, "cDefine36")    '表体自定义项
                oNode.setAttribute "cDefine37", GetNodeAtrVal(ele, "cDefine37")    '表体自定义项

                oNode.setAttribute "editprop", "A"
                oNodes.appendChild oNode
            Next
        End If


        If APBO.SaveVouch(APhead, APBody, errMsg) = False Then
            sReturnStr = GetString("U8.DZ.JA.Res1060") & errMsg
            MsgBox sReturnStr, vbInformation, GetString("U8.DZ.JA.Res030")
            WriteAPBill = False
            Exit Function
        End If

        If strError <> "" Then
            WriteAPBill = False
            Exit Function
        Else
            lrows = 0
            Set rs = conn.Execute("select clink from ap_vouch where cVouchID1='" & r!cCode & "'and cCoVouchType='" & gstrCardNumber & "'")    '"HY99"
            If Not rs Is Nothing Then
                If Not rs.EOF Then ApID = Null2Something(rs(0))
            End If
            strSql = "update " & MainTable & " set " & StriStatus & "=3,downstreamcode='" & ApID & _
                    "'," & StrIntoUser & "='" & login.cUserId & "'," & StrdIntoDate & "='" & login.CurDate & "' where id= '" & r.Fields("ID") & "' and CONVERT(nchar,CONVERT(money,ufts),2)='" & r!ufts & "'"
            conn.Execute strSql, lrows

        End If
        If lrows = 0 Then
            If Trim(strError) = "" Then strError = GetString("U8.DZ.JA.Res960")
            ReDim varArgs(0)
            varArgs(0) = strErr
            voucherErrMsg = voucherErrMsg & IIf(voucherErrMsg = "", GetStringPara("U8.DZ.JA.Res970", varArgs(0)), vbCrLf & GetStringPara("U8.DZ.JA.Res970", varArgs(0)))

            '            voucherErrMsg = voucherErrMsg & IIf(voucherErrMsg = "", "生单失败，原因：" & strError, vbCrLf & "生单失败，原因：" & strError)
        Else
            vouchID = vouchID & " " & ApID
            ApID = ""
            voucherSuccSize = voucherSuccSize + 1
        End If

        r.MoveNext
    Next

    If rsize = 0 Then voucherErrMsg = GetString("U8.DZ.JA.Res960")
    Screen.MousePointer = vbDefault
    Load FrmMsgBox
    ReDim varArgs(1)
    varArgs(0) = voucherSuccSize
    varArgs(1) = vouchID
    FrmMsgBox.Text1 = GetStringPara("U8.DZ.JA.Res1050", varArgs(0), varArgs(1))
    ' FrmMsgBox.Text1 = "成功生成 " & voucherSuccSize & " 张应收应付单据" & IIf(voucherSuccSize > 0, "，单号 " & vouchID, "") & vbCrLf & voucherErrMsg
    FrmMsgBox.Show 1

    Set r = Nothing
    WriteAPBill = True

    Exit Function

ErrHandler:
    If Err.Description <> "" Then MsgBox Err.Description, vbInformation, GetString("U8.DZ.JA.Res030")
    WriteAPBill = False

End Function


'参照生单处理
Public Function ProcessDatapro(ByRef Voucher As Object)

    On Error GoTo ErrHandler:

    Dim retvalue As Variant
    Dim referpara As UAPVoucherControl85.ReferParameter
    Dim eleline As IXMLDOMElement
    Dim echeck As Long
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim i As Integer

    '提供两种处理模式，表头使用recordset处理，表体使用xml解析处理
    Dim rshead As New ADODB.Recordset
    rshead.Open gDomReferHead
    If Not rshead.EOF And Not rshead.BOF Then
        If rshead("selcol") = "Y" Then

            Voucher.headerText("acccode") = Null2Something(rshead("acccode"))    '
            Voucher.headerText("accname") = Null2Something(rshead("accname"))    '
 
            Voucher.headerText("cmemo") = Null2Something(rshead("cmemo"))    '
            Voucher.headerText("engcode") = Null2Something(rshead("cname"))    '
            Voucher.headerText("engname") = Null2Something(rshead("ccode"))
'            '

            Voucher.headerText("ecustcode") = Null2Something(rshead("ecustcode"))
            Voucher.headerText("cCusAbbName") = Null2Something(rshead("cCusAbbName"))
            Voucher.headerText("custcontacta") = Null2Something(rshead("custcontacta"))
            Voucher.headerText("custcontactb") = Null2Something(rshead("custcontactb"))
            Voucher.headerText("contacta") = Null2Something(rshead("contacta"))
            Voucher.headerText("contactb") = Null2Something(rshead("contactb"))
            Voucher.headerText("engdescripta") = Null2Something(rshead("prodescriptiona"))
            Voucher.headerText("engdescriptb") = Null2Something(rshead("prodescriptionb"))
            Voucher.headerText("engdescriptc") = Null2Something(rshead("prodescriptionc"))
            Voucher.headerText("iStatus") = 1   '
            Voucher.headerText("cMaker") = g_oLogin.cUserName   '
            Voucher.headerText("dmDate") = Format(g_oLogin.CurDate, "yyyy-mm-dd")
            Voucher.headerText("ddate") = Format(g_oLogin.CurDate, "yyyy-mm-dd")  '

            For i = 1 To 16

                Voucher.headerText("cdefine" & i) = Null2Something(rshead("cdefine" & i))    '表头自定义项

            Next

            
        End If
    End If


     

    Voucher.RemoveEmptyRow
    Set rshead = Nothing
    Exit Function

ErrHandler:
    Set rshead = Nothing
    MsgBox Err.Description, vbInformation, GetString("U8.DZ.JA.Res030")

End Function


'其中：发货地址、联系人、电话、编码均由客户名称档案自动带出。
Private Sub setallinforbycus(Voucher As Object, Index As Variant, retvalue As String, _
                             bChanged As UAPVoucherControl85.CheckRet, _
                             referpara As UAPVoucherControl85.ReferParameter)
    On Error GoTo Err_Handler
    If Voucher.headerText("cType") <> "客户" Or Voucher.headerText("bObjectCode") = "" Then Exit Sub

    Dim rs As New ADODB.Recordset
    Dim strSql As String

    strSql = " SELECT cusadd.caddcode,cusadd.cdeliveradd, " & vbCrLf & _
            " ccontactname=case isnull(crm.ccontactname,'') when '' then t1.cCusPerson else crm.ccontactname end , " & vbCrLf & _
            " cmobilephone=case isnull(crm.cmobilephone,'') when '' then t1.cCusPhone else crm.cmobilephone end , " & vbCrLf & _
            " cofficephone=case isnull(crm.cofficephone,'') when '' then t1.cCusHand else crm.cofficephone end , " & vbCrLf & _
            " cZipcode=case isnull(crm.cZipcode,'') when '' then t1.cCusPostCode else crm.cZipcode end " & vbCrLf & _
            " from customer t1 " & vbCrLf & _
            " left join ShippingChoice on t1.cCusOType=ShippingChoice.cSCCode " & vbCrLf & _
            " left join cusdeliveradd cusadd on (cusadd.ccuscode = t1.ccuscode and cusadd.bdefault=1) " & vbCrLf & _
            " left join crm_contact crm on (crm.ccontactcode=cusadd.clinkperson) " & vbCrLf & _
            " WHERE t1.cCusCode='" & Voucher.headerText("bObjectCode") & "" & "'"
    rs.Open strSql, g_Conn, 1, 1
    If Not rs.EOF Then
        '  Voucher.headerText("caddcode") = rs.Fields("caddcode") & ""
        '  Voucher.headerText("cshipaddress") = rs.Fields("cdeliveradd") & ""
        '   Voucher.headerText("cZipcode") = rs.Fields("cZipcode") & ""
        Voucher.headerText("cContactperson") = rs.Fields("ccontactname") & ""
        Voucher.headerText("cContactWay") = rs.Fields("cmobilephone") & _
        IIf(IsNull(rs.Fields("cofficephone")), "", "   " & rs.Fields("cofficephone"))
    End If

    rs.Close
    Set rs = Nothing
    Exit Sub
Err_Handler:
    rs.Close
    Set rs = Nothing
    '    CheckCustomer = False
    MsgBox Err.Description, vbInformation, GetString("U8.DZ.JA.Res030")
End Sub


Public Sub NotifySrvSend(ByVal CardNum As String, NumType As String, voucherID As String, login As Object, Optional CN As Connection)
    
    On Error Resume Next
    If NotifySend Is Nothing Then
        Set NotifySend = CreateObject("UFIDA.U8.BizNotify.NotifyService")
    End If

    Call NotifySend.send(CardNum, NumType, voucherID, login)

End Sub

'890增加表体一次性取现存量功能hucl
Public Sub ShowBodyStockAll(ByRef Voucher As UAPVoucherControl85.ctlVoucher)
    Dim bShowIt As Boolean
    Dim sInvCode As String
    
    bShowIt = Voucher.ItemState("ipresentnum", sibody).showIt Or _
            Voucher.ItemState("ipresent", sibody).showIt Or _
            Voucher.ItemState("iavaquantity", sibody).showIt Or _
            Voucher.ItemState("iavanum", sibody).showIt
    
    If Voucher.BodyRows > 0 And bShowIt Then
        For i = 1 To Voucher.BodyRows
            sInvCode = Voucher.bodyText(i, "cinvcode")
            If sInvCode <> "" Then
                Call ShowBodyStock(Voucher, sInvCode, i)
            End If
        Next
    End If
End Sub


'刷新表体现存量和可用量
Public Sub ShowBodyStock(ByRef Voucher As UAPVoucherControl85.ctlVoucher, ByVal sInvCode As String, ByVal nRow As Long)
    Dim bShowIt         As Boolean
    Dim iquantity       As Double
    Dim iNum            As Double
    Dim iAvaQuantity       As Double
    Dim iAvaNum            As Double
    Dim bLoad As Boolean
    Dim sPosition As String
    Dim sWhCode As String
    Dim sError          As String
    
    Dim ClsAccount As USERPVO.Account
    Dim clsStockCo As USERPCO.StockCO
    Dim tmpInventory As USERPVO.Inventory
    Dim ClsInventoryCO As USERPCO.InventoryCO
    
    sWhCode = Voucher.bodyText(nRow, "cwhcode")
    sPosition = Voucher.bodyText(nRow, "cposition")
    bShow = Voucher.ItemState("ipresentnum", sibody).showIt Or _
            Voucher.ItemState("ipresent", sibody).showIt Or _
            Voucher.ItemState("iavaquantity", sibody).showIt Or _
            Voucher.ItemState("iavanum", sibody).showIt
            
    If bShow Then
        Set ClsAccount = mologin.Account
        Set clsStockCo = New USERPCO.StockCO
        Set tmpInventory = New USERPVO.Inventory
        Set ClsInventoryCO = New USERPCO.InventoryCO
        
        ClsInventoryCO.login = mologin
        clsStockCo.login = mologin
        Call ClsInventoryCO.Load(sInvCode, tmpInventory, sError, , sWhCode)
    Else
        Exit Sub
    End If
    
    With Voucher
        '现存量
        If .ItemState("ipresentnum", sibody).showIt Or .ItemState("ipresent", sibody).showIt Then
            If sPosition = "" Then
    
                clsStockCo.GetStockQTYandAvaQty mologin.Account.ControlFormula, iquantity, iNum, iAvaQuantity, iAvaNum, sError, sWhCode, sInvCode, .bodyText(nRow, "cfree1"), .bodyText(nRow, "cfree2"), .bodyText(nRow, "cfree3"), .bodyText(nRow, "cfree4"), .bodyText(nRow, "cfree5"), .bodyText(nRow, "cfree6"), .bodyText(nRow, "cfree7"), .bodyText(nRow, "cfree8"), .bodyText(nRow, "cfree9"), .bodyText(nRow, "cfree10"), .bodyText(nRow, "cbatch"), , , , , True
    
                If tmpInventory.GroupType = 2 Then
                    .bodyText(nRow, "ipresentnum") = Format(iNum, ClsAccount.FormatNumDecString)
                ElseIf tmpInventory.GroupType = 1 And val(.bodyText(nRow, "iinvexchrate")) <> 0 Then
                    .bodyText(nRow, "ipresentnum") = Format(iquantity / .bodyText(nRow, "iinvexchrate"), ClsAccount.FormatNumDecString)
                End If
                .bodyText(nRow, "ipresent") = Format(iquantity, ClsAccount.FormatQuanDecString)
                If tmpInventory.GroupType = 2 Then
                    .bodyText(nRow, "iavanum") = Format(iAvaNum, ClsAccount.FormatNumDecString)
                ElseIf tmpInventory.GroupType = 1 And val(.bodyText(nRow, "iinvexchrate")) <> 0 Then
                    .bodyText(nRow, "iavanum") = Format(iAvaQuantity / .bodyText(nRow, "iinvexchrate"), ClsAccount.FormatNumDecString)
                End If
                .bodyText(nRow, "iavaquantity") = Format(iAvaQuantity, ClsAccount.FormatQuanDecString)
                Exit Sub
            Else
                If GetPosStock(iquantity, iNum, sInvCode, _
                        .bodyText(nRow, "cfree1"), .bodyText(nRow, "cfree2"), .bodyText(nRow, "cfree3"), _
                        .bodyText(nRow, "cfree4"), .bodyText(nRow, "cfree5"), .bodyText(nRow, "cfree6"), _
                        .bodyText(nRow, "cfree7"), .bodyText(nRow, "cfree8"), .bodyText(nRow, "cfree9"), _
                        .bodyText(nRow, "cfree10"), .bodyText(nRow, "cbatch"), sPosition, sWhCode) Then
                        
                End If
            End If
            
            If tmpInventory.GroupType = 2 Then
                .bodyText(nRow, "ipresentnum") = Format(iNum, ClsAccount.FormatNumDecString)
            ElseIf tmpInventory.GroupType = 1 And val(.bodyText(nRow, "iinvexchrate")) <> 0 Then
                .bodyText(nRow, "ipresentnum") = Format(iquantity / .bodyText(nRow, "iinvexchrate"), ClsAccount.FormatNumDecString)
            End If
            .bodyText(nRow, "ipresent") = Format(iquantity, ClsAccount.FormatQuanDecString)
        End If
        
        '可用量
        If (.ItemState("iavaquantity", sibody).showIt Or .ItemState("iavanum", sibody).showIt) Then
            iNum = 0
            iquantity = 0
            
            If sPosition = "" Then
                clsStockCo.ControlStock mologin.Account.ControlFormula, iquantity, iNum, sError, sWhCode, sInvCode, .bodyText(nRow, "cfree1"), .bodyText(nRow, "cfree2"), .bodyText(nRow, "cfree3"), .bodyText(nRow, "cfree4"), .bodyText(nRow, "cfree5"), .bodyText(nRow, "cfree6"), .bodyText(nRow, "cfree7"), .bodyText(nRow, "cfree8"), .bodyText(nRow, "cfree9"), .bodyText(nRow, "cfree10"), .bodyText(nRow, "cbatch"), , , , , True
            Else
                If .ItemState("ipresentnum", sibody).showIt Or .ItemState("ipresent", sibody).showIt Then
                    .bodyText(nRow, "iavaquantity") = Format(.bodyText(nRow, "ipresent"), ClsAccount.FormatQuanDecString)
                    .bodyText(nRow, "iavanum") = Format(.bodyText(nRow, "ipresentnum"), ClsAccount.FormatNumDecString)
                    Exit Sub
                Else
                    If GetPosStock(iquantity, iNum, sInvCode, _
                            .bodyText(nRow, "cfree1"), .bodyText(nRow, "cfree2"), .bodyText(nRow, "cfree3"), _
                            .bodyText(nRow, "cfree4"), .bodyText(nRow, "cfree5"), .bodyText(nRow, "cfree6"), _
                            .bodyText(nRow, "cfree7"), .bodyText(nRow, "cfree8"), .bodyText(nRow, "cfree9"), _
                            .bodyText(nRow, "cfree10"), .bodyText(nRow, "cbatch"), sPosition, sWhCode) Then
                            
                    End If
                End If
            End If
            
            If tmpInventory.GroupType = 2 Then
                .bodyText(nRow, "iavanum") = Format(iNum, ClsAccount.FormatNumDecString)
            ElseIf tmpInventory.GroupType = 1 And val(.bodyText(nRow, "iinvexchrate")) <> 0 Then
                .bodyText(nRow, "iavanum") = Format(iquantity / .bodyText(nRow, "iinvexchrate"), ClsAccount.FormatNumDecString)
            End If
            .bodyText(nRow, "iavaquantity") = Format(iquantity, ClsAccount.FormatQuanDecString)
        End If
    End With
End Sub

'刷新表头和表体现存量和可用量
Public Sub ShowStock(ByRef Voucher As UAPVoucherControl85.ctlVoucher, ByVal sInvCode As String, ByVal nRow As Long)
    On Error Resume Next
    '现存量
    Dim sError As String
    Dim sPosition As String
    Dim sWhCode As String
    Dim iquantity As Double
    Dim iNum As Double
    Dim bShow As Boolean
    
    Dim ClsAccount As USERPVO.Account
    Dim clsStockCo As USERPCO.StockCO
    Dim tmpInventory As USERPVO.Inventory
    Dim ClsInventoryCO As USERPCO.InventoryCO
    
    sWhCode = Voucher.bodyText(nRow, "cwhcode")
    sPosition = Voucher.bodyText(nRow, "cposition")
    bShow = Voucher.ItemState("ipresentnum", siHeader).showIt Or _
            Voucher.ItemState("ipresent", siHeader).showIt Or _
            Voucher.ItemState("iavaquantity", siHeader).showIt Or _
            Voucher.ItemState("iavanum", siHeader).showIt Or _
            Voucher.ItemState("ipresentnum", sibody).showIt Or _
            Voucher.ItemState("ipresent", sibody).showIt Or _
            Voucher.ItemState("iavaquantity", sibody).showIt Or _
            Voucher.ItemState("iavanum", sibody).showIt
            
    If bShow Then
        Set ClsAccount = mologin.Account
        Set clsStockCo = New USERPCO.StockCO
        Set tmpInventory = New USERPVO.Inventory
        Set ClsInventoryCO = New USERPCO.InventoryCO
        
        ClsInventoryCO.login = mologin
        clsStockCo.login = mologin
        Call ClsInventoryCO.Load(sInvCode, tmpInventory, sError, , sWhCode)
    End If
    
    With Voucher
   
        If .ItemState("ipresentnum", siHeader).showIt Or .ItemState("ipresent", siHeader).showIt Or _
           .ItemState("ipresentnum", sibody).showIt Or .ItemState("ipresent", sibody).showIt Then
            If sPosition = "" Then
                clsStockCo.CurrentStock iquantity, iNum, sError, sWhCode, .bodyText(nRow, "cinvcode"), _
                .bodyText(nRow, "cfree1"), .bodyText(nRow, "cfree2"), .bodyText(nRow, "cfree3"), _
                    .bodyText(nRow, "cfree4"), .bodyText(nRow, "cfree5"), .bodyText(nRow, "cfree6"), _
                    .bodyText(nRow, "cfree7"), .bodyText(nRow, "cfree8"), .bodyText(nRow, "cfree9"), _
                    .bodyText(nRow, "cfree10"), .bodyText(nRow, "cbatch"), , , , , True
            Else
                If GetPosStock(iquantity, iNum, .bodyText(nRow, "cinvcode"), _
                        .bodyText(nRow, "cfree1"), .bodyText(nRow, "cfree2"), .bodyText(nRow, "cfree3"), _
                        .bodyText(nRow, "cfree4"), .bodyText(nRow, "cfree5"), .bodyText(nRow, "cfree6"), _
                        .bodyText(nRow, "cfree7"), .bodyText(nRow, "cfree8"), .bodyText(nRow, "cfree9"), _
                        .bodyText(nRow, "cfree10"), .bodyText(nRow, "cbatch"), sPosition, sWhCode) Then
                End If
            End If
            
            If tmpInventory.GroupType = 2 Then
                .headerText("ipresentnum") = Format(iNum, ClsAccount.FormatNumDecString)
                .bodyText(nRow, "ipresentnum") = Format(iNum, ClsAccount.FormatNumDecString)
            ElseIf tmpInventory.GroupType = 1 And val(.bodyText(.row, "iinvexchrate")) <> 0 Then
                .headerText("ipresentnum") = Format(iquantity / .bodyText(.row, "iinvexchrate"), ClsAccount.FormatNumDecString)
                .bodyText(nRow, "ipresentnum") = Format(iquantity / .bodyText(nRow, "iinvexchrate"), ClsAccount.FormatNumDecString)
            Else
                .headerText("ipresentnum") = ""
            End If
            .headerText("ipresent") = Format(iquantity, ClsAccount.FormatQuanDecString)
            .bodyText(nRow, "ipresent") = Format(iquantity, ClsAccount.FormatQuanDecString)
        End If
        '可用量
        iNum = 0
        iquantity = 0
    
        If .ItemState("iavaquantity", siHeader).showIt Or .ItemState("iavanum", siHeader).showIt Or _
            .ItemState("iavaquantity", sibody).showIt Or .ItemState("iavanum", sibody).showIt Then
            If sPosition = "" Then
                clsStockCo.ControlStock mologin.Account.ControlFormula, iquantity, iNum, sError, sWhCode, .bodyText(nRow, "cinvcode"), _
                    .bodyText(nRow, "cfree1"), .bodyText(nRow, "cfree2"), .bodyText(nRow, "cfree3"), _
                    .bodyText(nRow, "cfree4"), .bodyText(nRow, "cfree5"), .bodyText(nRow, "cfree6"), _
                    .bodyText(nRow, "cfree7"), .bodyText(nRow, "cfree8"), .bodyText(nRow, "cfree9"), _
                    .bodyText(nRow, "cfree10"), .bodyText(nRow, "cbatch"), , , , , True
            Else
                If .ItemState("ipresentnum", siHeader).showIt Or .ItemState("ipresent", siHeader).showIt Then
                    .headerText("iavaquantity") = Format(.headerText("ipresent"), ClsAccount.FormatQuanDecString)
                    .headerText("iavanum") = Format(.headerText("ipresentnum"), ClsAccount.FormatNumDecString)
                    .bodyText(nRow, "iavaquantity") = Format(.headerText("ipresent"), ClsAccount.FormatQuanDecString)
                    .bodyText(nRow, "iavanum") = Format(.headerText("ipresentnum"), ClsAccount.FormatNumDecString)
                    Exit Sub
                ElseIf .ItemState("ipresentnum", sibody).showIt Or .ItemState("ipresent", sibody).showIt Then
                    .headerText("iavaquantity") = Format(.bodyText(nRow, "ipresent"), ClsAccount.FormatQuanDecString)
                    .headerText("iavanum") = Format(.bodyText(nRow, "ipresentnum"), ClsAccount.FormatNumDecString)
                    .bodyText(nRow, "iavaquantity") = Format(.bodyText(nRow, "ipresent"), ClsAccount.FormatQuanDecString)
                    .bodyText(nRow, "iavanum") = Format(.bodyText(nRow, "ipresentnum"), ClsAccount.FormatNumDecString)
                    Exit Sub
                Else
                    If GetPosStock(iquantity, iNum, .bodyText(nRow, "cinvcode"), _
                        .bodyText(nRow, "cfree1"), .bodyText(nRow, "cfree2"), .bodyText(nRow, "cfree3"), _
                        .bodyText(nRow, "cfree4"), .bodyText(nRow, "cfree5"), .bodyText(nRow, "cfree6"), _
                        .bodyText(nRow, "cfree7"), .bodyText(nRow, "cfree8"), .bodyText(nRow, "cfree9"), _
                        .bodyText(nRow, "cfree10"), .bodyText(nRow, "cbatch"), sPosition, sWhCode) Then
                    End If
                End If
            End If
            
            If tmpInventory.GroupType = 2 Then
                .headerText("iavanum") = Format(iNum, ClsAccount.FormatNumDecString)
                .bodyText(nRow, "iavanum") = Format(iNum, ClsAccount.FormatNumDecString)
            ElseIf tmpInventory.GroupType = 1 And val(.bodyText(.row, "iinvexchrate")) <> 0 Then
                .headerText("iavanum") = Format(iquantity / .bodyText(.row, "iinvexchrate"), ClsAccount.FormatNumDecString)
                .bodyText(nRow, "iavanum") = Format(iquantity / .bodyText(nRow, "iinvexchrate"), ClsAccount.FormatNumDecString)
            Else
                .headerText("iavanum") = ""
            End If
            .headerText("iavaquantity") = Format(iquantity, ClsAccount.FormatQuanDecString)
            .bodyText(nRow, "iavaquantity") = Format(iquantity, ClsAccount.FormatQuanDecString)
        End If
    End With
    Set ClsAccount = Nothing
    Set clsStockCo = Nothing
    Set tmpInventory = Nothing
    Set ClsInventoryCO = Nothing
    VBA.Err.Clear
End Sub

Public Function GetPosStock(ByRef iquantity As Double, ByRef iNum As Double, _
                             ByVal sInvCode As String, _
                             Optional sfree1 As String = "", _
                             Optional sfree2 As String = "", _
                             Optional sfree3 As String = "", _
                             Optional sfree4 As String = "", _
                             Optional sfree5 As String = "", _
                             Optional sfree6 As String = "", _
                             Optional sfree7 As String = "", _
                             Optional sfree8 As String = "", _
                             Optional sfree9 As String = "", _
                             Optional sfree10 As String = "", _
                             Optional sBatch As String = "", _
                             Optional sPosition As String, _
                             Optional sWhCode As String = "") As Boolean
    On Error GoTo Error_General_Handler:
    Dim sSql As String
    Dim i As Long
    Dim cnn As ADODB.Connection
    Dim rst As ADODB.Recordset
    Set cnn = mologin.AccountConnection
    Set rst = New ADODB.Recordset
    rst.CursorLocation = adUseClient
    sSql = "select SUM(ISNULL(iQuantity,0)) as iquantity,"
    sSql = sSql + " sum(CASE WHEN I.iGroupType = 0 THEN 0 WHEN I.iGroupType = 2 THEN ISNULL(iNum,0) "
    sSql = sSql + " WHEN I.iGroupType = 1 THEN iQuantity/ CU_G.iChangRate END) as inum "
    sSql = sSql + " from invpositionsum CS left join inventory I on CS.cinvcode=I.cinvcode "
    sSql = sSql + "LEFT JOIN ComputationUnit CU_G ON  I.cSTComUnitCode = CU_G.cComUnitCode where CS.cinvcode=N'" & sInvCode & "'"
    sSql = sSql + IIf(sfree1 = "", "", "and cfree1=N'" & sfree1 & "'")
    sSql = sSql + IIf(sfree2 = "", "", "and cfree2=N'" & sfree2 & "'")
    sSql = sSql + IIf(sfree3 = "", "", "and cfree3=N'" & sfree3 & "'")
    sSql = sSql + IIf(sfree4 = "", "", "and cfree4=N'" & sfree4 & "'")
    sSql = sSql + IIf(sfree5 = "", "", "and cfree5=N'" & sfree5 & "'")
    sSql = sSql + IIf(sfree6 = "", "", "and cfree6=N'" & sfree6 & "'")
    sSql = sSql + IIf(sfree7 = "", "", "and cfree7=N'" & sfree7 & "'")
    sSql = sSql + IIf(sfree8 = "", "", "and cfree8=N'" & sfree8 & "'")
    sSql = sSql + IIf(sfree9 = "", "", "and cfree9=N'" & sfree9 & "'")
    sSql = sSql + IIf(sfree10 = "", "", "and cfree10=N'" & sfree10 & "'")
    sSql = sSql + IIf(sBatch = "", "", "and cbatch=N'" & sBatch & "'")
    sSql = sSql + IIf(sVmiCode = "", "", "and cvmivencode=N'" & sVmiCode & "'")
    sSql = sSql + IIf(sWhCode = "", "", "and cwhcode=N'" & sWhCode & "'")
    sSql = sSql + IIf(sPosition = "", "", "and cPoscode=N'" & sPosition & "'")
    sSql = sSql + IIf(cInVouchType = "", "", "and cInVouchType=N'" & cInVouchType & "'")
    sSql = sSql + IIf(iTrackId = 0, "", "and iTrackId=" & iTrackId)
    rst.Open sSql, cnn, adOpenStatic, adLockReadOnly
    If Not (rst.EOF And rst.BOF) Then
        iquantity = FormatToDouble(vFieldVal(rst.Fields("iquantity")), mologin.Account.FormatQuanDecString)
        iNum = FormatToDouble(vFieldVal(rst.Fields("inum")), mologin.Account.FormatNumDecString)
        Set rst = Nothing
        Set cnn = Nothing
    End If
    GetPosStock = True
    Exit Function
Error_General_Handler:
    GetPosStock = False
    MsgBox Err.Description, vbInformation, GetString("U8.DZ.JA.Res030")
    Exit Function
End Function

Public Sub QueryStock(ByRef Voucher As UAPVoucherControl85.ctlVoucher)
    Dim i As Integer
    Dim iRow As Long
    Dim sInvCode As String
    Dim StrReportName As String
    Dim oInventory As USERPVO.Inventory
    Dim oInventoryPst As USERPDMO.InventoryPst
    
    
    'Dim StrHWReportName As String
    StrReportName = "现存量查询"
    'StrHWReportName = "货位存量查询"
    iRow = Voucher.row
    sInvCode = Voucher.bodyText(iRow, "cInvCode")
    
    If sInvCode <> "" Then
    
        If Voucher.bodyText(iRow, "cposition") <> "" Then
            StrReportName = "货位存量查询"
        End If
        
        Dim poRep As Object
        Dim oFltSrv As New UFGeneralFilter.FilterSrv
        Set poRep = CreateObject("ReportService.clsReportManager")
        Set oInventoryPst = New InventoryPst
        oInventoryPst.login = mologin
        Set oInventory = New USERPVO.Inventory
        oInventoryPst.Load sInvCode, oInventory
        oFltSrv.OpenFilter mologin.OldLogin, "", StrReportName, "ST", , True
        Call FillFltSrvForStock(oFltSrv, iRow, StrReportName, oInventory, Voucher)
        poRep.HelpPath = App.HelpFile
        poRep.OpenReportNoneFilterUI StrReportName, mologin.OldLogin, "ST", oFltSrv
    Else
        'poRep.OpenReport StrReportName, ologin, "ST"
        'ST_OpenReport StrReportName, ologin, "ST"
    End If
End Sub

Private Sub FillFltSrvForStock(ByRef oFltSrv As Object, ByVal iRow As Long, ByVal StrReportName As String, oInventory As Object, ByRef Voucher As UAPVoucherControl85.ctlVoucher)
    Dim sInvCode As String
    Dim sWhCode As String
    sInvCode = Voucher.bodyText(iRow, "cInvCode")
    sWhCode = Voucher.bodyText(iRow, "cwhcode")
    
    oFltSrv.FilterList.Item("I.cInvCode").varValue = sInvCode
    oFltSrv.FilterList.Item("I.cInvCode").varValue2 = sInvCode
    oFltSrv.FilterList.Item("W.cWhCode").varValue = sWhCode
    oFltSrv.FilterList.Item("W.cWhCode").varValue2 = sWhCode
    
    For i = 1 To 10
        oFltSrv.FilterList.Item("CS.cFree" & CStr(i)).varValue = Voucher.bodyText(iRow, "cfree" & CStr(i))
        oFltSrv.FilterList.Item("CS.cFree" & CStr(i)).varValue2 = Voucher.bodyText(iRow, "cfree" & CStr(i))
    Next
    
    oFltSrv.FilterList.Item("CS.cBatch").varValue = Voucher.bodyText(iRow, "cbatch")
    oFltSrv.FilterList.Item("CS.cBatch").varValue2 = Voucher.bodyText(iRow, "cbatch")
    
    If StrReportName = "货位存量查询" Then
        oFltSrv.FilterList.Item("CS.cPosCode").varValue = Voucher.bodyText(iRow, "cposition")
        oFltSrv.FilterList.Item("CS.cPosCode").varValue2 = Voucher.bodyText(iRow, "cposition")
    End If
    
    If oInventory.IsLP And StrReportName = "现存量查询" Then
        oFltSrv.FilterList.Item("CS.iSoType").varValue = val(Voucher.bodyText(iRow, "isotype"))
        If val(Voucher.bodyText(iRow, "isoseq")) <> 0 Then
            oFltSrv.FilterList.Item("CS.iSoSeq").varValue = Voucher.bodyText(iRow, "isoseq")
            oFltSrv.FilterList.Item("CS.iSoSeq").varValue2 = Voucher.bodyText(iRow, "isoseq")
        End If
        oFltSrv.FilterList.Item("CS.cSoCode").varValue = Voucher.bodyText(iRow, "csocode")
        oFltSrv.FilterList.Item("CS.cSoCode").varValue2 = Voucher.bodyText(iRow, "csocode")
    End If
    
End Sub

Public Sub QueryStockAll(ByRef Voucher As UAPVoucherControl85.ctlVoucher)
    Dim iRow As Long
    Dim sInvCode As String
    Dim sWhCode As String
    Dim StrReportName As String
    Dim i As Integer
    Dim j As Integer
    Dim iSotype As Integer
    Dim iSodid As String
    Dim cVmivenCode As String
    Dim sSql As String
    Dim iCount As Integer
    Dim bQuery As Boolean
    
    Dim sTmpTableName As String
    Dim oInv As New Inventory
    Dim oInvPst As New InventoryPst
    
    oInvPst.login = mologin
    
    StrReportName = "现存量查询"
    
    sTmpTableName = CreateGUID("", False) & "_StockTmpTable_ST"
    
    DropTable "tempdb..[" & sTmpTableName & "]"
    sSql = ""
    sSql = sSql + " select cwhcode,cinvcode,cbatch,cfree1,cfree2,cfree3,cfree4,cfree5,cfree6,cfree7,cfree8,cfree9,cfree10,cvmivencode,isotype,isodid into "
    sSql = sSql + " tempdb..[" & sTmpTableName & "] from currentstock where 1=2 "
    
    mologin.AccountConnection.Execute sSql

    iCount = 0
    sSql = ""
    
    For i = 1 To Voucher.BodyRows
        iSotype = 0
        iSodid = ""
        cVmivenCode = ""
        sWhCode = Voucher.bodyText(i, "cwhcode")
        sInvCode = Voucher.bodyText(i, "cinvcode")
        If sInvCode <> "" Then
            bQuery = True
           
            oInvPst.Load sInvCode, oInv, , , "R"
            
            sSql = sSql & " insert into tempdb..[" & sTmpTableName & "] (cwhcode,cinvcode,cbatch,cfree1,cfree2,cfree3,cfree4,cfree5,cfree6,cfree7,cfree8,cfree9,cfree10,cvmivencode,isotype,isodid ) values ("
            sSql = sSql & "N'" & sWhCode & "',N'" & sInvCode & "',N'" & Voucher.bodyText(i, "cbatch") & "',"
            For j = 1 To 10
              sSql = sSql & "N'" & Voucher.bodyText(i, "cfree" & j) & "',"
            Next
            
            sSql = sSql & "N'" & cVmivenCode & "',"
            sSql = sSql & iSotype & ","
            sSql = sSql & "N'" & iSodid & "'"
            sSql = sSql & " )" & vbCrLf
            
            iCount = iCount + 1
           
            If iCount >= 50 Then
                mologin.AccountConnection.Execute sSql
                iCount = 0
                sSql = ""
            End If
        End If
    Next
    If sSql <> "" Then
       mologin.AccountConnection.Execute sSql
    End If
    
  
   
    If bQuery Then
        
        Dim poRep As Object
        Dim oFltSrv As New UFGeneralFilter.FilterSrv
        Set poRep = CreateObject("ReportService.clsReportManager")
        
        oFltSrv.OpenFilter mologin.OldLogin, "", StrReportName, "ST", , True
        oFltSrv.FilterList.Item("stocktmptablename").varValue = sTmpTableName
        poRep.HelpPath = App.HelpFile
        poRep.OpenReportNoneFilterUI StrReportName, mologin.OldLogin, "ST", oFltSrv
        
    End If
    
End Sub

'审核自动生成其他出库单
Public Function ExecPushOtherOut(lngVoucherID As Long) As String
    On Error GoTo Err_Handler:
    Dim errStr As String
    
    ExecPushOtherIn = ""
    
    Dim tmpCol As Collection
    Set tmpCol = New Collection
    
    If clsbill Is Nothing Then
        Set clsbill = New USERPCO.VoucherCO
        clsbill.IniLogin g_oLogin, errStr
        Set mologin = clsbill.login
    End If

    If clsbill.MakeOtherOutVouch(CStr(lngVoucherID), tmpCol, errStr) = False Then
        ExecPushOtherOut = GetStringPara("U8.ST.MakeVouchByRefering.00062", Replace(errStr, vbCrLf, ""))
    Else
        ReDim varArgs(2)
        Dim sTmp As String
        Dim sCode As String
        varArgs(0) = tmpCol.Count
        varArgs(1) = GetString("U8.DZ.JA.Res1950") '其他出库单
        For i = 1 To tmpCol.Count
            If GetFieldValue(g_Conn, "rdrecord09", "ccode", "id", CStr(tmpCol.Item(i)), sCode) Then
                sTmp = sTmp & sCode & ","
            Else
                sTmp = sTmp & CStr(tmpCol.Item(i)) & ","
            End If
        Next
        sTmp = Left(sTmp, Len(sTmp) - 1)
        varArgs(2) = sTmp
        ExecPushOtherOut = GetStringPara("U8.DZ.JA.Res985", varArgs(0), varArgs(1), varArgs(2)) & vbCrLf
    End If
    
    Set tmpCol = Nothing
    Exit Function
Err_Handler:
    Set tmpCol = Nothing
    ExecPushOtherOut = Err.Description
End Function

Public Function CheckCanBack(lngVoucherID As Long, cCode As String, sCreateType As String, strMsg As String) As Boolean
    Dim vouStatus As String
    vouStatus = CheckVoucherStatus(lngVoucherID, sCreateType)
    
    CheckCanBack = False
    '只有已出库，才能归还
    If vouStatus = "生单" Then
        ReDim varArgs(0)
        varArgs(0) = cCode
        If VoucherIsOut(lngVoucherID) = False Then
           CheckCanBack = False
           varArgs(0) = cCode
           strMsg = strMsg & GetStringPara("U8.ST.V870.00757", varArgs(0)) & vbCrLf '
        Else
        
            If VoucherIsAllBack(lngVoucherID) Then
                CheckCanBack = False
                strMsg = strMsg & GetStringPara("U8.ST.V870.00758", varArgs(0)) & vbCrLf
                'strMsg = strMsg & "单据 " & cCode & " 已经归还！" & vbCrLf
            Else
                CheckCanBack = True
            End If
        End If
    ElseIf vouStatus = "审核" Then
        '期初单据不用出库
        If sCreateType = "期初单据" Then
            CheckCanBack = True
        Else
            CheckCanBack = False
            ReDim varArgs(0)
            varArgs(0) = cCode
            strMsg = strMsg & GetStringPara("U8.ST.V870.00757", varArgs(0)) & vbCrLf '
            'strMsg = strMsg & "单据 " & cCode & " 未出库！" & vbCrLf
        End If
    ElseIf vouStatus = "新建" Then
        CheckCanBack = False
        ReDim varArgs(0)
        varArgs(0) = cCode
        strMsg = strMsg & GetStringPara("U8.DZ.JA.Res460", varArgs(0)) & vbCrLf
        'strMsg = strMsg & "单据 " & ccode & " 没有审核！" & vbCrLf
    ElseIf vouStatus = "关闭" Then
        CheckCanBack = False
        ReDim varArgs(0)
        varArgs(0) = cCode
        strMsg = strMsg & GetStringPara("U8.DZ.JA.Res445", varArgs(0)) & vbCrLf
        'strMsg = strMsg & "单据 " & ccode & " 已关闭！" & vbCrLf
    Else
        CheckCanBack = False
        ReDim varArgs(0)
        varArgs(0) = cCode
        strMsg = strMsg & GetStringPara("U8.DZ.JA.Res440", varArgs(0)) & vbCrLf
        'strMsg = strMsg & "单据 " & ccode & " 不存在！" & vbCrLf
    End If
End Function


Public Function ExecReturn(ByRef lngVoucherID As Long, ByRef sMsg As String, ByRef IsBackWfcontrolled As Boolean, Optional ByVal sUfts As String = "") As Boolean
    On Error GoTo Err_Handler:
    
    Dim sErrMsg As String
    Dim lngReturnVoucherID As Long
    Dim oBorrowOutBack As Object
    Set oBorrowOutBack = CreateObject("HY_DZ_BorrowOutBack.clsBorrowOutSrv")
    oBorrowOutBack.Init g_oLogin
    
    '生成归还单成功
    'If oBorrowOutBack.MakeVouchFromBorrowOut(lngVoucherID, sErrMsg, lngReturnVoucherID, GetTimeStamp(g_Conn, MainTable, lngVoucherID)) Then
    If oBorrowOutBack.MakeVouchFromBorrowOut(lngVoucherID, sErrMsg, lngReturnVoucherID, sUfts) Then
        ExecReturn = True
        
        ReDim varArgs(2)
        Dim sTmp As String
        varArgs(0) = 1
        varArgs(1) = GetString("U8.ST.V870.00756") '借出归还单
        
        If Not GetFieldValue(g_Conn, "HY_DZ_BorrowOutBack", "ccode", "id", CStr(lngReturnVoucherID), sTmp) Then
            sTmp = lngReturnVoucherID
        End If
        varArgs(2) = sTmp
        'sMsg = sMsg & "已生成借出归还单" & sTmp & "!" & vbCrLf
        sMsg = sMsg & GetStringPara("U8.DZ.JA.Res985", varArgs(0), varArgs(1), varArgs(2)) & vbCrLf
        
        '借出归还单单不是工作流控制，则审核归还单，生成入库单
        If Not IsBackWfcontrolled And Not IsBlank(lngReturnVoucherID) Then
            If oBorrowOutBack.Verify(lngReturnVoucherID, sErrMsg) Then
            
                '审核成功
                If sErrMsg <> "" Then sMsg = sMsg & sErrMsg & vbCrLf
                
                '再推其他入库单
                sErrMsg = oBorrowOutBack.PushOtherIn(lngReturnVoucherID)
                If sErrMsg <> "" Then sMsg = sMsg & sErrMsg & vbCrLf
                
            Else
                '归还单审核失败
                If sErrMsg <> "" Then sMsg = sMsg & sErrMsg & vbCrLf
            End If
        End If
    Else
        ExecReturn = False
        If sErrMsg <> "" Then sMsg = sMsg & sErrMsg & vbCrLf
        Exit Function
    End If
    
    ExecReturn = True
    Exit Function
Err_Handler:
    ExecReturn = False
    sMsg = Err.Description
End Function

'获取时间戳
Public Function Getufts(MainTable As String, pk As String, voucherID As String) As String
    Dim strSql As String, rs As ADODB.Recordset
    strSql = "select CONVERT(nchar,CONVERT(money,ufts),2) as ufts from " & MainTable & " where " & pk & "= " & voucherID & ""
    Set rs = New ADODB.Recordset
    rs.Open strSql, g_Conn
    If Not rs.EOF Then
        Getufts = rs!ufts
    End If
    rs.Close
End Function

Private Function GetMDeptCode(moid As String) As String
    Dim strSql As String, rs As ADODB.Recordset
    strSql = "select top 1 MDeptCode  from mom_orderdetail where moid='" & moid & "'"
    Set rs = New ADODB.Recordset
    rs.Open strSql, g_Conn
    If Not rs.EOF Then
        GetMDeptCode = rs!MDeptCode & ""
    End If
    rs.Close
End Function
