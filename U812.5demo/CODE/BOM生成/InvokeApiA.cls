VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "InvokeApiA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit



Private objLoginable As UFPortalProxyInterface.ILoginable
Public Business2 As Object
Private AddState As Integer
Public errMsg As String
'变量初始化


'自动获取单据号
Private Sub BillNumberChecksucceed(Voucher As Object)
    Dim oDomHead As DOMDocument
    Dim oDomFormat As DOMDocument
    Dim oelement As IXMLDOMElement
    Dim bManualCode As Boolean
    Dim bCanModyCode As Boolean
    Dim strVoucherNo As String

    Dim sKey As String

    Dim sError As String

    Set oDomHead = Voucher.GetHeadDom

    If GetVoucherNO(g_Conn, oDomHead, gstrCardNumber, sError, strVoucherNo, oDomFormat, True, , False) = True Then
        Voucher.SetBillNumberRule oDomFormat.xml

        Set oelement = oDomFormat.selectSingleNode("//单据编号")

        '支持完全手工编号
        '允许手工修改得含义为 完全手工编号， 重号自动重取的含义为 手工修改，重号自动重取
        bManualCode = oelement.getAttribute("允许手工修改")
        bCanModyCode = oelement.getAttribute("允许手工修改") Or oelement.getAttribute("重号自动重取")
    Else
        MsgBox sError, vbInformation, GetString("U8.DZ.JA.Res030")
    End If

    '支持完全手工编号，此时不取单据号 2003-07-16 黄朝阳
    If Not bManualCode Then
        With Voucher
            sKey = strcCode

            Set oDomHead = Voucher.GetHeadDom

            If GetVoucherNO(g_Conn, oDomHead, gstrCardNumber, sError, strVoucherNo, , , , False) = False Then
                MsgBox sError, vbInformation, GetString("U8.DZ.JA.Res030")
            Else
                .headerText(sKey) = strVoucherNo
            End If
        End With
    End If
End Sub

'字段有效性检查
Public Function APiExecFunSaveCheck(Voucher As Object) As Boolean
    Dim strMsgHead As String

    On Error GoTo Err_Handler
    '      Voucher.VoucherStatus = AddState

    '判断单据编号是否已存在
    If AddState = 2 And IsExisted(Voucher.headerText("cCode")) Then
        '    If IsExisted(Voucher.headerText("cCode")) Then
        errMsg = GetString("U8.DZ.JA.Res1080")
        '        MsgBox GetString("U8.DZ.JA.Res1080"), vbInformation, getstring("U8.DZ.JA.Res030")
        APiExecFunSaveCheck = False
        '        Voucher.SetFocus
        Exit Function
    End If

    '检查单据表头表体是否有数据
    If Voucher.headVaildIsNull2(strMsgHead) = False Then
        errMsg = strMsgHead
        '        MsgBox strMsgHead, vbCritical, getstring("U8.DZ.JA.Res030")
        APiExecFunSaveCheck = False
        '        Voucher.SetFocus
        Exit Function
    End If

    '单据退出编辑状态
    Voucher.ProtectUnload2
    If Voucher.BodyRows <= 0 Then
        errMsg = "表体没有数据，请输入"
        '        MsgBox GetString("U8.DZ.JA.Res1090"), vbCritical, getstring("U8.DZ.JA.Res030")
        APiExecFunSaveCheck = False
        Exit Function
    End If



    '检查表体数据的有效性,主要检查:
    '必填项是否空
    '自由项组合是否合法
    '批次交验
    '入库单号校验
    '项目编码校验
    '有效期校验

    If ExecFunEffectiveCheck(Voucher) = False Then
        APiExecFunSaveCheck = False
        Exit Function
    End If

    Voucher.ProtectUnload

    APiExecFunSaveCheck = True
    Exit Function

Err_Handler:
    MsgBox Err.Description, vbInformation, GetString("U8.DZ.JA.Res030")

End Function
'检查表体数据的有效性,主要检查:
'必填项是否空
'自由项组合是否合法
'批次交验
'入库单号校验
'项目编码校验
'有效期校验
Private Function apiExecFunEffectiveCheck(Voucher As Object) As Boolean
    On Error GoTo Err_Handler

    '表头校验:单据编号,制单人,单据日期为必填项,
    '仓库,部门等其他项目根据实际业务单据而定
    If Voucher.headerText(strcCode) = "" Then
        errMsg = GetString("U8.DZ.JA.Res1120")
        '        MsgBox GetString("U8.DZ.JA.Res1120"), vbInformation, getstring("U8.DZ.JA.Res030")
        apiExecFunEffectiveCheck = False
        Exit Function
    End If
    If Voucher.headerText(StrcMaker) = "" Then
        errMsg = GetString("U8.DZ.JA.Res1130")
        '        MsgBox GetString("U8.DZ.JA.Res1130"), vbInformation, getstring("U8.DZ.JA.Res030")
        apiExecFunEffectiveCheck = False
        Exit Function
    End If

    If Voucher.headerText(StrdDate) = "" Then
        errMsg = GetString("U8.DZ.JA.Res1140")
        '        MsgBox GetString("U8.DZ.JA.Res1140"), vbInformation, getstring("U8.DZ.JA.Res030")
        apiExecFunEffectiveCheck = False
        Exit Function
        '    Else
        '        '借用日期 不做大于制单日期控制，因为有补单需求
        '        If Voucher.VoucherStatus = VSeAddMode Then
        '        End If
    End If

    If Voucher.headerText("cfreight") = "是" Then
        If Voucher.headerText("MycdefineT2") = "" Or Voucher.headerText("cfreightType") = "" _
           Or Voucher.headerText("cfreightCost") = "" Then
            errMsg = GetString("U8.DZ.JA.Res1150")
            '            MsgBox GetString("U8.DZ.JA.Res1150"), vbInformation, getstring("U8.DZ.JA.Res030")
            apiExecFunEffectiveCheck = False
            Exit Function
        End If
    End If

    '单据表体校验
    If ExecFunBodyCheck(Voucher) = False Then
        apiExecFunEffectiveCheck = False
        Exit Function
    End If

    apiExecFunEffectiveCheck = True
    Exit Function

Err_Handler:
    MsgBox Err.Description, vbInformation, GetString("U8.DZ.JA.Res030")

End Function    '检查表体数据的有效性,主要检查:
'必填项是否空
'自由项组合是否合法
'批次交验
'入库单号校验
'项目编码校验
'有效期校验
Private Function ExecFunEffectiveCheck(Voucher As Object) As Boolean
    On Error GoTo Err_Handler

    '表头校验:单据编号,制单人,单据日期为必填项,
    '仓库,部门等其他项目根据实际业务单据而定
    If Voucher.headerText(strcCode) = "" Then
        errMsg = GetString("U8.DZ.JA.Res1120")
        '        MsgBox GetString("U8.DZ.JA.Res1120"), vbInformation, getstring("U8.DZ.JA.Res030")
        ExecFunEffectiveCheck = False
        Exit Function
    End If
    If Voucher.headerText(StrcMaker) = "" Then
        errMsg = GetString("U8.DZ.JA.Res1130")
        '        MsgBox GetString("U8.DZ.JA.Res1130"), vbInformation, getstring("U8.DZ.JA.Res030")
        ExecFunEffectiveCheck = False
        Exit Function
    End If

    If Voucher.headerText(StrdDate) = "" Then
        errMsg = GetString("U8.DZ.JA.Res1140")
        '        MsgBox GetString("U8.DZ.JA.Res1140"), vbInformation, getstring("U8.DZ.JA.Res030")
        ExecFunEffectiveCheck = False
        Exit Function
        '    Else
        '        '借用日期 不做大于制单日期控制，因为有补单需求
        '        If Voucher.VoucherStatus = VSeAddMode Then
        '        End If
    End If

    If Voucher.headerText("cfreight") = "是" Then
        If Voucher.headerText("MycdefineT2") = "" Or Voucher.headerText("cfreightType") = "" _
           Or Voucher.headerText("cfreightCost") = "" Then
            errMsg = GetString("U8.DZ.JA.Res1150")
            '            MsgBox GetString("U8.DZ.JA.Res1150"), vbInformation, getstring("U8.DZ.JA.Res030")
            ExecFunEffectiveCheck = False
            Exit Function
        End If
    End If

    '单据表体校验
    If ExecFunBodyCheck(Voucher) = False Then
        ExecFunEffectiveCheck = False
        Exit Function
    End If

    ExecFunEffectiveCheck = True
    Exit Function

Err_Handler:
    MsgBox Err.Description, vbInformation, GetString("U8.DZ.JA.Res030")

End Function

'表体有效性校验
'必填项是否空
'启用的自由项是否都输入了
'自由项组合是否合法
'批次交验
'入库单号校验
'有效期校验
Private Function ExecFunBodyCheck(Voucher As Object) As Boolean
    On Error GoTo Err_Handler
    '表体有效性校验,由于表体是多行需要循环
    Dim oDomHead, oDomBody As DOMDocument
    Dim bodyele As IXMLDOMElement
    Dim iRow As Integer
    Dim rs As New ADODB.Recordset
    Dim tRs As New ADODB.Recordset
    Dim sql As String
    Dim iSTConMode As Integer
    Dim iQ As Integer

    Voucher.getVoucherDataXML oDomHead, oDomBody

    iRow = 0

    For Each bodyele In oDomBody.selectNodes("//z:row[@editprop != 'D']")
        iRow = iRow + 1
        '临时赋值,作为行标
        '此处赋值不管用
        '    bodyele.setAttribute "AutoID", irow

        '1
        '比填项,简单校验：存货编码，数量
        '其他的校验根据实际业务而定
        If bodyele.getAttribute("cinvcode") = "" Then
            errMsg = "第" & iRow & "行存货编码不能为空"
            '            MsgBox "第" & iRow & "行存货编码不能为空", vbInformation, getstring("U8.DZ.JA.Res030")
            ExecFunBodyCheck = False
            Exit Function
        End If

        If Val(bodyele.getAttribute("iquantity") & "") = 0 And Val(bodyele.getAttribute("inum") & "") = 0 Then
            '            MsgBox "第" & iRow & "行 数量件数不能同时为空或等于0", vbInformation, getstring("U8.DZ.JA.Res030")
            errMsg = "第" & iRow & "行 数量/件数不能同时为空或等于0"
            '            MsgBox "第" & iRow & "行 数量/件数不能同时为空或等于0", vbInformation, getstring("U8.DZ.JA.Res030")
            '标识当前行,着蓝色
            Voucher.row = iRow
            ExecFunBodyCheck = False
            Exit Function
        End If

        'dxb 仓库必须录入！
        If bodyele.getAttribute("cwhcode") = "" Then
            errMsg = "第" & iRow & "行仓库不能为空"
            '            MsgBox "第" & iRow & "行仓库不能为空", vbInformation, getstring("U8.DZ.JA.Res030")
            ExecFunBodyCheck = False
            Exit Function
        End If

        '2
        '启用的自由项是否都输入了
        '结构性，非结构性自由项都必须输入
        '自由项组合是否合法
        If ExecFunFreeCheck(bodyele.getAttribute("cinvcode"), bodyele, iRow) = False Then
            '标识当前行,着蓝色
            Voucher.row = iRow
            'Voucher.SetCurrentRow ("@AutoID=" & bodyele.getAttribute("AutoID") & "")
            ExecFunBodyCheck = False
            Exit Function
        End If


        '３
        '批次
        '出库跟踪入库
        If ExecFuncbatch(bodyele.getAttribute("cinvcode"), _
                         bodyele.getAttribute("cbatch") & "", _
                         bodyele.getAttribute("cinvouchcode") & "", _
                         iRow) = False Then
            '标识当前行,着蓝色
            Voucher.row = iRow
            'Voucher.SetCurrentRow ("@AutoID=" & bodyele.getAttribute("AutoID") & "")
            ExecFunBodyCheck = False
            Exit Function
        End If



        '5 预计归还日期 dxb
        If bodyele.getAttribute("backdate") = "" Then
            errMsg = "第" & iRow & "行预计归还日期不能为空"
            '            MsgBox "第" & iRow & "行预计归还日期不能为空", vbInformation, getstring("U8.DZ.JA.Res030")
            ExecFunBodyCheck = False
            Exit Function
        Else
            If CDate(Mid(bodyele.getAttribute("backdate"), 1, 10)) < CDate(Mid(Voucher.headerText(StrdDate), 1, 10)) Then
                errMsg = "第" & iRow & "行预计归还日期不能早于借用日期"
                ExecFunBodyCheck = False
                Exit Function
            End If
        End If
    Next

    If CheckNumOutLimit(Voucher) = "" Then
        MsgBox GetString("U8.DZ.JA.Res1220"), vbInformation, GetString("U8.DZ.JA.Res030")
        ExecFunBodyCheck = False
        Exit Function
    End If

    ExecFunBodyCheck = True
    Exit Function

Err_Handler:
    MsgBox Err.Description, vbInformation, GetString("U8.DZ.JA.Res030")
End Function

'批次
'入库单号
Private Function ExecFuncbatch(cinvcode As String, cbatch As String, cTrackCode As String, iRow As Integer) As Boolean
    On Error GoTo Err_Handler

    Dim sql As String
    Dim rs As New ADODB.Recordset

    sql = "select bInvBatch,bTrack from inventory where cinvcode='" & cinvcode & "'"
    rs.Open sql, g_Conn, 1, 1
    If Not rs.EOF Then
        If CBool(rs("bInvBatch")) And cbatch = "" Then
            MsgBox "第" & iRow & "行存货启用了批次管理，批号不能为空", vbInformation, GetString("U8.DZ.JA.Res030")
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
        MsgBox "第" & iRow & "行存货编码" & cinvcode & "不存在", vbInformation, GetString("U8.DZ.JA.Res030")
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
    MsgBox Err.Description, vbInformation, GetString("U8.DZ.JA.Res030")
End Function
Private Function CheckNumOutLimit(Voucher As Object) As String
    On Error GoTo lerr:
    CheckNumOutLimit = ""

    Dim i As Long
    For i = 1 To Voucher.BodyRows
        '('iquantity','inum')
        If Voucher.bodyText(i, "iquantity") <> "" Then
            If Len(Voucher.bodyText(i, "iquantity")) > 15 Or CDbl(Voucher.bodyText(i, "iquantity")) > 999999999999999# Then
                Exit Function
            End If
        End If

        If Voucher.bodyText(i, "inum") <> "" Then
            If Len(Voucher.bodyText(i, "inum")) > 15 Or CDbl(Voucher.bodyText(i, "inum")) > 999999999999999# Then
                Exit Function
            End If
        End If
    Next

    CheckNumOutLimit = "OK"
lerr:

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
    rs.Open sql, g_Conn, 1, 1
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

        bRS.Open sFind, g_Conn, 1, 1
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
    MsgBox Err.Description, vbInformation, GetString("U8.DZ.JA.Res030")
End Function

'读取时间戳，并与旧时间戳比较
Private Function ExecFunCompareUfts(strcode As String) As Boolean

'读取时间戳
    TimeStamp = GetTimeStamp(g_Conn, MainTable, lngVoucherID)

    If TimeStamp = RecordDeleted Then
        errMsg = "单据(" & strcode & ")已被其他用户删除,不可修改"
        '        MsgBox "单据(" & strcode & ")已被其他用户删除,不可修改", vbInformation, getstring("U8.DZ.JA.Res030")
        ExecFunCompareUfts = False
        Exit Function
    ElseIf TimeStamp = RecordError Then
        errMsg = "单据(" & strcode & ")数据出现错误,请刷新"
        '         MsgBox "单据(" & strcode & ")数据出现错误,请刷新", vbInformation, getstring("U8.DZ.JA.Res030")
        ExecFunCompareUfts = False
        Exit Function
    ElseIf OldTimeStamp <> TimeStamp Then
        errMsg = "单据(" & strcode & ")已被其他用户修改,请刷新"
        '        MsgBox "单据(" & strcode & ")已被其他用户修改,请刷新", vbInformation, getstring("U8.DZ.JA.Res030")
        ExecFunCompareUfts = False
        Exit Function
    Else
        OldTimeStamp = TimeStamp
        ExecFunCompareUfts = True
    End If
End Function
'保存
Private Function APIExecSubSave(octVoucher As Object) As Boolean
    On Error GoTo Err_Handler

    Dim HeadData As New CDataVO
    Dim BodyData As New CDataVO
    Dim DMO As CDMO
    Dim Rult As CResult
    Dim i As Integer


    '    Me.Voucher.RemoveEmptyRow   '清除空行

CHECK:
    For i = 1 To octVoucher.BodyRows
        '        Voucher.BodyRowIsEmpty i
        If Trim(octVoucher.bodyText(i, "cinvcode")) = "" Then
            octVoucher.row = i
            octVoucher.DelLine
            GoTo CHECK
        End If
    Next


    '有效性校验
    If APiExecFunSaveCheck(octVoucher) = False Then Exit Function

    '修改保存，需要比较时间戳，避免出现并发
    If AddState = VSeEditMode Then
        lngVoucherID = octVoucher.headerText("ID")
        OldTimeStamp = octVoucher.headerText("ufts")
        If ExecFunCompareUfts(octVoucher.headerText(strcCode)) = False Then Exit Function
    End If


    '读取单据字段和数据
    ' 读取字段
    Set HeadData = GetHeadVouchData(g_Conn, octVoucher, MainTable)
    Set BodyData = GetBodyVouchData(g_Conn, octVoucher, DetailsTable)

    g_Conn.BeginTrans

    '新增保存,获得最大id,autoid
    '修改保存时,不需要更新id
    'dxb  2009 6 15
    g_Conn.Execute "update ufsystem..ua_identity set ichildid=(select isnull(max(autoid),1) from " & DetailsTable & _
                   ") where cvouchtype='" & gstrCardNumber & "' and cAcc_Id='" & g_oLogin.cAcc_Id & "'"
    Call GetMaxID

    If AddState = VSeAddMode Then
        '        Call GetMaxID
        HeadData.Item(1).Item("ID").Value = sID
    End If


    '更新单据号流水号
    Dim oDomHead As New DOMDocument
    Dim oDomFormat As DOMDocument
    '    Dim sError As String
    Set oDomHead = octVoucher.GetHeadDom
    If Not BOGetVoucherNO(g_Conn, oDomHead, gstrCardNumber, errMsg, octVoucher.headerText(strcCode), oDomFormat, False, , , True) Then
        g_Conn.RollbackTrans
        errMsg = GetString("U8.DZ.JA.Res100")
        '            Err.Raise 0,
        GoTo Err_Handler
    End If



    '保存表头
    HeadData.Item(1).Item("dmDate").Value = Now()    '制单时间 Format(Now(), "YYYY-MM-DD HH:MM:SS")
    HeadData.Item(1).Item("iStatus").Value = 1    '状态
    Set DMO = New CDMO
    '新增
    If AddState = VSeAddMode Then
        Set Rult = DMO.Insert(g_Conn, HeadData)
        '修改
    Else
        Set Rult = DMO.Update(g_Conn, HeadData)
    End If

    If Rult.Succeed = False Then
        errMsg = Rult.MsgCode & ",保存失败"
        '        MsgBox Rult.MsgCode & ",保存失败", vbInformation, getstring("U8.DZ.JA.Res030")
        g_Conn.RollbackTrans
        Exit Function
    End If


    '新增保存,获得最大id,autoid
    '修改保存时,不需要更新id
    If AddState = VSeAddMode Then
        For i = 1 To BodyData.Count
            sAutoId = sAutoId + i - 1
            BodyData.Item(i).Item("AutoId").Value = sAutoId
            BodyData.Item(i).Item("ID").Value = sID
        Next i

        '更新ua_identity表中的字表标识,确保与单据字表autoid一致
        g_Conn.Execute "update ufsystem..ua_identity  set ichildid=" & sAutoId & " where cacc_id='" & g_oLogin.cAcc_Id & "' and cvouchtype='" & gstrCardNumber & "'"

        Set Rult = DMO.Insert(g_Conn, BodyData)

        '修改保存,不更新id
    Else
        Dim j As Integer: j = 1
        sID = HeadData.Item(1).Item("ID").Value
        For i = 1 To BodyData.Count
            If Null2Something(BodyData.Item(i).Item("AutoID").Value) = "" Then
                sAutoId = sAutoId + j - 1
                j = j + 1
                BodyData.Item(i).Item("AutoId").Value = sAutoId
            End If
            BodyData.Item(i).Item("ID").Value = sID
        Next i

        '更新ua_identity表中的字表标识,确保与单据字表autoid一致
        g_Conn.Execute "update ufsystem..ua_identity  set ichildid=" & sAutoId & " where cacc_id='" & g_oLogin.cAcc_Id & "' and cvouchtype='" & gstrCardNumber & "'"

        Set Rult = DMO.Insert(g_Conn, BodyData)
        Set Rult = DMO.DeleteByWhere(g_Conn, DetailsTable, "id=" & lngVoucherID)
        Set Rult = DMO.Insert(g_Conn, BodyData)

    End If



    If Rult.Succeed = False Then
        errMsg = Rult.MsgCode & " 保存失败"
        '        MsgBox , vbInformation, getstring("U8.DZ.JA.Res030")
        g_Conn.RollbackTrans
        Exit Function
    End If

    g_Conn.CommitTrans

    '**********************************************************
    '新增(增加/复制),删除更新全局变量
    '保存,重取lngvoucherid
    '**********************************************************


    Exit Function

Err_Handler:
    errMsg = GetString("U8.DZ.JA.Res110")
    g_Conn.RollbackTrans

End Function
Public Function VoucherAddSave(ByVal ologin As Object, ByVal oDomHead As Object, ByVal oDomBody As Object) As Boolean

    Dim octVoucher As Object
    Dim VchSrv As New clsVouchServer
    Dim TempID As Long
    Dim TempIDs As String
    Dim rs As New ADODB.Recordset
    Dim oDomHeadA As New DOMDocument
    Dim oDomBodyA As New DOMDocument
    Dim sCache As String
    Dim sql As String



    Set g_oLogin = ologin    '获得login对象
    If g_Conn.State = 1 Then g_Conn.Close
    g_Conn.ConnectionString = g_oLogin.UfDbName
    g_Conn.CursorLocation = adUseClient
    g_Conn.Open

    Call GlobalInit(g_oLogin)
    Set octVoucher = CreateObject("UAPVoucherControl85.ctlVoucher")

    TempIDs = GetVouchTempID(g_Conn, gstrCardNumber)
    Set rs = GetVouchTempRs(g_Conn, gstrCardNumber, TempIDs)
    Call octVoucher.SetTemplateData(rs)
    errMsg = ""
    AddState = 2
    Call InitVoucher(octVoucher)
    '表头
    Set rs = Nothing
    sql = "select * from v_HY_DZ_BorrowOut where 1=2"
    Call rs.Open(sql, g_Conn, 1, 1)
    rs.Save oDomHeadA, adPersistXML
    rs.Close
    Set rs = Nothing


    '表体
    sql = "select * from v_HY_DZ_BorrowOuts where 1=2"
    Call rs.Open(sql, g_Conn, 1, 1)
    rs.Save oDomBodyA, adPersistXML
    rs.Close
    Set rs = Nothing
    octVoucher.AddNew ANMNormalAdd, oDomHeadA, oDomBodyA

    octVoucher.SetBillNumberRule sCache
    octVoucher.EnableHead strcCode, GetbCanModifyVCode(manual)

    octVoucher.setVoucherDataXML oDomHead, oDomBody
    Call BillNumberChecksucceed(octVoucher)
    APIExecSubSave octVoucher
End Function
Public Function VoucherEditSave(ByVal ologin As Object, ByVal oDomHead As Object, ByVal oDomBody As Object) As Boolean

    Dim octVoucher As Object
    Dim VchSrv As New clsVouchServer
    Dim TempID As Long
    Dim TempIDs As String
    Dim rs As New ADODB.Recordset


    Set g_oLogin = ologin    '获得login对象
    If g_Conn.State = 1 Then g_Conn.Close
    g_Conn.ConnectionString = g_oLogin.UfDbName
    g_Conn.CursorLocation = adUseClient
    g_Conn.Open

    Call GlobalInit(g_oLogin)
    Set octVoucher = CreateObject("UAPVoucherControl85.ctlVoucher")

    TempIDs = GetVouchTempID(g_Conn, gstrCardNumber)
    Set rs = GetVouchTempRs(g_Conn, gstrCardNumber, TempIDs)
    Call octVoucher.SetTemplateData(rs)
    errMsg = ""
    AddState = 1
    Call InitVoucher(octVoucher)
    octVoucher.setVoucherDataXML oDomHead, oDomBody
    APIExecSubSave octVoucher

End Function
Public Function VoucherDelete(ByVal ologin As Object, oVoucherCode As Long) As Boolean
    Dim octVoucher As Object
    Dim VchSrv As New clsVouchServer
    Dim TempID As Long
    Dim TempIDs As String
    Dim rs As New ADODB.Recordset
    Dim sql As String

    Call GlobalInit(g_oLogin)

    Set g_oLogin = ologin    '获得login对象
    If g_Conn.State = 1 Then g_Conn.Close
    g_Conn.ConnectionString = g_oLogin.UfDbName
    g_Conn.CursorLocation = adUseClient
    g_Conn.Open
    sql = "select convert(nchar,convert(money,ufts),2) ufts from " _
        & MainTable & " WHERE (ID = " & oVoucherCode & ") "
    Set rs = g_Conn.Execute(sql)

    If Not rs.EOF Then

        '删除字表
        g_Conn.Execute "delete from " & DetailsTable & " where " & HeadPKFld & "=" & oVoucherCode

        '删除字表
        g_Conn.Execute "delete from " & MainTable & " where " & HeadPKFld & "=" & oVoucherCode
    End If
End Function

'获的单据摸板数据集
Private Function GetVouchTempRs(ByRef conn As Connection, CardNum As String, TempID As String) As Recordset
    Dim TempSrv As New clsVoucherTemplate
    Set GetVouchTempRs = TempSrv.GetTemplateData2(conn, CardNum, TempID)
End Function

'Function VoucherAddSave(connString As String, oDomHead As Object, oDomBody As Object, errMsg As String) As Boolean
'      Function VoucherEditSave(connString As String, oDomHead As Object, oDomBody As Object, errMsg As String) As Boolean
'      Function VoucherDelete(connString As String, oVoucherCode As String, errMsg As String) As Boolean
'获的单据显示摸板号
Private Function GetVouchTempID(ByRef conn As Connection, CardNum As String) As String
    Dim strSql As String
    Dim rs As New ADODB.Recordset
    Dim TempID As String

    strSql = "SELECT VT_ID From VoucherTemplates WHERE (VT_CardNumber = '" & CardNum & "') AND (VT_TemplateMode = '0')"
    rs.Open strSql, conn
    If Not rs.EOF Then
        TempID = rs!VT_ID
    End If

    rs.Close
    Set rs = Nothing
    GetVouchTempID = TempID
End Function
Private Sub LoadVTID()

    Dim sql As String
    Dim oRecordset As ADODB.Recordset

    'On Error GoTo Err_Handler

    sql = "SELECT [DEF_ID], " _
        & "[DEF_ID_PRN] " _
        & "FROM [Vouchers] " _
        & "WHERE ([CardNumber] = '" & Trim(gstrCardNumber) & "') "

    If oRecordset Is Nothing Then _
       Set oRecordset = New ADODB.Recordset

    Call oRecordset.Open( _
         Source:=sql, _
         ActiveConnection:=g_Conn, _
         CursorType:=adOpenStatic, _
         LockType:=adLockReadOnly, _
         Options:=adCmdText)

    If oRecordset.RecordCount < 1 Then
        Call Err.Raise( _
             Number:=vbObjectError + 512 + 6002, _
             Description:=GetString("U8.DZ.JA.Res200"))
    End If

    '    m_strVT_ID = Null2Something( _
         '            vTarget:=oRecordset.Fields("DEF_ID").Value, _
         '            vReplace:=0)

    '    m_strVT_PRN_ID = Null2Something( _
         '            vTarget:=oRecordset.Fields("DEF_ID_PRN").Value, _
         '            vReplace:=0)

Exit_Label:
    On Error GoTo 0
    If Not oRecordset Is Nothing Then
        If oRecordset.State = adStateOpen Then _
           Call oRecordset.Close
    End If
    Set oRecordset = Nothing

    Exit Sub
Err_Handler:
    If Not oRecordset Is Nothing Then
        If oRecordset.State = adStateOpen Then _
           Call oRecordset.Close
    End If
    Set oRecordset = Nothing

    Err.Raise _
            Number:=Err.Number, _
            Source:="Sub LoadVTID of Form frmVoucher", _
            Description:=Err.Description

End Sub


Private Sub InitVoucher(octVoucher As Object)

    Dim oDataSource As Object
    Dim oRecordset As ADODB.Recordset
    Dim objVoucherTemplate As Object

    On Error GoTo Err_Handler

    '重置标题

    Set objVoucherTemplate = CreateObject("UFVoucherServer85.clsVoucherTemplate")

    ' *******************************************************
    ' * 读取当前表单的模板ID (VT_ID) 值
    '
    Call LoadVTID

    ' *******************************************************
    ' * 创建单据后台服务对象
    '
    If objVoucherTemplate Is Nothing Then _
       Set objVoucherTemplate = _
       New UFVoucherServer85.clsVoucherTemplate



    '    ' 创建单据数据源对象
    Set oDataSource = CreateObject("IDataSource.DefaultDataSource")

    If oDataSource Is Nothing Then
        MsgBox GetString("U8.DZ.JA.Res160"), vbExclamation, GetString("U8.DZ.JA.Res030")
    End If

    Set oDataSource.SetLogin = g_oLogin

    Set octVoucher.SetDataSource = oDataSource

    '请注意:SetTemplateData  必须放在 Set oDataSource.SetLogin = g_oLogin 之后, 即必须先给单据数据源初始化
    '    Call SetTemplateData(octVoucher)

    octVoucher.LoginObj = g_oLogin
    octVoucher.InitDataSource

Exit_Label:
    On Error GoTo 0
    If Not oRecordset Is Nothing Then
        If oRecordset.State = adStateOpen Then _
           Call oRecordset.Close
    End If
    Set oRecordset = Nothing

    Exit Sub
Err_Handler:
    Call ShowErrorInfo( _
         sHeaderMessage:=GetString("U8.DZ.JA.Res180"), _
         lMessageType:=vbInformation, _
         lErrorLevel:=ufsELOnlyHeader _
                    )
    GoTo Exit_Label

End Sub




