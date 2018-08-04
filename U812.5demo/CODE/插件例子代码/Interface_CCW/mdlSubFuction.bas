Attribute VB_Name = "mdlSubFuction"
'本模块主要是实现发货计划的保存验证
Option Explicit

Dim i As Integer '循环变量
Dim strSaleListID As String '销售订单子表ID
Dim strSql As String 'SQL语句
Dim strFlag As String '行打开、关闭的标识
Dim Rs As New ADODB.Recordset '数据集
Dim strAutoID As String '发货计划单子表id
Dim iSendQuantity As Single  '发货总数量
Dim iAllQuantity As Single  '订单总数量
Dim iCloseQty As Single  '关闭订单已发货数量
Dim iCurrentQuantity As Single '本次发货数量
Dim iSLimit As Single  '发货充超上限
Dim strOperFlag As String  '单据操作标志，M：修改，A：增加
Dim iBeforeModify As Single '修改前某条发货计划单的发货数量（从数据库中取出）
Dim strInvCode As String  '产品编码
Dim strInvName As String ' 产品名称
Public deps As String   '部门
Public banci As String  '班次
Public cls_Public As Object
 

'发货计划单保存验证函数
Public Function SendPlan_SaveCheck(DBconn As ADODB.Connection, domHead As DOMDocument, domBody As DOMDocument, Optional strUserErr As String, Optional bsuc As Boolean) As String
    On Error GoTo ExitSavecheck
    
    For i = 0 To domBody.selectNodes("//z:row").Length - 1
        strFlag = GetBodyItemValue(domBody, "b_int4", i)
        
        '界面上的打开回写到数据库的值为0，关闭的值为1
        Select Case strFlag
            Case "打开"
                domBody.selectNodes("//z:row").Item(i).Attributes.getNamedItem("b_int4").nodeValue = "0"
            Case "关闭"
                domBody.selectNodes("//z:row").Item(i).Attributes.getNamedItem("b_int4").nodeValue = "1"
        End Select
        
        strSaleListID = GetBodyItemValue(domBody, "b_int5", i)
        strInvCode = GetBodyItemValue(domBody, "b_cinvcode", i)
        strInvName = GetBodyItemValue(domBody, "b_str9", i)
        '判断销售订单子表ID有没有空值
        If strSaleListID = "" Or IsNull(strSaleListID) Then
            strUserErr = "第" & CStr(i + 1) & "行" & strInvCode & strInvName & "没有参照销售订单生单！"
            bsuc = False
            Exit Function
        End If
        
        '取出发货计划单子表ID
        strAutoID = GetBodyItemValue(domBody, "autoid", i)
        If IsNull(strAutoID) Or strAutoID = "" Then
            strAutoID = "0"
        End If
        '判断界面中的销售订单子表ID是否存在并且有没有被关闭
        If Rs.State <> 0 Then Rs.Close
        strSql = " select isnull((select isnull(b_float4,0) from EF_plan_DispatchLists where autoid = " & strAutoID & "),0) as BMCount ," & vbCrLf & _
                " isnull((select sum(isnull(b_float4,0)) from EF_plan_DispatchLists where isnull(b_int4,0) = 0 and b_int5 = " & strSaleListID & "),0) as SendCount," & vbCrLf & _
                " isnull((select isnull(fInvOutUpLimit,0) from so_sodetails a inner join Inventory_sub b on a.cinvcode = b.cInvSubCode and autoid =" & strSaleListID & "),0) as SLimit ," & vbCrLf & _
                " isnull( (select sum(c.iquantity) from so_sodetails a" & vbCrLf & _
                " left join EF_plan_DispatchLists b on a.autoid=b.b_int5 " & vbCrLf & _
                " left join DispatchLists c on b.autoid=c.cDefine34 where b.b_int4 = 1 and a.autoid = " & strSaleListID & "),0) as CloseQty " & vbCrLf & _
                " from SO_SODetails " & vbCrLf & _
                " where autoid = " & strSaleListID  '& " and isnull(cSCloser,'') = ''"
                
        Rs.CursorLocation = adUseClient
        Rs.Open strSql, DBconn.ConnectionString, 3, 4
        
        If Rs.EOF Then
            strUserErr = "第" & CStr(i + 1) & "行" & strInvCode & strInvName & "没有参照销售订单生单！"
            bsuc = False
            Exit Function
        Else
            iSendQuantity = Rs.Fields("SendCount") '取出该销售订单子ID总的计划发货总数量
            iCloseQty = Rs.Fields("CloseQty") '取出该销售订单子ID已做发货计划已关闭但做了发货单的数量
            iAllQuantity = GetBodyItemValue(domBody, "b_float5", i) '取出该销售订单子ID订单数量
            iCurrentQuantity = GetBodyItemValue(domBody, "b_float4", i) '取出该子ID本次发货数量
            iSLimit = Rs.Fields("SLimit")  '发货充超上限
            iSLimit = 1 + iSLimit
            
             '判断累计发货数量是否大于订单数量
            strOperFlag = GetBodyItemValue(domBody, "editprop", i) '增加、修改标志
            
            iSendQuantity = iSendQuantity + iCloseQty '已发货总数量＝未关闭总数量＋已关闭发货计划已发货数量
            
            Select Case strOperFlag
                Case "M" '修改
                    iBeforeModify = Rs.Fields("BMCount") '取出该子ID修改前发货数量
                    If iAllQuantity * iSLimit < (iCurrentQuantity - iBeforeModify) + iSendQuantity Then
                        strUserErr = "第" & CStr(i + 1) & "行的预发货总数量超过发货允超上限！"
                        bsuc = False
                        Exit Function
                    End If
                Case "A" '增加
                    If iAllQuantity * iSLimit < iSendQuantity + iCurrentQuantity Then
                        strUserErr = "第" & CStr(i + 1) & "行的预发货总数量超过发货允超上限！"
                        bsuc = False
                        Exit Function
                    End If
            End Select
        End If
    Next
ExitSavecheck:
    SendPlan_SaveCheck = Err.Description
End Function
'设置机台信息
Public Function SetMachineStation(DBconn As ADODB.Connection, strDepName As String, oVoucher As Object, Optional strUserErr As String, Optional bsuc As Boolean) As String
    On Error GoTo ExitSetMachineStation

    Dim inum As Long  '机台编号字段名称序号
    Dim strColumnName As String '字段名称
    Dim L As Long
    
    If Rs.State <> 0 Then Rs.Close
    
    strSql = " select b_str2,isnull(b_str3,'') as b_str3  from V_EF_Machines where b_cdepcode = '" & deps & "' order by b_str2 "
    
    Rs.CursorLocation = adUseClient
    Rs.Open strSql, DBconn.ConnectionString, 3, 4
    
    If Rs.EOF Then
        strUserErr = strDepName & "没有对应的机台！"
        bsuc = False
        Exit Function
    Else
        inum = 5  '机台一字段名称从b_str5开始
        For i = 1 To Rs.RecordCount
            strColumnName = "b_str" & CStr(inum)
            For L = 1 To oVoucher.BodyRows
                oVoucher.bodyText(L, strColumnName) = Rs.Fields("b_str2").Value '机台编码
                oVoucher.bodyText(L, "b_str" & CStr(64 + i)) = Rs.Fields("b_str3").Value  '机台名称(b_str65 --- b_str84)
            Next
            inum = inum + 3 '每个机台字段之间差3,机台一字段名为b_str5,机台二字段名称为b_str8,以此类推
            
            Rs.MoveNext '指针移到下一条记录
        Next
    End If
    
ExitSetMachineStation:
    SetMachineStation = Err.Description
End Function
'设置部门机台班组信息
Public Function SetClassGroup(DBconn As ADODB.Connection, strDepName As String, oVoucher As Object, Optional strUserErr As String, Optional bsuc As Boolean) As String
    On Error GoTo ExitSetClassGroup

    Dim inum As Long  '班组编码字段序号
    Dim strColumnName As String '字段名称
    Dim j As Long

    If Rs.State <> 0 Then Rs.Close
    
    strSql = " select distinct b_str1,isnull(b_str2,'') as b_str2 from V_list_EF_dep_Bz where t_cdepcode = '" & deps & "' order by b_str1 "
    
    Rs.CursorLocation = adUseClient
    Rs.Open strSql, DBconn.ConnectionString, 3, 4
    
    If Rs.EOF Then
        strUserErr = strDepName & "没有对应的班组！"
        bsuc = False
        Exit Function
    Else
        inum = 6  '部门机台班组一字段名称从b_str6开始
        For i = 1 To Rs.RecordCount
            strColumnName = "b_str" & CStr(inum)  '拼班组编码字段名称
            
            For j = 1 To oVoucher.BodyRows
                oVoucher.bodyText(j, strColumnName) = Rs.Fields("b_str1").Value
                oVoucher.bodyText(j, "b_str" & CStr(84 + i)) = Rs.Fields("b_str2").Value '班组名称(b_str85 --- b_str104)
            Next
            inum = inum + 3 '每个班组字段之间差3,班组一字段名为b_str6,班组二字段名称为b_str9,以此类推
            
            Rs.MoveNext '指针移到下一条记录
        Next
    End If
    
ExitSetClassGroup:
    SetClassGroup = Err.Description
End Function
'根据派工计划单生成派工单
Public Function CreateTaskBill(DBconn As ADODB.Connection, domHead As Object, Optional strUserErr As String, Optional bsuc As Boolean) As String
    On Error GoTo ExitCreateTaskBill

    Dim iID  '派工计划单主表ID
    Dim oDomBody As New DOMDocument  '子表dom
    Dim j As Long
    
    iID = GetHeadItemValue(domHead, "id") '得到主表ID，根据主表ID从库中取出子表的记录
    
    strSql = " select * from V_EF_Plan_Tasks where id = " & iID
    
    If Rs.State <> 0 Then Rs.Close
    Rs.CursorLocation = adUseClient
    Rs.Open strSql, DBconn.ConnectionString, 3, 4
    
    If Rs.EOF Then
        strUserErr = "无法生成派工单！"
        bsuc = False
        Exit Function
    Else

    End If
    
ExitCreateTaskBill:
    CreateTaskBill = Err.Description
End Function
'在正品报工单审核通过后，把相关信息写入系统报工单
Public Function CreateSystemReportWork(oLogin As Object, DBconn As ADODB.Connection, strFlag As String, Cardnumber As String, domHead As Object, Optional strUserErr As String, Optional bsuc As Boolean) As String
    On Error GoTo ExitCreateSystemReportWork
    
    Dim L As Long
    Dim iID As Long  '正品报工单主表ID
    Dim iMID As Long  '系统报工单主表ID
    Dim strBillCode As String '系统报工单主表单据号
    Dim RsDetail As New ADODB.Recordset '实现从表数据操作的记录集
    Dim strSqlDetail As String   '操作子表的SQL语句
    Dim sErr As String
    Dim strIsPJ  As String  '是否品检
    
    '--------------生成新的系统报工单主表ID-------------------'
    strSql = " select isnull(Max(mid),0)  as mid from fc_moroutingbill "
    
    If Rs.State <> 0 Then Rs.Close
    Rs.CursorLocation = adUseClient
    Rs.Open strSql, DBconn.ConnectionString, 3, 4
    
    If Rs.EOF Then
        iMID = 1
    Else
        iMID = Rs.Fields("mid").Value + 1
    End If
    '------------------------------------------------------'
    '--------------生成新的系统报工单主表单据号-------------------'
    strSql = " select isnull( Max(cvouchcode),'0') as cvouchcode from fc_moroutingbill "
    
    If Rs.State <> 0 Then Rs.Close
    Rs.CursorLocation = adUseClient
    Rs.Open strSql, DBconn.ConnectionString, 3, 4
    
    If Rs.EOF Then
        strBillCode = "0000000001"
    Else
        strBillCode = Right("0000000000" & CStr(CDec(Rs.Fields("cvouchcode").Value) + 1), 10)
    End If
    
'    Dim objBillNo As New UFBillComponent.clsBillComponent
'    If objBillNo.InitBill(DBconn.ConnectionString, "FC91") = False Then
''        ErrMsg = "初始化单据号码失败!"
''        GetVouchNO = False
''        Exit Function
'    End If
'
'    strBillCode = objBillNo.GetNumber(objBillNo.GetBillFormat, True)
    
    '------------------------------------------------------'
    
    iID = GetHeadItemValue(domHead, "id") '从界面上取出正品派工单主表ID
    
    
    '生成系统派工单主表SQL语句
    If strFlag = "f" Then '废品报工单生单
        If LCase(Cardnumber) = LCase("YXEF9131") Then
            If GetHeadItemValue(domHead, "str14") = "是" Then '如果是品检的印刷废品报工单，则不生成系统报工单
                bsuc = True
                Exit Function
            End If
        End If
        strSql = " insert into fc_moroutingbill (mid,cvouchcode,cvouchdate,createuser,createdate,define2,define4,define5,define6,define7,define15,define16,wcid,vt_id,issingle ,createtime) " & vbCrLf & _
                 " select top 1 " & CStr(iMID) & ", '" & strBillCode & "',a.datetime1 ,'" & oLogin.cUserName & "', " & vbCrLf & _
                 " CONVERT(varchar(100), GETDATE(), 23),rtrim('f' + convert(char(20),a.id)),'1900-01-01 00:00:00:000',0,'1900-01-01 00:00:00:000',0,0,0,c.WcId ,31062,1,getdate() from EF_Dust_Reportedwork a " & vbCrLf & _
                 " left join sfc_moroutingdetail c on c.MoRoutingDId = a.int10 where a.id = " & iID
    Else '正品报工单生单
        If LCase(Cardnumber) = LCase("YXEF9115") Then
            If GetHeadItemValue(domHead, "str14") = "是" Then '如果是品检的印刷正品报工单，则不生成系统报工单
                bsuc = True
                Exit Function
            End If
        End If
        strSql = " insert into fc_moroutingbill (mid,cvouchcode,cvouchdate,createuser,createdate,define2,define4,define5,define6,define7,define15,define16,wcid,vt_id,issingle ,createtime) " & vbCrLf & _
                 " select top 1 " & CStr(iMID) & ", '" & strBillCode & "',a.datetime1 ,'" & oLogin.cUserName & "', " & vbCrLf & _
                 " CONVERT(varchar(100), GETDATE(), 23),rtrim('z' + convert(char(20),a.id)),'1900-01-01 00:00:00:000',0,'1900-01-01 00:00:00:000',0,0,0,c.WcId ,31062,1,getdate() from EF_Reportedwork a " & vbCrLf & _
                 " left join sfc_moroutingdetail c on c.MoRoutingDId = a.int10 where a.id = " & iID
    End If
    
    
    sErr = Update(strSql, DBconn)
    
    
    If sErr <> "" Then
        strUserErr = "系统报工单生成失败！" & sErr
        bsuc = False
        Exit Function
    Else
         '写日志表
         cls_Public.WrtDBlog DBconn, oLogin.cUserId, "生工序报工ccw", strSql
        
        If strFlag = "f" Then '废品报工单生单
             strSql = " insert into fc_moroutingbilldetail (mid,wcid,moid,modid,moroutingdid,moroutingshiftid,opseq,opcode ,opdescription," & vbCrLf & _
                "InOpUnitCode,resid1,resid2,resid3,resid4,resid5,define26,define27,define34,define35,Define28,ScrapQty,define22,inchangerate,WorkHrOp) " & vbCrLf & _
                " select top 1 " & CStr(iMID) & ", c.WcId,a.int7 ,a.int8,int10,0,a.str11,d.opcode,d.Description,c.AuxUnitCode,0,0,0,0,0,0,0,0,0," & vbCrLf & _
                " a.t_cdepcode,(select sum(b_float3) from EF_Dust_Reportedworks where id = a.id ),a.str13," & vbCrLf & _
                "( case isnull(c.ChangeRate,0) when 0 then 1 else c.changerate end) ,a.float4 from EF_Dust_Reportedwork a " & vbCrLf & _
                " left join sfc_moroutingdetail c on c.MoRoutingDId = a.int10 " & vbCrLf & _
                " left join sfc_operation d on d.OperationId = c.OperationId where a.id = " & iID

        Else '正品报工单生单
        
            strSql = " insert into fc_moroutingbilldetail (mid,wcid,moid,modid,moroutingdid,moroutingshiftid,opseq,opcode ,opdescription," & vbCrLf & _
                    "InOpUnitCode,resid1,resid2,resid3,resid4,resid5,define26,define27,define34,define35,Define28,qualifiedqty,define22,inchangerate,WorkHrOp) " & vbCrLf & _
                    " select top 1 " & CStr(iMID) & ", c.WcId,a.int7 ,a.int8,int10,0,a.str11,d.opcode,d.Description,c.AuxUnitCode,0,0,0,0,0,0,0,0,0," & vbCrLf & _
                    " a.t_cdepcode,(select sum(b_float3) from ef_reportedworks where id = a.id ),a.str12," & vbCrLf & _
                    "( case isnull(c.ChangeRate,0) when 0 then 1 else c.changerate end) ,a.float4 from EF_Reportedwork a " & vbCrLf & _
                    " left join sfc_moroutingdetail c on c.MoRoutingDId = a.int10 " & vbCrLf & _
                    " left join sfc_operation d on d.OperationId = c.OperationId where a.id = " & iID
        End If
        
       
        sErr = Update(strSql, DBconn)
        If sErr <> "" Then
            strUserErr = "系统报工单明细生成失败！" & sErr
            bsuc = False
            Exit Function
        Else
             '写日志表
            cls_Public.WrtDBlog DBconn, oLogin.cUserId, "生工序报工明细ccw", strSql

        End If
    End If
    
ExitCreateSystemReportWork:
    CreateSystemReportWork = Err.Description
End Function
'保存数据
Public Function Update(ByVal cSql As String, DBconn As ADODB.Connection) As String
    On Error GoTo ExitUpdate
    Dim L As Integer
    '更新数据库
    
    DBconn.Execute cSql, L
    If L <> 1 Then
    
    End If
    
    Exit Function
ExitUpdate:
    Update = Err.Description
End Function

'生成系统报工单前作
Public Function Checkdata(Types As String, DBconn As ADODB.Connection, domHead As Object, strUserErr As String, Optional bsuc As Boolean) As Boolean
    Dim PBillCode As String '生产订单号
    
    Checkdata = True
    
    Select Case Types
        '正品报工单、废品报工单
        Case LCase("YXEF9104"), LCase("YXEF9115"), LCase("YXEF9116"), LCase("YXEF9117"), LCase("YXEF9118"), LCase("YXEF9119"), LCase("YXEF9120"), _
             LCase("YXEF9114"), LCase("YXEF9131"), LCase("YXEF9132"), LCase("YXEF9133"), LCase("YXEF9134"), LCase("YXEF9135"), LCase("YXEF9136")
            PBillCode = GetHeadItemValue(domHead, "str1") '从界面上取出生产订单号
            If IsNull(PBillCode) Or PBillCode = "" Then '判断有没有对应的生产订单号，有则生系统报工单
                bsuc = False
                strUserErr = "没有对应的生产订单号，请检查!"
                Checkdata = False
            Else
                Checkdata = True
            End If

            '判断印刷正品、废品报工单是否是品检工序，是，则不生单
            If LCase(Types) = LCase("YXEF9115") Or LCase(Types) = LCase("YXEF9131") Then
                If GetHeadItemValue(domHead, "str14") = "是" Then
                    bsuc = False
'                    strUserErr = "没有对应的生产订单号，请检查!"
                    Checkdata = False
                Else
                    Checkdata = True
                End If
            End If
        Case Else
    End Select
End Function
