Attribute VB_Name = "mdlSubFuction"
'��ģ����Ҫ��ʵ�ַ����ƻ��ı�����֤
Option Explicit

Dim i As Integer 'ѭ������
Dim strSaleListID As String '���۶����ӱ�ID
Dim strSql As String 'SQL���
Dim strFlag As String '�д򿪡��رյı�ʶ
Dim Rs As New ADODB.Recordset '���ݼ�
Dim strAutoID As String '�����ƻ����ӱ�id
Dim iSendQuantity As Single  '����������
Dim iAllQuantity As Single  '����������
Dim iCloseQty As Single  '�رն����ѷ�������
Dim iCurrentQuantity As Single '���η�������
Dim iSLimit As Single  '�����䳬����
Dim strOperFlag As String  '���ݲ�����־��M���޸ģ�A������
Dim iBeforeModify As Single '�޸�ǰĳ�������ƻ����ķ��������������ݿ���ȡ����
Dim strInvCode As String  '��Ʒ����
Dim strInvName As String ' ��Ʒ����
Public deps As String   '����
Public banci As String  '���
Public cls_Public As Object
 

'�����ƻ���������֤����
Public Function SendPlan_SaveCheck(DBconn As ADODB.Connection, domHead As DOMDocument, domBody As DOMDocument, Optional strUserErr As String, Optional bsuc As Boolean) As String
    On Error GoTo ExitSavecheck
    
    For i = 0 To domBody.selectNodes("//z:row").Length - 1
        strFlag = GetBodyItemValue(domBody, "b_int4", i)
        
        '�����ϵĴ򿪻�д�����ݿ��ֵΪ0���رյ�ֵΪ1
        Select Case strFlag
            Case "��"
                domBody.selectNodes("//z:row").Item(i).Attributes.getNamedItem("b_int4").nodeValue = "0"
            Case "�ر�"
                domBody.selectNodes("//z:row").Item(i).Attributes.getNamedItem("b_int4").nodeValue = "1"
        End Select
        
        strSaleListID = GetBodyItemValue(domBody, "b_int5", i)
        strInvCode = GetBodyItemValue(domBody, "b_cinvcode", i)
        strInvName = GetBodyItemValue(domBody, "b_str9", i)
        '�ж����۶����ӱ�ID��û�п�ֵ
        If strSaleListID = "" Or IsNull(strSaleListID) Then
            strUserErr = "��" & CStr(i + 1) & "��" & strInvCode & strInvName & "û�в������۶���������"
            bsuc = False
            Exit Function
        End If
        
        'ȡ�������ƻ����ӱ�ID
        strAutoID = GetBodyItemValue(domBody, "autoid", i)
        If IsNull(strAutoID) Or strAutoID = "" Then
            strAutoID = "0"
        End If
        '�жϽ����е����۶����ӱ�ID�Ƿ���ڲ�����û�б��ر�
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
            strUserErr = "��" & CStr(i + 1) & "��" & strInvCode & strInvName & "û�в������۶���������"
            bsuc = False
            Exit Function
        Else
            iSendQuantity = Rs.Fields("SendCount") 'ȡ�������۶�����ID�ܵļƻ�����������
            iCloseQty = Rs.Fields("CloseQty") 'ȡ�������۶�����ID���������ƻ��ѹرյ����˷�����������
            iAllQuantity = GetBodyItemValue(domBody, "b_float5", i) 'ȡ�������۶�����ID��������
            iCurrentQuantity = GetBodyItemValue(domBody, "b_float4", i) 'ȡ������ID���η�������
            iSLimit = Rs.Fields("SLimit")  '�����䳬����
            iSLimit = 1 + iSLimit
            
             '�ж��ۼƷ��������Ƿ���ڶ�������
            strOperFlag = GetBodyItemValue(domBody, "editprop", i) '���ӡ��޸ı�־
            
            iSendQuantity = iSendQuantity + iCloseQty '�ѷ�����������δ�ر����������ѹرշ����ƻ��ѷ�������
            
            Select Case strOperFlag
                Case "M" '�޸�
                    iBeforeModify = Rs.Fields("BMCount") 'ȡ������ID�޸�ǰ��������
                    If iAllQuantity * iSLimit < (iCurrentQuantity - iBeforeModify) + iSendQuantity Then
                        strUserErr = "��" & CStr(i + 1) & "�е�Ԥ�������������������ʳ����ޣ�"
                        bsuc = False
                        Exit Function
                    End If
                Case "A" '����
                    If iAllQuantity * iSLimit < iSendQuantity + iCurrentQuantity Then
                        strUserErr = "��" & CStr(i + 1) & "�е�Ԥ�������������������ʳ����ޣ�"
                        bsuc = False
                        Exit Function
                    End If
            End Select
        End If
    Next
ExitSavecheck:
    SendPlan_SaveCheck = Err.Description
End Function
'���û�̨��Ϣ
Public Function SetMachineStation(DBconn As ADODB.Connection, strDepName As String, oVoucher As Object, Optional strUserErr As String, Optional bsuc As Boolean) As String
    On Error GoTo ExitSetMachineStation

    Dim inum As Long  '��̨����ֶ��������
    Dim strColumnName As String '�ֶ�����
    Dim L As Long
    
    If Rs.State <> 0 Then Rs.Close
    
    strSql = " select b_str2,isnull(b_str3,'') as b_str3  from V_EF_Machines where b_cdepcode = '" & deps & "' order by b_str2 "
    
    Rs.CursorLocation = adUseClient
    Rs.Open strSql, DBconn.ConnectionString, 3, 4
    
    If Rs.EOF Then
        strUserErr = strDepName & "û�ж�Ӧ�Ļ�̨��"
        bsuc = False
        Exit Function
    Else
        inum = 5  '��̨һ�ֶ����ƴ�b_str5��ʼ
        For i = 1 To Rs.RecordCount
            strColumnName = "b_str" & CStr(inum)
            For L = 1 To oVoucher.BodyRows
                oVoucher.bodyText(L, strColumnName) = Rs.Fields("b_str2").Value '��̨����
                oVoucher.bodyText(L, "b_str" & CStr(64 + i)) = Rs.Fields("b_str3").Value  '��̨����(b_str65 --- b_str84)
            Next
            inum = inum + 3 'ÿ����̨�ֶ�֮���3,��̨һ�ֶ���Ϊb_str5,��̨���ֶ�����Ϊb_str8,�Դ�����
            
            Rs.MoveNext 'ָ���Ƶ���һ����¼
        Next
    End If
    
ExitSetMachineStation:
    SetMachineStation = Err.Description
End Function
'���ò��Ż�̨������Ϣ
Public Function SetClassGroup(DBconn As ADODB.Connection, strDepName As String, oVoucher As Object, Optional strUserErr As String, Optional bsuc As Boolean) As String
    On Error GoTo ExitSetClassGroup

    Dim inum As Long  '��������ֶ����
    Dim strColumnName As String '�ֶ�����
    Dim j As Long

    If Rs.State <> 0 Then Rs.Close
    
    strSql = " select distinct b_str1,isnull(b_str2,'') as b_str2 from V_list_EF_dep_Bz where t_cdepcode = '" & deps & "' order by b_str1 "
    
    Rs.CursorLocation = adUseClient
    Rs.Open strSql, DBconn.ConnectionString, 3, 4
    
    If Rs.EOF Then
        strUserErr = strDepName & "û�ж�Ӧ�İ��飡"
        bsuc = False
        Exit Function
    Else
        inum = 6  '���Ż�̨����һ�ֶ����ƴ�b_str6��ʼ
        For i = 1 To Rs.RecordCount
            strColumnName = "b_str" & CStr(inum)  'ƴ��������ֶ�����
            
            For j = 1 To oVoucher.BodyRows
                oVoucher.bodyText(j, strColumnName) = Rs.Fields("b_str1").Value
                oVoucher.bodyText(j, "b_str" & CStr(84 + i)) = Rs.Fields("b_str2").Value '��������(b_str85 --- b_str104)
            Next
            inum = inum + 3 'ÿ�������ֶ�֮���3,����һ�ֶ���Ϊb_str6,������ֶ�����Ϊb_str9,�Դ�����
            
            Rs.MoveNext 'ָ���Ƶ���һ����¼
        Next
    End If
    
ExitSetClassGroup:
    SetClassGroup = Err.Description
End Function
'�����ɹ��ƻ��������ɹ���
Public Function CreateTaskBill(DBconn As ADODB.Connection, domHead As Object, Optional strUserErr As String, Optional bsuc As Boolean) As String
    On Error GoTo ExitCreateTaskBill

    Dim iID  '�ɹ��ƻ�������ID
    Dim oDomBody As New DOMDocument  '�ӱ�dom
    Dim j As Long
    
    iID = GetHeadItemValue(domHead, "id") '�õ�����ID����������ID�ӿ���ȡ���ӱ�ļ�¼
    
    strSql = " select * from V_EF_Plan_Tasks where id = " & iID
    
    If Rs.State <> 0 Then Rs.Close
    Rs.CursorLocation = adUseClient
    Rs.Open strSql, DBconn.ConnectionString, 3, 4
    
    If Rs.EOF Then
        strUserErr = "�޷������ɹ�����"
        bsuc = False
        Exit Function
    Else

    End If
    
ExitCreateTaskBill:
    CreateTaskBill = Err.Description
End Function
'����Ʒ���������ͨ���󣬰������Ϣд��ϵͳ������
Public Function CreateSystemReportWork(oLogin As Object, DBconn As ADODB.Connection, strFlag As String, Cardnumber As String, domHead As Object, Optional strUserErr As String, Optional bsuc As Boolean) As String
    On Error GoTo ExitCreateSystemReportWork
    
    Dim L As Long
    Dim iID As Long  '��Ʒ����������ID
    Dim iMID As Long  'ϵͳ����������ID
    Dim strBillCode As String 'ϵͳ�����������ݺ�
    Dim RsDetail As New ADODB.Recordset 'ʵ�ִӱ����ݲ����ļ�¼��
    Dim strSqlDetail As String   '�����ӱ��SQL���
    Dim sErr As String
    Dim strIsPJ  As String  '�Ƿ�Ʒ��
    
    '--------------�����µ�ϵͳ����������ID-------------------'
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
    '--------------�����µ�ϵͳ�����������ݺ�-------------------'
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
''        ErrMsg = "��ʼ�����ݺ���ʧ��!"
''        GetVouchNO = False
''        Exit Function
'    End If
'
'    strBillCode = objBillNo.GetNumber(objBillNo.GetBillFormat, True)
    
    '------------------------------------------------------'
    
    iID = GetHeadItemValue(domHead, "id") '�ӽ�����ȡ����Ʒ�ɹ�������ID
    
    
    '����ϵͳ�ɹ�������SQL���
    If strFlag = "f" Then '��Ʒ����������
        If LCase(Cardnumber) = LCase("YXEF9131") Then
            If GetHeadItemValue(domHead, "str14") = "��" Then '�����Ʒ���ӡˢ��Ʒ��������������ϵͳ������
                bsuc = True
                Exit Function
            End If
        End If
        strSql = " insert into fc_moroutingbill (mid,cvouchcode,cvouchdate,createuser,createdate,define2,define4,define5,define6,define7,define15,define16,wcid,vt_id,issingle ,createtime) " & vbCrLf & _
                 " select top 1 " & CStr(iMID) & ", '" & strBillCode & "',a.datetime1 ,'" & oLogin.cUserName & "', " & vbCrLf & _
                 " CONVERT(varchar(100), GETDATE(), 23),rtrim('f' + convert(char(20),a.id)),'1900-01-01 00:00:00:000',0,'1900-01-01 00:00:00:000',0,0,0,c.WcId ,31062,1,getdate() from EF_Dust_Reportedwork a " & vbCrLf & _
                 " left join sfc_moroutingdetail c on c.MoRoutingDId = a.int10 where a.id = " & iID
    Else '��Ʒ����������
        If LCase(Cardnumber) = LCase("YXEF9115") Then
            If GetHeadItemValue(domHead, "str14") = "��" Then '�����Ʒ���ӡˢ��Ʒ��������������ϵͳ������
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
        strUserErr = "ϵͳ����������ʧ�ܣ�" & sErr
        bsuc = False
        Exit Function
    Else
         'д��־��
         cls_Public.WrtDBlog DBconn, oLogin.cUserId, "�����򱨹�ccw", strSql
        
        If strFlag = "f" Then '��Ʒ����������
             strSql = " insert into fc_moroutingbilldetail (mid,wcid,moid,modid,moroutingdid,moroutingshiftid,opseq,opcode ,opdescription," & vbCrLf & _
                "InOpUnitCode,resid1,resid2,resid3,resid4,resid5,define26,define27,define34,define35,Define28,ScrapQty,define22,inchangerate,WorkHrOp) " & vbCrLf & _
                " select top 1 " & CStr(iMID) & ", c.WcId,a.int7 ,a.int8,int10,0,a.str11,d.opcode,d.Description,c.AuxUnitCode,0,0,0,0,0,0,0,0,0," & vbCrLf & _
                " a.t_cdepcode,(select sum(b_float3) from EF_Dust_Reportedworks where id = a.id ),a.str13," & vbCrLf & _
                "( case isnull(c.ChangeRate,0) when 0 then 1 else c.changerate end) ,a.float4 from EF_Dust_Reportedwork a " & vbCrLf & _
                " left join sfc_moroutingdetail c on c.MoRoutingDId = a.int10 " & vbCrLf & _
                " left join sfc_operation d on d.OperationId = c.OperationId where a.id = " & iID

        Else '��Ʒ����������
        
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
            strUserErr = "ϵͳ��������ϸ����ʧ�ܣ�" & sErr
            bsuc = False
            Exit Function
        Else
             'д��־��
            cls_Public.WrtDBlog DBconn, oLogin.cUserId, "�����򱨹���ϸccw", strSql

        End If
    End If
    
ExitCreateSystemReportWork:
    CreateSystemReportWork = Err.Description
End Function
'��������
Public Function Update(ByVal cSql As String, DBconn As ADODB.Connection) As String
    On Error GoTo ExitUpdate
    Dim L As Integer
    '�������ݿ�
    
    DBconn.Execute cSql, L
    If L <> 1 Then
    
    End If
    
    Exit Function
ExitUpdate:
    Update = Err.Description
End Function

'����ϵͳ������ǰ��
Public Function Checkdata(Types As String, DBconn As ADODB.Connection, domHead As Object, strUserErr As String, Optional bsuc As Boolean) As Boolean
    Dim PBillCode As String '����������
    
    Checkdata = True
    
    Select Case Types
        '��Ʒ����������Ʒ������
        Case LCase("YXEF9104"), LCase("YXEF9115"), LCase("YXEF9116"), LCase("YXEF9117"), LCase("YXEF9118"), LCase("YXEF9119"), LCase("YXEF9120"), _
             LCase("YXEF9114"), LCase("YXEF9131"), LCase("YXEF9132"), LCase("YXEF9133"), LCase("YXEF9134"), LCase("YXEF9135"), LCase("YXEF9136")
            PBillCode = GetHeadItemValue(domHead, "str1") '�ӽ�����ȡ������������
            If IsNull(PBillCode) Or PBillCode = "" Then '�ж���û�ж�Ӧ�����������ţ�������ϵͳ������
                bsuc = False
                strUserErr = "û�ж�Ӧ�����������ţ�����!"
                Checkdata = False
            Else
                Checkdata = True
            End If

            '�ж�ӡˢ��Ʒ����Ʒ�������Ƿ���Ʒ�칤���ǣ�������
            If LCase(Types) = LCase("YXEF9115") Or LCase(Types) = LCase("YXEF9131") Then
                If GetHeadItemValue(domHead, "str14") = "��" Then
                    bsuc = False
'                    strUserErr = "û�ж�Ӧ�����������ţ�����!"
                    Checkdata = False
                Else
                    Checkdata = True
                End If
            End If
        Case Else
    End Select
End Function
