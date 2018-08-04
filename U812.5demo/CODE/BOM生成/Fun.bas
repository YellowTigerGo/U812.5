Attribute VB_Name = "Fun"

' ������
Dim iverifystate As Integer
Dim ufts As String
Dim IsWFControlled As Boolean                              '�Ƿ�����������
Dim vstate As Integer                                      '�Ƿ������־
Dim vouchercode As String                                  '���ݺ�
Dim ireturncount As Integer                                '���ݱ��˻ش���
Dim flag As Boolean                                        '������ն�ѡ

Public strwhereVou  As String
Dim NotifySend As Object 'ҵ��֪ͨ����



'��ȡ�����־id
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

'��ȡ�����־id
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

'���ݱ���Ƿ�ɱ༭
Public Function GetbCanModifyVCode(manual As Boolean) As Boolean
    Dim oEle As IXMLDOMElement
    Dim myobj As New UFBillComponent.clsBillComponent
    Dim xmlDOMObj As New DOMDocument
    Dim strTemp As String

    '��ʼ�����ݱ�Ź���
    myobj.InitBill g_Conn, gstrCardNumber
    Set xmlDOMObj = New DOMDocument
    strTemp = myobj.GetBillFormat
    '    m_sVouchRuler = strTemp
    xmlDOMObj.loadXML strTemp

    Set oEle = xmlDOMObj.selectSingleNode("//���ݱ��")

    If LCase(oEle.getAttribute("�����ֹ��޸�")) = "true" Then manual = True

    If LCase(oEle.getAttribute("�����ֹ��޸�")) = "true" Or LCase(oEle.getAttribute("�غ��Զ���ȡ")) = "true" Then
        GetbCanModifyVCode = True
    End If

End Function

'�ֹ�����ʱ,У���Ƿ����
Public Function CheckCellValue(sql As String) As Boolean
    On Error GoTo Err_Handler

    Dim rs As New ADODB.Recordset
    rs.Open sql, g_Conn, 1, 1
    If Not rs.EOF Then
        '��ǰ������Ǳ���,��������
        '��ǰ�����������,���ر���
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

'�ͻ�����У��,У���Ƿ����
Public Function CheckCustomer(sql As String, cName As String, Address As String) As Boolean
    On Error GoTo Err_Handler

    Dim rs As New ADODB.Recordset
    rs.Open sql, g_Conn, 1, 1
    If Not rs.EOF Then
        '��ǰ������Ǳ���,��������
        '��ǰ�����������,���ر���
        strCellCode = rs("cCusCode") & ""
        strCellName = rs("cCusAbbName") & ""
        cName = rs("cCusName") & ""                        '�ͻ�����
        Address = rs("cCusAddress") & ""                   '�ͻ���ַ

        CheckCustomer = True
    Else
        strCellCode = ""
        strCellName = ""
        cName = ""                                         '�ͻ�����
        Address = ""                                       '�ͻ���ַ
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

'�ͻ�����У��,У���Ƿ����
Public Function CheckVendor(sql As String, vencode As String, venname As String) As Boolean
    On Error GoTo Err_Handler

    Dim rs As New ADODB.Recordset
    rs.Open sql, g_Conn, 1, 1
    If Not rs.EOF Then
        '��ǰ������Ǳ���,��������
        '��ǰ�����������,���ر���
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
'���й�������
Public Function GetFilterList(ologin As U8Login.clsLogin, Optional o_Filter As Object = Nothing, Optional sMenuPubFilter As String) As Boolean

    On Error GoTo Err_Handler
    
    Dim lngReturn As Long
    Dim objfltint As New UFGeneralFilter.FilterSrv
    Dim filtername, filtersa As String                     '��������

    filtername = "PD010301"
    filtersa = "EF"
    
    '11.0�б��� wangfb 2012-06-11
    Dim filterItf As New UFGeneralFilter.FilterSrv
    Dim sError As Variant
    Dim iRet As Boolean
    If o_Filter Is Nothing Then
        'ԭ��ֻ�Ǽ򵥵ĸ�ֵ1=2�������˵���Ҫ���ݷ���ȡֵ
        'strWhere = " (1=2) "
        '���ǲ˵��������õ�ֱ���˳� 11.0 �����б�̸����
        If sMenuPubFilter = "" Then
'            strWhere = " (1=2) "
            GetFilterList = True
            Exit Function
        Else
            '11.0�˵�����ֱ�Ӵ���������id��sMenuPubFilter,Ȼ����������Զ����ء�wangfb
            filterItf.InitSolutionID = sMenuPubFilter
            '11.0�˵����� bHiddenFilter��Ϊ����(Ĭ��false)���룬wangfb
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

    'by zhangwchb 20110720 ���ӹ����������Ƿ��ύ��
 

'    If strWhere = "" Then
'        strWhere = sAuth_AllList                           'Replace(sAuth_AllList, "HY_DZ_BorrowOut.id", "V_HY_DZ_BorrowOutList.id")
'    Else
'        strWhere = strWhere & " and " & sAuth_AllList      'Replace(sAuth_ALL, "HY_DZ_BorrowOut.id", "V_HY_DZ_BorrowOutList.id")
'    End If

    '�����������
    '״̬
    Call FilteriStatus(strWhere)


    GetFilterList = True

    Exit Function

Err_Handler:
    MsgBox Err.Description, vbInformation, GetString("U8.DZ.JA.Res030")

End Function

'���й�������
Public Function GetFilter(ologin As U8Login.clsLogin) As Boolean

    On Error GoTo Err_Handler

    Dim lngReturn As Long
    Dim objfltint As New UFGeneralFilter.FilterSrv
    Dim filtername, filtersa As String                     '��������

    'filtername = "ST[__]������õ�"
    filtername = "������õ�����"
    filtersa = "ST"
    
    lngReturn = objfltint.OpenFilter(ologin, "", filtername, filtersa)

    If lngReturn = False Then
        GetFilter = False
        '        strwhereVou = ""
        Exit Function
    End If

    strwhereVou = objfltint.GetSQLWhere

    '�����������
    '״̬
    Call FilteriStatus(strwhereVou)


    GetFilter = True

    Exit Function

Err_Handler:
    MsgBox Err.Description, vbInformation, GetString("U8.DZ.JA.Res030")

End Function
Private Sub FilteriStatus(ByRef str As String)
    'ȫ��

    'by zhangwchb 20110720 ���ӹ����������Ƿ��ύ��
    str = Replace(str, "And (bSubmitted = N'1')", "")
    str = Replace(str, "And (bSubmitted = N'0')", "")

    str = Replace(str, "iPrintCount", "isNull(iPrintCount,0)")

    'enum by modify
    If InStr(1, str, "ȫ��") > 0 Then
        str = Replace(str, "And (iStatus = N'ȫ��')", "")

        '    ElseIf InStr(1, str, "�½�") > 0 Then
        '        str = Replace(str, "�½�", "1")
        '
        '    ElseIf InStr(1, str, "���") > 0 Then
        '        str = Replace(str, "���", "2")
        '
        ''    ElseIf InStr(1, str, "����") > 0 Then
        ''        str = Replace(str, "����", "3")
        '
        '    ElseIf InStr(1, str, "�ر�") > 0 Then
        '        str = Replace(str, "�ر�", "4")

    End If
End Sub

'���µ�ǰҳ����
Public Sub UpdatePageCurrent(iID As Long)
    Dim sql As String
    Dim rs As New ADODB.Recordset

    '    If tmpLinkTbl <> "" Then '�������� ʱ ��ť״̬���� by zhangwchb 20110809
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



'��ȡ����Ȩ��
'constring �����ַ�
'user ����Ա
'cBusObId ҵ����� Department ���� Customer �ͻ�
'cClassCode ��Ŀ���ࣨ��Ŀר�ã�
'cFuncId ��д��˵�Ȩ��

Public Function GetRowAuth(constring As String, user As String, _
                           cBusObId As String, Optional ByVal cClassCode As String, Optional ByVal cFuncId As String) As String
    '�ж�����Ȩ��
    Dim oRow As New clsRowAuth
    Dim Ret As String

    Ret = ""


    On Error Resume Next


    If oRow.Init(constring, user, False) = False Then
        GetRowAuth = ""
        Exit Function
    End If

    '����"Department""R"
    Ret = oRow.getAuthString(cBusObId, "", cFuncId)

    GetRowAuth = Ret


    Set oRow = Nothing

End Function

'conn �����ַ�
'user ����Ա
'cFuncId ��д��˵�Ȩ��

'�˴������۶���Ϊ��
Public Function GetRowAuthAlls(conn As Connection, user As String, _
                               Optional ByVal cFuncId As String, Optional ByVal AuthID As String = "") As String


    Dim sRet As String
    Dim sql As String
    Dim strSql As String
    Dim rs As New ADODB.Recordset

    strSql = ""

    On Error Resume Next

    '�ж����۹�����Ҫ������Щ����Ȩ��
    sql = "select * from  accinformation where csysid=N'sa' and  ((cname in(N'bAuth_Dep',N'bAuth_Per',N'bAuth_Inv',N'bAuth_Cus','bAuth_Wh') and cvalue='true') or " & _
            " (cname='bMaker' and cvalue='false')) "
    If AuthID <> "" Then
        sql = sql & " and cname ='" & AuthID & "'"
    End If

    If rs.State = adStateOpen Then Set rs = Nothing
    rs.Open sql, conn, 1, 1

    Do While Not rs.EOF

        Select Case rs("cName")
                '�ͻ�
            Case "bAuth_Cus"
                sRet = GetRowAuth(conn.ConnectionString, user, "Customer", "", cFuncId)
                If sRet <> "" And Trim(sRet) <> "1=2" Then
                    strSql = strSql & " AND (isnull(cCusCode,N'')=N'' or cCusCode in (select cCusCode from customer where iId in (" & sRet & ")))"
                End If

                '����
            Case "bAuth_Dep"
                sRet = GetRowAuth(conn.ConnectionString, user, "Department", "", cFuncId)
                If sRet <> "" Then
                    strSql = strSql & " AND (isnull(cDepCode,N'')=N'' or cDepCode in (" & sRet & "))"
                End If

                '���
                '            Case "bAuth_Inv"
                '                sret = GetRowAuth(conn.ConnectionString, user, "Inventory", "", cFuncId)
                '                If sret <> "" And Trim(sret) <> "1=2"  Then
                '                    strSql = strSql & " AND cinvcode in (" & sret & ")"
                '                End If


                'ҵ��Ա
            Case "bAuth_Per"
                sRet = GetRowAuth(conn.ConnectionString, user, "Person", "", cFuncId)
                If sRet <> "" Then
                    strSql = strSql & " AND (isnull(cPersonCode,N'')=N'' or  cPersonCode in (" & sRet & "))"
                End If

                '�ֿ�
            Case "bAuth_Wh"
                sRet = GetRowAuth(conn.ConnectionString, user, "Warehouse", "", cFuncId)
                If sRet <> "" Then
                    strSql = strSql & " AND (isnull(cwhcode,N'')=N'' or cwhcode in (" & sRet & "))"
                End If

                '����Ա
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

    '��ȡ��ǰ�༭���ֶ���
    sHeadItemName = Voucher.ItemState(Index, siHeader).sFieldName

    If LCase(sHeadItemName) Like "cdefine*" Or LCase(sHeadItemName) Like "chdefine*" Then
        '��ͷ�Զ��������
        Dim oDefPro As Object
        referpara.Cancel = True
        Set oDefPro = New U8DefPro.clsDefPro
        If oDefPro.Init(False, g_oLogin.UfDbName, g_oLogin.cUserId) Then
            '0:'�ֹ�����;1:'ϵͳ����;2:'����
            Dim arr As Variant
            arr = Split(Voucher.ItemState(Index, 0).sDataRule, ",")
            '(1)�������Զ�������Դ�ڻ�������ʱ��arr(0) �ǻ��������ı�����(2)�������Զ�������Դ�ڵ���ʱ��arr(0) �ǵ��ݵ����ͣ��磺�ɹ���ⵥ(24)��
            '���ӿڣ�GetRefVal ��(1)ʱ����sCardNumber ��û��ʵ������ģ���(2)ʱ����sTableName ��û��ʵ������ģ�
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

'��ͷ ���š��ͻ����ֿ⡢ҵ��Ա������ ����
Private Function Refer_T(Voucher As Object, _
                         ByVal Index As Variant, _
                         sRet As Variant, _
                         referpara As UAPVoucherControl85.ReferParameter)

    On Error GoTo ErrHandler

    '��ȡ��ǰ�༭���ֶ���
    Dim sHeadItemName As String

    Dim btype         As Long

    Dim rst           As New ADODB.Recordset

    Dim sqlstr        As String

    sHeadItemName = LCase(Voucher.ItemState(Index, siHeader).sFieldName)

    referpara.Cancel = False

    '���ݱ����Դ��Ĳ��շ���
    '/*B*/ ���ݵ��ݱ�ͷģ������ȷ���Ƿ���Ҫ������Ŀ

    Select Case sHeadItemName

            '��λ
        Case LCase("ecustcode"), LCase("cCusAbbName")
     
            referpara.ReferMetaXML = "<Ref><RefSet bAuth='" & IIf(bCus_ControlAuth, "1", "0") & "' authFunID='W' /></Ref>"
            referpara.id = "Customer_AA"
            referpara.RetField = "ccuscode"        '= " cCusCode   like '%" & sRet & "%' or cCusname like '%" & sRet & "%' or cCusAbbName like '%" & sRet & "%' or cCusMnemCode like '%" & sRet & "%' "
            referpara.sSql = " isnull(#FN[dEndDate],'9999-12-31')>N'" & g_oLogin.CurDate & "'"

            If sAuth_CusW <> "" Then
                referpara.sSql = referpara.sSql & "  and  #FN[cCusCode] in (" & sAuth_CusW & ")"
            End If
         '��Ƶ�λ
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
            '��������
            referpara.RetField = "cdepcode"
            referpara.sSql = " isnull(#FN[dDepEndDate],'9999-12-31')>N'" & g_oLogin.CurDate & "'"

            If sAuth_depW <> "" Then
                referpara.sSql = referpara.sSql & " and #FN[cDepCode] IN (" & sAuth_depW & ") "
            End If
            '����
        Case LCase("edepmentcode"), LCase("cDepName")
            referpara.id = "Department_AA"
            referpara.ReferMetaXML = "<Ref><RefSet bAuth='" & IIf(bDep_ControlAuth, "1", "0") & "' authFunID='W' /></Ref>"
            '��������
            referpara.RetField = "cdepcode"
            referpara.sSql = " isnull(#FN[dDepEndDate],'9999-12-31')>N'" & g_oLogin.CurDate & "'"

            If sAuth_depW <> "" Then
                referpara.sSql = referpara.sSql & " and #FN[cDepCode] IN (" & sAuth_depW & ") "
            End If
            
            '����
        Case LCase("consubject"), LCase("consubname")
            referpara.id = "Department_AA"
            referpara.ReferMetaXML = "<Ref><RefSet bAuth='" & IIf(bDep_ControlAuth, "1", "0") & "' authFunID='W' /></Ref>"
            '��������
            referpara.RetField = "cdepcode"
            referpara.sSql = " isnull(#FN[dDepEndDate],'9999-12-31')>N'" & g_oLogin.CurDate & "'"

            If sAuth_depW <> "" Then
                referpara.sSql = referpara.sSql & " and #FN[cDepCode] IN (" & sAuth_depW & ") "
            End If
            
        Case "checkproscode", "checkpername"
            referpara.id = "Person_AA"
            referpara.RetField = "cpersoncode"
            referpara.sSql = " dPValidDate <='" & g_oLogin.CurDate & "' and  isnull(dPInValidDate ,'2099-12-31') >='" & g_oLogin.CurDate & "' "
          
            '����
        Case LCase("enfdepcode"), LCase("enfdepname")
            referpara.id = "Department_AA"
            referpara.ReferMetaXML = "<Ref><RefSet bAuth='" & IIf(bDep_ControlAuth, "1", "0") & "' authFunID='W' /></Ref>"
            '��������
            referpara.RetField = "cdepcode"
            referpara.sSql = " isnull(#FN[dDepEndDate],'9999-12-31')>N'" & g_oLogin.CurDate & "'"

            If sAuth_depW <> "" Then
                referpara.sSql = referpara.sSql & " and #FN[cDepCode] IN (" & sAuth_depW & ") "
            End If
            
        Case "conproscode", "conpername"
            referpara.id = "Person_AA"
            referpara.RetField = "cpersoncode"
            referpara.sSql = " dPValidDate <='" & g_oLogin.CurDate & "' and  isnull(dPInValidDate ,'2099-12-31') >='" & g_oLogin.CurDate & "' "
             
            ' ͳ�Ʒ���conproscode
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

            If Ref.StrRefInit_SetColWidth(g_oLogin, False, "", sqlstr, "ҵ�����ͱ���,ҵ����������", "2000,4000") = True Then
                Ref.Show
            End If

            If Not (Ref.recmx Is Nothing) Then
                sRet = Ref.recmx("ccode")
            End If

            Ref.SetRWAuth "", "", True
        
            ' ͳ�Ʒ���conproscode
        Case LCase("statcode"), LCase("stcname")   '"cfreightType", "MycdefineT6"
            referpara.Cancel = True
            Set Ref = New UFReferC.UFReferClient
            Ref.SetLogin g_oLogin

            Ref.SetRWAuth "", "", False

            sqlstr = "select  distinct ccode,cname from HY_FYSL_Statclass where   isnull(islevel,1)=1 "

            If Ref.StrRefInit_SetColWidth(g_oLogin, False, "", sqlstr, "ͳ�Ʒ������,ͳ�Ʒ�������", "2000,4000") = True Then
                Ref.Show
            End If

            If Not (Ref.recmx Is Nothing) Then
                sRet = Ref.recmx("ccode")
            End If

            Ref.SetRWAuth "", "", True

            '�������
        Case LCase("acccode"), LCase("accname")   '"cfreightType", "MycdefineT6"
            referpara.Cancel = True
            Set Ref = New UFReferC.UFReferClient
            Ref.SetLogin g_oLogin

            Ref.SetRWAuth "", "", False

            sqlstr = "select  distinct ccode,cname from HY_FYSL_Accounting where   isnull(islevel,1)=1 "

            If Ref.StrRefInit_SetColWidth(g_oLogin, False, "", sqlstr, "����������,�����������", "2000,4000") = True Then
                Ref.Show
            End If

            If Not (Ref.recmx Is Nothing) Then
                sRet = Ref.recmx("ccode")
            End If

            Ref.SetRWAuth "", "", True

            '��������
        Case LCase("engproperties"), LCase("procname")   '"cfreightType", "MycdefineT6"
            referpara.Cancel = True
            Set Ref = New UFReferC.UFReferClient
            Ref.SetLogin g_oLogin

            Ref.SetRWAuth "", "", False

            sqlstr = "select  distinct ccode,cname from HY_FYSL_Properties where   isnull(islevel,1)=1"

            If Ref.StrRefInit_SetColWidth(g_oLogin, False, "", sqlstr, "�������Ա���,������������", "2000,4000") = True Then
                Ref.Show
            End If

            If Not (Ref.recmx Is Nothing) Then
                sRet = Ref.recmx("ccode")
            End If

            Ref.SetRWAuth "", "", True
            
            '�ϼ�������
        Case LCase("engcode"), LCase("engname")   '"cfreightType", "MycdefineT6"
            referpara.Cancel = True
            Set Ref = New UFReferC.UFReferClient
            Ref.SetLogin g_oLogin

            Ref.SetRWAuth "", "", False

            sqlstr = "select  distinct ccode,cname from V_HY_FYSL_Contract_refer2 where  ccode not in (select engcode from HY_FYSL_Contract) and id<>'" & lngVoucherID & "'"

            If Ref.StrRefInit_SetColWidth(g_oLogin, False, "", sqlstr, "�� �� �� ��,�� �� �� ��", "2000,4000") = True Then
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

            If Ref.StrRefInit_SetColWidth(g_oLogin, False, "", sqlstr, "�� �� �� ��,�� �� �� ��", "2000,4000") = True Then
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

            If Ref.StrRefInit_SetColWidth(g_oLogin, False, "", sqlstr, "�����˷������,�����˷�������", "2000,4000") = True Then
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
            If Ref.StrRefInit_SetColWidth(g_oLogin, False, "", sqlstr, "��������,��������,��Ŀ����,��Ŀ����", "2000,2000,2000,2000") = True Then
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


    '��ͷ�Զ�����У��

    If LCase(sMetaItemName) Like "cdefine*" Or LCase(sMetaItemName) Like "chdefine*" Then
        Call DefineCheck_T(Voucher, Index, retvalue, bChanged, referpara)

    Else

        '�������,���Ʋ��ո�ֵ,���ֹ�����
        If Not referpara.rstGrid Is Nothing Then

            Call ReferCheck_T(Voucher, Index, retvalue, bChanged, referpara)

            referpara.rstGrid.Close
            Set referpara.rstGrid = Nothing

        Else
            '�ֹ�������롢����У��
            Call HandRecord_T(Voucher, Index, retvalue, bChanged, referpara)

        End If


    End If


    Exit Function

Err_Handler:
    MsgBox Err.Description, vbInformation, GetString("U8.DZ.JA.Res030")

End Function


'��ͷ�Զ�����У��

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
            'Result:Row=4792        Col=76  Content="]�����˶�����" ID=88797142-25e8-4611-be75-a8586a97d0c2
            MsgBox GetResString("U8.ST.USKCGLSQL.frmqc.01806", Array("[" & Voucher.ItemState(Index, 0).sCardFormula1)), vbOKOnly + vbInformation, STMsgTitle
            bChanged = Cancel
            Exit Function
        End If
    End If

    If Voucher.ItemState(Index, 0).bValidityCheck Then
        '0:'�ֹ�����;1:'ϵͳ����;2:'����
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
                    'Result:Row=4814        Col=78  Content="���Ϸ�,������¼�룡"   ID=6d0a4805-7f50-4a25-a795-b499fe42d6b1
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

'��ͷ����У��
'���š��ֿ⡢ҵ��Ա���ͻ������֡�����
Private Function ReferCheck_T(Voucher As Object, Index As Variant, retvalue As String, bChanged As UAPVoucherControl85.CheckRet, referpara As UAPVoucherControl85.ReferParameter)
    On Error GoTo ErrHandler

    Dim rs As ADODB.Recordset
    Dim sql As String

    Dim sMetaItemName As String
    sMetaItemName = LCase(Voucher.ItemState(Index, siHeader).sFieldName)

    If Not referpara.rstGrid.EOF Then

        '/*B*/ ���ݵ���ģ���ͷ����ȷ���Ƿ���Ҫ������Ŀ,�Լ�����Ŀ�����ơ���Сд bObjectCode bObjectName
        Select Case sMetaItemName
           
                '��λ
            Case LCase("ecustcode"), LCase("cCusAbbName")
                
                Voucher.headerText("ecustcode") = referpara.rstGrid.Fields("ccuscode")
                Voucher.headerText("cCusAbbName") = referpara.rstGrid.Fields("ccusabbname")
                 Voucher.headerText("cCusName") = referpara.rstGrid.Fields("ccusname")
                If sMetaItemName = LCase("ecustcode") Then
                    retvalue = referpara.rstGrid.Fields("ccuscode")
                Else
                    retvalue = referpara.rstGrid.Fields("ccusabbname")
                End If
             '��Ƶ�λ
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
                '����
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
                If retvalue = "��" Then
                    Voucher.headerText("supengcode") = ""
                    Voucher.headerText("supcname") = ""

                    Voucher.EnableHead "supengcode", False
                    Voucher.EnableHead "supcname", False
                    '                  Voucher.SetCurrentRow ("@AutoID=" & bodyele.getAttribute("AutoID") & "")
                ElseIf retvalue = "��" Then
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
'                    Case "�ͻ�"
'                        referpara.id = "Customer_AA"
'                        referpara.sSql = sHeadItemName & "  like '%" & sRet & "%'"
'                    Case "��Ӧ��"
'                        referpara.id = "Vendor_AA"
'                        referpara.sSql = sHeadItemName & "  like '%" & sRet & "%'"
'                    Case "����"
'                         referpara.id = "Department_AA"
'                         '��������
'                         referpara.sSql = sHeadItemName & "  like '%" & sRet & "%'"
'                    Case "��Ա"
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

    '�ֹ�����ʱ,��ҪУ���Ƿ����
    '/*B*/ ���ݵ���ģ���ͷ����ȷ���Ƿ���Ҫ������Ŀ
    Select Case Voucher.headerText("cType")
            '����
            'enum by modify
        Case "����"
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

                    '                MsgBox "����" & RetValue & "�����ڻ���û��Ȩ��,����������!", vbInformation, GetString("U8.DZ.JA.Res030")
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
            'ҵ��Ա
        Case "��Ա"
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

                    '                MsgBox "ҵ��Ա" & RetValue & "�����ڻ���û��Ȩ��,����������!", vbInformation, GetString("U8.DZ.JA.Res030")
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
            '�ͻ�
        Case "�ͻ�"
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

                    '                MsgBox "�ͻ�" & RetValue & "�����ڻ���û��Ȩ��,����������!", vbInformation, GetString("U8.DZ.JA.Res030")
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
        Case "��Ӧ��"
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

                    '                MsgBox "��Ӧ��" & RetValue & "�����ڻ���û��Ȩ��,����������!", vbInformation, GetString("U8.DZ.JA.Res030")
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

'�ֹ�¼���ͷ�ֶ�У��
'���š��ֿ⡢ҵ��Ա���ͻ�
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

    '�ֹ�����ʱ,��ҪУ���Ƿ����

    '/*B*/ ���ݵ���ģ���ͷ����ȷ���Ƿ���Ҫ������Ŀ
    Select Case sMetaItemName
            '��λ
        
        Case "ecustcode", "cCusAbbName"

            If retvalue = "" Then
                Voucher.headerText("ecustcode") = ""
                Voucher.headerText("cCusAbbName") = ""
                Voucher.headerText("cCusName") = ""
 
            Else

                sql = "select cCusCode ,cCusAbbName,cCusName,cCusAddress  from Customer where (cCusCode='" & retvalue & "' or cCusAbbName='" & retvalue & "' or cCusMnemCode='" & retvalue & "' or cCusAbbName ='" & retvalue & "' )"
                ' sql = sql & IIf(sAuth_CusW = "", "", " and cCusCode in (" & sAuth_CusW & ")")

                'Ȩ��
                '                    Authstr = GetRowAuthAlls(g_Conn, g_oLogin.cUserId, "R", "bAuth_Cus")
                '                    If Authstr <> "" Then sql = sql & Authstr

                strValue = CheckCustomer(sql, cName, Address)

                If strValue = False Then
                    ReDim varArgs(0)
                    varArgs(0) = retvalue
                    MsgBox GetStringPara("U8.DZ.JA.Res580", varArgs(0)), vbInformation, GetString("U8.DZ.JA.Res030")

                    'MsgBox "�ͻ�" & RetValue & "�����ڻ���û��Ȩ��,����������!", vbInformation, GetString("U8.DZ.JA.Res030")
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

                'Ȩ��
                '                    Authstr = GetRowAuthAlls(g_Conn, g_oLogin.cUserId, "R", "bAuth_Cus")
                '                    If Authstr <> "" Then sql = sql & Authstr

                strValue = CheckCustomer(sql, cName, Address)

                If strValue = False Then
                    ReDim varArgs(0)
                    varArgs(0) = retvalue
                    MsgBox GetStringPara("U8.DZ.JA.Res580", varArgs(0)), vbInformation, GetString("U8.DZ.JA.Res030")

                    'MsgBox "�ͻ�" & RetValue & "�����ڻ���û��Ȩ��,����������!", vbInformation, GetString("U8.DZ.JA.Res030")
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
                '                     'Ȩ��
                '                    Authstr = GetRowAuthAlls(g_Conn, g_oLogin.cUserId, "R", "bAuth_Dep")
                '                    If Authstr <> "" Then sql = sql & Authstr
                sql = sql & IIf(sAuth_depW = "", "", " and cdepcode in (" & sAuth_depW & ")")

                strValue = CheckCellValue(sql)

                If strValue = False Then
                    ReDim varArgs(0)
                    varArgs(0) = retvalue
                    MsgBox GetStringPara("U8.DZ.JA.Res560", varArgs(0)), vbInformation, GetString("U8.DZ.JA.Res030")

                    'MsgBox "����" & RetValue & "�����ڻ���û��Ȩ��,����������!", vbInformation, GetString("U8.DZ.JA.Res030")
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
                '                     'Ȩ��
                '                    Authstr = GetRowAuthAlls(g_Conn, g_oLogin.cUserId, "R", "bAuth_Dep")
                '                    If Authstr <> "" Then sql = sql & Authstr
                sql = sql & IIf(sAuth_depW = "", "", " and cdepcode in (" & sAuth_depW & ")")

                strValue = CheckCellValue(sql)

                If strValue = False Then
                    ReDim varArgs(0)
                    varArgs(0) = retvalue
                    MsgBox GetStringPara("U8.DZ.JA.Res560", varArgs(0)), vbInformation, GetString("U8.DZ.JA.Res030")

                    'MsgBox "����" & RetValue & "�����ڻ���û��Ȩ��,����������!", vbInformation, GetString("U8.DZ.JA.Res030")
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
                '                     'Ȩ��
                '                    Authstr = GetRowAuthAlls(g_Conn, g_oLogin.cUserId, "R", "bAuth_Dep")
                '                    If Authstr <> "" Then sql = sql & Authstr
                sql = sql & IIf(sAuth_depW = "", "", " and cdepcode in (" & sAuth_depW & ")")

                strValue = CheckCellValue(sql)

                If strValue = False Then
                    ReDim varArgs(0)
                    varArgs(0) = retvalue
                    MsgBox GetStringPara("U8.DZ.JA.Res560", varArgs(0)), vbInformation, GetString("U8.DZ.JA.Res030")

                    'MsgBox "����" & RetValue & "�����ڻ���û��Ȩ��,����������!", vbInformation, GetString("U8.DZ.JA.Res030")
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

                'Ȩ��
                '                    Authstr = GetRowAuthAlls(g_Conn, g_oLogin.cUserId, "R", "bAuth_Per")
                '                    If Authstr <> "" Then sql = sql & Authstr

                strValue = CheckCellValue(sql)

                If strValue = False Then
                    ReDim varArgs(0)
                    varArgs(0) = retvalue
                    MsgBox GetStringPara("U8.DZ.JA.Res570", varArgs(0)), vbInformation, GetString("U8.DZ.JA.Res030")

                    'MsgBox "ҵ��Ա" & RetValue & "�����ڻ���û��Ȩ��,����������!", vbInformation, GetString("U8.DZ.JA.Res030")
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
                '                     'Ȩ��
                '                    Authstr = GetRowAuthAlls(g_Conn, g_oLogin.cUserId, "R", "bAuth_Dep")
                '                    If Authstr <> "" Then sql = sql & Authstr
                sql = sql & IIf(sAuth_depW = "", "", " and cdepcode in (" & sAuth_depW & ")")

                strValue = CheckCellValue(sql)

                If strValue = False Then
                    ReDim varArgs(0)
                    varArgs(0) = retvalue
                    MsgBox GetStringPara("U8.DZ.JA.Res560", varArgs(0)), vbInformation, GetString("U8.DZ.JA.Res030")

                    'MsgBox "����" & RetValue & "�����ڻ���û��Ȩ��,����������!", vbInformation, GetString("U8.DZ.JA.Res030")
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

                'Ȩ��
                '                    Authstr = GetRowAuthAlls(g_Conn, g_oLogin.cUserId, "R", "bAuth_Per")
                '                    If Authstr <> "" Then sql = sql & Authstr

                strValue = CheckCellValue(sql)

                If strValue = False Then
                    ReDim varArgs(0)
                    varArgs(0) = retvalue
                    MsgBox GetStringPara("U8.DZ.JA.Res570", varArgs(0)), vbInformation, GetString("U8.DZ.JA.Res030")

                    'MsgBox "ҵ��Ա" & RetValue & "�����ڻ���û��Ȩ��,����������!", vbInformation, GetString("U8.DZ.JA.Res030")
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
                MsgBox "�����ڴ˷�����Ϣ", vbInformation, "��ʾ"
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
                MsgBox "�����ڴ˷�����Ϣ", vbInformation, "��ʾ"
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
                MsgBox "�����ڴ˷�����Ϣ", vbInformation, "��ʾ"
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
                MsgBox "�����ڴ˷�����Ϣ", vbInformation, "��ʾ"
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
                MsgBox "�����ڴ˷�����Ϣ", vbInformation, "��ʾ"
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
                MsgBox "���̱�Ų�����", vbInformation, "��ʾ"
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
                MsgBox "��Ŀ��Ų�����", vbInformation, "��ʾ"
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
                MsgBox "��Ŀ�������ȵ��Ų�����", vbInformation, "��ʾ"
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
           If Voucher.headerText("sourcetype") = "���պ�ͬ" Then
            If retvalue <= 0 Then
                MsgBox "�������С�ڵ���0,���޸�"
                Voucher.headerText("appprice") = ""

                Exit Function

            End If
          End If
            If Voucher.headerText("contype") = "��ͨ��ͬ" And Voucher.headerText("sourcetype") = "���պ�ͬ" Then
                If Null2Something(Voucher.headerText("conpaymoney")) <> "" Then
                    'If Null2Something(Voucher.headerText("totalappmoney"), 0) > Null2Something(Voucher.headerText("totalpaymoney"), 0) Then
                        If val(Null2Something(Voucher.headerText("appprice"), 0)) > val(Null2Something(Voucher.headerText("conpaymoney"), 0)) + val(Null2Something(Voucher.headerText("designmoney"), 0)) - val(Null2Something(Voucher.headerText("addesignmoney"), 0)) - val(Null2Something(Voucher.headerText("totalappmoney"), 0)) + val(numappprice) Then
                            MsgBox "������ܴ��ں�ͬ�˶���ͬ�ܶ���ͬδ������Ϊ��ͬ��- �ۼ�������-Ԥ����Ʒ� ", vbInformation, "��ʾ"
                             bChanged = Cancel
                            Exit Function

                        End If

'                    Else
'
'                        If Null2Something(Voucher.headerText("appprice"), 0) > Val(Null2Something(Voucher.headerText("conpaymoney"), 0)) + Val(Null2Something(Voucher.headerText("designmoney"), 0)) - Val(Null2Something(Voucher.headerText("totalpaymoney"), 0)) - Val(Null2Something(Voucher.headerText("addesignmoney"), 0)) + Val(numappprice) Then
'                            MsgBox "������ܴ��ں�ͬ�˶���ͬ�ܶ���ͬδ������Ϊ��ͬ��- �ۼ�������-Ԥ����Ʒ�", vbInformation, "��ʾ"
'                             bChanged = Cancel
'                            Exit Function
'
'                        End If
'                    End If
                    
                Else
               
                   ' If Null2Something(Voucher.headerText("totalappmoney"), 0) > Null2Something(Voucher.headerText("totalpaymoney"), 0) Then
                        If val(Null2Something(Voucher.headerText("appprice"), 0)) > val(Null2Something(Voucher.headerText("conmoney"), 0)) + val(Null2Something(Voucher.headerText("designmoney"), 0)) - val(Null2Something(Voucher.headerText("totalappmoney"), 0)) - val(Null2Something(Voucher.headerText("addesignmoney"), 0)) + val(numappprice) Then
                            MsgBox "������ܴ��ں�ͬ�˶���ͬ�ܶ���ͬδ������Ϊ��ͬ��- �ۼ�������-Ԥ����Ʒ� ", vbInformation, "��ʾ"
                             bChanged = Cancel
                            Exit Function

                        End If

'                    Else
'
'                        If Null2Something(Voucher.headerText("appprice"), 0) > Val(Null2Something(Voucher.headerText("conmoney"), 0)) + Val(Null2Something(Voucher.headerText("designmoney"), 0)) - Val(Null2Something(Voucher.headerText("totalpaymoney"), 0)) - Val(Null2Something(Voucher.headerText("addesignmoney"), 0)) + Val(numappprice) Then
'                            MsgBox "������ܴ��ں�ͬ�˶���ͬ�ܶ���ͬδ������Ϊ��ͬ��- �ۼ�������-Ԥ����Ʒ� ", vbInformation, "��ʾ"
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

'�������
Public Function VoucherbodyBrowUser(Voucher As Object, ByVal row As Long, ByVal Col As Long, sRet As Variant, referpara As UAPVoucherControl85.ReferParameter)

    On Error GoTo Err_Handler
    Dim oDefPro As Object
    Dim sMetaItemXML As String
    Dim sMetaXML As String
    Dim Ref As UFReferC.UFReferClient
    Dim sql As String
    Dim oRefSelect As Object                               '���ţ���ⵥ�Ų���
    Dim rstTmp As New ADODB.Recordset
    Dim sWhCode As String                                  '�ֿ����
    Dim sInvCode As String                                 '�������

    sMetaItemXML = Voucher.ItemState(Col, sibody).sFieldName

    If sMetaItemXML <> "cinvcode" And Voucher.bodyText(row, "cinvcode") & "" = "" Then
        MsgBox GetString("U8.DZ.JA.Res630"), vbInformation, GetString("U8.DZ.JA.Res030")
        sRet = ""
        referpara.Cancel = True
        Exit Function
    End If

    '�����ֿ�
    sWhCode = Voucher.bodyText(row, "cwhcode")
    '�������
    sInvCode = Voucher.bodyText(row, "cinvcode")

    If LCase(sMetaItemXML) = "cbatch" Or LCase(sMetaItemXML) = "cinvouchcode" Then    '���ź���ⵥ�Ų������账��
        '��������ⵥ�Ų�����ز���  -chenliangc
        Dim i As Integer
        Dim sFree As Collection                            '�������
        Dim errStr As String
        Dim sSql As String
        Dim strFilter As String                            '����������
        Dim iquantity As Double                            '����
        Dim iNum As Double                                 '����
        Dim iExchange As Double                            '������
        Dim sFreeName As String                            '�������ֶ���
        Dim sBatch As String                               '����

        '********************************************
        '2008-11-17
        'Ϊƥ��872��LP���������۸��ٷ�ʽ�Ĵ���
        Dim sSosId As String                               '���۶�����ID
        Dim sDemandType As String                          '���۶�������
        Dim sDemandCode As String                          '���۶��������
        Dim lDemandCode As Long                            '�����Ͷ����к�
        Dim j As Long
        'Dim domline As DOMDocument

        '********************************************
        '���۶�����ID

        Set oRefSelect = CreateObject("USCONTROL.RefSelect")    '���Ų������

        '/************���Ų�����ز�����ʼ��*******************/ chenliangc
        sSosId = Voucher.bodyText(row, "isosid")

        Call GetSoDemandType(sSosId, sDemandType, sDemandCode, g_Conn)
        If IsNumeric(sDemandCode) Then
            lDemandCode = CLng(sDemandCode)
        Else
            lDemandCode = 0
        End If

        '���ε�������
        iquantity = ConvertStrToDbl(Voucher.bodyText(row, "iquantity"))
        '������
        iExchange = ConvertStrToDbl(Voucher.bodyText(row, "iinvexchrate"))
        '����
        If iExchange = 0 Then
            iNum = 0
        Else
            iNum = iquantity / iExchange
        End If

        If Col > 0 Then
            sBatch = Voucher.bodyText(row, "cbatch")
        End If

        '�������
        Set sFree = New Collection
        For i = 1 To 10
            sFree.Add Null2Something(Voucher.bodyText(row, "cfree" & i))
        Next
        '/*****************************************************/
    End If


    '�����Զ��������
    If LCase(sMetaItemXML) Like "cdefine*" Or LCase(sMetaItemXML) Like "cbdefine*" Then

        referpara.Cancel = True
        Set oDefPro = New U8DefPro.clsDefPro
        If oDefPro.Init(False, g_oLogin.UfDbName, g_oLogin.cUserId) Then
            '0:'�ֹ�����;1:'ϵͳ����;2:'����
            Dim arr As Variant
            arr = Split(Voucher.ItemState(Col, 1).sDataRule, ",")
            '(1)�������Զ�������Դ�ڻ�������ʱ��arr(0) �ǻ��������ı�����(2)�������Զ�������Դ�ڵ���ʱ��arr(0) �ǵ��ݵ����ͣ��磺�ɹ���ⵥ(24)��
            '���ӿڣ�GetRefVal ��(1)ʱ����sCardNumber ��û��ʵ������ģ���(2)ʱ����sTableName ��û��ʵ������ģ�
            If UBound(arr) > 0 Then
                sRet = oDefPro.GetRefVal(Voucher.ItemState(Col, 1).nDataSource, sibody, Voucher.ItemState(Col, 1).sFieldName, arr(0), arr(1), arr(0), sRet, False, 5, 1)
            Else
                sRet = oDefPro.GetRefVal(Voucher.ItemState(Col, 1).nDataSource, sibody, Voucher.ItemState(Col, 1).sFieldName, Voucher.ItemState(Col, 1).sTableName, Voucher.ItemState(Col, 1).sFieldName, gstrCardNumber, sRet, False, 5, 1)
            End If
        End If



        '�������������
    ElseIf LCase(sMetaItemXML) Like "cfree*" Then
        referpara.Cancel = True
        Set oDefPro = New U8DefPro.clsDefPro
        If oDefPro.InitNew(g_oLogin, False) Then
            sRet = oDefPro.GetStruFreeRefVal(sInvCode, Voucher.ItemState(Col, 1).sFieldName, sRet, False, 5, 1)
        End If



        '��Ŀ����
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
            sRet = Ref.recmx("��Ŀ����")
            Voucher.bodyText(row, "citem_cname") = Ref.recmx("��������")
            Voucher.bodyText(row, "cItemCode") = ""
            Voucher.bodyText(row, "cName") = ""
        End If
        Ref.SetRWAuth "", "", True


        '��Ŀ����
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
        ' ����Ȳ�����Ŀ����,�򹫹����տؼ����ṩģ������,����Ҫ�Ȳ���
        ' ��Ŀ�������
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
            ' ���ǿ�и����ݿؼ����������¸�ֵ��,���²���BODYCELLCHECK�¼�.
            Voucher.bodyText(row, sMetaItemXML) = sRet
            Voucher.ProtectUnload2
        End If
        Ref.SetRWAuth "", , True





        '�������
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
        '�ֿ�
    ElseIf LCase(sMetaItemXML) = "cwhcode" Or LCase(sMetaItemXML) = "cwhname" Then

        sMetaXML = "<Ref><RefSet bAuth='" & IIf(bWareHouse_ControlAuth, "1", "0") & "' authFunID='W' bMultiSel= '0' /></Ref>"
        referpara.id = "Warehouse_AA"                      ' "Warehouse_AA"
        referpara.RetField = "cwhcode"
        referpara.sSql = " ('" & CDate(Mid(IIf(IsBlank(Null2Something(Voucher.headerText(StrdDate))), g_oLogin.CurDate, Voucher.headerText(StrdDate)), 1, 10)) & "' < isnull(dWhEndDate,'2099-12-31'))  and bProxyWh=0"
        If sAuth_WareHouseW <> "" Then
            referpara.sSql = referpara.sSql & " and (#FN[cwhcode] in (" & sAuth_WareHouseW & "))"
        End If
       
        referpara.ReferMetaXML = sMetaXML

        '��λ-
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




        '��������λ,Ĭ��ȡ��������λ
        'ֻ�й̶������ʿ����޸ĸ���������λ
    ElseIf LCase(sMetaItemXML) = "cinva_unit" Then
        sMetaXML = "<Ref><RefSet bAuth='0' bMultiSel= '0' /></Ref>"
        referpara.id = "ComputationUnit_AA"
        referpara.ReferMetaXML = sMetaXML
        referpara.sSql = " cgroupcode='" & Voucher.bodyText(row, "cGroupCode") & "' and (cComunitCode like '%" & sRet & "%' or cComunitName like '%" & sRet & "%')"

        '�۸����
    ElseIf LCase(sMetaItemXML) = "iquotedprice" Or LCase(sMetaItemXML) = "itaxunitprice" Or LCase(sMetaItemXML) = "iunitprice" Or LCase(sMetaItemXML) = "kl" Then

        referpara.RetField = sMetaItemXML
        BrowsePrice referpara, sMetaItemXML, row, Voucher, "97"    '97���۶���


        '���β��� chenliangc
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


        '871��872���ÿ�溯��ʱ�����ݵĲ������ͷ����仯
        '871�Ķ����в���Ҫ���������ͣ�872�ĸĳ��ַ��ͣ������Ҫ��������
        If gU8Version = "872" Then
            clsbill.RefBatchList sSql, errStr, sWhCode, sInvCode, sFree, sBatch, , CLng(sDemandType), sDemandCode
            '����ͨ�õ����β��ս���
            oRefSelect.Refer g_oLogin, sWhCode, sInvCode, sFree, iquantity, iNum, iExchange, , sSql & strFilter, , True, "", CLng(sDemandType), sDemandCode, ""
        Else
            clsbill.RefBatchList sSql, errStr, sWhCode, sInvCode, sFree, sBatch, , CLng(sDemandType), lDemandCode
            '����ͨ�õ����β��ս���
            oRefSelect.Refer g_oLogin, sWhCode, sInvCode, sFree, iquantity, iNum, iExchange, , sSql & strFilter, , True, "", CLng(sDemandType), lDemandCode, ""
        End If

        If Not oRefSelect.ReturnData Is Nothing Then
            Set rstTmp = oRefSelect.ReturnData
            If rstTmp.RecordCount = 1 Then
                sRet = rstTmp.Fields("����")

                '����������
                For i = 0 To rstTmp.Fields.Count - 1
                    sFreeName = IIf(IsNull(rstTmp.Fields(i).Properties("BASECOLUMNNAME")), "", rstTmp.Fields(i).Properties("BASECOLUMNNAME"))
                    If LCase(Mid(sFreeName, 1, 5)) = "cfree" Then
                        Voucher.bodyText(row, sFreeName) = Null2Something(rstTmp.Fields(i).Value)
                    End If
                    'by liwqa ������������
                    If LCase(Mid(sFreeName, 1, 14)) = "cbatchproperty" Then
                        Voucher.bodyText(row, sFreeName) = Null2Something(rstTmp.Fields(i).Value)
                    End If
                Next
                
                '�����д��벻֪Ϊ��ע���ˣ�20120702�ſ�
                Voucher.bodyText(row, "iquantity") = Null2Something(rstTmp.Fields("��������"))
                Voucher.bodyText(row, "inum") = Null2Something(rstTmp.Fields("�������"))

                Voucher.bodyText(row, "dmadedate") = Null2Something(rstTmp.Fields("��������"))
                Voucher.bodyText(row, "dvdate") = Null2Something(rstTmp.Fields("ʧЧ����"))
                Voucher.bodyText(row, "dexpirationdate") = Null2Something(rstTmp.Fields("��Ч�ڼ�����"))
                Voucher.bodyText(row, "cexpirationdate") = Null2Something(rstTmp.Fields("��Ч����"))
                
                Voucher.bodyText(row, "imassdate") = Null2Something(rstTmp.Fields("������"))
                Voucher.bodyText(row, "cmassunit") = Null2Something(rstTmp.Fields("�����ڵ�λ"))
                Voucher.bodyText(row, "iexpiratdatecalcu") = Null2Something(rstTmp.Fields("��Ч�����㷽ʽ"))



            ElseIf rstTmp.RecordCount > 1 Then
                For j = 1 To rstTmp.RecordCount
                    Voucher.DuplicatedLine row

                    '                Set domline = Voucher.GetLineDom(row)
                    '                Voucher.AddLine Voucher.BodyRows + 1
                    '                Voucher.UpdateLineData domline, Voucher.BodyRows
                    '����������
                    For i = 0 To rstTmp.Fields.Count - 1
                        sFreeName = IIf(IsNull(rstTmp.Fields(i).Properties("BASECOLUMNNAME")), "", rstTmp.Fields(i).Properties("BASECOLUMNNAME"))
                        If LCase(Mid(sFreeName, 1, 5)) = "cfree" Then
                            Voucher.bodyText(Voucher.BodyRows, sFreeName) = Null2Something(rstTmp.Fields(i).Value)
                        End If
                        'by liwqa ������������
                        If LCase(Mid(sFreeName, 1, 14)) = "cbatchproperty" Then
                            Voucher.bodyText(Voucher.BodyRows, sFreeName) = Null2Something(rstTmp.Fields(i).Value)
                        End If
                    Next

                    '����������Ϣ
                    If Not IsNull(Voucher.bodyText(Voucher.BodyRows, "cbatch")) Then
                        Voucher.bodyText(Voucher.BodyRows, "cbatch") = Null2Something(rstTmp.Fields("����"))
                    End If
                    Voucher.bodyText(Voucher.BodyRows, "iquantity") = Null2Something(rstTmp.Fields("��������"))
                    Voucher.bodyText(Voucher.BodyRows, "inum") = Null2Something(rstTmp.Fields("�������"))

                    Voucher.bodyText(Voucher.BodyRows, "dmadedate") = Null2Something(rstTmp.Fields("��������"))
                    Voucher.bodyText(Voucher.BodyRows, "dvdate") = Null2Something(rstTmp.Fields("ʧЧ����"))
                    Voucher.bodyText(Voucher.BodyRows, "dexpirationdate") = Null2Something(rstTmp.Fields("��Ч�ڼ�����"))
                    Voucher.bodyText(Voucher.BodyRows, "cexpirationdate") = Null2Something(rstTmp.Fields("��Ч����"))
                    
                    Voucher.bodyText(Voucher.BodyRows, "imassdate") = Null2Something(rstTmp.Fields("������"))
                    Voucher.bodyText(Voucher.BodyRows, "cmassunit") = Null2Something(rstTmp.Fields("�����ڵ�λ"))
                    Voucher.bodyText(Voucher.BodyRows, "iexpiratdatecalcu") = Null2Something(rstTmp.Fields("��Ч�����㷽ʽ"))
                    rstTmp.MoveNext
                Next
                Voucher.DelLine row
            End If
        End If

        '��ⵥ�Ų���-chenliangc
    ElseIf LCase(sMetaItemXML) = "cinvouchcode" Then
        referpara.Cancel = True
        If sWhCode = "" Then
            MsgBox GetString("U8.DZ.JA.Res640"), vbInformation, GetString("U8.DZ.JA.Res030")
            retvalue = ""
            Exit Function
        End If

        Dim sInVouchCode As String                         '��ⵥ��


        sInVouchCode = Voucher.bodyText(row, "cinvouchcode")


        '871��872���ÿ�溯��ʱ�����ݵĲ������ͷ����仯
        '871�Ķ����в���Ҫ���������ͣ�872�ĸĳ��ַ��ͣ������Ҫ��������
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
                Voucher.bodyText(row, "cinvouchcode") = rstTmp.Fields("��ⵥ��")
                Voucher.bodyText(row, "rdsid") = rstTmp.Fields("���ϵͳ���")

                '����������
                For i = 0 To rstTmp.Fields.Count - 1
                    sFreeName = IIf(IsNull(rstTmp.Fields(i).Properties("BASECOLUMNNAME")), "", rstTmp.Fields(i).Properties("BASECOLUMNNAME"))
                    If LCase(Mid(sFreeName, 1, 5)) = "cfree" Then
                        Voucher.bodyText(row, sFreeName) = Null2Something(rstTmp.Fields(i).Value)
                    End If
                Next

                '����������Ϣ
                If Not IsNull(Voucher.bodyText(row, "cbatch")) Then
                    Voucher.bodyText(row, "cbatch") = Null2Something(rstTmp.Fields("����"))
                End If

                Voucher.bodyText(row, "iquantity") = Null2Something(rstTmp.Fields("�������"))
                Voucher.bodyText(row, "inum") = Null2Something(rstTmp.Fields("������"))
                Voucher.bodyText(row, "dmadedate") = Null2Something(rstTmp.Fields("��������"))
                Voucher.bodyText(row, "dvdate") = Null2Something(rstTmp.Fields("ʧЧ����"))
                Voucher.bodyText(row, "dexpirationdate") = Null2Something(rstTmp.Fields("��Ч�ڼ�����"))
                Voucher.bodyText(row, "cexpirationdate") = Null2Something(rstTmp.Fields("��Ч����"))
                
                Voucher.bodyText(row, "imassdate") = Null2Something(rstTmp.Fields("������"))
                Voucher.bodyText(row, "cmassunit") = Null2Something(rstTmp.Fields("�����ڵ�λ"))
                Voucher.bodyText(row, "iexpiratdatecalcu") = Null2Something(rstTmp.Fields("��Ч�����㷽ʽ"))

            ElseIf rstTmp.RecordCount > 1 Then
                For j = 1 To rstTmp.RecordCount
                    Voucher.DuplicatedLine row
                    '                 Set domline = Voucher.GetLineDom(row)
                    '                 Voucher.AddLine Voucher.BodyRows + 1
                    '                 Voucher.UpdateLineData domline, Voucher.BodyRows
                    '����������
                    For i = 0 To rstTmp.Fields.Count - 1
                        sFreeName = IIf(IsNull(rstTmp.Fields(i).Properties("BASECOLUMNNAME")), "", rstTmp.Fields(i).Properties("BASECOLUMNNAME"))
                        If LCase(Mid(sFreeName, 1, 5)) = "cfree" Then
                            Voucher.bodyText(Voucher.BodyRows, sFreeName) = Null2Something(rstTmp.Fields(i).Value)
                        End If
                    Next

                    '����������Ϣ
                    If Not IsNull(Voucher.bodyText(Voucher.BodyRows, "cbatch")) Then
                        Voucher.bodyText(Voucher.BodyRows, "cbatch") = Null2Something(rstTmp.Fields("����"))
                    End If
                    Voucher.bodyText(Voucher.BodyRows, "iquantity") = Null2Something(rstTmp.Fields("��������"))
                    Voucher.bodyText(Voucher.BodyRows, "inum") = Null2Something(rstTmp.Fields("�������"))
                    Voucher.bodyText(Voucher.BodyRows, "dmadedate") = Null2Something(rstTmp.Fields("��������"))
                    Voucher.bodyText(Voucher.BodyRows, "dvdate") = Null2Something(rstTmp.Fields("ʧЧ����"))
                    Voucher.bodyText(Voucher.BodyRows, "dexpirationdate") = Null2Something(rstTmp.Fields("��Ч�ڼ�����"))
                    Voucher.bodyText(Voucher.BodyRows, "cexpirationdate") = Null2Something(rstTmp.Fields("��Ч����"))
                    
                    Voucher.bodyText(Voucher.BodyRows, "imassdate") = Null2Something(rstTmp.Fields("������"))
                    Voucher.bodyText(Voucher.BodyRows, "cmassunit") = Null2Something(rstTmp.Fields("�����ڵ�λ"))
                    Voucher.bodyText(Voucher.BodyRows, "iexpiratdatecalcu") = Null2Something(rstTmp.Fields("��Ч�����㷽ʽ"))
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

'�۸�ť
'���۲����б�

Public Sub PriceList(Voucher As Object)
    On Error GoTo Err_Handler

    Dim clsRefSrv As New U8RefService.IService
    Dim strError As String
    Dim referpara As UAPVoucherControl85.ReferParameter

    '�۸����
    '97 ���ۼ۸�
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


'����У��
Public Function VoucherbodyCellCheck(Voucher As Object, retvalue As Variant, bChanged As Long, ByVal r As Long, ByVal c As Long, referpara As UAPVoucherControl85.ReferParameter)
    Dim sInvCode As String
    Dim sError As String
    Dim tmpstr As String
    Dim nRow As Long
    nRow = Voucher.row
    
    '��ס����� for U8dp202764834
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

    '�Զ�����,������,��Ŀ����,��Ŀ����У��
    If LCase(sMetaItemXML) Like "cdefine*" Or LCase(sMetaItemXML) Like "cbdefine*" Or LCase(sMetaItemXML) Like "cfree*" Then

        Call VoucherbodyCellCheckDefine(Voucher, retvalue, bChanged, r, c, referpara)

    Else
        '�����Զ���������� �����������Ŀ����
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

'����У�飨�Զ���������� ��
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

    '�Զ�����
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
            '0:'�ֹ�����;1:'ϵͳ����;2:'����
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

        '������
    Else
        Set oDefPro = New U8DefPro.clsDefPro
        If oDefPro.Init(False, g_oLogin.UfDbName, g_oLogin.cUserId) Then
            If cDefValue <> "" Then iRet = oDefPro.ValidateFreeAr(Voucher.ItemState(c, 1).sFieldName, cDefValue, Voucher.ItemState(c, 1).bBuildArchives)
        End If
        If bChanged = 2 Then                               '�������
            '            Voucher.bodyText(r, "cbatch") = ""
            Voucher.bodyText(r, "cinvouchcode") = ""
            Voucher.bodyText(r, "cvouchcode") = ""
            Voucher.bodyText(r, "dmadedate") = ""
            Voucher.bodyText(r, "dvdate") = ""
            Voucher.bodyText(r, "dexpirationdate") = ""
            Voucher.bodyText(r, "cexpirationdate") = ""
        End If
    End If

    'iRet :0 У��ɹ���1 �����ɹ���-1 У�鲻�ɹ���-2 �������ɹ�(ֻ�ܷ����ĸ�ֵ)
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

'���Ϸ��� dxb
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

'���Ϸ��� dxb
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

'����У�飨�Զ���������� ֮�����Ŀ��
'�������
'��Ҫ��ֵ���ֶΣ�������ơ�������롢�����ͺ�
'               ������λ����롢���ơ���������λ���롢���ơ���������λ���ơ����롢������
'               �����ڵ�λ������������
'               1-16����Զ�����

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


    Dim sWhCode As String                                  '�ֿ����
    Dim sInvCode As String                                 '�������
    '�������
    sInvCode = Voucher.bodyText(r, "cinvcode")
    sWhCode = Voucher.bodyText(r, "cwhcode")

    If LCase(sMetaItemXML) = "cbatch" Or LCase(sMetaItemXML) = "cinvouchcode" Then    '���ź���ⵥ�Ų������账��

        '********************************************
        '2008-11-17
        'Ϊƥ��872��LP���������۸��ٷ�ʽ�Ĵ���
        Dim sSosId As String                               '���۶�����ID
        Dim sDemandType As String                          '���۶�������
        Dim sDemandCode As String                          '���۶��������
        Dim lDemandCode As Long                            '�����Ͷ����к�

        '********************************************
        '���۶�����ID

        Set oRefSelect = CreateObject("USCONTROL.RefSelect")    '���Ų������

        '/************���Ų�����ز�����ʼ��*******************/ chenliangc
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
            '�ֿ�
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

                    '                MsgBox "�ֿ�" & RetValue & "�����ڣ�����������", vbInformation, GetString("U8.DZ.JA.Res030")
                    bChanged = Cancel
                    Exit Function
                Else
                    If rs.RecordCount > 0 Then

                        '������ֿ���չ�ϵ

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

                        'name����ʱ����name,code����ʱҲҪ����code
                        'If sMetaItemXML = "cwhname" Then retvalue = rs.Fields("cwhname").Value
                        retvalue = rs.Fields(sMetaItemXML).Value
                    End If
                End If
            End If


            '��λ
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

                    '                MsgBox "��λ" & RetValue & "�����ڣ�����������", vbInformation, GetString("U8.DZ.JA.Res030")
                    bChanged = Cancel
                    Exit Function
                Else
                    If rs.RecordCount > 0 Then

                        '�������λ���չ�ϵ

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

                        'name����ʱ����name,code����ʱҲҪ����code
                        If sMetaItemXML = "cPosition2" Then retvalue = rs.Fields("cposname").Value
                        If sMetaItemXML = "cPosition" Then retvalue = rs.Fields("cposcode").Value
                    End If
                End If
            End If
        Case "cinvcode"
            Dim i As Integer
            Dim iRow As Long
            iRow = r                                       '��ǰ��

            If retvalue = "" Then
                Voucher.DelLine iRow
                retvalue = Voucher.bodyText(iRow, sMetaItemXML)
            Else
                Set rs = cInvCodeRefer(CStr(retvalue))
                If rs Is Nothing Or rs.State = 0 Then
                    ReDim varArgs(0)
                    varArgs(0) = retvalue
                    MsgBox GetStringPara("U8.DZ.JA.Res690", varArgs(0)), vbInformation, GetString("U8.DZ.JA.Res030")

                    '                 MsgBox GetString("U8.DZ.JA.Res780") & RetValue & "�����ڻ���û���������Ի���û��Ȩ�޻���ͣ�ã�����������", vbInformation, GetString("U8.DZ.JA.Res030")
                    bChanged = Cancel
                    Exit Function

                Else

                    '������ֿ��λ���չ�ϵ
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



        '��ֵ
        '���������ȸ�ֵ
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
            '���մ������ʱ,��ѡ
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

        '1�޻�����
        If Voucher.bodyText(r, "igrouptype") = 0 Then
            GoTo lsuccess
        End If

        '2�̶�������
        If Voucher.bodyText(r, "igrouptype") = 1 Then
            Voucher.bodyText(r, "inum") = Voucher.bodyText(r, "iquantity") / IIf(ConvertStrToDbl(Voucher.bodyText(r, "iinvexchrate")) = 0, 1, ConvertStrToDbl(Voucher.bodyText(r, "iinvexchrate")))

            GoTo lsuccess
        End If
        '3����������
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

        '����
    Case "inum"
        If Not myCheckNumValue2(retvalue, GetString("U8.DZ.JA.Res720"), 0) Then
            bChanged = Cancel
            Exit Function
        End If

        '1�޻�����
        If Voucher.bodyText(r, "igrouptype") = 0 Then
            GoTo lsuccess
        End If

        '2�̶�������
        If Voucher.bodyText(r, "igrouptype") = 1 Then
            If Voucher.bodyText(r, "inum") <> "" Then
                Voucher.bodyText(r, "iquantity") = Voucher.bodyText(r, "inum") * Voucher.bodyText(r, "iinvexchrate")
            Else
                Voucher.bodyText(r, "iquantity") = ""
            End If

            GoTo lsuccess
        End If

        '3����������
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
        '������
    Case "iinvexchrate"
        '           '����������
        '           If Voucher.bodyText(r, "igrouptype") <> 2 Then
        '                bChanged = Cancel
        '                Exit Function
        '           End If

        '�վ������Ŀ
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

        '�仯��ϵ
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

        '�޸ļ���,��������λ������,������,��Ŀ����,��Ŀ����
        '���ۡ�������˰����,ԭ�Һ�˰����,ԭ����˰����
        '       Case "inum", "cinva_unit", "iquantity", "iinvexchrate", "citem_class", "citemcode", _
                '            "iquotedprice", "inatunitprice", "inatmoney", "inattax", "inatsum", "itaxrate", "inatdiscount", _
                '            "itaxunitprice", "iunitprice", "imoney", "itax", "isum", "idiscount", "kl", "kl2", "dkl1", "dkl2", "fsalecost", "fsaleprice", "fcusminprice"
        '       Case "iquantity", "inum"
        '                Dim A As USERPCO.VoucherCO
        '                Dim StLogin As New USCOMMON.login
        '                A.IniLogin g_oLogin, errmsg
        '                Set StLogin = A.login
        '                A.CheckBody "0301", nOther, r, "", dombody, errmsg, domHead

        '��Ч��У�� chenliangc
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

                    Case "1":                              '"����":
                        Voucher.bodyText(r, "dexpirationdate") = CDate(Voucher.bodyText(r, "dvdate")) - DatePart("d", CDate(Voucher.bodyText(r, "dvdate")))
                        Voucher.bodyText(r, "cexpirationdate") = Format(Voucher.bodyText(r, "dexpirationdate"), "yyyy-mm")
                    Case "2":                              '"����":
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


        '���ŵ�У�� -chenliangc
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

            '�����������
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

            '871��872���ÿ�溯��ʱ�����ݵĲ������ͷ����仯
            '871�Ķ����в���Ҫ���������ͣ�872�ĸĳ��ַ��ͣ������Ҫ��������
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

            If rs.RecordCount = 0 Then                     'û���ҵ�����

                If Not mologin.Account.BatchAllowZeroOut Then

                    '                                   MsgBox GetString("U8.DZ.JA.Res780") & Voucher.bodyText(r, "cinvname") & "û���ҵ����ν�棬�����¼��" & vbCrLf, vbInformation, getstring("U8.DZ.JA.Res030")
                    '                                    retvalue = ""
                End If

                '                    'û���ҵ������ by liwq
                '                    For i = 0 To rs.Fields.Count - 1
                '                        sFreeName = IIf(IsNull(rs.Fields(i).Properties("BASECOLUMNNAME")), "", rs.Fields(i).Properties("BASECOLUMNNAME"))
                '                        If LCase(Mid(sFreeName, 1, 5)) = "cfree" Then
                '                            Voucher.bodyText(r, sFreeName) = ""
                '                        End If
                '                        'by liwqa ������������
                '                        If LCase(Mid(sFreeName, 1, 14)) = "cbatchproperty" Then
                '                            Voucher.bodyText(r, sFreeName) = ""
                '                        End If
                '                    Next

                '                             Voucher.bodyText(r, "iquantity") = Null2Something(Rs.Fields("��������"))
                '                             Voucher.bodyText(r, "inum") = Null2Something(Rs.Fields("�������"))

                Voucher.bodyText(r, "dmadedate") = ""
                Voucher.bodyText(r, "dvdate") = ""
                Voucher.bodyText(r, "dexpirationdate") = ""
                Voucher.bodyText(r, "cexpirationdate") = ""

            Else

                If rs.RecordCount = 1 Then

                    '����������
                    For i = 0 To rs.Fields.Count - 1
                        sFreeName = IIf(IsNull(rs.Fields(i).Properties("BASECOLUMNNAME")), "", rs.Fields(i).Properties("BASECOLUMNNAME"))
                        If LCase(Mid(sFreeName, 1, 5)) = "cfree" Then
                            Voucher.bodyText(r, sFreeName) = Null2Something(rs.Fields(i).Value)
                        End If
                        'by liwqa ������������
                        If LCase(Mid(sFreeName, 1, 14)) = "cbatchproperty" Then
                            Voucher.bodyText(r, sFreeName) = Null2Something(rs.Fields(i).Value)
                        End If
                    Next

                    '                             Voucher.bodyText(r, "iquantity") = Null2Something(Rs.Fields("��������"))
                    '                             Voucher.bodyText(r, "inum") = Null2Something(Rs.Fields("�������"))

                    Voucher.bodyText(r, "dmadedate") = Null2Something(rs.Fields("��������"))
                    Voucher.bodyText(r, "dvdate") = Null2Something(rs.Fields("ʧЧ����"))
                    Voucher.bodyText(r, "dexpirationdate") = Null2Something(rs.Fields("��Ч�ڼ�����"))
                    Voucher.bodyText(r, "cexpirationdate") = Null2Something(rs.Fields("��Ч����"))
                    
                    Voucher.bodyText(r, "imassdate") = Null2Something(rs.Fields("������"))
                    Voucher.bodyText(r, "cmassunit") = Null2Something(rs.Fields("�����ڵ�λ"))
                    Voucher.bodyText(r, "iexpiratdatecalcu") = Null2Something(rs.Fields("��Ч�����㷽ʽ"))

                    '��������ʹ���ж�����¼����ʾ���մ��ڡ�
                Else
                    Call VoucherbodyBrowUser(Voucher, r, c, retvalue, referpara)
                End If

            End If
            rs.Close
            Set rs = Nothing
        End If

        '����������У�� chenliangc
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

        Dim sInVouchCode As String                         '��ⵥ��
        Dim sRdsID As String                               '��ⵥ��ID


        '��ⵥ��
        If c > 0 Then
            sInVouchCode = Voucher.bodyText(r, "cinvouchcode")
        End If

        '���ε�������
        iquantity = ConvertStrToDbl(Voucher.bodyText(r, "iquantity"))

        If c > 0 Then
            If sInVouchCode <> "" Then
                Dim moBatchPst As New BatchPst
                moBatchPst.login = mologin

                '871��872���ÿ�溯��ʱ�����ݵĲ������ͷ����仯
                '871�Ķ����в���Ҫ���������ͣ�872�ĸĳ��ַ��ͣ������Ҫ��������
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
                    '                    MsgBox "�����ʹ��" & Voucher.bodyText(r, "cinvcode") & "ָ������ⵥ��" & sInVouchCode & "�����ڣ������������ⵥ�š�", vbInformation, GetString("U8.DZ.JA.Res030")
                    retvalue = ""
                    Voucher.bodyText(r, "rdsid") = ""
                Else
                    '��������ʹ��ֻ��һ����¼
                    If rs.RecordCount = 1 Then
                        Voucher.bodyText(r, "rdsid") = rs("���ϵͳ���")

                        '����������
                        For i = 0 To rs.Fields.Count - 1
                            sFreeName = IIf(IsNull(rs.Fields(i).Properties("BASECOLUMNNAME")), "", rs.Fields(i).Properties("BASECOLUMNNAME"))
                            If LCase(Mid(sFreeName, 1, 5)) = "cfree" Then
                                Voucher.bodyText(r, sFreeName) = Null2Something(rs.Fields(i).Value)
                            End If
                        Next

                        '����������Ϣ
                        If Not IsNull(Voucher.bodyText(r, "cbatch")) Then
                            Voucher.bodyText(r, "cbatch") = Null2Something(rs.Fields("����"))
                        End If


                        Voucher.bodyText(r, "iquantity") = Null2Something(rs.Fields("��������"))
                        Voucher.bodyText(r, "inum") = Null2Something(rs.Fields("�������"))

                        Voucher.bodyText(r, "dmadedate") = Null2Something(rs.Fields("��������"))
                        Voucher.bodyText(r, "dvdate") = Null2Something(rs.Fields("ʧЧ����"))
                        Voucher.bodyText(r, "dexpirationdate") = Null2Something(rs.Fields("��Ч�ڼ�����"))
                        Voucher.bodyText(r, "cexpirationdate") = Null2Something(rs.Fields("��Ч����"))
                        Voucher.bodyText(r, "cinvouchcode") = Null2Something(rs.Fields("��ⵥ��"))
                        
                        Voucher.bodyText(r, "imassdate") = Null2Something(rs.Fields("������"))
                        Voucher.bodyText(r, "cmassunit") = Null2Something(rs.Fields("�����ڵ�λ"))
                        Voucher.bodyText(r, "iexpiratdatecalcu") = Null2Something(rs.Fields("��Ч�����㷽ʽ"))

                        '��������ʹ���ж�����¼����ʾ���մ��ڡ�
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

        strGrid = " Select  cComUnitcode,ccomUnitName,ComputationUnit.cGroupCode,iChangRate,case when bMainUnit=1 then '��' else '��' end bMainUnit " & _
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


'�۸����
'��ѭ���۹���[����ѡ��]�ļ۸����
'skey �ؼ��� iquotedprice�� "iunitprice", "itaxunitprice"

Public Function BrowsePrice(referpara As UAPVoucherControl85.ReferParameter, _
                            sKey As String, _
                            row As Long, _
                            Voucher As Object, _
                            strVouchType As String)    '�������

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
    Dim myinfo As USSAServer.MyInformation                 '����ѡ��

    On Error GoTo Err_Handler

    '����ѡ����Ϣ myinfo
    Set clsSAWeb = Nothing
    Set clsSAWeb = New USSAServer.clsSystem
    clsSAWeb.Init g_oLogin
    clsSAWeb.INIMyInfor
    myinfo = clsSAWeb.SysInformation



    With Voucher

        '�������
        cCusCode = Voucher.headerText("ccuscode")
        cinvcode = Voucher.bodyText(.row, "cinvcode")
        If cinvcode = "" Then
            MsgBox GetString("U8.DZ.JA.Res630"), vbInformation, GetString("U8.DZ.JA.Res030")
            referpara.Cancel = True
            Set clsSAWeb = Nothing
            Exit Function
        End If

        '����
        strExchName = .headerText("cexch_name")
        If strExchName = "" Then
            MsgBox GetString("U8.DZ.JA.Res760"), vbInformation, GetString("U8.DZ.JA.Res030")
            referpara.Cancel = True
            Set clsSAWeb = Nothing
            Exit Function
        End If


        Select Case LCase(sKey)

                '����
            Case "iquotedprice"
                If myinfo.CostRefType = 0 Then             ''�����ۼ�
                    Select Case myinfo.CostReferVouch
                        Case 0
                            strColumnKey = "SA_REF_SaleOrder_SA"
                            strAuthKey = "17"              '���۶���

                        Case 1
                            strColumnKey = "SA_REF_Dispatchlist_SA"

                            strAuthKey = "01"              '������
                        Case 2
                            strColumnKey = "SA_REF_SaleBillVouch_SA"

                            strAuthKey = "07"              '���۷�Ʊ
                        Case 3
                            strColumnKey = "SA_REF_Quo_SA"

                            strAuthKey = "16"              '���۵�
                    End Select

                    strWhere = "cinvcode='" & cinvcode & IIf(myinfo.CostRefCustomer = True, "' and ccuscode='" & cCusCode & "'", "'") & " and cexch_name='" & strExchName & "'"

                Else                                       ''���ֱ���
                    strColumnKey = "SA_REF_InvPrice_SA"
                    strWhere = " binvalid=0 and cinvcode='" & cinvcode & "'"
                    strAuthKey = "invprice"

                End If


                '��˰���ۡ���˰����
            Case "iunitprice", "itaxunitprice"
                Select Case strVouchType
                    Case "97"
                        strColumnKey = "SA_REF_SaleOrder_SA"
                        strAuthKey = "17"                  '���۶���

                    Case "05", "06"
                        strColumnKey = "SA_REF_Dispatchlist_SA"

                        strAuthKey = "05"                  'ί�д���������

                    Case "26", "27", "28", "29"
                        strColumnKey = "SA_REF_SaleBillVouch_SA"
                        strAuthKey = "07"

                    Case "16"
                        strColumnKey = "SA_REF_Quo_SA"
                        strAuthKey = "16"

                End Select
                strWhere = "cinvcode='" & cinvcode & IIf(myinfo.CostRefCustomer = True, "' and ccuscode='" & cCusCode & "'", "'") & " and cexch_name='" & strExchName & "'"

                '����
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



'ȡ�ۣ����С����ţ�
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

    '����ѡ����Ϣ myinfo
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

'��ݼ�¼���������Ӧ��ⵥ��
Public Sub GetBatchInfoFun(Voucher As ctlVoucher, KeyCode As Integer, Shift As Integer)
    Dim sFree As Collection                                '����������
    Dim sWhCode As String                                  '�ֿ�
    Dim sInvCode As String                                 '�������
    Dim iSosID As String                                   '���۶�����ID
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
    Dim iNum As Double                                     '��������
    Dim iExchRate As Double                                '������
    Dim sBatch As String                                   '����
    Dim sFreeName As String                                '����������
    Dim sDemandType As String                              '���۶�������
    Dim sDemandCode As String                              '���۶��������
    Dim lDemandCode As Long                                '�����͵����۶����к�
    'Dim domline As DOMDocument
    Dim r As Long                                          '��¼���ݱ�������
    Dim Quantity As Double                                 ' ����
    Dim row As Long


    'ȡ�ÿ��е�DOM
    row = Voucher.row                                      '��¼��ǰ��
    Dim domEmpty As DOMDocument
    Voucher.AddLine Voucher.BodyRows + 1
    Set domEmpty = Voucher.GetLineDom
    Voucher.DelLine Voucher.BodyRows
    Voucher.row = row                                      '�ָ�����ǰ��

    If Voucher.headerText("cwhcode") = "" Then
        MsgBox GetString("U8.DZ.JA.Res640"), vbInformation, GetString("U8.DZ.JA.Res030")
        Exit Sub
    End If

    '��ݼ�Ctrl+E����Ctrl+B���Զ�ָ������
    If KeyCode = vbKeyE Or KeyCode = vbKeyB Then
        If Shift = vbCtrlMask Then

            If Voucher.rows > 1 Then
                'Set oRefSelect = New RefSelect
                Set oRefSelect = CreateObject("USCONTROL.RefSelect")

                oRefSelect.CreateAndDropTmpCurrentStock g_oLogin, True

                r = Voucher.BodyRows
                For i = 1 To r
                    'ctrl+Bָ������
                    If KeyCode = vbKeyB And i <> Voucher.row Then GoTo SearchNextBatch

                    '�����ֿ�
                    sWhCode = Voucher.headerText("cwhcode")
                    '�������
                    sInvCode = Voucher.bodyText(i, "cinvcode")
                    '���۶���������ID
                    iSosID = Voucher.bodyText(i, "isosid")

                    '����
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

                    '������
                    If Voucher.bodyText(i, "iinvexchrate") <> "" Then
                        iExchRate = CDbl(Voucher.bodyText(i, "iinvexchrate"))
                    Else
                        iExchRate = 0
                    End If


                    '�õ�������Զ���
                    Set oInventoryPst = New InventoryPst
                    oInventoryPst.login = mologin
                    oInventoryPst.Load sInvCode, moInventory

                    '���������ι���Ĵ��,�Զ�ָ������
                    If moInventory.IsBatch = True Then
                        '********************************************
                        '2008-11-17
                        'Ϊƥ��872��LP���������۸��ٷ�ʽ�Ĵ���
                        Call GetSoDemandType(iSosID, sDemandType, sDemandCode, g_Conn)
                        If IsNumeric(sDemandCode) Then
                            lDemandCode = CLng(sDemandCode)
                        Else
                            lDemandCode = 0
                        End If
                        '********************************************

                        '�������
                        Set sFree = New Collection
                        For j = 1 To 10
                            sFree.Add Null2Something(Voucher.bodyText(i, "cfree" & j))
                        Next j

                        '871��872���ÿ�溯��ʱ�����ݵĲ������ͷ����仯
                        '871�Ķ����в���Ҫ���������ͣ�872�ĸĳ��ַ��ͣ������Ҫ��������
                        If gU8Version = "872" Then
                            clsbill.RefBatchList sSql, "", sWhCode, sInvCode, sFree, sRet, False, CLng(sDemandType), sDemandCode, ""
                        Else
                            clsbill.RefBatchList sSql, "", sWhCode, sInvCode, sFree, sRet, False, CLng(sDemandType), lDemandCode, ""
                        End If

                        sSql = oRefSelect.GetAllBSQL(sSql)

                        errStr = "����ָ������"

                        '871��872���ÿ�溯��ʱ�����ݵĲ������ͷ����仯
                        '871�Ķ����в���Ҫ���������ͣ�872�ĸĳ��ַ��ͣ������Ҫ��������
                        If gU8Version = "872" Then
                            oRefSelect.AutoRefer g_oLogin, sWhCode, sInvCode, sFree, Quantity, iNum, iExchRate, RefBatch, sSql, False, True, "12", errStr, CLng(sDemandType), sDemandCode, ""
                        Else
                            oRefSelect.AutoRefer g_oLogin, sWhCode, sInvCode, sFree, Quantity, iNum, iExchRate, RefBatch, sSql, False, True, "12", errStr, CLng(sDemandType), lDemandCode, ""
                        End If

                        If Not oRefSelect.ReturnData Is Nothing Then
                            Set recRef = oRefSelect.ReturnData
                            If recRef.RecordCount = 1 Then
                                Voucher.bodyText(i, "iquantity") = recRef.Fields("��������")
                                Voucher.bodyText(i, "inum") = recRef.Fields("�������")
                                Voucher.bodyText(i, "cbatch") = recRef.Fields("����")
                                Voucher.bodyText(i, "dmadedate") = recRef.Fields("��������")
                                Voucher.bodyText(i, "dvdate") = recRef.Fields("ʧЧ����")
                                Voucher.bodyText(i, "dexpirationdate") = recRef.Fields("��Ч�ڼ�����")
                                Voucher.bodyText(i, "cexpirationdate") = recRef.Fields("��Ч����")
                                
                                Voucher.bodyText(r, "imassdate") = recRef.Fields("������")
                                Voucher.bodyText(r, "cmassunit") = recRef.Fields("�����ڵ�λ")
                                Voucher.bodyText(r, "iexpiratdatecalcu") = recRef.Fields("��Ч�����㷽ʽ")
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

                                '871��872���ÿ�溯��ʱ�����ݵĲ������ͷ����仯
                                '871�Ķ����в���Ҫ���������ͣ�872�ĸĳ��ַ��ͣ������Ҫ��������
                                If gU8Version = "872" Then
                                    oRefSelect.UpdateTmpCurrentStock g_oLogin, IIf(IsNull(recRef.Fields("��������")), 0, recRef.Fields("��������")), IIf(IsNull(recRef.Fields("�������")), 0, recRef.Fields("�������")), oRefSelect.GetAutoID(g_oLogin, sWhCode, sInvCode, sFree(1), sFree(2), sFree(3), sFree(4), sFree(5), sFree(6), sFree(7), sFree(8), sFree(9), sFree(10), recRef.Fields("����"), CLng(sDemandType), sDemandCode)
                                Else
                                    oRefSelect.UpdateTmpCurrentStock g_oLogin, IIf(IsNull(recRef.Fields("��������")), 0, recRef.Fields("��������")), IIf(IsNull(recRef.Fields("�������")), 0, recRef.Fields("�������")), oRefSelect.GetAutoID(g_oLogin, sWhCode, sInvCode, sFree(1), sFree(2), sFree(3), sFree(4), sFree(5), sFree(6), sFree(7), sFree(8), sFree(9), sFree(10), recRef.Fields("����"), CLng(sDemandType), lDemandCode)
                                End If

                            ElseIf recRef.RecordCount > 1 Then    '������δ������㣬���в����

                                While Not recRef.EOF

                                    '���Ʊ��������ݣ������������
                                    '                                            Voucher.AddLine Voucher.BodyRows + 1
                                    '                                            '���Ƶ�ǰ��
                                    '                                            Set domline = Voucher.GetLineDom(i)
                                    '                                            Voucher.UpdateLineData domline, Voucher.BodyRows
                                    Voucher.DuplicatedLine i

                                    Voucher.bodyText(Voucher.BodyRows, "iquantity") = recRef.Fields("��������")
                                    Voucher.bodyText(Voucher.BodyRows, "inum") = recRef.Fields("�������")
                                    Voucher.bodyText(Voucher.BodyRows, "cbatch") = recRef.Fields("����")
                                    Voucher.bodyText(Voucher.BodyRows, "dmadedate") = recRef.Fields("��������")
                                    Voucher.bodyText(Voucher.BodyRows, "dvdate") = recRef.Fields("ʧЧ����")
                                    Voucher.bodyText(Voucher.BodyRows, "dexpirationdate") = recRef.Fields("��Ч�ڼ�����")
                                    Voucher.bodyText(Voucher.BodyRows, "cexpirationdate") = recRef.Fields("��Ч����")
                                    
                                    Voucher.bodyText(Voucher.BodyRows, "imassdate") = recRef.Fields("������")
                                    Voucher.bodyText(Voucher.BodyRows, "cmassunit") = recRef.Fields("�����ڵ�λ")
                                    Voucher.bodyText(Voucher.BodyRows, "iexpiratdatecalcu") = recRef.Fields("��Ч�����㷽ʽ")
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


                                    '871��872���ÿ�溯��ʱ�����ݵĲ������ͷ����仯
                                    '871�Ķ����в���Ҫ���������ͣ�872�ĸĳ��ַ��ͣ������Ҫ��������
                                    If gU8Version = "872" Then
                                        oRefSelect.UpdateTmpCurrentStock g_oLogin, IIf(IsNull(recRef.Fields("��������")), 0, recRef.Fields("��������")), IIf(IsNull(recRef.Fields("�������")), 0, recRef.Fields("�������")), oRefSelect.GetAutoID(g_oLogin, sWhCode, sInvCode, sFree(1), sFree(2), sFree(3), sFree(4), sFree(5), sFree(6), sFree(7), sFree(8), sFree(9), sFree(10), recRef.Fields("����"), CLng(sDemandType), sDemandCode)
                                    Else
                                        oRefSelect.UpdateTmpCurrentStock g_oLogin, IIf(IsNull(recRef.Fields("��������")), 0, recRef.Fields("��������")), IIf(IsNull(recRef.Fields("�������")), 0, recRef.Fields("�������")), oRefSelect.GetAutoID(g_oLogin, sWhCode, sInvCode, sFree(1), sFree(2), sFree(3), sFree(4), sFree(5), sFree(6), sFree(7), sFree(8), sFree(9), sFree(10), recRef.Fields("����"), CLng(sDemandType), lDemandCode)
                                    End If

                                    recRef.MoveNext
                                Wend

                                'ɾ�������м�¼�����������ظ���¼
                                Voucher.UpdateLineData domEmpty, CLng(i)

                            End If
                        End If


                        If errStr <> "" Then
                            errStr = GetString("U8.DZ.JA.Res780") & sInvCode & errStr & vbCrLf
                        End If

                    End If                                 'ƥ��moInventory.IsBatch = True��End If
SearchNextBatch:
                Next i

                Voucher.RemoveEmptyRow
                If errStr <> "" Then
                    MsgBox errStr, vbCritical, GetString("U8.DZ.JA.Res030")
                End If

                'ɾ����ʱ��
                oRefSelect.CreateAndDropTmpCurrentStock g_oLogin, False
                Set oRefSelect = Nothing
            End If
        End If                                             'ƥ�� Shift = vbCtrlMask ��End If

        '�Զ�ָ�������ʹ����ⵥ��
    ElseIf KeyCode = vbKeyQ Or KeyCode = vbKeyO Then

        If Shift = vbCtrlMask And Voucher.rows > 1 Then

            Set oRefSelect = CreateObject("USCONTROL.RefSelect")

            For i = 1 To Voucher.rows - 1
                'ctrl+Qָ������
                If KeyCode = vbKeyQ And i <> Voucher.row Then GoTo SearchNextInVouchCode:
                '�����ֿ�
                sWhCode = Voucher.headerText("cwhcode")
                '�������
                sInvCode = Voucher.bodyText(i, "cinvcode")
                '���۶���������ID
                iSosID = Voucher.bodyText(i, "isosid")

                '��������
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
                '������
                If Voucher.bodyText(i, "iinvexchrate") <> "" Then
                    iExchRate = CDbl(Voucher.bodyText(i, "iinvexchrate"))
                Else
                    iExchRate = 0
                End If

                '�õ�������Զ���
                Set oInventoryPst = New InventoryPst
                oInventoryPst.login = mologin
                oInventoryPst.Load sInvCode, moInventory

                '�����Ǹ����ʹ��,�Զ�ָ����ⵥ��
                If moInventory.IsTrack = True Then
                    '********************************************
                    '2008-11-17
                    'Ϊƥ��872��LP���������۸��ٷ�ʽ�Ĵ���
                    Call GetSoDemandType(iSosID, sDemandType, sDemandCode, g_Conn)
                    If IsNumeric(sDemandCode) Then
                        lDemandCode = CLng(sDemandCode)
                    Else
                        lDemandCode = 0
                    End If
                    '********************************************


                    '�������
                    Set sFree = New Collection
                    For j = 1 To 10
                        sFree.Add Null2Something(Voucher.bodyText(i, "cfree" & j))
                    Next j

                    '                        ClsBill.RefInVouchList sSql, "", sWhCode, sInvCode, sFree, sRet, True, voucher.bodytext(I, "cbatch")), 0, 0, ""
                    '                        oRefSelect.AutoRefer g_oLogin, sWhCode, sInvCode, sFree, voucher.bodytext(I, "iquantity")), inum, iExchRate, RefInVouch, sSql, False, True, "12", errStr, 0, 0, ""

                    '871��872���ÿ�溯��ʱ�����ݵĲ������ͷ����仯
                    '871�Ķ����в���Ҫ���������ͣ�872�ĸĳ��ַ��ͣ������Ҫ��������
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
                                Voucher.bodyText(i, "iquantity") = Null2Something(recRef.Fields("��������"))
                                Voucher.bodyText(i, "inum") = Null2Something(recRef.Fields("�������"))
                                If Not IsNull(Voucher.bodyText(i, "cbatch")) Then
                                    Voucher.bodyText(i, "cbatch") = Null2Something(recRef.Fields("����"))
                                End If

                                '����������
                                For j = 0 To recRef.Fields.Count - 1
                                    sFreeName = IIf(IsNull(recRef.Fields(j).Properties("BASECOLUMNNAME")), "", recRef.Fields(j).Properties("BASECOLUMNNAME"))
                                    If LCase(Mid(sFreeName, 1, 5)) = "cfree" Then
                                        If Not IsNull(Voucher.bodyText(i, sFreeName)) > 0 Then
                                            Voucher.bodyText(i, sFreeName) = Null2Something(recRef.Fields(i).Value)
                                        End If
                                    End If
                                Next

                                If Not IsNull(Voucher.bodyText(i, "cinvouchcode")) > 0 Then
                                    Voucher.bodyText(i, "cinvouchcode") = Null2Something(recRef.Fields("��ⵥ��"))
                                End If

                                '�������Id
                                If Not IsNull(Voucher.bodyText(i, "rdsid")) Then
                                    Voucher.bodyText(i, "rdsid") = Null2Something(recRef.Fields("���ϵͳ���"))
                                End If


                                '�����м�¼����ⵥ�š����ϵͳ��š����Ρ���������Ϣ��ѡ�б�ǡ��������

                                Voucher.bodyText(i, "iquantity") = recRef.Fields("��������")
                                Voucher.bodyText(i, "inum") = recRef.Fields("�������")
                                Voucher.bodyText(i, "cbatch") = recRef.Fields("����")
                                Voucher.bodyText(i, "cinvouchcode") = recRef.Fields("��ⵥ��")
                                Voucher.bodyText(i, "rdsid") = recRef.Fields("���ϵͳ���")
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
                                Voucher.bodyText(i, "dmadedate") = recRef.Fields("��������")
                                Voucher.bodyText(i, "dvdate") = recRef.Fields("ʧЧ����")
                                Voucher.bodyText(i, "dexpirationdate") = recRef.Fields("��Ч�ڼ�����")
                                Voucher.bodyText(i, "cexpirationdate") = recRef.Fields("��Ч����")
                                
                                Voucher.bodyText(i, "imassdate") = recRef.Fields("������")
                                Voucher.bodyText(i, "cmassunit") = recRef.Fields("�����ڵ�λ")
                                Voucher.bodyText(i, "iexpiratdatecalcu") = recRef.Fields("��Ч�����㷽ʽ")



                            ElseIf recRef.RecordCount > 1 Then    '������δ������㣬���в����

                                While Not recRef.EOF

                                    '���Ʊ��������ݣ�����������
                                    '                                        Voucher.AddLine Voucher.BodyRows + 1
                                    '                                            '���Ƶ�ǰ��
                                    '                                        Set domline = Voucher.GetLineDom(i)
                                    '                                        Voucher.UpdateLineData domline, Voucher.BodyRows
                                    Voucher.DuplicatedLine i

                                    Voucher.bodyText(Voucher.BodyRows, "iquantity") = recRef.Fields("��������")
                                    Voucher.bodyText(Voucher.BodyRows, "inum") = recRef.Fields("�������")
                                    Voucher.bodyText(Voucher.BodyRows, "cbatch") = recRef.Fields("����")
                                    Voucher.bodyText(Voucher.BodyRows, "cinvouchcode") = recRef.Fields("��ⵥ��")
                                    Voucher.bodyText(Voucher.BodyRows, "rdsid") = recRef.Fields("���ϵͳ���")
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
                                    Voucher.bodyText(Voucher.BodyRows, "dmadedate") = recRef.Fields("��������")
                                    Voucher.bodyText(Voucher.BodyRows, "dvdate") = recRef.Fields("ʧЧ����")
                                    Voucher.bodyText(Voucher.BodyRows, "dexpirationdate") = recRef.Fields("��Ч�ڼ�����")
                                    Voucher.bodyText(Voucher.BodyRows, "cexpirationdate") = recRef.Fields("��Ч����")
                                    
                                    Voucher.bodyText(Voucher.BodyRows, "imassdate") = recRef.Fields("������")
                                    Voucher.bodyText(Voucher.BodyRows, "cmassunit") = recRef.Fields("�����ڵ�λ")
                                    Voucher.bodyText(Voucher.BodyRows, "iexpiratdatecalcu") = recRef.Fields("��Ч�����㷽ʽ")

                                    recRef.MoveNext
                                Wend

                                'ɾ�������м�¼�����������ظ���¼
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

            Voucher.RemoveEmptyRow                         '�������
            If errStr <> "" Then
                MsgBox errStr, vbInformation, GetString("U8.DZ.JA.Res030")
            End If

            Set oRefSelect = Nothing

        End If                                             'ƥ��If Shift = vbCtrlMask And voucher.Rows > 1 Then
    End If                                                 'ƥ�� KeyCode = vbKeyE ��End If

End Sub


'����
Public Sub ExecViewVerify(guid As String)
    SendShowViewMessage guid, "UFIDA.U8.Audit.AuditHistoryView"
    SendMessgeToPortal "DocQueryAuditHistory", guid
End Sub

'�����ύ�Լ����ù���������ʱ�Ĳ���
Public Sub ExecRequestAudit(guid As String)
    SendShowViewMessage guid, "UFIDA.U8.Audit.AuditViews.TreatTaskViewPart"
    SendMessgeToPortal "DocRequestAudit", guid
End Sub

'���ù���������ʱ�Ĳ���
Public Sub ExecCancelAudit(guid As String)
    SendShowViewMessage guid, "UFIDA.U8.Audit.AuditViews.TreatTaskViewPart"
    SendMessgeToPortal "DocRequestCancelAudit", guid
End Sub


''��ƽ̨����Ϣ
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

'��ƽ̨����Ϣ
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
'��������
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

' ��������ֵ�����ص�ǰ��¼�Ƿ���빤����
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
    'sViewID:="UFIDA.U8.Audit.AuditViews.TreatTaskViewPart",������ͼ,
    'sViewID:="UFIDA.U8.Audit.AuditHistoryView",�������̱��,��ʱ����
    'SHOWVIEW��ʾ��ͼ��HIDEVIEW������ͼ
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

'870 added �ж��Ƿ����ù�����
Public Function getIsWfControl(login As clsLogin, myConn As ADODB.Connection, ByRef errMsg As String, cardnumber As String) As Boolean
    Dim clsisWfCtl As Object
    Set clsisWfCtl = CreateObject("SCMWorkFlowCommon.clsWFController")
    Dim isWfCtl As Boolean
    Call clsisWfCtl.GetIsWFControlled(myConn, cardnumber, cardnumber & ".Submit", login.cIYear, login.cAcc_Id, isWfCtl, errMsg)
    getIsWfControl = isWfCtl
End Function

'12.0 added �ж��Ƿ񼤻��������
Public Function getIsWFHasActivated(login As clsLogin, myConn As ADODB.Connection, ByRef errMsg As String, cardnumber As String) As Boolean
    Dim clsisWfCtl As Object
    Set clsisWfCtl = CreateObject("SCMWorkFlowCommon.clsWFController")
    Dim isWfCtl As Boolean
    'Call clsisWfCtl.GetIsWFControlled(myConn, cardnumber, cardnumber & ".Submit", login.cIYear, login.cAcc_Id, isWfCtl, errMsg)
    Call clsisWfCtl.getIsWFHasActivated(myConn, cardnumber, cardnumber & ".Submit", isWfCtl, errMsg)
    getIsWFHasActivated = isWfCtl
End Function

'���ù�������ذ�ť
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

'�������ύ�볷��
Public Sub ExecSubmit(DoOrUndo As Boolean, table As String, pk As String, id As Long)
    Dim retDoUndoSubmit As Boolean
    Dim strErrorResId As String                            '�������Ĵ�����Ϣ870 added

    Screen.MousePointer = vbHourglass
    Call CheckSubmit(table, pk, CStr(id))

    If CBool(IsWFControlled) And ((DoOrUndo And (iverifystate = 0 Or (iverifystate = 1 And ireturncount > 0))) Or (DoOrUndo = False And iverifystate <> 0)) Then

        retDoUndoSubmit = DoUndoSubmit(DoOrUndo, gstrCardNumber, CStr(id), table, ufts, CBool(IsWFControlled), strErrorResId, vouchercode)
        If retDoUndoSubmit = False Then
            MsgBox strErrorResId, vbInformation, GetString("U8.DZ.JA.Res030")
        Else
            If DoOrUndo Then
                MsgBox GetString("U8.SA.xsglsql_2.saworkflowsrv.011"), vbInformation, GetString("U8.DZ.JA.Res030")    '"�����ύ�ɹ���"
            Else
                MsgBox GetString("U8.SA.xsglsql_2.saworkflowsrv.012"), vbInformation, GetString("U8.DZ.JA.Res030")    '�����ɹ���
            End If
        End If
    Else
        If DoOrUndo Then
            MsgBox GetString("U8.SA.xsglsql_2.saworkflowsrv.001"), vbInformation, GetString("U8.DZ.JA.Res030")    '"�õ����Ѿ��ύ����δ������������"
        Else
            MsgBox GetString("U8.SA.xsglsql_2.saworkflowsrv.002"), vbInformation, GetString("U8.DZ.JA.Res030")    '"�õ����Ѿ���������δ������������"
        End If
    End If
    Screen.MousePointer = vbDefault
End Sub

'�ύ�ͳ����Ĵ���
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

''���յ���
Public Function ReferVouch() As Boolean

    Dim frm As New frmVouchRefers
    Dim clsReferVoucher As New clsReferVoucher

    ' ���ò��������ؼ�������
    clsReferVoucher.HelpID = "0"                           '���� 10151180
    clsReferVoucher.pageSize = 20                          'Ĭ�Ϸ�ҳ��С
    clsReferVoucher.strMainKey = "ID"                      '����Ψһ��������Ϊ���ӱ����������
    clsReferVoucher.strDetailKey = " "                '�ӱ�Ψһ����
    clsReferVoucher.FrmCaption = "�ƻ����ճа���ͬ"
    clsReferVoucher.FilterKey = "�ƻ����ճа���ͬ"                '"������õ�����"                  '���������� SA26
    clsReferVoucher.FilterSubID = "ST"
isfyflg = False

    clsReferVoucher.HeadKey = "FYSL0035"               '���������Ϣ AA_ColumnDic
    clsReferVoucher.BodyKey = ""               '�ӱ������Ϣ ,������ֻ�б�ͷʱ�������ÿ�
    '����Զ��尴ť
    'clsReferVoucher.Buttons = "<root><row type='toolbar' name='tlbTest' caption='�Զ�ƥ��' index='26' /></root>"
    '����ʱ
    'strButtons ="<root><row type='toolbar' name='tlbTest' caption='{res:U8.SA.xsglsql_2.frmvouchrefers.automatching}' index='26' /></root>"
    'strButtons =ReplaceResId("<root><row type='toolbar' name='tlbTest' caption='{res:U8.SA.xsglsql_2.frmvouchrefers.automatching}' index='26' /></root>")
    clsReferVoucher.MainDataSource = "V_HY_FYSL_Payment_refer"    '��������Դ��ͼ
    clsReferVoucher.DetailDataSource = " "    '�ӱ�����Դ��ͼ
    clsReferVoucher.DefaultFilter = ""                     '"IsNULL(cVerifier,'')<>'' and IsNULL(cCloser,'')=''And IsNull(cSCloser,'')=''"           'Ĭ�Ϲ�������

    clsReferVoucher.strAuth = Replace(sMakeAuth_ALL, "HY_FYSL_Payment.id", "V_HY_FYSL_Payment_refer.id")    '����Ȩ��SQL��, ���ﲻ֧�ֲֿ�Ȩ�޿���

    clsReferVoucher.OtherFilter = ""                       '������������

    clsReferVoucher.HeadEnabled = False                    '�����Ƿ�ɱ༭
    clsReferVoucher.BodyEnabled = False                    '�ӱ��Ƿ�ɱ༭

    'clsReferVoucher.bSelectSingle = True                                           '��ͷ�Ƿ�ֻ��ȡΨһ��¼

    clsReferVoucher.bSelectSingle = False                  '��ͷ�Ƿ�ֻ��ȡΨһ��¼
    clsReferVoucher.strCheckFlds = "ccode"                 '"cType,bObjectCode"    '
    clsReferVoucher.strCheckMsg = "ѡ���˲�ͬ�ĺ�ͬ���"
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
''���յ���
Public Function ReferVouchpro() As Boolean

    Dim frm As New frmVouchRefers
    Dim clsReferVoucher As New clsReferVoucher

    ' ���ò��������ؼ�������
    clsReferVoucher.HelpID = "0"                           '���� 10151180
    clsReferVoucher.pageSize = 20                          'Ĭ�Ϸ�ҳ��С
    clsReferVoucher.strMainKey = "ID"                      '����Ψһ��������Ϊ���ӱ����������
    clsReferVoucher.strDetailKey = " "             '�ӱ�Ψһ����
    clsReferVoucher.FrmCaption = "��ͬ������Ŀ���� "
    clsReferVoucher.FilterKey = "��ͬ������Ŀ����"                '"������õ�����"                  '���������� SA26
    clsReferVoucher.FilterSubID = "ST"


    clsReferVoucher.HeadKey = "FYSL0009"               '���������Ϣ AA_ColumnDic
    clsReferVoucher.BodyKey = ""               '�ӱ������Ϣ ,������ֻ�б�ͷʱ�������ÿ�
    '����Զ��尴ť
    'clsReferVoucher.Buttons = "<root><row type='toolbar' name='tlbTest' caption='�Զ�ƥ��' index='26' /></root>"
    '����ʱ
    'strButtons ="<root><row type='toolbar' name='tlbTest' caption='{res:U8.SA.xsglsql_2.frmvouchrefers.automatching}' index='26' /></root>"
    'strButtons =ReplaceResId("<root><row type='toolbar' name='tlbTest' caption='{res:U8.SA.xsglsql_2.frmvouchrefers.automatching}' index='26' /></root>")
    clsReferVoucher.MainDataSource = "V_HY_FYSL_Contract_refer3"    '��������Դ��ͼ
    clsReferVoucher.DetailDataSource = " "    '�ӱ�����Դ��ͼ
    clsReferVoucher.DefaultFilter = ""                     '"IsNULL(cVerifier,'')<>'' and IsNULL(cCloser,'')=''And IsNull(cSCloser,'')=''"           'Ĭ�Ϲ�������

    clsReferVoucher.strAuth = Replace(sMakeAuth_ALL, "HY_FYSL_Contract.id", "V_HY_FYSL_Contract_refer3.id")    '����Ȩ��SQL��, ���ﲻ֧�ֲֿ�Ȩ�޿���

    clsReferVoucher.OtherFilter = ""                       '������������

    clsReferVoucher.HeadEnabled = False                    '�����Ƿ�ɱ༭
    clsReferVoucher.BodyEnabled = False                    '�ӱ��Ƿ�ɱ༭

    'clsReferVoucher.bSelectSingle = True                                           '��ͷ�Ƿ�ֻ��ȡΨһ��¼

    clsReferVoucher.bSelectSingle = False                  '��ͷ�Ƿ�ֻ��ȡΨһ��¼
    clsReferVoucher.strCheckFlds = "ccode"                 '"cType,bObjectCode"    '
    clsReferVoucher.strCheckMsg = "ѡ���˲�ͬ����Ŀ���"
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

''���յ���
Public Function ReferVoucheng() As Boolean

    Dim frm As New frmVouchRefers
    Dim clsReferVoucher As New clsReferVoucher

    ' ���ò��������ؼ�������
    clsReferVoucher.HelpID = "0"                           '���� 10151180
    clsReferVoucher.pageSize = 20                          'Ĭ�Ϸ�ҳ��С
    clsReferVoucher.strMainKey = "ID"                      '����Ψһ��������Ϊ���ӱ����������
    clsReferVoucher.strDetailKey = " "                '�ӱ�Ψһ����
    clsReferVoucher.FrmCaption = "�а���ͬ����"
    clsReferVoucher.FilterKey = "�а���ͬ����"                '"������õ�����"                  '���������� SA26
    clsReferVoucher.FilterSubID = "ST"
    
    isfyflg = True
    
    clsReferVoucher.HeadKey = "FYSL0035"               '���������Ϣ AA_ColumnDic
    clsReferVoucher.BodyKey = ""               '�ӱ������Ϣ ,������ֻ�б�ͷʱ�������ÿ�
    '����Զ��尴ť
    'clsReferVoucher.Buttons = "<root><row type='toolbar' name='tlbTest' caption='�Զ�ƥ��' index='26' /></root>"
    '����ʱ
    'strButtons ="<root><row type='toolbar' name='tlbTest' caption='{res:U8.SA.xsglsql_2.frmvouchrefers.automatching}' index='26' /></root>"
    'strButtons =ReplaceResId("<root><row type='toolbar' name='tlbTest' caption='{res:U8.SA.xsglsql_2.frmvouchrefers.automatching}' index='26' /></root>")
    clsReferVoucher.MainDataSource = "V_HY_FYSL_Payment_refer1"    '��������Դ��ͼ
    clsReferVoucher.DetailDataSource = " "    '�ӱ�����Դ��ͼ
    clsReferVoucher.DefaultFilter = ""                     '"IsNULL(cVerifier,'')<>'' and IsNULL(cCloser,'')=''And IsNull(cSCloser,'')=''"           'Ĭ�Ϲ�������

    clsReferVoucher.strAuth = Replace(sMakeAuth_ALL, "HY_FYSL_Payment.id", "V_HY_FYSL_Payment_refer1.id")    '����Ȩ��SQL��, ���ﲻ֧�ֲֿ�Ȩ�޿���

    clsReferVoucher.OtherFilter = ""                       '������������

    clsReferVoucher.HeadEnabled = False                    '�����Ƿ�ɱ༭
    clsReferVoucher.BodyEnabled = False                    '�ӱ��Ƿ�ɱ༭

    'clsReferVoucher.bSelectSingle = True                                           '��ͷ�Ƿ�ֻ��ȡΨһ��¼

    clsReferVoucher.bSelectSingle = False                  '��ͷ�Ƿ�ֻ��ȡΨһ��¼
    clsReferVoucher.strCheckFlds = "ccode"                 '"cType,bObjectCode"    '
    clsReferVoucher.strCheckMsg = "ѡ���˲�ͬ�ĺ�ͬ���"
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

'��������-��֯Dom
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

        If GetNodeAtrVal(ele, "iStatus") = "���" Then
            strSel = strSel + "," + GetNodeAtrVal(ele, "ID")
        Else
            ReDim varArgs(1)
            varArgs(0) = GetNodeAtrVal(ele, "cCODE")
            varArgs(1) = GetNodeAtrVal(ele, "iStatus")
            errMsg = errMsg & GetStringPara("U8.DZ.JA.Res800", varArgs(0), varArgs(1)) & vbCrLf
            '            errMsg = errMsg & "����" & GetNodeAtrVal(ele, "cCODE") & "��ǰ״̬Ϊ" & GetNodeAtrVal(ele, "iStatus") & ",�����Ƶ���" & vbCrLf
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
    rs.Save oDomBody, adPersistXML                         '�õ���ⵥ��ͷDOM�ṹ����
    rs.Close
    ExecmakeDom = True
    Set rs = Nothing
    Exit Function
ErrHandler:
    ExecmakeDom = False
    MsgBox Err.Description, vbInformation, GetString("U8.DZ.JA.Res030")
    Set rs = Nothing
End Function

'�ж���������־
Public Function SdFlg(ByVal idStr As String, ByVal uftsStr As String) As String
    Dim strSql As String
    Dim oRs As New ADODB.Recordset
    'by lg081106 �޸�
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

'�Ƶ�  ��֯Dom,���ö�Ӧ����������
Public Function ExecCreateVoucher(ByRef oDomHead As DOMDocument, ByRef oDomBody As DOMDocument, conn As Object, login As clsLogin, bill As BillType, Optional VoucherType As String) As Boolean
    Select Case bill
            'enum by modify
        Case ����
            ExecCreateVoucher = WriteSABill(oDomHead, oDomBody, conn, login, VoucherType)
        Case �ɹ�
            ExecCreateVoucher = WritePUBill(oDomHead, oDomBody, conn, login, VoucherType, "88")
        Case ���
            ExecCreateVoucher = WriteSCBill(oDomHead, oDomBody, conn, login, VoucherType, "0301")
        Case Ӧ��
            ExecCreateVoucher = WriteAPBill(oDomHead, oDomBody, conn, login, "AP", "AP04")    'Ӧ�ն�ӦAR��AR04
    End Select
End Function

'������������
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

    '�ṩ���ִ���ģʽ����ͷʹ��recordset��������ʹ��xml��������
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
            Voucher.headerText("sourcetype") = "���պ�ͬ"
             Voucher.headerText("addesignmoney") = Null2Something(rshead("addesignmoney"))
              Voucher.headerText("contype") = Null2Something(rshead("contype"))
            
              
              
             If val(Null2Something(rshead("conpaymoney"))) <> 0 Then
             
             Voucher.headerText("appprice") = val(Null2Something(rshead("conpaymoney"))) + val(Null2Something(rshead("accdesignmoney"))) - val(Null2Something(rshead("addesignmoney"))) - val(Null2Something(rshead("totalappmoney")))
             
             
             Else
             
             Voucher.headerText("appprice") = val(Null2Something(rshead("conmoney"))) + val(Null2Something(rshead("designmoney"))) - val(Null2Something(rshead("addesignmoney"))) - val(Null2Something(rshead("totalappmoney")))
             
             End If
            
          

            For i = 1 To 16

                Voucher.headerText("cdefine" & i) = Null2Something(rshead("cdefine" & i))    '��ͷ�Զ�����

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
'������������
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

    '�ṩ���ִ���ģʽ����ͷʹ��recordset��������ʹ��xml��������
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
             Voucher.headerText("sourcetype") = "������Ʒ�"
             Voucher.headerText("addesignmoney") = Null2Something(rshead("addesignmoney"))
             Voucher.headerText("contype") = Null2Something(rshead("contype"))
             
            If val(Null2Something(rshead("accdesignmoney"))) <> 0 Then
             Voucher.headerText("appprice") = val(Null2Something(rshead("addesignmoney"))) - val(Null2Something(rshead("designmoney")))
             
             Else
             Voucher.headerText("appprice") = val(Null2Something(rshead("addesignmoney"))) - val(Null2Something(rshead("accdesignmoney")))
             
             
             End If
          

            For i = 1 To 16

                Voucher.headerText("cdefine" & i) = Null2Something(rshead("cdefine" & i))    '��ͷ�Զ�����

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

'��ʼ��ȫ�ֶ������
Public Sub GlobalInit(login As clsLogin)

    'TODO:
    '������ʼ��
    gstrCardNumberlist = "PD010301"
    gstrCardNumber = ""

    '���ݱ���
    MainTable = ""                          '��������
    DetailsTable = ""                      '�����ֱ�
    HeadPKFld = "ID"                                       '���������ֶ�
    MainView = ""                         '��ͷ��ͼ
    DetailsView = ""                     '������ͼ
'    VoucherList = "EF_v_InScanMaintmp"                  '�б���ͼ
    
    strcCode = "cCode"                                     '���ݱ��
    StrcMaker = "cMaker"                                   '�Ƶ���
    StrdDate = "dDate"                                     '�������� ��������
    StrcHandler = "cHandler"                               '�����
    StrdVeriDate = "dVeriDate"                             '�������
    StrCloseUser = "CloseUser"                             '�ر���
    StrdCloseDate = "dCloseDate"                           '�ر�����
    StrIntoUser = "IntoUser"                               '������
    StrdIntoDate = "dIntoDate"                             '��������
    StriStatus = "iStatus"                                 '״̬

    Set clsInfor = CreateObject("Info_PU.ClsS_Infor")      'New Info_PU.ClsS_Infor
    Call clsInfor.Init(login)

    Set m_SysInfor = clsInfor.Information

    ' ����С��λ
    m_sQuantityFmt = "#,##0" & IIf(m_SysInfor.iQuantityBit = 0, "", ".") & GetPrecision(m_SysInfor.iQuantityBit)

    ' ����С��λ��
    m_sNumFmt = "#,##0" & IIf(m_SysInfor.iNumBit = 0, "", ".") & GetPrecision(m_SysInfor.iNumBit)

    ' ������С��λ��
    m_iExchRateFmt = "#,##0" & IIf(m_SysInfor.iExchRateBit = 0, "", ".") & GetPrecision(m_SysInfor.iExchRateBit)

    ' ˰��С��λ��
    m_iRateFmt = "#,##0" & IIf(m_SysInfor.iRateBit = 0, "", ".") & GetPrecision(m_SysInfor.iRateBit)

    ' �������С��λ(�ɹ���)������ã�
    m_sPriceFmt = "#,##0" & IIf(m_SysInfor.iCostBit = 0, "", ".") & GetPrecision(m_SysInfor.iCostBit)

    ' ��Ʊ����С��λ(������)
    m_sPriceFmtSA = "#,##0" & IIf(m_SysInfor.iBillCostBit = 0, "", ".") & GetPrecision(m_SysInfor.iBillCostBit)

End Sub

'���ݾ��ȿ���
Public Sub FormatVouchList(rs As ADODB.Recordset)

    On Error GoTo ErrHandler:


    Dim sQuantityFmt As String                             ' ����С��λ
    Dim sNumFmt As String                                  ' ����С��λ��
    Dim iExchRateFmt As String                             ' ������С��λ��
    Dim m_iRateFmt As String                               ' ˰��С��λ��
    Dim sPriceFmt As String                                ' �������С��λ(�ɹ���)������ã�
    Dim sPriceFmtSA As String                              ' ��Ʊ����С��λ(������)

    sQuantityFmt = m_SysInfor.iQuantityBit
    sNumFmt = m_SysInfor.iNumBit
    iExchRateFmt = m_SysInfor.iExchRateBit
    m_iRateFmt = m_SysInfor.iRateBit
    sPriceFmt = m_SysInfor.iCostBit
    sPriceFmtSA = m_SysInfor.iBillCostBit

    Dim DomFormat As New DOMDocument
    rs.Save DomFormat, adPersistXML
    rs.Close

    '    SetFormat DomFormat, "cfreightCost", sQuantityFmt  '�˷� ��ͷ
    '    SetFormat DomFormat, "iquantity", sQuantityFmt     '����
    '    SetFormat DomFormat, "inum", sQuantityFmt          '������
    '    SetFormat DomFormat, "iinvexchrate", iExchRateFmt  '������
    '    SetFormat DomFormat, "iQtyOutSum", sQuantityFmt       '�ۼƳ�������
    '    SetFormat DomFormat, "iQtyOut2Sum", sQuantityFmt      '
    '    SetFormat DomFormat, "iQtyBackSum", sQuantityFmt      '�ۼƹ黹����
    '    SetFormat DomFormat, "iQtyBack2Sum", sQuantityFmt
    '    SetFormat DomFormat, "iQtyCOutSum", sQuantityFmt       '�ۼ�ת�������
    '    SetFormat DomFormat, "iQtyCOut2Sum", sQuantityFmt      '
    '    SetFormat DomFormat, "iQtyCSaleSum", sQuantityFmt      '�ۼ�ת��������
    '    SetFormat DomFormat, "iQtyCSale2Sum", sQuantityFmt
    '    SetFormat DomFormat, "iQtyCFreeSum", sQuantityFmt       '�ۼ�ת��Ʒ����
    '    SetFormat DomFormat, "iQtyCFree2Sum", sQuantityFmt      '
    '    SetFormat DomFormat, "iQtyCOverSum", sQuantityFmt       '�ۼ�ת���ø�����
    '    SetFormat DomFormat, "iQtyCOver2Sum", sQuantityFmt

    '    Dim sQuantityFmt As String  ' ����С��λ
    '    Dim sNumFmt As String ' ����С��λ��
    '    Dim iExchRateFmt As String ' ������С��λ��
    '    Dim m_iRateFmt As String  ' ˰��С��λ��
    '    Dim sPriceFmt As String  ' �������С��λ(�ɹ���)������ã�
    '    Dim sPriceFmtSA  As String  ' ��Ʊ����С��λ(������)

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
    If gcCreateType = "�ڳ�����" Then
        strShowFormat = "False"
    Else
        strShowFormat = "True"
    End If
    SetFormat2 DomFormat, "iQtyOutSum", strShowFormat
    SetFormat2 DomFormat, "iQtyOut2Sum", strShowFormat

    If gcCreateType = "�ڳ�����" Then
        SetFormat3 DomFormat, "iquantity", "��������"
        SetFormat3 DomFormat, "inum", "�������"
    End If

    rs.Open DomFormat
    Set DomFormat = Nothing
    Exit Sub
ErrHandler:
    MsgBox Err.Description, vbInformation, GetString("U8.DZ.JA.Res030")
    Set DomFormat = Nothing
End Sub

'���õ���ģ��ı�ͷ������Ŀ�Ŀɼ��� dixingben 2009/5/21
Public Sub SetFormat3(DomFormat As DOMDocument, FieldName As String, FormatStr As String)
    Dim ele As IXMLDOMElement
    Set ele = DomFormat.selectSingleNode("//z:row[@FieldName='" & FieldName & "']")
    ele.setAttribute "cardformula1", FormatStr
    Set ele = Nothing
End Sub

'���õ���ģ��ı�ͷ������Ŀ�Ŀɼ��� dixingben 2009/5/21
Public Sub SetFormat2(DomFormat As DOMDocument, FieldName As String, FormatStr As String)
    Dim ele As IXMLDOMElement
    Set ele = DomFormat.selectSingleNode("//z:row[@FieldName='" & FieldName & "']")
    ele.setAttribute "ShowIt", FormatStr
    Set ele = Nothing
End Sub

'���õ���ģ���С��λ��
Public Sub SetFormat(DomFormat As DOMDocument, FieldName As String, FormatStr As String)
    Dim ele As IXMLDOMElement
    Set ele = DomFormat.selectSingleNode("//z:row[@FieldName='" & FieldName & "']")
    ele.setAttribute "NumPoint", FormatStr
    Set ele = Nothing
End Sub

'
'д�����൥��
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
    Dim lstBody As IXMLDOMNodeList                         '����ڵ��б�


    Dim sVoucherID As String                               '����ID

    Dim errMsg As String

    Dim strSql As String
    Dim txtSQL As String
    Dim i As Integer


    Dim lrows As Long
    '
    Dim cSoCodeStr As String, sqlstr As String, vidstr As String
    Dim RsTemp As ADODB.Recordset, vouchID As String, oneVouchID As String

    Dim voucherSuccSize As Integer                         '��¼���ɹ��ĵ�������  2008-01-30
    Dim voucherErrMsg As String                            '��¼����ʧ�ܵ���Ϣ����2008-01-30
    Dim exchNameStr As String
    Dim rsKL As ADODB.Recordset
    Dim CurrentId As String


    '/*********************************************************************************/'
    Dim pco As Object                                      'New VoucherCO_Sa.ClsVoucherCO_SA
    Dim clsSysSa As Object                                 'USSAServer.clsSystem

    Set pco = CreateObject("VoucherCO_Sa.ClsVoucherCO_SA")
    Set clsSysSa = CreateObject("USSAServer.clsSystem")
    '��ʼ������
    'Set clsSysSa = New USSAServer.clsSystem
    clsSysSa.Init login
    clsSysSa.INIMyInfor
    clsSysSa.bManualTrans = True                           '2008-01-10
    '��ʼ�����������ӿ�,"97"��ʾ���۶���
    'Pco.Init VoucherTypeSA.SODetails, login, conn, "CS", clsSysSa
    pco.Init VoucherType, login, conn, "CS", clsSysSa
    '/*********************************************************************************/'
    voucherSuccSize = 0
    voucherErrMsg = ""
    vouchID = ""

    '���ݲ�ͬ����������֯��ͬ��dom

    '��ͷ��recordset����,������xml��������
    r.Open oDomHead
    rsize = r.RecordCount
    For i = 1 To rsize

        rows = 0
        dateStr = Format(Now(), "YYYY-MM-DD HH:MM:SS")

        Set SAHead = New DOMDocument
        Set SABody = New DOMDocument
        ViewHead = GetViewHead(conn, "17")
        ViewBody = GetViewBody(conn, "17")

        '��ͬ���+Ԥ��������+Ԥ�깤����������ͬ�ɹ����������¼��������������
        'д��ͷ----------------------------------------------------------------------------------
        strSql = "select *,'' as editprop from " & ViewHead & " where 1>2"
        rs.Open strSql, conn, adOpenDynamic, adLockOptimistic
        rs.Save SAHead, adPersistXML                       '�õ���ⵥ��ͷDOM�ṹ����
        rs.Close
        Set oNodes = SAHead.selectSingleNode("//rs:data")
        Set oNode = SAHead.createElement("z:row")

        CurrentId = r!id                                   '��¼��ǰ�������ʾ

        oNode.setAttribute "ccuscode", Null2Something(r!cCusCode)    '�ͻ�����
        '2008-01-23 �ͻ����ƣ����˷�ʽ���룬���˷�ʽ,�����������룬��������,������ַ,�����
        Dim ftmpData As String
        If fRs.State = adStateOpen Then fRs.Close
        Set fRs = conn.Execute("Select cCusSAProtocol,ccusname,cCusOType,cCusPayCond,cCusOAddress,cModifyPerson From Customer Where ccuscode='" & r!cCusCode & "'")
        If Not fRs.EOF Then
            oNode.setAttribute "ccusname", Null2Something(fRs!ccusname)    '�ͻ�����
            oNode.setAttribute "csccode", Null2Something(fRs!cCusOType)    '���˷�ʽ����
            oNode.setAttribute "cpaycode", Null2Something(fRs!cCusPayCond)    '������������
            oNode.setAttribute "ccusoaddress", Null2Something(fRs!ccusoaddress)    '������ַ
            oNode.setAttribute "cgatheringplan", Null2Something(fRs!cCusSAProtocol)    '����Э�����
            'oNode.setAttribute "cchanger", null2something(fRs!cModifyPerson)                     '�����
            ftmpData = Null2Something(fRs!cCusPayCond)
            '���˷�ʽ
            sqlstr = "Select cscName From ShippingChoice Where cscCode='" & fRs!cCusOType & "'"
            If fRs.State = adStateOpen Then fRs.Close
            Set fRs = conn.Execute(sqlstr)
            If Not fRs.EOF Then oNode.setAttribute "cscname", Null2Something(fRs!cscname)
            '��������
            sqlstr = "Select cPayName From PayCondition Where cPayCode='" & ftmpData & "'"
            If fRs.State = adStateOpen Then fRs.Close
            Set fRs = conn.Execute(sqlstr)
            If Not fRs.EOF Then oNode.setAttribute "cpayname", Null2Something(fRs!cPayName)
        End If
        If fRs.State = adStateOpen Then fRs.Close

        oNode.setAttribute "cbustype", "��ͨ����"              '�������� "��ͨ����"

        'ȡĬ����������,���ȡ��һ����������
        If fRs.State = adStateOpen Then fRs.Close
        Set fRs = conn.Execute("select * from saleType where bdefault=1")
        If fRs.EOF Then
            fRs.Close
            Set fRs = conn.Execute("select top1 * from saleType")
            If fRs.EOF Then
                MsgBox GetString("U8.DZ.JA.Res830"), vbInformation, GetString("U8.DZ.JA.Res030")
                Exit Function
            Else
                oNode.setAttribute "cstcode", Null2Something(fRs!cstcode)    '��������
                fRs.Close
            End If
        Else
            oNode.setAttribute "cstcode", Null2Something(fRs!cstcode)    '��������
            fRs.Close
        End If

        oNode.setAttribute "cpersoncode", Null2Something(r!cpersoncode)    'ҵ��Ա
        oNode.setAttribute "cdepcode", Null2Something(r!cDepcode)    '���ű���

        oNode.setAttribute "cdefine1", Null2Something(r!cDefine1)    '��ͷ�Զ�����
        oNode.setAttribute "cdefine2", Null2Something(r!cDefine2)    '��ͷ�Զ�����
        oNode.setAttribute "cdefine3", Null2Something(r!cDefine3)    '��ͷ�Զ�����
        oNode.setAttribute "cdefine4", Null2Something(r!cDefine4)    '��ͷ�Զ�����
        oNode.setAttribute "cdefine5", Null2Something(r!cDefine5)    '��ͷ�Զ�����
        oNode.setAttribute "cdefine6", Null2Something(r!cDefine6)    '��ͷ�Զ�����
        oNode.setAttribute "cdefine7", Null2Something(r!cdefine7)    '��ͷ�Զ�����
        oNode.setAttribute "cdefine8", Null2Something(r!cDefine8)    '��ͷ�Զ�����
        oNode.setAttribute "cdefine9", Null2Something(r!cDefine9)    '��ͷ�Զ�����
        oNode.setAttribute "cdefine10", Null2Something(r!cDefine10)    '��ͷ�Զ�����
        oNode.setAttribute "cdefine11", Null2Something(r!cDefine11)    '��ͷ�Զ�����
        oNode.setAttribute "cdefine12", Null2Something(r!cDefine12)    '��ͷ�Զ�����
        oNode.setAttribute "cdefine13", Null2Something(r!cDefine13)    '��ͷ�Զ�����
        oNode.setAttribute "cdefine14", Null2Something(r!cDefine14)    '��ͷ�Զ�����
        oNode.setAttribute "cdefine15", Null2Something(r!cDefine15)    '��ͷ�Զ�����
        oNode.setAttribute "cdefine16", Null2Something(r!cdefine16)    '��ͷ�Զ�����


        oNode.setAttribute "ivtid", GetVoucherID(conn, "17")
        oNode.setAttribute "cmaker", login.cUserName
        oNode.setAttribute "cmemo", GetString("U8.DZ.JA.Res840") & Format(Time, "HH:MM:SS")    ' & "���:"
        oNode.setAttribute "cvouchtype", "97"              '��������

        oNode.setAttribute "cexch_name", Null2Something(r!cexch_name)    '����
        oNode.setAttribute "iexchrate", Null2Something(r!iExchRate)    '����

        oNode.setAttribute "dpredatebt", ""                'Ԥ��������
        oNode.setAttribute "dpremodatebt", ""              'Ԥ�깤����
        oNode.setAttribute "ddate", dateStr

        'oNode.setAttribute "cchanger", ""           '�����
        oNode.setAttribute "ccrechpname", ""               '���������
        'oNode.setAttribute "csccode", ""            '���䷽ʽ����
        'oNode.setAttribute "cpaycode", ""           '������������
        oNode.setAttribute "istatus", ""                   '״̬
        oNode.setAttribute "cverifier", ""                 '�����

        oNode.setAttribute "itaxrate", "17"                '˰��
        oNode.setAttribute "imoney", ""                    '����
        oNode.setAttribute "ccloser", ""                   '�ر���
        oNode.setAttribute "cstname", "��ͨ����"               '������������
        oNode.setAttribute "iarmoney", "0"                 'Ӧ�����
        oNode.setAttribute "bdisflag", "0"                 '�Ƿ���������
        oNode.setAttribute "clocker", ""                   '������
        oNode.setAttribute "breturnflag", "0"              '�˻���־
        oNode.setAttribute "icuscreline", "0"              '���ö��
        oNode.setAttribute "coppcode", ""                  '�̻�����
        oNode.setAttribute "caddcode", ""                  '�ջ���ַ����
        oNode.setAttribute "iverifystate", ""              'iVerifyState ��������
        oNode.setAttribute "ireturncount", ""              'iReturnCount ��������
        oNode.setAttribute "icreditstate", ""              '��������״̬
        oNode.setAttribute "iswfcontrolled", ""            'IsWFControlled ��������
        oNode.setAttribute "editprop", "A"


        oNodes.appendChild oNode



        '����currentid�������----------------------------------------

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

                If Trim(GetNodeAtrVal(ele, "iquantity")) <> "0" Then    '�������������0��ʱ�򣬲�д����ⵥ��
                    oNode.setAttribute "cconfigstatus", "" '����
                    oNode.setAttribute "ikpquantity", ""   '�ѿ�Ʊ����
                    oNode.setAttribute "iprekeepquantity", ""    'Ԥ������
                    oNode.setAttribute "ccontractrowguid", ""    '��ͬcGuid
                    oNode.setAttribute "ccontracttagcode", ""    '��ͬ���
                    oNode.setAttribute "bproductbill", ""  '������������
                    oNode.setAttribute "ccontractid", ""   '��ͬ��
                    oNode.setAttribute "dEDate", ""        '��ҵ������ǩ������
                    oNode.setAttribute "cquocode", ""      '���۵���
                    oNode.setAttribute "iprekeeptotquantity", ""    'Ԥ��������
                    oNode.setAttribute "bTrack", ""        '�Ƿ����������
                    oNode.setAttribute "bInvBatch", ""     '�Ƿ����ι���
                    oNode.setAttribute "bspecialorder", "" '�ͻ�����ר��
                    oNode.setAttribute "ballpurchase", ""  '�Ѿ����ײɹ���־
                    oNode.setAttribute "iprekeeptotnum", "" 'Ԥ���ܼ���
                    oNode.setAttribute "bProxyForeign", "" '�Ƿ�ί��
                    oNode.setAttribute "fomquantity", ""   'ί������
                    oNode.setAttribute "fimquantity", ""   '��������
                    oNode.setAttribute "dreleasedate", ""  '�ͷ�����
                    oNode.setAttribute "iadvancedate", ""  '�ۼ���ǰ��
                    oNode.setAttribute "iprekeepnum", ""   'Ԥ������
                    oNode.setAttribute "fpurquan", ""      '�ɹ�����
                    oNode.setAttribute "iquoid", ""        '���۵�id
                    oNode.setAttribute "csrpolicy", "PE"   '��������
                    oNode.setAttribute "binvmodel", "��"    'ģ��
                    oNode.setAttribute "ippartqty", ""     'ĸ������
                    oNode.setAttribute "ippartid", ""      'ĸ������ID
                    oNode.setAttribute "ippartseqid", ""   'PTOĸ��˳���
                    oNode.setAttribute "citem_class", GetNodeAtrVal(ele, "citem_class")    '��Ŀ�������
                    oNode.setAttribute "citemcode", GetNodeAtrVal(ele, "citemcode")    '��Ŀ����
                    oNode.setAttribute "citemname", GetNodeAtrVal(ele, "citemname")    '��Ŀ����
                    oNode.setAttribute "icusbomid", ""     '�ͻ�bomid
                    oNode.setAttribute "imoquantity", ""   '�´���������
                    oNode.setAttribute "irowno", " 1"      '�����к�
                    oNode.setAttribute "binvtype", "0"     '�Ƿ��ۿ�
                    'oNode.setAttribute "iquotedprice", "0" '"1000.00"   '����
                    oNode.setAttribute "cscloser", ""      '�رձ��
                    oNode.setAttribute "bservice", "0"     '�Ƿ�Ӧ˰����
                    'oNode.setAttribute "inum", "0"                      '����
                    oNode.setAttribute "itax", "0"         '"12815.38"          'ԭ��˰��
                    oNode.setAttribute "isum", "0"         '"88200.00"          'ԭ�Ҽ�˰�ϼ�
                    oNode.setAttribute "inatsum", "0"      ' "88200.00"      '���Ҽ�˰�ϼ�
                    oNode.setAttribute "inatdiscount", "0" ' "11800.00" '�����ۿ۶�
                    oNode.setAttribute "fsaleprice", "0"   '���۽��
                    oNode.setAttribute "ikpnum", ""        '�ۼƿ�Ʊ����������
                    oNode.setAttribute "ikpmoney", ""      '�ۼ�ԭ�ҿ�Ʊ���
                    oNode.setAttribute "iinvexchrate", "0" ' "50.00"    '������
                    oNode.setAttribute "idiscount", "0"    ' "11800.00"    'ԭ���ۿ۶�
                    oNode.setAttribute "fsalecost", "0"    '���۵���
                    oNode.setAttribute "ifhquantity", ""   '�ۼƷ�������
                    oNode.setAttribute "ifhnum", ""        '�ۼƷ�������������
                    oNode.setAttribute "itaxunitprice", "0" ' "882.00"  'ԭ�Һ�˰����
                    oNode.setAttribute "iunitprice", "0"   '"753.85"      'ԭ����˰����
                    oNode.setAttribute "inatunitprice", "0" '"753.85"   '������˰����
                    oNode.setAttribute "inatmoney", "0"    ' "75384.62"    '������˰���
                    oNode.setAttribute "inattax", "0"      ' "12815.38"      '����˰��
                    oNode.setAttribute "imoney", "0"       ' "75384.62"       '���ҽ��
                    oNode.setAttribute "itaxrate", "0"     '��ͷ˰��
                    oNode.setAttribute "ifhmoney", ""      '�ۼ�ԭ�ҷ������
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


                    '����ۼ�
                    Dim fiInvLSCost As Double
                    sqlstr = "Select iInvLSCost From Inventory Where cinvcode='" & Null2Something(GetNodeAtrVal(ele, "cinvcode")) & "'"
                    If fRs.State = adStateOpen Then fRs.Close
                    Set fRs = conn.Execute(sqlstr)
                    '����ۼ�
                    If Not fRs.EOF Then
                        oNode.setAttribute "iinvlscost", Null2Something(fRs!iInvLSCost)
                        fiInvLSCost = val(Null2Something(fRs!iInvLSCost))
                    Else
                        oNode.setAttribute "iinvlscost", ""
                    End If
                    '�ͻ�����ۼ�
                    sqlstr = "Select fCusminPrice From SA_CusPriceJustdetail Where ccuscode='" & Null2Something(r!cCusCode) & "' And cinvcode='" & Null2Something(GetNodeAtrVal(ele, "cinvcode")) & "'"
                    If fRs.State = adStateOpen Then fRs.Close
                    Set fRs = conn.Execute(sqlstr)
                    '�ͻ�����ۼ�
                    If Not fRs.EOF Then
                        If IsNull(fRs!fCusMinPrice) Then
                            oNode.setAttribute "fcusminprice", str(fiInvLSCost)
                        Else
                            oNode.setAttribute "fcusminprice", Null2Something(fRs!fCusMinPrice)    '�ͻ�����ۼ�
                        End If
                    Else
                        oNode.setAttribute "fcusminprice", ""
                    End If


                    '�쿴�۸�ѡ���Ƿ����ü۸�����             xin            2008-10-22
                    Dim klValue As Double, kl2Value As Double
                    Set rsKL = conn.Execute("select cvalue from accinformation where cname='bquantitydisrate'")
                    If Not (rsKL Is Nothing) Then
                        If rsKL.Fields(0) Then
                            'ȡ����2
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
                        Case 0, 1                          '0:���¼۸�  1�����³ɱ��۸�
                            oNode.setAttribute "kl", "100" '����
                            oNode.setAttribute "kl2", kl2Value    '���ο���
                            oNode.setAttribute "dkl1", "0" '������1
                            oNode.setAttribute "dkl2", "0" '������2

                        Case 2                             '�۸�����

                            'ȡ����
                            If gU8Version = "870" Then
                                getKL conn, Null2Something(r!cCusCode), Null2Something(GetNodeAtrVal(ele, "cinvcode")), Null2Something(GetNodeAtrVal(ele, "cfree1")), Null2Something(GetNodeAtrVal(ele, "cfree2")), Format(dateStr, "YYYY-MM-DD"), GetNodeAtrVal(ele, "iquantity"), clsSysSa, klValue
                            ElseIf gU8Version = "871" Then
                                getKL871 conn, Null2Something(r!cCusCode), Null2Something(GetNodeAtrVal(ele, "cinvcode")), Null2Something(GetNodeAtrVal(ele, "cfree1")), Null2Something(GetNodeAtrVal(ele, "cfree2")), Format(dateStr, "YYYY-MM-DD"), GetNodeAtrVal(ele, "iquantity"), exchNameStr, clsSysSa, klValue
                                '872����   xin 2008-10-22
                            ElseIf gU8Version = "872" Then
                                getKL872 conn, Null2Something(r!cCusCode), Null2Something(GetNodeAtrVal(ele, "cinvcode")), Null2Something(GetNodeAtrVal(ele, "cfree1")), Null2Something(GetNodeAtrVal(ele, "cfree2")), GetNodeAtrVal(ele, "cfree3") & "", GetNodeAtrVal(ele, "cfree4") & "", GetNodeAtrVal(ele, "cfree5") & "", GetNodeAtrVal(ele, "cfree6") & "", GetNodeAtrVal(ele, "cfree7") & "", GetNodeAtrVal(ele, "cfree8") & "", GetNodeAtrVal(ele, "cfree9") & "", GetNodeAtrVal(ele, "cfree10") & "", Format(dateStr, "YYYY-MM-DD"), GetNodeAtrVal(ele, "iquantity"), exchNameStr, clsSysSa, klValue
                            End If


                            oNode.setAttribute "kl", klValue    '����
                            oNode.setAttribute "kl2", kl2Value    '���ο���
                            oNode.setAttribute "dkl1", str(100 - klValue)    '������1
                            oNode.setAttribute "dkl2", str(100 - kl2Value)    '������2
                    End Select


                    oNode.setAttribute "iinvid", ""
                    oNode.setAttribute "cdefine22", GetNodeAtrVal(ele, "cdefine22")    '�����Զ�����
                    oNode.setAttribute "cdefine23", GetNodeAtrVal(ele, "cdefine23")    '�����Զ�����
                    oNode.setAttribute "cdefine24", GetNodeAtrVal(ele, "cdefine24")    '�����Զ�����
                    oNode.setAttribute "cdefine25", GetNodeAtrVal(ele, "cdefine25")    '�����Զ�����
                    oNode.setAttribute "cdefine26", GetNodeAtrVal(ele, "cdefine26")    '�����Զ�����
                    oNode.setAttribute "cdefine27", GetNodeAtrVal(ele, "cdefine27")    '�����Զ�����
                    oNode.setAttribute "cdefine28", GetNodeAtrVal(ele, "cdefine28")    '�����Զ�����
                    oNode.setAttribute "cdefine29", GetNodeAtrVal(ele, "cdefine29")    '�����Զ�����
                    oNode.setAttribute "cdefine30", GetNodeAtrVal(ele, "cdefine30")    '�����Զ�����
                    oNode.setAttribute "cdefine31", GetNodeAtrVal(ele, "cdefine31")    '�����Զ�����
                    oNode.setAttribute "cdefine32", GetNodeAtrVal(ele, "cdefine32")    '�����Զ�����
                    oNode.setAttribute "cdefine33", GetNodeAtrVal(ele, "cdefine33")    '�����Զ�����
                    oNode.setAttribute "cdefine34", GetNodeAtrVal(ele, "cdefine34")    '�����Զ�����
                    oNode.setAttribute "cdefine35", GetNodeAtrVal(ele, "cdefine35")    '�����Զ�����
                    oNode.setAttribute "cdefine36", GetNodeAtrVal(ele, "cdefine36")    '�����Զ�����
                    oNode.setAttribute "cdefine37", GetNodeAtrVal(ele, "cdefine37")    '�����Զ�����

                    '2008-01-23 ������,����
                    sqlstr = "Select iChangRate From ComputationUnit Where cComunitCode In (Select cSAComUnitCode From Inventory Where cinvcode='" & GetNodeAtrVal(ele, "cinvcode") & "')"
                    If fRs.State = adStateOpen Then fRs.Close
                    Set fRs = conn.Execute(sqlstr)
                    If Not fRs.EOF Then
                        '�޹̶�������                 2008-11-06            xin
                        If Not IsNull(fRs.Fields(0)) Then
                            oNode.setAttribute "iinvexchrate", fRs.Fields("iChangRate")    '������
                            If fRs.Fields("iChangRate") <> 0 Then oNode.setAttribute "inum", GetNodeAtrVal(ele, "iquantity") / fRs.Fields("iChangRate")    '����
                        End If
                    End If
                    '2008-01-15 ˰��
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
                            '��ͷ��������
                            oNode.setAttribute "itaxrate", Null2Something(SAHead.selectSingleNode("//z:row").Attributes.getNamedItem("itaxrate").nodeValue)
                        Else
                            '��ͷȡĬ��ֵ
                            Set taxRateNode = SAHead.selectSingleNode("//z:row")
                            taxRateNode.setAttribute "itaxrate", "0"
                        End If
                    End If
                    fRs.Close


                    oNode.setAttribute "dpredate", dateStr 'Ԥ��������
                    oNode.setAttribute "dpremodate", dateStr    'Ԥ�깤����
                    '��PartID ���Ҵ�����뼰������-------------------------
                    sqlstr = "select cComUnitCode,cAssComUnitCode,Free1,Free2,Free3,Free4,Free5,Free6, " & _
                            " Free7,Free8,Free9,Free10 ,cSTComUnitCode,iGroupType,cGroupCode " & _
                            " from v_bas_part  " & _
                            " where cinvcode ='" & GetNodeAtrVal(ele, "cinvcode") & "' And Free1='" & GetNodeAtrVal(ele, "cfree1") & "' And Free2='" & GetNodeAtrVal(ele, "cfree2") & "' And Free3='" & GetNodeAtrVal(ele, "cfree3") & "'And " & _
                            " Free4='" & GetNodeAtrVal(ele, "cfree4") & "' And Free5='" & GetNodeAtrVal(ele, "cfree5") & "' And Free6='" & GetNodeAtrVal(ele, "cfree6") & "' And Free7='" & GetNodeAtrVal(ele, "cfree7") & "' And Free8='" & GetNodeAtrVal(ele, "cfree8") & "' And Free9='" & GetNodeAtrVal(ele, "cfree9") & "' And Free10='" & GetNodeAtrVal(ele, "cfree10") & "'"
                    Set fRs = conn.Execute(sqlstr)

                    If fRs.EOF = False Then
                        oNode.setAttribute "cbarcode", ""  '��Ӧ���������
                        oNode.setAttribute "cinvcode", Null2Something(GetNodeAtrVal(ele, "cinvcode"))    '�������
                        oNode.setAttribute "iquantity", Null2Something(GetNodeAtrVal(ele, "iquantity"))    '����
                        oNode.setAttribute "igrouptype", Null2Something(fRs!iGroupType)    '������λ�����
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
                        oNode.setAttribute "ccomunitcode", Null2Something(fRs!cComUnitCode)    '��������λ����
                        oNode.setAttribute "cgroupcode", Null2Something(fRs!cGroupCode)    '"05"           '������λ�����

                        fRs.Close

                        sqlstr = "Select cSAComUnitCode From Inventory Where cinvcode='" & Trim(GetNodeAtrVal(ele, "cinvcode")) & "'"
                        Set fRs = conn.Execute(sqlstr)
                        oNode.setAttribute "cunitid", Null2Something(fRs!cSAComUnitCode)    '������λ

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
                '�Ա����д������ȡ��
                '***************ȡ�۽ӿ���Ҫ�޸�         -2008.10.13 -���
                strErr = pco.VoucherGetPrice(conn, SAHead, SABody)
                '�������۵���
                GetPrice2 SABody

                Set domHead = SAHead.cloneNode(True)
                Set domBody = SABody.cloneNode(True)

ToSave:
                strErr = pco.Save(SAHead, SABody, 0, sVoucherID, retDom)
                If strErr <> "" Then
                    '                    MsgBox "���ʱ�������۵�����!" & strErr, vbInformation, pustrMsgTitle
                    '                    WriteSABill = False
                    Set frmCheckCredit.myinfo = clsSysSa

                    If SAHead.selectNodes("//���ü�鲻ͨ��").Length > 0 Then

                        If frmCheckCredit.CheckShow(SAHead, errMsg) = False Then
                            'MsgBox errMsg, vbExclamation, GetString("U8.SA.xsglsql.01.frmbillvouch.00402")  'zh-CN�����ü��
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
                        If SAHead.selectNodes("//����ۼ�").Length > 0 Then
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
                            'MsgBox strError                            ''����ۼ۷��ش���
                        Else
                            If SAHead.selectNodes("//��������鲻��").Length > 0 Then
                                If frmCheckCredit.CheckShow(SAHead, errMsg, 1) = False Then
                                    'MsgBox errMsg, vbExclamation, GetString("U8.SA.xsglsql.01.frmbillvouch.00403")  'zh-CN�����������
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
                                '                                        'GoTo ToSave ' ���Ա���
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

                    'MsgBox "���ʱ�������۵�����!" & strErr, vbInformation, pustrMsgTitle
missPass:

                    WriteSABill = False
                Else

                    '���۵����ɳɹ��󣬻�д��Ӧ����Ϣ-----------------------------------------------------------------------------------------------------------------------------

                    'ȡ������id
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
                    '�Զ����
                    RsTemp.Close

                    If True Then                           '�Զ����
                        '�Զ����,�ж��Ƿ��ܹ���������
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
                '                 voucherErrMsg = voucherErrMsg & IIf(voucherErrMsg = "", "����ʧ�ܣ�ԭ��" & strErr, vbCrLf & "����ʧ�ܣ�ԭ��" & strErr)
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
    '    FrmMsgBox.Text1 = "�ɹ����� " & voucherSuccSize & " �����۶���" & IIf(voucherSuccSize > 0, "������ " & vouchID, "") & vbCrLf & voucherErrMsg
    FrmMsgBox.Show 1
    'MsgBox "���� " & vouchID & " �����۶����ɹ�!", vbInformation, pustrMsgTitle
    'End If
    Set r = Nothing
    WriteSABill = True
    Exit Function
ErrHandler:
    MsgBox Err.Description, vbInformation, GetString("U8.DZ.JA.Res030")
    WriteSABill = False

End Function

'д�ɹ��൥��
Public Function WritePUBill(ByRef oDomHead As DOMDocument, _
                            ByRef oDomBody As DOMDocument, _
                            conn As Object, _
                            login As clsLogin, _
                            VoucherType As String, _
                            sBillType As String) As Boolean

    On Error GoTo ErrHandler

    Dim ele           As IXMLDOMElement, eleList As IXMLDOMNodeList

    Dim sBusType      As String                                 'ҵ������:��ͨ�ɹ�,ֱ�˲ɹ�,���д���

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
   
    strSql = "select distinct cdeptcode,cDepName from V_HY_LSDG_InputpuAppdata  where  isnull(istats,0) ='δ����' and  id in (" & idtmp & ")"
    Set r = New ADODB.Recordset
    r.Open strSql, conn, 1, 1

    If r.EOF Then
        
        MsgBox "���ݲ����ڻ���������", vbInformation, "��ʾ"
        WritePUBill = False

        Exit Function

    Else

        While Not r.EOF

            sBusType = "��ͨ�ɹ�"
            '2008-01-31 ��ʼ���ɹ������ӿ�
            Set mVouchCO = CreateObject("VoucherCO_PU.clsVoucherCO_PU")
            '   Sub Init(enmVoucherType As VoucherType, [Login As clsLogin], [conn As Connection], [clsInfor As ClsS_Infor], [bPositive As Boolean = True], [sBillType As String], [sBusType As String], [emnUseMode As UseMode])
            mVouchCO.Init VoucherType, login, conn, , True, sBillType, sBusType    'sbilltype=88Ϊ���ݱ�ʾ ����ɹ�����
            mVouchCO.bOutTrans = True

            '   ��֯��odomhead��

            Set CGHead = New DOMDocument
            Set CGBody = New DOMDocument

            'д��ͷ----------------------------------------------------------------------------------
            strSql = "select *,'' as editprop from " & ViewHead & " where 1>2"
            rs.Open strSql, conn, adOpenDynamic, adLockOptimistic
            rs.Save CGHead, adPersistXML                       '�õ���ⵥ��ͷDOM�ṹ����
            rs.Close
            Set oNodes = CGHead.selectSingleNode("//rs:data")
            Set oNode = CGHead.createElement("z:row")
 
            oNode.setAttribute "cdepcode", Null2Something(r!cdeptcode)    '  cDepCode ���ű���  varchar 12  True
        
            oNode.setAttribute "cbustype", sBusType          '  cBusType ҵ������  varchar 8  True
       
            oNode.setAttribute "ivtid", GetVoucherID(conn, sBillType)

            oNode.setAttribute "cdefine1", ""    '��ͷ�Զ�����
            oNode.setAttribute "cdefine2", ""   '��ͷ�Զ�����
            oNode.setAttribute "cdefine3", ""    '��ͷ�Զ�����
            oNode.setAttribute "cdefine4", ""   '��ͷ�Զ�����
            oNode.setAttribute "cdefine5", ""    '��ͷ�Զ�����
            oNode.setAttribute "cdefine6", ""   '��ͷ�Զ�����
            oNode.setAttribute "cdefine7", ""    '��ͷ�Զ�����
            oNode.setAttribute "cdefine8", ""    '��ͷ�Զ�����
            oNode.setAttribute "cdefine9", ""    '��ͷ�Զ�����
            oNode.setAttribute "cdefine10", ""    '��ͷ�Զ�����
            oNode.setAttribute "cdefine11", ""   '��ͷ�Զ�����
            oNode.setAttribute "cdefine12", ""    '��ͷ�Զ�����
            oNode.setAttribute "cdefine13", ""    '��ͷ�Զ�����
            oNode.setAttribute "cdefine14", ""    '��ͷ�Զ�����
            oNode.setAttribute "cdefine15", ""    '��ͷ�Զ�����
            oNode.setAttribute "cdefine16", ""    '��ͷ�Զ�����

            oNodes.appendChild oNode

            If mVouchCO.GetVoucherNO(CGHead, sBillType, errMsg, POID) = False Then
                WritePUBill = False

                Exit Function

            End If

            '����ͷ���嵥�ݱ��
            Set ele = CGHead.selectSingleNode("//z:row")
            ele.setAttribute "ccode", POID
            ele.setAttribute "ufts", ""
            ele.setAttribute "ddate", login.CurDate
            ele.setAttribute "cbustype", sBusType

            '����R!id�������----------------------------------------

            strSql = "select *,'' as editprop from " & ViewBody & " where 1>2"
            rs.Open strSql, conn, adOpenDynamic, adLockOptimistic
            rs.Save CGBody, adPersistXML
            rs.Close

            strSql = "select * from V_HY_LSDG_InputpuAppdata  where  isnull(istats,0) ='δ����' and  id in (" & idtmp & ") and cdeptcode='" & r.Fields("cdeptcode") & "'"
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
                oNode.setAttribute "cexch_name", "�����"
                oNode.setAttribute "iexchrate", 1
                oNode.setAttribute "ivouchrowno", rows
                oNode.setAttribute "fquantity", CDbl(Format(Null2Something(rst.Fields("iqty")), m_sQuantityFmt))

                oNode.setAttribute "cdefine22", ""    '�����Զ�����1
                oNode.setAttribute "cdefine23", ""    '�����Զ�����2
                oNode.setAttribute "cdefine24", ""     '�����Զ�����3
                oNode.setAttribute "cdefine25", ""    '�����Զ�����4
                oNode.setAttribute "cdefine26", ""   '�����Զ�����5
                oNode.setAttribute "cdefine27", ""     '�����Զ�����6
                oNode.setAttribute "cdefine28", ""     '�����Զ�����7
                oNode.setAttribute "cdefine29", ""     '�����Զ�����8
                oNode.setAttribute "cdefine30", ""    '�����Զ�����9
                oNode.setAttribute "cdefine31", ""     '�����Զ�����10
                oNode.setAttribute "cdefine32", ""    '�����Զ�����11
                oNode.setAttribute "cdefine33", Null2Something(rst.Fields("id"))     '�����Զ�����12
                oNode.setAttribute "cdefine34", ""     '�����Զ�����13
                oNode.setAttribute "cdefine35", ""    '�����Զ�����14
                oNode.setAttribute "cdefine36", ""    '�����Զ�����15
                oNode.setAttribute "cdefine37", ""    '�����Զ�����16

                oNode.setAttribute "editprop", "A"
                oNodes.appendChild oNode
                rst.MoveNext
                rows = rows + 1

            Wend
            rst.Close

            '        '�Ǵ��ܲɹ���ҵ�����ͽ�������޼ۿ���
            '        If sBusType <> "���ܲɹ�" Then
            '            If Not bGetMPService(sBillType, CGHead, CGBody, conn, login) Then
            '                strError = GetString("U8.DZ.JA.Res1000")
            '                WritePUBill = False
            '                Exit Function
            '            End If
            '        End If
            '2008-01-31 ���òɹ��ӿ�����
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
    FrmMsgBox.Text1 = "�ɹ����� " & voucherSuccSize & " ���빺��" & IIf(voucherSuccSize > 0, "������ " & vouchID, "") & vbCrLf & voucherErrMsg
    FrmMsgBox.Show 1
    'MsgBox "���� " & vouchID & " �����۶����ɹ�!", vbInformation, pustrMsgTitle
    'End If
    Set r = Nothing
    WritePUBill = True

    Exit Function

ErrHandler:
    MsgBox Err.Description, vbInformation, GetString("U8.DZ.JA.Res030")
    WritePUBill = False

End Function

'д����൥��
Public Function WriteSCBill(ByRef oDomHead As DOMDocument, ByRef oDomBody As DOMDocument, conn As Object, login As clsLogin, VoucherType As String, sBillType As String) As Boolean
    On Error GoTo ErrHandler

    Dim pco As Object
    Dim errMsg As String
    Dim domMsg As DOMDocument
    Dim Postion As New DOMDocument                         '��λ��Ϣ
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
        
        MsgBox "���ݲ����ڻ���������", vbInformation, "��ʾ"
        WriteSCBill = False

        Exit Function

    Else

        While Not r.EOF

        '   ��֯��odomhead
        Set SCHead = New DOMDocument
        Set SCBody = New DOMDocument
        ViewHead = GetViewHead(conn, sBillType)
        ViewBody = GetViewBody(conn, sBillType)

        'д��ͷ----------------------------------------------------------------------------------
        strSql = "select *,'' as editprop from " & ViewHead & " where 1>2"
          Set rs = New ADODB.Recordset
        rs.Open strSql, conn, adOpenDynamic, adLockOptimistic
        rs.Save SCHead, adPersistXML                       '�õ���ⵥ��ͷDOM�ṹ����
        rs.Close
        Set oNodes = SCHead.selectSingleNode("//rs:data")
        Set oNode = SCHead.createElement("z:row")

        oNode.setAttribute "cwhcode", getAccinformation("QR", "cWhCode", conn) '"01" 'IIf(r!cwhcode & "" = "", "001", r!cwhcode) '"001"    '�ֿ��� (�����ֶ�)
        oNode.setAttribute "ddate", login.CurDate                    '��������
        oNode.setAttribute "crdcode", getAccinformation("QR", "cRdCode", conn) '"102"                   '�շ����
        oNode.setAttribute "btransflag", "0"               '�Ƿ񴫵�
        oNode.setAttribute "cmaker", login.cUserName       '�Ƶ���
        oNode.setAttribute "cbustype", "��Ʒ���"              'ҵ������
'        oNode.setAttribute "inetlock", "0"                 '������
        oNode.setAttribute "brdflag", "1"                  '�շ���ʶ
        oNode.setAttribute "cvouchtype", "10"              '��������(������ⵥ)
        oNode.setAttribute "csource", "��������"    '��ԭ����
        oNode.setAttribute "bpufirst", "0"                 '�ɹ��ڳ���־
        oNode.setAttribute "biafirst", "0"                 '����ڳ���־
        oNode.setAttribute "bisstqc", "0"                  '����ڳ���־
        oNode.setAttribute "bomfirst", "0"                 'ί�������־

        oNode.setAttribute "cmemo", ""    '��ע
        oNode.setAttribute "iexchrate", 1   '����
        oNode.setAttribute "cexch_name", "�����"    '����
        oNode.setAttribute "ccode", "0000000001"           '�շ����ݺ�
        oNode.setAttribute "iproorderid", r!moid '�������������ʶ
        oNode.setAttribute "cmpocode", r!cmocode '�����������
        oNode.setAttribute "cdepcode", GetMDeptCode(r!moid) 'Null2Something(r!cdeptcode)    '���ű���
'        oNode.setAttribute "cCusCode", Null2Something(r!cdeptcode)
        

        oNode.setAttribute "cdefine1", ""    '��ͷ�Զ�����
            oNode.setAttribute "cdefine2", ""   '��ͷ�Զ�����
            oNode.setAttribute "cdefine3", ""    '��ͷ�Զ�����
            oNode.setAttribute "cdefine4", ""   '��ͷ�Զ�����
            oNode.setAttribute "cdefine5", ""    '��ͷ�Զ�����
            oNode.setAttribute "cdefine6", ""   '��ͷ�Զ�����
            oNode.setAttribute "cdefine7", ""    '��ͷ�Զ�����
            oNode.setAttribute "cdefine8", ""    '��ͷ�Զ�����
            oNode.setAttribute "cdefine9", ""    '��ͷ�Զ�����
            oNode.setAttribute "cdefine10", ""    '��ͷ�Զ�����
            oNode.setAttribute "cdefine11", ""   '��ͷ�Զ�����
            oNode.setAttribute "cdefine12", ""    '��ͷ�Զ�����
            oNode.setAttribute "cdefine13", ""    '��ͷ�Զ�����
            oNode.setAttribute "cdefine14", ""    '��ͷ�Զ�����
            oNode.setAttribute "cdefine15", ""    '��ͷ�Զ�����
            oNode.setAttribute "cdefine16", ""    '��ͷ�Զ�����

        oNode.setAttribute "vt_id", GetVoucherID(conn, sBillType)    '������ʾģ���
        oNodes.appendChild oNode
        
         Dim oDomFormat As DOMDocument
     Dim sError As String
    Dim strVoucherNo As String
        
         If GetVoucherNO(conn, SCHead, sBillType, sError, strVoucherNo, , , , False) = True Then
             Set ele = SCHead.selectSingleNode("//z:row")
            ele.setAttribute "ccode", strVoucherNo
         End If
         
        
        '����R!id�������----------------------------------------

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
                oNode.setAttribute "cinvcode", Null2Something(rst.Fields("cinvcode"))   '�������
'
'                oNode.setAttribute "cinvm_unit", GetNodeAtrVal(ele, "ccomunitcode")    '            ������
'
'                oNode.setAttribute "cassunit", GetNodeAtrVal(ele, "cunitid")    '            ������

                oNode.setAttribute "iquantity", CDbl(Format(Null2Something(rst.Fields("iqty")), m_sQuantityFmt))   '����

'                If GetNodeAtrVal(ele, "inum") <> "" Then _
'                        oNode.setAttribute "inum", Format(CDbl(GetNodeAtrVal(ele, "inum")), m_sNumFmt)    '    ����
'                If GetNodeAtrVal(ele, "iinvexchrate") <> "" Then _
'                        oNode.setAttribute "iinvexchrate", Format(CDbl(GetNodeAtrVal(ele, "iinvexchrate")), m_iExchRateFmt)    '������
'                oNode.setAttribute "cbatch", GetNodeAtrVal(ele, "cBatch")    '����

                oNode.setAttribute "bcosting", "1"         '�Ƿ����
'                If Val(Format(Null2Something(rst.Fields("price")), m_sPriceFmt)) <> 0 Then _
'                        oNode.setAttribute "iunitcost", Val(Format(Null2Something(rst.Fields("price")), m_sPriceFmt))    '���ҵ���
               'If GetNodeAtrVal(ele, "inatmoney") <> "" Then
'                        oNode.setAttribute "iprice", Val(Format(Null2Something(rst.Fields("price")), m_sPriceFmt)) * Val(Format(Null2Something(rst.Fields("iqty")), m_sQuantityFmt))    '���ҽ��
               ' If GetNodeAtrVal(ele, "itaxrate") <> "" _
                        Then
                        oNode.setAttribute "itaxrate", 17    '˰��
'
'                oNode.setAttribute "isotype", GetNodeAtrVal(ele, "sotype")    ' ������ٷ�ʽ
'                oNode.setAttribute "csocode", GetNodeAtrVal(ele, "socode")    ' ������ٺ�
'                oNode.setAttribute "cdemandmemo", GetNodeAtrVal(ele, "cdemandmemo")    '����������˵��

'
'                oNode.setAttribute "iexpiratdatecalcu", GetNodeAtrVal(ele, "iexpiratdatecalcu")    '��Ч�����㷽ʽ
'                oNode.setAttribute "cexpirationdate", GetNodeAtrVal(ele, "cexpirationdate")    '��Ч����
'                oNode.setAttribute "dexpirationdate", GetNodeAtrVal(ele, "dexpirationdate")    '��Ч�ڼ�����
'                oNode.setAttribute "dmadedate", GetNodeAtrVal(ele, "dmadedate")    '��������
'                oNode.setAttribute "imassdate", GetNodeAtrVal(ele, "imassdate")    '������
'                oNode.setAttribute "cmassunit", GetNodeAtrVal(ele, "cmassunit")    '�����ڵ�λ
                    
                oNode.setAttribute "imoseq", rst!imoseq    '���������к�
                oNode.setAttribute "impoids", rst!modid    '���������ӱ��ʶ
                oNode.setAttribute "cmocode", r!cmocode    '����������
                 oNode.setAttribute "cdefine22", ""    '�����Զ�����1
                oNode.setAttribute "cdefine23", ""    '�����Զ�����2
                oNode.setAttribute "cdefine24", ""     '�����Զ�����3
                oNode.setAttribute "cdefine25", ""    '�����Զ�����4
                oNode.setAttribute "cdefine26", ""   '�����Զ�����5
                oNode.setAttribute "cdefine27", ""     '�����Զ�����6
                oNode.setAttribute "cdefine28", ""     '�����Զ�����7
                oNode.setAttribute "cdefine29", ""     '�����Զ�����8
                oNode.setAttribute "cdefine30", ""    '�����Զ�����9
                oNode.setAttribute "cdefine31", ""     '�����Զ�����10
                oNode.setAttribute "cdefine32", ""    '�����Զ�����11
'                oNode.setAttribute "cdefine33", Null2Something(rst.Fields("id"))     '�����Զ�����12
                oNode.setAttribute "cdefine34", ""     '�����Զ�����13
                oNode.setAttribute "cdefine35", ""    '�����Զ�����14
                oNode.setAttribute "cdefine36", ""    '�����Զ�����15
                oNode.setAttribute "cdefine37", ""    '�����Զ�����16

                oNode.setAttribute "ufts", ""
                oNode.setAttribute "editprop", "A"
                oNodes.appendChild oNode
                rst.MoveNext
                rows = rows + 1
                  
            Wend
            rst.Close

        '2008-01-31 ���ÿ��ӿ�����
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
      FrmMsgBox.Text1 = "�ɹ����� " & voucherSuccSize & " �ſ�浥��" & IIf(voucherSuccSize > 0, "������ " & vouchID, "") & vbCrLf & voucherErrMsg
    FrmMsgBox.Show 1
    'MsgBox "���� " & vouchID & " �����۶����ɹ�!", vbInformation, pustrMsgTitle
    'End If
    Set r = Nothing
    WriteSCBill = True

    Exit Function

ErrHandler:
    If Err.Description <> "" Then MsgBox Err.Description, vbInformation, GetString("U8.DZ.JA.Res030")
    WriteSCBill = False

End Function



'дӦ������
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

        '   ��֯��odomhead
        Set APhead = New DOMDocument
        Set APBody = New DOMDocument
        ViewHead = GetViewHead(conn, sBillType)
        ViewBody = GetViewBody(conn, sBillType)

        'д��ͷ----------------------------------------------------------------------------------
        strSql = "select *,'' as editprop from " & ViewHead & " where 1>2"
        rs.Open strSql, conn, adOpenDynamic, adLockOptimistic
        rs.Save APhead, adPersistXML                       '�õ���ⵥ��ͷDOM�ṹ����
        rs.Close
        Set oNodes = APhead.selectSingleNode("//rs:data")
        Set oNode = APhead.createElement("z:row")

        oNode.setAttribute "cVouchType", "P0"              '�������� -Ap_VouchType��   0Ϊ������
        '            oNode.setAttribute "cVouchID", "0000000001"                  '��Ӧ���ݺ�
        oNode.setAttribute "cVouchID1", Null2Something(r!cCode)    '��Ӧ��������
        oNode.setAttribute "cCoVouchType", gstrCardNumber  '"HY99"                  '��Ӧ���ݺ�
        oNode.setAttribute "dVouchDate", Date              '��������
        oNode.setAttribute "cDeptCode", Null2Something(r!cDepcode)    '���ű���
        oNode.setAttribute "cPerson", Null2Something(r!cpersoncode)    'ҵ��Ա����
        oNode.setAttribute "cCode", ""                     '��Ŀ����
        oNode.setAttribute "cexch_name", Null2Something(r!cexch_name)    '����
        oNode.setAttribute "iExchRate", Null2Something(r!iExchRate)    '����
        oNode.setAttribute "cDigest", GetString("U8.DZ.JA.Res990") & Time    'ժҪ
        oNode.setAttribute "cPayCode", ""                  '��������
        oNode.setAttribute "cOperator", login.cUserName    '¼����
        oNode.setAttribute "bStartFlag", "0"               '�ڳ���־


        oNode.setAttribute "cFlag", VoucherType            '�շ���ʶ
        oNode.setAttribute "bd_c", IIf(VoucherType = "AP", "0", "1")    '�������
        oNode.setAttribute "cDwCode", IIf(VoucherType = "AP", Null2Something(r!cvencode), Null2Something(r!cCusCode))    '��λ

        sumQ = 0: sumM = 0: sumM_f = 0
        Set eleList = oDomBody.selectNodes("//z:row[@" + HeadPKFld + "='" & r!id & "']")
        For Each ele In eleList
            If GetNodeAtrVal(ele, "iquantity") <> "" Then sumQ = sumQ + CDbl(GetNodeAtrVal(ele, "iquantity"))
            If GetNodeAtrVal(ele, "inatmoney") <> "" Then sumM = sumM + CDbl(GetNodeAtrVal(ele, "inatsum"))
            If GetNodeAtrVal(ele, "isum") <> "" Then sumM_f = sumM_f + CDbl(GetNodeAtrVal(ele, "isum"))
        Next
        oNode.setAttribute "iAmount_s", Format(CDbl(sumQ), m_sQuantityFmt)    '����
        oNode.setAttribute "iAmount", Format(CDbl(sumM), m_sPriceFmt)    '���ҽ��
        oNode.setAttribute "iAmount_f", Format(CDbl(sumM_f), m_sPriceFmt)    'ԭ�ҽ��

        oNode.setAttribute "cDefine1", Null2Something(r!cDefine1)    '��ͷ�Զ�����
        oNode.setAttribute "cDefine2", Null2Something(r!cDefine2)    '��ͷ�Զ�����
        oNode.setAttribute "cDefine3", Null2Something(r!cDefine3)    '��ͷ�Զ�����
        oNode.setAttribute "cDefine4", Null2Something(r!cDefine4)    '��ͷ�Զ�����
        oNode.setAttribute "cDefine5", Null2Something(r!cDefine5)    '��ͷ�Զ�����
        oNode.setAttribute "cDefine6", Null2Something(r!cDefine6)    '��ͷ�Զ�����
        oNode.setAttribute "cDefine7", Null2Something(r!cdefine7)    '��ͷ�Զ�����
        oNode.setAttribute "cDefine8", Null2Something(r!cDefine8)    '��ͷ�Զ�����
        oNode.setAttribute "cDefine9", Null2Something(r!cDefine9)    '��ͷ�Զ�����
        oNode.setAttribute "cDefine10", Null2Something(r!cDefine10)    '��ͷ�Զ�����
        oNode.setAttribute "cDefine11", Null2Something(r!cDefine11)    '��ͷ�Զ�����
        oNode.setAttribute "cDefine12", Null2Something(r!cDefine12)    '��ͷ�Զ�����
        oNode.setAttribute "cDefine13", Null2Something(r!cDefine13)    '��ͷ�Զ�����
        oNode.setAttribute "cDefine14", Null2Something(r!cDefine14)    '��ͷ�Զ�����
        oNode.setAttribute "cDefine15", Null2Something(r!cDefine15)    '��ͷ�Զ�����
        oNode.setAttribute "cDefine16", Null2Something(r!cdefine16)    '��ͷ�Զ�����

        oNode.setAttribute "vt_id", GetVoucherID(conn, sBillType)    '������ʾģ���
        oNodes.appendChild oNode


        'Function GetVouchID(cType As String, oDom As DOMDocument, xmlErrMsg As String) As String
        '����R!id�������----------------------------------------

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

                oNode.setAttribute "cinvm_unit", GetNodeAtrVal(ele, "citemcode")    '��Ŀ����
                oNode.setAttribute "cItemCode", GetNodeAtrVal(ele, "citem_class")    '   ��Ŀ�������
                oNode.setAttribute "cItemName", GetNodeAtrVal(ele, "citemname")    '��Ŀ����
                oNode.setAttribute "cPerson", Null2Something(r!cpersoncode)    'ҵ��Ա����
                oNode.setAttribute "cDeptCode", Null2Something((r!cDepcode))    '���ű���
                oNode.setAttribute "cDwCode", IIf(VoucherType = "AP", Null2Something(r!cvencode), Null2Something(r!cCusCode))    '��λ
                oNode.setAttribute "iAmt_s", Format(CDbl(GetNodeAtrVal(ele, "iquantity")), m_sQuantityFmt)    '����
                oNode.setAttribute "iTaxRate", Format(CDbl(GetNodeAtrVal(ele, "itaxrate")), m_iRateFmt)    '˰��
                oNode.setAttribute "iTax", Format(CDbl(GetNodeAtrVal(ele, "itax")), m_sPriceFmt)    ' ˰��
                oNode.setAttribute "iNatTax", Format(CDbl(GetNodeAtrVal(ele, "inattax")), m_sPriceFmt)    ' ����˰��
                oNode.setAttribute "cexch_name", Null2Something(r!cexch_name)    '����
                oNode.setAttribute "iExchRate", Format(CDbl(Null2Something(r!iExchRate, 0)), m_iRateFmt)    '����
                oNode.setAttribute "bd_c", IIf(VoucherType = "AP", "1", "0")    '�������
                oNode.setAttribute "iAmount", Format(CDbl(GetNodeAtrVal(ele, "inatsum")), m_sPriceFmt)    '���ҽ��
                oNode.setAttribute "iAmount_f", Format(CDbl(GetNodeAtrVal(ele, "isum")), m_sPriceFmt)    'ԭ�ҽ��

                oNode.setAttribute "cDefine22", GetNodeAtrVal(ele, "cDefine22")    '�����Զ�����
                oNode.setAttribute "cDefine23", GetNodeAtrVal(ele, "cDefine23")    '�����Զ�����
                oNode.setAttribute "cDefine24", GetNodeAtrVal(ele, "cDefine24")    '�����Զ�����
                oNode.setAttribute "cDefine25", GetNodeAtrVal(ele, "cDefine25")    '�����Զ�����
                oNode.setAttribute "cDefine26", GetNodeAtrVal(ele, "cDefine26")    '�����Զ�����
                oNode.setAttribute "cDefine27", GetNodeAtrVal(ele, "cDefine27")    '�����Զ�����
                oNode.setAttribute "cDefine28", GetNodeAtrVal(ele, "cDefine28")    '�����Զ�����
                oNode.setAttribute "cDefine29", GetNodeAtrVal(ele, "cDefine29")    '�����Զ�����
                oNode.setAttribute "cDefine30", GetNodeAtrVal(ele, "cDefine30")    '�����Զ�����
                oNode.setAttribute "cDefine31", GetNodeAtrVal(ele, "cDefine31")    '�����Զ�����
                oNode.setAttribute "cDefine32", GetNodeAtrVal(ele, "cDefine32")    '�����Զ�����
                oNode.setAttribute "cDefine33", GetNodeAtrVal(ele, "cDefine33")    '�����Զ�����
                oNode.setAttribute "cDefine34", GetNodeAtrVal(ele, "cDefine34")    '�����Զ�����
                oNode.setAttribute "cDefine35", GetNodeAtrVal(ele, "cDefine35")    '�����Զ�����
                oNode.setAttribute "cDefine36", GetNodeAtrVal(ele, "cDefine36")    '�����Զ�����
                oNode.setAttribute "cDefine37", GetNodeAtrVal(ele, "cDefine37")    '�����Զ�����

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

            '            voucherErrMsg = voucherErrMsg & IIf(voucherErrMsg = "", "����ʧ�ܣ�ԭ��" & strError, vbCrLf & "����ʧ�ܣ�ԭ��" & strError)
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
    ' FrmMsgBox.Text1 = "�ɹ����� " & voucherSuccSize & " ��Ӧ��Ӧ������" & IIf(voucherSuccSize > 0, "������ " & vouchID, "") & vbCrLf & voucherErrMsg
    FrmMsgBox.Show 1

    Set r = Nothing
    WriteAPBill = True

    Exit Function

ErrHandler:
    If Err.Description <> "" Then MsgBox Err.Description, vbInformation, GetString("U8.DZ.JA.Res030")
    WriteAPBill = False

End Function


'������������
Public Function ProcessDatapro(ByRef Voucher As Object)

    On Error GoTo ErrHandler:

    Dim retvalue As Variant
    Dim referpara As UAPVoucherControl85.ReferParameter
    Dim eleline As IXMLDOMElement
    Dim echeck As Long
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim i As Integer

    '�ṩ���ִ���ģʽ����ͷʹ��recordset��������ʹ��xml��������
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

                Voucher.headerText("cdefine" & i) = Null2Something(rshead("cdefine" & i))    '��ͷ�Զ�����

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


'���У�������ַ����ϵ�ˡ��绰��������ɿͻ����Ƶ����Զ�������
Private Sub setallinforbycus(Voucher As Object, Index As Variant, retvalue As String, _
                             bChanged As UAPVoucherControl85.CheckRet, _
                             referpara As UAPVoucherControl85.ReferParameter)
    On Error GoTo Err_Handler
    If Voucher.headerText("cType") <> "�ͻ�" Or Voucher.headerText("bObjectCode") = "" Then Exit Sub

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

'890���ӱ���һ����ȡ�ִ�������hucl
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


'ˢ�±����ִ����Ϳ�����
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
        '�ִ���
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
        
        '������
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

'ˢ�±�ͷ�ͱ����ִ����Ϳ�����
Public Sub ShowStock(ByRef Voucher As UAPVoucherControl85.ctlVoucher, ByVal sInvCode As String, ByVal nRow As Long)
    On Error Resume Next
    '�ִ���
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
        '������
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
    StrReportName = "�ִ�����ѯ"
    'StrHWReportName = "��λ������ѯ"
    iRow = Voucher.row
    sInvCode = Voucher.bodyText(iRow, "cInvCode")
    
    If sInvCode <> "" Then
    
        If Voucher.bodyText(iRow, "cposition") <> "" Then
            StrReportName = "��λ������ѯ"
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
    
    If StrReportName = "��λ������ѯ" Then
        oFltSrv.FilterList.Item("CS.cPosCode").varValue = Voucher.bodyText(iRow, "cposition")
        oFltSrv.FilterList.Item("CS.cPosCode").varValue2 = Voucher.bodyText(iRow, "cposition")
    End If
    
    If oInventory.IsLP And StrReportName = "�ִ�����ѯ" Then
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
    
    StrReportName = "�ִ�����ѯ"
    
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

'����Զ������������ⵥ
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
        varArgs(1) = GetString("U8.DZ.JA.Res1950") '�������ⵥ
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
    'ֻ���ѳ��⣬���ܹ黹
    If vouStatus = "����" Then
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
                'strMsg = strMsg & "���� " & cCode & " �Ѿ��黹��" & vbCrLf
            Else
                CheckCanBack = True
            End If
        End If
    ElseIf vouStatus = "���" Then
        '�ڳ����ݲ��ó���
        If sCreateType = "�ڳ�����" Then
            CheckCanBack = True
        Else
            CheckCanBack = False
            ReDim varArgs(0)
            varArgs(0) = cCode
            strMsg = strMsg & GetStringPara("U8.ST.V870.00757", varArgs(0)) & vbCrLf '
            'strMsg = strMsg & "���� " & cCode & " δ���⣡" & vbCrLf
        End If
    ElseIf vouStatus = "�½�" Then
        CheckCanBack = False
        ReDim varArgs(0)
        varArgs(0) = cCode
        strMsg = strMsg & GetStringPara("U8.DZ.JA.Res460", varArgs(0)) & vbCrLf
        'strMsg = strMsg & "���� " & ccode & " û����ˣ�" & vbCrLf
    ElseIf vouStatus = "�ر�" Then
        CheckCanBack = False
        ReDim varArgs(0)
        varArgs(0) = cCode
        strMsg = strMsg & GetStringPara("U8.DZ.JA.Res445", varArgs(0)) & vbCrLf
        'strMsg = strMsg & "���� " & ccode & " �ѹرգ�" & vbCrLf
    Else
        CheckCanBack = False
        ReDim varArgs(0)
        varArgs(0) = cCode
        strMsg = strMsg & GetStringPara("U8.DZ.JA.Res440", varArgs(0)) & vbCrLf
        'strMsg = strMsg & "���� " & ccode & " �����ڣ�" & vbCrLf
    End If
End Function


Public Function ExecReturn(ByRef lngVoucherID As Long, ByRef sMsg As String, ByRef IsBackWfcontrolled As Boolean, Optional ByVal sUfts As String = "") As Boolean
    On Error GoTo Err_Handler:
    
    Dim sErrMsg As String
    Dim lngReturnVoucherID As Long
    Dim oBorrowOutBack As Object
    Set oBorrowOutBack = CreateObject("HY_DZ_BorrowOutBack.clsBorrowOutSrv")
    oBorrowOutBack.Init g_oLogin
    
    '���ɹ黹���ɹ�
    'If oBorrowOutBack.MakeVouchFromBorrowOut(lngVoucherID, sErrMsg, lngReturnVoucherID, GetTimeStamp(g_Conn, MainTable, lngVoucherID)) Then
    If oBorrowOutBack.MakeVouchFromBorrowOut(lngVoucherID, sErrMsg, lngReturnVoucherID, sUfts) Then
        ExecReturn = True
        
        ReDim varArgs(2)
        Dim sTmp As String
        varArgs(0) = 1
        varArgs(1) = GetString("U8.ST.V870.00756") '����黹��
        
        If Not GetFieldValue(g_Conn, "HY_DZ_BorrowOutBack", "ccode", "id", CStr(lngReturnVoucherID), sTmp) Then
            sTmp = lngReturnVoucherID
        End If
        varArgs(2) = sTmp
        'sMsg = sMsg & "�����ɽ���黹��" & sTmp & "!" & vbCrLf
        sMsg = sMsg & GetStringPara("U8.DZ.JA.Res985", varArgs(0), varArgs(1), varArgs(2)) & vbCrLf
        
        '����黹�������ǹ��������ƣ�����˹黹����������ⵥ
        If Not IsBackWfcontrolled And Not IsBlank(lngReturnVoucherID) Then
            If oBorrowOutBack.Verify(lngReturnVoucherID, sErrMsg) Then
            
                '��˳ɹ�
                If sErrMsg <> "" Then sMsg = sMsg & sErrMsg & vbCrLf
                
                '����������ⵥ
                sErrMsg = oBorrowOutBack.PushOtherIn(lngReturnVoucherID)
                If sErrMsg <> "" Then sMsg = sMsg & sErrMsg & vbCrLf
                
            Else
                '�黹�����ʧ��
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

'��ȡʱ���
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
