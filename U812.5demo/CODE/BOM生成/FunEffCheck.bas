Attribute VB_Name = "FunEffCheck"



'**************************************��Ч�Լ��***********************************************8


'�ֶ���Ч�Լ��
Public Function ExecFunSaveCheck(Voucher As Object) As Boolean
    Dim strMsgHead As String
    Dim strMsgBody As String

    On Error GoTo Err_Handler

    '    '�жϵ��ݱ���Ƿ��Ѵ���
    '    If Voucher.VoucherStatus = VSeAddMode And IsExisted(Voucher.headerText("cCode")) Then
    '        MsgBox GetString("U8.DZ.JA.Res1080"), vbInformation, GetString("U8.DZ.JA.Res030")
    '        ExecFunSaveCheck = False
    '        Voucher.SetFocus
    '        Exit Function
    '    End If

    '��鵥�ݱ�ͷ�����Ƿ�������
    If Voucher.headVaildIsNull2(strMsgHead) = False Then
        MsgBox strMsgHead, vbCritical, GetString("U8.DZ.JA.Res030")
        ExecFunSaveCheck = False
        Voucher.SetFocus
        Exit Function
   
    End If


    '�����˳��༭״̬
    Voucher.ProtectUnload2
'    If Voucher.BodyRows <= 0 Then
'        MsgBox GetString("U8.DZ.JA.Res1090"), vbCritical, GetString("U8.DZ.JA.Res030")
'        ExecFunSaveCheck = False
'        Exit Function
'    End If
'
'    If (Voucher.BodyRowIsEmpty(1) = True) Then
'        If Voucher.bodyVaildIsNull = False Then
'            MsgBox GetString("U8.DZ.JA.Res1100"), vbCritical, GetString("U8.DZ.JA.Res030")
'            ExecFunSaveCheck = False
'            Voucher.SetFocus
'            Exit Function
'        End If
'    End If
'
'    If Not Voucher.bodyVaildIsNull2(strMsgBody) Then
'        MsgBox strMsgBody, vbOKOnly + vbInformation, GetString("U8.DZ.JA.Res030")
'        ExecFunSaveCheck = False
'        Voucher.SetFocus
'        Exit Function
'    End If

    '���������ݵ���Ч��,��Ҫ���:
    '�������Ƿ��
    '����������Ƿ�Ϸ�
    '���ν���
    '��ⵥ��У��
    '��Ŀ����У��
    '��Ч��У��

    If ExecFunEffectiveCheck(Voucher) = False Then
        ExecFunSaveCheck = False
        Exit Function
    End If

    Voucher.ProtectUnload

    ExecFunSaveCheck = True
    Exit Function

Err_Handler:
    MsgBox Err.Description, vbInformation, GetString("U8.DZ.JA.Res030")

End Function

Public Function IsExisted(cCode As String) As Boolean
    On Error GoTo errhanlder:
    Dim rs As ADODB.Recordset
    Set rs = g_Conn.Execute("select 1 from " & MainTable & " where ccode='" & cCode & "'")
    If rs.EOF And rs.BOF Then
        IsExisted = False
    Else
        IsExisted = True
    End If

    Exit Function
errhanlder:
    MsgBox GetString("U8.DZ.JA.Res1110") & Err.Description, vbInformation, GetString("U8.DZ.JA.Res030")
End Function

'���������ݵ���Ч��,��Ҫ���:
'�������Ƿ��
'����������Ƿ�Ϸ�
'���ν���
'��ⵥ��У��
'��Ŀ����У��
'��Ч��У��
Private Function ExecFunEffectiveCheck(Voucher As Object) As Boolean

    On Error GoTo Err_Handler

    Dim sql As String

    Dim rs  As New ADODB.Recordset

    Dim rst As New ADODB.Recordset

    '��ͷУ��:���ݱ��,�Ƶ���,��������Ϊ������,
    '�ֿ�,���ŵ�������Ŀ����ʵ��ҵ�񵥾ݶ���
    If Voucher.headerText(strcCode) = "" Then
        MsgBox GetString("U8.DZ.JA.Res1120"), vbInformation, GetString("U8.DZ.JA.Res030")
        ExecFunEffectiveCheck = False

        Exit Function

    End If

    If Voucher.headerText(StrcMaker) = "" Then
        MsgBox GetString("U8.DZ.JA.Res1130"), vbInformation, GetString("U8.DZ.JA.Res030")
        ExecFunEffectiveCheck = False

        Exit Function

    End If
    
    '     If Voucher.headerText("proccode") = "" Then
    '        MsgBox "��Ŀ��Ų���Ϊ��,�������Ŀ����������", vbInformation, GetString("U8.DZ.JA.Res030")
    '        ExecFunEffectiveCheck = False
    '        Exit Function
    '    End If
    
    '��鲿����Ч��
    If Voucher.headerText("chdepartcode") <> "" Then
        sql = "select dDepEndDate from department where cdepcode='" + Voucher.headerText("chdepartcode") + "'"
       
        rs.Open sql, g_Conn

        If Not rs.EOF Or rs.BOF Then
            If Not IsNull(rs("dDepEndDate")) Then
                If DateDiff("d", CDate(rs("dDepEndDate")), CDate(Voucher.headerText("ddate"))) >= 0 Then
                    MsgBox GetString("U8.ST.V870.00290"), vbInformation, GetString("U8.DZ.JA.Res030")
                    ExecFunEffectiveCheck = False

                    Exit Function

                End If
            End If
        End If

        rs.Close
    End If
    
    
      If Voucher.headerText("contype") = "��ͨ��ͬ" And Voucher.headerText("sourcetype") = "���պ�ͬ" Then
               If Null2Something(Voucher.headerText("conpaymoney")) <> "" Then
'                    If Null2Something(Voucher.headerText("totalappmoney"), 0) > Null2Something(Voucher.headerText("totalpaymoney"), 0) Then
                        If Val(Null2Something(Voucher.headerText("appprice"), 0)) - Val(Val(Null2Something(Voucher.headerText("conpaymoney"), 0)) + Val(Null2Something(Voucher.headerText("designmoney"), 0)) - Val(Null2Something(Voucher.headerText("addesignmoney"), 0)) - Val(Null2Something(Voucher.headerText("totalappmoney"), 0)) + Val(numappprice)) > 0 Then
                            MsgBox "������ܴ��ں�ͬ�˶���ͬ�ܶ���ͬδ������Ϊ��ͬ�ܶ- �ۼ�������-Ԥ����Ʒ� ", vbInformation, "��ʾ"
                              ExecFunEffectiveCheck = False
                            Exit Function

                        End If

'                    Else
'
'                        If Null2Something(Voucher.headerText("appprice"), 0) > Val(Null2Something(Voucher.headerText("conpaymoney"), 0)) + Val(Null2Something(Voucher.headerText("designmoney"), 0)) - Val(Null2Something(Voucher.headerText("totalpaymoney"), 0)) - Val(Null2Something(Voucher.headerText("addesignmoney"), 0)) + Val(numappprice) Then
'                            MsgBox "������ܴ��ں�ͬ�˶���ͬ�ܶ���ͬδ������Ϊ��ͬ��- �ۼ�������-Ԥ����Ʒ�", vbInformation, "��ʾ"
'                             ExecFunEffectiveCheck = False
'                            Exit Function
'
'                        End If
'                    End If
                    
                Else
                
          
              
                  '  If Null2Something(Voucher.headerText("totalappmoney"), 0) > Null2Something(Voucher.headerText("totalpaymoney"), 0) Then
                        If Val(Null2Something(Voucher.headerText("appprice"), 0)) - Val(Val(Null2Something(Voucher.headerText("conmoney"), 0)) + Val(Null2Something(Voucher.headerText("designmoney"), 0)) - Val(Null2Something(Voucher.headerText("totalappmoney"), 0)) - Val(Null2Something(Voucher.headerText("addesignmoney"), 0)) + Val(numappprice)) > 0 Then
                            MsgBox "������ܴ��ں�ͬ�˶���ͬ�ܶ���ͬδ������Ϊ��ͬ�ܶ- �ۼ�������-Ԥ����Ʒ� ", vbInformation, "��ʾ"
                             ExecFunEffectiveCheck = False
                            Exit Function

                        End If
'
'                    Else
'
'                        If Null2Something(Voucher.headerText("appprice"), 0) > Val(Null2Something(Voucher.headerText("conmoney"), 0)) + Val(Null2Something(Voucher.headerText("designmoney"), 0)) - Val(Null2Something(Voucher.headerText("totalpaymoney"), 0)) - Val(Null2Something(Voucher.headerText("addesignmoney"), 0)) + Val(numappprice) Then
'                            MsgBox "������ܴ��ں�ͬ�˶���ͬ�ܶ���ͬδ������Ϊ��ͬ��- �ۼ�������-Ԥ����Ʒ� ", vbInformation, "��ʾ"
'                              ExecFunEffectiveCheck = False
'                            Exit Function
'
'                        End If
'                    End If
               
                End If
        End If
    
        If Voucher.headerText("sourcetype") = "������Ʒ�" Then
                          If Abs(Val(Null2Something(Voucher.headerText("appprice"), 0))) - Val(Abs(Val(Null2Something(Voucher.headerText("addesignmoney"), 0)) - Val(Null2Something(Voucher.headerText("designmoney"), 0)))) > 0 Then
                            MsgBox "������ܴ��� Ԥ����Ʒ�-��Ʒ�", vbInformation, "��ʾ"
                              ExecFunEffectiveCheck = False
                            Exit Function

                        End If
        
        End If
    
    
'    sql = "select  isnull(btype,0) as btype, isnull(stype,0) as stype from  HY_FYSL_Accounting where  ccode='" & Voucher.headerText("acccode") & "'"
'    rs.Open sql, g_Conn
'
'    If Not rs.EOF Then
'        If rs.Fields("stype") = 1 And Voucher.headerText("contype") = "��ͨ��ͬ" And Voucher.headerText("procode") = "" And Voucher.headerText("engcode") = "" Then
'
'            MsgBox "��ͨ��ͬ���������ĺ�ͬ����Ŀ����͹��̱��벻��Ϊ��", vbInformation, GetString("U8.DZ.JA.Res030")
'            ExecFunEffectiveCheck = False
'
'            Exit Function
'
'        End If
'
'        If rs.Fields("btype") = 1 And Voucher.headerText("contype") = "��ͨ��ͬ" And Voucher.headerText("engcode") = "" And Voucher.headerText("engcode") = "" Then
'
'            MsgBox "��ͨ��ͬ����������ĺ�ͬ�� ���̱��벻��Ϊ��", vbInformation, GetString("U8.DZ.JA.Res030")
'            ExecFunEffectiveCheck = False
'
'            Exit Function
'
'        End If
'
'    End If
'
    If Voucher.headerText("appprogm") = "���ȿ�" And Voucher.headerText("procode") = "" Then

        MsgBox "�տ����Ϊ���ȿ�,��Ŀ���ȼ������Ų���Ϊ��", vbInformation, GetString("U8.DZ.JA.Res030")
        ExecFunEffectiveCheck = False

        Exit Function

    End If
    
    

    '    If gcCreateType = "�ڳ�����" Then
    '        sql = "select CVALUE  from AccInformation where csysid='ST' AND CNAME='dSTFirstDate'"
    '        rs.Open sql, g_Conn
    '        If Not rs.BOF Or Not rs.EOF Then
    '            If DateDiff("d", CDate(rs.Fields("CVALUE")), CDate(Voucher.headerText("ddate"))) >= 0 Then
    '                MsgBox GetString("U8.DZ.JA.Res1860"), vbInformation, GetString("U8.DZ.JA.Res030")
    '                ExecFunEffectiveCheck = False
    '                Exit Function
    '            End If
    '        End If
    '    End If

    '    If Voucher.headerText("cfreight") = "��" Then
    '        If Voucher.headerText("MycdefineT2") = "" Or Voucher.headerText("cfreightType") = "" _
    '                Or Voucher.headerText("cfreightCost") = "" Then
    '            MsgBox GetString("U8.DZ.JA.Res1150"), vbInformation, GetString("U8.DZ.JA.Res030")
    '            ExecFunEffectiveCheck = False
    '            Exit Function
    '        End If
    '    End If
    '
    '    '���ݱ���У��
    '    If ExecFunBodyCheck(Voucher) = False Then
    '        ExecFunEffectiveCheck = False
    '        Exit Function
    '    End If

    ExecFunEffectiveCheck = True

    Exit Function

Err_Handler:
    MsgBox Err.Description, vbInformation, GetString("U8.DZ.JA.Res030")

End Function

'������Ч��У��
'�������Ƿ��
'���õ��������Ƿ�������
'����������Ƿ�Ϸ�
'���ν���
'��ⵥ��У��
'��Ч��У��
Private Function ExecFunBodyCheck(Voucher As Object) As Boolean
    On Error GoTo Err_Handler
    '������Ч��У��,���ڱ����Ƕ�����Ҫѭ��
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
        '��ʱ��ֵ,��Ϊ�б�
        '�˴���ֵ������
        '    bodyele.setAttribute "AutoID", irow

        '1
        '������,��У�飺������룬����
        '������У�����ʵ��ҵ�����
        If bodyele.getAttribute("cinvcode") = "" Then
            ReDim varArgs(0)
            varArgs(0) = iRow
            MsgBox GetStringPara("U8.DZ.JA.Res1160", varArgs(0)), vbInformation, GetString("U8.DZ.JA.Res030")
            '            MsgBox "��" & iRow & "�д�����벻��Ϊ��", vbInformation, GetString("U8.DZ.JA.Res030")
            ExecFunBodyCheck = False
            Exit Function
        End If

        If Val(bodyele.getAttribute("iquantity") & "") = 0 And Val(bodyele.getAttribute("inum") & "") = 0 Then
            '            MsgBox "��" & iRow & "�� ������������ͬʱΪ�ջ����0", vbInformation, getstring("U8.DZ.JA.Res030")
            ReDim varArgs(0)
            varArgs(0) = iRow
            MsgBox GetStringPara("U8.DZ.JA.Res1170", varArgs(0)), vbInformation, GetString("U8.DZ.JA.Res030")
            '            MsgBox "��" & iRow & "�� ����/��������ͬʱΪ�ջ����0", vbInformation, GetString("U8.DZ.JA.Res030")
            '��ʶ��ǰ��,����ɫ
            Voucher.row = iRow
            ExecFunBodyCheck = False
            Exit Function
        End If

  If gcCreateType = "�ڳ�����" Then
        'dxb �ֿ����¼�룡
        If bodyele.getAttribute("cwhcode") = "" Then
            ReDim varArgs(0)
            varArgs(0) = iRow
            MsgBox GetStringPara("U8.DZ.JA.Res1180", varArgs(0)), vbInformation, GetString("U8.DZ.JA.Res030")
            'MsgBox "��" & iRow & "�вֿⲻ��Ϊ��", vbInformation, GetString("U8.DZ.JA.Res030")
            ExecFunBodyCheck = False
            Exit Function
        End If
        sql = " select bwhpos from Warehouse where cwhcode='" & bodyele.getAttribute("cwhcode") & "' and bwhpos=1 "
        Set tRs = New ADODB.Recordset
        tRs.Open sql, g_Conn
        If Not tRs.BOF Or Not tRs.EOF Then
            '���ֿ����û�λ������������λ
            If bodyele.getAttribute("cPosition") & "" = "" Then
                ReDim varArgs(0)
                varArgs(0) = iRow
                MsgBox GetStringPara("U8.DZ.JA.Res1920", varArgs(0)), vbInformation, GetString("U8.DZ.JA.Res030")
                'MsgBox "��" & iRow & "�вֿⲻ��Ϊ��", vbInformation, GetString("U8.DZ.JA.Res030")
                ExecFunBodyCheck = False
                Exit Function
            End If
        End If
End If

        '2
        '���õ��������Ƿ�������
        '�ṹ�ԣ��ǽṹ���������������
        '����������Ƿ�Ϸ�
        If ExecFunFreeCheck(bodyele.getAttribute("cinvcode"), bodyele, iRow) = False Then
            '��ʶ��ǰ��,����ɫ
            Voucher.row = iRow
            'Voucher.SetCurrentRow ("@AutoID=" & bodyele.getAttribute("AutoID") & "")
            ExecFunBodyCheck = False
            Exit Function
        End If


        '��
        '����
        '����������

        If ExecFuncbatch(bodyele.getAttribute("cinvcode"), _
                bodyele.getAttribute("cbatch") & "", _
                bodyele.getAttribute("cinvouchcode") & "", _
                iRow) = False Then
            '��ʶ��ǰ��,����ɫ
            Voucher.row = iRow
            'Voucher.SetCurrentRow ("@AutoID=" & bodyele.getAttribute("AutoID") & "")
            ExecFunBodyCheck = False
            Exit Function
        End If
        '        If Voucher.headerText("dmDate") <> "" Then
        If gcCreateType = "�ڳ�����" Then
            sql = " select * from warehouse where cwhcode='" & bodyele.getAttribute("cwhcode") & "' and (isnull(dWhEndDate,'')='' or datediff(d,'" & g_oLogin.CurDate & "', dWhEndDate) >0 )"
            Set rs = Nothing
            rs.Open sql, g_Conn, 1, 1
            If rs.EOF Then
                ReDim varArgs(0)
                varArgs(0) = iRow
                MsgBox GetStringPara("U8.DZ.JA.Res1190", varArgs(0)), vbInformation, GetString("U8.DZ.JA.Res030")
                'MsgBox "��" & iRow & "�������ڴ��ڵ��ڲֿ��ʧЧ����", vbInformation, GetString("U8.DZ.JA.Res030")
                ExecFunBodyCheck = False
                Exit Function
            End If
        End If

        '        End If

        '        '4 ������У��
        '         If ExecFunMassDate(bodyele.getAttribute("cinvcode"), _
                  '                            iRow, bodyele.getAttribute("cMassDateUnit") & "", _
                  '                            bodyele.getAttribute("imassdate") & "", _
                  '                            bodyele.getAttribute("dmadedate") & "", _
                  '                            bodyele.getAttribute("dvdate") & "") = False Then
        '             '��ʶ��ǰ��,����ɫ
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
        '                                 MsgBox "��Ʒ[" + bodyele.getAttribute("cinvcode") + "]�ֿ�[" + bodyele.getAttribute("cwhcode") + "]����[" + bodyele.getAttribute("cbatch") + "]�������������", vbInformation, GetString("U8.DZ.JA.Res030")
        '                                 ExecFunBodyCheck = False
        '                                 Exit Function
        '                              End If
        '
        '                            End If
        '                        End If
        '                  Case 2
        '                       If iQ = -1 Then
        '                             MsgBox "��Ʒ[" + bodyele.getAttribute("cinvcode") + "]�ֿ�[" + bodyele.getAttribute("cwhcode") + "]����[" + bodyele.getAttribute("cbatch") + "]�������������", vbInformation, GetString("U8.DZ.JA.Res030")
        ''                                 ExecFunBodyCheck = False
        ''                                 Exit Function
        '                          End If
        '
        '                End Select

        '5 Ԥ�ƹ黹���� dxb
        If bodyele.getAttribute("backdate") = "" Then
            ReDim varArgs(0)
            varArgs(0) = iRow
            MsgBox GetStringPara("U8.DZ.JA.Res1200", varArgs(0)), vbInformation, GetString("U8.DZ.JA.Res030")
            'MsgBox "��" & iRow & "��Ԥ�ƹ黹���ڲ���Ϊ��", vbInformation, GetString("U8.DZ.JA.Res030")
            ExecFunBodyCheck = False
            Exit Function
        Else
            If CDate(Mid(bodyele.getAttribute("backdate"), 1, 10)) < CDate(Mid(Voucher.headerText(StrdDate), 1, 10)) Then
                ReDim varArgs(0)
                varArgs(0) = iRow
                MsgBox GetStringPara("U8.DZ.JA.Res1210", varArgs(0)), vbInformation, GetString("U8.DZ.JA.Res030")
                'MsgBox "��" & iRow & "��Ԥ�ƹ黹���ڲ������ڽ�������", vbInformation, GetString("U8.DZ.JA.Res030")
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




' 4 ������У��
Private Function ExecFunMassDate(cinvcode As String, _
                                 iRow As Integer, _
                                 cMassDateUnit As String, _
                                 imassdate As String, _
                                 dMadeDate As String, _
                                 dvdate As String) As Boolean
    Dim sql As String
    Dim rs As New ADODB.Recordset

    sql = "SELECT bInvQuality FROM inventory where cinvcode='" & cinvcode & "'"
    rs.Open sql, g_Conn, 1, 1
    If Not rs.EOF Then
        If rs("bInvQuality") = "1" Then

            If cMassDateUnit = "" Then
                ReDim varArgs(1)
                varArgs(0) = iRow
                varArgs(1) = cinvcode
                MsgBox GetStringPara("U8.DZ.JA.Res1230", varArgs(0), varArgs(1)), vbInformation, GetString("U8.DZ.JA.Res030")
                'MsgBox "��" & iRow & "�д��" & cinvcode & "���ñ����ڹ���,�����ڵ�λ����Ϊ��", vbInformation, GetString("U8.DZ.JA.Res030")
                ExecFunMassDate = False
                rs.Close
                Set rs = Nothing
                Exit Function
            End If
            If imassdate = "" Or Val(imassdate) = 0 Then
                ReDim varArgs(1)
                varArgs(0) = iRow
                varArgs(1) = cinvcode
                MsgBox GetStringPara("U8.DZ.JA.Res1240", varArgs(0), varArgs(1)), vbInformation, GetString("U8.DZ.JA.Res030")
                'MsgBox "��" & iRow & "�д��" & cinvcode & "���ñ����ڹ���,��������������Ϊ��", vbInformation, GetString("U8.DZ.JA.Res030")
                ExecFunMassDate = False
                rs.Close
                Set rs = Nothing
                Exit Function
            End If
            If dMadeDate = "" Then
                ReDim varArgs(1)
                varArgs(0) = iRow
                varArgs(1) = cinvcode
                MsgBox GetStringPara("U8.DZ.JA.Res1250", varArgs(0), varArgs(1)), vbInformation, GetString("U8.DZ.JA.Res030")
                'MsgBox "��" & iRow & "�д��" & cinvcode & "���ñ����ڹ���,�������ڲ���Ϊ��", vbInformation, GetString("U8.DZ.JA.Res030")
                ExecFunMassDate = False
                rs.Close
                Set rs = Nothing
                Exit Function
            End If
            If dvdate = "" Then
                ReDim varArgs(1)
                varArgs(0) = iRow
                varArgs(1) = cinvcode
                MsgBox GetStringPara("U8.DZ.JA.Res1260", varArgs(0), varArgs(1)), vbInformation, GetString("U8.DZ.JA.Res030")
                'MsgBox "��" & iRow & "�д��" & cinvcode & "���ñ����ڹ���,ʧЧ���ڲ���Ϊ��", vbInformation, GetString("U8.DZ.JA.Res030")
                ExecFunMassDate = False
                rs.Close
                Set rs = Nothing
                Exit Function
            End If

        End If
    Else
        ReDim varArgs(1)
        varArgs(0) = iRow
        varArgs(1) = cinvcode
        MsgBox GetStringPara("U8.DZ.JA.Res1270", varArgs(0), varArgs(1)), vbInformation, GetString("U8.DZ.JA.Res030")
        'MsgBox "��" & iRow & "�д������" & cinvcode & "������", vbInformation, GetString("U8.DZ.JA.Res030")
        ExecFunMassDate = False
        rs.Close
        Set rs = Nothing
        Exit Function
    End If

    ExecFunMassDate = True
    rs.Close
    Set rs = Nothing

End Function


'�ṹ��������

'���õ��������Ƿ�������
'�ṹ�ԣ��ǽṹ���������������
'����������Ƿ�Ϸ�
Private Function ExecFunFreeCheck(cinvcode As String, _
                                  bodyele As IXMLDOMElement, _
                                  iRow As Integer _
                                ) As Boolean
    On Error GoTo Err_Handler

    'cinvcode :bodyele.getAttribute("cinvcode")
    'cFree:bodyele.getAttribute("cfree1")
    'irow:irow �к�
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
                bRS.Open "select cItemName from userdef where cClass = '���' and cDicDbname='cFree" & i & "'", g_Conn, 1, 1
                ReDim varArgs(2)
                varArgs(0) = iRow
                varArgs(1) = cinvcode
                varArgs(2) = bRS("cItemName")
                MsgBox GetStringPara("U8.DZ.JA.Res1280", varArgs(0), varArgs(1), varArgs(2)), vbInformation, GetString("U8.DZ.JA.Res030")
                'MsgBox "��" & iRow & "�д��" & cinvcode & "������������" & bRS("cItemName") & ",��������", vbInformation, GetString("U8.DZ.JA.Res030")
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
            MsgBox GetStringPara("U8.DZ.JA.Res1290", varArgs(0)), vbInformation, GetString("U8.DZ.JA.Res030")
            'MsgBox "��" & iRow & "�д���ṹ����������ϲ��Ϸ�", vbInformation, GetString("U8.DZ.JA.Res030")
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
        MsgBox GetStringPara("U8.DZ.JA.Res1300", varArgs(0), varArgs(1)), vbInformation, GetString("U8.DZ.JA.Res030")
        'MsgBox "��" & iRow & "�д������" & cinvcode & "������", vbInformation, GetString("U8.DZ.JA.Res030")
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

'����
'��ⵥ��
Private Function ExecFuncbatch(cinvcode As String, cbatch As String, cTrackCode As String, iRow As Integer) As Boolean
    On Error GoTo Err_Handler

    Dim sql As String
    Dim rs As New ADODB.Recordset

    sql = "select bInvBatch,bTrack from inventory where cinvcode='" & cinvcode & "'"
    rs.Open sql, g_Conn, 1, 1
    If Not rs.EOF Then
        If gcCreateType = "�ڳ�����" And CBool(rs("bInvBatch")) And cbatch = "" Then

            ReDim varArgs(0)
            varArgs(0) = iRow
            MsgBox GetStringPara("U8.DZ.JA.Res1310", varArgs(0)), vbInformation, GetString("U8.DZ.JA.Res030")
            ' MsgBox "��" & iRow & "�д�����������ι������Ų���Ϊ��", vbInformation, GetString("U8.DZ.JA.Res030")
            ExecFuncbatch = False
            rs.Close
            Set rs = Nothing
            Exit Function
        End If
        '        If CBool(Rs("bTrack")) And cTrackCode = "" Then
        '            MsgBox "��" & iRow & "�д�������˳��������⣬��ⵥ�Ų���Ϊ��", vbInformation, getstring("U8.DZ.JA.Res030")
        '            ExecFuncbatch = False
        '            Rs.Close
        '            Set Rs = Nothing
        '            Exit Function
        '        End If
    Else
        ReDim varArgs(1)
        varArgs(0) = iRow
        varArgs(1) = cinvcode
        MsgBox GetStringPara("U8.DZ.JA.Res1300", varArgs(0), varArgs(1)), vbInformation, GetString("U8.DZ.JA.Res030")
        'MsgBox "��" & iRow & "�д������" & cinvcode & "������", vbInformation, GetString("U8.DZ.JA.Res030")
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




'******************************************************************************
'                           �������У�� begin
'******************************************************************************




'�������У��
'�ɴ���������������Ŀ��
'               ������ơ�������롢�����ͺ�
'               ������λ����롢���ơ���������λ���롢���ơ���������λ���ơ����롢������
'               �����ڵ�λ������������
'               1-16����Զ�����

Public Function cInvCodeRefer(cinvcode As String) As ADODB.Recordset
    On Error GoTo Err_Handler

    Dim sql As String
    Dim rs As New ADODB.Recordset
    
    rs.CursorLocation = adUseClient
    
    Dim bByCode As Boolean
    rs.Open "select top 1 cinvcode from inventory with (nolock) where cinvcode =N'" & cinvcode & "'", g_Conn, adOpenDynamic, adLockReadOnly
    If rs.RecordCount > 0 Then
       bByCode = True
    Else
       bByCode = False
    End If
    
    rs.Close
    

    '��������λ,Ĭ��ȡ��������λ,�ɸ��ݵ��ݵ�ҵ��Ҫ�����

    '/*B*/ �˴����ݵ���ģ�����õ��ֶ���Ҫ���޸Ĳ�ѯ���

    sql = "SELECT i.cInvCode,i.cInvName,i.cInvStd,i.cInvAddCode," & _
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
            " left outer join ComputationUnit u2 on i.cSTComUnitCode=u2.cComunitCode "
            
    If bByCode = False Then
         sql = sql + "  where (i.cinvcode=N'" & cinvcode & "' or cinvname=N'" & cinvcode & "' or cinvstd=N'" & cinvcode & "' or cInvMnemCode =N'" & cinvcode & "') and '" & g_oLogin.CurDate & "' <=isnull(dEDate,'2099-12-31')"
    Else
         sql = sql + "  where (i.cinvcode=N'" & cinvcode & "') and '" & g_oLogin.CurDate & "' <=isnull(dEDate,'2099-12-31')"
    End If

'    If isQualityInv = False Then
'        sql = sql + " and i.bInvQuality <> 1 "
'    End If
            '" where (i.cinvcode='" & cinvcode & "' or cinvname='" & cinvcode & "' or cinvstd='" & cinvcode & "' or cInvMnemCode ='" & cinvcode & "') and i.bTrack<>1 and i.bInvQuality <> 1 and i.bSerial <> 1   and '" & g_oLogin.CurDate & "' <=isnull(dEDate,'2099-12-31')"
    '" where (i.cinvcode='" & cinvcode & "' or cinvname='" & cinvcode & "' or cinvstd='" & cinvcode & "' or cInvMnemCode ='" & cinvcode & "') and i.bTrack<>1 and i.bInvQuality <> 1 and i.bSerial <> 1 and i.bPropertyCheck <> 1   and '" & g_oLogin.CurDate & "' <=isnull(dEDate,'2099-12-31')"
    '" where (i.cinvcode='" & cinvcode & "' or cinvname='" & cinvcode & "' or cinvstd='" & cinvcode & "' or cInvMnemCode ='" & cinvcode & "') and i.bTrack<>1 and i.bInvQuality <> 1 and i.bSerial <> 1 and i.bPropertyCheck <> 1 and ( '" & g_oLogin.CurDate & "' >= dSDate and '" & g_oLogin.CurDate & "' <=isnull(dEDate,'2099-12-31'))"
    sql = sql & IIf(sAuth_invW = "", "", " and i.iid in (" & sAuth_invW & ")")
    rs.Open sql, g_Conn, 1, 1
    If Not rs.EOF Then
        Set cInvCodeRefer = rs
    Else
        Set cInvCodeRefer = Nothing
        rs.Close
        Set rs = Nothing
        Exit Function
    End If

    'rs.Close
    Set rs = Nothing
    Exit Function

Err_Handler:
    MsgBox Err.Description, vbInformation, GetString("U8.DZ.JA.Res030")

End Function

Public Function cWhCodeRefer(ccwhcode As String, sDate As String) As ADODB.Recordset
    On Error GoTo Err_Handler

    Dim sql As String
    Dim rs As New ADODB.Recordset
    sql = "SELECT * from Warehouse where (cwhcode='" & ccwhcode & "' or cwhname ='" & ccwhcode & "') and (isnull(dWhEndDate,'')='' or datediff(d,'" & CDate(Mid(sDate, 1, 10)) & "',IsNull(dWhEndDate,'2099-12-31'))>0)  and bProxyWh=0"
    sql = sql & IIf(sAuth_WareHouseW = "", "", " and cwhcode in (" & sAuth_WareHouseW & ")")

    rs.Open sql, g_Conn, 1, 1
    If Not rs.EOF Then
        Set cWhCodeRefer = rs
    Else
        Set cWhCodeRefer = Nothing
        rs.Close
        Set rs = Nothing
        Exit Function
    End If

    Set rs = Nothing
    Exit Function
Err_Handler:
    MsgBox Err.Description, vbInformation, GetString("U8.DZ.JA.Res030")
End Function


Public Function cPosionRefer(ccPosion As String, cwhcode As String) As ADODB.Recordset
    On Error GoTo Err_Handler

    Dim sql As String
    Dim rs As New ADODB.Recordset
    sql = "SELECT * from Position  where (cPosCode ='" & ccPosion & "' or cposname ='" & ccPosion & "') and cwhcode='" & cwhcode & "' "
    sql = sql & IIf(sAuth_PositionW = "", "", " and cPosCode in (" & sAuth_PositionW & ")")

    rs.Open sql, g_Conn, 1, 1
    If Not rs.EOF Then
        Set cPosionRefer = rs
    Else
        Set cPosionRefer = Nothing
        rs.Close
        Set rs = Nothing
        Exit Function
    End If

    Set rs = Nothing
    Exit Function
Err_Handler:
    MsgBox Err.Description, vbInformation, GetString("U8.DZ.JA.Res030")
End Function

'����������,�Զ����帳ֵ

Public Sub SetBodyCellValue(Voucher As Object, rs As ADODB.Recordset, iRow As Long)
    Dim i As Integer
    Dim iele As IXMLDOMElement
    Dim domBodyLine As DOMDocument

    Set domBodyLine = Voucher.GetLineDom(iRow)

    '/*B*/ ���ݵ���ģ����Ҫ���ֶθ�ֵ

    For Each iele In domBodyLine.selectNodes("//z:row")
        iele.setAttribute "cinvcode", rs("cinvcode") & ""
        iele.setAttribute "cinvname", rs("cInvName")
        iele.setAttribute "cinvstd", rs("cInvStd") & ""
        iele.setAttribute "cinvaddcode", rs("cInvAddCode") & ""
        iele.setAttribute "igrouptype", rs("iGroupType") & ""
        iele.setAttribute "cgroupcode", rs("cGroupCode") & ""
        iele.setAttribute "cgroupname", rs("cGroupName") & ""
        iele.setAttribute "ccomunitcode", rs("cComUnitCode") & ""    '��������λ
        iele.setAttribute "ccomunitname", rs("cComunitName") & ""
        iele.setAttribute "cunitid", rs("cAssUnit") & ""   '��������λ����
        iele.setAttribute "cinva_unit", rs("cAssUnitName") & ""    '��������λ����
        iele.setAttribute "iinvexchrate", rs("iinvexchrate") & ""    '������
        iele.setAttribute "cmassunit", rs("cMassUnit") & ""
        iele.setAttribute "imassdate", rs("imassdate") & ""
        iele.setAttribute "itaxrate", 17                   '˰��,Ĭ��17
        iele.setAttribute "kl", 100                        '����
        iele.setAttribute "kl2", 100                       '����2
        iele.setAttribute "iExpiratDateCalcu", Null2Something(rs("iExpiratDateCalcu"))    ' ��Ч�����㷽ʽ

        If iele.getAttribute("iquantity") <> "" Then iele.setAttribute "iquantity", ""
        If iele.getAttribute("inum") <> "" Then iele.setAttribute "inum", ""
        If iele.getAttribute("backdate") = "" Then
            iele.setAttribute "backdate", g_oLogin.CurDate
            If Voucher.headerText("ddate") <> "" Then
                If CDate(Voucher.headerText("ddate")) > g_oLogin.CurDate Then
                    iele.setAttribute "backdate", CDate(Voucher.headerText("ddate"))
                End If
            End If
        End If
        '        If Voucher.bodyText(iRow, "iquantity") <> "" Then iele.setAttribute "iquantity", ""
        '        If Voucher.bodyText(iRow, "inum") <> "" Then iele.setAttribute "inum", ""

        If Voucher.VoucherStatus = VSeAddMode Then
            Voucher.bodyText(iRow, sAutoId) = iRow
        Else
            Call GetMaxIDs
            Voucher.bodyText(iRow, sID) = lngVoucherID
            Voucher.bodyText(iRow, sAutoId) = sAutoId
            '�˻�����id,�����˷�id��
            '            g_Conn.Execute "update ufsystem..ua_identity  set Ifatherid=" & sID - 1 & " where cacc_id='" & g_oLogin.cAcc_Id & "' and cvouchtype='" & gstrCardNumber & "'"
        End If

        '������
        For i = 1 To 10
            iele.setAttribute "cfree" & i, ""
        Next i

        '�Ƿ����۶���
        For i = 1 To 10
            iele.setAttribute "bsalepricefree" & i, Null2Something(rs("bSalePriceFree" & i))
        Next i

        '����Զ�����
        For i = 1 To 16
            iele.setAttribute "cInvDefine" & i, Null2Something(rs("cInvDefine" & i)) & ""
        Next i
        
        '�����ι��������������
        If Not CBool(Null2Something(rs("bInvBatch"), "0")) Then
            iele.setAttribute "cbatch", ""
        End If
        
        Voucher.UpdateLineData domBodyLine, iRow

        '���Ʊ����ֶεĿɱ༭��
        SetBodyControl Voucher, rs, iRow

    Next

    Set iele = Nothing
    Set domBodyLine = Nothing

    '���Ʊ����ֶεĿɱ༭��
    SetBodyControl Voucher, rs, iRow
End Sub


'���ݵ����¼�
Public Sub voucher_click(section As UAPVoucherControl85.SectionsConstants, ByVal Index As Long, Voucher As Object)
    '����,�����Ƿ�ɱ༭
    If Index = sibody Then
        Dim rs As New ADODB.Recordset
        Dim cinvcode As String
        cinvcode = Voucher.bodyText(Voucher.row, "cinvcode")

        Set rs = cInvCodeRefer(CStr(cinvcode))
        If rs Is Nothing Or rs.State = 0 Then
            ReDim varArgs(0)
            varArgs(0) = cinvcode
            MsgBox GetStringPara("U8.DZ.JA.Res1320", varArgs(0)), vbInformation, GetString("U8.DZ.JA.Res030")
            ' MsgBox GetString("U8.DZ.JA.Res780") & cinvcode & "�����ڻ���û���������Ի���û��Ȩ�޻���ͣ�ã�����������", vbInformation, GetString("U8.DZ.JA.Res030")
            Exit Sub
        End If
        '���Ʊ����ֶεĿɱ༭��
        SetBodyControl Voucher, rs, Voucher.row

        rs.Close
        Set rs = Nothing

    End If
End Sub



'���Ʊ����ֶεĿɱ༭��
Public Sub SetBodyControl(Voucher As Object, rs As ADODB.Recordset, iRow As Long)

    '��������λ�Ƿ�ɱ༭
    'ֻ�й̶������ʵĴ�����������ſ��Ա༭
    If rs("iGroupType") = 1 Then                           '�̶�������
        Voucher.ClearDisibleColor2 iRow, "cunitid", RGB(255, 255, 255)
        Voucher.ClearDisibleColor2 iRow, "cinva_unit", RGB(255, 255, 255)
            Voucher.ClearDisibleColor2 iRow, "inum", RGB(255, 255, 255)
        Voucher.SetDisibleColor2 iRow, "iinvexchrate"
    Else
        Voucher.SetDisibleColor2 iRow, "cunitid"
        Voucher.SetDisibleColor2 iRow, "cinva_unit"

        '����������
        If rs("iGroupType") = "2" Then
            Voucher.ClearDisibleColor2 iRow, "iinvexchrate", RGB(255, 255, 255)
            Voucher.ClearDisibleColor2 iRow, "inum", RGB(255, 255, 255)
            '�޻�����
        Else
            Voucher.SetDisibleColor2 iRow, "iinvexchrate"
            Voucher.SetDisibleColor2 iRow, "inum"
        End If

    End If



    '�Ƿ����ι���
    If rs("bInvBatch") = True Then
        Voucher.ClearDisibleColor2 iRow, "cbatch", RGB(255, 255, 255)
    Else
        Voucher.SetDisibleColor2 iRow, "cbatch"
    End If

    '�Ƿ����ڹ���
    If rs("bInvQuality") = True Then
        Voucher.ClearDisibleColor2 iRow, "dmadedate", RGB(255, 255, 255)
        Voucher.ClearDisibleColor2 iRow, "dvdate", RGB(255, 255, 255)
    Else
        Voucher.SetDisibleColor2 iRow, "dmadedate"
        Voucher.SetDisibleColor2 iRow, "dvdate"
    End If

    '�Ƿ����������
    If rs("bTrack") = True Then
        Voucher.ClearDisibleColor2 iRow, "cinvouchcode", RGB(255, 255, 255)
    Else
        Voucher.SetDisibleColor2 iRow, "cinvouchcode"
    End If

    '������
    Dim i As Integer
    For i = 1 To 10
        If rs("bFree" & i) = 0 Then
            Voucher.SetDisibleColor2 iRow, "cfree" & i
        Else
            Voucher.ClearDisibleColor2 iRow, "cfree" & i, RGB(255, 255, 255)
        End If
    Next i

End Sub
'******************************************************************************
'                           �������У�� end
'******************************************************************************



'��ȡ��ͷ����
Public Function GetRateValue(retvalue As String, dDate As String) As Double
    On Error GoTo Err_Handler

    '�������е�[���ʷ�ʽ]����ʹ�õ���[�̶�����]����[��������]
    'true:�̶����� ,false:��������
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim RateType As Boolean
    sql = "SELECT cValue from accinformation where cname='iXchgRateStl' and csysid='AA'"
    rs.Open sql, g_Conn, 1, 1
    If Not rs.EOF Then
        RateType = CBool(rs("cvalue"))
    Else
        RateType = True
    End If

    rs.Close
    Set rs = Nothing

    '�̶����� itype=2
    If RateType = True Then
        sql = "select nflat,case when f.bcal=1  then '*' else '/' end symbol from exch e" & _
                " inner join foreigncurrency f on e.cexch_name=f.cexch_name " & _
                " where e.cexch_name='" & retvalue & "' and iperiod=" & Month(dDate) & " and itype=2"

        '�������� itype=1
    Else
        sql = "select nflat ,case when f.bcal=1  then '*' else '/' end symbol from exch e" & _
                " inner join foreigncurrency f on e.cexch_name=f.cexch_name " & _
                " where e.cexch_name='" & retvalue & "' and iperiod=" & Month(dDate) & " and itype=1 and cdate='" & Format(dDate, "YYYY.MM.DD") & "'"
    End If

    rs.Open sql, g_Conn, 1, 1
    If Not rs.EOF Then
        GetRateValue = rs("nflat")
        symbol = rs("symbol")
    Else
        GetRateValue = 1
        symbol = "*"
    End If

    rs.Close
    Set rs = Nothing
    Exit Function


Err_Handler:
    rs.Close
    Set rs = Nothing
    GetRateValue = 1
    MsgBox Err.Description, vbInformation, GetString("U8.DZ.JA.Res030")

End Function

Public Function isQualityInv() As Boolean
    On Error GoTo Error_General_Handler
    isQualityInv = False
    '�޸�ʧЧ����ֻ���㱣���ڣ���������������
    
    Dim oDom As DOMDocument
    Dim sPath As String
    Dim oEle1 As IXMLDOMElement
   
    
    sPath = App.Path
    
    If Right(sPath, 1) <> "\" Then
        sPath = sPath & "\"
    End If
    sPath = sPath & "STVoucherSet.xml"
    Set oDom = New DOMDocument
    If oDom.Load(sPath) Then
        Set oEle1 = oDom.selectSingleNode("//Borrow")
        If Not oEle1 Is Nothing Then
            If Not IsNull(oEle1.getAttribute("isQualityInv")) Then
                isQualityInv = CBool(oEle1.getAttribute("isQualityInv"))
            End If
        End If
    End If
exit_handle:
    
    Set oDom = Nothing
    Exit Function
    
Error_General_Handler:

    isQualityInv = False
    GoTo exit_handle

End Function




