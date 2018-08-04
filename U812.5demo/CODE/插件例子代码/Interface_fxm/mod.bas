Attribute VB_Name = "mod"
Public nvLogin As Object
Public rows As Integer
Public con As New ADODB.Connection '数据库连接
Public ccode As String             '销售订单号
Public iSoId As String              '生产订单ID
Public SoKey As Long                '销售订单ID
Public cInv As String               '产品编码
Public nvRs As New ADODB.Recordset
Public editprop As String           '编辑状态
Public nvsql As String
Public cls_Public As Object
Public BtnErr As String



'取主OM对象值
'函数功能 ：得到DOM对象中指定元素的值
'domBody  dom 对象
'sKey   关键字名称
Public Function GetHeadItemValue(ByVal domHead As DOMDocument, ByVal sKey As String) As String
    sKey = LCase(sKey)
    If Not domHead.selectSingleNode("//z:row").Attributes.getNamedItem(sKey) Is Nothing Then
        GetHeadItemValue = domHead.selectSingleNode("//z:row").Attributes.getNamedItem(sKey).nodeValue
    Else
        GetHeadItemValue = ""
    End If
End Function


'取子DOM对象值
'函数功能 ：得到DOM对象中指定元素的值
'domBody  dom 对象
'sKey   关键字名称
'R      行号
Public Function GetBodyItemValue(ByVal domBody As DOMDocument, ByVal sKey As String, ByVal r As Long) As String
    sKey = LCase(sKey)
    If Not domBody.selectNodes("//z:row").Item(r).Attributes.getNamedItem(sKey) Is Nothing Then
        GetBodyItemValue = domBody.selectNodes("//z:row").Item(r).Attributes.getNamedItem(sKey).nodeValue
    Else
        GetBodyItemValue = ""
    End If
End Function
'不区分大小写
'取子DOM对象值
'函数功能 ：得到DOM对象中指定元素的值
'domBody  dom 对象
'sKey   关键字名称
'R      行号
Public Function GetBodyItemValue1(ByVal domBody As DOMDocument, ByVal sKey As String, ByVal r As Long) As String
    
    If Not domBody.selectNodes("//z:row").Item(r).Attributes.getNamedItem(sKey) Is Nothing Then
        GetBodyItemValue1 = domBody.selectNodes("//z:row").Item(r).Attributes.getNamedItem(sKey).nodeValue
    Else
        GetBodyItemValue1 = ""
    End If
End Function



'修改
Public Function Update(ByVal cSql As String, DBconn As ADODB.Connection) As String
    On Error GoTo ExitUpdate
    Dim L As Integer
    '更新数据库
'    DBConn.Execute cSql, L
    
    DBconn.Execute cSql, L
    If L <> 1 Then
    
    End If
    
    Exit Function
ExitUpdate:
    Update = Err.Description
End Function

'保存
Public Function Save(ByVal cSql As String, DBconn As ADODB.Connection) As String
    On Error GoTo ExitSave
    
    
    
    
    
    
ExitSave:
    Save = Err.Description
End Function

'删除
Public Function Delete(ByVal cSql As String, DBconn As ADODB.Connection) As String
    On Error GoTo ExitDelete
        
    DBconn.Execute cSql
    Exit Function
ExitDelete:
    Delete = Err.Description
End Function


'查询
Public Function Query(ByVal cSql As String, DBconn As ADODB.Connection, Optional strUserErr As String) As ADODB.Recordset
    Dim QueryRs As New ADODB.Recordset
    On Error GoTo ExitQuerys
    
    If QueryRs.State = 1 Then QueryRs.Close
    QueryRs.Open cSql, DBconn.ConnectionString, adOpenKeyset, adLockPessimistic
    
    Set Query = QueryRs
    
    Exit Function
ExitQuerys:
    Set Query = Nothing
    Set QueryRs = Nothing
    strUserErr = Err.Description
End Function

'检查数据合法性
Public Function CheckData(DBconn As ADODB.Connection, CardNum As String, domHead As DOMDocument, domBody As DOMDocument, Optional strUserErr As String) As Boolean
    Dim sql As String
    Dim Rs As New ADODB.Recordset
    Dim cInvCode As String
    On Error GoTo Errhandle
    
    Select Case LCase(CardNum)
        Case LCase("YXEF9101")
            cInvCode = GetHeadItemValue(domHead, "t_cinvcode")
            '检验存货编码是否合法
            sql = "select cInvCode from Inventory  where cInvCode='" & cInvCode & "'"
            Rs.Open sql, DBconn.ConnectionString, 3, 4
            If Not (Rs.RecordCount > 0) Then
                strUserErr = "该存货编码不存在，请选择正确的存货编码！！！"
                CheckData = False
                Exit Function
            End If
            If Rs.State = 1 Then Rs.Close
            Set Rs = Nothing
            
            '产品档案参照 存货档案生单时 存货编码不可以重复
            sql = "select t_cinvcode from EF_Inventory_Information where t_cinvcode='" & cInvCode & "'"
            Rs.Open sql, DBconn.ConnectionString, 3, 4
            If Rs.RecordCount > 0 Then
                strUserErr = "该存货编码已被参照，请另选存货编码进行参照！！！"
                CheckData = False
                Exit Function
            End If
            If Rs.State = 1 Then Rs.Close
    End Select
    CheckData = True
    
Exit Function
Errhandle:
    CheckData = False
    strUserErr = strUserErr & Err.Description
End Function


Public Function isCheck(DBconn As ADODB.Connection, CardNum As String, strKey As String, Optional voucher As Object) As String
    Dim sql As String
    Dim Rs As New ADODB.Recordset
    Dim person() As String
    On Error GoTo Errhandle
    Select Case LCase(CardNum)
        Case LCase("YXEF9101"), LCase("YXEF9122"), LCase("YXEF9123"), LCase("YXEF9124"), LCase("YXEF9125"), LCase("YXEF9126")
            sql = "SELECT refertype FROM VoucherItems_base WHERE CARDNUM ='" & CardNum & "' AND FIELDNAME ='" & strKey & "'"
            
            If Rs.State = 1 Then Rs.Close
            Rs.Open sql, DBconn.ConnectionString, adOpenKeyset, adLockPessimistic
            
            If Not Rs.BOF And Not Rs.EOF Then
                isCheck = CStr(Rs.Fields("refertype"))
            Else
                isCheck = ""
            End If
           
            If Rs.State = 1 Then Rs.Close
            
        Case LCase("YXEF9104"), LCase("YXEF9114"), LCase("YXEF9115"), LCase("YXEF9116"), LCase("YXEF9117"), LCase("YXEF9118"), LCase("YXEF9119"), LCase("YXEF9120"), LCase("YXEF9131"), LCase("YXEF9132"), LCase("YXEF9133"), LCase("YXEF9134"), LCase("YXEF9135"), LCase("YXEF9136")
            If strKey = "t_cpersoncode" Then
                
                If voucher.headerText(strKey) = "" Or IsNull(voucher.headerText(strKey)) Then
                    isCheck = "true"
                    Exit Function
                End If
                
                person = Split(voucher.headerText(strKey), "/")

                If UBound(person) > 0 Then
                    For i = 0 To (UBound(person) - LBound(person))
                        sql = "select * from hr_hi_person where cPsn_Num ='" & person(i) & "'"
                        If Rs.State = 1 Then Rs.Close
                        Rs.Open sql, DBconn.ConnectionString, adOpenKeyset, adLockPessimistic
                        
                        If Not Rs.BOF And Not Rs.EOF Then
                            isCheck = "true"   '系统中有当前的记录
                        Else
                            isCheck = "false"    '系统中无当前的记录
                            Exit Function
                        End If
                    Next i
                Else
                    If UBound(person) = 0 Then
                        sql = "select * from hr_hi_person where cPsn_Num ='" & person(i) & "'"
                        If Rs.State = 1 Then Rs.Close
                        Rs.Open sql, DBconn.ConnectionString, adOpenKeyset, adLockPessimistic
                        
                        If Not Rs.BOF And Not Rs.EOF Then
                            isCheck = "true"   '系统中有当前的记录
                        Else
                            isCheck = "false"    '系统中无当前的记录
                            Exit Function
                        End If
                    End If
                End If
                            
                
            End If
            
            If strKey = "t_cpersonname" Then
                
                If voucher.headerText(strKey) = "" Or IsNull(voucher.headerText(strKey)) Then
                    isCheck = "true"
                    Exit Function
                End If
                person = Split(voucher.headerText(strKey), "/")
                If UBound(person) > 0 Then
                    For i = 0 To (UBound(person) - LBound(person))
                        sql = "select * from hr_hi_person where cPsn_Name ='" & person(i) & "'"
                        If Rs.State = 1 Then Rs.Close
                        Rs.Open sql, DBconn.ConnectionString, adOpenKeyset, adLockPessimistic
                        
                        If Not Rs.BOF And Not Rs.EOF Then
                            isCheck = "true"   '系统中有当前的记录
                        Else
                            isCheck = "false"    '系统中无当前的记录
                            Exit Function
                        End If
                    Next i
                Else
                    If UBound(person) = -1 Then
                        sql = "select * from hr_hi_person where cPsn_Name ='" & person(i) & "'"
                        If Rs.State = 1 Then Rs.Close
                        Rs.Open sql, DBconn.ConnectionString, adOpenKeyset, adLockPessimistic
                        
                        If Not Rs.BOF And Not Rs.EOF Then
                            isCheck = "true"   '系统中有当前的记录
                        Else
                            isCheck = "false"    '系统中无当前的记录
                            Exit Function
                        End If
                    End If
                End If
                
                
            End If
            
            
            
            If strKey = "b_cpersoncode" Then
                
                If voucher.bodyText(r, strKey) = "" Or IsNull(voucher.bodyText(r, strKey)) Then
                    isCheck = "true"
                    Exit Function
                End If
                
                person = Split(voucher.bodyText(r, strKey), "/")

                If UBound(person) > 0 Then
                    For i = 0 To (UBound(person) - LBound(person))
                        sql = "select * from hr_hi_person where cPsn_Num ='" & person(i) & "'"
                        If Rs.State = 1 Then Rs.Close
                        Rs.Open sql, DBconn.ConnectionString, adOpenKeyset, adLockPessimistic
                        
                        If Not Rs.BOF And Not Rs.EOF Then
                            isCheck = "true"   '系统中有当前的记录
                        Else
                            isCheck = "false"    '系统中无当前的记录
                            Exit Function
                        End If
                    Next i
                Else
                    If UBound(person) = 0 Then
                        sql = "select * from hr_hi_person where cPsn_Num ='" & person(i) & "'"
                        If Rs.State = 1 Then Rs.Close
                        Rs.Open sql, DBconn.ConnectionString, adOpenKeyset, adLockPessimistic
                        
                        If Not Rs.BOF And Not Rs.EOF Then
                            isCheck = "true"   '系统中有当前的记录
                        Else
                            isCheck = "false"    '系统中无当前的记录
                            Exit Function
                        End If
                    End If
                End If
                            
                
            End If
            
            If strKey = "b_cpersonname" Then
                
                If voucher.bodyText(r, strKey) = "" Or IsNull(voucher.bodyText(r, strKey)) Then
                    isCheck = "true"
                    Exit Function
                End If
                person = Split(voucher.headerText(strKey), "/")
                If UBound(person) > 0 Then
                    For i = 0 To (UBound(person) - LBound(person))
                        sql = "select * from hr_hi_person where cPsn_Name ='" & person(i) & "'"
                        If Rs.State = 1 Then Rs.Close
                        Rs.Open sql, DBconn.ConnectionString, adOpenKeyset, adLockPessimistic
                        
                        If Not Rs.BOF And Not Rs.EOF Then
                            isCheck = "true"   '系统中有当前的记录
                        Else
                            isCheck = "false"    '系统中无当前的记录
                            Exit Function
                        End If
                    Next i
                Else
                    If UBound(person) = -1 Then
                        sql = "select * from hr_hi_person where cPsn_Name ='" & person(i) & "'"
                        If Rs.State = 1 Then Rs.Close
                        Rs.Open sql, DBconn.ConnectionString, adOpenKeyset, adLockPessimistic
                        
                        If Not Rs.BOF And Not Rs.EOF Then
                            isCheck = "true"   '系统中有当前的记录
                        Else
                            isCheck = "false"    '系统中无当前的记录
                            Exit Function
                        End If
                    End If
                End If
                
                
            End If
            
            
    End Select
    
    Exit Function
    
Errhandle:
     isCheck = "Err" & Err.Description
     MsgBox "FXM001：" & Err.Description
End Function

 
 '设置单据项目的值
 Public Function setVoucher(DBconn As ADODB.Connection, CardNum As String, strKey As String, where As String, Optional voucher As Object)
    Dim sql As String
    Dim Rs As New ADODB.Recordset
    Dim tmp As String
    Dim pername As String
    Dim arr_Per() As String
    Dim i As Integer
    
    On Error GoTo Errhandle
    
    Select Case LCase(CardNum)
        Case LCase("YXEF9104"), LCase("YXEF9105"), LCase("YXEF9114"), LCase("YXEF9115"), LCase("YXEF9116"), LCase("YXEF9117"), LCase("YXEF9118"), LCase("YXEF9119"), LCase("YXEF9120"), LCase("YXEF9131"), LCase("YXEF9132"), LCase("YXEF9133"), LCase("YXEF9134"), LCase("YXEF9135"), LCase("YXEF9136")
            Select Case LCase(strKey)
                Case LCase("t_cpersonname"), LCase("t_cpersoncode")
                     
                    tmp = ""
                    pername = ""
                    If where <> "" Then
'                        If LCase(strKey) = LCase("t_cpersoncode") Then
'                            sql = "select * from hr_hi_person where cPsn_Num in (" & where & ")"
'                        End If
'
'                        If LCase(strKey) = LCase("t_cpersonname") Then
'                            sql = "select * from hr_hi_person where cPsn_Name in (" & where & ")"
'                        End If
'                        If Rs.State = 1 Then Rs.Close
'
'                        Rs.Open sql, DBconn.ConnectionString, adOpenKeyset, adLockPessimistic
'                        Do Until Rs.EOF
'                            tmp = tmp & Rs.Fields("cPsn_Num")
'                            pername = pername & Rs.Fields("cPsn_Name")
'                            Rs.MoveNext
'                            If Not Rs.BOF And Not Rs.EOF Then
'                                tmp = tmp & "/"
'                                pername = pername & "/"
'                            End If
'                        Loop
                        
                        arr_Per = Split(where, ",")
                        
                        If UBound(arr_Per) >= 0 Then
                            For i = 0 To (UBound(arr_Per) - LBound(arr_Per))
                                where = arr_Per(i)
                                
                                If LCase(strKey) = LCase("t_cpersoncode") Then
                                    sql = "select * from hr_hi_person where cPsn_Num = " & where
                                End If
                                
                                If LCase(strKey) = LCase("t_cpersonname") Then
                                    sql = "select * from hr_hi_person where cPsn_Name = " & where
                                End If
                                
                                If Rs.State = 1 Then Rs.Close
                                Rs.Open sql, DBconn.ConnectionString, adOpenKeyset, adLockPessimistic
                                
                                Do Until Rs.EOF
                                    
                                    If tmp <> "" Or pername <> "" Then
                                        tmp = tmp & "/"
                                        pername = pername & "/"
                                    End If
                                    
                                    tmp = tmp & Rs.Fields("cPsn_Num")
                                    pername = pername & Rs.Fields("cPsn_Name")
                                    Rs.MoveNext
                                Loop
                                
                                
                            Next i
                        End If
                        
                        If Rs.State = 1 Then Rs.Close
'                        If strKey = "t_cpersonname" Then
'                            voucher.headerText("t_cpersoncode") = tmp
'                        End If
'
'                        If strKey = "t_cpersoncode" Then
'                            voucher.headerText("t_cpersonname") = tmp
'                        End If
                        
                        voucher.headerText("t_cpersoncode") = tmp
                        voucher.headerText("t_cpersonname") = pername
                        
                    End If
                    
                Case LCase("b_cpersonname"), LCase("b_cpersoncode")
                     
                    If where <> "" Then
'                        If LCase(strKey) = LCase("b_cpersoncode") Then
'                            sql = "select * from hr_hi_person where cPsn_Num in (" & where & ")"
'                        End If
'
'                        If LCase(strKey) = LCase("b_cpersonname") Then
'                            sql = "select * from hr_hi_person where cPsn_Name in (" & where & ")"
'                        End If
'                        If Rs.State = 1 Then Rs.Close
'
'                        Rs.Open sql, DBconn.ConnectionString, adOpenKeyset, adLockPessimistic
'                        Do Until Rs.EOF
'                            tmp = tmp & Rs.Fields("cPsn_Num")
'                            pername = pername & Rs.Fields("cPsn_Name")
'                            Rs.MoveNext
'                            If Not Rs.BOF And Not Rs.EOF Then
'                                tmp = tmp & "/"
'                                pername = pername & "/"
'                            End If
'                        Loop
'
'                        If Rs.State = 1 Then Rs.Close
                        
                        
                        
                        
                        
                        arr_Per = Split(where, ",")
                        
                        If UBound(arr_Per) >= 0 Then
                            For i = 0 To (UBound(arr_Per) - LBound(arr_Per))
                                where = arr_Per(i)
                                
                                If LCase(strKey) = LCase("b_cpersoncode") Then
                                    sql = "select * from hr_hi_person where cPsn_Num = " & where
                                End If
                                
                                If LCase(strKey) = LCase("b_cpersonname") Then
                                    sql = "select * from hr_hi_person where cPsn_Name = " & where
                                End If
                                
                                If Rs.State = 1 Then Rs.Close
                                Rs.Open sql, DBconn.ConnectionString, adOpenKeyset, adLockPessimistic
                                
                                Do Until Rs.EOF
                                
                                    If tmp <> "" Or pername <> "" Then
                                        tmp = tmp & "/"
                                        pername = pername & "/"
                                    End If
                                    
                                    tmp = tmp & Rs.Fields("cPsn_Num")
                                    pername = pername & Rs.Fields("cPsn_Name")
                                    Rs.MoveNext
                                Loop
                                
                            Next i
                        End If
                        
                        If Rs.State = 1 Then Rs.Close
                        
'                        If strKey = "b_cpersonname" Then
'                            voucher.headerText(rows, "b_cpersoncode") = tmp
'                        End If
'
'                        If strKey = "b_cpersoncode" Then
'                            voucher.headerText(rows, "b_cpersonname") = tmp
'                        End If
                        voucher.bodyText(rows, "b_cpersoncode") = tmp
                        voucher.bodyText(rows, "b_cpersonname") = pername
                        
                    End If
            End Select
            
    End Select
    Exit Function
    
Errhandle:
     MsgBox "FXM001：" & Err.Description
End Function

 
'检查数据合法性
Public Function CheckData_YXEF9110(DBconn As ADODB.Connection, CardNum As String, domHead As DOMDocument, domBody As DOMDocument, Optional strUserErr As String) As Boolean
   
    Dim checkSTRb As String
    Dim checkSTRa As String
    Dim sql As String
    Dim QueryRs As New ADODB.Recordset
    Dim str As String
    
    Dim id As String
    
    On Error GoTo ExitCheck
     
    Select Case LCase(CardNum)
    
        Case LCase("YXEF9110")
            '检查设备参数配置单表体设备是否重复
'            For i = 0 To domBody.selectNodes("//z:row").Length - 1
'                checkSTRb = GetBodyItemValue(domBody, "b_str2", i)
'                If i = domBody.selectNodes("//z:row").Length - 1 Then Exit For
'                For j = i + 1 To domBody.selectNodes("//z:row").Length - 1
'                    If LCase(GetBodyItemValue(domBody, "editprop", j)) <> LCase("D") Then
'                        checkSTRa = GetBodyItemValue(domBody, "b_str2", j)
'                    Else
'                        checkSTRa = ""
'                    End If
'                    If LCase(checkSTRb) = LCase(checkSTRa) Then
'                        CheckData_YXEF9110 = False
'                        strUserErr = "编码为：" & checkSTRb & "的设备重复！"
'                        Exit Function
'                    Else
'                        CheckData_YXEF9110 = True
'                    End If
'                Next j
'            Next i
            CheckData_YXEF9110 = True
            
        Case LCase("YXEF9105")
        
        
'            str = GetHeadItemValue(domHead, "str1")
'            '检查班组是否为空
'            If GetHeadItemValue(domHead, "str1") = "" Or IsNull(GetHeadItemValue(domHead, "str1")) Then
'                CheckData_YXEF9110 = False
'                strUserErr = "班组不能为空，请确认！"
'                Exit Function
'            End If
            
            '检查班组是否为空
            
            str = str & GetHeadItemValue(domHead, "str2")
            If GetHeadItemValue(domHead, "str2") = "" Or IsNull(GetHeadItemValue(domHead, "str2")) Then
                CheckData_YXEF9110 = False
                strUserErr = "班次不能为空，请确认！"
                Exit Function
            End If
            
            str = str & Format(Replace(GetHeadItemValue(domHead, "datetime1"), "T", " "), "yyyy-mm-dd")
            '检查班组是否为空
            If GetHeadItemValue(domHead, "datetime1") = "" Or IsNull(GetHeadItemValue(domHead, "datetime1")) Then
                CheckData_YXEF9110 = False
                strUserErr = "汇报日期不能为空，请确认！"
                Exit Function
            End If
            
            
            '保存前检查班组记录单表头 班组+班次+汇报日期 是否唯一
            id = GetHeadItemValue(domHead, "id")
            If id <> "" Then
                sql = "SELECT ID  FROM EF_CLASSGROUP WHERE  id =" & id & " and str5+str2+convert(varchar ,Year(datetime1) ,4)+'-'+ RIGHT('00'+convert(varchar,month(datetime1) ,2),2)+'-'+RIGHT('00'+convert(varchar,day(datetime1), 2),2) ='" & str & "'"
                If QueryRs.State = 1 Then QueryRs.Close
                QueryRs.Open sql, DBconn.ConnectionString, adOpenKeyset, adLockPessimistic
                If QueryRs.EOF Then
                    sql = "SELECT ID  FROM EF_CLASSGROUP WHERE  str5+str2+convert(varchar ,Year(datetime1) ,4)+'-'+ RIGHT('00'+convert(varchar,month(datetime1) ,2),2)+'-'+RIGHT('00'+convert(varchar,day(datetime1), 2),2) ='" & str & "'"
                    If QueryRs.State = 1 Then QueryRs.Close
                    QueryRs.Open sql, DBconn.ConnectionString, adOpenKeyset, adLockPessimistic
                    If Not QueryRs.EOF Then
                        CheckData_YXEF9110 = False
                        strUserErr = "班组+班次+汇报日期，已存在，请确认！"
                        If QueryRs.State = 1 Then QueryRs.Close
                        Exit Function
                    End If
                End If
            Else
                sql = "SELECT ID  FROM EF_CLASSGROUP WHERE str5+str2+convert(varchar ,Year(datetime1) ,4)+'-'+ RIGHT('00'+convert(varchar,month(datetime1) ,2),2)+'-'+RIGHT('00'+convert(varchar,day(datetime1), 2),2) ='" & str & "'"
                If QueryRs.State = 1 Then QueryRs.Close
                QueryRs.Open sql, DBconn.ConnectionString, adOpenKeyset, adLockPessimistic
                If Not QueryRs.EOF Then
                    CheckData_YXEF9110 = False
                    strUserErr = "班组+班次+汇报日期，已存在，请确认！"
                    If QueryRs.State = 1 Then QueryRs.Close
                    Exit Function
                End If
            End If
'
        Case LCase("YXEF9104")
        
        Case LCase("YXEF9122"), LCase("YXEF9123"), LCase("'YXEF9124"), LCase("'YXEF9125"), LCase("'YXEF9126")
        
            If GetHeadItemValue(domHead, "t_cinvcode") = "" Then
                
                CheckData_YXEF9110 = False
                strUserErr = "产品编码不可为空，请确认！"
                Exit Function
            
            End If
            
                       
            
            If GetHeadItemValue(domHead, "t_cvencode") = "" Then
                
'                CheckData_YXEF9110 = True
'                strUserErr = "" ' "制版厂编码不可为空，请确认！"
                CheckData_YXEF9110 = False
                strUserErr = "制版厂编码不可为空，请确认！"
                Exit Function
            
            End If
            
            
            If GetHeadItemValue(domHead, "t_ccuscode") = "" Then
                
'                CheckData_YXEF9110 = True
'                strUserErr = "" ' "客户编码不可为空，请确认！"
                CheckData_YXEF9110 = False
                strUserErr = "客户编码不可为空，请确认！"
                Exit Function
            
            End If
            
            
            sql = "select * from Inventory where cinvcode ='" & GetHeadItemValue(domHead, "t_cinvcode") & "'"
            If QueryRs.State = 1 Then QueryRs.Close
            QueryRs.Open sql, DBconn.ConnectionString, adOpenKeyset, adLockPessimistic
            
            If QueryRs.EOF Then
                CheckData_YXEF9110 = False
                strUserErr = "编号为：" & GetHeadItemValue(domHead, "t_cinvcode") & "产成品在系统中不存在" & "，请确认！"
                If QueryRs.State = 1 Then QueryRs.Close
                Exit Function
            End If
            
            If QueryRs.State = 1 Then QueryRs.Close
            
            If editprop <> "M" Then
                If GetHeadItemValue(domHead, "t_cinvcode") <> "" Then
                
                    sql = "select * from EF_Inventory_Information where t_cinvcode ='" & GetHeadItemValue(domHead, "t_cinvcode") & "'"
                    If QueryRs.State = 1 Then QueryRs.Close
                    
                    QueryRs.Open sql, DBconn.ConnectionString, adOpenKeyset, adLockPessimistic
                    
                    If Not QueryRs.EOF Then
                        CheckData_YXEF9110 = False
                        strUserErr = "编号为：" & GetHeadItemValue(domHead, "t_cinvcode") & "产成品档案已存在" & "，请确认！"
                        If QueryRs.State = 1 Then QueryRs.Close
                        Exit Function
                    End If
                    
                    If QueryRs.State = 1 Then QueryRs.Close
                
                End If
            End If
            
            sql = "select * from Vendor where cvencode ='" & GetHeadItemValue(domHead, "t_cvencode") & "'"
            If QueryRs.State = 1 Then QueryRs.Close
            QueryRs.Open sql, DBconn.ConnectionString, adOpenKeyset, adLockPessimistic
            
            If QueryRs.EOF Then
                CheckData_YXEF9110 = False
                strUserErr = "编号为：" & GetHeadItemValue(domHead, "t_cvencode") & "制版厂在系统中不存在" & "，请确认！"
                If QueryRs.State = 1 Then QueryRs.Close
                Exit Function
            End If
            
            If QueryRs.State = 1 Then QueryRs.Close
            
            If GetHeadItemValue(domHead, "t_ccuscode") <> "" Then
                
                sql = "select * from Customer where ccuscode ='" & GetHeadItemValue(domHead, "t_ccuscode") & "'"
                If QueryRs.State = 1 Then QueryRs.Close
                QueryRs.Open sql, DBconn.ConnectionString, adOpenKeyset, adLockPessimistic
                
                If QueryRs.EOF Then
                    CheckData_YXEF9110 = False
                    strUserErr = "编号为：" & GetHeadItemValue(domHead, "t_ccuscode") & "客户在系统中不存在" & "，请确认！"
                    If QueryRs.State = 1 Then QueryRs.Close
                    Exit Function
                End If
                If QueryRs.State = 1 Then QueryRs.Close
            
            End If
            
        Case Else
             CheckData_YXEF9110 = True
             
    End Select
    
    CheckData_YXEF9110 = True
 
    Exit Function
ExitCheck:
    CheckData_YXEF9110 = False
    strUserErr = strUserErr & Err.Description
End Function

Public Function GetID(tablename As String, moid As String) As String
   
    nvsql = " select Max(" & moid & ") as " & moid & " from  " & tablename
    
    If nvRs.State <> 0 Then nvRs.Close
    nvRs.CursorLocation = adUseClient
    nvRs.Open nvsql, con.ConnectionString, 3, 4
    If nvRs.EOF Then
        GetID = 1
    Else
        
        If IsNull(nvRs.Fields(moid)) Then
            GetID = 1
        Else
            GetID = nvRs.Fields(moid).Value + 1
        End If
    End If
    
End Function

Public Function GetCode(tablename As String, mCode As String) As String

    '--------------生成新的系统报工单主表单据号-------------------'
    nvsql = " select Max(" & mCode & ") as " & mCode & " from  " & tablename
    
    If nvRs.State <> 0 Then nvRs.Close
    nvRs.CursorLocation = adUseClient
    nvRs.Open nvsql, con.ConnectionString, 3, 4
    
    If nvRs.EOF Then
        GetCode = "0000000001"
    Else
        GetCode = Right("0000000000" & CStr(CDec(IIf(IsNull(nvRs.Fields(mCode).Value), 0, nvRs.Fields(mCode).Value)) + 1), 10)
    End If
    
End Function
Public Function RefProduct(soId As Long, OrderCode As String)
    Dim mRs  As New ADODB.Recordset
    Dim bomRs As New ADODB.Recordset
    Dim byRs As New ADODB.Recordset
    Dim zxRs As New ADODB.Recordset
    Dim rds As New ADODB.Recordset
    Dim invRS As New ADODB.Recordset
    
    Dim trues As Boolean
    Dim Dom As New DOMDocument
    Dim bomErrstr As String
    Dim iQuantity As Double
                
    Dim seq As Integer  '行号
    Dim moaSeq As Integer '序号
    
    Dim bDate As String   '开工日期
    Dim eDate As String   '完工日期
    
    Dim errInfo As String
    
    Dim ddID As Integer          '预生产订单主表id
    Dim ddAutoID As Integer      '预生产订单子表id,'预生产订单子件主表id
    Dim zjAutoID As Integer      '预生产订单子件子表id
    Dim mCode As String          '预生产订单编码
    
    Dim PRoutingId As Integer    '工艺路线id
    Dim BomId As Integer        'bom物料清单id
    Dim RountingType As Integer '工艺路线类别
    Dim BomType As Integer      'bom类别
    
    
    Dim vt_id_main As Integer   '预订单模板号
    Dim vt_id_son  As Integer   '预订单子件模板号
    
    On Error GoTo ExitSub1
    

'   获取模板号
    If mRs.State <> 0 Then mRs.Close
    mRs.Open "select vt_id from  VoucherTemplates_Base where vt_templatemode=0 and vt_cardnumber='YXEF9141'", con.ConnectionString, 3, 4
    If Not mRs.BOF And Not mRs.EOF Then
        vt_id_main = CInt(mRs.Fields("vt_id"))
    End If
    
    If mRs.State <> 0 Then mRs.Close
    mRs.Open "select vt_id from  VoucherTemplates_Base where vt_templatemode=0 and vt_cardnumber='YXEF9142'", con.ConnectionString, 3, 4
    If Not mRs.BOF And Not mRs.EOF Then
        vt_id_son = CInt(mRs.Fields("vt_id"))
    End If
     
'    开始事务
    If isTrans = False Then
'        con.BeginTrans
        isTrans = True
    End If
    
    
    
    '修改预生产订单资料
    If Not (IsNull(pubMoid)) Then
        errInfo = Delete("delete from ef_mom_morder where moid in (select moid from ef_mom_order where SoID =" & soId & ")", con)
        errInfo = Delete("delete from ef_mom_moallocate where modid in (select modid from ef_mom_orderdetail where moid in (select moid from ef_mom_order where  SoID =" & soId & "))", con)
        errInfo = Delete("delete from ef_mom_orderdetail where moid in (select moid from ef_mom_order where SoID =" & soId & ")", con)
        errInfo = Delete("delete from ef_mom_order where SoID =" & soId & "", con)
    End If
    
'    计算生产订单id，ccode
    ddID = CInt(GetID("ef_mom_order", "MoId"))
    iSoId = ddID
    ddAutoID = GetID("ef_mom_orderdetail", "autoid")
    zjAutoID = GetID("ef_mom_moallocate", "autoid")
    mCode = CStr(GetCode("ef_mom_order", "MoCode"))
  
        
'   对预生产订单表头写操作
    nvsql = "insert into ef_mom_order([MoId],[id],[MoCode],[ccode],[CreateUser],[cmaker],[CreateDate],[cmakerddate],[cvouchtype],[SoID],[vt_id]) values ("
    nvsql = nvsql & ddID & "," & ddID & ",'" & mCode & "','" & mCode & "','" & nvLogin.cUserName & "','" & nvLogin.cUserName & "','" & nvLogin.CurDate & "','" & nvLogin.CurDate & "','"
    nvsql = nvsql & "YXEF9141','" & soId & "'," & vt_id_main & ")"
    con.Execute nvsql

    
    '获取销售评审单信息
    If mRs.State <> 0 Then mRs.Close
    mRs.Open "select * from EF_V_List_BOMChenage where id =" & soId, con.ConnectionString, 3, 4
    seq = 1
    moaSeq = 10

    Do Until mRs.EOF
        
        '获取工艺路线id、bom id
        nvsql = "select sfc_prouting.PRoutingId,bom_parent.bomid as parentBomid,bom_opcomponent.bomid as componentBomid,sfc_prouting.RountingType,bom_bom.BomType "
        nvsql = nvsql & " from sfc_prouting left outer join sfc_proutingpart on sfc_prouting.PRoutingId = sfc_proutingpart.PRoutingId "
        nvsql = nvsql & " left outer join bas_part on sfc_proutingpart.PartId = bas_part.PartId left outer join Inventory on bas_part.InvCode = Inventory.cInvCode "
        nvsql = nvsql & " left outer join bom_parent on bas_part.partid=bom_parent.parentid left outer join bom_opcomponent on bas_part.partid=bom_opcomponent.componentid"
        nvsql = nvsql & " left outer join bom_bom on bom_parent.bomid = bom_bom.bomid "
        nvsql = nvsql & " where Inventory.cInvCode ='" & mRs.Fields("b_cinvcode") & "'"""
        
        If invRS.State <> 0 Then invRS.Close
        invRS.Open nvsql, con.ConnectionString, 3, 4
        
        If Not invRS.EOF And Not invRS.BOF Then
            If Not IsNull(invRS.Fields("PRoutingId")) Then
                PRoutingId = invRS.Fields("PRoutingId")
            End If
            
            If Not IsNull(invRS.Fields("parentBomid")) Then
                BomId = invRS.Fields("parentBomid")
            End If
            
            If Not IsNull(invRS.Fields("RountingType")) Then
                RountingType = invRS.Fields("RountingType")
            End If
            
            If Not IsNull(invRS.Fields("BomType")) Then
                BomType = invRS.Fields("BomType")
            End If
            
            
        End If
        If invRS.State <> 0 Then invRS.Close
                
                
        If zxRs.State <> 0 Then zxRs.Close
        zxRs.Open "select * from Inventory where cinvcode ='" & mRs.Fields("b_cinvcode") & "'", con.ConnectionString, 3, 4
        If Not zxRs.BOF And Not zxRs.EOF Then
            '生产部门未确定
            nvsql = "insert into ef_mom_orderdetail([vt_id],[autoid],[id],[MoDId],[MoId],[SortSeq],[InvCode],[Qty],[MDeptCode],[MoTypeId],[Status],[OrderType],[OrderCode],[OrderSeq],[WhCode],[qcflag],[MrpQty],[RoutingId],"
            nvsql = nvsql & " [cvouchtype],[bomid],[RoutingType],[BomType]) values(" & vt_id_son & "," & ddAutoID & "," & ddAutoID & "," & ddAutoID & "," & ddID & "," & seq & ",'" & IIf(IsNull(mRs.Fields("b_cinvcode")), "Null", mRs.Fields("b_cinvcode"))
            nvsql = nvsql & "'," & IIf(IsNull(mRs.Fields("b_float2")), "Null", mRs.Fields("b_float2")) & ",'" & IIf(IsNull(zxRs.Fields("cInvDepCode")), "Null", zxRs.Fields("cInvDepCode")) & "'," & "Null,'" & 1 & "','"
'            nvsql = nvsql & 1 & "','" & ordercode & "'," & mRs.Fields("b_int1") & ",'" & zxRs.Fields("cDefWareHouse") & "'," & IIf(zxRs.Fields("bPropertyCheck"), 1, 0) & "," & IIf(IsNull(mRs.Fields("b_float2")), "Null", mRs.Fields("b_float2"))
            nvsql = nvsql & 1 & "'," & "Null" & "," & mRs.Fields("b_int1") & ",'" & zxRs.Fields("cDefWareHouse") & "'," & IIf(zxRs.Fields("bPropertyCheck"), 1, 0) & "," & IIf(IsNull(mRs.Fields("b_float2")), "Null", mRs.Fields("b_float2"))
            
            '获取工艺路线id
            If IsNull(PRoutingId) Then
                nvsql = nvsql & ",Null" & ",'YXEF9142',"
            Else
                nvsql = nvsql & "," & PRoutingId & ",'YXEF9142',"
            End If
            
             '获取物料清单（bom）id
            If IsNull(BomId) Then
                nvsql = nvsql & "Null,"
            Else
                nvsql = nvsql & BomId & ","
            End If
            
             '获取工艺路线类别
            If IsNull(RountingType) Then
                nvsql = nvsql & "Null,"
            Else
                nvsql = nvsql & RountingType & ","
            End If
            
             '获取BOM类别
            If IsNull(BomType) Then
                nvsql = nvsql & "Null )"
            Else
                nvsql = nvsql & BomType & ")"
            End If
            
            con.Execute nvsql
        
            eDate = mRs.Fields("b_datetime1")
            bDate = CStr(CDate(eDate) - IIf(IsNull(zxRs.Fields("iAdvanceDate")), 0, zxRs.Fields("iAdvanceDate")))
            nvsql = "insert into ef_mom_morder ([MoDId],[MoId],[StartDate],[DueDate]) values(" & ddAutoID & "," & ddID & ",'" & bDate & "','" & eDate & "'" & ")"
            con.Execute nvsql
        Else

            nvsql = "insert into ef_mom_orderdetail([autoid],[id],[MoDId],[MoId],[SortSeq],[InvCode],[Qty],[MDeptCode],[RoutingType],[BomType],[MoTypeId],[Status],[OrderType],[OrderCode],[OrderSeq],[cvouchtype]) values("
            nvsql = nvsql & ddAutoID & "," & ddAutoID & "," & ddAutoID & "," & ddID & "," & seq & ",'" & mRs.Fields("b_cinvcode") & "'," & IIf(IsNull(mRs.Fields("b_float2")), "", mRs.Fields("b_float2")) & ",'" & "" & "'," & 0 & "," & 0 & ",'01','" & 1 & "','"
            nvsql = nvsql & 1 & "','" & OrderCode & "'," & mRs.Fields("b_int1") & ",'YXEF9142'" & ")"
            con.Execute nvsql
        End If
                 
        '获取bom展开数据
        nvsql = "select * from EF_T_V_BOMChenageGroup where 销售订单ID='" & mRs.Fields("id") & "' and autoid ='" & mRs.Fields("autoid") & "'"
        If bomRs.State <> 0 Then bomRs.Close
        bomRs.CursorLocation = adUseClient
        bomRs.Open nvsql, con.ConnectionString, 3, 4

        
        Do Until bomRs.EOF
            
            If byRs.State <> 0 Then byRs.Close
            nvsql = "select * from EF_BOM where id='" & bomRs.Fields("销售订单ID") & "' and InvCode='" & bomRs.Fields("材料编码") & "' and autoid=" & bomRs.Fields("autoid")
            byRs.Open nvsql, con.ConnectionString, 3, 4
            If Not byRs.EOF Then
            
                moa_AllocateId = zjAutoID
                moa_MoDId = ddAutoID
                moa_SortSeq = moaSeq
                moaSeq = moaSeq + 10
                moa_OpSeq = byRs.Fields("OpSeq")
                moa_ComponentId = 0
                moa_FVFlag = byRs.Fields("FVQty")
                moa_BaseQtyN = byRs.Fields("BaseQtyN")
                moa_BaseQtyD = byRs.Fields("BaseQtyD")
                moa_ParentScrap = byRs.Fields("ParentScrap")
                moa_CompScrap = byRs.Fields("CompScrap")
                moa_Qty = bomRs.Fields("总需求量")
                moa_WhCode = byRs.Fields("WhCode")
                moa_WIPType = byRs.Fields("WIPType")
                If LCase(byRs.Fields("ByproductFlag")) = LCase("") Then
                    moa_ByproductFlag = 1
                Else
                    moa_ByproductFlag = 0
                End If
                moa_Offset = byRs.Fields("Offset")
                moa_InvCode = byRs.Fields("InvCode")
                moa_OpComponentId = byRs.Fields("OpComponentId")
                
                moa_IssQty = 0                              '已领量
                moa_DeclaredQty = 0                         'DeclaredQty 报检数量
                moa_StartDemDate = CStr(bDate)              '开始需求日期
                moa_EndDemDate = CStr(eDate)                '结束需求日期"
                moa_LotNo = ""                              '批号
                moa_QcFlag = 1                              '检查否
                moa_Free1 = ""
                moa_Free2 = ""
                moa_Free3 = ""
                moa_Free4 = ""
                moa_Free5 = ""
                moa_Free6 = ""
                moa_Free7 = ""
                moa_Free8 = ""
                moa_Free9 = ""
                moa_Free10 = ""
                moa_Define22 = ""
                moa_Define23 = ""
                moa_Define24 = ""
                moa_Define25 = ""
                moa_Define26 = 0
                moa_Define27 = 0
                moa_Define28 = ""
                moa_Define29 = ""
                moa_Define30 = ""
                moa_Define31 = ""
                moa_Define32 = ""
                moa_Define33 = ""
                moa_Define34 = 0
                moa_Define35 = 0
                moa_Define36 = ""
                moa_Define37 = ""
                moa_AuxUnitCode = ""                        'AuxUnitCode 辅助计量单位
                moa_ChangeRate = 0                          'ChangeRate 换算率
                moa_AuxBaseQtyN = 0                         'AuxBaseQtyN辅助基本用量
                moa_AuxQty = 0                              'AuxQty 应领辅助量
                moa_ReplenishQty = 0                        'ReplenishQty 补料量
                moa_Remark = ""                             '备注
                moa_TransQty = 0                            '已调拨量
                moa_SoType = 5                              '需求跟踪方式 0无来源 1销售订单行 3出口订单行 4需求分类 5 销售订单 6出口订单
                moa_SoCode = ""                             '需求跟踪号
                moa_SoSeq = 0                               '需求跟踪行号
                moa_SoDId = ""                              '需求跟踪DId
                moa_DemandCode = ""                         '需求分类单号
                
                nvsql = "INSERT INTO ef_mom_moallocate(AllocateId, MoDId, SortSeq, OpSeq, ComponentId, FVFlag, BaseQtyN, BaseQtyD, ParentScrap, CompScrap, Qty , IssQty, DeclaredQty, StartDemDate,"
                nvsql = nvsql & "EndDemDate, WhCode, LotNo, WIPType, ByproductFlag, ProductType, QcFlag, Offset, InvCode , Free1, Free2, Free3, Free4, Free5, Free6, Free7, Free8, Free9, Free10 , Define22,"
                nvsql = nvsql & "Define23, Define24, Define25, Define26, Define27, Define28, Define29, Define30, Define31, Define32, Define33, Define34, Define35, Define36, Define37 , OpComponentId,"
                nvsql = nvsql & "AuxUnitCode,ChangeRate,AuxBaseQtyN,AuxQty,ReplenishQty,Remark,TransQty,SoType,SoCode,SoSeq,SoDId,DemandCode,autoid,id) values ( "
                nvsql = nvsql & moa_AllocateId & "," & moa_MoDId & "," & moa_SortSeq & ",'" & moa_OpSeq & "'," & moa_ComponentId & "," & moa_FVFlag & "," & moa_BaseQtyN & "," & moa_BaseQtyD & ","
                nvsql = nvsql & moa_ParentScrap & "," & moa_CompScrap & "," & moa_Qty & "," & moa_IssQty & "," & moa_DeclaredQty & ",'" & moa_StartDemDate & "','" & moa_EndDemDate & "','" & moa_WhCode & "','"
                nvsql = nvsql & moa_LotNo & "'," & moa_WIPType & "," & moa_ByproductFlag & "," & moa_ProductType & "," & moa_QcFlag & ",'" & moa_Offset & "','" & moa_InvCode & "','" & moa_Free1 & "','"
                nvsql = nvsql & moa_Free2 & "','" & moa_Free3 & "','" & moa_Free4 & "','" & moa_Free5 & "','" & moa_Free6 & "','" & moa_Free7 & "','" & moa_Free8 & "','" & moa_Free9 & "','" & moa_Free10 & "','"
                nvsql = nvsql & moa_Define22 & "','" & moa_Define23 & "','" & moa_Define24 & "','" & moa_Define25 & "'," & moa_Define26 & "," & moa_Define27 & ",'" & moa_Define28 & "','" & moa_Define29 & "','"
                nvsql = nvsql & moa_Define30 & "','" & moa_Define31 & "','" & moa_Define32 & "','" & moa_Define33 & "'," & moa_Define34 & "," & moa_Define35 & ",'" & moa_Define36 & "','" & moa_Define37 & "'," & moa_OpComponentId & ",'"
                nvsql = nvsql & moa_AuxUnitCode & "'," & moa_ChangeRate & "," & moa_AuxBaseQtyN & "," & moa_AuxQty & "," & moa_ReplenishQty & ",'" & moa_Remark & "'," & moa_TransQty & "," & moa_SoType & ",'" & moa_SoCode & "',"
                nvsql = nvsql & moa_SoSeq & ",'" & moa_SoDId & "','" & moa_DemandCode & "'," & moa_AllocateId & "," & moa_MoDId & ")"
                
                con.Execute nvsql
            End If
            
 
            bomRs.MoveNext
            zjAutoID = zjAutoID + 1
        Loop

        mRs.MoveNext
        seq = seq + 1
        ddAutoID = ddAutoID + 1
    Loop
    
    
    nvsql = "update EF_SellAccreditation set moid = " & iSoId & "Where SoID = '" & OrderCode & "'"
    con.Execute nvsql
    

    If isTrans Then
'        con.CommitTrans
        isTrans = False
    End If
    
    Exit Function
ExitSub1:

    If isTrans Then
'        con.RollbackTrans
        isTrans = False
    End If
    If nvRs.State <> 0 Then nvRs.Close
    If bomRs.State <> 0 Then bomRs.Close
    If mRs.State <> 0 Then mRs.Close
    If zxRs.State <> 0 Then zxRs.Close
    If rds.State <> 0 Then rds.Close
    MsgBox "生成预生产订单失败，错误：" & Err.Description
    
End Function

Public Function SetNVBom(mGrid As SuperGrid)
    
    Dim momRs As New ADODB.Recordset
        
    Dim strSql As String
    Dim allBool As Boolean      '是否全部替换标志
    Dim i As Integer            'For 循环变量
    Dim j As Integer            '
    Dim qty As Double           '订单数量
    Dim bQty As Double          '原物料保留量
    
    Dim sql As String
    
'    For i = 1 To mGrid.rows - 1
'        If CDbl(mGrid.TextMatrix(i, 14)) <= 0 Then
'            allBool = True
'        End If
'    Next i
'
'    If Not allBool Then
'        bQty = CDbl(mGrid.TextMatrix(1, 14))
'        If mGrid.rows = 2 Then
'            bQty = CDbl(mGrid.TextMatrix(1, 14))
'            j = 1
'        End If
'        For i = 1 To mGrid.rows - 1
'            If bQty > CDbl(mGrid.TextMatrix(i, 14)) Then
'                bQty = CDbl(mGrid.TextMatrix(i, 14))
'                j = i
'            End If
'        Next i
'
'    End If
           
    For i = 1 To mGrid.rows - 1
        
        sql = "select * from ef_bom where id=" & mGrid.TextMatrix(i, 18) & " and autoid=" & mGrid.TextMatrix(i, 19) & " and minvcode='" & mGrid.TextMatrix(i, 3) & "' and cinvcode ='" & mGrid.TextMatrix(i, 5) & "'"
       
        If nvRs.State <> 0 Then nvRs.Close
        nvRs.CursorLocation = adUseClient
        nvRs.Open sql, con.ConnectionString, 3, 4
        
        If Not nvRs.EOF And Not nvRs.BOF Then
            
            If CDbl(mGrid.TextMatrix(i, 14)) <= 0 Then
                con.Execute "delete ef_bom where invcode = '" & mGrid.TextMatrix(1, 5) + "' and CCode ='" & mGrid.TextMatrix(1, 1) & "'"
            Else
                strSql = "update ef_bom set BaseQtyD = 1, BaseQtyN = " & CDbl(nvRs.Fields("mQty")) / CDbl(mGrid.TextMatrix(1, 15)) & " where id=" & mGrid.TextMatrix(i, 18) & " and autoid=" & mGrid.TextMatrix(i, 19) & " and minvcode='" & mGrid.TextMatrix(i, 3) & "' and cinvcode ='" & mGrid.TextMatrix(i, 5) & "'"
                con.Execute strSql
            End If
        
            bom_cInvCode = mGrid.TextMatrix(i, 11)
            bom_CCode = nvRs.Fields("ID")
            bom_OpComponentId = 9999
            bom_mInvCode = nvRs.Fields("minvcode")
            bom_OpSeq = nvRs.Fields("OpSeq")
            bom_CompId = nvRs.Fields("CompId")
            bom_UnitId = nvRs.Fields("UnitId")
            bom_BaseQtyD = 1                                                        '基础用量/
            bom_BaseQtyN = CDbl(mGrid.TextMatrix(i, 15)) / nvRs.Fields("mQty")      '基本用量
            bom_mQty = bom_BaseQtyN / bom_BaseQtyD * nvRs.Fields("mQty")
            bom_ParentScrap = nvRs.Fields("ParentScrap")
            bom_CompScrap = nvRs.Fields("CompScrap")
            If LCase(nvRs.Fields("FVQty")) = "true" Then
                bom_FVQty = 1
            ElseIf LCase(nvRs.Fields("FVQty")) = "false" Then
                bom_FVQty = 0
            End If
            bom_Cqty = nvRs.Fields("Cqty")
            bom_Cqty1 = nvRs.Fields("Cqty1")
            bom_UseQty = nvRs.Fields("UseQty")
            bom_Offset = nvRs.Fields("Offset")
            bom_WIPtype = nvRs.Fields("WIPtype")
            bom_WhCode = nvRs.Fields("WhCode")
            bom_InvCode = mGrid.TextMatrix(i, 11)
            bom_Free1 = nvRs.Fields("Free1")
            bom_Free2 = nvRs.Fields("Free2")
            bom_Free3 = nvRs.Fields("Free3")
            bom_Free4 = nvRs.Fields("Free4")
            bom_Free5 = nvRs.Fields("Free5")
            bom_Free6 = nvRs.Fields("Free6")
            bom_Free7 = nvRs.Fields("Free7")
            bom_Free8 = nvRs.Fields("Free8")
            bom_Free9 = nvRs.Fields("Free9")
            bom_Free10 = nvRs.Fields("Free10")
            bom_Dept = nvRs.Fields("Dept")
            bom_DepName = nvRs.Fields("DepName")
            If LCase(nvRs.Fields("ByproductFlag")) = "true" Then
                bom_ByproductFlag = 1
            ElseIf LCase(nvRs.Fields("ByproductFlag")) = "false" Then
                bom_ByproductFlag = 1
            End If
            
        '    bom_AccuCostFlag = GetBodyItemValue1(domtemp, "AccuCostFlag ", j)
        '    bom_SubFlag = GetBodyItemValue1(domtemp, "SubFlag", j)
        '    bom_BomType = GetBodyItemValue1(domtemp, "BomType", j)
        '    bom_iGrade = GetBodyItemValue1(domtemp, "iGrade", j)
        '    bom_DemDate = GetBodyItemValue1(domtemp, "DemDate", j)
        '    bom_AuxUnitCode = GetBodyItemValue1(domtemp, "AuxUnitCode", j)
        '    bom_ChangeRate = GetBodyItemValue1(domtemp, "ChangeRate", j)
        '    bom_AuxBaseQtyN = GetBodyItemValue1(domtemp, "AuxBaseQtyN", j)
        '    bom_AuxCqty = GetBodyItemValue1(domtemp, "AuxCqty", j)
        '    bom_AuxUseQty = GetBodyItemValue1(domtemp, "AuxUseQty", j)
        '    bom_AuxUnitName = GetBodyItemValue1(domtemp, "AuxUnitName", j)
        '    bom_Define1 = GetBodyItemValue1(domtemp, "Define1", j)
        '    bom_Define2 = GetBodyItemValue1(domtemp, "Define2", j)
        '    bom_Define3 = GetBodyItemValue1(domtemp, "Define3", j)
        '    bom_Define4 = GetBodyItemValue1(domtemp, "Define4", j)
        '    bom_Define5 = GetBodyItemValue1(domtemp, "Define5", j)
        '    bom_Define6 = GetBodyItemValue1(domtemp, "Define6", j)
        '    bom_Define7 = GetBodyItemValue1(domtemp, "Define7", j)
        '    bom_Define8 = GetBodyItemValue1(domtemp, "Define8", j)
        '    bom_Define9 = GetBodyItemValue1(domtemp, "Define9", j)
        '    bom_Define10 = GetBodyItemValue1(domtemp, "Define10", j)
        '    bom_Define11 = GetBodyItemValue1(domtemp, "Define11", j)
        '    bom_Define12 = GetBodyItemValue1(domtemp, "Define12", j)
        '    bom_Define13 = GetBodyItemValue1(domtemp, "Define13", j)
        '    bom_Define14 = GetBodyItemValue1(domtemp, "Define14", j)
        '    bom_Define15 = GetBodyItemValue1(domtemp, "Define15", j)
        '    bom_Define16 = GetBodyItemValue1(domtemp, "Define16", j)
        '    bom_Define22 = GetBodyItemValue1(domtemp, "Define22", j)
        '    bom_Define23 = GetBodyItemValue1(domtemp, "Define23", j)
        '    bom_Define24 = GetBodyItemValue1(domtemp, "Define24", j)
        '    bom_Define25 = GetBodyItemValue1(domtemp, "Define25", j)
        '    bom_Define26 = GetBodyItemValue1(domtemp, "Define26", j)
        '    bom_Define27 = GetBodyItemValue1(domtemp, "Define27", j)
        '    bom_Define28 = GetBodyItemValue1(domtemp, "Define28", j)
        '    bom_Define29 = GetBodyItemValue1(domtemp, "Define29", j)
        '    bom_Define30 = GetBodyItemValue1(domtemp, "Define30", j)
        '    bom_Define31 = GetBodyItemValue1(domtemp, "Define31", j)
        '    bom_Define32 = GetBodyItemValue1(domtemp, "Define32", j)
        '    bom_Define33 = GetBodyItemValue1(domtemp, "Define33", j)
        '    bom_Define34 = GetBodyItemValue1(domtemp, "Define34", j)
        '    bom_Define35 = GetBodyItemValue1(domtemp, "Define35", j)
        '    bom_Define36 = GetBodyItemValue1(domtemp, "Define36", j)
        '    bom_Define37 = GetBodyItemValue1(domtemp, "Define37", j)
        
            sql = "insert into ef_bom(OpComponentId,OpSeq,CompId,UnitId,BaseQtyN,BaseQtyD,ParentScrap,CompScrap,FVQty,Cqty,Cqty1,UseQty,Offset,WIPtype,WhCode,InvCode,Free1,Free2,Free3,Free4,Free5,Free6,Free7,"
            sql = sql & "Free8,Free9,Free10,Dept,DepName,ByproductFlag,AccuCostFlag ,SubFlag,BomType,iGrade,DemDate,AuxUnitCode,ChangeRate,AuxBaseQtyN,AuxCqty,AuxUseQty,AuxUnitName,Define1,Define2,"
            sql = sql & "Define3,Define4,Define5,Define6,Define7,Define8,Define9,Define10,Define11,Define12,Define13,Define14,Define15,Define16,Define22,Define23,Define24,Define25,Define26,Define27,"
            sql = sql & "Define28,Define29,Define30,Define31,Define32,Define33,Define34,Define35,Define36,Define37,ID,cInvCode,autoid,mInvcode,mQty,guid) values ('"
            sql = sql & bom_OpComponentId & "','" & bom_OpSeq & "','" & bom_CompId & "','" & bom_UnitId & "','" & bom_BaseQtyN & "','" & bom_BaseQtyD & "','" & bom_ParentScrap & "','" & bom_CompScrap & "',"
            sql = sql & bom_FVQty & "," & bom_Cqty & "," & bom_Cqty1 & "," & bom_UseQty & "," & bom_Offset & "," & bom_WIPtype & ",'" & bom_WhCode & "','" & bom_InvCode & "','" & bom_Free1 & "','" & bom_Free2 & "','"
            sql = sql & bom_Free3 & "','" & bom_Free4 & "','" & bom_Free5 & "','" & bom_Free6 & "','" & bom_Free7 & "','" & bom_Free8 & "','" & bom_Free9 & "','" & bom_Free10 & "','" & bom_Dept & "','" & bom_DepName & "',"
            sql = sql & bom_ByproductFlag & "," & bom_AccuCostFlag & "," & bom_SubFlag & "," & bom_BomType & "," & bom_iGrade & ",'" & bom_DemDate & "','" & bom_AuxUnitCode & "','" & bom_ChangeRate & "'," & bom_AuxBaseQtyN & ",'"
            sql = sql & bom_AuxCqty & "'," & bom_AuxUseQty & ",'" & bom_AuxUnitName & "','" & bom_Define1 & "','" & bom_Define2 & "','" & bom_Define3 & "','" & bom_Define4 & "','" & bom_Define5 & "','" & bom_Define6 & "','"
            sql = sql & bom_Define7 & "','" & bom_Define8 & "','" & bom_Define9 & "','" & bom_Define10 & "','" & bom_Define11 & "','" & bom_Define12 & "','" & bom_Define13 & "','" & bom_Define14 & "','" & bom_Define15 & "','"
            sql = sql & bom_Define16 & "','" & bom_Define22 & "','" & bom_Define23 & "','" & bom_Define24 & "','" & bom_Define25 & "','" & bom_Define26 & "','" & bom_Define27 & "','" & bom_Define28 & "','" & bom_Define29 & "','"
            sql = sql & bom_Define30 & "','" & bom_Define31 & "','" & bom_Define32 & "','" & bom_Define33 & "','" & bom_Define34 & "','" & bom_Define35 & "','" & bom_Define36 & "','" & bom_Define37 & "','" & bom_CCode & "','"
            sql = sql & bom_cInvCode & "'," & CInt(nvRs.Fields("autoid")) & ",'" & bom_mInvCode & "'," & bom_mQty & ",newid()" & ")"
        
            con.Execute sql
        End If
    Next i
     
'    If allBool Then
'        con.Execute "delete ef_bom where invcode = '" & mGrid.TextMatrix(1, 5) + "' and CCode ='" & mGrid.TextMatrix(1, 1) & "'"
'    Else
'        strSql = "update ef_bom set BaseQtyD = 1, BaseQtyN = " & CDbl(mGrid.TextMatrix(j, 14)) / "& mGrid.TextMatrix(1, 15) &" & "where invcode = '" & mGrid.TextMatrix(1, 5) + "' and CCode ='" & mGrid.TextMatrix(1, 1) & "'"
'        con.Execute strSql
'    End If
'
'    For i = 1 To mGrid.rows - 1
'        bom_cInvCode = mGrid.TextMatrix(i, 11)
'        bom_CCode = nvRs.Fields("ID")
'        bom_OpComponentId = 9999
'        bom_mInvCode = nvRs.Fields("minvcode")
'        bom_OpSeq = nvRs.Fields("OpSeq")
'        bom_CompId = nvRs.Fields("CompId")
'        bom_UnitId = nvRs.Fields("UnitId")
'        bom_BaseQtyD = 1                                                '基础用量/
'        bom_BaseQtyN = CDbl(mGrid.TextMatrix(i, 15)) / nvRs.Fields("mQty")  '基本用量
'        bom_mQty = CDbl(mGrid.TextMatrix(i, 15)) * nvRs.Fields("mQty")
'        bom_ParentScrap = nvRs.Fields("ParentScrap")
'        bom_CompScrap = nvRs.Fields("CompScrap")
'        If LCase(nvRs.Fields("FVQty")) = "true" Then
'            bom_FVQty = 1
'        ElseIf LCase(nvRs.Fields("FVQty")) = "false" Then
'            bom_FVQty = 0
'        End If
'        bom_Cqty = nvRs.Fields("Cqty")
'        bom_Cqty1 = nvRs.Fields("Cqty1")
'        bom_UseQty = nvRs.Fields("UseQty")
'        bom_Offset = nvRs.Fields("Offset")
'        bom_WIPtype = nvRs.Fields("WIPtype")
'        bom_WhCode = nvRs.Fields("WhCode")
'        bom_InvCode = mGrid.TextMatrix(i, 11)
'        bom_Free1 = nvRs.Fields("Free1")
'        bom_Free2 = nvRs.Fields("Free2")
'        bom_Free3 = nvRs.Fields("Free3")
'        bom_Free4 = nvRs.Fields("Free4")
'        bom_Free5 = nvRs.Fields("Free5")
'        bom_Free6 = nvRs.Fields("Free6")
'        bom_Free7 = nvRs.Fields("Free7")
'        bom_Free8 = nvRs.Fields("Free8")
'        bom_Free9 = nvRs.Fields("Free9")
'        bom_Free10 = nvRs.Fields("Free10")
'        bom_Dept = nvRs.Fields("Dept")
'        bom_DepName = nvRs.Fields("DepName")
'        If LCase(nvRs.Fields("ByproductFlag")) = "true" Then
'            bom_ByproductFlag = 1
'        ElseIf LCase(nvRs.Fields("ByproductFlag")) = "false" Then
'            bom_ByproductFlag = 1
'        End If
'
'    '    bom_AccuCostFlag = GetBodyItemValue1(domtemp, "AccuCostFlag ", j)
'    '    bom_SubFlag = GetBodyItemValue1(domtemp, "SubFlag", j)
'    '    bom_BomType = GetBodyItemValue1(domtemp, "BomType", j)
'    '    bom_iGrade = GetBodyItemValue1(domtemp, "iGrade", j)
'    '    bom_DemDate = GetBodyItemValue1(domtemp, "DemDate", j)
'    '    bom_AuxUnitCode = GetBodyItemValue1(domtemp, "AuxUnitCode", j)
'    '    bom_ChangeRate = GetBodyItemValue1(domtemp, "ChangeRate", j)
'    '    bom_AuxBaseQtyN = GetBodyItemValue1(domtemp, "AuxBaseQtyN", j)
'    '    bom_AuxCqty = GetBodyItemValue1(domtemp, "AuxCqty", j)
'    '    bom_AuxUseQty = GetBodyItemValue1(domtemp, "AuxUseQty", j)
'    '    bom_AuxUnitName = GetBodyItemValue1(domtemp, "AuxUnitName", j)
'    '    bom_Define1 = GetBodyItemValue1(domtemp, "Define1", j)
'    '    bom_Define2 = GetBodyItemValue1(domtemp, "Define2", j)
'    '    bom_Define3 = GetBodyItemValue1(domtemp, "Define3", j)
'    '    bom_Define4 = GetBodyItemValue1(domtemp, "Define4", j)
'    '    bom_Define5 = GetBodyItemValue1(domtemp, "Define5", j)
'    '    bom_Define6 = GetBodyItemValue1(domtemp, "Define6", j)
'    '    bom_Define7 = GetBodyItemValue1(domtemp, "Define7", j)
'    '    bom_Define8 = GetBodyItemValue1(domtemp, "Define8", j)
'    '    bom_Define9 = GetBodyItemValue1(domtemp, "Define9", j)
'    '    bom_Define10 = GetBodyItemValue1(domtemp, "Define10", j)
'    '    bom_Define11 = GetBodyItemValue1(domtemp, "Define11", j)
'    '    bom_Define12 = GetBodyItemValue1(domtemp, "Define12", j)
'    '    bom_Define13 = GetBodyItemValue1(domtemp, "Define13", j)
'    '    bom_Define14 = GetBodyItemValue1(domtemp, "Define14", j)
'    '    bom_Define15 = GetBodyItemValue1(domtemp, "Define15", j)
'    '    bom_Define16 = GetBodyItemValue1(domtemp, "Define16", j)
'    '    bom_Define22 = GetBodyItemValue1(domtemp, "Define22", j)
'    '    bom_Define23 = GetBodyItemValue1(domtemp, "Define23", j)
'    '    bom_Define24 = GetBodyItemValue1(domtemp, "Define24", j)
'    '    bom_Define25 = GetBodyItemValue1(domtemp, "Define25", j)
'    '    bom_Define26 = GetBodyItemValue1(domtemp, "Define26", j)
'    '    bom_Define27 = GetBodyItemValue1(domtemp, "Define27", j)
'    '    bom_Define28 = GetBodyItemValue1(domtemp, "Define28", j)
'    '    bom_Define29 = GetBodyItemValue1(domtemp, "Define29", j)
'    '    bom_Define30 = GetBodyItemValue1(domtemp, "Define30", j)
'    '    bom_Define31 = GetBodyItemValue1(domtemp, "Define31", j)
'    '    bom_Define32 = GetBodyItemValue1(domtemp, "Define32", j)
'    '    bom_Define33 = GetBodyItemValue1(domtemp, "Define33", j)
'    '    bom_Define34 = GetBodyItemValue1(domtemp, "Define34", j)
'    '    bom_Define35 = GetBodyItemValue1(domtemp, "Define35", j)
'    '    bom_Define36 = GetBodyItemValue1(domtemp, "Define36", j)
'    '    bom_Define37 = GetBodyItemValue1(domtemp, "Define37", j)
'
'        sql = "insert into ef_bom(OpComponentId,OpSeq,CompId,UnitId,BaseQtyN,BaseQtyD,ParentScrap,CompScrap,FVQty,Cqty,Cqty1,UseQty,Offset,WIPtype,WhCode,InvCode,Free1,Free2,Free3,Free4,Free5,Free6,Free7,"
'        sql = sql & "Free8,Free9,Free10,Dept,DepName,ByproductFlag,AccuCostFlag ,SubFlag,BomType,iGrade,DemDate,AuxUnitCode,ChangeRate,AuxBaseQtyN,AuxCqty,AuxUseQty,AuxUnitName,Define1,Define2,"
'        sql = sql & "Define3,Define4,Define5,Define6,Define7,Define8,Define9,Define10,Define11,Define12,Define13,Define14,Define15,Define16,Define22,Define23,Define24,Define25,Define26,Define27,"
'        sql = sql & "Define28,Define29,Define30,Define31,Define32,Define33,Define34,Define35,Define36,Define37,ID,cInvCode,autoid,mInvcode,mQty,guid) values ('"
'        sql = sql & bom_OpComponentId & "','" & bom_OpSeq & "','" & bom_CompId & "','" & bom_UnitId & "','" & bom_BaseQtyN & "','" & bom_BaseQtyD & "','" & bom_ParentScrap & "','" & bom_CompScrap & "',"
'        sql = sql & bom_FVQty & "," & bom_Cqty & "," & bom_Cqty1 & "," & bom_UseQty & "," & bom_Offset & "," & bom_WIPtype & ",'" & bom_WhCode & "','" & bom_InvCode & "','" & bom_Free1 & "','" & bom_Free2 & "','"
'        sql = sql & bom_Free3 & "','" & bom_Free4 & "','" & bom_Free5 & "','" & bom_Free6 & "','" & bom_Free7 & "','" & bom_Free8 & "','" & bom_Free9 & "','" & bom_Free10 & "','" & bom_Dept & "','" & bom_DepName & "',"
'        sql = sql & bom_ByproductFlag & "," & bom_AccuCostFlag & "," & bom_SubFlag & "," & bom_BomType & "," & bom_iGrade & ",'" & bom_DemDate & "','" & bom_AuxUnitCode & "','" & bom_ChangeRate & "'," & bom_AuxBaseQtyN & ",'"
'        sql = sql & bom_AuxCqty & "'," & bom_AuxUseQty & ",'" & bom_AuxUnitName & "','" & bom_Define1 & "','" & bom_Define2 & "','" & bom_Define3 & "','" & bom_Define4 & "','" & bom_Define5 & "','" & bom_Define6 & "','"
'        sql = sql & bom_Define7 & "','" & bom_Define8 & "','" & bom_Define9 & "','" & bom_Define10 & "','" & bom_Define11 & "','" & bom_Define12 & "','" & bom_Define13 & "','" & bom_Define14 & "','" & bom_Define15 & "','"
'        sql = sql & bom_Define16 & "','" & bom_Define22 & "','" & bom_Define23 & "','" & bom_Define24 & "','" & bom_Define25 & "','" & bom_Define26 & "','" & bom_Define27 & "','" & bom_Define28 & "','" & bom_Define29 & "','"
'        sql = sql & bom_Define30 & "','" & bom_Define31 & "','" & bom_Define32 & "','" & bom_Define33 & "','" & bom_Define34 & "','" & bom_Define35 & "','" & bom_Define36 & "','" & bom_Define37 & "','" & bom_CCode & "','"
'        sql = sql & bom_cInvCode & "'," & CInt(nvRs.Fields("autoid")) & ",'" & bom_mInvCode & "'," & bom_mQty & ",newid()" & ")"
'
'        con.Execute sql

'    Next i

End Function


Public Sub makeorrder(i_Soid As Long, i_SoDid As Long, m_Cinv As String, CreateUser As String, CreateDate As String)
    '生成生产订单API接口
    Dim Import As Object   '导入
    Dim xmldoc As New DOMDocument
    Dim Errstr As String
    Dim pi As IXMLDOMProcessingInstruction
    
    On Error GoTo makeorrderErr
    
    If Import Is Nothing Then
    Set Import = CreateObject("MRPAPI.API_interface")
    End If
'    Set xmldoc = GetProdectXml(m_MOid)
'    Set xmldoc = GetBomToXml(62, 126, "999-PCBA", "F010-998000-001", "demo", "2010-08-03")
    
    Set xmldoc = GetBomToXml(i_Soid, i_SoDid, m_Cinv, CreateUser, CreateDate)
    Set pi = xmldoc.createProcessingInstruction("xml", "version='1.0'")
    Call xmldoc.insertBefore(pi, xmldoc.childNodes(0))
    
    
    xmldoc.Save "c:\GetProdectXml" & Date & "-" & Hour(Time) & "-" & Minute(Time) & "-" & Second(Time) & ".xml"
    
    cls_Public.WrtDBlog con, nvLogin.cUserId, "Interface_fxm", "开始 Import.init nvLogin！"
    Import.init nvLogin
    
    cls_Public.WrtDBlog con, nvLogin.cUserId, "Interface_fxm", "开始 Import.Add！"
    
    Import.Add 2, xmldoc, Errstr
    
    cls_Public.WrtDBlog con, nvLogin.cUserId, "Interface_fxm", "Import.Add  Errstr1： " & Errstr
    Exit Sub
makeorrderErr:
    cls_Public.WrtDBlog con, nvLogin.cUserId, "Interface_fxm", "Import.Add  Errstr2： " & Err.Description
End Sub

Private Function GetBomToXml(i_Soid As Long, i_SoDid As Long, m_Cinv As String, CreateUser As String, CreateDate As String) As DOMDocument
    Dim MOrder As IXMLDOMNode
    Dim Order As IXMLDOMNode
    Dim OrderDetail As IXMLDOMNode
    Dim MOrderDetail As IXMLDOMNode
    Dim Allocate As IXMLDOMNode
    Dim Dom As New DOMDocument
    Dim domxml As New DOMDocument
    Dim i As Long
    
    Dim AllocateId As Long
    Dim AllRow As Integer

    Dim OrderRs As New ADODB.Recordset  '订单主
    Dim MrderRs As New ADODB.Recordset  '订单表体
    Dim ArderRs As New ADODB.Recordset  '订单子件

    Dim OrderSql As String

    On Error GoTo ErrGetBomToXml
    
    AllocateId = 1
    AllRow = 10

    Set MOrder = domxml.createElement("MOrder").cloneNode(True)
     '增加根节点 MOrder
    Dom.appendChild MOrder

    '创建生产订单Order节点
    Set Order = GET_space_Orderxml

    '节点赋值 Moid
    SET_IXMLDOMNode_text Order, "MoId", "1"
    '节点赋值 MoCode
    SET_IXMLDOMNode_text Order, "MoCode", "0000000001"
    '节点赋值 CreateDate
    SET_IXMLDOMNode_text Order, "CreateDate", CreateDate
    '节点赋值 CreateUser
    SET_IXMLDOMNode_text Order, "CreateUser", CreateUser
    '节点赋值 Define14
    SET_IXMLDOMNode_text Order, "Define14", CStr(i_Soid) & CStr(i_SoDid) & m_Cinv

    '节点追加
    Dom.selectSingleNode("//MOrder").appendChild Order '增加MOrder节点的下已节点 MOrder

    OrderSql = "select * from ref_v_mom_orderdetail where id=" & i_Soid & " and autoid=" & i_SoDid & " and invcode='" & m_Cinv & "'"
    MrderRs.Open OrderSql, con.ConnectionString, 3, 4

    Do While Not MrderRs.EOF
        '创建生产订单OrderDetail节点
        Set OrderDetail = GET_space_OrderDetailxml
        '节点赋值
        For i = 0 To MrderRs.Fields.count - 1
            SET_IXMLDOMNode_text OrderDetail, MrderRs.Fields(i).Name, IIf(IsNull(MrderRs.Fields(i).Value), "", MrderRs.Fields(i).Value)
        Next i
        '节点追加
        Dom.selectSingleNode("//MOrder").appendChild OrderDetail '增加MOrder节点的下已节点 OrderDetail


        '创建生产订单MOrderDetail节点
        Set MOrderDetail = GET_space_MOrderDetailxml
        '节点赋值
        For i = 0 To MrderRs.Fields.count - 1
            SET_IXMLDOMNode_text MOrderDetail, MrderRs.Fields(i).Name, IIf(IsNull(MrderRs.Fields(i).Value), "", MrderRs.Fields(i).Value)
        Next i
        '节点追加
        Dom.selectSingleNode("//MOrder").appendChild MOrderDetail '增加MOrder节点的下已节点 MOrderDetail


        OrderSql = "select * from ref_v_mom_moallocate where id=" & i_Soid & " and autoid=" & i_SoDid & " and minvcode='" & m_Cinv & "'"
        If ArderRs.State <> 0 Then ArderRs.Close
        ArderRs.Open OrderSql, con.ConnectionString, 3, 4

        Do While Not ArderRs.EOF
            '创建生产订单Allocate节点
            Set Allocate = GET_space_Allocatexml
            '节点赋值
            For i = 0 To ArderRs.Fields.count - 1
                If ArderRs.Fields(i).Name = "AllocateId" Then
                        SET_IXMLDOMNode_text Allocate, ArderRs.Fields(i).Name, CStr(AllocateId)
                ElseIf ArderRs.Fields(i).Name = "SortSeq" Then
                    SET_IXMLDOMNode_text Allocate, ArderRs.Fields(i).Name, CStr(AllRow)
                Else
                    SET_IXMLDOMNode_text Allocate, ArderRs.Fields(i).Name, IIf(IsNull(ArderRs.Fields(i).Value), "", ArderRs.Fields(i).Value)
                End If
            Next i
            '节点追加
            Dom.selectSingleNode("//MOrder").appendChild Allocate '增加MOrder节点的下已节点 Allo

            AllocateId = AllocateId + 1
            AllRow = AllRow + 10
            ArderRs.MoveNext
        Loop
        MrderRs.MoveNext
    Loop
    Set GetBomToXml = Dom.cloneNode(True)

    If OrderRs.State <> 0 Then OrderRs.Close
    If MrderRs.State <> 0 Then MrderRs.Close
    If ArderRs.State <> 0 Then ArderRs.Close
    Exit Function

ErrGetBomToXml:
    If OrderRs.State <> 0 Then OrderRs.Close
    If MrderRs.State <> 0 Then MrderRs.Close
    If ArderRs.State <> 0 Then ArderRs.Close
    MsgBox Err.Description
End Function



Private Function GetProdectXml(m_MOid As Long) As DOMDocument
    Dim MOrder As IXMLDOMNode
    Dim Order As IXMLDOMNode
    Dim OrderDetail As IXMLDOMNode
    Dim MOrderDetail As IXMLDOMNode
    Dim Allocate As IXMLDOMNode
    Dim Dom As New DOMDocument
    Dim domxml As New DOMDocument
    Dim i As Long
    
    Dim OrderRs As New ADODB.Recordset  '订单主
    Dim MrderRs As New ADODB.Recordset  '订单表体
    Dim ArderRs As New ADODB.Recordset  '订单子件
    
    Dim OrderSql As String
    
    On Error GoTo ErrGetProdectXml
    

    
    Set MOrder = domxml.createElement("MOrder").cloneNode(True)
     '增加根节点 MOrder
    Dom.appendChild MOrder
    
    OrderSql = "select OrderCode,MoId,MoCode,CreateDate,CreateUser,Define1 as Define1 ,Define2 as Define2,Define3 as Define3,Define4 as Define4,Define5 as Define5,Define6 as Define6,Define7  as Define7,Define8 as Define8,Define9 as Define9,Define10 as Define10,Define11 as Define11,Define12 as Define12,Define13 as Define13,OrderCode as Define14,Define15  as Define15,Define16 as Define16"
    OrderSql = OrderSql & " from ef_mom_order where Moid =" & m_MOid & ""
    
    OrderRs.Open OrderSql, con.ConnectionString, 3, 4
    '创建生产订单Order节点
    Set Order = GET_space_Orderxml
    Do While Not OrderRs.EOF
        For i = 0 To OrderRs.Fields.count - 1
            '节点赋值
            SET_IXMLDOMNode_text Order, OrderRs.Fields(i).Name, IIf(IsNull(OrderRs.Fields(i).Value), "", OrderRs.Fields(i).Value)
        Next i
        '节点追加
        Dom.selectSingleNode("//MOrder").appendChild Order '增加MOrder节点的下已节点 MOrder
        
        OrderSql = "select * from ref_v_mom_orderdetail where moid = " & OrderRs.Fields("MoId").Value
        MrderRs.Open OrderSql, con.ConnectionString, 3, 4
        
        Do While Not MrderRs.EOF
            '创建生产订单OrderDetail节点
            Set OrderDetail = GET_space_OrderDetailxml
            '节点赋值
            For i = 0 To MrderRs.Fields.count - 1
                SET_IXMLDOMNode_text OrderDetail, MrderRs.Fields(i).Name, IIf(IsNull(MrderRs.Fields(i).Value), "", MrderRs.Fields(i).Value)
            Next i
            '节点追加
            Dom.selectSingleNode("//MOrder").appendChild OrderDetail '增加MOrder节点的下已节点 OrderDetail
    
        
            '创建生产订单MOrderDetail节点
            Set MOrderDetail = GET_space_MOrderDetailxml
            '节点赋值
            For i = 0 To MrderRs.Fields.count - 1
                SET_IXMLDOMNode_text MOrderDetail, MrderRs.Fields(i).Name, IIf(IsNull(MrderRs.Fields(i).Value), "", MrderRs.Fields(i).Value)
            Next i
            '节点追加
            Dom.selectSingleNode("//MOrder").appendChild MOrderDetail '增加MOrder节点的下已节点 MOrderDetail
        
        
            OrderSql = "select * from ref_v_mom_moallocate where modid = " & MrderRs.Fields("MoDid").Value
            ArderRs.Open OrderSql, con.ConnectionString, 3, 4
            
            Do While Not ArderRs.EOF
                '创建生产订单Allocate节点
                Set Allocate = GET_space_Allocatexml
                '节点赋值
                For i = 0 To ArderRs.Fields.count - 1
                    
                    SET_IXMLDOMNode_text Allocate, ArderRs.Fields(i).Name, IIf(IsNull(ArderRs.Fields(i).Value), "", ArderRs.Fields(i).Value)
                                     
                Next i

                '节点追加
                Dom.selectSingleNode("//MOrder").appendChild Allocate '增加MOrder节点的下已节点 Allo
            
                ArderRs.MoveNext
            Loop
            MrderRs.MoveNext
        Loop
        OrderRs.MoveNext
    Loop
    Set GetProdectXml = Dom.cloneNode(True)
    
    If OrderRs.State <> 0 Then OrderRs.Close
    If MrderRs.State <> 0 Then MrderRs.Close
    If ArderRs.State <> 0 Then ArderRs.Close
    Exit Function
    
ErrGetProdectXml:
    If OrderRs.State <> 0 Then OrderRs.Close
    If MrderRs.State <> 0 Then MrderRs.Close
    If ArderRs.State <> 0 Then ArderRs.Close
    MsgBox Err.Description
End Function

Public Function GET_space_Orderxml() As IXMLDOMNode
    Dim Order As IXMLDOMNode
    Dim domxml As New DOMDocument
    Set Order = domxml.createElement("Order").cloneNode(True)
    Order.appendChild domxml.createElement("MoId").cloneNode(True)
    Order.appendChild domxml.createElement("MoCode").cloneNode(True)
    Order.appendChild domxml.createElement("CreateDate").cloneNode(True)
    Order.appendChild domxml.createElement("CreateUser").cloneNode(True)
    Order.appendChild domxml.createElement("Define1").cloneNode(True)
    Order.appendChild domxml.createElement("Define2").cloneNode(True)
    Order.appendChild domxml.createElement("Define3").cloneNode(True)
    Order.appendChild domxml.createElement("Define4").cloneNode(True)
    Order.appendChild domxml.createElement("Define5").cloneNode(True)
    Order.appendChild domxml.createElement("Define6").cloneNode(True)
    Order.appendChild domxml.createElement("Define7").cloneNode(True)
    Order.appendChild domxml.createElement("Define8").cloneNode(True)
    Order.appendChild domxml.createElement("Define9").cloneNode(True)
    Order.appendChild domxml.createElement("Define10").cloneNode(True)
    Order.appendChild domxml.createElement("Define11").cloneNode(True)
    Order.appendChild domxml.createElement("Define12").cloneNode(True)
    Order.appendChild domxml.createElement("Define13").cloneNode(True)
    Order.appendChild domxml.createElement("Define14").cloneNode(True)
    Order.appendChild domxml.createElement("Define15").cloneNode(True)
    Order.appendChild domxml.createElement("Define16").cloneNode(True)
    Set GET_space_Orderxml = Order
End Function



Public Function GET_space_OrderDetailxml() As IXMLDOMNode
    Dim Order As IXMLDOMNode
    Dim domxml As New DOMDocument
    Set Order = domxml.createElement("OrderDetail").cloneNode(True)
    Order.appendChild domxml.createElement("MoDId").cloneNode(True)
    Order.appendChild domxml.createElement("MoId").cloneNode(True)
    Order.appendChild domxml.createElement("SortSeq").cloneNode(True)
    Order.appendChild domxml.createElement("MoClass").cloneNode(True)
    Order.appendChild domxml.createElement("MoTypeCode").cloneNode(True)
    Order.appendChild domxml.createElement("Qty").cloneNode(True)
    Order.appendChild domxml.createElement("MrpQty").cloneNode(True)
    Order.appendChild domxml.createElement("AuxUnitCode").cloneNode(True)
    Order.appendChild domxml.createElement("AuxQty").cloneNode(True)
    Order.appendChild domxml.createElement("ChangeRate").cloneNode(True)
    Order.appendChild domxml.createElement("MoLotCode").cloneNode(True)
    Order.appendChild domxml.createElement("WhCode").cloneNode(True)
    Order.appendChild domxml.createElement("MDeptCode").cloneNode(True)
    Order.appendChild domxml.createElement("OrderCode").cloneNode(True)
    Order.appendChild domxml.createElement("OrderSeq").cloneNode(True)
    Order.appendChild domxml.createElement("DeclaredQty").cloneNode(True)
    Order.appendChild domxml.createElement("QualifiedInQty").cloneNode(True)
    Order.appendChild domxml.createElement("Status").cloneNode(True)
    Order.appendChild domxml.createElement("OrgStatus").cloneNode(True)
    Order.appendChild domxml.createElement("BomId").cloneNode(True)
    Order.appendChild domxml.createElement("RoutingId").cloneNode(True)
    Order.appendChild domxml.createElement("BomType").cloneNode(True)
    Order.appendChild domxml.createElement("BomVersion").cloneNode(True)
    Order.appendChild domxml.createElement("BomIdent").cloneNode(True)
    Order.appendChild domxml.createElement("RoutingType").cloneNode(True)
    Order.appendChild domxml.createElement("RoutingVersion").cloneNode(True)
    Order.appendChild domxml.createElement("RoutingIdent").cloneNode(True)
    Order.appendChild domxml.createElement("InvCode").cloneNode(True)
    Order.appendChild domxml.createElement("OrderType").cloneNode(True)
    Order.appendChild domxml.createElement("Free1").cloneNode(True)
    Order.appendChild domxml.createElement("Free2").cloneNode(True)
    Order.appendChild domxml.createElement("Free3").cloneNode(True)
    Order.appendChild domxml.createElement("Free4").cloneNode(True)
    Order.appendChild domxml.createElement("Free5").cloneNode(True)
    Order.appendChild domxml.createElement("Free6").cloneNode(True)
    Order.appendChild domxml.createElement("Free7").cloneNode(True)
    Order.appendChild domxml.createElement("Free8").cloneNode(True)
    Order.appendChild domxml.createElement("Free9").cloneNode(True)
    Order.appendChild domxml.createElement("Free10").cloneNode(True)
    Order.appendChild domxml.createElement("RelsDate").cloneNode(True)
    Order.appendChild domxml.createElement("RelsUser").cloneNode(True)
    Order.appendChild domxml.createElement("CloseDate").cloneNode(True)
    Order.appendChild domxml.createElement("OrgClsDate").cloneNode(True)
    Order.appendChild domxml.createElement("Define22").cloneNode(True)
    Order.appendChild domxml.createElement("Define23").cloneNode(True)
    Order.appendChild domxml.createElement("Define24").cloneNode(True)
    Order.appendChild domxml.createElement("Define25").cloneNode(True)
    Order.appendChild domxml.createElement("Define26").cloneNode(True)
    Order.appendChild domxml.createElement("Define27").cloneNode(True)
    Order.appendChild domxml.createElement("Define28").cloneNode(True)
    Order.appendChild domxml.createElement("Define29").cloneNode(True)
    Order.appendChild domxml.createElement("Define30").cloneNode(True)
    Order.appendChild domxml.createElement("Define31").cloneNode(True)
    Order.appendChild domxml.createElement("Define32").cloneNode(True)
    Order.appendChild domxml.createElement("Define33").cloneNode(True)
    Order.appendChild domxml.createElement("Define34").cloneNode(True)
    Order.appendChild domxml.createElement("Define35").cloneNode(True)
    Order.appendChild domxml.createElement("Define36").cloneNode(True)
    Order.appendChild domxml.createElement("Define37").cloneNode(True)
    Order.appendChild domxml.createElement("LeadTime").cloneNode(True)
    Order.appendChild domxml.createElement("WIPType").cloneNode(True)
    Order.appendChild domxml.createElement("OrdFlag").cloneNode(True)
    Order.appendChild domxml.createElement("SupplyWhCode").cloneNode(True)
    Order.appendChild domxml.createElement("ReasonCode").cloneNode(True)
    Order.appendChild domxml.createElement("SourceMoCode").cloneNode(True)
    Order.appendChild domxml.createElement("SourceQCCode").cloneNode(True)
    Order.appendChild domxml.createElement("SourceMoSeq").cloneNode(True)
    Order.appendChild domxml.createElement("CostItemCode").cloneNode(True)
    Order.appendChild domxml.createElement("CostItemName").cloneNode(True)
    Order.appendChild domxml.createElement("Remark").cloneNode(True)
    Order.appendChild domxml.createElement("AuditStatus").cloneNode(True)
    Order.appendChild domxml.createElement("IsWFControlled").cloneNode(True)
    Order.appendChild domxml.createElement("iVerifyState").cloneNode(True)
    Order.appendChild domxml.createElement("SoType").cloneNode(True)
    Order.appendChild domxml.createElement("SoCode").cloneNode(True)
    Order.appendChild domxml.createElement("SoSeq").cloneNode(True)
    Order.appendChild domxml.createElement("SoDId").cloneNode(True)
    Order.appendChild domxml.createElement("DemandCode").cloneNode(True)
    Order.appendChild domxml.createElement("ManualCode").cloneNode(True)
    Order.appendChild domxml.createElement("ReformFlag").cloneNode(True)
    Order.appendChild domxml.createElement("SourceQCVouchType").cloneNode(True)
    Order.appendChild domxml.createElement("QcFlag").cloneNode(True)
    Order.appendChild domxml.createElement("CollectiveFlag").cloneNode(True)
    Order.appendChild domxml.createElement("OpScheduleType").cloneNode(True)
    
Set GET_space_OrderDetailxml = Order
End Function


Public Function GET_space_MOrderDetailxml() As IXMLDOMNode
    Dim Order As IXMLDOMNode
    Dim domxml As New DOMDocument
    Set Order = domxml.createElement("MOrderDetail").cloneNode(True)
    Order.appendChild domxml.createElement("MoDId").cloneNode(True)
    Order.appendChild domxml.createElement("StartDate").cloneNode(True)
    Order.appendChild domxml.createElement("DueDate").cloneNode(True)
    Set GET_space_MOrderDetailxml = Order
End Function


Public Function GET_space_Allocatexml() As IXMLDOMNode
    Dim Order As IXMLDOMNode
    Dim domxml As New DOMDocument
    Set Order = domxml.createElement("Allocate").cloneNode(True)
    Order.appendChild domxml.createElement("AllocateId").cloneNode(True)
    Order.appendChild domxml.createElement("MoDId").cloneNode(True)
    Order.appendChild domxml.createElement("SortSeq").cloneNode(True)
    Order.appendChild domxml.createElement("OpSeq").cloneNode(True)
    Order.appendChild domxml.createElement("FVFlag").cloneNode(True)
    Order.appendChild domxml.createElement("BaseQtyN").cloneNode(True)
    Order.appendChild domxml.createElement("BaseQtyD").cloneNode(True)
    Order.appendChild domxml.createElement("ParentScrap").cloneNode(True)
    Order.appendChild domxml.createElement("CompScrap").cloneNode(True)
    Order.appendChild domxml.createElement("Qty").cloneNode(True)
    Order.appendChild domxml.createElement("IssQty").cloneNode(True)
    Order.appendChild domxml.createElement("DeclaredQty").cloneNode(True)
    Order.appendChild domxml.createElement("StartDemDate").cloneNode(True)
    Order.appendChild domxml.createElement("EndDemDate").cloneNode(True)
    Order.appendChild domxml.createElement("WhCode").cloneNode(True)
    Order.appendChild domxml.createElement("LotNo").cloneNode(True)
    Order.appendChild domxml.createElement("WIPType").cloneNode(True)
    Order.appendChild domxml.createElement("ByproductFlag").cloneNode(True)
    Order.appendChild domxml.createElement("Offset").cloneNode(True)
    Order.appendChild domxml.createElement("InvCode").cloneNode(True)
    Order.appendChild domxml.createElement("Free1").cloneNode(True)
    Order.appendChild domxml.createElement("Free2").cloneNode(True)
    Order.appendChild domxml.createElement("Free3").cloneNode(True)
    Order.appendChild domxml.createElement("Free4").cloneNode(True)
    Order.appendChild domxml.createElement("Free5").cloneNode(True)
    Order.appendChild domxml.createElement("Free6").cloneNode(True)
    Order.appendChild domxml.createElement("Free7").cloneNode(True)
    Order.appendChild domxml.createElement("Free8").cloneNode(True)
    Order.appendChild domxml.createElement("Free9").cloneNode(True)
    Order.appendChild domxml.createElement("Free10").cloneNode(True)
    Order.appendChild domxml.createElement("AuxUnitCode").cloneNode(True)
    Order.appendChild domxml.createElement("ChangeRate").cloneNode(True)
    Order.appendChild domxml.createElement("AuxBaseQtyN").cloneNode(True)
    Order.appendChild domxml.createElement("AuxQty").cloneNode(True)
    Order.appendChild domxml.createElement("ReplenishQty").cloneNode(True)
    Order.appendChild domxml.createElement("ProductType").cloneNode(True)
    Order.appendChild domxml.createElement("Define22").cloneNode(True)
    Order.appendChild domxml.createElement("Define23").cloneNode(True)
    Order.appendChild domxml.createElement("Define24").cloneNode(True)
    Order.appendChild domxml.createElement("Define25").cloneNode(True)
    Order.appendChild domxml.createElement("Define26").cloneNode(True)
    Order.appendChild domxml.createElement("Define27").cloneNode(True)
    Order.appendChild domxml.createElement("Define28").cloneNode(True)
    Order.appendChild domxml.createElement("Define29").cloneNode(True)
    Order.appendChild domxml.createElement("Define30").cloneNode(True)
    Order.appendChild domxml.createElement("Define31").cloneNode(True)
    Order.appendChild domxml.createElement("Define32").cloneNode(True)
    Order.appendChild domxml.createElement("Define33").cloneNode(True)
    Order.appendChild domxml.createElement("Define34").cloneNode(True)
    Order.appendChild domxml.createElement("Define35").cloneNode(True)
    Order.appendChild domxml.createElement("Define36").cloneNode(True)
    Order.appendChild domxml.createElement("Define37").cloneNode(True)
    Order.appendChild domxml.createElement("TransQty").cloneNode(True)
    Order.appendChild domxml.createElement("Remark").cloneNode(True)
    Order.appendChild domxml.createElement("SoType").cloneNode(True)
    Order.appendChild domxml.createElement("SoCode").cloneNode(True)
    Order.appendChild domxml.createElement("SoSeq").cloneNode(True)
    Order.appendChild domxml.createElement("SoDId").cloneNode(True)
    Order.appendChild domxml.createElement("DemandCode").cloneNode(True)
    Order.appendChild domxml.createElement("QmFlag").cloneNode(True)
    Order.appendChild domxml.createElement("QcFlag").cloneNode(True)
    
    Set GET_space_Allocatexml = Order
End Function

Public Sub SET_IXMLDOMNode_text(Nodexml As IXMLDOMNode, NodeName As String, NodeText As String)
    If Not Nodexml.selectSingleNode("//" & NodeName) Is Nothing Then
        Nodexml.selectSingleNode("//" & NodeName).Text = NodeText
    End If
End Sub
