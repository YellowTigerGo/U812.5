VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Cls_UI_interface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'接口类说明
'本插件主要是实现在用户界面上的接口插件
Option Explicit

Dim mLogin As Object             '登陆对象
Dim DBconn As Object             '数据库连接对象
'Dim DBconn As ADODB.Connection             '数据库连接对象

Dim CardNum As String            '单据Key
Dim Vouchers As Object           '单据对象                    'UAPVoucherControl85.ctlVoucher
Dim Vouchers_Frm As Object
Dim DomH As New DOMDocument
Dim DomB As New DOMDocument
'----------------------------------------------
Dim rss As New ADODB.Recordset  '数据集
Dim sSql As String
Dim sErr As String
Dim strPub As String
Dim count As Integer
Dim isOC As Boolean
Dim inum As Double

'展bom数据dom

'展bom数据Rs
Dim rds As New ADODB.Recordset
'----------------------------------------------


Public Sub Userdll_Init(g_busines As Object, m_login As Object, D_Conn As Object, Frm As Object, strCardNum As String, strUserErr As String, bSuc As Boolean)
    Set mLogin = m_login
    Set DBconn = D_Conn
    Set nvLogin = m_login
    Set con = D_Conn
    Set Vouchers_Frm = Frm
    CardNum = strCardNum
    bSuc = True
    
    cls_Public.WrtDBlog con, nvLogin.cUserId, "Interface_fxm", "Userdll_Init 初始化！"
'    MsgBox " Userdll_Init 插件初始化配置成功！"
End Sub

'系统菜单按扭前事件
'Vouhcher 单据对象
'ButtonKey 按钮Key
'strUserErr 错误信息
'bSuc 返回结果，false 继走系统按钮功能，true到此结束
Public Sub Before_ButtonClick(voucher As Object, ButtonKey As String, strUserErr As String, bSuc As Boolean)
    Dim menuId As String
    Dim Name As String
    Dim subSysId As String
    Dim authid As String
    Dim voucherID As String
    Dim row As Integer  '第几行
    Dim rowsel As Integer '选择多少行
    Dim bodyrows As String '实际数据有多少行
    Dim CBATCH As String '批号
    Dim cInvCode As String '存货编码

'    cls_Public.WrtDBlog con, nvLogin.cUserId, "Interface_fxm", "Before_ButtonClick 初始化！"
    
    
    bSuc = False
    Select Case LCase(CardNum)
    
        Case LCase("YXEF9144")
            Select Case LCase(ButtonKey)
                
                Case LCase("cgps")  '采购评审
                    If voucher.headerText("ccode") = "" Then
                        MsgBox "当前无对应评审单据！"
                    Else
                        cls_Public.WrtDBlog con, nvLogin.cUserId, "Interface_fxm", "调用BOM......,展bom!"
                        BtnErr = GetBom(voucher)
                        If BtnErr <> "" Then
                            MsgBox BtnErr
                        Else
                            voucher.headerText("datetime3") = mLogin.CurDate
                            voucher.headerText("str6") = mLogin.cUserName
                            nvsql = "update EF_SellAccreditation set str6='" & mLogin.cUserName & "' ,datetime3='" & mLogin.CurDate & "'  where id =" & voucher.headerText("id")
                            DBconn.Execute nvsql
                        End If
                        
                    End If
                Case LCase("cgsh")
                    voucher.headerText("datetime1") = mLogin.CurDate
                    voucher.headerText("str1") = mLogin.cUserName
                    nvsql = "update EF_SellAccreditation set str1='" & mLogin.cUserName & "' ,datetime1='" & mLogin.CurDate & "'  where id =" & voucher.headerText("id")
                    DBconn.Execute nvsql
                Case LCase("scsh")
                    voucher.headerText("datetime2") = mLogin.CurDate
                    voucher.headerText("str2") = mLogin.cUserName
                    nvsql = "update EF_SellAccreditation set str2='" & mLogin.cUserName & "' ,datetime2='" & mLogin.CurDate & "'  where id =" & voucher.headerText("id")
                    DBconn.Execute nvsql
                Case LCase("modify")
                    voucher.headerText("datetime4") = mLogin.CurDate
                    voucher.headerText("str7") = mLogin.cUserName
'                    nvsql = "update EF_SellAccreditation set str7='" & mLogin.cUserName & "' ,datetime4='" & mLogin.CurDate & "'  where id =" & voucher.headerText("id")
'                    DBconn.Execute nvsql
'                    bSuc = False
                Case LCase("refsh")
                    CreateOrder voucher.headerText("ccode")
            End Select
            
    
        Case LCase("YXEF9104")
        
            If ButtonKey = "dust" Then
                '菜单ID
                menuId = "YXEF9114"
                '单据名称
                Name = "废品报工单"
                '子产品号
                subSysId = " "
                '权限ID
                authid = "YXEF911401"
                '单据ID
                voucherID = 0
                
                Call Vouchers_Frm.UFToolbar1.Business.ResetCallFunctionFlag
                Dim strNewCommandline As String
                strNewCommandline = "ID:" & menuId & "&&&Name:" & Name & "&&&SubSysID:" & subSysId & "&&&AuthID:" & authid & "&&&CmdLine:" & vbTab & voucherID
                Call Vouchers_Frm.UFToolbar1.Business.RunCommand(strNewCommandline)
            Else
                bSuc = False
                Exit Sub
            End If
        Case Else
            bSuc = False
            Exit Sub
    End Select
     
End Sub

'系统菜单按扭后事件
Public Sub After_ButtonClick(voucher As Object, ButtonKey As String, strUserErr As String, bSuc As Boolean)
    bSuc = False
    
    If LCase(CardNum) = LCase("YXEF9105") And LCase(ButtonKey) = LCase("copy") Then
        voucher.headerText("id") = ""
    End If
    
    Select Case LCase(CardNum)
        Case LCase("YXEF9122"), LCase("YXEF9123"), LCase("'YXEF9124"), LCase("'YXEF9125"), LCase("'YXEF9126")
            If LCase(ButtonKey) = LCase("modify") Then
                editprop = "M"
            End If
            
            If LCase(ButtonKey) = LCase("cancel") Then
                editprop = ""
            End If
    End Select
      
'    If LCase(CardNum) = LCase("YXEF9144") And LCase(ButtonKey) = LCase("cgps") Then
'        If BtnErr = "" Then
'            voucher.headerText("datetime3") = mLogin.CurDate
'            voucher.headerText("str6") = mLogin.cUserName
'            nvsql = "update EF_SellAccreditation set str6='" & mLogin.cUserName & "' ,datetime3='" & mLogin.CurDate & "'  where id =" & voucher.headerText("id")
'            DBconn.Execute nvsql
'        End If
'    End If
'
'    If LCase(CardNum) = LCase("YXEF9144") And LCase(ButtonKey) = LCase("cgsh") Then
'        voucher.headerText("datetime1") = mLogin.CurDate
'        voucher.headerText("str1") = mLogin.cUserName
'        nvsql = "update EF_SellAccreditation set str1='" & mLogin.cUserName & "' ,datetime1='" & mLogin.CurDate & "'  where id =" & voucher.headerText("id")
'        DBconn.Execute nvsql
'    End If
'
'    If LCase(CardNum) = LCase("YXEF9144") And LCase(ButtonKey) = LCase("scsh") Then
'        voucher.headerText("datetime2") = mLogin.CurDate
'        voucher.headerText("str2") = mLogin.cUserName
'        nvsql = "update EF_SellAccreditation set str2='" & mLogin.cUserName & "' ,datetime2='" & mLogin.CurDate & "'  where id =" & voucher.headerText("id")
'        DBconn.Execute nvsql
'    End If
'
'
'    If LCase(CardNum) = LCase("YXEF9144") And LCase(ButtonKey) = LCase("modify") Then
'        voucher.headerText("datetime4") = mLogin.CurDate
'        voucher.headerText("str7") = mLogin.cUserName
'        nvsql = "update EF_SellAccreditation set str7='" & mLogin.cUserName & "' ,datetime4='" & mLogin.CurDate & "'  where id =" & voucher.headerText("id")
'        DBconn.Execute nvsql
'        bSuc = False
'    End If
'
'    If LCase(CardNum) = LCase("YXEF9144") And LCase(ButtonKey) = LCase("refsh") Then
'        MsgBox "111111"
'        CreateOrder voucher.headerText("ccode")
'        MsgBox "222222"
'    End If
                    
'        MsgBox " After_ButtonClick 插件配置成功！"
End Sub

'提示：单据表头参照事件
Public Sub Voucher_headBrowUser(ByVal voucher As Object, ByVal sKey As String, sRet As Variant, strUserErr As String, bSuc As Boolean)
    Dim sMetaXML As String
    Dim retRstClass As ADODB.Recordset, retRstGrid As ADODB.Recordset
    Dim sErrMsg As String
    Dim dept As String
    Dim cpersoncode As String, cpersonname As String
    Dim U8ref As New U8RefService.IService
    Set Vouchers = voucher
      
    Select Case LCase(CardNum)
        Case LCase("YXEF9105"), LCase("YXEF9104"), LCase("YXEF9114"), LCase("YXEF9115"), LCase("YXEF9116"), LCase("YXEF9117"), LCase("YXEF9118"), LCase("YXEF9119"), LCase("YXEF9120"), LCase("YXEF9131"), LCase("YXEF9132"), LCase("YXEF9133"), LCase("YXEF9134"), LCase("YXEF9135"), LCase("YXEF9136")
            Select Case sKey
                Case "t_cpersoncode", "t_cpersonname"
'                    U8ref.RefID = "person_aa"
                    U8ref.RefID = "hr_hi_person_aa"
                    'U8ref.FilterSQL = "[test_tableEntity_test_tables].[cInvCode]='mb' "
                    
                    dept = Vouchers.headerText("t_cdepcode")
                    
                    If dept = "" Then
                        U8ref.FilterSQL = ""
                    Else
                        If LCase(CardNum) = LCase("YXEF9105") Then
                            U8ref.FilterSQL = ""
                        Else
                            U8ref.FilterSQL = "[HR_HI_PersonEntity_人员档案].cpsn_num in (select cpsn_num from hr_hi_person  where cdept_num = '" & dept & "')"
                        End If
                        
                        
                    End If
                    
                    
                    U8ref.MetaXML = "<Ref><RefSet bMultiSel= '1' /></Ref>"
                    If U8ref.ShowRef(mLogin, retRstClass, retRstGrid, sErrMsg) = False Then
                        MsgBox sErrMsg
                    Else
                        If Not (retRstGrid Is Nothing) Then
                            Do Until retRstGrid.EOF
                                cpersoncode = cpersoncode & retRstGrid.Fields(0) '& "/"
                                cpersonname = cpersonname & retRstGrid.Fields(1) '& "/"
                                retRstGrid.MoveNext
                                If Not retRstGrid.BOF And Not retRstGrid.EOF Then
                                    cpersoncode = cpersoncode & "/"
                                    cpersonname = cpersonname & "/"
                                End If
                            Loop
                        End If
                        Vouchers.headerText("t_cpersoncode") = cpersoncode
                        Vouchers.headerText("t_cpersonname") = cpersonname
                        sRet = Vouchers.headerText(sKey)
                        bSuc = True
                        Exit Sub
                    End If

            End Select
    End Select
    bSuc = False
End Sub

'初始化用户插件
'Voucher    当前单据对象
'sKey       当前操作关键字
'sRet       返回值
'strUserErr  错误信息
'bChanged
'bSuc  =true 表示成功并终止系统流程   =false表示失败并继续系统流程
'表头校验事件
Public Sub Voucher_headCellCheck(ByVal voucher As Object, sKey As String, RetValue As String, bChanged As Long, strUserErr As String, bSuc As Boolean)
    Dim person() As String
    Dim msg As String
    Dim i As Integer
    Dim where As String
    
    Dim personcode As String
    Dim personname As String
    
    Set Vouchers = voucher
     
    Select Case LCase(CardNum)
        Case LCase("YXEF9101"), LCase("YXEF9122"), LCase("YXEF9123"), LCase("YXEF9124"), LCase("YXEF9125"), LCase("YXEF9126")
            msg = isCheck(DBconn, CardNum, sKey)
            If msg = "4" Then
                bSuc = True
            End If
        Case LCase("YXEF9104"), LCase("YXEF9105"), LCase("YXEF9114"), LCase("YXEF9115"), LCase("YXEF9116"), LCase("YXEF9117"), LCase("YXEF9118"), LCase("YXEF9119"), LCase("YXEF9120"), LCase("YXEF9131"), LCase("YXEF9132"), LCase("YXEF9133"), LCase("YXEF9134"), LCase("YXEF9135"), LCase("YXEF9136")
            bSuc = False
            Select Case sKey
                Case "t_cpersoncode", "t_cpersonname"
                    
                    If Vouchers.headerText(sKey) = "" Then
                    
                        Vouchers.headerText("t_cpersoncode") = ""
                        Vouchers.headerText("t_cpersonname") = ""
                        RetValue = Vouchers.headerText(sKey)
                        bSuc = True
                        Exit Sub
                        
                    End If
                    
                    
                    msg = isCheck(DBconn, CardNum, sKey, Vouchers)
                    If msg = "false" Then
                        MsgBox "人员录入不正确或者没有权限或者被停用，请检查！"
                        Vouchers.headerText("t_cpersoncode") = ""
                        Vouchers.headerText("t_cpersonname") = ""
                        RetValue = Vouchers.headerText(sKey)
                        bSuc = True
                    Else
                        If Vouchers.headerText(sKey) <> "" And Not IsNull(Vouchers.headerText(sKey)) Then
                            person = Split(Vouchers.headerText(sKey), "/")
                            where = ""
                            
                            If UBound(person) > 0 Then
                                For i = 0 To (UBound(person) - LBound(person))
                                    where = where & "'" & person(i) & "'"
                                    If i < UBound(person) Then
                                        where = where & ","
                                    End If
                                Next i
                            Else
                                If UBound(person) = 0 Then
                                    where = "'" & Vouchers.headerText(sKey) & "'"
                                End If
                            End If
                            
                            
                            If where <> "" Then
                                setVoucher DBconn, CardNum, sKey, where, voucher
                                bSuc = True
                                Exit Sub
                            End If
                        End If
                    End If
                Case "str5", "str6"
                
                    sSql = "select a.b_str1,a.b_str2,b.cpersoncode,b.cpersonname from EF_dep_Bzs a left join person b on a.b_cpersoncode = b.cpersoncode"
                    sSql = sSql & " where a.b_str1 = '" & Vouchers.headerText("str5") & "'"
                    Set rss = Query(sSql, DBconn, strUserErr)
                    
                    Do Until rss.EOF
                        personcode = personcode & rss.Fields("cpersoncode")
                        personname = personname & rss.Fields("cpersonname")
                        rss.MoveNext
                        If Not rss.BOF And Not rss.EOF Then
                            personcode = personcode & "/"
                            personname = personname & "/"
                        End If
                    Loop
                    
                    Vouchers.headerText("t_cpersoncode") = personcode
                    Vouchers.headerText("t_cpersonname") = personname
                    
            End Select
              
              
        Case LCase("YXEF9122"), LCase("YXEF9123"), LCase("YXEF9124"), LCase("YXEF9125"), LCase("YXEF9126")
            Dim cInv As String
            cInv = Mid(Vouchers.headerText("t_cinvcode"), 2, 5)
'            ssql = "select cust.cCusCode,cust.cCusName,per.cPersonCode,per.cPersonName from Customer cust left join Person per on cust.cCusPPerson = per.cpersoncode where cust.cCusCode ='" & cinv & "'"
            sSql = "select cust.cCusCode,cust.cCusName from Customer cust where cust.cCusCode ='" & cInv & "'"

            Set rss = Query(sSql, DBconn, strUserErr)
            If Not rss.BOF And Not rss.EOF Then
            Vouchers.headerText("t_ccuscode") = rss.Fields("cCusCode")
            Vouchers.headerText("t_ccusname") = rss.Fields("cCusName")
'            voucher.headerText("cPersonCode") = Rss.Fields("cPersonCode")
'            voucher.headerText("cPersonName") = Rss.Fields("cPersonName")
            End If
                
        Case Else
            bSuc = False
    End Select
  
'    bSuc = False
'        MsgBox " Voucher_headCellCheck 插件配置成功！"
End Sub

 

'表体参照事件
Public Sub Voucher_bodyBrowUser(ByVal voucher As Object, sKey As String, ByVal r As Long, sRet As Variant, strUserErr As String, bSuc As Boolean)
    
    Dim sMetaXML As String
    Dim retRstClass As ADODB.Recordset, retRstGrid As ADODB.Recordset
    Dim sErrMsg As String
    Dim dept As String
    Dim cpersoncode As String, cpersonname As String
    Dim U8ref As New U8RefService.IService
    Set Vouchers = voucher
      
    Select Case LCase(CardNum)
        Case LCase("YXEF9104"), LCase("YXEF9114"), LCase("YXEF9115"), LCase("YXEF9116"), LCase("YXEF9117"), LCase("YXEF9118"), LCase("YXEF9119"), LCase("YXEF9120"), LCase("YXEF9131"), LCase("YXEF9132"), LCase("YXEF9133"), LCase("YXEF9134"), LCase("YXEF9135"), LCase("YXEF9136")
            Select Case sKey
                Case "b_cpersoncode", "b_cpersonname"
                    U8ref.RefID = "hr_hi_person_aa"
                    
'                    dept = Vouchers.bodyText(R, "t_cdepcode")
                    U8ref.FilterSQL = ""
'                    If dept = "" Then
'                        U8ref.FilterSQL = ""
'                    Else
'                        U8ref.FilterSQL = "[HR_HI_PersonEntity_人员档案].cpsn_num in (select cpsn_num from hr_hi_person  where cdept_num = '" & dept & "')"
'                    End If
                    
                    
                    U8ref.MetaXML = "<Ref><RefSet bMultiSel= '1' /></Ref>"
                    If U8ref.ShowRef(mLogin, retRstClass, retRstGrid, sErrMsg) = False Then
                        MsgBox sErrMsg
                    Else
                        If Not (retRstGrid Is Nothing) Then
                            Do Until retRstGrid.EOF
                                cpersoncode = cpersoncode & retRstGrid.Fields(0) '& "/"
                                cpersonname = cpersonname & retRstGrid.Fields(1) '& "/"
                                retRstGrid.MoveNext
                                If Not retRstGrid.BOF And Not retRstGrid.EOF Then
                                    cpersoncode = cpersoncode & "/"
                                    cpersonname = cpersonname & "/"
                                End If
                            Loop
                        End If
                        Vouchers.bodyText(r, "b_cpersoncode") = cpersoncode
                        Vouchers.bodyText(r, "b_cpersonname") = cpersonname
                        sRet = Vouchers.bodyText(r, sKey)
                        bSuc = True
                        Exit Sub
                    End If

            End Select
    End Select
    bSuc = False
End Sub


'表体校验事件
Public Sub Voucher_bodyCellCheck(ByVal voucher As Object, RetValue As Variant, bChanged As Long, sKey As String, ByVal r As Long, strUserErr As String, bSuc As Boolean)
    Dim float1 As Double '宽度b_float10
    Dim float2 As Double '厚度b_float11
    Dim float3 As Double '密度b_float9
    Dim float4 As Double '来料重量b_float2
    
    Dim person() As String
    Dim msg As String
    Dim i As Integer
    Dim where As String
    
    Dim personcode As String
    Dim personname As String
    
    rows = r
    Set Vouchers = voucher
    
    On Error GoTo ErrorMsg
    
    Set Vouchers = voucher
    bSuc = False
    Select Case LCase(CardNum)
    
        Case LCase("YXEF9115")
        
            Select Case LCase(sKey)
        
                Case LCase("b_float10"), LCase("b_float11"), LCase("b_float9"), LCase("b_float2")
        
                    float1 = CDbl(IIf(Vouchers.bodyText(r, "b_float10") = "", 0, Vouchers.bodyText(r, "b_float10")))
                    float2 = CDbl(IIf(Vouchers.bodyText(r, "b_float11") = "", 0, Vouchers.bodyText(r, "b_float11")))
                    float3 = CDbl(IIf(Vouchers.bodyText(r, "b_float9") = "", 0, Vouchers.bodyText(r, "b_float9")))
                    float4 = CDbl(IIf(Vouchers.bodyText(r, "b_float2") = "", 0, Vouchers.bodyText(r, "b_float2")))
        
        
                    If float1 <> 0 And float2 <> 0 And float3 <> 0 And float4 <> 0 Then
                    voucher.bodyText(r, "b_float1") = float3 / float2 / float1 / float4 * 1000000
                    RetValue = float3 / float2 / float1 / float4 * 1000000
                    End If
            End Select
        
        Case LCase("YXEF9104"), LCase("YXEF9114"), LCase("YXEF9115"), LCase("YXEF9116"), LCase("YXEF9117"), LCase("YXEF9118"), LCase("YXEF9119"), LCase("YXEF9120"), LCase("YXEF9131"), LCase("YXEF9132"), LCase("YXEF9133"), LCase("YXEF9134"), LCase("YXEF9135"), LCase("YXEF9136")
            bSuc = False
            Select Case sKey
                Case "b_cpersoncode", "b_cpersonname"
                
                    If Vouchers.bodyText(r, sKey) = "" Then
                        Vouchers.bodyText(r, "b_cpersoncode") = "" 'Vouchers.headerText("t_cpersoncode") = ""
                        Vouchers.bodyText(r, "b_cpersonname") = ""
                        RetValue = Vouchers.bodyText(r, sKey)
                        bSuc = True
                        Exit Sub
                    End If
            
                    msg = isCheck(DBconn, CardNum, sKey, Vouchers)
                    If msg = "false" Then
                        MsgBox "人员录入不正确或者没有权限或者被停用，请检查！"
                        Vouchers.bodyText(r, "b_cpersoncode") = ""
                        Vouchers.bodyText(r, "b_cpersonname") = ""
                        RetValue = Vouchers.bodyText(r, sKey)
                        bSuc = True
                    Else
                        If Vouchers.bodyText(r, sKey) <> "" And Not IsNull(Vouchers.bodyText(r, sKey)) Then
                            person = Split(Vouchers.bodyText(r, sKey), "/")
                            where = ""
                
                            If UBound(person) > 0 Then
                            For i = 0 To (UBound(person) - LBound(person))
                                where = where & "'" & person(i) & "'"
                                If i < UBound(person) Then
                                    where = where & ","
                                End If
                            Next i
                        Else
                            If UBound(person) = 0 Then
                                where = "'" & Vouchers.bodyText(r, sKey) & "'"
                            End If
                        End If
                    
                        If where <> "" Then
                            setVoucher DBconn, CardNum, sKey, where, voucher
                            bSuc = True
                            Exit Sub
                        End If
                    End If
                End If
        
            End Select
    
        End Select
    Exit Sub
              
ErrorMsg:
    strUserErr = Err.Description
    bSuc = False
End Sub



'提示：单据表头列表式项目参照事件
Public Sub Voucher_FillHeadComboBox(ByVal voucher As Object, sKey As String, pCom As Object, strUserErr As String, bSuc As Boolean)
    Set Vouchers = voucher
'    Dim i As Integer
'    Dim rds As New ADODB.Recordset
'    Select Case LCase(Me.Voucher.ItemState(Index, siheader).sFieldName)
'
'        Case "iyear" '编制年度
'                pCom.Clear
'                For i = 0 To 4
'                  pCom.AddItem CStr(Val(m_login.cIYear) + i)
'                Next

'    End Select


         bSuc = False
'        MsgBox " Voucher_FillHeadComboBox 插件配置成功！"

End Sub
 
 
 
 '提示：单据表体列表式项目参照事件
Public Sub Voucher_FillList(ByVal voucher As Object, sKey As String, ByVal r As Long, pCom As Object, strUserErr As String, bSuc As Boolean)
    Set Vouchers = voucher
    'Dim sFieldName As String
    'sFieldName = LCase(Me.Voucher.ItemState(C, sibody).sFieldName)
    'Select Case sFieldName
    '    Case "adds", "lenssen" '控制方向
    '        pCom.Clear
    '        pCom.AddItem ""
    '        pCom.AddItem "借方"
    '        pCom.AddItem "贷方"
 
    '

         bSuc = True
'        MsgBox " Voucher_FillList 插件配置成功！"
    
End Sub

''
'Source_Cardnumber    来源单句
'Dest_Cardnumber      目标单据

Public Sub AddNew_Voucher(Source_Cardnumber As String, Dest_Cardnumber As String, Dest_domHead As DOMDocument, Dest_domBody As DOMDocument, strUserErr As String, bSuc As Boolean)
On Error GoTo ExitSub
    Dim FrmR As New frmRefernew ' frmRefer
    bSuc = False
    With FrmR
        
        Call .init(mLogin, DBconn, FrmR, Dest_domHead, Dest_domBody, Source_Cardnumber, Dest_Cardnumber)
        If .filter Then
            FrmR.SetData
            FrmR.Show 1
            Set Dest_domHead = .DomHead_Dest.cloneNode(True)
            Set Dest_domBody = .DomBody_Dest.cloneNode(True)
            Vouchers_Frm.voucher.setVoucherDataXML Dest_domHead, Dest_domBody
        End If
    End With
    bSuc = True
    Set FrmR = Nothing
    Exit Sub
ExitSub:
    bSuc = False
    strUserErr = strUserErr & Err.Description
End Sub


'个性化生单填充方法
'Source_Cardnumber  来源参照类型
'Dest_Cardnumber    目标单据的cardnumber
'Source_domHead     参照生单界面上半部分数据对象
'Source_domBody     参照生单界面上半部分数据对象
'Dest_domHead       目标单据 表头数据对象
'Dest_domHead       目标单据 表头数据对象
'strUserErr         错误信息
'bSuc               =true 表示个性化填充       =false 表示系统自动填中

Public Sub Full_Voucher(Source_Cardnumber As String, Dest_Cardnumber As String, Source_domHead As DOMDocument, Source_domBody As DOMDocument, Dest_domHead As DOMDocument, Dest_domBody As DOMDocument, strUserErr As String, Optional bSuc As Boolean = False)
On Error GoTo ExitSub
    bSuc = False

    '    FrmR.ctlVoucher.setVoucherDataXML Dest_domHead, Dest_domBody
    '    bSuc = True
    Exit Sub
ExitSub:
    bSuc = False
    strUserErr = strUserErr & Err.Description
End Sub


Public Function RunCommand(ByVal objLogin As Object, ByVal objForm As Object, ByVal objVoucher As Object, _
                            ByVal sKey As String, ByVal VarentValue As Variant, ByVal other As String)
    Dim id As String        '销售订单主表ID
    Dim ccode As String     '销售订单号
    Dim sql As String
    Dim BeginTrans As Boolean
    Dim DBconns  As New ADODB.Connection
    
    On Error GoTo ExitFunction
    
    id = objVoucher.headerText("id")
    
    If id = "" Then
        
        MsgBox "提交失败！当前无销售订单！"
        Exit Function
    
    End If
    
    If objVoucher.headerText("cverifier") = "" Then
    
        MsgBox "提交失败！当前销售订单未审核，不可以提交评审！"
        Exit Function
    
    End If
        
    If DBconns.State <> 0 Then DBconns.Close
    DBconns.Open objLogin.UfDbName
    If BeginTrans = False Then
        DBconns.BeginTrans
        BeginTrans = True
    End If
    
    sql = "delete ef_sellaccreditations where id =" & id
    DBconns.Execute sql
    sql = "delete ef_sellaccreditation where id =" & id
    DBconns.Execute sql
    
'    sql = "insert into ef_sellaccreditation ([id],[ccode]) select id,cSOCode  from SO_SOMain  where id=" & id
    sql = "insert into ef_sellaccreditation ([id],[SoID]) select id,ID  from SO_SOMain  where id=" & id
    DBconns.Execute sql
    sql = "insert into ef_sellaccreditations ([id],[autoid]) select id,autoid from SO_SODetails where id =" & id
    DBconns.Execute sql
    
    If BeginTrans Then
        DBconns.CommitTrans
        BeginTrans = False
    End If
    
    
    MsgBox "提交成功！"
    
    Exit Function
    
ExitFunction:

    If BeginTrans Then
        DBconns.RollbackTrans
        BeginTrans = False
    End If
    MsgBox "提交失败！ 【" & Err.Description & "】"
End Function

Public Function init(ByVal objLogin As Object, ByVal objForm As Object, ByVal objVoucher As Object, msbar As Object)
'MsgBox "Init"
End Function

Private Function CreateOrder(soCode As String)
    Dim QuerySql As String
    Dim QueryRs As New ADODB.Recordset
    
    QuerySql = "select moid from mom_order where define14 = '" & soCode & "'"
    
    If nvRs.State <> 0 Then nvRs.Close
    nvRs.Open QuerySql, con.ConnectionString, 3, 4
    
    If Not nvRs.BOF And Not nvRs.EOF Then
        MsgBox "编号【" & soCode & "】的销售订单已生成生产订单，需手工删除对应单据，才可重新生成生产订单！"
    Else
        
        QueryRs.Open "select ccode,moid from EF_SellAccreditation where ccode = '" & soCode & "'", con.ConnectionString, 3, 4
        If Not QueryRs.BOF And Not QueryRs.EOF Then
            '生成系统生产订单
            Call makeorrder(QueryRs.Fields("moid").Value)
        Else
            MsgBox "编号【" & soCode & "】的销售订单评审结果丢失，请重新评审！"
        End If
    End If
    
    
End Function

Public Function GetBom(t_voucher As Object) As String
    Dim soInvRs As New ADODB.Recordset
    Dim blRs As New ADODB.Recordset
    
    On Error GoTo ExitSub
    
    If soInvRs.State <> 0 Then soInvRs.Close
    soInvRs.Open "select b_cinvcode as InvCode,* from V_EF_SO_SellAccreditations where id =" & t_voucher.headerText("id"), con.ConnectionString, 3, 4
    
    isTrans = False
'        开始事务
    If isTrans = False Then
'        con.BeginTrans
        isTrans = True
    End If
    
    If blRs.State <> 0 Then blRs.Close
    blRs.Open "select * from ef_bomchenage where cSOID =" & t_voucher.headerText("id"), con.ConnectionString, 3, 4
    
    If Not blRs.BOF And Not blRs.EOF Then
        
        If MsgBox("已做过采购评审，是否重新评审", vbYesNo, "") = vbYes Then
              
            nvsql = "delete ef_bomchenage where cSOID =" & t_voucher.headerText("id")
            con.Execute nvsql
            cls_Public.WrtDBlog con, nvLogin.cUserId, "Interface_fxm", "调用CreateBomChenage,展bom!"
        
            Do Until soInvRs.EOF
                CreateBomChenage soInvRs.Fields("InvCode"), soInvRs.Fields("b_float1"), soInvRs.Fields("id"), soInvRs.Fields("autoid")
                soInvRs.MoveNext
            Loop
            
'            RefProduct t_voucher.headerText("id"), t_voucher.headerText("ccode")
'            cgpsForm.init t_voucher, DBconn, t_voucher.headerText("ccode"), t_voucher.headerText("id")
'            cgpsForm.Show 1
         End If
     
    Else
             
        nvsql = "delete ef_bomchenage where cSOID =" & t_voucher.headerText("id")
        con.Execute nvsql
        cls_Public.WrtDBlog con, nvLogin.cUserId, "Interface_fxm", "调用CreateBomChenage,展bom!"
    
        Do Until soInvRs.EOF
            CreateBomChenage soInvRs.Fields("InvCode"), soInvRs.Fields("b_float1"), soInvRs.Fields("id"), soInvRs.Fields("autoid")
            soInvRs.MoveNext
        Loop
        
'        RefProduct t_voucher.headerText("id"), t_voucher.headerText("ccode")
'        cgpsForm.init t_voucher, DBconn, t_voucher.headerText("ccode"), t_voucher.headerText("id")
'        cgpsForm.Show 1
    End If
    
    '提交操作
    If isTrans Then
'        con.CommitTrans
        isTrans = False
    End If
    
    
    cgpsForm.init t_voucher, DBconn, t_voucher.headerText("ccode"), t_voucher.headerText("id")
    cgpsForm.Show 1

    GetBom = ""
    Exit Function
        
ExitSub:
    If isTrans Then
'        con.RollbackTrans
        isTrans = False
    End If
    GetBom = "提交失败！------ 【" & Err.Description & "】"
End Function

Private Sub Class_Initialize()
    Set cls_Public = CreateObject("UF_Public_base.cls_log")
End Sub

Private Sub CreateBomChenage(cInvCodes As String, iQuantitys As Double, soCode As Long, soAutoid As Long)
    Dim ufu8mbom As New UFU8MBOMISrv.clsBom
    Dim trues As Boolean
    Dim cInvCode As String
    Dim bomErrstr As String
    Dim iQuantity As Double

    Dim sql As String
    Dim domtemp As New DOMDocument
            
    Dim i As Integer
    Dim j As Integer
    
    Dim rds2 As New ADODB.Recordset

    cInvCode = cInvCodes
    iQuantity = iQuantitys
    
    
'   展BOM
    cls_Public.WrtDBlog con, nvLogin.cUserId, "Interface_fxm  ufu8mbom.GetBomComponent", cInvCode & "展BOM"
    trues = ufu8mbom.GetBomComponent(mLogin, cInvCode, "", "", "", "", "", "", "", "", "", "", False, rds2, bomErrstr, iQuantity, , , True, , , , , True)
    cls_Public.WrtDBlog con, nvLogin.cUserId, "Interface_fxm  ufu8mbom.GetBomComponent   ", "trues=" & trues
    
'           判断是否成功展开BOM，并做相应操作
    If trues Then
        rds2.Save domtemp, adPersistXML
        cls_Public.WrtDBlog con, nvLogin.cUserId, "Interface_fxm  调用CreateBomChenage 1", "rds2.RecordCount=" & rds2.RecordCount & "    " & domtemp.xml
        
        Do Until rds2.EOF
            cls_Public.WrtDBlog con, nvLogin.cUserId, "Interface_fxm  递归 调用CreateBomChenage 2", "[" & rds2.Fields("InvCode").Value & "]" & rds2.Fields("Cqty").Value
            CreateBomChenage rds2.Fields("InvCode").Value, rds2.Fields("Cqty").Value, soCode, soAutoid
            
            rds2.MoveNext
            
        Loop
        
        cls_Public.WrtDBlog con, nvLogin.cUserId, "Interface_fxm 调用CreateBomChenage3 ", "开始循环写入母子件关系"
        
        insetBOM domtemp, con, cInvCodes, iQuantitys, soCode, soAutoid

    End If

End Sub

Public Function insetBOM(domtemp As DOMDocument, domConn As Connection, cInvCodes As String, iQuantitys As Double, soCode As Long, soAutoid As Long)
    'BOM展开
    Dim bom_OpComponentId    As String
    Dim bom_OpSeq            As String
    Dim bom_CompId           As String
    Dim bom_UnitId           As String
    Dim bom_BaseQtyN         As String
    Dim bom_BaseQtyD         As String
    Dim bom_ParentScrap      As String
    Dim bom_CompScrap        As String
    Dim bom_FVQty            As Integer
    Dim bom_Cqty             As Double
    Dim bom_Cqty1            As Double
    Dim bom_UseQty           As Double
    Dim bom_Offset           As Double
    Dim bom_WIPtype          As Integer
    Dim bom_WhCode           As String
    Dim bom_InvCode          As String
    Dim bom_Free1            As String
    Dim bom_Free2            As String
    Dim bom_Free3            As String
    Dim bom_Free4            As String
    Dim bom_Free5            As String
    Dim bom_Free6            As String
    Dim bom_Free7            As String
    Dim bom_Free8            As String
    Dim bom_Free9            As String
    Dim bom_Free10           As String
    Dim bom_Dept             As String
    Dim bom_DepName          As String
    Dim bom_ByproductFlag    As Integer
    Dim bom_AccuCostFlag     As Integer
    Dim bom_SubFlag          As Integer
    Dim bom_BomType          As Integer
    Dim bom_iGrade           As Integer
    Dim bom_DemDate          As String
    Dim bom_AuxUnitCode      As String
    Dim bom_ChangeRate       As Double
    Dim bom_AuxBaseQtyN      As Double
    Dim bom_AuxCqty          As String
    Dim bom_AuxUseQty        As Double
    Dim bom_AuxUnitName      As String
    Dim bom_Define1      As String
    Dim bom_Define2      As String
    Dim bom_Define3      As String
    Dim bom_Define4      As String
    Dim bom_Define5      As String
    Dim bom_Define6      As String
    Dim bom_Define7      As String
    Dim bom_Define8      As String
    Dim bom_Define9      As String
    Dim bom_Define10     As String
    Dim bom_Define11     As String
    Dim bom_Define12     As String
    Dim bom_Define13     As String
    Dim bom_Define14     As String
    Dim bom_Define15     As String
    Dim bom_Define16     As String
    Dim bom_Define22     As String
    Dim bom_Define23     As String
    Dim bom_Define24     As String
    Dim bom_Define25     As String
    Dim bom_Define26     As String
    Dim bom_Define27     As String
    Dim bom_Define28     As String
    Dim bom_Define29     As String
    Dim bom_Define30     As String
    Dim bom_Define31     As String
    Dim bom_Define32     As String
    Dim bom_Define33     As String
    Dim bom_Define34     As String
    Dim bom_Define35     As String
    Dim bom_Define36     As String
    Dim bom_Define37     As String
    Dim bom_CCode        As String
    Dim bom_cInvCode       As String
    Dim bom_mInvCode       As String
    Dim bom_mQty       As Double
    
    Dim i As Integer
    Dim j As Integer
    
    Dim qty As Double       '订单数量
    Dim unitQty As Double
    
    Dim sql As String

    For j = 0 To domtemp.selectNodes("//z:row").Length - 1
        cls_Public.WrtDBlog con, nvLogin.cUserId, "Interface_fxm 调用CreateBomChenage3 ", "母[cInvCode]=" & cInvCodes & "   子[InvCode]" & GetBodyItemValue1(domtemp, "InvCode", j)
    
    
        unitQty = CDbl(GetBodyItemValue1(domtemp, "BaseQtyN", j)) / CDbl(GetBodyItemValue1(domtemp, "BaseQtyD", j))
        qty = unitQty * iQuantitys
        
        sql = "insert into ef_bomchenage ([id],[autoid],[cSOID],[cinvcode],[unitqty],[minvcode]) values (newid(),"
        sql = sql & soAutoid & ",'" & soCode & "','" & GetBodyItemValue1(domtemp, "InvCode", j) & "',"
        sql = sql & unitQty & ",'" & cInvCodes & "')"
        
    
        domConn.Execute sql
        bom_mInvCode = cInvCodes
        bom_mQty = iQuantitys
        bom_cInvCode = GetBodyItemValue1(domtemp, "InvCode", j)
        bom_CCode = soCode
        bom_OpComponentId = GetBodyItemValue1(domtemp, "OpComponentId", j)
        bom_OpSeq = GetBodyItemValue1(domtemp, "OpSeq", j)
        bom_CompId = GetBodyItemValue1(domtemp, "CompId", j)
        bom_UnitId = GetBodyItemValue1(domtemp, "UnitId", j)
        bom_BaseQtyN = GetBodyItemValue1(domtemp, "BaseQtyN", j)
        bom_BaseQtyD = GetBodyItemValue1(domtemp, "BaseQtyD", j)
        bom_ParentScrap = GetBodyItemValue1(domtemp, "ParentScrap", j)
        bom_CompScrap = GetBodyItemValue1(domtemp, "CompScrap", j)
        If LCase(GetBodyItemValue1(domtemp, "FVQty", j)) = "true" Then
            bom_FVQty = 1
        ElseIf LCase(GetBodyItemValue1(domtemp, "FVQty", j)) = "false" Then
            bom_FVQty = 0
        End If
        bom_Cqty = GetBodyItemValue1(domtemp, "Cqty", j)
        bom_Cqty1 = GetBodyItemValue1(domtemp, "Cqty1", j)
        bom_UseQty = GetBodyItemValue1(domtemp, "UseQty", j)
        bom_Offset = GetBodyItemValue1(domtemp, "Offset", j)
        bom_WIPtype = GetBodyItemValue1(domtemp, "WIPtype", j)
        bom_WhCode = GetBodyItemValue1(domtemp, "WhCode", j)
        bom_InvCode = GetBodyItemValue1(domtemp, "InvCode", j)
        bom_Free1 = GetBodyItemValue1(domtemp, "Free1", j)
        bom_Free2 = GetBodyItemValue1(domtemp, "Free2", j)
        bom_Free3 = GetBodyItemValue1(domtemp, "Free3", j)
        bom_Free4 = GetBodyItemValue1(domtemp, "Free4", j)
        bom_Free5 = GetBodyItemValue1(domtemp, "Free5", j)
        bom_Free6 = GetBodyItemValue1(domtemp, "Free6", j)
        bom_Free7 = GetBodyItemValue1(domtemp, "Free7", j)
        bom_Free8 = GetBodyItemValue1(domtemp, "Free8", j)
        bom_Free9 = GetBodyItemValue1(domtemp, "Free9", j)
        bom_Free10 = GetBodyItemValue1(domtemp, "Free10", j)
        bom_Dept = GetBodyItemValue1(domtemp, "Dept", j)
        bom_DepName = GetBodyItemValue1(domtemp, "DepName", j)
        If LCase(GetBodyItemValue1(domtemp, "ByproductFlag", j)) = "true" Then
            bom_ByproductFlag = 1
        ElseIf LCase(GetBodyItemValue1(domtemp, "ByproductFlag", j)) = "false" Then
            bom_ByproductFlag = 1
        End If
    '                    bom_AccuCostFlag = GetBodyItemValue1(domtemp, "AccuCostFlag ", j)
    '                    bom_SubFlag = GetBodyItemValue1(domtemp, "SubFlag", j)
    '                    bom_BomType = GetBodyItemValue1(domtemp, "BomType", j)
    '                    bom_iGrade = GetBodyItemValue1(domtemp, "iGrade", j)
    '                    bom_DemDate = GetBodyItemValue1(domtemp, "DemDate", j)
    '                    bom_AuxUnitCode = GetBodyItemValue1(domtemp, "AuxUnitCode", j)
    '                    bom_ChangeRate = GetBodyItemValue1(domtemp, "ChangeRate", j)
    '                    bom_AuxBaseQtyN = GetBodyItemValue1(domtemp, "AuxBaseQtyN", j)
    '                    bom_AuxCqty = GetBodyItemValue1(domtemp, "AuxCqty", j)
    '                    bom_AuxUseQty = GetBodyItemValue1(domtemp, "AuxUseQty", j)
    '                    bom_AuxUnitName = GetBodyItemValue1(domtemp, "AuxUnitName", j)
    '                    bom_Define1 = GetBodyItemValue1(domtemp, "Define1", j)
    '                    bom_Define2 = GetBodyItemValue1(domtemp, "Define2", j)
    '                    bom_Define3 = GetBodyItemValue1(domtemp, "Define3", j)
    '                    bom_Define4 = GetBodyItemValue1(domtemp, "Define4", j)
    '                    bom_Define5 = GetBodyItemValue1(domtemp, "Define5", j)
    '                    bom_Define6 = GetBodyItemValue1(domtemp, "Define6", j)
    '                    bom_Define7 = GetBodyItemValue1(domtemp, "Define7", j)
    '                    bom_Define8 = GetBodyItemValue1(domtemp, "Define8", j)
    '                    bom_Define9 = GetBodyItemValue1(domtemp, "Define9", j)
    '                    bom_Define10 = GetBodyItemValue1(domtemp, "Define10", j)
    '                    bom_Define11 = GetBodyItemValue1(domtemp, "Define11", j)
    '                    bom_Define12 = GetBodyItemValue1(domtemp, "Define12", j)
    '                    bom_Define13 = GetBodyItemValue1(domtemp, "Define13", j)
    '                    bom_Define14 = GetBodyItemValue1(domtemp, "Define14", j)
    '                    bom_Define15 = GetBodyItemValue1(domtemp, "Define15", j)
    '                    bom_Define16 = GetBodyItemValue1(domtemp, "Define16", j)
    '                    bom_Define22 = GetBodyItemValue1(domtemp, "Define22", j)
    '                    bom_Define23 = GetBodyItemValue1(domtemp, "Define23", j)
    '                    bom_Define24 = GetBodyItemValue1(domtemp, "Define24", j)
    '                    bom_Define25 = GetBodyItemValue1(domtemp, "Define25", j)
    '                    bom_Define26 = GetBodyItemValue1(domtemp, "Define26", j)
    '                    bom_Define27 = GetBodyItemValue1(domtemp, "Define27", j)
    '                    bom_Define28 = GetBodyItemValue1(domtemp, "Define28", j)
    '                    bom_Define29 = GetBodyItemValue1(domtemp, "Define29", j)
    '                    bom_Define30 = GetBodyItemValue1(domtemp, "Define30", j)
    '                    bom_Define31 = GetBodyItemValue1(domtemp, "Define31", j)
    '                    bom_Define32 = GetBodyItemValue1(domtemp, "Define32", j)
    '                    bom_Define33 = GetBodyItemValue1(domtemp, "Define33", j)
    '                    bom_Define34 = GetBodyItemValue1(domtemp, "Define34", j)
    '                    bom_Define35 = GetBodyItemValue1(domtemp, "Define35", j)
    '                    bom_Define36 = GetBodyItemValue1(domtemp, "Define36", j)
    '                    bom_Define37 = GetBodyItemValue1(domtemp, "Define37", j)
    
    '                    con.Execute "delete ef_bom where invcode = '" & bom_CCode + "'"
    
        sql = "insert into ef_bom(OpComponentId,OpSeq,CompId,UnitId,BaseQtyN,BaseQtyD,ParentScrap,CompScrap,FVQty,Cqty,Cqty1,UseQty,Offset,WIPtype,WhCode,InvCode,Free1,Free2,Free3,Free4,Free5,Free6,Free7,"
        sql = sql & "Free8,Free9,Free10,Dept,DepName,ByproductFlag,AccuCostFlag ,SubFlag,BomType,iGrade,DemDate,AuxUnitCode,ChangeRate,AuxBaseQtyN,AuxCqty,AuxUseQty,AuxUnitName,Define1,Define2,"
        sql = sql & "Define3,Define4,Define5,Define6,Define7,Define8,Define9,Define10,Define11,Define12,Define13,Define14,Define15,Define16,Define22,Define23,Define24,Define25,Define26,Define27,"
        sql = sql & "Define28,Define29,Define30,Define31,Define32,Define33,Define34,Define35,Define36,Define37,ID,cInvCode,autoid,mInvcode,mQty) values ('"
        sql = sql & bom_OpComponentId & "','" & bom_OpSeq & "','" & bom_CompId & "','" & bom_UnitId & "','" & bom_BaseQtyN & "','" & bom_BaseQtyD & "','" & bom_ParentScrap & "','" & bom_CompScrap & "',"
        sql = sql & bom_FVQty & "," & bom_Cqty & "," & bom_Cqty1 & "," & bom_UseQty & "," & bom_Offset & "," & bom_WIPtype & ",'" & bom_WhCode & "','" & bom_InvCode & "','" & bom_Free1 & "','" & bom_Free2 & "','"
        sql = sql & bom_Free3 & "','" & bom_Free4 & "','" & bom_Free5 & "','" & bom_Free6 & "','" & bom_Free7 & "','" & bom_Free8 & "','" & bom_Free9 & "','" & bom_Free10 & "','" & bom_Dept & "','" & bom_DepName & "',"
        sql = sql & bom_ByproductFlag & "," & bom_AccuCostFlag & "," & bom_SubFlag & "," & bom_BomType & "," & bom_iGrade & ",'" & bom_DemDate & "','" & bom_AuxUnitCode & "','" & bom_ChangeRate & "'," & bom_AuxBaseQtyN & ",'"
        sql = sql & bom_AuxCqty & "'," & bom_AuxUseQty & ",'" & bom_AuxUnitName & "','" & bom_Define1 & "','" & bom_Define2 & "','" & bom_Define3 & "','" & bom_Define4 & "','" & bom_Define5 & "','" & bom_Define6 & "','"
        sql = sql & bom_Define7 & "','" & bom_Define8 & "','" & bom_Define9 & "','" & bom_Define10 & "','" & bom_Define11 & "','" & bom_Define12 & "','" & bom_Define13 & "','" & bom_Define14 & "','" & bom_Define15 & "','"
        sql = sql & bom_Define16 & "','" & bom_Define22 & "','" & bom_Define23 & "','" & bom_Define24 & "','" & bom_Define25 & "','" & bom_Define26 & "','" & bom_Define27 & "','" & bom_Define28 & "','" & bom_Define29 & "','"
        sql = sql & bom_Define30 & "','" & bom_Define31 & "','" & bom_Define32 & "','" & bom_Define33 & "','" & bom_Define34 & "','" & bom_Define35 & "','" & bom_Define36 & "','" & bom_Define37 & "','" & bom_CCode & "','" & bom_cInvCode & "',"
        sql = sql & soAutoid & ",'" & bom_mInvCode & "'," & bom_mQty & ")"
        
        cls_Public.WrtDBlog con, nvLogin.cUserId, "Interface_fxm  调用CreateBomChenage 3 ", "写入母子件对应关系！"
        domConn.Execute sql

    Next j
End Function
