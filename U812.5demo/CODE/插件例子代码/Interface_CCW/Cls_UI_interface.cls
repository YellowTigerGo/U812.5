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

Dim g_busines1 As Object
Dim m_login As Object           '当前登陆对象
Dim DBconn As Object            '当前数据库连接对象
Dim StrCardNum As String           '当前单据的CardNum
'Dim Vouchers As UAPVoucherControl85.ctlVoucher
Dim Vouchers As Object          '当前单据对象
Dim Vouchers_Frm As Object      '当前窗体
Dim DomH As New DOMDocument     '
Dim DomB As New DOMDocument     '
'----------------------------------------------
Dim Rss As New ADODB.Recordset  '数据集
Dim ssql As String 'SQL语句
Dim sErr As String '保存错误返回的错误信息
Dim Count As Integer '循环变量
Dim isOC As Boolean '判断行是否有打开的，如果有为true,否则为false，false时则行全部关闭
'----------------------------------------------


'初始化用户插件
'mLogin 当前登陆对象
'DBcon 当前数据库连接对象
'Frm 当前单据的CardNum
'CardNum 当前单据对象
'strUserErr  错误信息
'bSuc  =true 表示初始化成功   =false 表示初始化失败
Public Sub Userdll_Init(g_busines As Object, mLogin As Object, DBcon As Object, Frm As Object, CardNum As String, strUserErr As String, bsuc As Boolean)
    Set g_busines1 = g_busines
    Set m_login = mLogin
    Set DBconn = DBcon
    Set Vouchers_Frm = Frm
    StrCardNum = CardNum
    gstrVoucherType = CardNum '用于在刷新界面的时候增加过滤条件
    bsuc = True
'    MsgBox " Userdll_Init 插件初始化配置成功！"
End Sub

 '初始化
Public Function Init(ByVal objLogin As Object, ByVal objForm As Object, ByVal objVoucher As Object, msbar As Object)
'MsgBox "Init"
End Function

'Voucher   当前单据对象
'ButtonKey 按扭关键字
'strUserErr  错误信息
'bSuc  =true 表示成功并终止系统流程   =false表示失败并继续系统流程
Public Sub Before_ButtonClick(Voucher As Object, ButtonKey As String, strUserErr As String, bsuc As Boolean)
    'Voucher.Row            当前行
    'Voucher.BodyRows       总行数
    '    bSuc = False
    Set Vouchers = Voucher
    
    
'
'     Dim EF As New EFMain.clsProductFacade
' EF.Set_Business g_busines1
' EF.Login m_login
' EF.CallFunction "YXEF9122", "卷膜类", "YXEF912201", vbTab & "12633"
    
    
'    gstrKeyName = ButtonKey
    
    bsuc = False
    Select Case LCase(StrCardNum)
        Case LCase("YXEF9102")  '发货计划单
            Select Case LCase(ButtonKey)
                Case LCase("openline")  '行打开
                    Call SendPaln_OpenColumn(Voucher, strUserErr, bsuc)
                Case LCase("closeline") '行关闭
                    Call SendPlan_CloseColumn(Voucher, strUserErr, bsuc)
                Case LCase("closeorder") '点击表头半闭后，将所有表体行关闭标识置为关闭
                    Call SendPlan_MainClose(Voucher, strUserErr, bsuc)
                    
                Case LCase("reforder")  '处理发货计划参照生单
                    Vouchers.getVoucherDataXML DomH, DomB
                    AddNew_Voucher "YXEF9202A", "YXEF9202", DomH, DomB, strUserErr, bsuc
                    bsuc = True
            End Select
            
        Case LCase("YXEF9103")  '派工计划单
            If LCase(ButtonKey) = LCase("reforder") Then  '处理派工计划参照生单
                Vouchers.getVoucherDataXML DomH, DomB
                AddNew_Voucher_PGD "YXEF9103_ref", "YXEF9103", DomH, DomB, strUserErr, bsuc
                bsuc = True
            End If
        Case LCase("YXEF9114"), LCase("YXEF9115"), LCase("YXEF9116"), LCase("YXEF9117"), LCase("YXEF9118"), LCase("YXEF9119"), LCase("YXEF9120") '正品报工单
            If LCase(ButtonKey) = LCase("reforder") Then  '处理正品报工单参照生产订单
                Vouchers.getVoucherDataXML DomH, DomB
                AddNew_Voucher_ZPBGD "YXEF9115_refjd", "YXEF9115", DomH, DomB, strUserErr, bsuc
                bsuc = True
            End If
            If LCase(ButtonKey) = LCase("reforderp") Then  '处理正品报工单参照派工单
                Vouchers.getVoucherDataXML DomH, DomB
                AddNew_Voucher_ZPBGD "YXEF9115_refpg", "YXEF9115", DomH, DomB, strUserErr, bsuc
                bsuc = True
            End If
            If LCase(StrCardNum) = LCase("YXEF9117") Then  '处理挤复正品报工单填写时，默认将双挤置
               Voucher.headerText("str13") = "否"
            End If
    
            
        Case LCase("YXEF9131"), LCase("YXEF9132"), LCase("YXEF9133"), LCase("YXEF9134"), LCase("YXEF9135"), LCase("YXEF9136"), LCase("YXEF9137") '废品报工单
            If LCase(ButtonKey) = LCase("reforder") Then  '处理废品报工单参照生产订单
                Vouchers.getVoucherDataXML DomH, DomB
                AddNew_Voucher_FPBGD StrCardNum, "YXEF9115_refjd", "YXEF9131", DomH, DomB, strUserErr, bsuc
                bsuc = True
            End If
            If LCase(ButtonKey) = LCase("reforderp") Then  '处理废品报工单参照派工单
                Vouchers.getVoucherDataXML DomH, DomB
                AddNew_Voucher_FPBGD StrCardNum, "YXEF9115_refpg", "YXEF9131", DomH, DomB, strUserErr, bsuc
                bsuc = True
            End If
            If LCase(ButtonKey) = LCase("reforderz") Then   '处理废品报工单参照正品派工单生单
                Vouchers.getVoucherDataXML DomH, DomB
                AddNew_Voucher_FPBGD StrCardNum, "YXEF9131_refbg", "YXEF9131", DomH, DomB, strUserErr, bsuc
                bsuc = True
            End If
    End Select
End Sub

 
'Voucher   当前单据对象
'ButtonKey 按扭关键字
'strUserErr  错误信息
'bSuc  =true 表示成功并终止系统流程   =false表示失败并继续系统流程
Public Sub After_ButtonClick(Voucher As Object, ButtonKey As String, strUserErr As String, bsuc As Boolean)
    bsuc = False
    
'    Set Vouchers = Voucher
'
'    Select Case LCase(StrCardNum)
'        '发货计划单
'        Case LCase("YXEF9102")
'            Select Case LCase(ButtonKey)
'                Case LCase("add")  '处理发货计划参照生单
'                    Vouchers.getVoucherDataXML DomH, DomB
'                    AddNew_Voucher "YXEF9202A", "YXEF9202", DomH, DomB, strUserErr, bSuc
'                    bSuc = True
'            End Select
'
'    End Select
    
    
End Sub

 
'Voucher    当前单据对象
'sKey       当前操作关键字
'sRet       返回值
'strUserErr  错误信息
'bSuc  =true 表示成功并终止系统流程   =false表示失败并继续系统流程
Public Sub Voucher_headBrowUser(ByVal Voucher As Object, ByVal sKey As String, sRet As Variant, strUserErr As String, bsuc As Boolean)

    bsuc = False
End Sub


'初始化用户插件
'Voucher    当前单据对象
'sKey       当前操作关键字
'sRet       返回值
'strUserErr  错误信息
'bChanged
'bSuc  =true 表示成功并终止系统流程   =false表示失败并继续系统流程
Public Sub Voucher_headCellCheck(ByVal Voucher As Object, sKey As String, RetValue As String, bChanged As Long, strUserErr As String, bsuc As Boolean)
    Dim i As Long
    bsuc = False
    
    If LCase(sKey) = LCase("datetime1") And StrCardNum = "YXEF9102" Then
        For i = 1 To Voucher.BodyRows
            Voucher.bodyText(i, "b_datetime1") = Voucher.headerText("datetime1")
        Next i
        bsuc = True
    End If
    
End Sub

 
 
'Voucher    当前单据对象
'sKey       当前操作关键字
'R          当前行数
'sRet       返回值
'strUserErr  错误信息
'bSuc  =true 表示成功并终止系统流程   =false表示失败并继续系统流程
Public Sub Voucher_bodyBrowUser(ByVal Voucher As Object, sKey As String, ByVal R As Long, sRet As Variant, strUserErr As String, bsuc As Boolean)
'    Set Vouchers = Voucher


'        MsgBox " Voucher_bodyBrowUser 插件配置成功！"

    bsuc = False
End Sub


'初始化用户插件
'Voucher    当前单据对象
'sRet       返回值
'sKey       当前操作关键字
'R          当前行数
'strUserErr  错误信息
'bChanged
'bSuc  =true 表示成功并终止系统流程   =false表示失败并继续系统流程
Public Sub Voucher_bodyCellCheck(ByVal Voucher As Object, RetValue As Variant, bChanged As Long, sKey As String, ByVal R As Long, strUserErr As String, bsuc As Boolean)
'    Set Vouchers = Voucher
'    MsgBox " Voucher_bodyCellCheck 插件配置成功！"
    
    bsuc = False
    
    Select Case LCase(StrCardNum)
        Case LCase("YXEF9115"), LCase("YXEF9116"), LCase("YXEF9117"), LCase("YXEF9118"), LCase("YXEF9119"), LCase("YXEF9120") '正品报工单
            Dim i As Long
            
            If LCase(StrCardNum) = LCase("YXEF9119") And LCase(sKey) = LCase("b_float3") Then '制袋正品报工单,计算完工折算米数
                If IsNull(Voucher.headerText("float25")) Or Voucher.headerText("float25") = "" Or Voucher.headerText("float25") = 0 Then
                    bsuc = True
                    Exit Sub
                End If
                Voucher.bodyText(Voucher.row, "b_float6") = Voucher.headerText("float25") * Voucher.bodyText(Voucher.row, "b_float3") / 1000
            End If
            '设置版辊类型默认值
            If LCase(StrCardNum) = LCase("YXEF9115") And LCase(sKey) = LCase("b_float3") Then Voucher.bodyText(Voucher.row, "b_str5") = "凹版"
            
            
            '复制行
            For i = 1 To Voucher.Cols
                If Voucher.bodyText(Voucher.row, i) = "" Then
                    Voucher.bodyText(Voucher.row, i) = Voucher.bodyText(Voucher.row - 1, i)
                End If
            Next
        Case LCase("YXEF9131"), LCase("YXEF9132"), LCase("YXEF9133"), LCase("YXEF9134"), LCase("YXEF9135")  '其他废品报工单
        
            copy_line Voucher
            
            If LCase(sKey) = LCase("b_str4") Then '选择工序性质为本工序，则取出表头工序说明赋值给工序
                If Voucher.bodyText(Voucher.row, "b_str4") = "本工序" Then
                    Voucher.bodyText(Voucher.row, "b_str5") = Voucher.headerText("str3")
                    Voucher.bodyText(Voucher.row, "b_cpersoncode") = Voucher.headerText("t_cpersoncode")
                    Voucher.bodyText(Voucher.row, "b_cpersonname") = Voucher.headerText("t_cpersonname")
                    Voucher.bodyText(Voucher.row, "b_cdepcode") = Voucher.headerText("t_cdepcode")
                    Voucher.bodyText(Voucher.row, "b_cdepname") = Voucher.headerText("t_cdepname")
                    Voucher.bodyText(Voucher.row, "b_str11") = "正常"
                    If LCase(StrCardNum) = LCase("YXEF9131") Then Voucher.bodyText(Voucher.row, "b_str14") = "凹版"
                Else
                    Voucher.bodyText(Voucher.row, "b_str5") = ""
                    Voucher.bodyText(Voucher.row, "b_cpersoncode") = ""
                    Voucher.bodyText(Voucher.row, "b_cpersonname") = ""
                    Voucher.bodyText(Voucher.row, "b_cdepcode") = ""
                    Voucher.bodyText(Voucher.row, "b_cdepname") = ""
                    Voucher.bodyText(Voucher.row, "b_str11") = "正常"
                    If LCase(StrCardNum) = LCase("YXEF9131") Then Voucher.bodyText(Voucher.row, "b_str14") = "凹版"
                End If
            ElseIf LCase(sKey) = LCase("b_float1") Then '填写重量时，自动根据换算率填写相应的数量
                If IsNull(Voucher.headerText("float22")) Or Voucher.headerText("float22") = "" Or Voucher.headerText("float22") = 0 Then
                    bsuc = True
                    Exit Sub
                End If
                Voucher.bodyText(Voucher.row, "b_float3") = Voucher.headerText("float22") * Voucher.bodyText(Voucher.row, "b_float1")
                Voucher.bodyText(Voucher.row, "b_float6") = Voucher.headerText("float22") * Voucher.bodyText(Voucher.row, "b_float1") * Voucher.headerText("float25") / 1000
            ElseIf LCase(sKey) = LCase("b_float3") Then '填写数量时，自动根据换算率填写相应的重量(先注销，还可能启用)
'                If IsNull(Voucher.headerText("float22")) Or Voucher.headerText("float22") = "" Or Voucher.headerText("float22") = 0 Then
'                    bSuc = True
'                    Exit Sub
'                End If
'                Voucher.bodyText(Voucher.row, "b_float1") = Voucher.bodyText(Voucher.row, "b_float3") / Voucher.headerText("float22")
                If LCase(StrCardNum) = LCase("YXEF9135") Then '制袋废品报工单,计算完工折算米数
                    If IsNull(Voucher.headerText("float25")) Or Voucher.headerText("float25") = "" Or Voucher.headerText("float25") = 0 Then
                        bsuc = True
                        Exit Sub
                    End If
                    Voucher.bodyText(Voucher.row, "b_float6") = Voucher.headerText("float25") * Voucher.bodyText(Voucher.row, "b_float3") / 1000
                End If
            End If
        Case LCase("YXEF9136") '包材废品报工单
            copy_line Voucher
            If LCase(sKey) = LCase("b_str4") Then '选择工序性质为本工序，则取出表头工序说明赋值给工序
                If Voucher.bodyText(Voucher.row, "b_str4") = "本工序" Then
                    Voucher.bodyText(Voucher.row, "b_str5") = Voucher.headerText("str3")
                    Voucher.bodyText(Voucher.row, "b_cpersoncode") = Voucher.headerText("t_cpersoncode")
                    Voucher.bodyText(Voucher.row, "b_cpersonname") = Voucher.headerText("t_cpersonname")
                    Voucher.bodyText(Voucher.row, "b_cdepcode") = Voucher.headerText("t_cdepcode")
                    Voucher.bodyText(Voucher.row, "b_cdepname") = Voucher.headerText("t_cdepname")
                    Voucher.bodyText(Voucher.row, "b_str11") = "正常"
                Else
                    Voucher.bodyText(Voucher.row, "b_str5") = ""
                    Voucher.bodyText(Voucher.row, "b_cpersoncode") = ""
                    Voucher.bodyText(Voucher.row, "b_cpersonname") = ""
                    Voucher.bodyText(Voucher.row, "b_cdepcode") = ""
                    Voucher.bodyText(Voucher.row, "b_cdepname") = ""
                    Voucher.bodyText(Voucher.row, "b_str11") = "正常"
                End If
            ElseIf LCase(sKey) = LCase("b_float3") Then '填写重量时，自动根据换算率填写相应的数量
                If IsNull(Voucher.headerText("float22")) Or Voucher.headerText("float22") = "" Or Voucher.headerText("float22") = 0 Then
                    bsuc = True
                    Exit Sub
                End If
                Voucher.bodyText(Voucher.row, "b_float2") = Voucher.headerText("float22") * Voucher.bodyText(Voucher.row, "b_float3")
            ElseIf LCase(sKey) = LCase("b_float2") Then '填写数量时，自动根据换算率填写相应的重量
'                If IsNull(Voucher.headerText("float22")) Or Voucher.headerText("float22") = "" Or Voucher.headerText("float22") = 0 Then
'                    bSuc = True
'                    Exit Sub
'                End If
'                Voucher.bodyText(Voucher.row, "b_float3") = Voucher.bodyText(Voucher.row, "b_float2") / Voucher.headerText("float22")
            End If
        Case Else
        
    End Select
    
'    bSuc = False
End Sub

'初始化用户插件
'Voucher        当前单据对象
'sKey           当前操作关键字
'pCom           下拉控件
'strUserErr     错误信息
'bChanged
'bSuc  =true 表示成功并终止系统流程   =false表示失败并继续系统流程
Public Sub Voucher_FillHeadComboBox(ByVal Voucher As Object, sKey As String, pCom As Object, strUserErr As String, bsuc As Boolean)
    Set Vouchers = Voucher
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
'        MsgBox " Voucher_FillHeadComboBox 插件配置成功！"



    bsuc = False
End Sub
 
 '初始化用户插件
'Voucher        当前单据对象
'sKey           当前操作关键字
'R          当前行数
'pCom           下拉控件
'strUserErr     错误信息
'bChanged
'bSuc  =true 表示成功并终止系统流程   =false表示失败并继续系统流程
Public Sub Voucher_FillList(ByVal Voucher As Object, sKey As String, ByVal R As Long, pCom As Object, strUserErr As String, bsuc As Boolean)
    Set Vouchers = Voucher
    'Dim sFieldName As String
    'sFieldName = LCase(Me.Voucher.ItemState(C, sibody).sFieldName)
    'Select Case sFieldName
    '    Case "adds", "lenssen" '控制方向
    '        pCom.Clear
    '        pCom.AddItem ""
    '        pCom.AddItem "借方"
    '        pCom.AddItem "贷方"
 
    '

         bsuc = False
'        MsgBox " Voucher_FillList 插件配置成功！"
    
End Sub

''
'Source_Cardnumber    来源单句
'Dest_Cardnumber      目标单据
'Dest_domHead         目标单据表头数据
'Dest_domBody         目标单据表体数据
'strUserErr           错误信息
'bSuc               =true 表示成功    =false 表示失败
Public Sub AddNew_Voucher_FHD(Source_Cardnumber As String, Dest_Cardnumber As String, Dest_domHead As DOMDocument, Dest_domBody As DOMDocument, strUserErr As String, bsuc As Boolean)
On Error GoTo ExitSub
    Dim frmR As New frmReferFHD ' frmRefer
    bsuc = False
    With frmR
        Call .Init(m_login, DBconn, frmR, Dest_domHead, Dest_domBody, Source_Cardnumber, Dest_Cardnumber)
        If .filter Then
            frmR.SetData
            frmR.Show 1
            Set Dest_domHead = .DomHead_Dest.cloneNode(True)
            Set Dest_domBody = .DomBody_Dest.cloneNode(True)
            Vouchers.setVoucherDataXML Dest_domHead, Dest_domBody
        End If
    End With
    bsuc = True
    Set frmR = Nothing
    Exit Sub
ExitSub:
    bsuc = False
    strUserErr = strUserErr & Err.Description
End Sub

''
'Source_Cardnumber    来源单句
'Dest_Cardnumber      目标单据
'Dest_domHead         目标单据表头数据
'Dest_domBody         目标单据表体数据
'strUserErr           错误信息
'bSuc               =true 表示成功    =false 表示失败
Public Sub AddNew_Voucher(Source_Cardnumber As String, Dest_Cardnumber As String, Dest_domHead As DOMDocument, Dest_domBody As DOMDocument, strUserErr As String, bsuc As Boolean)
On Error GoTo ExitSub
    Dim frmR As New frmRefernew
    
    bsuc = False
    With frmR
        Call .Init(m_login, DBconn, frmR, Dest_domHead, Dest_domBody, Source_Cardnumber, Dest_Cardnumber)
        If .filter Then
            frmR.SetData
            frmR.Show 1
            Set Dest_domHead = .DomHead_Dest.cloneNode(True)
            Set Dest_domBody = .DomBody_Dest.cloneNode(True)
            Vouchers.setVoucherDataXML Dest_domHead, Dest_domBody
        End If
    End With
    bsuc = True
    Set frmR = Nothing
    Exit Sub
ExitSub:
    bsuc = False
    strUserErr = strUserErr & Err.Description
End Sub
''正品报工单
'Source_Cardnumber    来源单句
'Dest_Cardnumber      目标单据
'Dest_domHead         目标单据表头数据
'Dest_domBody         目标单据表体数据
'strUserErr           错误信息
'bSuc               =true 表示成功    =false 表示失败
Public Sub AddNew_Voucher_ZPBGD(Source_Cardnumber As String, Dest_Cardnumber As String, Dest_domHead As DOMDocument, Dest_domBody As DOMDocument, strUserErr As String, bsuc As Boolean)
On Error GoTo ExitSub
    Dim frmR As New frmReferZPBGD ' frmRefer
    Dim Dom As New DOMDocument

    bsuc = False
    With frmR
        Call .Init(m_login, DBconn, frmR, Dest_domHead, Dest_domBody, Source_Cardnumber, Dest_Cardnumber)
        If .filter Then
            frmR.SetData
            Set Dom = frmR.ctlReferMakeVouch1.GetHeadDom(False)
            If Dom.selectNodes("//z:row").Length = 1 Then
                frmR.ctlReferMakeVouch1.SelectHeadAll
                frmR.ButtonClick "tlbmakevouch"
            Else
                frmR.Show 1
            End If
            
            
            Set Dest_domHead = .DomHead_Dest.cloneNode(True)
            Set Dest_domBody = .DomBody_Dest.cloneNode(True)
            Vouchers.setVoucherDataXML Dest_domHead, Dest_domBody
        End If
    End With
    bsuc = True
    Set frmR = Nothing
    Exit Sub
ExitSub:
    bsuc = False
    strUserErr = strUserErr & Err.Description
End Sub


'废品报工单
'Source_Cardnumber    来源单句
'Dest_Cardnumber      目标单据
'Dest_domHead         目标单据表头数据
'Dest_domBody         目标单据表体数据
'strUserErr           错误信息
'bSuc               =true 表示成功    =false 表示失败
Public Sub AddNew_Voucher_FPBGD(Cardnumber As String, Source_Cardnumber As String, Dest_Cardnumber As String, Dest_domHead As DOMDocument, Dest_domBody As DOMDocument, strUserErr As String, bsuc As Boolean)
On Error GoTo ExitSub
    Dim frmR As New frmReferFPBGD ' frmRefer
    Dim Dom As New DOMDocument

    bsuc = False
    With frmR
        Call .Init(m_login, Cardnumber, DBconn, frmR, Dest_domHead, Dest_domBody, Source_Cardnumber, Dest_Cardnumber)
        If .filter Then
            frmR.SetData
            Set Dom = frmR.ctlReferMakeVouch1.GetHeadDom(False)
            If Dom.selectNodes("//z:row").Length = 1 Then
                frmR.ctlReferMakeVouch1.SelectHeadAll
                frmR.ButtonClick "tlbmakevouch"
            Else
                frmR.Show 1
            End If
            
            
            Set Dest_domHead = .DomHead_Dest.cloneNode(True)
            Set Dest_domBody = .DomBody_Dest.cloneNode(True)
            Vouchers.setVoucherDataXML Dest_domHead, Dest_domBody
        End If
    End With
    bsuc = True
    Set frmR = Nothing
    Exit Sub
ExitSub:
    bsuc = False
    strUserErr = strUserErr & Err.Description
End Sub


''
'Source_Cardnumber    来源单句
'Dest_Cardnumber      目标单据
'Dest_domHead         目标单据表头数据
'Dest_domBody         目标单据表体数据
'strUserErr           错误信息
'bSuc               =true 表示成功    =false 表示失败
Public Sub AddNew_Voucher_PGD(Source_Cardnumber As String, Dest_Cardnumber As String, Dest_domHead As DOMDocument, Dest_domBody As DOMDocument, strUserErr As String, bsuc As Boolean)
On Error GoTo ExitSub
    Dim frmR As New frmReferPGD ' frmRefer
    
    Dim strDepName As String '部门名称
    
    bsuc = False
    With frmR
        Call .Init(m_login, DBconn, frmR, Dest_domHead, Dest_domBody, Source_Cardnumber, Dest_Cardnumber)
        If .filter Then
            frmR.SetData
            frmR.Show 1
            Set Dest_domHead = .DomHead_Dest.cloneNode(True)
            Set Dest_domBody = .DomBody_Dest.cloneNode(True)
            'Vouchers.setVoucherDataXML Dest_domHead, Dest_domBody
            
            '设置表头信息
            Vouchers.setVoucherDataXML Dest_domHead, Dest_domBody
            Vouchers.headerText("t_cdepcode") = deps
            
            ssql = " select cdepname from department where cdepcode = '" & deps & "'"
                
            If Rss.State <> 0 Then Rss.Close
            Rss.CursorLocation = adUseClient
            Rss.Open ssql, DBconn.ConnectionString, 3, 4
            
            If Rss.EOF Then
                bsuc = False
            Else
                strDepName = Rss.Fields("cdepname") '取出部门名称
                Vouchers.headerText("t_cdepname") = strDepName  '把部门名称设置到单据上
            End If
            
            Vouchers.headerText("str1") = banci '把班次设置到单据上
            
            '---------------------把机台、班组信息设置到单据上---------------------------
            Call SetMachineStation(DBconn, strDepName, Vouchers, strUserErr, bsuc) '机台
            Call SetClassGroup(DBconn, strDepName, Vouchers, strUserErr, bsuc) '班组
            '----------------------------------------------------------------------
        End If
    End With
    bsuc = True
    Set frmR = Nothing
    Exit Sub
ExitSub:
    bsuc = False
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
Public Sub Full_Voucher(Source_Cardnumber As String, Dest_Cardnumber As String, Source_domHead As DOMDocument, Source_domBody As DOMDocument, Dest_domHead As DOMDocument, Dest_domBody As DOMDocument, strUserErr As String, Optional bsuc As Boolean = False)
On Error GoTo ExitSub
    bsuc = False

    '    FrmR.ctlVoucher.setVoucherDataXML Dest_domHead, Dest_domBody
    '    bSuc = True
    Exit Sub
ExitSub:
    bsuc = False
    strUserErr = strUserErr & Err.Description
End Sub


Public Sub Before_But_key(Voucher As Object, ButtonKey As String, strUserErr As String, bsuc As Boolean)


End Sub
'发货计划单行打开
Public Sub SendPaln_OpenColumn(Voucher As Object, strUserErr As String, bsuc As Boolean)
    Set Vouchers = Voucher
    ssql = "UPDATE EF_plan_DispatchLists SET b_int4 = 0 WHERE autoid = " & Vouchers.bodyText(Voucher.row, "autoid")
    sErr = Update(ssql, DBconn)
    If sErr <> "" Then
        strUserErr = sErr
        bsuc = False
        Exit Sub
    Else
        Vouchers.bodyText(Voucher.row, "b_int4") = "打开" '回填界面
        '更新数据库
        ssql = "UPDATE EF_plan_DispatchList SET ccloser = null,dcloserdate = null WHERE id = " & Vouchers.headerText("id")
        sErr = Update(ssql, DBconn)
        If sErr <> "" Then
            strUserErr = sErr
            bsuc = False
            Exit Sub
        Else '回填界面
            '如打开，将表头关闭人、关闭日期清空
            Vouchers.headerText("ccloser") = ""
            Vouchers.headerText("dcloserdate") = ""
        End If
    End If
    
'    Vouchers_Frm.ButtonClick "refresh", ""
End Sub

'发货计划单行关闭
Public Sub SendPlan_CloseColumn(Voucher As Object, strUserErr As String, bsuc As Boolean)
    Set Vouchers = Voucher
    
    ssql = "UPDATE EF_plan_DispatchLists SET b_int4 = 1 WHERE autoid = " & Vouchers.bodyText(Voucher.row, "autoid")
    sErr = Update(ssql, DBconn)
    If sErr <> "" Then
        strUserErr = sErr
        bsuc = False
        Exit Sub
    Else
        Vouchers.bodyText(Voucher.row, "b_int4") = "关闭" '回填界面

        '判断表体行是否都关闭
        isOC = False
        For Count = 1 To Voucher.BodyRows
            If Vouchers.bodyText(Count, "b_int4") = "打开" Then
                isOC = True
                Exit For
            End If
        Next Count
        
        '如果全部关闭，则将主表关闭
        
        If isOC = False Then
            ssql = "UPDATE EF_plan_DispatchList SET ccloser = '" & m_login.cUserName & _
            "',dcloserdate = '" & m_login.CurDate & "' WHERE id = " & Vouchers.headerText("id")
            sErr = Update(ssql, DBconn)
            If sErr <> "" Then
                strUserErr = sErr
                bsuc = False
                Exit Sub
            Else '回填界面
                Vouchers.headerText("ccloser") = m_login.cUserName
                Vouchers.headerText("dcloserdate") = m_login.CurDate
            End If
        End If
    End If

'    Vouchers_Frm.ButtonClick "refresh", ""
End Sub

'发货计划单表头关闭
Public Sub SendPlan_MainClose(Voucher As Object, strUserErr As String, bsuc As Boolean)
    Set Vouchers = Voucher
    
    ssql = "UPDATE EF_plan_DispatchList SET ccloser = '" & m_login.cUserName & _
            "',dcloserdate = '" & m_login.CurDate & "' WHERE id = " & Vouchers.headerText("id")
    sErr = Update(ssql, DBconn)
    If sErr <> "" Then
        strUserErr = sErr
        bsuc = False
        Exit Sub
    Else
        '回填界面
        Vouchers.headerText("ccloser") = m_login.cUserName
        Vouchers.headerText("dcloserdate") = m_login.CurDate
        
        '将表体所有行关闭提交到数据库
        ssql = "UPDATE EF_plan_DispatchLists SET b_int4 = 1 WHERE id = " & Vouchers.headerText("id")
        sErr = Update(ssql, DBconn)
        If sErr <> "" Then
            strUserErr = sErr
            bsuc = False
            Exit Sub
        Else

            '将界面表体行全部设置为关闭
            For Count = 1 To Voucher.BodyRows
                 Vouchers.bodyText(Count, "b_int4") = "关闭"
            Next Count
        
        End If
    End If

'    Vouchers_Frm.ButtonClick "refresh", ""
End Sub

Public Function RunCommand(ByVal objLogin As Object, ByVal objForm As Object, ByVal objVoucher As Object, _
                            ByVal sKey As String, ByVal VarentValue As Variant, ByVal other As String)
                            
                            
Dim strUserErr As String
Dim bsuc As Boolean
Set Vouchers = objVoucher
Set m_login = objLogin
Set DBconn = New ADODB.Connection
DBconn.Open objLogin.UfDbName
'
'    Vouchers.getVoucherDataXML DomH, DomB
''
'    DomH.Save ("c:\DomH.xml")
'    DomB.Save ("c:\DomB.xml")
''
'    If MsgBox("aaaaaaa", vbYesNo) = vbNo Then
'        Exit Function
'    End If
 
    Select Case LCase(sKey)
        '发货参照发货计划单
        Case LCase("_CUSTDEFbtnRefer")
            Vouchers.getVoucherDataXML DomH, DomB
            AddNew_Voucher_FHD "YXEF9203A", "01", DomH, DomB, strUserErr, bsuc
 
        Case Else
        
    End Select
End Function
                                                              

'复制行
Public Sub copy_line(ByVal Voucher As Object)

'Dim i As Long
'    For i = 1 To Voucher.Cols
'        Debug.Print i & " ----------" & Voucher.ItemState(i, sibody).sFieldName & " =" & Voucher.bodyText(Voucher.row, i)
'        If Voucher.bodyText(Voucher.row, i) = "" Then
'            Voucher.bodyText(Voucher.row, i) = Voucher.bodyText(Voucher.row - 1, i)
'        End If
'    Next

If Voucher.BodyRows > 1 Then
    If Voucher.bodyText(Voucher.row, "b_str4") = "" Then
        Voucher.bodyText(Voucher.row, "b_str4") = Voucher.bodyText(Voucher.row - 1, "b_str4")
    End If
    If Voucher.bodyText(Voucher.row, "b_str5") = "" Then
        Voucher.bodyText(Voucher.row, "b_str5") = Voucher.bodyText(Voucher.row - 1, "b_str5")
    End If
    If Voucher.bodyText(Voucher.row, "b_str11") = "" Then
        Voucher.bodyText(Voucher.row, "b_str11") = Voucher.bodyText(Voucher.row - 1, "b_str11")
    End If
    If Voucher.bodyText(Voucher.row, "b_str13") = "" Then
        Voucher.bodyText(Voucher.row, "b_str13") = Voucher.bodyText(Voucher.row - 1, "b_str13")
    End If
    If Voucher.bodyText(Voucher.row, "b_float6") = "" Then
        Voucher.bodyText(Voucher.row, "b_float6") = Voucher.bodyText(Voucher.row - 1, "b_float6")
    End If
    If Voucher.bodyText(Voucher.row, "b_float7") = "" Then
        Voucher.bodyText(Voucher.row, "b_float7") = Voucher.bodyText(Voucher.row - 1, "b_float7")
    End If
    If Voucher.bodyText(Voucher.row, "b_float1") = "" Then
        Voucher.bodyText(Voucher.row, "b_float1") = Voucher.bodyText(Voucher.row - 1, "b_float1")
    End If
    If Voucher.bodyText(Voucher.row, "b_float3") = "" Then
        Voucher.bodyText(Voucher.row, "b_float3") = Voucher.bodyText(Voucher.row - 1, "b_float3")
    End If
    If Voucher.bodyText(Voucher.row, "b_cdepcode") = "" Then
        Voucher.bodyText(Voucher.row, "b_cdepcode") = Voucher.bodyText(Voucher.row - 1, "b_cdepcode")
    End If
    If Voucher.bodyText(Voucher.row, "b_cdepname") = "" Then
        Voucher.bodyText(Voucher.row, "b_cdepname") = Voucher.bodyText(Voucher.row - 1, "b_cdepname")
    End If
    If Voucher.bodyText(Voucher.row, "b_cpersoncode") = "" Then
        Voucher.bodyText(Voucher.row, "b_cpersoncode") = Voucher.bodyText(Voucher.row - 1, "b_cpersoncode")
    End If
    If Voucher.bodyText(Voucher.row, "b_cpersonname") = "" Then
        Voucher.bodyText(Voucher.row, "b_cpersonname") = Voucher.bodyText(Voucher.row - 1, "b_cpersonname")
    End If
    If Voucher.bodyText(Voucher.row, "b_str9") = "" Then
        Voucher.bodyText(Voucher.row, "b_str9") = Voucher.bodyText(Voucher.row - 1, "b_str9")
    End If
    If Voucher.bodyText(Voucher.row, "b_float2") = "" Then
        Voucher.bodyText(Voucher.row, "b_float2") = Voucher.bodyText(Voucher.row - 1, "b_float2")
    End If

End If


End Sub

Private Sub Class_Initialize()


    Set cls_Public = CreateObject("UF_Public_base.cls_log")
End Sub
