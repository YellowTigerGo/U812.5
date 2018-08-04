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

'Dim m_login As Object           '当前登陆对象
'Dim DBconn As Object            '当前数据库连接对象
Dim StrCardNum As String           '当前单据的CardNum
'Dim Vouchers As UAPVoucherControl85.ctlVoucher
Dim Vouchers As Object          '当前单据对象
Dim Vouchers_Frm As Object      '当前窗体
Dim DomH As New DOMDocument     '
Dim DomB As New DOMDocument     '
Dim g_busines1 As Object

' '初始化
'Public Function Init(ByVal objLogin As Object, ByVal objForm As Object, ByVal objVoucher As Object, msbar As Object)
''MsgBox "Init"
'End Function

'单据数据装载后事件
Public Sub LoadAfter_VoucherData(strUserErr As String, bSuc As Boolean)
 On Error GoTo Errhandle
     Dim i As Integer
     Vouchers_Frm.picSure.Visible = True
     
     '注意：如果没有任何的用户自定义动态库，则需要返回TRUE
'     bSuc = True
ExitSub:
     Exit Sub
Errhandle:
     bSuc = False
     strUserErr = Err.Description
End Sub


'初始化用户插件
'mLogin 当前登陆对象
'DBcon 当前数据库连接对象
'Frm 当前单据的CardNum
'CardNum 当前单据对象
'strUserErr  错误信息
'bSuc  =true 表示初始化成功   =false 表示初始化失败
Public Sub Userdll_Init(g_busines As Object, mLogin As Object, DBcon As Object, Frm As Object, CardNum As String, strUserErr As String, bSuc As Boolean)
    Set g_busines1 = g_busines
    Set g_oLogin = mLogin
    Set DBconn = DBcon
    Set Vouchers_Frm = Frm
    If Frm.Name = "frmVouchNew" Then
        Set Vouchers = Vouchers_Frm.Voucher
    Else
        Set voucherForm = Frm
    End If
    StrCardNum = CardNum
    bSuc = True
'    MsgBox " Userdll_Init 插件初始化配置成功！"
End Sub

 
'Voucher   当前单据对象
'ButtonKey 按扭关键字
'strUserErr  错误信息
'bSuc  =true 表示成功并终止系统流程   =false表示失败并继续系统流程
Public Sub Before_ButtonClick(Voucher As Object, ButtonKey As String, strUserErr As String, bSuc As Boolean)
    Dim id As String
    Dim strsql As String
    Dim i As Long
    On Error GoTo hErr
    bSuc = False
    Select Case StrCardNum
        Case "EFJHSC0504"
            Select Case ButtonKey
                Case "doAction"
                    DoRouting
                Case "tlbdelete"
                    DoDelete Voucher
            End Select
        Case "EFJHSC0505"
            Select Case ButtonKey
                Case "doAction"
                    DoQualityCheck Voucher
                Case "tlbdelete"
                    DoDeleteQC Voucher
            End Select
    End Select
        
    Exit Sub
hErr:
    MsgBox Err.Description, vbCritical, "提示"
End Sub

Private Function DoDelete(VchLst As Object) As Boolean
    Dim i As Long
    Dim id As String
    If MsgBox("确认是否删除", vbOKCancel, "提示") = vbCancel Then
        Exit Function
    End If
    For i = 1 To VchLst.rows
        If VchLst.TextMatrix(i, VchLst.GridColIndex("selcol")) = "Y" Then
            If VchLst.TextMatrix(i, VchLst.GridColIndex("cqcid")) <> "" Then
                MsgBox "第" & i & "行记录已检验，不能删除。", vbInformation, "提示"
                Exit For
            End If
            id = VchLst.TextMatrix(i, VchLst.GridColIndex("id"))
            DBconn.Execute "delete EF_Routing where id='" & id & "'"
        End If
    Next
    voucherForm.UFToolbar1.FireSysCommand enumButton, "Refresh"
End Function

'删除检验
Private Function DoDeleteQC(VchLst As Object) As Boolean
    Dim i As Long
    Dim id As String
    If MsgBox("确认是否清除检验信息", vbOKCancel, "提示") = vbCancel Then
        Exit Function
    End If
    For i = 1 To VchLst.rows
        If VchLst.TextMatrix(i, VchLst.GridColIndex("selcol")) = "Y" Then
            If VchLst.TextMatrix(i, VchLst.GridColIndex("iinqty")) <> "" Then
                MsgBox "第" & i & "行记录已有入库记录，不能清除。", vbInformation, "提示"
                Exit For
            End If
            id = VchLst.TextMatrix(i, VchLst.GridColIndex("cqcid"))
            DBconn.Execute "delete EF_QualityCheck where id='" & id & "'"
            DBconn.Execute "update EF_Routing set cqcid=null where cqcid='" & id & "'"
        End If
    Next
    voucherForm.UFToolbar1.FireSysCommand enumButton, "Refresh"
End Function

'报工检验
Private Sub DoQualityCheck(VchLst As Object)
    Dim i As Long
    Dim id As String
    For i = 1 To VchLst.rows
        If VchLst.TextMatrix(i, VchLst.GridColIndex("selcol")) = "Y" Then
            id = VchLst.TextMatrix(i, VchLst.GridColIndex("wxid"))
            Exit For
        End If
    Next
    If id <> "" Then
        frmQualityCheck.id = id
        frmQualityCheck.formStatus = 1
        frmQualityCheck.Show 1
        voucherForm.UFToolbar1.FireSysCommand enumButton, "Refresh"
    End If
End Sub

Private Sub DoRouting()
    frmRouting.formStatus = 1
    frmRouting.Show 1
    voucherForm.UFToolbar1.FireSysCommand enumButton, "Refresh"
End Sub

Private Function GetVouchID(strTableName As String, lngIDs As String, lngsTableCount As Long, errMsg As String) As String
    Dim AdoComm As ADODB.Command
    On Error GoTo DoErr
    Set AdoComm = New ADODB.Command
    With AdoComm
        .ActiveConnection = DBconn
        .CommandText = "sp_GetID"
        .CommandType = adCmdStoredProc
        .Prepared = False
        .Parameters.Append .CreateParameter("RemoteId", adVarChar, adParamInput, 3, "00")
        .Parameters.Append .CreateParameter("cAcc_Id", adVarChar, adParamInput, 3, g_oLogin.cAcc_Id)
        .Parameters.Append .CreateParameter("VouchType", adVarChar, adParamInput, 50, strTableName)
        .Parameters.Append .CreateParameter("iAmount", adInteger, adParamInput, 8, lngsTableCount)
        .Parameters.Append .CreateParameter("MaxID", adBigInt, adParamOutput)
        .Parameters.Append .CreateParameter("MaxIDs", adBigInt, adParamOutput)
        .Execute
        GetVouchID = CStr(.Parameters("MaxID"))
        lngIDs = .Parameters("MaxIDs") - lngsTableCount + 1
    End With
    Set AdoComm = Nothing
    Exit Function
DoErr:
    errMsg = "获取单据ID发生错误：" & Err.Description
    Set AdoComm = Nothing
End Function

'Voucher   当前单据对象
'ButtonKey 按扭关键字
'strUserErr  错误信息
'bSuc  =true 表示成功并终止系统流程   =false表示失败并继续系统流程
Public Sub After_ButtonClick(Voucher As Object, ButtonKey As String, strUserErr As String, bSuc As Boolean)
    bSuc = False
    Select Case ButtonKey
        Case "refervouch"
            If StrCardNum = "PD010401" Then
                Vouchers_Frm.Voucher.getVoucherDataXML DomH, DomB
                AddNew_Voucher "PD010301", "PD010401", DomH, DomB, strUserErr, bSuc
            ElseIf StrCardNum = "PD010501" Then
                Vouchers_Frm.Voucher.getVoucherDataXML DomH, DomB
                AddNew_Voucher "PD010301", "PD010501", DomH, DomB, strUserErr, bSuc
            End If
        Case "modify"
            If StrCardNum = "EP0205" Then
                Voucher.ItemState("cpmrpcode", 0).bCanModify = False
            End If
        Case Else
'            If StrCardNum = "EP0205" Then
'                Voucher.ItemState("cpmrpcode", 0).bCanModify = True
'            End If
    End Select
    
End Sub

'关闭、打开
Private Function DoClose(bClose As Boolean, VchLst As Object) As Boolean
    Dim i As Long
    Dim strsql As String
    Dim id As String
    On Error GoTo hErr
    For i = 1 To VchLst.rows - 1
        If VchLst.TextMatrix(i, VchLst.GridColIndex("selcol")) = "Y" Then
            id = VchLst.TextMatrix(i, VchLst.GridColIndex("id"))
            If bClose Then
                strsql = "update hzland_ChangedTask set status='关闭' where id='" & id & "' and isnull(status,'')=''"
            Else
                strsql = "update hzland_ChangedTask set status='' where id='" & id & "' and status='关闭'"
            End If
            DBconn.Execute strsql
        End If
    Next
    DoClose = True
    Exit Function
hErr:
    MsgBox Err.Description, vbCritical, "提示"
End Function

'Voucher    当前单据对象
'sKey       当前操作关键字
'sRet       返回值
'strUserErr  错误信息
'bSuc  =true 表示成功并终止系统流程   =false表示失败并继续系统流程
Public Sub Voucher_headBrowUser(ByVal Voucher As Object, ByVal Skey As String, sRet As Variant, strUserErr As String, bSuc As Boolean)

    bSuc = False
End Sub


'初始化用户插件
'Voucher    当前单据对象
'sKey       当前操作关键字
'sRet       返回值
'strUserErr  错误信息
'bChanged
'bSuc  =true 表示成功并终止系统流程   =false表示失败并继续系统流程
Public Sub Voucher_headCellCheck(ByVal Voucher As Object, Skey As String, RetValue As String, bChanged As Long, strUserErr As String, bSuc As Boolean)
    Dim strsql As String
    Dim rs As ADODB.Recordset
    Dim i As Long
    On Error GoTo hErr
    Select Case StrCardNum
        Case "EP0207"
            Select Case Skey
                Case "citemcode"
                    '清空表体
                    For i = Voucher.rows To 1 Step -1
                        Voucher.DelLine i
                    Next
                    strsql = "select SUM(p.iSum) as isum,i.cInvDefine2 from PurBillVouchs p " & _
                            " left outer join inventory i on p.cInvCode=i.cInvCode " & _
                            " where p.citemcode='" & RetValue & "' group by i.cInvDefine2 "
                    Set rs = New ADODB.Recordset
                    rs.Open strsql, DBconn
                    i = 1
                    While Not rs.EOF
                        Voucher.bodyText(i, "cmaterialclass") = rs!cInvDefine2 & ""
                        Voucher.bodyText(i, "icb") = str2Dbl(rs!isum & "")
                        Voucher.AddLine
                        rs.MoveNext
                        i = i + 1
                    Wend
                    Voucher.RemoveEmptyRow
            End Select
        
    End Select
    
    bSuc = False
    GoTo Finish
hErr:
    strUserErr = Err.Description
    bSuc = True
Finish:
    If Not rs Is Nothing Then
        If rs.State <> 0 Then
            rs.Close
        End If
        Set rs = Nothing
    End If
End Sub

 
'Voucher    当前单据对象
'sKey       当前操作关键字
'R          当前行数
'sRet       返回值
'strUserErr  错误信息
'bSuc  =true 表示成功并终止系统流程   =false表示失败并继续系统流程
Public Sub Voucher_bodyBrowUser(ByVal Voucher As Object, Skey As String, ByVal R As Long, sRet As Variant, strUserErr As String, bSuc As Boolean)
    bSuc = False

'    Set Vouchers = Voucher


'        MsgBox " Voucher_bodyBrowUser 插件配置成功！"

 
End Sub


'初始化用户插件
'Voucher    当前单据对象
'sRet       返回值
'sKey       当前操作关键字
'R          当前行数
'strUserErr  错误信息
'bChanged
'bSuc  =true 表示成功并终止系统流程   =false表示失败并继续系统流程
Public Sub Voucher_bodyCellCheck(ByVal Voucher As Object, RetValue As Variant, bChanged As Long, Skey As String, ByVal R As Long, strUserErr As String, bSuc As Boolean)
    bSuc = False
    Select Case Skey
        Case "iunitqty"
            If RetValue = "" Then
                Voucher.bodyText(R, "iqty") = ""
            Else
                Voucher.bodyText(R, "iqty") = str2Dbl(Voucher.headerText("ipqty")) * str2Dbl(RetValue)
            End If
    End Select
    

End Sub

'初始化用户插件
'Voucher        当前单据对象
'sKey           当前操作关键字
'pCom           下拉控件
'strUserErr     错误信息
'bChanged
'bSuc  =true 表示成功并终止系统流程   =false表示失败并继续系统流程
Public Sub Voucher_FillHeadComboBox(ByVal Voucher As Object, Skey As String, pCom As Object, strUserErr As String, bSuc As Boolean)
    bSuc = False
    
    
    Set Vouchers = Voucher
'    Dim i As Integer
'    Dim rds As New ADODB.Recordset
'    Select Case LCase(Me.Voucher.ItemState(Index, siheader).sFieldName)
'
'        Case "iyear" '编制年度
'                pCom.Clear
'                For i = 0 To 4
'                  pCom.AddItem CStr(Val(g_oLogin.cIYear) + i)
'                Next

'    End Select
'        MsgBox " Voucher_FillHeadComboBox 插件配置成功！"




End Sub
 
 '初始化用户插件
'Voucher        当前单据对象
'sKey           当前操作关键字
'R          当前行数
'pCom           下拉控件
'strUserErr     错误信息
'bChanged
'bSuc  =true 表示成功并终止系统流程   =false表示失败并继续系统流程
Public Sub Voucher_FillList(ByVal Voucher As Object, Skey As String, ByVal R As Long, pCom As Object, strUserErr As String, bSuc As Boolean)
    bSuc = False
    
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


'        MsgBox " Voucher_FillList 插件配置成功！"
    
End Sub



''
'Source_Cardnumber    来源单句
'Dest_Cardnumber      目标单据
'Dest_domHead         目标单据表头数据
'Dest_domBody         目标单据表体数据
'strUserErr           错误信息
'bSuc               =true 表示成功    =false 表示失败
Public Sub AddNew_Voucher(Source_Cardnumber As String, Dest_Cardnumber As String, Dest_domHead As DOMDocument, Dest_domBody As DOMDocument, strUserErr As String, bSuc As Boolean)
On Error GoTo ExitSub
    Dim FrmR As New frmRefernew ' frmRefer
    bSuc = False
    With FrmR
        Set FrmR.Dest_Voucher = Vouchers
        Call .init(g_oLogin, DBconn, FrmR, Dest_domHead, Dest_domBody, Source_Cardnumber, Dest_Cardnumber)
        If .filter Then
            FrmR.SetData
            FrmR.Show 1
'            Set Dest_domHead = .DomHead_Dest.cloneNode(True)
'            Set Dest_domBody = .DomBody_Dest.cloneNode(True)
'            Vouchers_Frm.Voucher.setVoucherDataXML Dest_domHead, Dest_domBody
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
Public Sub Full_Voucher(Voucher As Object, Source_Cardnumber As String, Dest_Cardnumber As String, Source_domHead As DOMDocument, Source_domBody As DOMDocument, Dest_domHead As DOMDocument, Dest_domBody As DOMDocument, strUserErr As String, Optional bSuc As Boolean = False)
    Dim list As IXMLDOMNodeList
    Dim ele As IXMLDOMElement
    Dim iRow As Integer
    On Error GoTo ExitSub
    bSuc = False
    '处理表头填充
'    Voucher.headerText("citemcode") = GetHeadItemValue(Source_domHead, "projectnumber")
'    Voucher.headerText("citemname") = GetHeadItemValue(Source_domHead, "projectname")
    
    
    '清除当前单据表体数据
    '处理表体填充
    Set list = Source_domHead.selectNodes("//z:row")
    Voucher.RemoveEmptyRow
    iRow = Voucher.BodyRows + 1
    For Each ele In list
        Voucher.AddLine Voucher.BodyRows + 1
        Voucher.SimulateInput SectionsConstants.sibody, Voucher.BodyRows, "b_cinvcode", GetNodeValue(ele, "stockcode"), 0
        If Voucher.bodyText(iRow, "b_cinvcode") <> "" Then
    '        Voucher.bodyText(iRow, "b_cinvcode") = GetNodeValue(ele, "stockcode")
    '        Voucher.bodyText(iRow, "b_cinvname") = GetNodeValue(ele, "stockname")
    '        Voucher.bodyText(iRow, "b_cinvstd") = GetNodeValue(ele, "specification")
    '        Voucher.bodyText(iRow, "b_ccomunitname") = GetNodeValue(ele, "unit")
            Voucher.bodyText(iRow, "iqty") = GetNodeValue(ele, "iwsendqty")
            Voucher.bodyText(iRow, "hzid") = GetNodeValue(ele, "id")
            Voucher.bodyText(iRow, "hzbilltype") = GetNodeValue(ele, "billtype")
            Voucher.bodyText(iRow, "b_citemcode") = GetNodeValue(ele, "projectnumber")
            Voucher.bodyText(iRow, "b_citemname") = GetNodeValue(ele, "projectname")
            Voucher.bodyText(iRow, "cabinetnumber") = GetNodeValue(ele, "cabinetnumber")
            Voucher.bodyText(iRow, "cabinetname") = GetNodeValue(ele, "cabinetname")
            iRow = iRow + 1
        End If
    Next
    bSuc = True
    Exit Sub
ExitSub:
    bSuc = False
    strUserErr = strUserErr & Err.Description
End Sub


 '提示：单据单击事件
Public Sub Voucher_Click(ByVal Voucher As Object, section As Variant, Skey As String, R As Long, strUserErr As String, bSuc As Boolean)
On Error GoTo Errhandle

'     Dim i As Integer
'     bSuc = False
'     For i = 0 To colInterface.Count - 1
'         '判断用户DLL是否有这个方法
'         strTemp = colInterface.Keys(i)
'         If bHasFunction("Voucher_Click", strTemp) Then
'            Set clsInterface = colInterface.Item(strTemp)
'            clsInterface.Voucher_Click Voucher, section, Skey, R, strUserErr, bSuc
'            If bSuc = True Then
'                GoTo ExitSub
'            End If
'         End If
'     Next i
'     '注意：如果没有任何的用户自定义动态库，则需要返回TRUE
''     bSuc = True
''    bSuc = False
ExitSub:
     Exit Sub
Errhandle:
     bSuc = False
     strUserErr = Err.Description
End Sub


 '提示：单据单击事件
Public Sub Voucher_DblClick(ByVal Voucher As Object, section As Variant, Skey As String, R As Long, strUserErr As String, bSuc As Boolean)
On Error GoTo Errhandle

'     Dim i As Integer
'     bSuc = False
'     For i = 0 To colInterface.Count - 1
'         '判断用户DLL是否有这个方法
'         strTemp = colInterface.Keys(i)
'         If bHasFunction("Voucher_DblClick", strTemp) Then
'            Set clsInterface = colInterface.Item(strTemp)
'            clsInterface.Voucher_DblClick Voucher, section, Skey, R, strUserErr, bSuc
'            If bSuc = True Then
'                GoTo ExitSub
'            End If
'         End If
'     Next i
'     '注意：如果没有任何的用户自定义动态库，则需要返回TRUE
''     bSuc = True
''    bSuc = False
ExitSub:
     Exit Sub
Errhandle:
     bSuc = False
     strUserErr = Err.Description
End Sub

 '提示：单据列表双击事件
Public Sub VchLst_DblClick(ByVal VchLst As Object, strUserErr As String, bSuc As Boolean)
On Error GoTo Errhandle
    Dim id As String
    Select Case StrCardNum
        Case "EFJHSC0504"
            id = VchLst.TextMatrix(VchLst.row, VchLst.GridColIndex("wxid"))
            frmRouting.id = id
            frmRouting.formStatus = 2
            frmRouting.Show 1
            voucherForm.UFToolbar1.FireSysCommand enumButton, "Refresh"
        Case "EFJHSC0505"
            id = VchLst.TextMatrix(VchLst.row, VchLst.GridColIndex("cqcid"))
            If id = "" Then
                id = VchLst.TextMatrix(VchLst.row, VchLst.GridColIndex("wxid"))
                frmQualityCheck.id = id
                frmQualityCheck.formStatus = 1
                frmQualityCheck.Show 1
                voucherForm.UFToolbar1.FireSysCommand enumButton, "Refresh"
            Else
                frmQualityCheck.id = id
                frmQualityCheck.formStatus = 2
                frmQualityCheck.Show 1
                voucherForm.UFToolbar1.FireSysCommand enumButton, "Refresh"
            End If
            
    End Select
    
    bSuc = True
ExitSub:
     Exit Sub
Errhandle:
     bSuc = False
     strUserErr = Err.Description
End Sub
