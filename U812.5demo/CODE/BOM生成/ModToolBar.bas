Attribute VB_Name = "ModToolBar"
Option Explicit

'u8常用快捷键

'增加F5
'修改F8
'删除DEL
'审核Ctrl+U
'弃审ALT+U
'保存F6
'放弃Ctrl+Z
'增行Ctrl+N
'删行Ctrl+D
'打印Ctrl+P
'预览Ctrl+V
'输出Alt+E
'刷新Ctrl+R
'定位Ctrl+F3
'首页Alt+PageUp
'上一页PageUp
'下一页PageDown
'末页Alt+PageDown
'生单Ctrl+G
'关闭Alt+C
'打开Alt+O

Public varArgs() As Variant
'工具栏按钮关键字
Public Const sKey_Batchprint = "PrintBatch"         '批打
Public Const sKey_Print = "Print"                   '打印
Public Const sKey_Preview = "Preview"               '预览
Public Const sKey_Output = "Output"                 '输出
Public Const sKey_Export = "Export"                 '输出
Public Const sKey_Copy = "Copy"                     '复制
Public Const sKey_Add = "Add"                       '增加
Public Const sKey_Modify = "Modify"                 '修改
Public Const sKey_Delete = "Delete"                 '删除
Public Const sKey_Save = "Save"                     '保存
Public Const sKey_Discard = "Discard"               '放弃
Public Const sKey_Addrecord = "AddRecord"           '增行
Public Const sKey_InsertRecord = "InsertRecord"     '插行
Public Const sKey_Deleterecord = "DeleteRecord"     '删行
Public Const sKey_AlterPO = "AlterPO"               '变更
Public Const skey_ReportCheck = "ReportCheck"       '报检
Public Const sKey_Close = "Close"                   '关闭
Public Const sKey_Open = "Open"                     '打开
Public Const sKey_Confirm = "Confirm"               '审核
Public Const sKey_Cancelconfirm = "Cancelconfirm"   '弃审
Public Const sKey_Return = "Return"                 '归还
Public Const sKey_BatchReturn = "BatchReturn"       '批量归还
Public Const sKey_QueryConfirm = "QueryConfirm"     '查审(审批流查询)
Public Const sKey_Payment = "Payment"               '现付
Public Const sKey_Cancelpayment = "CancelPayment"   '弃付
Public Const sKey_BatchBV = "BatchBV"               '批量生成发票
Public Const sKey_Settle = "Settle"                 '结算
Public Const sKey_Locate = "Locate"                 '定位
Public Const sKey_LocateSet = "LocateSet"           '定位设置
Public Const sKey_Load = "Load"                     '调入
Public Const sKey_First = "First"                   '首张
Public Const sKey_Previous = "Previous"             '上张
Public Const sKey_Next = "Next"                     '下张
Public Const sKey_Last = "Last"                     '末张
Public Const sKey_RefVoucher = "RefVoucher"         '关联单据 by zhangwchb 20110718
Public Const sKey_Refresh = "Refresh"               '刷新
Public Const sKey_Help = "Help"                     '帮助
Public Const sKey_Exit = "Exit"                     '退出
Public Const sKey_Lock = "Lock"                     '锁定
Public Const sKey_RLock = "removelock"              '解锁
Public Const sKey_Acc = "Accessories"               '附件
Public Const sKey_Link = "Link"                     '联查
Public Const sKey_Column = "Column"                 '栏目
Public Const sKey_Fetchprice = "Fetchprice"         '取价
'工作流按钮 chenliangc
Public Const sKey_Submit = "Submit"                 '提交
Public Const sKey_Unsubmit = "Unsubmit"             '撤销
Public Const sKey_Resubmit = "Resubmit"             '重新提交
Public Const sKey_ViewVerify = "ViewVerify"         '查审
'生单
Public Const sKey_ReferVoucher = "RererVoucher"     '生单(一对一)
Public Const sKey_ReferVouchers = "RererVouchers"   '生单(多对一)
Public Const sKey_CreateVoucher = "CreateVoucher"       '推单
Public Const sKey_CreateSAVoucher = "CreateSAVoucher"       '推销售单
Public Const sKey_CreatePUVoucher = "CreatePUVoucher"       '推采购单
Public Const sKey_CreateSCVoucher = "CreateSCVoucher"       '推库存单
Public Const sKey_CreateAPVoucher = "CreateAPVoucher"       '推应付单

'反选
Public Const sKey_ReverseSelection = "ReverseSelection"       '反选
Public Const sKey_VoucherDesign = "VoucherDesign"    '单据格式设置
Public Const sKey_SaveVoucherDesign = "SaveVoucherDesign"    '单据格式保存



'##ModelId=431947B203CD
Public Const gstrHelpCode As String = "Help"
'##ModelId=431947B203D6
Public gstrHelpText As String  '= "帮助"
'##ModelId=431947B203E1
Public gstrHelpTip As String  '= "帮助"
'##ModelId=431947B30002
Public Const gintHelpImg As Integer = 145


'单据列表工具条所用资源的关键字字符串
Public Const strKprintbill = "printbill"  '打印单据
Public Const strKfilter = "filter"   '过滤
Public Const strKfind = "find"    '查找
Public Const strKsetfield = "setfield"   '设置显示字段
Public Const strKsort = "sort"  '排序
Public Const strKhelp = "help"    '帮助
Public Const strKclose = "close"   '退出
Public Const strKCard = "card"    '单据
Public Const strKSelectAll = "SelectAll"    '全选
Public Const strKUnSelectAll = "UnSelectAll"    '全消

Public Const strKComparePrice = "ComparePrice"    '比价



Public Const strKLock = "lock"
Public Const strKRLock = "removelock"


Public Const sKey_Add1 = "Add1" ' "Add1"                       '增加 期初
Public strAdd1 As String  ' "期初"

Public Const sKey_Add2 = "Add2" ' "Add2"                       '增加 非期初
Public strAdd2 As String  ' "增加"
'工具栏按钮提示文字
Public strBatchprint As String  ' "批打"
Public strBatchOpen As String  ' "批开"
Public strBatchClose As String  ' "批关"
Public strBatchVeri As String  ' "批审"
Public strBatchUnVeri As String  ' "批弃"
Public strPrint As String  ' "打印"
Public strPreview As String  ' "预览"
Public strOutput As String  ' "输出"
Public strCopy As String  ' "复制"
Public strAdd As String  ' "增加"
Public strModify As String  ' "修改"
Public strdelete As String  ' "删除"
Public strSave As String  ' "保存"
Public strDiscard As String  ' "放弃"
Public strAddrecord As String  ' "增行"
Public strDeleterecord As String  ' "删行"
Public strAlterPO As String  ' "变更"
Public strReportCheck As String  ' "报检"
Public strClose As String  ' "关闭"
Public strOpen As String  ' "打开"
Public strConfirm As String  ' "审核"
Public strCancelconfirm As String  ' "弃审"
Public strQueryConfirm As String  ' "查审"
Public strPayment As String  ' "现付"
Public strCancelpayment As String  ' "弃付"
Public strBatchBV As String  ' "生成"
Public strSettle As String  ' "结算"
Public strLocate As String  ' "定位"
Public strLocateSet As String  ' "定位设置"
Public strFirst As String  ' "首张"
Public strPrevious As String  ' "上张"
Public strNext As String  ' "下张"
Public strLast As String  ' "末张"
Public strRefVoucher As String  ' "关联单据"         '关联单据 by zhangwchb 20110718
Public strRefresh As String  ' "刷新"
Public strHelp As String  ' "帮助"
Public strFilter As String  ' "过滤"
Public strExit As String  ' "退出"
Public strColumn As String  ' "栏目"
Public strSelectAll As String  ' "全选"
Public strUnSelectAll As String  ' "全消"
Public strLock As String  ' "锁定"
Public strRLock As String  ' "解锁"
Public strBatchLock As String  ' "批锁"
Public strBatchRLock As String  ' "批解"
Public strAcc As String  ' "附件"
Public strLink As String  ' "联查"
Public strFetchprice As String  ' "取价"
'工作流 chenliangc
Public strSubmit As String  ' "提交"                 '提交
Public strUnsubmit As String  ' "撤销"             '撤销
Public strResubmit As String  ' "重新提交"             '重新提交
Public strViewVerify As String  ' "查审"         '查审

Public strReferVoucher As String  ' "生单(一对一)"     '生单(一对一)
Public strReferVouchers As String  ' "生单(多对一)"   '生单(多对一)
Public strCreateVoucher As String  ' "推单"       '推单
Public strCreateSAVoucher As String  ' "销售单据"       '推单
Public strCreatePUVoucher As String  ' "采购单据"       '推单
Public strCreateSCVoucher As String  ' "库存单据"       '推单
Public strCreateAPVoucher As String  ' "应付单据"       '推单

Public strReverseSelection As String  ' "反选"
Public strVoucherDesign As String  ' "格式设置"         '单据格式设置
Public strSaveVoucherDesign As String  ' "保存布局"     '单据格式保存

Public Sub InitMulText()
    strBatchprint = GetString("U8.DZ.JA.btn010")    '") '批打"
    strBatchOpen = GetString("U8.DZ.JA.btn020")    '批开"
    strBatchClose = GetString("U8.DZ.JA.btn030")    '批关"
    strBatchVeri = GetString("U8.DZ.JA.btn035")    '批审"
    strBatchUnVeri = GetString("U8.DZ.JA.btn040")    '批弃"
    strPrint = GetString("U8.DZ.JA.btn045")    '打印"
    strPreview = GetString("U8.DZ.JA.btn050")    '预览"
    strOutput = GetString("U8.DZ.JA.btn055")    '输出"
    strCopy = GetString("U8.DZ.JA.btn060")    '复制"
    strAdd = GetString("U8.DZ.JA.btn065")    '增加"
    strAdd1 = GetString("U8.DZ.JA.btn760")
    strAdd2 = GetString("U8.DZ.JA.btn065")    '增加"
    strModify = GetString("U8.DZ.JA.btn070")    '修改"
    strdelete = GetString("U8.DZ.JA.btn075")    '删除"
    strSave = GetString("U8.DZ.JA.btn080")    '保存"
    strDiscard = GetString("U8.DZ.JA.btn090")    '放弃"
    strAddrecord = GetString("U8.DZ.JA.btn100")    '增行"
    strDeleterecord = GetString("U8.DZ.JA.btn110")    '删行"
    strAlterPO = GetString("U8.DZ.JA.btn120")    '变更"
    strReportCheck = GetString("U8.DZ.JA.btn130")    '报检"
    strClose = GetString("U8.DZ.JA.btn140")    '关闭"
    strOpen = GetString("U8.DZ.JA.btn150")    '打开"
    strConfirm = GetString("U8.DZ.JA.btn155")    '审核"
    strCancelconfirm = GetString("U8.DZ.JA.btn160")    '弃审"
    strQueryConfirm = GetString("U8.DZ.JA.btn170")    '查审"
    strPayment = GetString("U8.DZ.JA.btn180")    '现付"
    strCancelpayment = GetString("U8.DZ.JA.btn190")    '弃付"
    strBatchBV = GetString("U8.DZ.JA.btn200")    '生成"
    strSettle = GetString("U8.DZ.JA.btn210")    '结算"
    strLocate = GetString("U8.DZ.JA.btn220")    '定位"
    strLocateSet = GetString("U8.DZ.JA.btn230")    '定位设置"
    strFirst = GetString("U8.DZ.JA.btn240")    '首张"
    strPrevious = GetString("U8.DZ.JA.btn250")    '上张"
    strNext = GetString("U8.DZ.JA.btn260")    '下张"
    strLast = GetString("U8.DZ.JA.btn270")    '末张"
    strRefresh = GetString("U8.DZ.JA.btn280")    '刷新"
    strHelp = GetString("U8.DZ.JA.btn290")    '帮助"
    strFilter = GetString("U8.DZ.JA.btn300")    '过滤"
    strExit = GetString("U8.DZ.JA.btn310")    '退出"
    strColumn = GetString("U8.DZ.JA.btn320")    '栏目"
    strSelectAll = GetString("U8.DZ.JA.btn330")    '全选"
    strUnSelectAll = GetString("U8.DZ.JA.btn340")    '全消"
    strLock = GetString("U8.DZ.JA.btn350")    '锁定"
    strRLock = GetString("U8.DZ.JA.btn360")    '解锁"
    strBatchLock = GetString("U8.DZ.JA.btn370")    '批锁"
    strBatchRLock = GetString("U8.DZ.JA.btn380")    '批解"
    strAcc = GetString("U8.DZ.JA.btn390")   '附件"
    strVoucherDesign = GetString("U8.DZ.JA.btn540")
    strSaveVoucherDesign = GetString("U8.DZ.JA.btn550")
    strRefVoucher = GetString("U8.DZ.JA.btn620")
    gstrHelpText = GetString("U8.DZ.JA.btn290")   '"帮助"
    gstrHelpTip = GetString("U8.DZ.JA.btn290")   '"帮助"

    strLink = GetString("U8.DZ.JA.btn400")
    strFetchprice = GetString("U8.DZ.JA.btn410")  '"取价"
    '工作流 chenliangc
    strSubmit = GetString("U8.DZ.JA.btn420")  '"提交"                 '提交
    strUnsubmit = GetString("U8.DZ.JA.btn430")    '"撤销"             '撤销
    strResubmit = GetString("U8.DZ.JA.btn440")    ' "重新提交"             '重新提交
    strViewVerify = GetString("U8.DZ.JA.btn170")  '"查审"         '查审

    strReferVoucher = GetString("U8.DZ.JA.btn460")    ' "生单(一对一)"     '生单(一对一)
    strReferVouchers = GetString("U8.DZ.JA.btn470")  '"生单(多对一)"   '生单(多对一)
    strCreateVoucher = GetString("U8.DZ.JA.btn480")    '"推单"       '推单
    strCreateSAVoucher = GetString("U8.DZ.JA.btn490")    ' "销售单据"
    strCreatePUVoucher = GetString("U8.DZ.JA.btn500")    ' "采购单据"
    strCreateSCVoucher = GetString("U8.DZ.JA.btn510")    '"库存单据"
    strCreateAPVoucher = GetString("U8.DZ.JA.btn520")    '"应付单据"
    '
    strReverseSelection = GetString("U8.DZ.JA.btn530")    ' "反选"
End Sub


'合并工具条
Public Sub ChangeOneFormTbr(Frm As Form, objTbl As Toolbar, objU8Tbl As Control)
    Dim rs As New ADODB.Recordset
    Dim strSql As String
    strSql = "select * from UFMeta_" + g_oLogin.cAcc_Id + "..AA_CustomerButton  where cFormKey='" & gstrCardNumber & "' and cLocaleID='zh-cn' order by cButtonID"
    Set rs = g_Conn.Execute(strSql)
'    Set objU8Tbl.Business = g_oBusiness
'    With objTbl
'
'        .Buttons(sKey_Print).Tag = g_oBusiness.createportaltoolbartag("print", "ICOMMON", "PortalToolbar")
'        .Buttons(sKey_Preview).Tag = g_oBusiness.createportaltoolbartag("print preview", "ICOMMON", "PortalToolbar")
'        .Buttons(sKey_Output).Tag = g_oBusiness.createportaltoolbartag("Output", "ICOMMON", "PortalToolbar")
'
'        .Buttons(sKey_First).Tag = g_oBusiness.createportaltoolbartag("first page", "ICOMMON", "PortalToolbar")
'        .Buttons(sKey_Previous).Tag = g_oBusiness.createportaltoolbartag("previous page", "ICOMMON", "PortalToolbar")
'        .Buttons(sKey_Next).Tag = g_oBusiness.createportaltoolbartag("next page", "ICOMMON", "PortalToolbar")
'        .Buttons(sKey_Last).Tag = g_oBusiness.createportaltoolbartag("last page", "ICOMMON", "PortalToolbar")
'        .Buttons(sKey_RefVoucher).Tag = g_oBusiness.createportaltoolbartag("query", "ICOMMON", "PortalToolbar")  'zhangwchb
'
'        .Buttons(sKey_Add).Tag = g_oBusiness.createportaltoolbartag("add", "IEDIT", "PortalToolbar")
'        .Buttons(sKey_Modify).Tag = g_oBusiness.createportaltoolbartag("modify", "IEDIT", "PortalToolbar")
'        .Buttons(sKey_Delete).Tag = g_oBusiness.createportaltoolbartag("delete", "IEDIT", "PortalToolbar")
'
'
'        .Buttons(sKey_ReferVoucher).Tag = g_oBusiness.createportaltoolbartag("create", "IDEAL", "PortalToolbar")
'        .Buttons(sKey_Copy).Tag = g_oBusiness.createportaltoolbartag("Copy", "IEDIT", "PortalToolbar")
'        .Buttons(sKey_Save).Tag = g_oBusiness.createportaltoolbartag("Save", "IEDIT", "PortalToolbar")
'        .Buttons(sKey_Discard).Tag = g_oBusiness.createportaltoolbartag("back", "IEDIT", "PortalToolbar")
'
'
'        .Buttons(sKey_Confirm).Tag = g_oBusiness.createportaltoolbartag("column", "ICOMMON", "PortalToolbar")
'        .Buttons(sKey_Cancelconfirm).Tag = g_oBusiness.createportaltoolbartag("Unapprove", "ICOMMON", "PortalToolbar")
'        .Buttons(sKey_Open).Tag = g_oBusiness.createportaltoolbartag("Open", "ICOMMON", "PortalToolbar")
'        .Buttons(sKey_Close).Tag = g_oBusiness.createportaltoolbartag("Close", "ICOMMON", "PortalToolbar")
'
'
'        .Buttons(sKey_Locate).Tag = g_oBusiness.createportaltoolbartag("location", "ICOMMON", "PortalToolbar")
'        .Buttons(sKey_Refresh).Tag = g_oBusiness.createportaltoolbartag("refresh", "ICOMMON", "PortalToolbar")
'        .Buttons(gstrHelpCode).Tag = g_oBusiness.createportaltoolbartag("help", "ICOMMON", "PortalToolbar")
'
'
'
'
'        .Buttons(sKey_Addrecord).Tag = g_oBusiness.createportaltoolbartag("add a row", "IEDIT", "PortalToolbar")
'        .Buttons(sKey_Deleterecord).Tag = g_oBusiness.createportaltoolbartag("delete row", "IEDIT", "PortalToolbar")
'        .Buttons(sKey_Fetchprice).Tag = g_oBusiness.createportaltoolbartag("price", "IEDIT", "PortalToolbar")  '取价
'        .Buttons(sKey_Acc).Tag = g_oBusiness.createportaltoolbartag("accessories", "IEDIT", "PortalToolbar")    '附件
'
'        '工作流 chenliangc
'        .Buttons(sKey_Submit).Tag = g_oBusiness.createportaltoolbartag("Submit", "ICOMMON", "PortalToolbar")
'        .Buttons(sKey_Unsubmit).Tag = g_oBusiness.createportaltoolbartag("recover", "ICOMMON", "PortalToolbar")
'        .Buttons(sKey_Resubmit).Tag = g_oBusiness.createportaltoolbartag("Submit", "ICOMMON", "PortalToolbar")    '
'        .Buttons(sKey_ViewVerify).Tag = g_oBusiness.createportaltoolbartag("Relate query", "ICOMMON", "PortalToolbar")    '件
'        '生单
'        .Buttons(sKey_ReferVoucher).Tag = g_oBusiness.createportaltoolbartag("create", "IEDIT", "PortalToolbar")
'        .Buttons(sKey_CreateVoucher).Tag = g_oBusiness.createportaltoolbartag("creating", "ICOMMON", "PortalToolbar")
'        If Not rs.EOF Then
'            Do While Not rs.EOF
'                .Buttons(CStr(rs!cButtonkey)).Tag = g_oBusiness.createportaltoolbartag(CStr(rs!cImage), CStr(rs!cGroup), "PortalToolbar")
'                rs.MoveNext
'            Loop
'        End If
'        'U810.0适配， 添加格式设置和保存按钮
'        .Buttons(sKey_VoucherDesign).Tag = g_oBusiness.createportaltoolbartag("format", "IEDIT", "PortalToolbar")
'        .Buttons(sKey_SaveVoucherDesign).Tag = g_oBusiness.createportaltoolbartag("format", "IEDIT", "PortalToolbar")
'
'    End With
    'InitToolBar方法里已经设置过
    'objU8Tbl.SetToolbar objTbl
    objU8Tbl.SetDisplayStyle 0    'TextOnly
    objTbl.Visible = False
    objU8Tbl.Visible = True
    objU8Tbl.Left = objTbl.Left
    objU8Tbl.Top = objTbl.Top
    objU8Tbl.Width = Frm.Width - 6 * Screen.TwipsPerPixelX
    objU8Tbl.Height = objTbl.Height
End Sub

'------------------------------------------------------------
'初始化工具栏控件
'------------------------------------------------------------
Public Sub Init_Toolbar(tlbObj As Toolbar)
    Dim rs As New ADODB.Recordset
    Dim strSql As String
    strSql = "select * from UFMeta_" + g_oLogin.cAcc_Id + "..AA_CustomerButton  where cFormKey='" & gstrCardNumber & "' and cLocaleID='zh-cn' order by cButtonID"
    Set rs = g_Conn.Execute(strSql)


    With tlbObj.Buttons
        .Clear

        .Add , sKey_Print, strPrint
        .Item(sKey_Print).ToolTipText = strPrint + "Ctrl+P"

        .Add , sKey_Preview, strPreview
        .Item(sKey_Preview).ToolTipText = strPreview + "Ctrl+W"

        .Add , sKey_Output, strOutput
        .Item(sKey_Output).ToolTipText = strOutput + "Alt+E"

        .Add , "S1", , tbrSeparator


        .Add , sKey_Confirm, strConfirm
        .Item(sKey_Confirm).ToolTipText = strConfirm + " Ctrl+U"

        .Add , sKey_Cancelconfirm, strCancelconfirm
        .Item(sKey_Cancelconfirm).ToolTipText = strCancelconfirm + "Alt+U"
        
         .Add , sKey_ViewVerify, strViewVerify
        .Item(sKey_ViewVerify).ToolTipText = strViewVerify
        
                '工作流 chenliangc
        .Add , sKey_Submit, strSubmit
        .Item(sKey_Submit).ToolTipText = strSubmit + " Ctrl+J"

        .Add , sKey_Resubmit, strResubmit
        .Item(sKey_Resubmit).ToolTipText = strResubmit

        .Add , sKey_Unsubmit, strUnsubmit
        .Item(sKey_Unsubmit).ToolTipText = strUnsubmit + "Alt+J"
     
           .Add , sKey_Open, strOpen
        .Item(sKey_Open).ToolTipText = strOpen + "Alt+O"

        .Add , sKey_Close, strClose
        .Item(sKey_Close).ToolTipText = strClose + "Alt+C"
        
          .Add , "S2", , tbrSeparator




        .Add , sKey_First, strFirst
        .Item(sKey_First).ToolTipText = strFirst + "Alt+PageUp"

        .Add , sKey_Previous, strPrevious
        .Item(sKey_Previous).ToolTipText = strPrevious + "PageUp"

        .Add , sKey_Next, strNext
        .Item(sKey_Next).ToolTipText = strNext + " PageDown"

        .Add , sKey_Last, strLast
        .Item(sKey_Last).ToolTipText = strLast + "Alt+PageDown"



        .Add , "S3", , tbrSeparator


        .Add , sKey_Locate, strLocate
        .Item(sKey_Locate).ToolTipText = strLocate + "Ctrl+F3"
        
        .Add , sKey_RefVoucher, strRefVoucher 'zhangwchb
        .Item(sKey_RefVoucher).ToolTipText = strRefVoucher

        .Add , sKey_Refresh, strRefresh
        .Item(sKey_Refresh).ToolTipText = strRefresh + "Ctrl+R"
        
  
        .Add , gstrHelpCode, gstrHelpText
        .Item(gstrHelpCode).ToolTipText = gstrHelpText + " F1"


        .Add , "S4", , tbrSeparator


        .Add , sKey_Add, strAdd
        .Item(sKey_Add).ToolTipText = strAdd + " F5"
        .Item(sKey_Add).Style = tbrDropdown
        .Item(sKey_Add).ButtonMenus.Add , sKey_Add2, strAdd2
        .Item(sKey_Add).ButtonMenus.Add , sKey_Add1, strAdd1
       

        .Add , sKey_Modify, strModify
        .Item(sKey_Modify).ToolTipText = strModify + " F8"

        .Add , sKey_Delete, strdelete
        .Item(sKey_Delete).ToolTipText = strdelete + " Delete"

        '生单 chenliangc
        .Add , sKey_ReferVoucher, strReferVoucher
        .Item(sKey_ReferVoucher).ToolTipText = strReferVoucher
        .Item(sKey_ReferVoucher).Style = tbrDropdown
        Call .Item(sKey_ReferVoucher).ButtonMenus.Add(, sKey_ReferVouchers, strReferVouchers)

        '推单 chenliangc
        .Add , sKey_CreateVoucher, strCreateVoucher
        .Item(sKey_CreateVoucher).ToolTipText = strCreateVoucher
        .Item(sKey_CreateVoucher).Style = tbrDropdown

        .Item(sKey_CreateVoucher).ButtonMenus.Add , sKey_CreateSAVoucher, strCreateSAVoucher
        .Item(sKey_CreateVoucher).ButtonMenus.Add , sKey_CreatePUVoucher, strCreatePUVoucher
        .Item(sKey_CreateVoucher).ButtonMenus.Add , sKey_CreateSCVoucher, strCreateSCVoucher
        .Item(sKey_CreateVoucher).ButtonMenus.Add , sKey_CreateAPVoucher, strCreateAPVoucher

        .Add , sKey_Copy, strCopy
        .Item(sKey_Copy).ToolTipText = strCopy + "Ctrl+F5"

        .Add , sKey_Save, strSave
        .Item(sKey_Save).ToolTipText = strSave + " F6"

        .Add , sKey_Discard, strDiscard
        .Item(sKey_Discard).ToolTipText = strDiscard + " Ctrl+Z"


        .Add , "S5", , tbrSeparator


        .Add , sKey_Addrecord, strAddrecord
        .Item(sKey_Addrecord).ToolTipText = strAddrecord + "Ctrl+N"

        .Add , sKey_Deleterecord, strDeleterecord
        .Item(sKey_Deleterecord).ToolTipText = strDeleterecord + "Ctrl+D"
        
         '附件
        .Add , sKey_Acc, strAcc
        .Item(sKey_Acc).ToolTipText = strAcc

        '取价
        With .Add(, sKey_Fetchprice, strFetchprice, tbrDropdown)
            Call .ButtonMenus.Add(, "rowprice", GetString("U8.DZ.JA.btn770"))
            Call .ButtonMenus.Add(, "allprice", GetString("U8.DZ.JA.btn780"))
        End With
        'U810 适配  添加 单据格式设置和保存按钮  2011/03/04   LEW
        .Add , sKey_VoucherDesign, strVoucherDesign                  '单据格式设置按钮
        .Item(sKey_VoucherDesign).ToolTipText = strVoucherDesign


        .Add , sKey_SaveVoucherDesign, strSaveVoucherDesign          '格式保存按钮
        .Item(sKey_SaveVoucherDesign).ToolTipText = strSaveVoucherDesign

       
        If Not rs.EOF Then
            Do While Not rs.EOF
                .Add , CStr(rs!cButtonkey), CStr(rs!cCaption)
                '               .Item(sKey_Deleterecord).ToolTipText = Rs!cCaption + Rs!cHotKey
                .Item(CStr(rs!cButtonkey)).ToolTipText = rs!cCaption + rs!cHotKey
                rs.MoveNext
            Loop
        End If


    End With
End Sub

'设置 相关按钮 显示与否
Public Sub SetWFControlBrnsList(login As clsLogin, myConn As ADODB.Connection, Toolbar As Object, UFToolbar As Object, cardnumber As String)
    Dim rstfilter As String

    '   .Buttons(sKey_Batchprint).Tag = goBusiness.createportaltoolbartag("Print", "ICOMMON", "PortalToolbar")
    '   .Buttons(strKSelectAll).Tag = goBusiness.createportaltoolbartag("Select All", "ICOMMON", "PortalToolbar")
    '   .Buttons(sKey_ReverseSelection).Tag = goBusiness.createportaltoolbartag("Revise", "ICOMMON", "PortalToolbar")
    '   .Buttons(strKUnSelectAll).Tag = goBusiness.createportaltoolbartag("Select none", "ICOMMON", "PortalToolbar")
    '   .Buttons(sKey_Link).Tag = goBusiness.createportaltoolbartag("relate query", "ICOMMON", "PortalToolbar")
    '   .Buttons(sKey_Confirm).Tag = goBusiness.createportaltoolbartag("Approve", "ICOMMON", "PortalToolbar")
    '   .Buttons(sKey_Cancelconfirm).Tag = goBusiness.createportaltoolbartag("Unapprove", "ICOMMON", "PortalToolbar")
    '   .Buttons(sKey_Open).Tag = goBusiness.createportaltoolbartag("Open", "ICOMMON", "PortalToolbar")
    '   .Buttons(sKey_Close).Tag = goBusiness.createportaltoolbartag("Close", "ICOMMON", "PortalToolbar")


    Toolbar.Buttons(sKey_CreateVoucher).Visible = False
    Toolbar.Buttons(sKey_Close).Visible = False
    Toolbar.Buttons(sKey_Open).Visible = False
    '    Toolbar.Buttons(sKey_Cancelconfirm).Visible = False
    '    Toolbar.Buttons(sKey_Confirm).Visible = False
    '    Toolbar.Buttons(strKUnSelectAll).Visible = False
    '    Toolbar.Buttons(strKSelectAll).Visible = False
    '    Toolbar.Buttons(sKey_ReverseSelection).Visible = False

    UFToolbar.RefreshVisible
End Sub

'合并工具条
Public Sub ChangeOneFormTbrlist(Frm As Form, objTbl As Toolbar, objU8Tbl As Control)
    Set objU8Tbl.Business = goBusiness
    With objTbl

        .Buttons(sKey_Print).Tag = goBusiness.createportaltoolbartag("print", "ICOMMON", "PortalToolbar")
        .Buttons(sKey_Preview).Tag = goBusiness.createportaltoolbartag("print preview", "ICOMMON", "PortalToolbar")
        .Buttons(sKey_Output).Tag = goBusiness.createportaltoolbartag("Output", "ICOMMON", "PortalToolbar")

        .Buttons(sKey_Locate).Tag = goBusiness.createportaltoolbartag("Location", "ICOMMON", "PortalToolbar")
        .Buttons(strKfilter).Tag = goBusiness.createportaltoolbartag("filter", "ICOMMON", "PortalToolbar")
        .Buttons(sKey_Column).Tag = goBusiness.createportaltoolbartag("column", "ICOMMON", "PortalToolbar")
        .Buttons(sKey_Batchprint).Tag = goBusiness.createportaltoolbartag("Print", "ICOMMON", "PortalToolbar")


        .Buttons(strKSelectAll).Tag = goBusiness.createportaltoolbartag("Select All", "ICOMMON", "PortalToolbar")
        .Buttons(sKey_ReverseSelection).Tag = goBusiness.createportaltoolbartag("Revise", "ICOMMON", "PortalToolbar")
        .Buttons(strKUnSelectAll).Tag = goBusiness.createportaltoolbartag("Select none", "ICOMMON", "PortalToolbar")
       ' .Buttons(sKey_Link).Tag = goBusiness.createportaltoolbartag("relate query", "ICOMMON", "PortalToolbar")


        .Buttons(sKey_Confirm).Tag = goBusiness.createportaltoolbartag("Approve", "ICOMMON", "PortalToolbar")
        .Buttons(sKey_Cancelconfirm).Tag = goBusiness.createportaltoolbartag("Unapprove", "ICOMMON", "PortalToolbar")
        .Buttons(sKey_Open).Tag = goBusiness.createportaltoolbartag("Open", "ICOMMON", "PortalToolbar")
        .Buttons(sKey_Close).Tag = goBusiness.createportaltoolbartag("Close", "ICOMMON", "PortalToolbar")
        '推单 chenliangc
        .Buttons(sKey_CreateVoucher).Tag = goBusiness.createportaltoolbartag("creating", "ICOMMON", "PortalToolbar")
        .Buttons(sKey_Refresh).Tag = goBusiness.createportaltoolbartag("refresh", "ICOMMON", "PortalToolbar")
        .Buttons(gstrHelpCode).Tag = goBusiness.createportaltoolbartag("help", "ICOMMON", "PortalToolbar")

    End With
    'InitToolBar方法里已经设置过
    'objU8Tbl.SetToolbar objTbl
    objU8Tbl.SetDisplayStyle 0    'TextOnly
    objTbl.Visible = False
    objU8Tbl.Visible = True
    objU8Tbl.Left = objTbl.Left
    objU8Tbl.Top = objTbl.Top
    objU8Tbl.Width = Frm.Width - 6 * Screen.TwipsPerPixelX
    objU8Tbl.Height = objTbl.Height
End Sub

'------------------------------------------------------------
'初始化工具栏控件
'------------------------------------------------------------
Public Sub Init_Toolbarlist(tlbObj As Toolbar)

    With tlbObj.Buttons
        .Clear

        .Add , sKey_Print, strPrint
        .Item(sKey_Print).ToolTipText = strPrint

        .Add , sKey_Preview, strPreview
        .Item(sKey_Preview).ToolTipText = strPreview

        .Add , sKey_Output, strOutput
        .Item(sKey_Output).ToolTipText = strOutput

        '        .Add , "S1", , tbrSeparator

        .Add , strKfilter, strFilter
        .Item(strKfilter).ToolTipText = strFilter
        .Item(strKfilter).Visible = False
        .Add , sKey_Locate, strLocate
        .Item(sKey_Locate).ToolTipText = strLocate

'        .Add , sKey_Link, strLink
'        .Item(sKey_Link).ToolTipText = strLink


        .Add , sKey_Column, strColumn
        .Item(sKey_Column).ToolTipText = strColumn

        .Add , sKey_Batchprint, strBatchprint
        .Item(sKey_Batchprint).ToolTipText = strBatchprint





        .Add , strKSelectAll, strSelectAll
        .Item(strKSelectAll).ToolTipText = strSelectAll
        .Item(strKSelectAll).Visible = False

        .Add , sKey_ReverseSelection, strReverseSelection
        .Item(sKey_ReverseSelection).ToolTipText = strReverseSelection
        .Item(sKey_ReverseSelection).Visible = False
        
        .Add , strKUnSelectAll, strUnSelectAll
        .Item(strKUnSelectAll).ToolTipText = strUnSelectAll
        .Item(strKUnSelectAll).Visible = False



        .Add , sKey_Confirm, strBatchVeri
        .Item(sKey_Confirm).ToolTipText = strBatchVeri

        .Add , sKey_Cancelconfirm, strBatchUnVeri
        .Item(sKey_Cancelconfirm).ToolTipText = strBatchUnVeri

        .Add , sKey_Open, strBatchOpen
        .Item(sKey_Open).ToolTipText = strBatchOpen

        .Add , sKey_Close, strBatchClose
        .Item(sKey_Close).ToolTipText = strBatchClose

        '         .Add , "S3", , tbrSeparator



        .Add , sKey_Refresh, strRefresh
        .Item(sKey_Refresh).ToolTipText = strRefresh
        .Item(sKey_Refresh).Visible = False
        
        .Add , gstrHelpCode, gstrHelpText
        .Item(gstrHelpCode).ToolTipText = gstrHelpText + " F1"


        '        .Add , "S4", , tbrSeparator
        '推单 chenliangc
        .Add , sKey_CreateVoucher, strCreateVoucher, tbrDropdown
        .Item(sKey_CreateVoucher).ToolTipText = strCreateVoucher
        .Item(sKey_CreateVoucher).ButtonMenus.Add , sKey_CreateSAVoucher, strCreateSAVoucher
        .Item(sKey_CreateVoucher).ButtonMenus.Add , sKey_CreatePUVoucher, strCreatePUVoucher
        .Item(sKey_CreateVoucher).ButtonMenus.Add , sKey_CreateSCVoucher, strCreateSCVoucher
        .Item(sKey_CreateVoucher).ButtonMenus.Add , sKey_CreateAPVoucher, strCreateAPVoucher

        .Item(sKey_Batchprint).Visible = False   '暂不支持批打
    End With
End Sub


'工具栏状态
'工具栏状态
Public Sub SetCtlStyle(Frm As Form, Voucher As Object, Toolbar As Toolbar, UFToolbar As UFToolbar, mOpStatus As OpStatus)
    On Error Resume Next
    Dim sql As String
    Dim strSql As String
    Dim rs As New ADODB.Recordset

    With Toolbar
        Select Case mOpStatus
            '增加
        Case ADD_MAIN
            .Buttons(sKey_Print).Enabled = False  '打印
            .Buttons(sKey_Preview).Enabled = False    '预览
            .Buttons(sKey_Output).Enabled = False    '输出

            .Buttons(sKey_First).Enabled = False    '首页
            .Buttons(sKey_Previous).Enabled = False    '上一页
            .Buttons(sKey_Next).Enabled = False    '下一页
            .Buttons(sKey_Last).Enabled = False    '末页

            .Buttons(sKey_Add).Enabled = False    '增加
            .Buttons(sKey_Modify).Enabled = False    '修改
            .Buttons(sKey_Delete).Enabled = False    '删除


            .Buttons(sKey_ReferVoucher).Enabled = True    '生单
            .Buttons(sKey_Copy).Enabled = False    '复制
            .Buttons(sKey_Save).Enabled = True    '保存
            .Buttons(sKey_Discard).Enabled = True    '放弃
            .Buttons(sKey_Fetchprice).Enabled = True    '取价

            .Buttons(sKey_Submit).Enabled = False  '提交
            .Buttons(sKey_Resubmit).Enabled = False    '重新提交
            .Buttons(sKey_Unsubmit).Enabled = False   '撤销
            .Buttons(sKey_ViewVerify).Enabled = False    '查审
            .Buttons(sKey_Confirm).Enabled = False    '审核
            .Buttons(sKey_Cancelconfirm).Enabled = False    '弃审
            .Buttons(sKey_Return).Enabled = False   '归还
            .Buttons(sKey_CreateVoucher).Enabled = False    '推单
            .Buttons(sKey_Open).Enabled = False    '打开
            .Buttons(sKey_Close).Enabled = False    '关闭

            .Buttons(sKey_Locate).Enabled = False    '定位
            .Buttons(sKey_Refresh).Enabled = False    '刷新
            .Buttons(gstrHelpCode).Enabled = True    '帮助
            .Buttons(sKey_VoucherDesign).Enabled = False

            .Buttons(sKey_Addrecord).Enabled = True    '增行
            .Buttons(sKey_Deleterecord).Enabled = True    '删行
            .Buttons(sKey_Acc).Enabled = True    '附件
            .Buttons("Prorefer1").Enabled = False
            .Buttons("Prorefer2").Enabled = False
            .Buttons("Prorefer3").Enabled = False
            
            




            Voucher.VoucherStatus = VSeAddMode

            Frm.ComTemplatePRN.Visible = False
            Frm.ComTemplateShow.Visible = True
            Frm.LblTemplate.Caption = GetString("U8.DZ.JA.Res020")
            Frm.LblTemplate.Visible = True

            '修改
        Case MODIFY_MAIN

            .Buttons(sKey_Print).Enabled = False    '打印
            .Buttons(sKey_Preview).Enabled = False    '预览
            .Buttons(sKey_Output).Enabled = False    '输出
            .Buttons("Prorefer").Enabled = False

            .Buttons(sKey_First).Enabled = False    '首页
            .Buttons(sKey_Previous).Enabled = False    '上一页
            .Buttons(sKey_Next).Enabled = False    '下一页
            .Buttons(sKey_Last).Enabled = False    '末页

            .Buttons(sKey_Add).Enabled = False    '增加
            .Buttons(sKey_Modify).Enabled = False    '修改
            .Buttons(sKey_Delete).Enabled = False    '删除
            .Buttons(sKey_Fetchprice).Enabled = True    '取价

            .Buttons(sKey_ReferVoucher).Enabled = False    '生单
            .Buttons(sKey_Copy).Enabled = False    '复制
            .Buttons(sKey_Save).Enabled = True    '保存
            .Buttons(sKey_Discard).Enabled = True    '放弃

            .Buttons(sKey_Submit).Enabled = False  '提交
            .Buttons(sKey_Resubmit).Enabled = False    '重新提交
            .Buttons(sKey_Unsubmit).Enabled = False   '撤销
            .Buttons(sKey_ViewVerify).Enabled = False    '查审
            .Buttons(sKey_Confirm).Enabled = False    '审核
            .Buttons(sKey_Cancelconfirm).Enabled = False    '弃审
            .Buttons(sKey_Return).Enabled = False   '归还
            .Buttons(sKey_CreateVoucher).Enabled = False    '推单
            .Buttons(sKey_Confirm).Enabled = False    '审核
            .Buttons(sKey_Cancelconfirm).Enabled = False    '弃审
            .Buttons(sKey_Open).Enabled = False    '打开
            .Buttons(sKey_Close).Enabled = False    '关闭
            
            .Buttons("Prorefer1").Enabled = False
            .Buttons("Prorefer2").Enabled = False
            .Buttons("Prorefer3").Enabled = False


            .Buttons(sKey_Locate).Enabled = False    '定位
            .Buttons(sKey_Refresh).Enabled = False    '刷新
            .Buttons(gstrHelpCode).Enabled = True    '帮助


            .Buttons(sKey_Addrecord).Enabled = True    '增行
            .Buttons(sKey_Deleterecord).Enabled = True    '删行
            .Buttons(sKey_VoucherDesign).Enabled = False
            .Buttons(sKey_Acc).Enabled = True    '附件

            Voucher.VoucherStatus = VSeEditMode

            Frm.ComTemplatePRN.Visible = False
            Frm.ComTemplateShow.Visible = True
            Frm.LblTemplate.Caption = GetString("U8.DZ.JA.Res020")
            Frm.LblTemplate.Visible = True

            '显示
        Case SHOW_ALL

            .Buttons(sKey_Print).Enabled = True    '打印
            .Buttons(sKey_Preview).Enabled = True    '预览
            .Buttons(sKey_Output).Enabled = True    '输出

            .Buttons(sKey_First).Enabled = True    '首页
            .Buttons(sKey_Previous).Enabled = True    '上一页
            .Buttons(sKey_Next).Enabled = True    '下一页
            .Buttons(sKey_Last).Enabled = True    '末页

            .Buttons(sKey_Add).Enabled = True    '增加
            .Buttons(sKey_Modify).Enabled = True    '修改
            .Buttons(sKey_Delete).Enabled = True    '删除
            .Buttons("Prorefer").Enabled = True

            .Buttons(sKey_ReferVoucher).Enabled = False    '生单
            .Buttons(sKey_Copy).Enabled = True    '复制
            .Buttons(sKey_Save).Enabled = False    '保存
            .Buttons(sKey_Discard).Enabled = False    '放弃
            .Buttons(sKey_Fetchprice).Enabled = False    '取价
            .Buttons(sKey_VoucherDesign).Enabled = True

            .Buttons(sKey_Confirm).Enabled = True    '审核
            .Buttons(sKey_Cancelconfirm).Enabled = True    '弃审
            
            .Buttons(sKey_Locate).Enabled = True    '定位
            .Buttons(sKey_Refresh).Enabled = True    '刷新
            .Buttons(gstrHelpCode).Enabled = True    '帮助


            .Buttons(sKey_Addrecord).Enabled = False    '增行
            .Buttons(sKey_Deleterecord).Enabled = False    '删行
            .Buttons(sKey_Acc).Enabled = False    '附件
           ' .Buttons(sKey_Acc).Visible = False
           .Buttons("Prorefer1").Enabled = True
            .Buttons("Prorefer2").Enabled = True
            .Buttons("Prorefer3").Enabled = True



            '翻页按钮
            If pageCount <= 1 Then
                .Buttons(sKey_First).Enabled = False    '首页
                .Buttons(sKey_Previous).Enabled = False    '上一页
                .Buttons(sKey_Next).Enabled = False    '下一页
                .Buttons(sKey_Last).Enabled = False    '末页
            ElseIf PageCurrent = pageCount Then
                .Buttons(sKey_First).Enabled = True    '首页
                .Buttons(sKey_Previous).Enabled = True    '上一页
                .Buttons(sKey_Next).Enabled = False    '下一页
                .Buttons(sKey_Last).Enabled = False    '末页
            ElseIf PageCurrent = 1 Then
                .Buttons(sKey_First).Enabled = False    '首页
                .Buttons(sKey_Previous).Enabled = False    '上一页
                .Buttons(sKey_Next).Enabled = True    '下一页
                .Buttons(sKey_Last).Enabled = True    '末页
            Else
                .Buttons(sKey_First).Enabled = True    '首页
                .Buttons(sKey_Previous).Enabled = True    '上一页
                .Buttons(sKey_Next).Enabled = True    '下一页
                .Buttons(sKey_Last).Enabled = True    '末页

            End If


            'modify by chenliangc 添加工作流按钮显示

            '           SQL = "SELECT iStatus,iswfcontrolled,iverifystate,DownStreamcode,case when isnull(closeuser,N'')=N'' then 1 else 0 end as Closed FROM HY_DZ_BorrowOutChange WHERE ID=" & lngVoucherID
            sql = "SELECT iStatus,iswfcontrolled,iverifystate,case when isnull(closeuser,N'')=N'' then 1 else 0 end as Closed FROM " & MainTable & " WHERE ID=" & lngVoucherID
            Set rs = New ADODB.Recordset
            rs.Open sql, g_Conn, 1, 1
            If Not rs.EOF Then
            
                .Buttons(sKey_ViewVerify).Enabled = True    '查审
                
                Select Case rs("iStatus")
                    '正常
                Case 1
                    .Buttons(sKey_Open).Enabled = False    '打开
                    .Buttons(sKey_Close).Enabled = False    '关闭
                    .Buttons(sKey_CreateVoucher).Enabled = False    '推单
                    .Buttons(sKey_ReferVoucher).Enabled = False    '生单
                    .Buttons(sKey_Cancelconfirm).Enabled = False    '弃审
                    .Buttons(sKey_Return).Enabled = False   '归还
                    If Not CBool(Null2Something(rs("iswfcontrolled"), "0")) Then    '不进入工作流
                        .Buttons(sKey_Submit).Enabled = False  '提交
                        .Buttons(sKey_Resubmit).Enabled = False    '重新提交
                        .Buttons(sKey_Unsubmit).Enabled = False   '撤销
                        
                    ElseIf CInt(Null2Something(rs("iverifystate"), 0)) <= 0 Then  '工作流控制但未提交
                        .Buttons(sKey_Submit).Enabled = True  '提交
                        .Buttons(sKey_Resubmit).Enabled = False    '重新提交
                        .Buttons(sKey_Unsubmit).Enabled = False   '撤销
                        .Buttons(sKey_Confirm).Enabled = False    '审核
                        .Buttons(sKey_Cancelconfirm).Enabled = False    '弃审
                        .Buttons(sKey_Return).Enabled = False   '归还
                    ElseIf CInt(rs("iverifystate")) = 1 Then   '工作流控制已提交
                        .Buttons(sKey_Submit).Enabled = False  '提交
                        .Buttons(sKey_Resubmit).Enabled = True    '重新提交
                        .Buttons(sKey_Unsubmit).Enabled = True   '撤销
                        .Buttons(sKey_Confirm).Enabled = True    '审核
                        .Buttons(sKey_Cancelconfirm).Enabled = True    '弃审
                        .Buttons(sKey_Return).Enabled = True   '归还
                        .Buttons(sKey_Delete).Enabled = False    '删除
                    End If

                    '审核 审核人不为空
                Case 2
           
                    .Buttons(sKey_Open).Enabled = False    '打开
                    .Buttons(sKey_Close).Enabled = True    '关闭
                    .Buttons(sKey_Submit).Enabled = False  '提交
                    .Buttons(sKey_Resubmit).Enabled = True    '重新提交
                    .Buttons(sKey_Unsubmit).Enabled = True   '撤销
                    .Buttons(sKey_Modify).Enabled = False    '修改
                    .Buttons(sKey_Delete).Enabled = False    '删除
                    .Buttons(sKey_Confirm).Enabled = False    '审核
                    .Buttons(sKey_Submit).Enabled = False    '提交
                    .Buttons(sKey_Unsubmit).Enabled = True    '撤销
                    .Buttons(sKey_Resubmit).Enabled = True    '重新提交
                    .Buttons(sKey_Cancelconfirm).Enabled = True    '弃审
                    .Buttons(sKey_CreateVoucher).Enabled = True    '推单
                    .Buttons(sKey_Return).Enabled = True   '归还
                    
                  If Not CBool(Null2Something(rs("iswfcontrolled"), "0")) Then    '不进入工作流
                        .Buttons(sKey_Submit).Enabled = False  '提交
                        .Buttons(sKey_Resubmit).Enabled = False    '重新提交
                        .Buttons(sKey_Unsubmit).Enabled = False   '撤销
'                        .Buttons(sKey_ViewVerify).Enabled = False    '查审
                    End If

                    '生单
                Case 3
                    .Buttons(sKey_Submit).Enabled = False  '提交
                    .Buttons(sKey_Resubmit).Enabled = False    '重新提交
                    .Buttons(sKey_Unsubmit).Enabled = False   '撤销
                    .Buttons(sKey_Modify).Enabled = False    '修改
                    .Buttons(sKey_Delete).Enabled = False    '删除
                    .Buttons(sKey_Confirm).Enabled = False    '审核
                    .Buttons(sKey_Cancelconfirm).Enabled = False    '弃审
                    .Buttons(sKey_Open).Enabled = False    '打开
                    .Buttons(sKey_Close).Enabled = True    '关闭
                    .Buttons(sKey_Submit).Enabled = False    '提交
                    .Buttons(sKey_Unsubmit).Enabled = False    '撤销
                    .Buttons(sKey_Resubmit).Enabled = False    '重新提交
                    .Buttons(sKey_CreateVoucher).Enabled = False    '推单
                    .Buttons(sKey_Return).Enabled = True   '归还


                    '关闭
                Case 4
                    .Buttons(sKey_Submit).Enabled = False  '提交
                    .Buttons(sKey_Resubmit).Enabled = False    '重新提交
                    .Buttons(sKey_Unsubmit).Enabled = False   '撤销
                    .Buttons(sKey_Modify).Enabled = False    '修改
                    .Buttons(sKey_Delete).Enabled = False    '删除
                    .Buttons(sKey_ReferVoucher).Enabled = False    '生单
                    .Buttons(sKey_Confirm).Enabled = False    '审核
                    .Buttons(sKey_Cancelconfirm).Enabled = False    '弃审
                    .Buttons(sKey_Open).Enabled = True    '打开
                    .Buttons(sKey_Close).Enabled = False    '关闭
                    .Buttons(sKey_Submit).Enabled = False    '提交
                    .Buttons(sKey_Unsubmit).Enabled = False    '撤销
                    .Buttons(sKey_Resubmit).Enabled = False    '重新提交
                    .Buttons(sKey_CreateVoucher).Enabled = False    '推单
                    .Buttons(sKey_Return).Enabled = False   '归还


                End Select
                 
                        
                '没有记录
            Else
                .Buttons(sKey_Print).Enabled = False    '打印
                .Buttons(sKey_Preview).Enabled = False    '预览
                .Buttons(sKey_Output).Enabled = False    '输出

                .Buttons(sKey_First).Enabled = False    '首页
                .Buttons(sKey_Previous).Enabled = False    '上一页
                .Buttons(sKey_Next).Enabled = False    '下一页
                If pageCount >= 1 Then
                    .Buttons(sKey_Last).Enabled = True    '末页
                Else
                    .Buttons(sKey_Last).Enabled = False    '末页
                End If
                .Buttons(sKey_Add).Enabled = True    '增加
                .Buttons(sKey_Modify).Enabled = False    '修改
                .Buttons(sKey_Delete).Enabled = False    '删除


                .Buttons(sKey_ReferVoucher).Enabled = False    '生单
                .Buttons(sKey_CreateVoucher).Enabled = False    '推单

                .Buttons(sKey_Copy).Enabled = False    '复制
                .Buttons(sKey_Save).Enabled = False    '保存
                .Buttons(sKey_Discard).Enabled = False    '放弃


                .Buttons(sKey_Submit).Enabled = False    '提交
                .Buttons(sKey_Unsubmit).Enabled = False    '撤销
                .Buttons(sKey_Resubmit).Enabled = False    '重新提交
                .Buttons(sKey_ViewVerify).Enabled = False    '查审
                .Buttons(sKey_Confirm).Enabled = False    '审核
                .Buttons(sKey_Cancelconfirm).Enabled = False    '弃审
                .Buttons(sKey_Open).Enabled = False    '打开
                .Buttons(sKey_Close).Enabled = False    '关闭
                .Buttons(sKey_Return).Enabled = False   '归还
                .Buttons("Prorefer1").Enabled = True
                .Buttons("Prorefer2").Enabled = True
                .Buttons("Prorefer3").Enabled = True
    

                .Buttons(sKey_Locate).Enabled = True    '定位
                .Buttons(sKey_Refresh).Enabled = False    '刷新
                .Buttons(gstrHelpCode).Enabled = True    '帮助


                .Buttons(sKey_Addrecord).Enabled = False    '增行
                .Buttons(sKey_Deleterecord).Enabled = False    '删行
                .Buttons(sKey_CreateVoucher).Enabled = False    '推单
            End If

            Voucher.VoucherStatus = VSNormalMode

            Frm.ComTemplatePRN.Visible = True
            Frm.ComTemplateShow.Visible = False
            Frm.LblTemplate.Caption = GetString("U8.DZ.JA.Res010")
            Frm.LblTemplate.Visible = True


            rs.Close
            Set rs = Nothing


        End Select
    End With

    strSql = "select * from UFMeta_" + g_oLogin.cAcc_Id + "..AA_CustomerButton  where cFormKey='" & gstrCardNumber & "' and cLocaleID='zh-cn'   order by cButtonID"
    Set rs = g_Conn.Execute(strSql)
    If Not rs.EOF Then
        Do While Not rs.EOF
            Toolbar.Buttons(CStr(rs!cButtonkey)).Visible = Toolbar.Buttons(CStr(rs!cVisibleAsKey)).Visible
            Toolbar.Buttons(CStr(rs!cButtonkey)).Visible = Toolbar.Buttons(CStr(rs!cEnableAsKey)).Visible
            rs.MoveNext
        Loop
    End If

    If Voucher.headerText("cborrowouttype") = "2" Then
        '备机借出不能复制
        Toolbar.Buttons(sKey_Copy).Enabled = False
    End If

    Toolbar.Refresh
    UFToolbar.RefreshEnable    '注意,此处必须是RefreshEnable方法,Refresh不起作用


End Sub
