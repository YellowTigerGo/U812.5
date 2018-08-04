Attribute VB_Name = "Mod"
'****************************************
'工程基本功能说明：
'               1、实现单据的基本功能
'               2、实现列表的基本功能

'创建时间：2008-11-21
'创建人：xuyan
'****************************************
Option Explicit
' * 单据号
Public gstrCardNumber As String '=ua_menu表中的cMenu_Id字段的值，单据CARDNUM
Public gstrCardNumberlist As String '列表id

Public g_oBusiness As Object '单据business对象,由门户传入,该对象的作用：一是用来获取Login对象，二是用来对窗体的显示、关闭、激活等进行操作。
Public g_oLogin As U8Login.clsLogin '单据Login对象,与原861的Login对象同，但不再由门户传入,而是通过business对象间接得到。引用"u8注册部件[v8.71]"
Public g_bLogined As Boolean '单据

Public goBusiness As Object '列表business对象,由门户传入,该对象的作用：一是用来获取Login对象，二是用来对窗体的显示、关闭、激活等进行操作。
Public goLogin As U8Login.clsLogin '列表Login对象,与原861的Login对象同，但不再由门户传入,而是通过business对象间接得到。引用"u8注册部件[v8.71]"
Public gbLogined As Boolean '列表

Public g_Conn As New ADODB.Connection '单据连接
Public gConn As New ADODB.Connection '列表连接

Public clsbill As Object ' USERPCO.VoucherCO '?参照时获取sql串
Public mologin As USCOMMON.login '?

Public gsGUIDForVouch As String                '单据的GUID
Public gsGUIDForList As String                 '列表的GUID

Public idtmp As String
Public gMoCode As String
    
'参照生单获得的表头表体数据DOM
Public gDomReferHead As DOMDocument
Public gDomReferBody As DOMDocument
Public moneytmp As Double
Public numbertmp As Double
Public numappprice As Double
Public isfyflg As Boolean

Public iSinvCZ As Boolean

Public m_OK As Boolean
 
Public tmpTableName As String
Public bSwFlag As Boolean
  
 

'是否控制存货权限,存货权限分类 11-7-12 chenliangc
Public bInv_ControlAuth As Boolean
Public sAuth_invR As String, sAuth_invW As String
Attribute sAuth_invW.VB_VarUserMemId = 1073741841

'是否控制业务员权限,业务员权限字符串 11-7-12 chenliangc
Public bPerson_ControlAuth As Boolean
Attribute bPerson_ControlAuth.VB_VarUserMemId = 1073741843
Public sAuth_personR As String, sAuth_personW As String
Attribute sAuth_personR.VB_VarUserMemId = 1073741844
Attribute sAuth_personW.VB_VarUserMemId = 1073741844

'是否控制供应商权限,供应商权限字符串 11-7-13 chenliangc
Public bVendor_ControlAuth As Boolean
Attribute bVendor_ControlAuth.VB_VarUserMemId = 1073741846
Public sAuth_vendorR As String, sAuth_vendorW As String
Attribute sAuth_vendorR.VB_VarUserMemId = 1073741847
Attribute sAuth_vendorW.VB_VarUserMemId = 1073741847

'是否控制部门权限,部门权限字符串 11-7-13 chenliangc
Public bDep_ControlAuth As Boolean
Attribute bDep_ControlAuth.VB_VarUserMemId = 1073741849
Public sAuth_depR As String, sAuth_depW As String
Attribute sAuth_depR.VB_VarUserMemId = 1073741850
Attribute sAuth_depW.VB_VarUserMemId = 1073741850

'user auth
Public bCheckUser As Boolean
Public isfromcon As Boolean

'是否控制客户权限,客户权限字符串 11-7-13 chenliangc
Public bCus_ControlAuth As Boolean
Attribute bCus_ControlAuth.VB_VarUserMemId = 1073741852
Public sAuth_CusR As String, sAuth_CusW As String
Attribute sAuth_CusR.VB_VarUserMemId = 1073741853
Attribute sAuth_CusW.VB_VarUserMemId = 1073741853

'是否控制仓库权限,仓库权限字符串 11-7-13 chenliangc
Public bWareHouse_ControlAuth As Boolean
Attribute bWareHouse_ControlAuth.VB_VarUserMemId = 1073741855
Public sAuth_WareHouseR As String, sAuth_WareHouseW As String
Attribute sAuth_WareHouseR.VB_VarUserMemId = 1073741856
Attribute sAuth_WareHouseW.VB_VarUserMemId = 1073741856

'是否控制货位权限,货位权限字符串 11-7-13 chenliangc
Public bPosition_ControlAuth As Boolean
Attribute bPosition_ControlAuth.VB_VarUserMemId = 1073741858
Public sAuth_PositionR As String, sAuth_PositionW As String
Attribute sAuth_PositionR.VB_VarUserMemId = 1073741859
Attribute sAuth_PositionW.VB_VarUserMemId = 1073741859

'是否控制操作员权限,操作员权限字符串 11-7-12 chenliangc
Public sAuth_cmaker As String
Attribute sAuth_cmaker.VB_VarUserMemId = 1073741861

'可见单据 11-7-12 chenliangc
Public sAuth_ALL As String
Attribute sAuth_ALL.VB_VarUserMemId = 1073741862
Public sAuth_AllList As String

Public sMakeAuth_ALL As String

Public sAuth_UnitR As String

Public pageCount As Long   '总页数
Attribute pageCount.VB_VarUserMemId = 1073741863
Public PageCurrent As Long  '当前页
Attribute PageCurrent.VB_VarUserMemId = 1073741864
Public lngVoucherID As Long ' 单据主表ID值
Attribute lngVoucherID.VB_VarUserMemId = 1073741865
Public MainTable As String '单据主表
Attribute MainTable.VB_VarUserMemId = 1073741866
Public DetailsTable As String '单据字表
Attribute DetailsTable.VB_VarUserMemId = 1073741867
Public HeadPKFld As String '主表主键字段
Attribute HeadPKFld.VB_VarUserMemId = 1073741868
Public MainView As String '表头视图
Attribute MainView.VB_VarUserMemId = 1073741869
Public DetailsView As String '表体视图
Attribute DetailsView.VB_VarUserMemId = 1073741870

Public TblName As String
Public ViewDetailName As String
Public ViewMainName As String
Public VoucherList As String '列表视图
Attribute VoucherList.VB_VarUserMemId = 1073741871
Public VoucherList2 As String '列表视图

Public conid As String
'记录当前操作状态，有增加，修改，正常
Public mOpStatus As OpStatus

'单据编号,制单人,单据日期字段名
Public strcCode, StrcMaker, StrdDate As String
Attribute strcCode.VB_VarUserMemId = 1073741872
Attribute StrcMaker.VB_VarUserMemId = 1073741872
Attribute StrdDate.VB_VarUserMemId = 1073741872
'审核人,日期,状态,关闭人,关闭日期,生单人,生单日期
Public StrcHandler, StrdVeriDate, StriStatus, StrCloseUser, StrdCloseDate, StrIntoUser, StrdIntoDate As String
Attribute StrcHandler.VB_VarUserMemId = 1073741875
Attribute StrdVeriDate.VB_VarUserMemId = 1073741875
Attribute StriStatus.VB_VarUserMemId = 1073741875
Attribute StrCloseUser.VB_VarUserMemId = 1073741875
Attribute StrdCloseDate.VB_VarUserMemId = 1073741875
Attribute StrIntoUser.VB_VarUserMemId = 1073741875
Attribute StrdIntoDate.VB_VarUserMemId = 1073741875

Public sID, sAutoId As Long '最大的主表id,和子表autoid
Attribute sID.VB_VarUserMemId = 1073741882
Attribute sAutoId.VB_VarUserMemId = 1073741882

Public sTmpTableName As String '定位/列表定位 临时表名
Public sGUID As String '临时表名需要用到的guid

Public strCellCode As String '返回编码和名称
Attribute strCellCode.VB_VarUserMemId = 1073741884
Public strCellName As String '返回编码和名称
Attribute strCellName.VB_VarUserMemId = 1073741885
Public symbol As String '汇率折算方式
Attribute symbol.VB_VarUserMemId = 1073741886

Public strWhere As String '过滤条件
Attribute strWhere.VB_VarUserMemId = 1073741887
Public strWhere2 As String '过滤条件

Public TimeStamp As String   '时间戳
Attribute TimeStamp.VB_VarUserMemId = 1073741888
Public OldTimeStamp As String    '当前单据的时间戳 chenliangc，也可在单据模板预置ufts后使用voucher.headertext("ufts")
Attribute OldTimeStamp.VB_VarUserMemId = 1073741889

'Public Const HelpFile = "\Help\寄售管理_zh-CN.chm"    '帮助文件路径
'Public Const HelpFile = "\HY\client\HY_DZ_JA_JYGH\电子行业插件帮助.chm"    '帮助文件路径
Public Const HelpFile = "\Help\ST_zh-CN.chm"  '帮助文件路径


'功能权限
Public Const AuthBrowse = "FYSL02050301" '浏览
Public Const AuthBrowselist = "PD01030101" '浏览


'时间戳判断结果
Public Const RecordDeleted = 1 '已被其他用户删除
Public Const RecordModified = 2 '已被其他用户修改
Public Const RecordNoChanged = 0 '正常,可以操作
Public Const RecordError = -1     '异常


'U8系统信息
Public m_SysInfor As clsSystemInfo
Attribute m_SysInfor.VB_VarUserMemId = 1073741890
Public clsInfor As Object 'Info_PU.ClsS_Infor
Attribute clsInfor.VB_VarUserMemId = 1073741891

'格式相关内容
Public m_sQuantityFmt As String '数量格式
Attribute m_sQuantityFmt.VB_VarUserMemId = 1073741892
Public m_sNumFmt As String      ' 数值格式
Attribute m_sNumFmt.VB_VarUserMemId = 1073741893
Public m_iExchRateFmt As String   ' 换算率
Attribute m_iExchRateFmt.VB_VarUserMemId = 1073741894
Public m_iRateFmt As String   ' 税率
Attribute m_iRateFmt.VB_VarUserMemId = 1073741895
Public m_sPriceFmt As String  ' 金额格式
Attribute m_sPriceFmt.VB_VarUserMemId = 1073741896
Public m_sPriceFmtSA As String  ' 金额格式（销售用）
Attribute m_sPriceFmtSA.VB_VarUserMemId = 1073741897

Public gcCreateType As String     '新增 单据类型
Attribute gcCreateType.VB_VarUserMemId = 1073741898
'Public tmpLinkTbl As String '关联单据联查时ccmdline传过来的临时表名 '单据联查 时 按钮状态控制 by zhangwchb 20110809


' * 非法字符
'##ModelId=42F6FF0701F4
Public Const gstrBAD_STRING As String = " ~`!@#$%^&*()-:;+={}[]'\|<>?,./"

'操作状态
Public Enum OpStatus
    ADD_MAIN = 1              '增加主集
    ADD_SUB = 2               '增加子集
    MODIFY_MAIN = 3           '修改主集
    MODIFY_SUB = 4            '修改子集
    DELETE_MAIN = 5           '删除主集
    DELETE_SUB = 6            '删除子集
    SHOW_MAIN = 7             '主集所有只读显示
    SHOW_SUB = 8              '子集所有只读显示
    SHOW_ALL = 9              '所有只读显示
    ADD_MAIN_AFTER = 10       '增加主集保存
    ADD_SUB_AFTER = 11        '增加子集保存
    MODIFY_MAIN_AFTER = 12    '修改主集保存
    MODIFY_SUB_AFTER = 13     '修改子集保存
    DELETE_MAIN_AFTER = 14    '删除主集保存
    DELETE_SUB_AFTER = 15     '删除子集示保存
    SHOW_NOTHING = 16         '没有对应单据
End Enum
' ***********************************************************
' * 调试模式
'
#If DEBUG_MODE = 1 Then
    Public Const g_blnDEBUG_MODE As Boolean = True
#Else
    Public Const g_blnDEBUG_MODE As Boolean = False
#End If

' ***********************************************************
' * 错误级别
' *
' * (用于定制错误信息的消息框)
Public Enum ErrorLevelConstants
    ufsELAllInfo = 0                ' 包含所有信息(友好信息、错误号、 错误源、错误描述)
    ufsELOnlyHeader = 1             ' 只包含友好提示信息(自定义)
    ufsELHeaderAndDescription = 2   ' 只包含错误描述和友好信息
End Enum
' *
' * 窗体动作
Public Enum FormActionConstants
    ufsFANew = 1    ' 新增
    ufsFAEdit = 2   ' 编辑
    ufsFAView = 0   ' 浏览
End Enum

Public Enum BillType       '推单类型
    销售 = 0
    采购 = 1
    库存 = 2
    应付 = 3
End Enum

Public Enum SaleVoucherType   '销售推单模式
    专用发票 = 0
    红字专用发票 = 1
    普通发票 = 2
    红字普通发票 = 3
    销售发货单 = 9
    退货单 = 10
    销售订单 = 12
    委托代销发货单 = 15
    委托代销退货单 = 16
    销售报价单 = 21
    
End Enum

Public Enum PUVoucherType   '采购推单模式
    采购请购单 = 0
    采购订单 = 1
    采购到货单 = 2
    采购发票 = 4
End Enum

Public Enum COVoucherType   '库存推单模式
    采购入库单 = 1
    采购发票 = 4
    其他入库单 = 8
    其他出库单 = 9
    产成品入库单 = 10
    材料出库单 = 11
    调拨单 = 12
    组装单 = 13
    拆卸单 = 14
    形态转换单 = 15
    盘点单 = 18
    销售出库单 = 32
    不合格品处理单 = 46
    起初不合格品 = 55
    调拨申请单 = 62
End Enum


'获取权限相关设置
Public Sub getAuthString(conn As ADODB.Connection)
    sAuth_ALL = "(1=1)"
    sMakeAuth_ALL = "(1=1)"
    sAuth_AllList = "(1=1)"
    Dim sauth_unit As String                               '表头单位权限-有可能是客户，供应商，部门
    sauth_unit = "(1=2)"

    bPerson_ControlAuth = False
    bInv_ControlAuth = False
    bVendor_ControlAuth = False
    bDep_ControlAuth = False
    bCus_ControlAuth = False
    bWareHouse_ControlAuth = False
    bPosition_ControlAuth = False
    bCheckUser = False

'    '存货权限
'    sAuth_invR = ""
'    sAuth_invW = ""
'    If LCase(getAccinformation("ST", "bInventoryCheck", conn)) = "true" Then
'        bInv_ControlAuth = True
'        sAuth_invR = GetRowAuth(conn.ConnectionString, g_oLogin.cUserId, "Inventory", , "R")
'        If sAuth_invR = "1=2" Then sAuth_invR = "-1"
'        sAuth_invW = GetRowAuth(conn.ConnectionString, g_oLogin.cUserId, "Inventory", , "W")
'        If sAuth_invW = "1=2" Then sAuth_invW = "-1"
'
'        sAuth_ALL = sAuth_ALL & IIf(sAuth_invR = "", "", " and (not exists(select top 1 iid from HY_DZ_BorrowOuts a INNER JOIN inventory b ON a.cInvCode=b.cinvcode where a.id=HY_DZ_BorrowOut.id AND isnull(b.iid,0) not in (" & sAuth_invR & ")))")
'        sMakeAuth_ALL = sMakeAuth_ALL & IIf(sAuth_invW = "", "", " and (exists(select top 1 iid from HY_DZ_BorrowOuts a INNER JOIN inventory b ON a.cInvCode=b.cinvcode where a.id=HY_DZ_BorrowOut.id AND isnull(b.iid,0) in (" & sAuth_invW & ")))")
'        sAuth_AllList = sAuth_AllList & IIf(sAuth_invR = "", "", " and iid in (" & sAuth_invR & ")")
'
'    End If
'
'    '制单人权限
'    sAuth_cmaker = ""
'    If LCase(getAccinformation("ST", "bOperatorCheck", conn)) = "true" Then
'        bCheckUser = True
'        sAuth_cmaker = GetRowAuth(conn.ConnectionString, g_oLogin.cUserId, "user", , "R")
'        sAuth_ALL = sAuth_ALL & IIf(sAuth_cmaker = "", "", " and (cmaker in (" & sAuth_cmaker & ")) ")
'        sMakeAuth_ALL = sMakeAuth_ALL & IIf(sAuth_cmaker = "", "", " and (cmaker in (" & sAuth_cmaker & ")) ")
'        sAuth_AllList = sAuth_AllList & IIf(sAuth_cmaker = "", "", " and (cmaker in (" & sAuth_cmaker & ")) ")
'    End If

    '    '业务员权限
    '    sAuth_personR = ""
    '    sAuth_personW = ""
    '    If LCase(getAccinformation("ST", "bCheckPersonAuth", conn)) = "true" Then
    '        bPerson_ControlAuth = True
    '        sAuth_personR = GetRowAuth(conn.ConnectionString, g_oLogin.cUserId, "Person", , "R")
    '        sAuth_personW = GetRowAuth(conn.ConnectionString, g_oLogin.cUserId, "Person", , "W")
    '
    '        sAuth_ALL = sAuth_ALL & IIf(sAuth_personR = "", "", " and ( isnull(cpersoncode,'')='' or cpersoncode in (" & sAuth_personR & ")) ")
    '
    '    End If


    '部门权限
    sAuth_depR = ""
    sAuth_depW = ""
    If LCase(getAccinformation("ST", "bDepartmentCheck", conn)) = "true" Then
        bDep_ControlAuth = True
        sAuth_depR = GetRowAuth(conn.ConnectionString, g_oLogin.cUserId, "Department", , "R")
        sAuth_depW = GetRowAuth(conn.ConnectionString, g_oLogin.cUserId, "Department", , "W")
        If sAuth_depR = "1=2" Then sAuth_depR = "'|'"
        If sAuth_depW = "1=2" Then sAuth_depW = "'|'"
        sAuth_ALL = sAuth_ALL & IIf(sAuth_depR = "", "", " and ( isnull(chdepartcode,'')='' or chdepartcode in (" & sAuth_depR & ")) ")
        sMakeAuth_ALL = sMakeAuth_ALL & IIf(sAuth_depR = "", "", " and ( isnull(chdepartcode,'')='' or chdepartcode in (" & sAuth_depR & ")) ")
        sAuth_AllList = sAuth_AllList & IIf(sAuth_depR = "", "", " and ( isnull(chdepartcode,'')='' or chdepartcode in (" & sAuth_depR & ")) ")
        
        sauth_unit = sauth_unit

    Else
        sauth_unit = sauth_unit

    End If

'    '客户权限
'    sAuth_CusR = ""
'    sAuth_CusW = ""
'    If LCase(getAccinformation("ST", "bCustomerCheck", conn)) = "true" Then
'        bCus_ControlAuth = True
'        sAuth_CusR = GetRowAuth(conn.ConnectionString, g_oLogin.cUserId, "Customer", , "R")
'        sAuth_CusW = GetRowAuth(conn.ConnectionString, g_oLogin.cUserId, "Customer", , "W")
'
'        '11.0 客户iid改成nvarchar类型了
'        If sAuth_CusR = "1=2" Then sAuth_CusR = "'-1'"
'        If sAuth_CusW = "1=2" Then sAuth_CusW = "'-1'"
'        sauth_unit = sauth_unit & " or " & IIf(sAuth_CusR = "", "(ctype='客户')", " (ctype='客户' and ( isnull(bObjectCode,'')='' or bObjectCode in (select ccuscode from customer where iid in (" & sAuth_CusR & ")))) ")
'    Else
'        sauth_unit = sauth_unit & " or " & " (ctype='客户') "
'    End If
'
'    '供应商权限
'    sAuth_vendorR = ""
'    sAuth_vendorW = ""
'    If LCase(getAccinformation("ST", "bVendorCheck", conn)) = "true" Then
'        bVendor_ControlAuth = True
'        sAuth_vendorR = GetRowAuth(conn.ConnectionString, g_oLogin.cUserId, "Vendor", , "R")
'        sAuth_vendorW = GetRowAuth(conn.ConnectionString, g_oLogin.cUserId, "Vendor", , "W")
'        If sAuth_vendorR = "1=2" Then sAuth_vendorR = "-1"
'        If sAuth_vendorW = "1=2" Then sAuth_vendorW = "-1"
'        sauth_unit = sauth_unit & " or " & IIf(sAuth_vendorR = "", "(ctype='供应商')", " (ctype='供应商' and ( isnull(bObjectCode,'')='' or bObjectCode in (select cvencode from vendor where iid in (" & sAuth_vendorR & ")))) ")
'    Else
'        sauth_unit = sauth_unit & " or " & " (ctype='供应商') "
'    End If
'
'    '仓库权限
'    sAuth_WareHouseR = ""
'    sAuth_WareHouseW = ""
'    If LCase(getAccinformation("ST", "bWarehouseCheck", conn)) = "true" Then
'        bWareHouse_ControlAuth = True
'        sAuth_WareHouseR = GetRowAuth(conn.ConnectionString, g_oLogin.cUserId, "Warehouse", , "R")
'        sAuth_WareHouseW = GetRowAuth(conn.ConnectionString, g_oLogin.cUserId, "Warehouse", , "W")
'        If sAuth_WareHouseR = "1=2" Then sAuth_WareHouseR = "'|'"
'        If sAuth_WareHouseW = "1=2" Then sAuth_WareHouseW = "'|'"
'        sAuth_ALL = sAuth_ALL & IIf(sAuth_WareHouseR = "", "", " and (not exists(select top 1 1 from HY_DZ_BorrowOuts a where a.id=HY_DZ_BorrowOut.id and isnull(a.cwhcode,'')<>'' and a.cwhcode not in (" & sAuth_WareHouseR & ")))")
'        sMakeAuth_ALL = sMakeAuth_ALL & IIf(sAuth_WareHouseW = "", "", " and (exists(select top 1 1 from HY_DZ_BorrowOuts a where a.id=HY_DZ_BorrowOut.id and (ISNULL(a.cwhcode,N'')=N'' OR a.cwhcode in (" & sAuth_WareHouseW & "))))")
'
'        sAuth_AllList = sAuth_AllList & IIf(sAuth_WareHouseR = "", "", " and ( isnull(cwhcode,'')='' or cwhcode in (" & sAuth_WareHouseR & ")) ") '" and (not exists(select top 1 1 from HY_DZ_BorrowOuts a where a.id=HY_DZ_BorrowOut.id and a.cwhcode not in (" & sAuth_WareHouseR & ")))")
'    End If
'
'    '货位权限
'    sAuth_PositionR = ""
'    sAuth_PositionW = ""
'    If LCase(getAccinformation("ST", "bPostionCheck", conn)) = "true" Then
'        bPosition_ControlAuth = True
'        sAuth_PositionR = GetRowAuth(conn.ConnectionString, g_oLogin.cUserId, "Position", , "R")
'        sAuth_PositionW = GetRowAuth(conn.ConnectionString, g_oLogin.cUserId, "Position", , "W")
'        If sAuth_PositionR = "1=2" Then sAuth_PositionR = "'|'"
'        If sAuth_PositionW = "1=2" Then sAuth_PositionW = "'|'"
'        sAuth_ALL = sAuth_ALL & IIf(sAuth_PositionR = "", "", " and (not exists(select top 1 1 from HY_DZ_BorrowOuts a where a.id=HY_DZ_BorrowOut.id and (isnull(a.cposition,'')<>'' and a.cposition not in (" & sAuth_PositionR & "))))")
'        sMakeAuth_ALL = sMakeAuth_ALL & IIf(sAuth_PositionW = "", "", " and ( exists(select top 1 1 from HY_DZ_BorrowOuts a where a.id=HY_DZ_BorrowOut.id and (isnull(a.cposition,'')='' or a.cposition in (" & sAuth_PositionW & "))))")
'
'        sAuth_AllList = sAuth_AllList & IIf(sAuth_PositionR = "", "", " and ( isnull(cposition,'')='' or cposition in (" & sAuth_PositionR & ")) ")
'    End If
'
'    '12.0备机借出，单位类型支持最终用户
'    sauth_unit = sauth_unit & " or " & " (ctype='最终用户') "
    
    '12.0屏蔽样品借出（生单不能由此限制）
    'sauth_unit = "(" & sauth_unit & ") and isnull(cborrowouttype,0)!=1"
    
    If sauth_unit = "(1=2)" Then sauth_unit = "(1=1)"
    sAuth_ALL = sAuth_ALL & " and (" & sauth_unit & ")"
    sAuth_AllList = sAuth_AllList & " and (" & sauth_unit & ")"
    sMakeAuth_ALL = sMakeAuth_ALL & " and (" & sauth_unit & ")"
    sAuth_UnitR = sauth_unit
    
End Sub

'检查操作员权限 CheckUserAuth
'connstr 数据库连接串
'selfuserid 本身的操作员userid ,直接取 login.cuserid
'objuserid 单据的制单人，取单据上的制单人，此处直接是username
'cfunctionid 读的权限已经通过getauthstring控制了, 需要控制删改-“W”，审核-“V”,弃审-“U”，关闭-“C”，撤销-“A”
Public Function CheckUserAuth(connstr As String, selfuserid As String, objuserid As String, cfunctionid As String) As Boolean
'    Dim authsrv As Object
'    Set authsrv = CreateObject("U8RowAuthsvr.clsRowAuth")
    If bCheckUser = False Then
        CheckUserAuth = True
        Exit Function
    End If

    Dim authsrv As New U8RowAuthsvr.clsRowAuth
    If Not authsrv.Init(connstr, selfuserid, False, "ST") Then Exit Function
    CheckUserAuth = authsrv.IsHoldAuth("user", objuserid, , cfunctionid)
End Function


Public Function IsWFControlled() As Boolean
 
    On Error GoTo ErrHandle
    Dim cBizObjectId As String
    Dim bWFControlled As Boolean
    Dim errMsg As String
    Dim o As Object
    cBizObjectId = "HYJCGH001"
    IsWFControlled = False
    Set o = CreateObject("SCMWorkFlowCommon.clsWFController")
    'If o.getIsWfControl(mologin.AccountConnection, cBizObjectId, cBizObjectId & ".Submit", mologin.OldLogin.cIYear, mologin.OldLogin.cAcc_Id, bWFControlled, errMsg) Then
    If o.getIsWFHasActivated(mologin.AccountConnection, cBizObjectId, cBizObjectId & ".Submit", bWFControlled, errMsg) Then
       IsWFControlled = bWFControlled
    End If
    Set o = Nothing
    Exit Function
    
ErrHandle:
    IsWFControlled = False
    Exit Function
End Function
