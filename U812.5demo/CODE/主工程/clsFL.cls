VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'   软件著作权: 用友软件集团
'   系统名称：应收、应付系统
'   功能说明: 凭证分录定义
'   作者：翟万达
'------------------------------------------------------------------------------------
Public cvouchid        As String               '单据号
Public ibillno_id      As String               '用于存原单据ID号
Public cCancelNo       As String
Public cProcNo         As String               '此处用于存储制单批号
Public cProcStyle      As String
Public cTableName      As String               '用于存数据来源的表名
Public cFieldName      As String               '用于存数据来源的表的字段名
Public cBlueID         As String               '若是红票，存对应兰票的ID号
Public bTaxFlag        As Boolean
Public cmergeno        As String               '存合并项目标志的小编号
Public iLink           As Long                 '存iflag标志
Public dBillDate       As Date                 '制单日期

'自定义项1-16
Public cDefine1        As Variant
Public cDefine2        As Variant
Public cDefine3        As Variant
Public cDefine4        As Variant
Public cDefine5        As Variant
Public cDefine6        As Variant
Public cDefine7        As Variant
Public cDefine8        As Variant
Public cDefine9        As Variant
Public cDefine10       As Variant
Public cDefine11       As Variant
Public cDefine12       As Variant
Public cDefine13       As Variant
Public cDefine14       As Variant
Public cDefine15       As Variant
Public cDefine16       As Variant

Public coutsign        As String               '外部凭证业务类型'√
Public coutno_id       As String               '外部凭证业务号'√
Public inid            As Long                 '行号'
Public doutdate        As Date                 '外部凭证单据日期'
Public coutbillsign    As String               '外部凭证单据类型'
Public coutid          As String               '外部凭证单据号'
Public cSign           As String               '凭证类别字(与凭证主表关联)'
Public idoc            As Long                 '附单据数'
Public cDigest         As String               '摘要'
Public cCode           As String               '科目编码(与科目主表关联)'
Public md              As Currency             '金额借方'√
Public mc              As Currency             '金额贷方'√
Public md_f            As Currency             '外币借方金额(若无外币,写0)'√
Public mc_f            As Currency             '外币贷方金额(若无外币,写0)'√
Public nfrat           As Double               '汇率(若无外币,写0)'√
Public nd_s            As Double               '数量借方(若无数量,写0)'√
Public nc_s            As Double               '数量贷方(若无数量,写0)'√
Public csettle         As String               '结算方式(与结算方式目录关联)'
Public cn_id           As String               '票号'
Public dt_date         As Date                 '票号发生日期'
Public cdept_id        As String               '部门编码(与部门目录关联)'
Public cPerson_id      As String               '个人编码(与职员目录关联)'
Public ccus_id         As String               '客户编码(与客户目录关联)'
Public csup_id         As String               '供应商编码(与供应商目录关联)'
Public citem_id        As String               '项目编码(与项目目录关联)'
Public cItem_Class     As String               '项目大类(与项目大类关联)'
Public cname           As String               '业务员'

Public bPrepay         As Boolean
Public bvouchedit      As Boolean              '凭证是否可修改(True_可修改,False_不可修改)'
Public bvouchAddordele As Boolean              '凭证分录是否可增删(True_可增删,False_不可增删)'
Public bvouchmoneyhold As Boolean              '凭证合计金额是否保值(True_必须保值,False_可不保值)'
Public bvalueedit      As Boolean              '分录数值(金额/数量/外币)是否可修改(True_可修改,False_不可修改)'
Public bcodeedit       As Boolean              '分录科目是否可修改(True_可修改,False_不可修改)'
Public ccodecontrol    As String               '分录受控科目(Null_均不可用,****_均可用,!_指定不可用,ID_指定可用)'
Public bPCSedit        As Boolean              '分录往来项(个人/客户/供应商)是否可修改(True_可修改,False_不可修改)'
Public bDeptedit       As Boolean              '分录部门是否可修改(True_可修改,False_不可修改)'
Public bItemedit       As Boolean              '分录项目是否可修改(True_可修改,False_不可修改)'
Public bCusSupInput    As Boolean              '分录往来项(客户/供应商)是否必输(True_必输,False_不必输)'

'by zzc 2008-01-21
Public BG_AuditRemark    As String
Public cBudgetBuffer     As String
Public CErrReason        As String
Public iBG_ControlResult As Integer
Public iPeriod           As Integer               '会计期间
Public isignseq          As Integer               '凭证类别排序号
Public ino_id            As Integer               '凭证编号
Public cbill             As String                '制单人
Public ccheck            As String                '审核人
Public cbook             As String                '记账人
Public ibook             As Integer               '记账标志
Public ccashier          As String                '出纳签字人
Public iflag             As Integer               '凭证标志
Public ctext1            As String                '凭证头自定义项1
Public ctext2            As String                '凭证头自定义项2
Public cExch_Name        As String                '币种名称
Public ccode_equal       As String                '对方科目编码
Public iflagbank         As Integer               '银行账两清标志
Public iflagPerson       As Integer               '往来账两清标志
Public bdelete           As Integer               '是否核销
Public coutaccset        As String                '外部凭证账套号
Public ioutyear          As Integer               '外部凭证会计年度
Public coutsysname       As String                '外部凭证系统名称
Public coutsysver        As String                '外部凭证系统版本号
Public doutbilldate      As Date                  '外部凭证制单日期
Public ioutperiod        As Integer               '外部凭证会计期间
Public dReceive          As Date                  'u8/nc传输接受日期
Public cWLDZFlag         As String                '公司对账类型
Public dWLDZTime         As Date                  '公司对账日期
Public bFlagOut          As Integer               '公司对帐是否导出过对帐单
Public iBG_OverFlag      As Integer               '超预算标识
Public cBG_Auditor       As String                '超预算审批人
Public dBG_AuditTime     As Date                  '超预算审批时间
Public cBG_AuditOpinion  As String                '超预算审批结论
Public bWH_BgFlag        As Integer               '预算审核标志
Public ssxznum           As Integer               '实时核销号

