VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public coutsign        As String               '外部凭证业务类型'√
Public coutno_id       As String                 '外部凭证业务号'√
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

Public bPrepay        As Boolean
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

Public cBlueID         As String               ' 的ID号
Public cTableName      As String               '用于存数据来源的表名
Public cFieldName      As String               '用于存数据来源的表的字段名
Public ibillno_id      As String               '用于存原单据ID号
Public cProcNo         As String               '此处用于存储制单批号
Public cProcStyle      As String
Public cCancelNo       As String
Public bTaxFlag        As Boolean
Public iLink           As Long                 '存iflag标志


'////////////////////////////////////////////////////
Public Version As Long
Private o_dbSales As UfDatabase
Public clsItem As ItemDef.clsXmDY
Public strVoucherName As String      '显示单据名
Private m_PassIsTrue As Boolean


Public Property Get PassIsTrue() As Boolean
    PassIsTrue = m_PassIsTrue
End Property

Public Property Let PassIsTrue(ByVal New_PassIsTrue As Boolean)
    m_PassIsTrue = New_PassIsTrue
End Property

Private Sub Class_Terminate()
    On Error Resume Next
    Set clsItem = Nothing
End Sub

Public Function dbSales() As UfDatabase
    If o_dbSales Is Nothing Then
        Set o_dbSales = New UfDatabase
        o_dbSales.OpenDatabase m_login.UfDbName, False, False, ";PWD=" & m_login.SysPassword
    End If
    Set dbSales = o_dbSales
End Function
