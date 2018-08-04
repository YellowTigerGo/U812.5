VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IINterface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'--------------------------------------------
' 文 件 名: IINterface.cls
'
' 功 能: GSP档案的接口类
'
' 作 者: 王良彬    时  间:     2002
' 北就用友软件集团 版权所有 Copyright(c) 2002
'--------------------------------------------

Option Explicit

Dim objM    As Object
'new portal xiaojing
Private m_business As Object
Public Property Set o_business(V As Object)
    Set m_business = V
    Set g_business = m_business
End Property
Public Property Get o_business() As Object
    Set o_business = m_business
End Property


'-----------------------------------------------------------
'功能：设置帮助存放路径
'
'参数：path帮助存放路径
'
'返回：
'
'-----------------------------------------------------------
Public Sub putPath(ByVal path As String)
    AppPath = path
End Sub


'-----------------------------------------------------------
'功能：设置主窗体句柄
'
'参数：hwnd主窗体句柄
'
'返回：
'
'-----------------------------------------------------------
Public Sub putWnd(ByVal hwnd As Long)
    mhwndMain = hwnd
End Sub

'-----------------------------------------------------------
'功能：初始化修改、审核权限信息
'
'参数：bCanModifyOther 是否可修改别人编辑的档案；bCanAuditOwn 是否可审核自己编辑的档案
'
'返回：
'
'-----------------------------------------------------------
Public Sub Init(bCanModifyOther As Boolean, bCanAuditOwn As Boolean)
    mbCanModifyOther = bCanModifyOther
    mbCanAuditOwn = bCanAuditOwn
End Sub


'-----------------------------------------------------------
'功能：初始化DbGSP、oLogin等， 显示档案、分类
'
'参数：DbGSP数据库操作公共类，oLogin登陆Login公共类，sCaption显示标题，tbl toolbar风格，
'      sAdd_auth编辑权限，sAudit_auth审核权限，sUAudit_auth弃审权限，sAppro_auth审批权限，sUAppro_auth弃批权限
'
'返回：
'
'-----------------------------------------------------------
Public Sub Show(DbGSP As Object, oLogin As Object, lType As Long, sCaption As String, ByRef tbl As Long, sQuery_auth As String, sAdd_auth As String, sAudit_auth As String, sUAudit_auth As String, sAppro_auth As String, Optional sUAppro_auth As String)
    '设置Login对象
    Set g_oLogin = oLogin
    If g_oLogin Is Nothing Then
        ShowMsg "没有传入Login对象！"
        Exit Sub
    End If
    '设置DbGSP对象
    Set g_DbGSP = DbGSP
    If g_DbGSP Is Nothing Then
        ShowMsg "没有传入DbGSP对象！"
        Exit Sub
    End If
    g_DbGSP.DbConnect.CursorLocation = adUseClient         '设置数据库连接游标类型
    Set SrvDB = CreateObject("U8SrvTrans.IClsCommon")
    If SrvDB Is Nothing Then
        ShowMsg "没有找到连接数据库的动态库！"
        Exit Sub
    End If
    SrvDB.Init g_oLogin.UfDbName
    If SrvDB.Status = False Then
        ShowMsg "数据库连接失败！" + vbCrLf + "请检查数据库连接！"
        Exit Sub
    End If
    If g_oPub Is Nothing Then Set g_oPub = CreateObject("U8Pub.IPub")
    If g_oPub Is Nothing Then
        ShowMsg "没有找到公共动态库！"
        Exit Sub
    End If
    '///初始化公共参数
    Set AdoCnn = New ADODB.Connection
    AdoCnn.CursorLocation = adUseClient
    AdoCnn.Open g_oLogin.UfDbName
    
    g_DbGSP.DbConnect.CursorLocation = adUseClient
    App.HelpFile = AppPath
'    App.HelpFile = IIf(Right(App.path, 1) = "\", App.path, App.path & "\") & "出版管理系统.chm"
    mbYearEnd = bYearEnd        '年度封账
    Set clsAuth = New U8RowAuthsvr.clsRowAuth
    clsAuth.Init g_oLogin.UfDbName, g_oLogin.cUserId, , "GS"
    '///////
    mlngType = lType
    
    '判断是否有查询权限
    If Not UA_Task(sQuery_auth) Then Exit Sub
    
    Select Case lType
    Case 0, 1, 2, 3, 4, 6, 7, 8, 9, 10, 11, 12, 13, 14          '显示分类
        Set objM = New clsMannual
        Call objM.DOModal(lType, sCaption, tbl, sAdd_auth)
    Case 5              '显示合同条款
        Set objM = New clsArchive
        Call objM.DOModal(lType, sCaption, tbl, sAdd_auth)
    End Select
    
    tbl = TBLStyle
    '释放对象资源
    Set objM = Nothing
    
    Set g_oPub = Nothing
    Set g_oLogin = Nothing
    Set SrvDB = Nothing
End Sub




'<属性设置

'-----------------------------------------------------------
'功能：是否可以修改别人单据
'
'参数：vData 是否可修改别人编辑的档案
'
'返回：
'
'-----------------------------------------------------------
Public Property Let ChangeOther(vData As Variant)
    mbolChangeOther = vData
End Property
'-----------------------------------------------------------
'功能：是否可以审核自己单据
'
'参数：vData 是否可以审核自己单据
'
'返回：
'
'-----------------------------------------------------------
Public Property Let AuditOwner(vData As Variant)
    mbolAuditOwner = vData
End Property

'-----------------------------------------------------------
'功能：当前操作员
'
'参数：vData 当前操作员
'
'返回：
'
'-----------------------------------------------------------
Public Property Let Operator(vData As Variant)
    mstrOperator = vData
End Property
'/属性设置>


Private Sub Class_Initialize()
    Set ctlDate = New CalendarAPP.ICaleCom               '初始化日期控件类
End Sub

Public Sub SetBusiness(vValue As Object)
    Set g_business = vValue
End Sub
