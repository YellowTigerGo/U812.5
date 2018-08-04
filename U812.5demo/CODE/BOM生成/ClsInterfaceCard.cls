VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsInterfaceCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements UFPortalProxyInterface.ILoginable


Public Property Get Business() As Object
    Set Business = goBusiness
End Property

Public Property Set Business(ByRef bBus As Object)
    Set goBusiness = bBus  '获得business对象。
End Property

Public Function ILoginable_Login() As Boolean
    If ((Not goBusiness Is Nothing) And gbLogined = False) Then
        Set goLogin = goBusiness.GetVBclsLogin() '获得login对象
        If gConn.State = 1 Then gConn.Close
        gConn.CursorLocation = adUseClient
        gConn.ConnectionString = goLogin.UfDbName
        gConn.Open
        
        Set g_oLogin = goBusiness.GetVBclsLogin() '获得login对象
        If g_Conn.State = 1 Then g_Conn.Close
        g_Conn.CursorLocation = adUseClient
        g_Conn.ConnectionString = goLogin.UfDbName
        g_Conn.Open
        
        gbLogined = True
        ILoginable_Login = True
    Else
        ILoginable_Login = Not goBusiness Is Nothing
    End If

End Function

Public Function ILoginable_LogOff() As Boolean
    '程序退出前,添加自己的代码

    ILoginable_LogOff = True '本行代码必须。
    gbLogined = False '本行代码必须。
    Unload FrmList
End Function

Public Function ILoginable_CallFunction(ByVal cMenuId As String, ByVal cMenuName As String, ByVal cAuthId As String, ByVal cCmdLine As String) As Object
    '开始执行业务代码
    'cMenuID、cmenuName、cAuthId、cCmlLine均为前述在UA_Menu表中注册的cMenu_Id、cMenu_Name、cAuth_Id、Paramters字段的值。
    
    '加密判断 08-03-25
    Dim Encryption As Object
    Set Encryption = New clsEncryption
    
    Dim strAuditSQL As String  ' by zhangwchb 20110907 合并任务打开时不显示过滤界面
    
    
    
    If ZwTaskExec(goLogin, AuthBrowselist, 1) = False Then Exit Function
     
    
'        If DateDiff("D", "2016-01-30", goLogin.CurDate) >= 0 Then
'        MsgBox "演示期已到!", vbCritical, "项目管理"
'        Exit Function
'    End If
'
'    If DateDiff("D", "2016-01-30", Date) >= 0 Then
'        MsgBox "演示期已到!", vbCritical, "项目管理"
'        Exit Function
'    End If
    
    GlobalInit goLogin '初始化系统信息 chenliangc
    InitMulText
    '    获取权限字段
    Call getAuthString(gConn)

    '控制过滤、参照等不会因为数据权限而无法显示数据 11-7-12
    goLogin.AuthString = "warehouse=True,department=True,inventory=True,person=False,user=True,vendor=True,purchasetype=False,customer=True,position=True"

    
    Dim vfd As Object
    
    '11.0菜单发布 wangfb
    '传入的cmdline先用&&@@&&分割 前部分是原来的 后部分是菜单发布传入的
    Dim cPara
    Dim cPara1
    Dim cCmdLineMenu As String '菜单发布过滤方案传入的命令
    Dim sMenuPubFilter As String
    sMenuPubFilter = ""
    
    cPara = Split(cCmdLine, "&&@@&&")
    If UBound(cPara) > 0 Then
        '原命令和菜单发布传入的分开
        cCmdLine = cPara(0)
        cCmdLineMenu = cPara(1)
    End If
        
    '处理菜单发布传入的cmdline信息 11.0
    cPara1 = Split(cCmdLineMenu, vbTab)
    If UBound(cPara1) >= 1 Then
        '菜单发布的源menuid 和 过滤方案id
        cMenuId = Trim(cPara1(0))
        sMenuPubFilter = Trim(cPara1(1))
    End If
    
    '加载过滤条件
    If InStr(1, cCmdLine, "audit." & UCase$(gstrCardNumber)) <= 0 Then
    
        If GetFilterList(goLogin, , sMenuPubFilter) = False Then
            Call ZwTaskExec(goLogin, AuthBrowselist, 0)
            Exit Function
        End If
    
    Else
    
        strAuditSQL = Split(cCmdLine, vbTab)(1)
        strWhere = IIf(strAuditSQL = "", sAuth_AllList, " id in(" & strAuditSQL & ") and " & sAuth_AllList)
        
    End If
    
    If sGUID = "" Then
        sGUID = CreateGUID
    End If
    
'    If Not (goBusiness Is Nothing) Then
'      Set vfd = goBusiness.CreateFormEnv(sGUID, frmScan)
'      Call goBusiness.ShowForm(frmScan, "EF", sGUID, False, True, vfd)
'    Else
'      frmScan.Show
'      frmScan.ZOrder 0
'    End If
    Dim frm As New frmScan
    frm.Show vbModal
    frm.ZOrder 0
    Call ZwTaskExec(goLogin, AuthBrowselist, 0)

End Function



