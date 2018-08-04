VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProductFacade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements UFPortalProxyInterface.ILoginable

'接口ILoginable函数CallFunction
Private Function ILoginable_CallFunction(ByVal cmenuid As String, ByVal cmenuname As String, ByVal cAuthId As String, ByVal cCmdLine As String) As Object
    Wrtlog "ILoginable_CallFunction 方法开始！"
    Set ILoginable_CallFunction = CallFunction(cmenuid, cmenuname, cAuthId, cCmdLine)
End Function

'接口ILoginable函数Login
Private Function ILoginable_Login() As Boolean
    
    Set cls_Public = CreateObject("UF_Public_base.cls_log")
    If ((Not g_business Is Nothing) And g_bLogined = False) Then
        Dim oLoginTmp As Object
        Set oLoginTmp = g_business.GetVBclsLogin()
        Set m_login = oLoginTmp
        m_login.AuthString = "Department =TRUE "
        ILoginable_Login = Login(oLoginTmp)
        If ILoginable_Login = True Then
            g_bLogined = True
        End If
    Else
        ILoginable_Login = True
    End If
End Function

'接口ILoginable函数LogOff
Private Function ILoginable_LogOff() As Boolean
    ILoginable_LogOff = True
    If (g_bLogined = True) Then
        ILoginable_LogOff = LogOff()
        If (ILoginable_LogOff = True) Then
            g_bLogined = False
        End If
    End If
End Function

'获取门给定的业务属性g_business
Public Property Get Business() As Object
    Set Business = g_business
End Property

'门户给定业务属性Business
Public Property Set Business(busVal As Object)
    Set g_business = busVal
End Property



'门户给定业务属性Business
Public Sub Set_Business(busVal As Object)
    Set g_business = busVal
End Sub

'设置窗口管理器对象，用于协调子窗体与原MDI主窗体的交互，g_oMainFrmProxy为门户主窗体代理全局对象
Public Function SetFrmProxy(ByRef oFrmProxy As Object) As Boolean
    Set g_oMainFrmProxy = oFrmProxy
End Function
 

'设置Login对象，初始化产品，g_oLogin为产品使用的全局U8Login对象，此接口必须显式返回True或False
Public Function Login(oLogin As Object) As Boolean
    Dim Rst As New ADODB.Recordset
    Dim strsql As String
    Dim id As Long
    Set m_login = oLogin
    m_login.AuthString = "Department =TRUE "
    Wrtlog "login 方法开始！"
 
    If DBconn Is Nothing Then
        Set DBconn = New ADODB.Connection
        DBconn.ConnectionTimeout = 600
        DBconn.CommandTimeout = 1200
    End If
    If DBconn.State = 1 Then DBconn.Close
    DBconn.Open m_login.UfDbName
    
    
'   --更新附件中的字段长度"
'    strsql = "alter table VoucherAccessories drop constraint PK_VoucherAccessories  " & vbCrLf
'    strsql = strsql & "GO " & vbCrLf
'    strsql = "ALTER TABLE VoucherAccessories ALTER COLUMN vouchertypeid nvarchar(20) not null  "
'    DBconn.Execute strsql

    
    strsql = "/******************Contract structure *********************/" & vbCrLf
    strsql = strsql & " Print '1 dbo.EF_Log  标准EF_Log 日志表 ...'" & vbCrLf
    strsql = strsql & " if not exists (select * from sysobjects where id = object_id('dbo.EF_Log') and sysstat & 0xf = 3)" & vbCrLf
    strsql = strsql & " BEGIN " & vbCrLf
    strsql = strsql & " CREATE TABLE [EF_Log] (" & vbCrLf
    strsql = strsql & "    [id] [int] identity not null ," & vbCrLf
    strsql = strsql & "    [loguser]   [nvarchar](50) NULL," & vbCrLf
    strsql = strsql & "    [logType]   [nvarchar](50) NULL," & vbCrLf
    strsql = strsql & "    [logtext]   [nvarchar](4000) NULL," & vbCrLf
    strsql = strsql & "    [logdatetime] [datetime] NULL" & vbCrLf
    strsql = strsql & "    CONSTRAINT [PK_EF_Log] PRIMARY KEY  CLUSTERED" & vbCrLf
    strsql = strsql & "    (" & vbCrLf
    strsql = strsql & "        [ID]" & vbCrLf
    strsql = strsql & "    )  ON [PRIMARY]" & vbCrLf
    strsql = strsql & "   ) ON [PRIMARY]" & vbCrLf
    strsql = strsql & " End " & vbCrLf
'    strsql = strsql & " GO"
    DBconn.Execute strsql
    
    strsql = "select isnull(MAX(id)-MIN(id),0) ,MIN(id)  as minid from EF_Log     "
    If Rst.State <> 0 Then Rst.Close
    Rst.CursorLocation = adUseClient
    Rst.Open strsql, DBconn, 3, 4
    If Rst.Fields(0) > 10000 Then
        strsql = "delete EF_Log where ID<" & Rst.Fields(1) + 5000
        DBconn.Execute strsql
    End If
    
    cls_Public.WrtDBlog DBconn, m_login.cUserId, "EFmain", "系统登陆成功！"
'    Public Sub WrtDBlog(adoconn As Object, Optional loguser As String = "", Optional logType As String = "", Optional logtext As String = "", Optional logdatetime As Date = Now)
    
    
    If Not m_login.TaskExec(m_login.cSub_Id, 10) Then
        If bDemoOverTime Then
            Login = False
            MsgBox "[出版行业插件]登陆错误：" & vbCr & "账套[" & m_login.cacc_id & "]" & "年度[" & m_login.cIYear & "]" & "的演示期限已到期！", vbExclamation, "UFIDA ERP-U8"
            Exit Function
        End If
    Else
        Dim bRe As Boolean
        If m_login.GetAccInfo(10000, bRe) Then
        End If
        If Not bRe Then
              If bDemoOverTime Then
                  Login = False
                  MsgBox "[出版行业插件]登陆错误：" & vbCr & "加密已饱和！", vbExclamation, "UFIDA ERP-U8"
                  Exit Function
              End If
         End If
    End If

    
    Call Main
    
    '处理870不能对老报表的支持,在这里初始化这两个对象,这段代码不允许变动/////////////////////////////
'    On Error Resume Next
'    If pRepSysinfo Is Nothing Then
'      Set pRepSysinfo = New clsSysInterface
'      pRepSysinfo.systemId = m_Login.cSub_Id
'      pRepSysinfo.ServerRunmode = False
'      Set pRepSysinfo.objU8Login = m_Login
'      pRepSysinfo.InitInterFace m_Login.cSub_Id, , , DBConn, , , , , , , , , m_Login
'      Set pRepLst = pRepSysinfo.GetReportEngine()
'      If pRepLst Is Nothing Then MsgBox " pRepLst Is Nothing "
'      pRepSysinfo.HelpFile = App.HelpFile
'    End If
    '/////////////////////////////////////////////////////////////////////////////////////////////////
    bLoadmain = True
    Login = True
End Function

'发送注销消息
Public Function LogOff() As Boolean
    Wrtlog "LogOff 方法开始！"
    Unload frmMain
    g_bCanExit = True
    LogOff = True
End Function

'调用子窗体接口，与原左树菜单控件接口一致，gd_frmMain为原MDI主窗口对象
Public Function CallFunction(ByVal cmenuid As String, ByVal cmenuname As String, ByVal cAuthId As String, ByVal cCmdLine As String) As Object
    g_FormbillShow = False
    frmMain.OnCommand cmenuid, cmenuname, cAuthId, cCmdLine
End Function

'扩展解决与外部接口问题
Public Function SetUseModeName(sUseModeName As String) As Boolean
    cModeCode = UCase(sUseModeName)
End Function

Public Function GetVouchForm() As Object
   Set GetVouchForm = GetActiveForm()
End Function

'将系统类传出供外部系统更改部分参数
Public Function GetSASysInfObject() As Object
    Set GetSASysInfObject = clsSAWeb
End Function

'是否实现加载了单据窗体
Public Property Get bSuccessShowVoucher() As Boolean
    bSuccessShowVoucher = g_FormbillShow
End Property




''''''''''''''判断账套演示版是否到期
Public Function bDemoOverTime() As Boolean
    Dim bOverTime As Integer
    Dim strsql As String
    Dim RsTemp As New ADODB.Recordset
    Set RsTemp = New ADODB.Recordset
    bDemoOverTime = False
    RsTemp.CursorLocation = adUseClient
    
    Select Case UCase(m_login.cSub_Id)
 
        Case "1N"   '费用
            If RsTemp.State <> 0 Then RsTemp.Close
            RsTemp.Open "select datediff(month,min([ddate]),max([ddate])) as date1 from EFFYGL_Pcostbudget  having datediff(month,min([ddate]),max([ddate]))>=3 ", DBconn.ConnectionString, adOpenDynamic, adLockOptimistic
            If RsTemp.RecordCount > 0 Then
                bDemoOverTime = True: Exit Function
            End If
            
            If RsTemp.State <> 0 Then RsTemp.Close
            RsTemp.Open "select datediff(month,min([ddate]),max([ddate])) as date1 from EFFYGL_SettleVouch  having datediff(month,min([ddate]),max([ddate]))>=3 ", DBconn.ConnectionString, adOpenDynamic, adLockOptimistic
            If RsTemp.RecordCount > 0 Then
                bDemoOverTime = True: Exit Function
            End If
    End Select
    
    Exit Function
End Function
