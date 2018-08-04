VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Cls_MOD_interface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'接口类说明
'本插件主要是实现在mod 业务处理层业务处理插件
Option Explicit



'by ahzzd 保存前插件接口插件接口
' m_login  U8   登陆对象
'Dbconn         当前数据库连接
'Cardnumber     单据关键字
'domHead        表头数据
'domBody        表体数据
'strUserErr     错误信息
'bSuc   =true   表示处理成功继续系统流程           =false表示处理失败中断系统流程
Public Sub Before_Save(m_login As Object, DBconn As ADODB.Connection, Cardnumber As String, domHead As DOMDocument, domBody As DOMDocument, Optional strUserErr As String, Optional bsuc As Boolean)
On Error GoTo Errhandle
'    Dim i As Integer
'    Dim strSaleListID As String
'    Dim strSql As String
'    Dim strFlag As String
'    Dim Rs As New ADODB.Recordset
'    Dim StrAutoid As String
    Dim strErr As String

    bsuc = True

    Select Case LCase(Cardnumber)
        Case LCase("YXEF9102")  '验证发货计划单是否可以保存
            strErr = SendPlan_SaveCheck(DBconn, domHead, domBody, strUserErr, bsuc)
            If strErr <> "" Then
                strUserErr = strErr
                bsuc = False
                Exit Sub
            End If
    End Select



'    MsgBox "保存前插件接口插件接口 Before_Save 配置成功"
     
ExitSub:
     Exit Sub
Errhandle:
     bsuc = False
     strUserErr = strUserErr & Err.Description
End Sub

'by ahzzd 保存后插件接口插件接口
' m_login       U8 登陆对象
'Dbconn         当前数据库连接
'Cardnumber     单据关键字
'ID             当前保存的单据ID
'domHead        表头数据
'domBody        表体数据
'strUserErr     错误信息
'bSuc   =true   表示处理成功继续系统流程           =false表示处理失败中断系统流程
Public Sub After_Save(m_login As Object, DBconn As ADODB.Connection, Cardnumber As String, Id As String, domHead As DOMDocument, domBody As DOMDocument, Optional strUserErr As String, Optional bsuc As Boolean)
On Error GoTo Errhandle
    bsuc = True
     
'     MsgBox "保存后插件接口插件接口 After_Save 配置成功"

ExitSub:
     Exit Sub
Errhandle:
     bsuc = False
     strUserErr = strUserErr & Err.Description
End Sub

 
 
 
 
'by ahzzd 审核前插件接口插件接口
' m_login       U8 登陆对象
'Dbconn         当前数据库连接
'Cardnumber     单据关键字
'domHead        表头数据
'strUserErr     错误信息
'bSuc   =true   表示处理成功继续系统流程           =false表示处理失败中断系统流程
Public Sub Before_Verify(ByVal m_login As Object, ByVal cnnAccount As ADODB.Connection, ByVal Cardnumber As String, ByVal domHead As Object, strUserErr As String, bsuc As Boolean)
On Error GoTo Errhandle
      
     bsuc = True
 
'     MsgBox "审核前插件接口插件接口 Before_Verify 配置成功"
     '注意：如果没有任何的用户自定义动态库，则需要返回TRUE

ExitSub:
     Exit Sub
Errhandle:
     bsuc = False
     strUserErr = strUserErr & Err.Description
End Sub

 
'by ahzzd 审核后插件接口插件接口
' m_login       U8 登陆对象
'Dbconn         当前数据库连接
'Cardnumber     单据关键字
'domHead        表头数据
'strUserErr     错误信息
'bSuc   =true   表示处理成功继续系统流程           =false表示处理失败中断系统流程
Public Sub After_Verify(ByVal Login As Object, ByVal cnnAccount As ADODB.Connection, ByVal Cardnumber As String, ByVal domHead As Object, strUserErr As String, bsuc As Boolean)
On Error GoTo Errhandle
    bsuc = True
    
    Select Case LCase(Cardnumber)
        Case LCase("YXEF9103") '派工计划单审核后生成派工单
'            If Checkdata("", cnnAccount, strUserErr) Then
'                Call CreateTaskBill(cnnAccount, domHead, strUserErr, bSuc)
'            Else
'                bSuc = False
'            End If
            
        Case LCase("YXEF9104"), LCase("YXEF9115"), LCase("YXEF9116"), LCase("YXEF9117"), LCase("YXEF9118"), LCase("YXEF9119"), LCase("YXEF9120") '正品报工单审核后将相关信息写入系统报工单
            If Checkdata(Cardnumber, cnnAccount, domHead, strUserErr, bsuc) Then
                Call CreateSystemReportWork(Login, cnnAccount, "z", Cardnumber, domHead, strUserErr, bsuc)
            Else
                bsuc = False
            End If
         Case LCase("YXEF9114"), LCase("YXEF9131"), LCase("YXEF9132"), LCase("YXEF9133"), LCase("YXEF9134"), LCase("YXEF9135"), LCase("YXEF9136") '废品报工单审核后将相关信息写入系统报工单
            
            If Checkdata(Cardnumber, cnnAccount, domHead, strUserErr, bsuc) Then
                Call CreateSystemReportWork(Login, cnnAccount, "f", Cardnumber, domHead, strUserErr, bsuc)
            Else
                bsuc = False
            End If
        Case Else
        
    End Select
'    MsgBox "审核后插件接口插件接口 After_Verify 配置成功"
     '注意：如果没有任何的用户自定义动态库，则需要返回TRUE
ExitSub:
     Exit Sub
Errhandle:
     bsuc = False
     strUserErr = strUserErr & Err.Description
End Sub




 
'by ahzzd 弃审前插件接口插件接口
' m_login       U8 登陆对象
'Dbconn         当前数据库连接
'Cardnumber     单据关键字
'domHead        表头数据
'strUserErr     错误信息
'bSuc   =true   表示处理成功继续系统流程           =false表示处理失败中断系统流程
Public Sub Before_UnVerify(ByVal Login As Object, ByVal cnnAccount As ADODB.Connection, ByVal Cardnumber As String, ByVal domHead As Object, strUserErr As String, bsuc As Boolean)
On Error GoTo Errhandle
    Dim strAutoID As String ' 子表的autoid
    Dim strID As String  '主表ID
    Dim strSql As String 'SQL语句
    Dim Rs As New ADODB.Recordset '数据集
    
    bsuc = True
    '判断发货计划单有没有生成发货单，如果有，则发货计划单不能弃审
    Select Case LCase(Cardnumber)
        Case LCase("YXEF9102")
            strID = GetHeadItemValue(domHead, "id")
            strSql = " select cdefine34 from dispatchlists where cdefine34 in ( select autoid from EF_plan_DispatchLists where id = " & strID & ") "
            
            If Rs.State <> 0 Then Rs.Close
            Rs.CursorLocation = adUseClient
            Rs.Open strSql, cnnAccount.ConnectionString, 3, 4
            
            If Rs.EOF Then
            Else
                strUserErr = "该计划单已生成发货单，不能弃审！"
                bsuc = False
                Exit Sub
            End If
        Case LCase("YXEF9115"), LCase("YXEF9116"), LCase("YXEF9117"), LCase("YXEF9118"), LCase("YXEF9119"), LCase("YXEF9120") '正品报工单
            strID = "z" + GetHeadItemValue(domHead, "id")
            strSql = " select define2 from fc_moroutingbill where define2 = '" & strID & "'"
               
            If Rs.State <> 0 Then Rs.Close
            Rs.CursorLocation = adUseClient
            Rs.Open strSql, cnnAccount.ConnectionString, 3, 4
            
            If Rs.EOF Then
            Else
                strUserErr = "该正品报工单已生成工序报工单，不能弃审！"
                bsuc = False
                Exit Sub
            End If
        Case LCase("YXEF9131"), LCase("YXEF9132"), LCase("YXEF9133"), LCase("YXEF9134"), LCase("YXEF9135"), LCase("YXEF9136") '废品报工单
            strID = "f" + GetHeadItemValue(domHead, "id")
            strSql = " select define2 from fc_moroutingbill where define2 = '" & strID & "'"
                
            If Rs.State <> 0 Then Rs.Close
            Rs.CursorLocation = adUseClient
            Rs.Open strSql, cnnAccount.ConnectionString, 3, 4
            
            If Rs.EOF Then
            Else
                strUserErr = "该废品报工单已生成工序报工单，不能弃审！"
                bsuc = False
                Exit Sub
            End If
        Case Else
        
    End Select
'   MsgBox "弃审前插件接口插件接口 Before_UnVerify 配置成功"
     '注意：如果没有任何的用户自定义动态库，则需要返回TRUE

ExitSub:
     Exit Sub
Errhandle:
     bsuc = False
     strUserErr = strUserErr & Err.Description
End Sub

'by ahzzd 弃审后插件接口插件接口
' m_login       U8 登陆对象
'Dbconn         当前数据库连接
'Cardnumber     单据关键字
'domHead        表头数据
'strUserErr     错误信息
'bSuc   =true   表示处理成功继续系统流程           =false表示处理失败中断系统流程
Public Sub After_UnVerify(ByVal Login As Object, ByVal cnnAccount As ADODB.Connection, ByVal Cardnumber As String, ByVal domHead As Object, strUserErr As String, bsuc As Boolean)
On Error GoTo Errhandle
     bsuc = True
 
'  MsgBox "弃审后插件接口插件接口 After_UnVerify 配置成功"
     '注意：如果没有任何的用户自定义动态库，则需要返回TRUE

ExitSub:
     Exit Sub
Errhandle:
     bsuc = False
     strUserErr = strUserErr & Err.Description
End Sub





'by ahzzd 删除前插件接口插件接口
' m_login       U8 登陆对象
'Dbconn         当前数据库连接
'Cardnumber     单据关键字
'domHead        表头数据
'strUserErr     错误信息
'bSuc   =true   表示处理成功继续系统流程           =false表示处理失败中断系统流程
Public Sub Before_Delete(ByVal Login As Object, ByVal cnnAccount As ADODB.Connection, ByVal Cardnumber As String, ByVal domHead As Object, strUserErr As String, bsuc As Boolean)
On Error GoTo Errhandle
     bsuc = True
 

'  MsgBox "删除前插件接口插件接口 Before_Delete 配置成功"

ExitSub:
     Exit Sub
Errhandle:
     bsuc = False
     strUserErr = Err.Description
End Sub


'by ahzzd 删除后插件接口插件接口
' m_login       U8 登陆对象
'Dbconn         当前数据库连接
'Cardnumber     单据关键字
'domHead        表头数据
'strUserErr     错误信息
'bSuc   =true   表示处理成功继续系统流程           =false表示处理失败中断系统流程
Public Sub After_Delete(ByVal Login As Object, ByVal cnnAccount As ADODB.Connection, ByVal Cardnumber As String, ByVal domHead As Object, strUserErr As String, bsuc As Boolean)
On Error GoTo Errhandle
     bsuc = True
 
 
 
' MsgBox "删除后插件接口插件接口 After_Delete 配置成功"
 

ExitSub:
     Exit Sub
Errhandle:
     bsuc = False
     strUserErr = Err.Description
End Sub

Private Sub Class_Initialize()
    Set cls_Public = CreateObject("UF_Public_base.cls_log")
End Sub
