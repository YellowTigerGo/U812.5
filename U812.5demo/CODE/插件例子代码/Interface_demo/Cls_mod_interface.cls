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
Public Sub Before_Save(m_login As Object, DBconn As ADODB.Connection, Cardnumber As String, domHead As DOMDocument, domBody As DOMDocument, Optional strUserErr As String, Optional bSuc As Boolean)
    Dim ele As IXMLDOMElement
    Dim list As IXMLDOMNodeList
    Dim i As Long
    Dim cinvname_cad As String
    Dim cperform As String
    Dim editprop As String
On Error GoTo Errhandle
    
    bSuc = True

    If GetHeadItemValue(domHead, "citem_class") = "" Then
        Call SetHeadItemValue(domHead, "citem_class", "97")
    End If
    
'    MsgBox "保存前插件接口插件接口 Before_Save 配置成功"
     
ExitSub:
     Exit Sub
Errhandle:
     bSuc = False
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
Public Sub After_Save(m_login As Object, DBconn As ADODB.Connection, Cardnumber As String, id As String, domHead As DOMDocument, domBody As DOMDocument, Optional strUserErr As String, Optional bSuc As Boolean)
    Dim ele As IXMLDOMElement
    Dim list As IXMLDOMNodeList
    Dim i As Long
    Dim editprop As String
    Dim strsql As String
    Dim rs As New ADODB.Recordset
    Dim irowno As String
    Dim hzid As String
    Dim hzbilltype As String
On Error GoTo Errhandle
    bSuc = True
    Select Case Cardnumber
        Case "PD010401"
            Set list = domBody.selectNodes("//z:row")
            If list.length = 0 Then
                strUserErr = "表体记录不能为空。"
                GoTo Errhandle
            Else
                i = 1
                For Each ele In list
                    editprop = GetElementValue(ele, "editprop")
                    If editprop = "A" Or editprop = "M" Or editprop = "D" Then
                        hzid = GetElementValue(ele, "hzid")
                        hzbilltype = GetElementValue(ele, "hzbilltype")
                        irowno = GetElementValue(ele, "irowno")
'                        If hzbilltype = "材料清单" Then
'                            strsql = "update hzland_materialsdetail set isendqty=(select sum(iqty) as iqty from EF_Invoices where hzid='" & hzid & "') where id='" & hzid & "'"
'                            DBconn.Execute strsql
'                            '检验累计发货量不能大于设计清单量
'                            strsql = "select isendqty,num from hzland_materialsdetail where id='" & hzid & "' and isendqty>num"
'                            If rs.State <> 0 Then rs.Close
'                            rs.Open strsql, DBconn
'                            If Not rs.BOF And Not rs.EOF Then
'                                strUserErr = "第" & irowno & "行,累计发货量为" & rs!isendqty & "，已超出设计清单量" & rs!Num & "。"
'                                GoTo Errhandle
'                            End If
'                        ElseIf hzbilltype = "冷作清单" Then
'                            strsql = "update hzland_lzdetail set isendqty=(select sum(iqty) as iqty from EF_Invoices where hzid='" & hzid & "') where id='" & hzid & "'"
'                            DBconn.Execute strsql
'                            '检验累计发货量不能大于设计清单量
'                            strsql = "select isendqty,num from hzland_materialsdetail where id='" & hzid & "' and isendqty>num"
'                            If rs.State <> 0 Then rs.Close
'                            rs.Open strsql, DBconn
'                            If Not rs.BOF And Not rs.EOF Then
'                                strUserErr = "第" & irowno & "行,累计发货量为" & rs!isendqty & "，已超出设计清单量" & rs!Num & "。"
'                                GoTo Errhandle
'                            End If
'                        Else
                        If hzbilltype = "配置清单" Then
                            strsql = "update hzland_configurationdetail set isendqty=(select sum(iqty) as iqty from EF_Invoices where hzid='" & hzid & "') where id='" & hzid & "'"
                            DBconn.Execute strsql
                            '检验累计发货量不能大于设计清单量
                            strsql = "select isendqty,num from hzland_configurationdetail where id='" & hzid & "' and isendqty>num"
                            If rs.State <> 0 Then rs.Close
                            rs.Open strsql, DBconn
                            If Not rs.BOF And Not rs.EOF Then
                                strUserErr = "第" & irowno & "行,累计发货量为" & rs!isendqty & "，已超出设计清单量" & rs!Num & "。"
                                GoTo Errhandle
                            End If
                        End If
                        
                    End If
                    i = i + 1
                Next
            End If
        Case "PD010501"
            Set list = domBody.selectNodes("//z:row")
            If list.length = 0 Then
                strUserErr = "表体记录不能为空。"
                GoTo Errhandle
            Else
                i = 1
                For Each ele In list
                    editprop = GetElementValue(ele, "editprop")
                    If editprop = "A" Or editprop = "M" Or editprop = "D" Then
                        hzid = GetElementValue(ele, "hzid")
                        hzbilltype = GetElementValue(ele, "hzbilltype")
                        irowno = GetElementValue(ele, "irowno")
'                        If hzbilltype = "材料清单" Then
'                            strsql = "update hzland_materialsdetail set isendqty=(select sum(iqty) as iqty from EF_PackingLists where hzid='" & hzid & "') where id='" & hzid & "'"
'                            DBconn.Execute strsql
'                            '检验累计发货量不能大于设计清单量
'                            strsql = "select isendqty,num from hzland_materialsdetail where id='" & hzid & "' and isendqty>num"
'                            If rs.State <> 0 Then rs.Close
'                            rs.Open strsql, DBconn
'                            If Not rs.BOF And Not rs.EOF Then
'                                strUserErr = "第" & irowno & "行,累计装箱量为" & rs!isendqty & "，已超出设计清单量" & rs!Num & "。"
'                                GoTo Errhandle
'                            End If
'                        ElseIf hzbilltype = "冷作清单" Then
'                            strsql = "update hzland_lzdetail set isendqty=(select sum(iqty) as iqty from EF_PackingLists where hzid='" & hzid & "') where id='" & hzid & "'"
'                            DBconn.Execute strsql
'                            '检验累计发货量不能大于设计清单量
'                            strsql = "select isendqty,num from hzland_materialsdetail where id='" & hzid & "' and isendqty>num"
'                            If rs.State <> 0 Then rs.Close
'                            rs.Open strsql, DBconn
'                            If Not rs.BOF And Not rs.EOF Then
'                                strUserErr = "第" & irowno & "行,累计装箱量为" & rs!isendqty & "，已超出设计清单量" & rs!Num & "。"
'                                GoTo Errhandle
'                            End If
'                        Else
                        If hzbilltype = "配置清单" Then
                            strsql = "update hzland_configurationdetail set isendqty=(select sum(iqty) as iqty from EF_PackingLists where hzid='" & hzid & "') where id='" & hzid & "'"
                            DBconn.Execute strsql
                            '检验累计发货量不能大于设计清单量
                            strsql = "select isendqty,num from hzland_configurationdetail where id='" & hzid & "' and isendqty>num"
                            If rs.State <> 0 Then rs.Close
                            rs.Open strsql, DBconn
                            If Not rs.BOF And Not rs.EOF Then
                                strUserErr = "第" & irowno & "行,累计装箱量为" & rs!isendqty & "，已超出设计清单量" & rs!Num & "。"
                                GoTo Errhandle
                            End If
                        End If
                        
                    End If
                    i = i + 1
                Next
            End If
    End Select
    
'     MsgBox "保存后插件接口插件接口 After_Save 配置成功"

ExitSub:
     Exit Sub
Errhandle:
     bSuc = False
     strUserErr = strUserErr & Err.Description
End Sub

 
 
 
 
'by ahzzd 审核前插件接口插件接口
' m_login       U8 登陆对象
'Dbconn         当前数据库连接
'Cardnumber     单据关键字
'domHead        表头数据
'strUserErr     错误信息
'bSuc   =true   表示处理成功继续系统流程           =false表示处理失败中断系统流程
Public Sub Before_Verify(m_login As Object, cnnAccount As ADODB.Connection, Cardnumber As String, domHead As Object, strUserErr As String, bSuc As Boolean)

On Error GoTo Errhandle
      
     bSuc = True
     
ExitSub:
    
     Exit Sub
Errhandle:
     bSuc = False
     strUserErr = strUserErr & Err.Description
End Sub

 
'by ahzzd 审核后插件接口插件接口
' m_login       U8 登陆对象
'Dbconn         当前数据库连接
'Cardnumber     单据关键字
'domHead        表头数据
'strUserErr     错误信息
'bSuc   =true   表示处理成功继续系统流程           =false表示处理失败中断系统流程
Public Sub After_Verify(Login As Object, cnnAccount As ADODB.Connection, Cardnumber As String, domHead As Object, strUserErr As String, bSuc As Boolean)

    Dim strsql As String
    Dim rs As ADODB.Recordset
    Dim id As String
    
On Error GoTo Errhandle
    bSuc = True
    
    
'    MsgBox "审核后插件接口插件接口 After_Verify 配置成功"
     '注意：如果没有任何的用户自定义动态库，则需要返回TRUE

ExitSub:
'     If Not rs Is Nothing Then
'        If rs.State <> 0 Then
'            rs.Close
'        End If
'        Set rs = Nothing
'    End If
    Exit Sub
Errhandle:
     bSuc = False
     strUserErr = strUserErr & Err.Description
End Sub

'by ahzzd 弃审前插件接口插件接口
' m_login       U8 登陆对象
'Dbconn         当前数据库连接
'Cardnumber     单据关键字
'domHead        表头数据
'strUserErr     错误信息
'bSuc   =true   表示处理成功继续系统流程           =false表示处理失败中断系统流程
Public Sub Before_UnVerify(Login As Object, cnnAccount As ADODB.Connection, Cardnumber As String, domHead As Object, strUserErr As String, bSuc As Boolean)
    Dim strsql As String
    Dim rs As ADODB.Recordset
    Dim id As String
    Dim ccode As String
On Error GoTo Errhandle
      
     bSuc = True
     
ExitSub:
'    If Not rs Is Nothing Then
'        If rs.State <> 0 Then
'            rs.Close
'        End If
'        Set rs = Nothing
'    End If
     Exit Sub
Errhandle:
     bSuc = False
     strUserErr = strUserErr & Err.Description
End Sub

'by ahzzd 弃审后插件接口插件接口
' m_login       U8 登陆对象
'Dbconn         当前数据库连接
'Cardnumber     单据关键字
'domHead        表头数据
'strUserErr     错误信息
'bSuc   =true   表示处理成功继续系统流程           =false表示处理失败中断系统流程
Public Sub After_UnVerify(Login As Object, cnnAccount As ADODB.Connection, Cardnumber As String, domHead As Object, strUserErr As String, bSuc As Boolean)
On Error GoTo Errhandle
     bSuc = True
 
'  MsgBox "弃审后插件接口插件接口 After_UnVerify 配置成功"
     '注意：如果没有任何的用户自定义动态库，则需要返回TRUE

ExitSub:
     Exit Sub
Errhandle:
     bSuc = False
     strUserErr = strUserErr & Err.Description
End Sub





'by ahzzd 删除前插件接口插件接口
' m_login       U8 登陆对象
'Dbconn         当前数据库连接
'Cardnumber     单据关键字
'domHead        表头数据
'strUserErr     错误信息
'bSuc   =true   表示处理成功继续系统流程           =false表示处理失败中断系统流程
Public Sub Before_Delete(Login As Object, cnnAccount As ADODB.Connection, Cardnumber As String, domHead As Object, strUserErr As String, bSuc As Boolean)
    Dim id As String
    Dim autoid As String
    Dim ccode As String
    Dim strsql As String
    Dim rs As New ADODB.Recordset
    
    Dim hzid As String
    Dim hzbilltype As String
    On Error GoTo Errhandle
     bSuc = True
    Select Case Cardnumber
        Case "PD010401"
            id = GetHeadItemValue(domHead, "id")
            strsql = "select * from EF_Invoices where id='" & id & "'"
            rs.Open strsql, cnnAccount
            While Not rs.EOF
                autoid = rs!autoid
                hzid = rs!hzid & ""
                hzbilltype = rs!hzbilltype & ""
'                If hzbilltype = "材料清单" Then
'                    strsql = "update hzland_materialsdetail set isendqty=iqty from (select sum(iqty) as iqty from EF_Invoices where id<>'" & id & "' and hzid='" & hzid & "') s" & _
'                        " where hzland_materialsdetail.id='" & hzid & "'"
'                    cnnAccount.Execute strsql
'                ElseIf hzbilltype = "冷作清单" Then
'                    strsql = "update hzland_lzdetail set isendqty=iqty from (select sum(iqty) as iqty from EF_Invoices where id<>'" & id & "' and hzid='" & hzid & "') s" & _
'                        " where hzland_lzdetail.id='" & hzid & "'"
'                    cnnAccount.Execute strsql
'                Else
                If hzbilltype = "配置清单" Then
                    strsql = "update hzland_configurationdetail set isendqty=iqty from (select sum(iqty) as iqty from EF_Invoices where id<>'" & id & "' and hzid='" & hzid & "') s" & _
                        " where hzland_configurationdetail.id='" & hzid & "'"
                    cnnAccount.Execute strsql
                End If
                rs.MoveNext
            Wend
        Case "PD010501"
            id = GetHeadItemValue(domHead, "id")
            strsql = "select * from EF_PackingLists where id='" & id & "'"
            rs.Open strsql, cnnAccount
            While Not rs.EOF
                autoid = rs!autoid
                hzid = rs!hzid & ""
                hzbilltype = rs!hzbilltype & ""
'                If hzbilltype = "材料清单" Then
'                    strsql = "update hzland_materialsdetail set isendqty=iqty from (select sum(iqty) as iqty from EF_PackingLists where id<>'" & id & "' and hzid='" & hzid & "') s" & _
'                        " where hzland_materialsdetail.id='" & hzid & "'"
'                    cnnAccount.Execute strsql
'                ElseIf hzbilltype = "冷作清单" Then
'                    strsql = "update hzland_lzdetail set isendqty=iqty from (select sum(iqty) as iqty from EF_PackingLists where id<>'" & id & "' and hzid='" & hzid & "') s" & _
'                        " where hzland_lzdetail.id='" & hzid & "'"
'                    cnnAccount.Execute strsql
'                Else
                If hzbilltype = "配置清单" Then
                    strsql = "update hzland_configurationdetail set isendqty=iqty from (select sum(iqty) as iqty from EF_PackingLists where id<>'" & id & "' and hzid='" & hzid & "') s" & _
                        " where hzland_configurationdetail.id='" & hzid & "'"
                    cnnAccount.Execute strsql
                End If
                rs.MoveNext
            Wend
    End Select
ExitSub:
    If rs.State <> 0 Then
        rs.Close
    End If
    Set rs = Nothing
     Exit Sub
Errhandle:
     bSuc = False
     strUserErr = Err.Description
End Sub


'by ahzzd 删除后插件接口插件接口
' m_login       U8 登陆对象
'Dbconn         当前数据库连接
'Cardnumber     单据关键字
'domHead        表头数据
'strUserErr     错误信息
'bSuc   =true   表示处理成功继续系统流程           =false表示处理失败中断系统流程
Public Sub After_Delete(Login As Object, cnnAccount As ADODB.Connection, Cardnumber As String, domHead As Object, strUserErr As String, bSuc As Boolean)
    Dim id As String
    Dim ccode As String
    Dim strsql As String
On Error GoTo Errhandle
    bSuc = True
' MsgBox "删除后插件接口插件接口 After_Delete 配置成功"
 

ExitSub:
     Exit Sub
Errhandle:
     bSuc = False
     strUserErr = Err.Description
End Sub


