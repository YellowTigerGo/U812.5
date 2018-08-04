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

'by ahzzd 保存前插件接口插件接口
' m_login  U8   登陆对象
'Dbconn         当前数据库连接
'Cardnumber     单据关键字
'domHead        表头数据
'domBody        表体数据
'strUserErr     错误信息
'bSuc   =true   表示处理成功           =false表示处理失败
Public Sub Before_Save(m_Login As Object, Dbconn As ADODB.Connection, Cardnumber As String, domHead As DOMDocument, domBody As DOMDocument, Optional strUserErr As String, Optional bSuc As Boolean)
On Error GoTo Errhandle
bSuc = True


If LCase(Cardnumber) = LCase("YXEF9101") Then




'domBody.selectNodes("//z:row").length


    Dim t_cinvcode1 As String
   
     t_cinvcode1 = GetBodyItemValue(domHead, "t_cinvcode", 0)
      Dim str8_1 As String
      str8_1 = GetBodyItemValue(domHead, "str8", 0)
        
        
         Dim id As Long
     
      If GetBodyItemValue(domHead, "id", 0) <> "" Then
       id = GetBodyItemValue(domHead, "id", 0)
     
    
      
      End If
        

'检验存货编码是否合法
        Dim sql2 As String
         sql2 = "select cInvCode from Inventory  where cInvCode='" & t_cinvcode1 & "'"
         Dim rs2 As New ADODB.Recordset
         rs2.CursorLocation = adUseClient
         rs2.Open sql2, Dbconn.ConnectionString, 3, 4
             If rs2.RecordCount > 0 Then
             
                Else
                strUserErr = "该存货编码不存在，请选择正确的存货编码！！！"
                bSuc = False
             End If
         rs2.Close
         Set rs2 = Nothing
        
        
        

            Dim sql3 As String
            sql3 = "select id from EF_Inventory_Information where t_cinvcode='" & t_cinvcode1 & "'"
            Dim rs3 As New ADODB.Recordset
            rs3.CursorLocation = adUseClient
            rs3.Open sql3, Dbconn.ConnectionString, 3, 4
            
            Dim id2 As Long
                If rs3.RecordCount > 0 Then
                id2 = rs3.Fields(0)
                Else
                id2 = id
                End If
            rs3.Close
            Set rs3 = Nothing
        
        
        
         
        If CLng(id) <> 0 Then  '代表该情况下为 修改的情况
            If CLng(id) <> CLng(id2) Then
            
                        '产品档案参照 存货档案生单时 存货编码不可以重复
                        Dim sql1 As String
                        sql1 = "select t_cinvcode from EF_Inventory_Information where t_cinvcode='" & t_cinvcode1 & "'"
                        Dim rs1 As New ADODB.Recordset
                        rs1.CursorLocation = adUseClient
                        rs1.Open sql1, Dbconn.ConnectionString, 3, 4
                        
                            If rs1.RecordCount > 0 Then
                                strUserErr = "该存货编码已被参照，请另选存货编码进行参照！！！"
                                bSuc = False
                            
                            End If
                        rs1.Close
                        Set rs1 = Nothing
            
            
            End If
            
        
        
        Else  '此时是新增 的情况 需要校验的
        
                   '产品档案参照 存货档案生单时 存货编码不可以重复
                    Dim sql1_1 As String
                    sql1_1 = "select t_cinvcode from EF_Inventory_Information where t_cinvcode='" & t_cinvcode1 & "'"
                    Dim rs1_1 As New ADODB.Recordset
                    rs1_1.CursorLocation = adUseClient
                    rs1_1.Open sql1_1, Dbconn.ConnectionString, 3, 4
                    
                        If rs1_1.RecordCount > 0 Then
                            strUserErr = "该存货编码已被参照，请另选存货编码进行参照！！！"
                            bSuc = False
                        End If
                    rs1_1.Close
                    Set rs1_1 = Nothing
        
        
        End If
        '检验客户编码是否存在
        Dim t_ccuscode1 As String
        t_ccuscode1 = GetBodyItemValue(domHead, "t_ccuscode", 0)
        If t_ccuscode1 <> "" Then
        Dim sql4 As String
        sql4 = "select cCusCode from Customer where cCusCode='" & t_ccuscode1 & "'"
        Dim rs4 As New ADODB.Recordset
            rs4.CursorLocation = adUseClient
            
            rs4.Open sql4, Dbconn.ConnectionString, 3, 4
   
               If rs4.RecordCount > 0 Then
                
                   Else
                   strUserErr = "该客户编码不存在，请重新选择客户编码进行参照！！！"
                   bSuc = False
               End If
        rs4.Close
        Set rs4 = Nothing
        
       End If
     
       '检验成品状态是否合法
'        If str8_1 <> "袋料" Or str8_1 <> "卷膜" Or str8_1 <> "片膜" Then
'            strUserErr = "成品状态只能为 袋料，卷膜，片膜！！！"
'            bSuc = False
'        End If
        
Else
    bSuc = True
End If



If LCase(Cardnumber) = LCase("YXEF9111") Then



Dim b_str1_1 As String   '获取班组编码
Dim sql7 As String
Dim rs7 As New ADODB.Recordset
         
    Dim id1_1 As Long
     
    If GetBodyItemValue(domHead, "id", 0) <> "" Then '获取当前单据的ID 号
     id1_1 = GetBodyItemValue(domHead, "id", 0)
 
    End If
      

For i7 = 0 To domBody.selectNodes("//z:row").length - 1

b_str1_1 = GetBodyItemValue(domBody, "b_str1", i7)


If CLng(id1_1) <> 0 Then '代表该情况为修改的情况
   
   sql7 = "select b_str1 from EF_dep_Bzs where b_str1='" & b_str1_1 & "' and id<>'" & id1_1 & "'"
   
   
   Else
   sql7 = "select b_str1 from EF_dep_Bzs where b_str1='" & b_str1_1 & "'"
   
   
End If

'sql7 = "select b_str1 from EF_dep_Bzs where b_str1='" & b_str1_1 & "'"
rs7.Open sql7, Dbconn.ConnectionString, 3, 4
If rs7.RecordCount > 0 Then
    strUserErr = "该班组编码中有已被别的部门使用，不可以重复使用！！！"
bSuc = False


End If
    rs7.Close
    Set rs7 = Nothing


Next

Else
    bSuc = True

End If

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
'bSuc   =true   表示处理成功           =false表示处理失败
Public Sub After_Save(m_Login As Object, Dbconn As ADODB.Connection, Cardnumber As String, id As String, domHead As DOMDocument, domBody As DOMDocument, Optional strUserErr As String, Optional bSuc As Boolean)
On Error GoTo Errhandle
     Dim i As Integer
'     For i = 0 To colInterface.Count - 1
'         '判断用户DLL是否有这个方法
'         strTemp = colInterface.Keys(i)
'         If bHasFunction("PreviewVoucher", strTemp) Then
'            Set clsInterface = colInterface.Item(strTemp)
'            clsInterface.VoucherCheck_AfterSave m_login, Dbconn, Cardnumber, domHead, domBody, strUserErr, bSuc
'            If bSuc = False Then
'                GoTo ExitSub
'            End If
'         End If
'     Next i
     '注意：如果没有任何的用户自定义动态库，则需要返回TRUE
     
    ' MsgBox "保存后插件接口插件接口 After_Save 配置成功"
     bSuc = True
ExitSub:
     Exit Sub
Errhandle:
     bSuc = False
     strUserErr = Err.Description
End Sub

 
 
 
 
 '单据审核操作时候调用（事前）
'##ModelId=3FEA8A97023B
Public Sub Before_Verify(ByVal m_Login As Object, ByVal cnnAccount As ADODB.Connection, ByVal Cardnumber As String, ByVal domHead As Object, strUserErr As String, bSuc As Boolean)
On Error GoTo Errhandle
     Dim i As Integer

 
     'MsgBox "审核前插件接口插件接口 Before_Verify 配置成功"
     '注意：如果没有任何的用户自定义动态库，则需要返回TRUE
     bSuc = True
ExitSub:
     Exit Sub
Errhandle:
     bSuc = False
     strUserErr = Err.Description
End Sub

'单据审核操作成功以后调用（事后）
'##ModelId=3FEA8A97023B
Public Sub After_Verify(ByVal Login As Object, ByVal cnnAccount As ADODB.Connection, ByVal Cardnumber As String, ByVal domHead As Object, strUserErr As String, bSuc As Boolean)
On Error GoTo Errhandle
     Dim i As Integer

    'MsgBox "审核后插件接口插件接口 After_Verify 配置成功"
     '注意：如果没有任何的用户自定义动态库，则需要返回TRUE
     bSuc = True
ExitSub:
     Exit Sub
Errhandle:
     bSuc = False
     strUserErr = Err.Description
End Sub





'单据弃审操作时候调用（事前）
'##ModelId=3FEA8A97026A
Public Sub Before_UnVerify(ByVal Login As Object, ByVal cnnAccount As ADODB.Connection, ByVal Cardnumber As String, ByVal domHead As Object, strUserErr As String, bSuc As Boolean)
On Error GoTo Errhandle

 
     
   'MsgBox "弃审前插件接口插件接口 Before_UnVerify 配置成功"
     '注意：如果没有任何的用户自定义动态库，则需要返回TRUE
     bSuc = True
ExitSub:
     Exit Sub
Errhandle:
     bSuc = False
     strUserErr = Err.Description
End Sub

'单据弃审操作时候调用（事后）
'##ModelId=3FEA8A97026A
Public Sub After_UnVerify(ByVal Login As Object, ByVal cnnAccount As ADODB.Connection, ByVal Cardnumber As String, ByVal domHead As Object, strUserErr As String, bSuc As Boolean)
On Error GoTo Errhandle

 
  'MsgBox "弃审后插件接口插件接口 After_UnVerify 配置成功"
     '注意：如果没有任何的用户自定义动态库，则需要返回TRUE
     bSuc = True
ExitSub:
     Exit Sub
Errhandle:
     bSuc = False
     strUserErr = Err.Description
End Sub






'单据删除操作时候调用
'##ModelId=3FEA8A9701FC
Public Sub Before_Delete(ByVal Login As Object, ByVal cnnAccount As ADODB.Connection, ByVal Cardnumber As String, ByVal domHead As Object, strUserErr As String, bSuc As Boolean)
On Error GoTo Errhandle

 

 ' MsgBox "删除前插件接口插件接口 Before_Delete 配置成功"
     '注意：如果没有任何的用户自定义动态库，则需要返回TRUE
     bSuc = True
ExitSub:
     Exit Sub
Errhandle:
     bSuc = False
     strUserErr = Err.Description
End Sub


'单据删除操作时候调用
'##ModelId=3FEA8A9701FC
Public Sub After_Delete(ByVal Login As Object, ByVal cnnAccount As ADODB.Connection, ByVal Cardnumber As String, ByVal domHead As Object, strUserErr As String, bSuc As Boolean)
On Error GoTo Errhandle

 
 
 
 'MsgBox "删除后插件接口插件接口 After_Delete 配置成功"
     '注意：如果没有任何的用户自定义动态库，则需要返回TRUE
     bSuc = True
ExitSub:
     Exit Sub
Errhandle:
     bSuc = False
     strUserErr = Err.Description
End Sub



