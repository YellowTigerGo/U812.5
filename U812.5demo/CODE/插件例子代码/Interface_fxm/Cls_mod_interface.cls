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
'bSuc   =true   表示处理成功           =false表示处理失败
Public Sub Before_Save(m_login As Object, DBconn As ADODB.Connection, Cardnumber As String, domHead As DOMDocument, domBody As DOMDocument, Optional strUserErr As String, Optional bSuc As Boolean)

On Error GoTo Errhandle
    bSuc = True
    Select Case LCase(Cardnumber)
       Case LCase("YXEF9101")
            If Not CheckData(DBconn, Cardnumber, domHead, domBody, strUserErr) Then
                bSuc = False
                strUserErr = strUserErr
            Else
                strUserErr = ""
                bSuc = True
            End If
       Case LCase("YXEF9110")
            bSuc = CheckData_YXEF9110(DBconn, Cardnumber, domHead, domBody, strUserErr)
            If Not bSuc Then
                strUserErr = strUserErr
                bSuc = False
            End If
       Case LCase("YXEF9105")
            bSuc = CheckData_YXEF9110(DBconn, Cardnumber, domHead, domBody, strUserErr)
            If Not bSuc Then
                strUserErr = strUserErr
                bSuc = False
            End If
       Case LCase("YXEF9104")
            bSuc = CheckData_YXEF9110(DBconn, Cardnumber, domHead, domBody, strUserErr)
            If Not bSuc Then
                strUserErr = strUserErr
                bSuc = False
            End If
       Case LCase("YXEF9114")
            bSuc = CheckData_YXEF9110(DBconn, Cardnumber, domHead, domBody, strUserErr)
            If Not bSuc Then
                strUserErr = strUserErr
                bSuc = False
            End If
       Case LCase("YXEF9122"), LCase("YXEF9123"), LCase("'YXEF9124"), LCase("'YXEF9125"), LCase("'YXEF9126")
            bSuc = CheckData_YXEF9110(DBconn, Cardnumber, domHead, domBody, strUserErr)
            If Not bSuc Then
                strUserErr = strUserErr
                bSuc = False
            End If
       Case Else
      
    End Select
    
    Exit Sub
Errhandle:
    bSuc = True
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
Public Sub After_Save(m_login As Object, DBconn As ADODB.Connection, Cardnumber As String, id As String, domHead As DOMDocument, domBody As DOMDocument, Optional strUserErr As String, Optional bSuc As Boolean)
On Error GoTo Errhandle
    Dim i As Integer
    Dim strSql As String
    Dim DBRS As New ADODB.Recordset
    
    Dim invcode As String
    
    invcode = GetHeadItemValue(domHead, "t_cinvcode")
    Select Case LCase(Cardnumber)
    
        Case LCase("YXEF9122"), LCase("YXEF9123"), LCase("YXEF9124"), LCase("YXEF9125"), LCase("YXEF9126")
            If invcode = "" Then
                bSuc = True
                Exit Sub
            End If
            If invcode = "" Then
                bSuc = True
                Exit Sub
            End If
            strSql = "select * from EFFBgAndInv where cInvCode = '" & invcode & "'"
            If DBRS.State = 1 Then DBRS.Close
            DBRS.Open strSql, DBconn.ConnectionString, adOpenKeyset, adLockPessimistic
            
            If Not DBRS.EOF And Not DBRS.BOF Then
                strSql = "delete effbgandinv where cInvCode='" & invcode & "'"
                DBconn.Execute strSql
            End If
    
            Do While Not DBRS.EOF

                strSql = "insert into effbgandinv ([GUID],[CbgCode],[cInvCode]) values(newid(),'" & DBRS.Fields("CbgCode") & "','" & invcode & "')"
                DBconn.Execute strSql
                
                DBRS.MoveNext
            Loop
        Case LCase("YXEF9122"), LCase("YXEF9123"), LCase("'YXEF9124"), LCase("'YXEF9125"), LCase("'YXEF9126")
            If LCase(editprop) = LCase("M") Then
                editprop = ""
            End If
        Case Else
            
            
    End Select

    
    bSuc = True
ExitSub:
    Exit Sub
Errhandle:
     bSuc = False
     strUserErr = strUserErr & Err.Description
End Sub

 
 
 
 
 '单据审核操作时候调用（事前）
'##ModelId=3FEA8A97023B
Public Sub Before_Verify(ByVal m_login As Object, ByVal cnnAccount As ADODB.Connection, ByVal Cardnumber As String, ByVal domHead As Object, strUserErr As String, bSuc As Boolean)
On Error GoTo Errhandle
     Dim i As Integer

 
'     MsgBox "审核前插件接口插件接口 Before_Verify 配置成功"
     '注意：如果没有任何的用户自定义动态库，则需要返回TRUE
     bSuc = True
ExitSub:
     Exit Sub
Errhandle:
     bSuc = False
     strUserErr = strUserErr & Err.Description
End Sub

'单据审核操作成功以后调用（事后）
'##ModelId=3FEA8A97023B
Public Sub After_Verify(ByVal Login As Object, ByVal cnnAccount As ADODB.Connection, ByVal Cardnumber As String, ByVal domHead As Object, strUserErr As String, bSuc As Boolean)
On Error GoTo Errhandle
     Dim i As Integer

'    MsgBox "审核后插件接口插件接口 After_Verify 配置成功"
     '注意：如果没有任何的用户自定义动态库，则需要返回TRUE
     bSuc = True
ExitSub:
     Exit Sub
Errhandle:
     bSuc = False
     strUserErr = strUserErr & Err.Description
End Sub





'单据弃审操作时候调用（事前）
'##ModelId=3FEA8A97026A
Public Sub Before_UnVerify(ByVal Login As Object, ByVal cnnAccount As ADODB.Connection, ByVal Cardnumber As String, ByVal domHead As Object, strUserErr As String, bSuc As Boolean)
On Error GoTo Errhandle

 
     
   MsgBox "弃审前插件接口插件接口 Before_UnVerify 配置成功"
     '注意：如果没有任何的用户自定义动态库，则需要返回TRUE
     bSuc = True
ExitSub:
     Exit Sub
Errhandle:
     bSuc = False
     strUserErr = strUserErr & Err.Description
End Sub

'单据弃审操作时候调用（事后）
'##ModelId=3FEA8A97026A
Public Sub After_UnVerify(ByVal Login As Object, ByVal cnnAccount As ADODB.Connection, ByVal Cardnumber As String, ByVal domHead As Object, strUserErr As String, bSuc As Boolean)
On Error GoTo Errhandle

 
  MsgBox "弃审后插件接口插件接口 After_UnVerify 配置成功"
     '注意：如果没有任何的用户自定义动态库，则需要返回TRUE
     bSuc = True
ExitSub:
     Exit Sub
Errhandle:
     bSuc = False
     strUserErr = strUserErr & Err.Description
End Sub






'单据删除操作时候调用
'##ModelId=3FEA8A9701FC
Public Sub Before_Delete(ByVal Login As Object, ByVal cnnAccount As ADODB.Connection, ByVal Cardnumber As String, ByVal domHead As Object, strUserErr As String, bSuc As Boolean)
On Error GoTo Errhandle

 

  MsgBox "删除前插件接口插件接口 Before_Delete 配置成功"
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

 
 
 
 MsgBox "删除后插件接口插件接口 After_Delete 配置成功"
     '注意：如果没有任何的用户自定义动态库，则需要返回TRUE
     bSuc = True
ExitSub:
     Exit Sub
Errhandle:
     bSuc = False
     strUserErr = Err.Description
End Sub



