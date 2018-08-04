VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UserDefineDll_UI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

 
Dim colInterface As Dictionary
Dim clsInterface As Object
Dim strUserErr As String
Dim domdll As DOMDocument
Dim eleMent As IXMLDOMElement
'by lg060722 修改
'Dim i As Integer　i不能是全局
Dim strTemp As String
Dim mbHasUserDLL As Boolean

'by lg-2004-12-13
'调用条码序号管理主窗体
'参数说明:domHead 表头DOM
'          domBody 表体DOM'
'          icurrRow 当前行
'        cnnAccount:数据库连接
'        Login:用友登陆LOGIN
'        struserErr:错误字符串
'        bSuc:是否成功

'接口加B（before）,表示先调用该方法，然后走我们的方法
'接口加A（after），表示先走我们的方法，然后调用该方法

'添加接口：双击时调用该方法，由二次开发输入条形码
Public Sub ctlVoucher1_DblClick(ByVal domHead As Object, ByVal domBody As Object, icurrRow As Long, ByVal cnnAccount As ADODB.Connection, ByVal Login As Object, strUserErr As String, bSuc As Boolean)

On Error GoTo ErrHandle
     Dim i As Integer
     For i = 0 To colInterface.Count - 1
         '判断用户DLL是否有这个方法
         strTemp = colInterface.Keys(i)
         If bHasFunction("ctlVoucher1_DblClick", strTemp) Then
            Set clsInterface = colInterface.Item(strTemp)
            clsInterface.ctlVoucher1_DblClick domHead, domBody, icurrRow, cnnAccount, Login, strUserErr, bSuc
            If bSuc = False Then
                GoTo ExitSub
            End If
         End If
     Next i

     '注意：如果没有任何的用户自定义动态库，则需要返回TRUE
     bSuc = True
ExitSub:
     Exit Sub
ErrHandle:
     bSuc = False
     strUserErr = Err.Description
End Sub

'单据删除操作时候调用
'##ModelId=3FEA8A9701FC
Public Sub Before_Delete(ByVal m_login As Object, ByVal cnnAccount As ADODB.Connection, ByVal Cardnumber As String, ByVal domHead As Object, strUserErr As String, bSuc As Boolean)
On Error GoTo ErrHandle

     Dim i As Integer
     
     For i = 0 To colInterface.Count - 1
         '判断用户DLL是否有这个方法
         strTemp = colInterface.Keys(i)
         If bHasFunction("Before_Delete", strTemp) Then
            Set clsInterface = colInterface.Item(strTemp)
            clsInterface.Before_Delete m_login, cnnAccount, Cardnumber, domHead, strUserErr, bSuc
            If bSuc = False Then
                GoTo ExitSub
            End If
         End If
     Next i


     '注意：如果没有任何的用户自定义动态库，则需要返回TRUE
     bSuc = True
ExitSub:
     Exit Sub
ErrHandle:
     bSuc = False
     strUserErr = Err.Description
End Sub


'单据删除操作时候调用
'##ModelId=3FEA8A9701FC
Public Sub After_Delete(ByVal m_login As Object, ByVal cnnAccount As ADODB.Connection, ByVal Cardnumber As String, ByVal domHead As Object, strUserErr As String, bSuc As Boolean)
On Error GoTo ErrHandle
     Dim i As Integer
     For i = 0 To colInterface.Count - 1
         '判断用户DLL是否有这个方法
         strTemp = colInterface.Keys(i)
         If bHasFunction("After_Delete", strTemp) Then
            Set clsInterface = colInterface.Item(strTemp)
            clsInterface.After_Delete m_login, cnnAccount, Cardnumber, domHead, strUserErr, bSuc
            If bSuc = False Then
                GoTo ExitSub
            End If
         End If
     Next i
     '注意：如果没有任何的用户自定义动态库，则需要返回TRUE
     bSuc = True
ExitSub:
     Exit Sub
ErrHandle:
     bSuc = False
     strUserErr = Err.Description
End Sub



'单据行删除操作时候调用
'##ModelId=3FEA8A9701FC
'Public Sub DeleteLine(ByVal domHead As Object, ByVal domBody As Object, icurrRow As Long, ByVal cnnAccount As ADODB.Connection, ByVal Login As Object, strUserErr As String, bSuc As Boolean)
'On Error GoTo Errhandle
'     Dim i As Integer
'
'
'     For i = 0 To colInterface.Count - 1
'         '判断用户DLL是否有这个方法
'         strTemp = colInterface.Keys(i)
'         If bHasFunction("DeleteLine", strTemp) Then
'            Set clsInterface = colInterface.Item(strTemp)
'            clsInterface.DeleteLine domHead, domBody, icurrRow, cnnAccount, Login, strUserErr, bSuc
'            If bSuc = False Then
'                GoTo ExitSub
'            End If
'         End If
'     Next i
'
'
'
'     '注意：如果没有任何的用户自定义动态库，则需要返回TRUE
'     bSuc = True
'ExitSub:
'     Exit Sub
'Errhandle:
'     bSuc = False
'     strUserErr = Err.Description
'End Sub

'单据审核操作时候调用（事前）
'##ModelId=3FEA8A97023B
Public Sub Before_Verify(ByVal m_login As Object, ByVal cnnAccount As ADODB.Connection, ByVal Cardnumber As String, ByVal domHead As Object, strUserErr As String, bSuc As Boolean)
On Error GoTo ErrHandle
     Dim i As Integer

     For i = 0 To colInterface.Count - 1
         '判断用户DLL是否有这个方法
         strTemp = colInterface.Keys(i)
         If bHasFunction("Before_Verify", strTemp) Then
            Set clsInterface = colInterface.Item(strTemp)
            clsInterface.Before_Verify m_login, cnnAccount, Cardnumber, domHead, strUserErr, bSuc
            If bSuc = False Then
                GoTo ExitSub
            End If
         End If
     Next i
     '注意：如果没有任何的用户自定义动态库，则需要返回TRUE
     bSuc = True
ExitSub:
     Exit Sub
ErrHandle:
     bSuc = False
     strUserErr = Err.Description
End Sub

'单据审核操作成功以后调用（事后）
'##ModelId=3FEA8A97023B
Public Sub After_Verify(ByVal Login As Object, ByVal cnnAccount As ADODB.Connection, ByVal Cardnumber As String, ByVal domHead As Object, strUserErr As String, bSuc As Boolean)
On Error GoTo ErrHandle
     Dim i As Integer

     For i = 0 To colInterface.Count - 1
'        MsgBox colInterface.Count
         '判断用户DLL是否有这个方法
         strTemp = colInterface.Keys(i)
'         MsgBox strTemp
         If bHasFunction("After_Verify", strTemp) Then
            Set clsInterface = colInterface.Item(strTemp)
            clsInterface.After_Verify Login, cnnAccount, Cardnumber, domHead, strUserErr, bSuc
            If bSuc = False Then
                
                GoTo ExitSub
            End If
         End If
     Next i
    
  
     '注意：如果没有任何的用户自定义动态库，则需要返回TRUE
     bSuc = True
ExitSub:
     Exit Sub
ErrHandle:
     bSuc = False
     strUserErr = Err.Description
End Sub

'单据审核不成功时调用（回滚）
'Public Sub VerifyRollBack(ByVal domHead As Object, ByVal domBody As Object, ByVal cnnAccount As ADODB.Connection, ByVal Login As Object, strUserErr As String, bSuc As Boolean)
'On Error GoTo Errhandle
'
'     Dim i As Integer
'
'     For i = 0 To colInterface.Count - 1
'         '判断用户DLL是否有这个方法
'         strTemp = colInterface.Keys(i)
'         If bHasFunction("VerifyRollBack", strTemp) Then
'            Set clsInterface = colInterface.Item(strTemp)
'            clsInterface.VerifyRollBack domHead, domBody, cnnAccount, Login, strUserErr, bSuc
'            If bSuc = False Then
'                GoTo ExitSub
'            End If
'         End If
'     Next i
'
'
'     '注意：如果没有任何的用户自定义动态库，则需要返回TRUE
'     bSuc = True
'ExitSub:
'     Exit Sub
'Errhandle:
'     bSuc = False
'     strUserErr = Err.Description
'End Sub


'单据弃审操作时候调用（事前）
'##ModelId=3FEA8A97026A
Public Sub Before_UnVerify(ByVal Login As Object, ByVal cnnAccount As ADODB.Connection, ByVal Cardnumber As String, ByVal domHead As Object, strUserErr As String, bSuc As Boolean)
On Error GoTo ErrHandle

     Dim i As Integer


     For i = 0 To colInterface.Count - 1
         '判断用户DLL是否有这个方法
         strTemp = colInterface.Keys(i)
         If bHasFunction("Before_UnVerify", strTemp) Then
            Set clsInterface = colInterface.Item(strTemp)
            clsInterface.Before_UnVerify Login, cnnAccount, Cardnumber, domHead, strUserErr, bSuc
            If bSuc = False Then
                GoTo ExitSub
            End If
         End If
     Next i
     
  
     '注意：如果没有任何的用户自定义动态库，则需要返回TRUE
     bSuc = True
ExitSub:
     Exit Sub
ErrHandle:
     bSuc = False
     strUserErr = Err.Description
End Sub

'单据弃审操作时候调用（事后）
'##ModelId=3FEA8A97026A
Public Sub After_UnVerify(ByVal Login As Object, ByVal cnnAccount As ADODB.Connection, ByVal Cardnumber As String, ByVal domHead As Object, strUserErr As String, bSuc As Boolean)
On Error GoTo ErrHandle

     Dim i As Integer
     For i = 0 To colInterface.Count - 1
         '判断用户DLL是否有这个方法
         strTemp = colInterface.Keys(i)
         If bHasFunction("After_UnVerify", strTemp) Then
            Set clsInterface = colInterface.Item(strTemp)
            clsInterface.After_UnVerify Login, cnnAccount, Cardnumber, domHead, strUserErr, bSuc
            If bSuc = False Then
                GoTo ExitSub
            End If
         End If
     Next i
     '注意：如果没有任何的用户自定义动态库，则需要返回TRUE
     bSuc = True
ExitSub:
     Exit Sub
ErrHandle:
     bSuc = False
     strUserErr = Err.Description
End Sub

'单据弃审不成功时调用（回滚）
'Public Sub UnVerifyRollBack(ByVal domHead As Object, ByVal domBody As Object, ByVal cnnAccount As ADODB.Connection, ByVal Login As Object, strUserErr As String, bSuc As Boolean)
'On Error GoTo Errhandle
'     Dim i As Integer
'
'     For i = 0 To colInterface.Count - 1
'         '判断用户DLL是否有这个方法
'         strTemp = colInterface.Keys(i)
'         If bHasFunction("UnVerifyRollBack", strTemp) Then
'            Set clsInterface = colInterface.Item(strTemp)
'            clsInterface.UnVerifyRollBack domHead, domBody, cnnAccount, Login, strUserErr, bSuc
'            If bSuc = False Then
'                GoTo ExitSub
'            End If
'         End If
'     Next i
'
'     '注意：如果没有任何的用户自定义动态库，则需要返回TRUE
'     bSuc = True
'ExitSub:
'     Exit Sub
'Errhandle:
'     bSuc = False
'     strUserErr = Err.Description
'End Sub

'单据行修改行操作时候调用
'##ModelId=3FEA8A9701FC
'Public Sub CanModifyBody(ByVal domHead As Object, ByVal domBody As Object, icurrRow As Long, ByVal cnnAccount As ADODB.Connection, ByVal Login As Object, strUserErr As String, bSuc As Boolean)
'On Error GoTo Errhandle
'
'     Dim i As Integer
'
'
'     For i = 0 To colInterface.Count - 1
'         '判断用户DLL是否有这个方法
'         strTemp = colInterface.Keys(i)
'         If bHasFunction("CanModifyBody", strTemp) Then
'            Set clsInterface = colInterface.Item(strTemp)
'            clsInterface.CanModifyBody domHead, domBody, icurrRow, cnnAccount, Login, strUserErr, bSuc
'            If bSuc = False Then
'                GoTo ExitSub
'            End If
'         End If
'     Next i
'
'     '注意：如果没有任何的用户自定义动态库，则需要返回TRUE
'     bSuc = True
'ExitSub:
'     Exit Sub
'Errhandle:
'     bSuc = False
'     strUserErr = Err.Description
'End Sub

'单据行修改货位操作时候调用
'##ModelId=3FEA8A9701FC
'Public Sub CanModifyPosition(ByVal domHead As Object, ByVal domBody As Object, icurrRow As Long, ByVal cnnAccount As ADODB.Connection, ByVal Login As Object, strUserErr As String, bSuc As Boolean)
'On Error GoTo Errhandle
'
'     Dim i As Integer
'
'     For i = 0 To colInterface.Count - 1
'         '判断用户DLL是否有这个方法
'         strTemp = colInterface.Keys(i)
'         If bHasFunction("CanModifyPosition", strTemp) Then
'            Set clsInterface = colInterface.Item(strTemp)
'            clsInterface.CanModifyPosition domHead, domBody, icurrRow, cnnAccount, Login, strUserErr, bSuc
'            If bSuc = False Then
'                GoTo ExitSub
'            End If
'         End If
'     Next i
'
'
'     '注意：如果没有任何的用户自定义动态库，则需要返回TRUE
'     bSuc = True
'ExitSub:
'     Exit Sub
'Errhandle:
'     bSuc = False
'     strUserErr = Err.Description
'End Sub
'单据打印
'Public Sub PrintBill(ByVal domHead As Object, ByVal domBody As Object, icurrRow As Long, ByVal cnnAccount As ADODB.Connection, ByVal Login As Object, strUserErr As String, bSuc As Boolean)
'On Error GoTo Errhandle
'
'     Dim i As Integer
'     bSuc = False
'
'     For i = 0 To colInterface.Count - 1
'         '判断用户DLL是否有这个方法
'         strTemp = colInterface.Keys(i)
'         If bHasFunction("PrintBill", strTemp) Then
'            Set clsInterface = colInterface.Item(strTemp)
'            clsInterface.PrintBill domHead, domBody, icurrRow, cnnAccount, Login, strUserErr, bSuc
'            If bSuc = False Then
'                GoTo ExitSub
'            End If
'         End If
'     Next i
'
'
'     '注意：如果没有任何的用户自定义动态库，则需要返回TRUE
'ExitSub:
'     Exit Sub
'Errhandle:
'     bSuc = False
'     strUserErr = Err.Description
'End Sub
'单据审核不成功时调用（回滚）
'Public Sub DeleteRollBack(ByVal domHead As Object, ByVal domBody As Object, ByVal cnnAccount As ADODB.Connection, ByVal Login As Object, strUserErr As String, bSuc As Boolean)
'On Error GoTo Errhandle
'     Dim i As Integer
'
'     For i = 0 To colInterface.Count - 1
'         '判断用户DLL是否有这个方法
'         strTemp = colInterface.Keys(i)
'         If bHasFunction("DeleteRollBack", strTemp) Then
'            Set clsInterface = colInterface.Item(strTemp)
'            clsInterface.DeleteRollBack domHead, domBody, cnnAccount, Login, strUserErr, bSuc
'            If bSuc = False Then
'                GoTo ExitSub
'            End If
'         End If
'     Next i
'
'     '注意：如果没有任何的用户自定义动态库，则需要返回TRUE
'     bSuc = True
'ExitSub:
'     Exit Sub
'Errhandle:
'     bSuc = False
'     strUserErr = Err.Description
'End Sub

'单据审核不成功时调用（回滚）
'Public Sub Save(ByVal domHead As Object, ByVal domBody As Object, ByVal cnnAccount As ADODB.Connection, ByVal Login As Object, strUserErr As String, bSuc As Boolean)
'On Error GoTo Errhandle
'     Dim i As Integer
'
'     For i = 0 To colInterface.Count - 1
'         '判断用户DLL是否有这个方法
'         strTemp = colInterface.Keys(i)
'         If bHasFunction("Save", strTemp) Then
'            Set clsInterface = colInterface.Item(strTemp)
'            clsInterface.Save domHead, domBody, cnnAccount, Login, strUserErr, bSuc
'            If bSuc = False Then
'                GoTo ExitSub
'            End If
'         End If
'     Next i
'
'     '注意：如果没有任何的用户自定义动态库，则需要返回TRUE
'     bSuc = True
'ExitSub:
'     Exit Sub
'Errhandle:
'     bSuc = False
'     strUserErr = Err.Description
'End Sub

'add by tianxl 2005-12-12 武汉公交插件
'单据审核不成功时调用（回滚）
'Public Sub BuildVoucher(ByVal domHead As Object, ByVal domBody As Object, ByVal cnnAccount As ADODB.Connection, ByVal Login As Object, strUserErr As String, bSuc As Boolean)
'On Error GoTo Errhandle
'     Dim i As Integer
'
'     For i = 0 To colInterface.Count - 1
'         '判断用户DLL是否有这个方法
'         strTemp = colInterface.Keys(i)
'         If bHasFunction("BuildVoucher", strTemp) Then
'            Set clsInterface = colInterface.Item(strTemp)
'            clsInterface.BuildVoucher domHead, domBody, cnnAccount, Login, strUserErr, bSuc
'            If bSuc = False Then
'                GoTo ExitSub
'            End If
'         End If
'     Next i
'
'     '注意：如果没有任何的用户自定义动态库，则需要返回TRUE
'     bSuc = True
'ExitSub:
'     Exit Sub
'Errhandle:
'     bSuc = False
'     strUserErr = Err.Description
'End Sub
'add by tianxl 2005-12-12 武汉公交插件
'单据审核不成功时调用（回滚）
'Public Sub CheckVoucher(ByVal domHead As Object, ByVal domBody As Object, ByVal cnnAccount As ADODB.Connection, ByVal Login As Object, strUserErr As String, bSuc As Boolean)
'On Error GoTo Errhandle
'
'     Dim i As Integer
'     For i = 0 To colInterface.Count - 1
'         '判断用户DLL是否有这个方法
'         strTemp = colInterface.Keys(i)
'         If bHasFunction("CheckVoucher", strTemp) Then
'            Set clsInterface = colInterface.Item(strTemp)
'            clsInterface.CheckVoucher domHead, domBody, cnnAccount, Login, strUserErr, bSuc
'            If bSuc = False Then
'                GoTo ExitSub
'            End If
'         End If
'     Next i
'
'     '注意：如果没有任何的用户自定义动态库，则需要返回TRUE
'     bSuc = True
'ExitSub:
'     Exit Sub
'Errhandle:
'     bSuc = False
'     strUserErr = Err.Description
'End Sub


'Public Sub CanSplit(ByVal domHead As Object, ByVal domBody As Object, ByVal cnnAccount As ADODB.Connection, ByVal Login As Object, strUserErr As String, bSuc As Boolean)
'On Error GoTo Errhandle
'
'     Dim i As Integer
'     For i = 0 To colInterface.Count - 1
'         '判断用户DLL是否有这个方法
'         strTemp = colInterface.Keys(i)
'         If bHasFunction("CanSplit", strTemp) Then
'            Set clsInterface = colInterface.Item(strTemp)
'            clsInterface.CanSplit domHead, domBody, cnnAccount, Login, strUserErr, bSuc
'            If bSuc = False Then
'                GoTo ExitSub
'            End If
'         End If
'     Next i
'
'     '注意：如果没有任何的用户自定义动态库，则需要返回TRUE
'     bSuc = True
'ExitSub:
'     Exit Sub
'Errhandle:
'     bSuc = False
'     strUserErr = Err.Description
'End Sub

'新增限额领料单的计划、删除、生单接口 add by tianxl 2005-12-15
'计划。
'Public Sub MatchPlan(ByVal domHead As Object, ByVal domBody As Object, ByVal cnnAccount As ADODB.Connection, ByVal Login As Object, strUserErr As String, bSuc As Boolean)
'On Error GoTo Errhandle
'
'     Dim i As Integer
'     For i = 0 To colInterface.Count - 1
'         '判断用户DLL是否有这个方法
'         strTemp = colInterface.Keys(i)
'         If bHasFunction("MatchPlan", strTemp) Then
'            Set clsInterface = colInterface.Item(strTemp)
'            clsInterface.MatchPlan domHead, domBody, cnnAccount, Login, strUserErr, bSuc
'            If bSuc = False Then
'                GoTo ExitSub
'            End If
'         End If
'     Next i
'
'     '注意：如果没有任何的用户自定义动态库，则需要返回TRUE
'     bSuc = True
'ExitSub:
'     Exit Sub
'Errhandle:
'     bSuc = False
'     strUserErr = Err.Description
'End Sub
'限额领料单删除
'Public Sub MatchDelete(ByVal domHead As Object, ByVal domBody As Object, ByVal cnnAccount As ADODB.Connection, ByVal Login As Object, strUserErr As String, bSuc As Boolean)
'On Error GoTo Errhandle
'
'     Dim i As Integer
'     For i = 0 To colInterface.Count - 1
'         '判断用户DLL是否有这个方法
'         strTemp = colInterface.Keys(i)
'         If bHasFunction("MatchDelete", strTemp) Then
'            Set clsInterface = colInterface.Item(strTemp)
'            clsInterface.MatchDelete domHead, domBody, cnnAccount, Login, strUserErr, bSuc
'            If bSuc = False Then
'                GoTo ExitSub
'            End If
'         End If
'     Next i
'
'     '注意：如果没有任何的用户自定义动态库，则需要返回TRUE
'     bSuc = True
'ExitSub:
'     Exit Sub
'Errhandle:
'     bSuc = False
'     strUserErr = Err.Description
'End Sub

'限额领料单生单
'Public Sub MatchGenVouch(ByVal domHead As Object, ByVal domBody As Object, ByVal cnnAccount As ADODB.Connection, ByVal Login As Object, strUserErr As String, bSuc As Boolean)
'On Error GoTo Errhandle
'
'     Dim i As Integer
'     For i = 0 To colInterface.Count - 1
'         '判断用户DLL是否有这个方法
'         strTemp = colInterface.Keys(i)
'         If bHasFunction("MatchGenVouch", strTemp) Then
'            Set clsInterface = colInterface.Item(strTemp)
'            clsInterface.MatchGenVouch domHead, domBody, cnnAccount, Login, strUserErr, bSuc
'            If bSuc = False Then
'                GoTo ExitSub
'            End If
'         End If
'     Next i
'
'     '注意：如果没有任何的用户自定义动态库，则需要返回TRUE
'     bSuc = True
'ExitSub:
'     Exit Sub
'Errhandle:
'     bSuc = False
'     strUserErr = Err.Description
'End Sub



'用户编辑界面，在的编辑界面里面直接与数据交互
'##ModelId=3FEA8A97026A
'Public Sub show_userfrm(ByVal obj As Object, ByVal Login As Object, strUserErr As String, bSuc As Boolean)
'On Error GoTo Errhandle
'     Dim i As Integer
'     For i = 0 To colInterface.Count - 1
'         '判断用户DLL是否有这个方法
'         strTemp = colInterface.Keys(i)
'         If bHasFunction("show_userfrm", strTemp) Then
'            Set clsInterface = colInterface.Item(strTemp)
'            clsInterface.show_userfrm obj, Login, strUserErr, bSuc
'            If bSuc = False Then
'                GoTo ExitSub
'            End If
'         End If
'     Next i
'     '注意：如果没有任何的用户自定义动态库，则需要返回TRUE
'     bSuc = True
'ExitSub:
'     Exit Sub
'Errhandle:
'     bSuc = False
'     strUserErr = Err.Description
'End Sub


'by lg080617  销售发货单是否推式生成销售出库单（事前）
'##ModelId=3FEA8A97023B
'Public Sub MakeOutVouch(ByVal domHead As Object, ByVal domBody As Object, ByVal cnnAccount As ADODB.Connection, ByVal Login As Object, strUserErr As String, bSuc As Boolean)
'On Error GoTo Errhandle
'     Dim i As Integer
'
'     For i = 0 To colInterface.Count - 1
''        MsgBox colInterface.Count
'         '判断用户DLL是否有这个方法
'         strTemp = colInterface.Keys(i)
''         MsgBox strTemp
'         If bHasFunction("MakeOutVouch", strTemp) Then
'            Set clsInterface = colInterface.Item(strTemp)
'            clsInterface.MakeOutVouch domHead, domBody, cnnAccount, Login, strUserErr, bSuc
'            If bSuc = False Then
'                GoTo ExitSub
'            End If
'         End If
'     Next i
'
'
'     '注意：如果没有任何的用户自定义动态库，则需要返回TRUE
'     bSuc = True
'ExitSub:
'     Exit Sub
'Errhandle:
'     bSuc = False
'     strUserErr = Err.Description
'End Sub

Private Sub Class_Initialize()

On Error GoTo ErrHandle
mbHasUserDLL = False

'sVoucherType = ""

Set domdll = New DOMDocument
Set colInterface = New Dictionary
'销售管理
'domdll.Load App.path & "\" & m_login.cSub_Id & "_UserConfig.xml"


'需求规划管理
If domdll.Load(App.Path & "\UI_UserConfig.xml") Then
    cls_Public.WrtDBlog DBConn, m_login.cUserId, "EFmain", "UI接口配置文件" & App.Path & "\UI_UserConfig.xml　加载成功！"
Else
    cls_Public.WrtDBlog DBConn, m_login.cUserId, "EFmain", "UI接口配置文件" & App.Path & "\UI_UserConfig.xml　加载失败！"
End If
For Each eleMent In domdll.selectNodes("config/dll/userdll")
    Set clsInterface = CreateObject(eleMent.Text)
    colInterface.Add eleMent.Text, clsInterface
    mbHasUserDLL = True
Next
      
ExitSub:
     Exit Sub
ErrHandle:
    
     strUserErr = Err.Description
End Sub

Private Sub Class_Terminate()
On Error GoTo ErrHandle
     Dim i As Integer

     For i = 0 To colInterface.Count - 1
         '判断用户DLL是否有这个方法
         strTemp = colInterface.Keys(i)
         Set clsInterface = colInterface.Item(strTemp)
         Set clsInterface = Nothing
     Next i


'    For Each clsInterface In colInterface
'        Set clsInterface = Nothing
'    Next
ExitSub:
     Exit Sub
ErrHandle:
     strUserErr = Err.Description
End Sub



Public Function bHasFunction(strFunName As String, Skey As String) As Boolean

'by dcb-判断是否有该方法，如果有则执行，否则退出。主要为了

On Error GoTo ErrHandle
    Dim mElement As IXMLDOMElement
    Dim mElement2 As IXMLDOMElement
    bHasFunction = False
    Dim i As Integer
    For i = 0 To domdll.selectNodes("config/dll").length - 1
        Set mElement = domdll.selectNodes("config/dll").Item(i)
        If mElement.selectSingleNode("userdll").Text = Skey Then
            For Each mElement2 In mElement.selectNodes("function/userfunction")
               If UCase(mElement2.Text) = UCase(strFunName) Then
                    bHasFunction = True
                    GoTo ExitSub
               End If
           Next
        End If
    Next i
'    MsgBox strFunName & "----" & sKey & "-----" & bHasFunction
ExitSub:
    Exit Function
ErrHandle:
   bHasFunction = False
End Function


Public Property Get bHasUserDLL() As Boolean
    bHasUserDLL = mbHasUserDLL
End Property

Public Property Let bHasUserDLL(ByVal vNewValue As Boolean)
    mbHasUserDLL = vNewValue
End Property

'判断表头参照是否进行二次开发处理
'Public Sub isBrowHeadUser(ByVal indexname As String, VouchType As String, bSuc As Boolean, Optional m_login As Object, Optional domHead As DOMDocument, Optional domBody As DOMDocument)
'On Error GoTo Errhandle
'     Dim i As Integer
'     For i = 0 To colInterface.Count - 1
'         '判断用户DLL是否有这个方法
'         strTemp = colInterface.Keys(i)
'         If bHasFunction("isBrowHeadUser", strTemp) Then
'            Set clsInterface = colInterface.Item(strTemp)
'            clsInterface.isBrowHeadUser indexname, VouchType, bSuc, m_login, domHead, domBody
'            If bSuc = False Then
'                GoTo ExitSub
'            End If
'         End If
'     Next i
'     '注意：如果没有任何的用户自定义动态库，则需要返回TRUE
'     bSuc = True
'ExitSub:
'     Exit Sub
'Errhandle:
'     bSuc = False
'     strUserErr = Err.Description
'End Sub

'判断表体参照是否进行二次开发处理
'Public Sub isBrowbodyUser(ByVal R As Long, ByVal indexname As String, Vouchtypes As String, bSuc As Boolean, Optional m_login As Object, Optional domHead As DOMDocument, Optional domBody As DOMDocument)
'On Error GoTo Errhandle
'     Dim i As Integer
'     For i = 0 To colInterface.Count - 1
'         '判断用户DLL是否有这个方法
'         strTemp = colInterface.Keys(i)
'         If bHasFunction("isBrowbodyUser", strTemp) Then
'            Set clsInterface = colInterface.Item(strTemp)
'            clsInterface.isBrowbodyUser R, indexname, Vouchtypes, bSuc, m_login, domHead, domBody
'            If bSuc = False Then
'                GoTo ExitSub
'            End If
'         End If
'     Next i
'     '注意：如果没有任何的用户自定义动态库，则需要返回TRUE
'     bSuc = True
'ExitSub:
'     Exit Sub
'Errhandle:
'     bSuc = False
'     strUserErr = Err.Description
'End Sub



'Public Sub Brow_HeadUser(ByVal indexname As Variant, sRet As Variant, ByVal m_login As Object, ByVal voucherTypes As String, domHead As DOMDocument, domBody As DOMDocument, strUserErr As String, bSuc As Boolean)
'On Error GoTo Errhandle
'     Dim i As Integer
'     For i = 0 To colInterface.Count - 1
'         '判断用户DLL是否有这个方法
'         strTemp = colInterface.Keys(i)
'         If bHasFunction("Brow_HeadUser", strTemp) Then
'            Set clsInterface = colInterface.Item(strTemp)
'            clsInterface.Brow_HeadUser indexname, sRet, m_login, voucherTypes, domHead, domBody, strUserErr, bSuc
'            If bSuc = False Then
'                GoTo ExitSub
'            End If
'         End If
'     Next i
'     '注意：如果没有任何的用户自定义动态库，则需要返回TRUE
'     bSuc = True
'ExitSub:
'     Exit Sub
'Errhandle:
'     bSuc = False
'     strUserErr = Err.Description
'End Sub
'
'Public Sub Brow_bodyUser(ByVal R As Variant, ByVal indexname As Variant, sRet As Variant, ByVal m_login As Object, ByVal voucherTypes As String, domHead As DOMDocument, domBody As DOMDocument, strUserErr As String, bSuc As Boolean)
'On Error GoTo Errhandle
'     Dim i As Integer
'     For i = 0 To colInterface.Count - 1
'         '判断用户DLL是否有这个方法
'         strTemp = colInterface.Keys(i)
'         If bHasFunction("Brow_bodyUser", strTemp) Then
'            Set clsInterface = colInterface.Item(strTemp)
'            clsInterface.Brow_bodyUser R, indexname, sRet, m_login, voucherTypes, domHead, domBody, strUserErr, bSuc
'            If bSuc = False Then
'                GoTo ExitSub
'            End If
'         End If
'     Next i
'     '注意：如果没有任何的用户自定义动态库，则需要返回TRUE
'     bSuc = True
'ExitSub:
'     Exit Sub
'Errhandle:
'     bSuc = False
'     strUserErr = Err.Description
'End Sub

'Public Sub setVoucher(m_login As Object, voucherTypes As String, vouchers As Object, Optional strUserErr As String, Optional bSuc As Boolean)
'On Error GoTo Errhandle
'     Dim i As Integer
'     For i = 0 To colInterface.Count - 1
'         '判断用户DLL是否有这个方法
'         strTemp = colInterface.Keys(i)
'         If bHasFunction("setVoucher", strTemp) Then
'            Set clsInterface = colInterface.Item(strTemp)
'            clsInterface.setVoucher m_login, voucherTypes, vouchers, strUserErr, bSuc
'            If bSuc = False Then
'                GoTo ExitSub
'            End If
'         End If
'     Next i
'     '注意：如果没有任何的用户自定义动态库，则需要返回TRUE
'     bSuc = True
'ExitSub:
'     Exit Sub
'Errhandle:
'     bSuc = False
'     strUserErr = Err.Description
'End Sub


'by ahzzd 20071105 单据打印插件接口
'Public Sub PrintVoucher(m_login As Object, voucherTypes As String, domHead As Object, domBody As Object, Optional strUserErr As String, Optional bSuc As Boolean, Optional Voucher As Object, Optional sysprint As Boolean = True)
'On Error GoTo Errhandle
'     Dim i As Integer
'     For i = 0 To colInterface.Count - 1
'         '判断用户DLL是否有这个方法
'         strTemp = colInterface.Keys(i)
'         If bHasFunction("PrintVoucher", strTemp) Then
'            Set clsInterface = colInterface.Item(strTemp)
'            clsInterface.PrintVoucher m_login, voucherTypes, domHead, domBody, strUserErr, bSuc, Voucher, sysprint
'            If bSuc = False Then
'                GoTo ExitSub
'            End If
'         End If
'     Next i
'     '注意：如果没有任何的用户自定义动态库，则需要返回TRUE
'     bSuc = True
'ExitSub:
'     Exit Sub
'Errhandle:
'     bSuc = False
'     strUserErr = Err.Description
'End Sub


'by ahzzd 20071105 单据打印插件接口
'Public Sub PreviewVoucher(m_login As Object, voucherTypes As String, domHead As Object, domBody As Object, Optional strUserErr As String, Optional bSuc As Boolean, Optional Voucher As Object, Optional sysprint As Boolean = True)
'On Error GoTo Errhandle
'     Dim i As Integer
'     For i = 0 To colInterface.Count - 1
'         '判断用户DLL是否有这个方法
'         strTemp = colInterface.Keys(i)
'         If bHasFunction("PreviewVoucher", strTemp) Then
'            Set clsInterface = colInterface.Item(strTemp)
'            clsInterface.PreviewVoucher m_login, voucherTypes, domHead, domBody, strUserErr, bSuc, Voucher, sysprint
'            If bSuc = False Then
'                GoTo ExitSub
'            End If
'         End If
'     Next i
'     '注意：如果没有任何的用户自定义动态库，则需要返回TRUE
'     bSuc = True
'ExitSub:
'     Exit Sub
'Errhandle:
'     bSuc = False
'     strUserErr = Err.Description
'End Sub


'by ahzzd 保存前插件接口插件接口
'Public Sub Before_Save(m_login As Object, DBconn As ADODB.Connection, Cardnumber As String, domHead As Object, domBody As Object, Optional strUserErr As String, Optional bSuc As Boolean)
'On Error GoTo Errhandle
'     Dim i As Integer
'     For i = 0 To colInterface.Count - 1
'         '判断用户DLL是否有这个方法
'         strTemp = colInterface.Keys(i)
'         If bHasFunction("Before_Save", strTemp) Then
'            Set clsInterface = colInterface.Item(strTemp)
'            clsInterface.Before_Save m_login, DBconn, Cardnumber, domHead, domBody, strUserErr, bSuc
'            If bSuc = False Then
'                GoTo ExitSub
'            End If
'         End If
'     Next i
'     '注意：如果没有任何的用户自定义动态库，则需要返回TRUE
'     bSuc = True
'ExitSub:
'     Exit Sub
'Errhandle:
'     bSuc = False
'     strUserErr = Err.Description
'End Sub

''by ahzzd 保存后插件接口插件接口
'Public Sub After_Save(m_login As Object, DBconn As ADODB.Connection, Cardnumber As String, Id As String, domHead As Object, domBody As Object, Optional strUserErr As String, Optional bSuc As Boolean)
'On Error GoTo Errhandle
'     Dim i As Integer
'     For i = 0 To colInterface.Count - 1
'         '判断用户DLL是否有这个方法
'         strTemp = colInterface.Keys(i)
'         If bHasFunction("After_Save", strTemp) Then
'            Set clsInterface = colInterface.Item(strTemp)
'            clsInterface.After_Save m_login, DBconn, Cardnumber, Id, domHead, domBody, strUserErr, bSuc
'            If bSuc = False Then
'                GoTo ExitSub
'            End If
'         End If
'     Next i
'     '注意：如果没有任何的用户自定义动态库，则需要返回TRUE
'     bSuc = True
'ExitSub:
'     Exit Sub
'Errhandle:
'     bSuc = False
'     strUserErr = Err.Description
'End Sub





'系统菜单按扭前事件
Public Sub Before_ButtonClick(Voucher As Object, ButtonKey As String, strUserErr As String, bSuc As Boolean)
On Error GoTo ErrHandle
     Dim i As Integer
     bSuc = False
     For i = 0 To colInterface.Count - 1
         '判断用户DLL是否有这个方法
         strTemp = colInterface.Keys(i)
         If bHasFunction("Before_ButtonClick", strTemp) Then
            Set clsInterface = colInterface.Item(strTemp)
            clsInterface.Before_ButtonClick Voucher, ButtonKey, strUserErr, bSuc
            If bSuc = True Then
                GoTo ExitSub
            End If
         End If
     Next i
     '注意：如果没有任何的用户自定义动态库，则需要返回False 继续系统流程
'     bSuc = False
ExitSub:
     Exit Sub
ErrHandle:
     bSuc = False
     strUserErr = Err.Description
End Sub
 
 
'系统菜单按扭后事件
Public Sub After_ButtonClick(Voucher As Object, ButtonKey As String, strUserErr As String, bSuc As Boolean)
On Error GoTo ErrHandle
     Dim i As Integer
     bSuc = False
     For i = 0 To colInterface.Count - 1
         '判断用户DLL是否有这个方法
         strTemp = colInterface.Keys(i)
         If bHasFunction("After_ButtonClick", strTemp) Then
            Set clsInterface = colInterface.Item(strTemp)
            clsInterface.After_ButtonClick Voucher, ButtonKey, strUserErr, bSuc
            If bSuc = True Then
                GoTo ExitSub
            End If
         End If
     Next i
     '注意：如果没有任何的用户自定义动态库，则需要返回False 继续系统流程
'     bSuc = False
ExitSub:
     Exit Sub
ErrHandle:
     bSuc = False
     strUserErr = Err.Description
End Sub


 
'提示：单据表头参照事件
Public Sub Voucher_headBrowUser(ByVal Voucher As Object, ByVal Skey As String, sRet As Variant, strUserErr As String, bSuc As Boolean)
On Error GoTo ErrHandle
     Dim i As Integer
     bSuc = False
     For i = 0 To colInterface.Count - 1
         '判断用户DLL是否有这个方法
         strTemp = colInterface.Keys(i)
         If bHasFunction("Voucher_headBrowUser", strTemp) Then
            Set clsInterface = colInterface.Item(strTemp)
            clsInterface.Voucher_headBrowUser Voucher, Skey, sRet, strUserErr, bSuc
            If bSuc = True Then
                GoTo ExitSub
            End If
         End If
     Next i
     '注意：如果没有任何的用户自定义动态库，则需要返回TRUE
'     bSuc = True
'    bSuc = False
ExitSub:
     Exit Sub
ErrHandle:
     bSuc = False
     strUserErr = Err.Description
End Sub


'表头校验事件
Public Sub Voucher_headCellCheck(ByVal Voucher As Object, Skey As String, RetValue As String, bChanged As Long, strUserErr As String, bSuc As Boolean)

On Error GoTo ErrHandle
     Dim i As Integer
     bSuc = False
     For i = 0 To colInterface.Count - 1
         '判断用户DLL是否有这个方法
         strTemp = colInterface.Keys(i)
         If bHasFunction("Voucher_headCellCheck", strTemp) Then
            Set clsInterface = colInterface.Item(strTemp)
            clsInterface.Voucher_headCellCheck Voucher, Skey, RetValue, bChanged, strUserErr, bSuc
            If bSuc = True Then
                GoTo ExitSub
            End If
         End If
     Next i
     '注意：如果没有任何的用户自定义动态库，则需要返回TRUE
'     bSuc = True
'    bSuc = False
ExitSub:
     Exit Sub
ErrHandle:
     bSuc = False
     strUserErr = Err.Description
End Sub



'表体参照事件
Public Sub Voucher_bodyBrowUser(ByVal Voucher As Object, Skey As String, ByVal R As Long, sRet As Variant, strUserErr As String, bSuc As Boolean)
On Error GoTo ErrHandle
     Dim i As Integer
     bSuc = False
     For i = 0 To colInterface.Count - 1
         '判断用户DLL是否有这个方法
         strTemp = colInterface.Keys(i)
         If bHasFunction("Voucher_bodyBrowUser", strTemp) Then
            Set clsInterface = colInterface.Item(strTemp)
            clsInterface.Voucher_bodyBrowUser Voucher, Skey, R, sRet, strUserErr, bSuc
            If bSuc = True Then
                GoTo ExitSub
            End If
         End If
     Next i
     '注意：如果没有任何的用户自定义动态库，则需要返回TRUE
'     bSuc = True
'    bSuc = False
ExitSub:
     Exit Sub
ErrHandle:
     bSuc = False
     strUserErr = Err.Description
End Sub



'表体校验事件
Public Sub Voucher_bodyCellCheck(ByVal Voucher As Object, RetValue As Variant, bChanged As Long, Skey As String, ByVal R As Long, strUserErr As String, bSuc As Boolean)
On Error GoTo ErrHandle
     Dim i As Integer
     bSuc = False
     For i = 0 To colInterface.Count - 1
         '判断用户DLL是否有这个方法
         strTemp = colInterface.Keys(i)
         If bHasFunction("Voucher_bodyCellCheck", strTemp) Then
            Set clsInterface = colInterface.Item(strTemp)
            clsInterface.Voucher_bodyCellCheck Voucher, RetValue, bChanged, Skey, R, strUserErr, bSuc
            If bSuc = True Then
                GoTo ExitSub
            End If
         End If
     Next i
     '注意：如果没有任何的用户自定义动态库，则需要返回TRUE
'     bSuc = True
'    bSuc = False
ExitSub:
     Exit Sub
ErrHandle:
     bSuc = False
     strUserErr = Err.Description
End Sub


'提示：单据表头列表式项目参照事件
Public Sub Voucher_FillHeadComboBox(ByVal Voucher As Object, Skey As String, pCom As Object, strUserErr As String, bSuc As Boolean)
On Error GoTo ErrHandle
     Dim i As Integer
     bSuc = False
     For i = 0 To colInterface.Count - 1
         '判断用户DLL是否有这个方法
         strTemp = colInterface.Keys(i)
         If bHasFunction("Voucher_FillHeadComboBox", strTemp) Then
            Set clsInterface = colInterface.Item(strTemp)
            clsInterface.Voucher_FillHeadComboBox Voucher, Skey, pCom, strUserErr, bSuc
            If bSuc = True Then
                GoTo ExitSub
            End If
         End If
     Next i
     '注意：如果没有任何的用户自定义动态库，则需要返回TRUE
'     bSuc = True
'    bSuc = False
ExitSub:
     Exit Sub
ErrHandle:
     bSuc = False
     strUserErr = Err.Description
End Sub



 '提示：单据表体列表式项目参照事件
Public Sub Voucher_FillList(ByVal Voucher As Object, Skey As String, ByVal R As Long, pCom As Object, strUserErr As String, bSuc As Boolean)
On Error GoTo ErrHandle
     Dim i As Integer
     bSuc = False
     For i = 0 To colInterface.Count - 1
         '判断用户DLL是否有这个方法
         strTemp = colInterface.Keys(i)
         If bHasFunction("Voucher_FillList", strTemp) Then
            Set clsInterface = colInterface.Item(strTemp)
            clsInterface.Voucher_FillList Voucher, Skey, R, pCom, strUserErr, bSuc
            If bSuc = True Then
                GoTo ExitSub
            End If
         End If
     Next i
     '注意：如果没有任何的用户自定义动态库，则需要返回TRUE
'     bSuc = True
'    bSuc = False
ExitSub:
     Exit Sub
ErrHandle:
     bSuc = False
     strUserErr = Err.Description
End Sub

 '提示：单据单击事件
Public Sub Voucher_Click(ByVal Voucher As Object, section As Variant, Skey As String, R As Long, strUserErr As String, bSuc As Boolean)
On Error GoTo ErrHandle

     Dim i As Integer
     bSuc = False
     For i = 0 To colInterface.Count - 1
         '判断用户DLL是否有这个方法
         strTemp = colInterface.Keys(i)
         If bHasFunction("Voucher_Click", strTemp) Then
            Set clsInterface = colInterface.Item(strTemp)
            clsInterface.Voucher_Click Voucher, section, Skey, R, strUserErr, bSuc
            If bSuc = True Then
                GoTo ExitSub
            End If
         End If
     Next i
     '注意：如果没有任何的用户自定义动态库，则需要返回TRUE
'     bSuc = True
'    bSuc = False
ExitSub:
     Exit Sub
ErrHandle:
     bSuc = False
     strUserErr = Err.Description
End Sub

 '提示：单据单击事件
Public Sub Voucher_DblClick(ByVal Voucher As Object, section As Variant, Skey As String, R As Long, strUserErr As String, bSuc As Boolean)
On Error GoTo ErrHandle

     Dim i As Integer
     bSuc = False
     For i = 0 To colInterface.Count - 1
         '判断用户DLL是否有这个方法
         strTemp = colInterface.Keys(i)
         If bHasFunction("Voucher_DblClick", strTemp) Then
            Set clsInterface = colInterface.Item(strTemp)
            clsInterface.Voucher_DblClick Voucher, section, Skey, R, strUserErr, bSuc
            If bSuc = True Then
                GoTo ExitSub
            End If
         End If
     Next i
     '注意：如果没有任何的用户自定义动态库，则需要返回TRUE
'     bSuc = True
'    bSuc = False
ExitSub:
     Exit Sub
ErrHandle:
     bSuc = False
     strUserErr = Err.Description
End Sub


Public Sub VchLst_DblClick(ByVal VchLst As Object, strUserErr As String, bSuc As Boolean)
    On Error GoTo ErrHandle

     Dim i As Integer
     bSuc = False
     For i = 0 To colInterface.Count - 1
         '判断用户DLL是否有这个方法
         strTemp = colInterface.Keys(i)
         If bHasFunction("VchLst_DblClick", strTemp) Then
            Set clsInterface = colInterface.Item(strTemp)
            clsInterface.VchLst_DblClick VchLst, strUserErr, bSuc
            If bSuc = True Then
                GoTo ExitSub
            End If
         End If
     Next i
     '注意：如果没有任何的用户自定义动态库，则需要返回TRUE
'     bSuc = True
'    bSuc = False
ExitSub:
     Exit Sub
ErrHandle:
     bSuc = False
     strUserErr = Err.Description
End Sub
'用户插件初始化
Public Sub Userdll_Init(g_busines As Object, m_login As Object, DBConn As Object, Frm As Object, CardNum As String, strUserErr As String, bSuc As Boolean)
 On Error GoTo ErrHandle
     Dim i As Integer
     bSuc = True
     For i = 0 To colInterface.Count - 1
         '判断用户DLL是否有这个方法
         strTemp = colInterface.Keys(i)
         If bHasFunction("Userdll_Init", strTemp) Then
            Set clsInterface = colInterface.Item(strTemp)
            clsInterface.Userdll_Init g_busines, m_login, DBConn, Frm, CardNum, strUserErr, bSuc
            If bSuc = False Then
                GoTo ExitSub
            End If
         End If
     Next i
     '注意：如果没有任何的用户自定义动态库，则需要返回TRUE
'     bSuc = True
ExitSub:
     Exit Sub
ErrHandle:
     bSuc = False
     strUserErr = Err.Description
End Sub


'单据数据装载后数事件
Public Sub LoadAfter_VoucherData(strUserErr As String, bSuc As Boolean)
 On Error GoTo ErrHandle
     Dim i As Integer
     bSuc = True
     For i = 0 To colInterface.Count - 1
         '判断用户DLL是否有这个方法
         strTemp = colInterface.Keys(i)
         If bHasFunction("LoadAfter_VoucherData", strTemp) Then
            Set clsInterface = colInterface.Item(strTemp)
            clsInterface.LoadAfter_VoucherData strUserErr, bSuc
            If bSuc = False Then
                GoTo ExitSub
            End If
         End If
     Next i
     '注意：如果没有任何的用户自定义动态库，则需要返回TRUE
'     bSuc = True
ExitSub:
     Exit Sub
ErrHandle:
     bSuc = False
     strUserErr = Err.Description
End Sub


 
