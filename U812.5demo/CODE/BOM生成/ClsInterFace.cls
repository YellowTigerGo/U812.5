VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsInterFace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements UFPortalProxyInterface.ILoginable

Private objLoginable As UFPortalProxyInterface.ILoginable
Public Business2 As Object

Public Property Get Business() As Object
    Set Business = g_oBusiness
End Property

Public Property Set Business(ByRef bBus As Object)
    Set g_oBusiness = bBus  '获得business对象。
End Property
Public Function ILoginable_Login() As Boolean
    If ((Not g_oBusiness Is Nothing) And g_bLogined = False) Then
        Set g_oLogin = g_oBusiness.GetVBclsLogin() '获得login对象
        If g_Conn.State = 1 Then g_Conn.Close
        g_Conn.ConnectionString = g_oLogin.UfDbName
        g_Conn.CursorLocation = adUseClient
        g_Conn.Open

        g_bLogined = True
        ILoginable_Login = True
    Else
        ILoginable_Login = Not g_oBusiness Is Nothing
    End If

End Function
Public Function ILoginable_LogOff() As Boolean
    '程序退出前,添加自己的代码

    ILoginable_LogOff = True '本行代码必须。
    g_bLogined = False '本行代码必须。
    Unload FrmVoucher
End Function
Private Sub CClick()
  Dim ologin As Object
  Dim strSql As String
  Dim oDomHead As New DOMDocument
  Dim oDomBody As New DOMDocument
  Dim rs As New ADODB.Recordset
  Dim objc As New InvokeApiA
'  Dim gConn1 As New ADODB.Connection
'  Dim hyDZJA As Object
  Dim strMesg As String
'   set g_Conn.ConnectionString = "PROVIDER=SQLOLEDB;data source=4FB234B0B92744F;user id=sa;password="sa";initial catalog=UFDATA_999_2010;Connect Timeout=30;Persist Security Info=True ;Current Language=Simplified Chinese;

  
'  dim
'  set olobin as New U8Login.clsLogin
'  Set ologin = CreateObject("U8Login.clsLogin")
'  If Not ologin.login("AS", "(default)@999", "2010", "demo", "", "2010-02-01", "4FB234B0B92744F") Then
'      MsgBox ologin.ShareString
'  Else
'     g_Conn.ConnectionString = ologin.UfDbName
'    g_Conn.CursorLocation = adUseClient
'    g_Conn.Open
    
    strSql = "select *,'' as editprop from   V_HY_DZ_BorrowOut where ID=16"
    If rs Is Nothing Then Set rs = New ADODB.Recordset
    If rs.State = adStateOpen Then Call rs.Close
    Call rs.Open(strSql, g_Conn, adOpenDynamic, adLockReadOnly, adCmdText)
    rs.Save oDomHead, adPersistXML
    
    strSql = " select * ,'' as editprop from V_HY_DZ_BorrowOuts where id=16"
    If rs Is Nothing Then Set rs = New ADODB.Recordset
    If rs.State = adStateOpen Then Call rs.Close
    Call rs.Open(strSql, g_Conn, adOpenStatic, adLockReadOnly, adCmdText)
    rs.Save oDomBody, adPersistXML
'    Set hyDZJA = CreateObject("HY_DZ_BorrowOut.InvokeAPi")
    If objc.VoucherDelete(g_oLogin, 16) = True Then
        MsgBox GetString("U8.DZ.JA.Res1370"), vbInformation, GetString("U8.DZ.JA.Res030")
    Else
        MsgBox strMesg, vbInformation, GetString("U8.DZ.JA.Res030")
    End If
    

'  End If
      

End Sub

Public Function ILoginable_CallFunction(ByVal cMenuId As String, ByVal cMenuName As String, ByVal cAuthId As String, ByVal cCmdLine As String) As Object
    '开始执行业务代码
    'cMenuID、cmenuName、cAuthId、cCmlLine均为前述在UA_Menu表中注册的cMenu_Id、cMenu_Name、cAuth_Id、Paramters字段的值。
    
    '加密判断 08-03-25
    Dim Encryption As Object
    Set Encryption = New clsEncryption
    '此处要改为程序里定义的login对象
    'Encryption.Init g_ologin
'    If Encryption.IsOutOfLimit Then
'       ' MsgBox "请购买用友U8 V10.0【借出归还管理】", vbCritical, GetString("U8.DZ.JA.Res030")
'         ReDim varArgs(1)
'
'            varArgs(0) = g_oLogin.cAcc_Id
'            varArgs(1) = g_oLogin.cIYear
'            MsgBox GetStringPara("U8.AA.LoginBO.DemoInvalidate", varArgs(0), varArgs(1)), vbExclamation, GetString("U8.DZ.MRP_MOSub.Res030")
'        Exit Function
'    End If
    
'        If DateDiff("D", "2016-01-30", g_oLogin.CurDate) >= 0 Then
'        MsgBox "演示期已到!", vbCritical, "项目管理"
'        Exit Function
'    End If
'
'    If DateDiff("D", "2016-01-30", Date) >= 0 Then
'        MsgBox "演示期已到!", vbCritical, "项目管理"
'        Exit Function
'    End If
'    '权限查看
    If ZwTaskExec(g_oLogin, AuthBrowse, 1) = False Then Exit Function
    
'    If checkJCJYStart(g_Conn) = False Then Exit Function
'
'    '检查表是否建立 chenliangc
'    If checkTableExist(g_Conn, "HY_DZ_BorrowOut") = False Then
'        Exit Function
'    End If
'    If checkTableExist(g_Conn, "HY_DZ_BorrowOuts") = False Then
'        Exit Function
'    End If
    
    Call GlobalInit(g_oLogin)  '初始化系统信息 chenliangc

    InitMulText
    
    Dim vfd As Object
    
    '    获取权限字段
    Call getAuthString(g_Conn)

    '控制过滤、参照等不会因为数据权限而无法显示数据 11-7-12
    g_oLogin.AuthString = "warehouse=True,department=True,inventory=True,person=False,user=True,vendor=True,purchasetype=False,customer=True,position=True"
    
    Dim cmdArry
    Dim cCmdLineArr As Variant
     Dim strTemp As String
    '消息门户传回的审批流信息
'    If InStr(1, cCmdLine, gstrCardNumber) > 0 Then
'        cmdArry = Split(cCmdLine, vbTab)
'        cCmdLine = cmdArry(1)
'        lngVoucherID = CLng(Trim(cCmdLine))
'        sID = lngVoucherID
'
'    Else   '关联单据调用 by zhangwchb 20110718
'       If Len(cCmdLine) > 10 Then
'
'            '单据联查 时 按钮状态控制 by zhangwchb 20110809
'
'            cmdArry = Split(cCmdLine, vbTab)
''            tmpLinkTbl = cmdArry(2)
'            cCmdLine = cmdArry(1)
'            sID = cCmdLine
'
''        cCmdLine = Mid(cCmdLine, 2, 10)
'       End If
'    End If
'    CClick
    '由列表进入,并且单据界面已经打开
    isfromcon = False
    If Left(Trim(cCmdLine), 8) = "FYSL0006" Then
      isfromcon = True
      conid = Replace$(cCmdLine, "FYSL0006", "")
      
     cCmdLine = ""
     End If
    
    
    If cCmdLine <> "" Then
        Debug.Print "cCmdLine=" & cCmdLine
        If InStr(1, cCmdLine, Chr(9)) > 0 Then
            cCmdLineArr = Split(cCmdLine, Chr(9))
               ' If Replace(cCmdLineArr(0), Chr(9), "") = "" or Not IsNumeric(Trim(Replace(cCmdLineArr(0), Chr(9), ""))) Then
            If Not IsNumeric(Trim(Replace(cCmdLineArr(0), Chr(9), ""))) Then
                strTemp = Replace(cCmdLineArr(1), Chr(9), "")
            Else
                strTemp = Replace(cCmdLineArr(0), Chr(9), "")
                If Trim(Replace(cCmdLineArr(1), Chr(9), "")) <> "" Then
                    sTmpTableName = Trim(Replace(cCmdLineArr(1), Chr(9), ""))
                End If
            End If
            If strTemp <> "" Then
              
                    cCmdLine = GetcCode(strTemp)
                    
                    If cCmdLine = "" Then
                        MsgBox GetString("U8.pu.prjpu860.01614"), vbInformation, GetString("U8.DZ.JA.Res030")
                        Exit Function
                    End If
            End If
        End If
        '更新
        lngVoucherID = CLng(cCmdLine)
        sID = lngVoucherID
        Call UpdatePageCurrent(lngVoucherID)
        If PageCurrent <> 0 Then
            FrmVoucher.bexitload = True
            FrmVoucher.ExecSubRefresh
            '单据模板数据权限控制
            If FrmVoucher.bexitload = False Then
                Unload FrmVoucher
                Exit Function
            End If
            FrmVoucher.SetSearchState4List
        Else
            MsgBox GetString("U8.pu.prjpu860.01614"), vbInformation, GetString("U8.DZ.JA.Res030")
            Exit Function
        End If
    End If
    If sGUID = "" Then
        sGUID = CreateGUID
    End If
    
    If Not (g_oBusiness Is Nothing) Then
      gsGUIDForVouch = "HY_FYSL_Payment_FrmVoucher"
      Set vfd = g_oBusiness.CreateFormEnv(gsGUIDForVouch, FrmVoucher) 'Form11_key为任意字符串，建议以"窗体名_key"来表示，既防止重复，又清晰易懂。
      Call g_oBusiness.ShowForm(FrmVoucher, gsGUIDForVouch, "dp", False, True, vfd)
      
      
       If isfromcon = True Then
        Call FrmVoucher.ExecSubAdd
      
        If FrmVoucher.processdataforcon = False Then
                    
            Call FrmVoucher.ExecSubRefresh
            mOpStatus = SHOW_ALL

            Call SetCtlStyle(FrmVoucher, FrmVoucher.Voucher, FrmVoucher.Toolbar, FrmVoucher.UFToolbar, mOpStatus)
        End If
    
    End If
      
    Else
      FrmVoucher.Show
      FrmVoucher.ZOrder 0
    End If

    ZwTaskExec g_oLogin, AuthBrowse, 0

End Function

Public Function GetcCode(ByVal int1 As Long) As String
    Dim rs As New ADODB.Recordset
    Dim sql As String
    Set rs = Nothing
    'HY_DZ_BorrowOut
    sql = "select * from " & MainTable & "  where ID=" & int1 & IIf(sAuth_ALL = "", "", " and " & sAuth_ALL)
    rs.Open sql, g_Conn

    If rs.RecordCount > 0 Then
        GetcCode = rs.Fields("id")
    Else
        GetcCode = ""
    End If

    Set rs = Nothing
End Function

'Public Function Execute(e As ActionArgs)
'On Error GoTo Err
'    Dim str As String
'    Dim objVoucherInterface As Object
'    Dim SQL As String
'    Dim Rs As New ADODB.Recordset
'    Dim id As Long
'
'    id = e.RelateData.GetData("ID")
'
'    Set objVoucherInterface = CreateObject("HY_DZ_BorrowOut.ClsInterFace")
'    Set objLoginable = objVoucherInterface
'    Set objVoucherInterface.Business = g_oBusiness
'
'    If Not objLoginable.login() Then
'        MsgBox "登录验证失败，单据联查失败。", vbExclamation
'        Exit Function
'    End If
'
'    Call objLoginable.CallFunction("ST02JC0201", "借出借用单", "ST02JC0201", id)
'
'    Exit Function
'Err:
'   MsgBox "请选择数据行！", vbExclamation, getstring("U8.DZ.JA.Res030")
'End Function

'Public Function Execute(e As ActionArgs)
'On Error GoTo Err
'    Dim str As String
'    Dim objVoucherInterface As Object
'    Dim SQL As String
'    Dim Rs As New ADODB.Recordset
'    Dim id As Long
'
'    id = e.RelateData.GetData("ID")
'    Call ILoginable_CallFunction("ST02JC0201", "借出借用单", "ST02JC0201", id)
'
'    Exit Function
'Err:
'   MsgBox "请选择数据行！", vbExclamation, getstring("U8.DZ.JA.Res030")
'End Function
 
'是否样品借出，其他单据联查样品借出时，给出提示：样品借出单，请到CRM系统中查看!
Public Function IsSampleBorrowOut(ByVal int1 As Long) As String
    Dim rs As New ADODB.Recordset
    Dim sql As String
    Set rs = Nothing
    IsSampleBorrowOut = False
    'HY_DZ_BorrowOut
    sql = "select ID from " & MainTable & "  where ID=" & int1 & " and isnull(cborrowouttype,0)=1"
    rs.Open sql, g_Conn

    If rs.RecordCount > 0 Then
        IsSampleBorrowOut = True
    End If
    Set rs = Nothing
End Function

