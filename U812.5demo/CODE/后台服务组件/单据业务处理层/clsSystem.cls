VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSystem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public objlogin As U8Login.clsLogin                   '全局Login对象
Private DAccInfor As New Scripting.Dictionary
Public dbSales As New ADODB.Connection
Public objsys As New SystemInfo.cSysInfo
''系统信息参数
Public CurrentAccID As String
Public cSysVersion As String
Public CurrentUserName As String, CurrentUserID As String
Public Connectstr As String, SysDBstr As String, SysPassword As String, sDBName As String
Public iWorkMonth As Integer, dWorkDate As String, iWorkYear As Integer
Public bAdmin As Boolean
Public sRemoteID As String, sComputerName As String ''工作站名
Public iCurMonth As Integer  ''当前会计月
   
''环境变量(本系统和相关系统的选项及公用选项)
Public scName As String             '单位名称
Public scAddName As String            '单位简称
Public scAddress As String        '单位地址
Public scPostCode As String         '邮政编码
Public scPhone As String           '电话
Public scFax As String             '传真
Public scEMail As String           'Email
Public scTax As String             '税号
Public scLegalPerson As String     '法人
Public scEcoProperty As String     '经济性质
Public scEcoStyle As String        '经济类型
Public scAccountant As String      '主管会计
Public scBussiness As String       '主营业务
Public scAddBussiness As String     '兼营业务
Public sbBussiness As Boolean       ''企业类型 FALSE-工业 TRUE-商
Public bExchange As Boolean         ''是否外币核算
Public bAuth_Cus As Boolean         '是否控制客户权限
Public bAuth_Dep As Boolean         '是否控制部门权限
Public bAuth_Per As Boolean         '是否控制业务员权限
Public bAuth_Inv As Boolean         '是否控制项目权限
Public bAuth_Wh As Boolean          '是否仓库权限

Public sUseMode As String       '调用者 "" :CS,"BS":BS AR:"AR"
Public sCurrencyName As String      ''本币名称
Public clsAuth As Object                    'U8RowAuthsvr.clsRowAuth

'
Public bARUse As Boolean
 
Public obj_EA As Object    ''审批流对象
Public bManualTrans As Boolean ''U860 增加是否手动事务
Public cModeCode As String
Public bVouOrderByDate As Boolean '单据是否按日期排序
Private DateST(1 To 12) As Date
Private DateEnd(1 To 12) As Date

''获得月份开始日期
Public Function getBeginDate(i) As Date
    If i >= 1 And i <= 12 Then
        getBeginDate = DateST(i)
    End If
End Function

''获得月份结束日期
Public Function getEndDate(i) As Date
    If i >= 1 And i <= 12 Then
        getEndDate = DateEnd(i)
    End If
End Function
 
'' 初始化环境变量
Private Function InitEnv(ErrMsg As String) As Boolean
    Dim strSql As String
    Dim Rs As New ADODB.Recordset
    Dim i As Long, iCount As Long

    On Error GoTo Errhandle
    InitEnv = False
    
    bAuth_Dep = IIf(LCase(GetSysOption("KI", "bAuth_Dep")) = "true", True, False)
    bAuth_Inv = IIf(LCase(GetSysOption("KI", "bAuth_Inv")) = "true", True, False)
    InitEnv = True
    Exit Function
     
Errhandle:
    InitEnv = False
    ErrMsg = err.Description
End Function



Public Function Init(obj As U8Login.clsLogin, Optional strErrMsg As String) As Boolean
    Dim strSql As String, sDateST As String
    Dim DbSys As New ADODB.Connection
    Dim bGetAccDate As Boolean
    Dim bInitEnv As Boolean
    Dim Rs As New ADODB.Recordset
    Dim tmpsDate As String
    On Error GoTo Errhandle
 
    If Not obj Is Nothing Then Set objlogin = obj
    
    Set clsAuth = objlogin.RowAuth
    If clsAuth Is Nothing Then
        Set clsAuth = CreateObject("U8RowAuthsvr.clsRowAuth")
        clsAuth.Init obj.UfDbName, obj.cUserId
    End If
    
    bManualTrans = False
        
    sComputerName = objsys.ComputerName
    CurrentAccID = objlogin.cacc_id
    CurrentUserID = objlogin.cUserId
    CurrentUserName = objlogin.cUserName
    iWorkMonth = objlogin.iMonth
    iWorkYear = CInt(objlogin.cIYear)
    dWorkDate = Format(objlogin.CurDate, "YYYY-MM-DD")
    SysPassword = IIf(objlogin.SysPassword = """", "", objlogin.SysPassword)
    Connectstr = objlogin.UfDbName
    bAdmin = objlogin.isAdmin
    sDBName = ExtractStr(Connectstr, "Initial Catalog=", ";") ''数据库名称
    '' DbSales 数据库对象针对公用函数有效,可以防止多次打开连接
    If dbSales.State = 1 Then dbSales.Close
    dbSales.ConnectionTimeout = 600
    dbSales.CommandTimeout = 1200
    dbSales.Open objlogin.UfDbName          ', , objlogin.SysPassword
    If bARUse = False Then
        iCurMonth = CurrentAccMonth(dbSales)                   ''获得当前会计月
    Else
        bGetAccDate = True
    End If
    bGetAccDate = GetAccDateZone
    cSysVersion = GetSysDicOption("AA", "cSysVersion", , dbSales)
    bInitEnv = InitEnv(strErrMsg)
    If bInitEnv = False Then
        strErrMsg = "获取系统参数发生错误！"
    End If
    Init = True
   Exit Function
Errhandle:
    Init = False
    strErrMsg = err.Description
    
End Function

Public Property Get saLogin() As U8Login.clsLogin
    Set saLogin = objlogin
End Property
Public Property Let saLogin(ByVal vNewValue As U8Login.clsLogin)
  Set objlogin = saLogin
End Property

'' 系统清理
Public Sub CloseSys()
On Error Resume Next
   dbSales.oClose
   Set dbSales = Nothing
   Set objsys = Nothing
   Erase DateST
   Erase DateEnd
   ClearCol LockDLID
   Set LockDLID = Nothing
   ClearCol LockSOCode
   Set LockSOCode = Nothing
End Sub


''获得每个会计月的起止期间
Private Function GetAccDateZone() As Boolean
  Dim CN As New ADODB.Connection
  Dim Rs As New ADODB.Recordset
  Dim strFilter As String
  Dim iCount As Long, i As Long
  On Error GoTo Errhandle
     strFilter = "Select dbegin,dEnd From ufsystem..UA_Period Where iYear=" & CInt(iWorkYear) & " And cAcc_Id='" & CurrentAccID & "'"
     'CN.Open SysDBstr, , SysPassword
     Set CN = dbSales
     If CN.State = 1 Then
        Rs.CursorLocation = 3
        Rs.ActiveConnection = CN
        Rs.Open strFilter, , 0, 1
     Else
        GetAccDateZone = False
        Exit Function
     End If
     If Rs.State = 1 Then
        If Rs.BOF And Rs.EOF Then
           GetAccDateZone = False
           Exit Function
        Else
           iCount = Rs.RecordCount
           Rs.MoveFirst
           For i = 1 To iCount
               DateST(i) = Rs.Fields("dbegin")    '开始日期
               DateEnd(i) = Rs.Fields("dend")   '结束日期
               Rs.MoveNext
           Next
           Rs.Close
           Set Rs = Nothing
           GetAccDateZone = True
           Exit Function
        End If
     Else
        GetAccDateZone = False
        Exit Function
     End If
     Exit Function
Errhandle:
     GetAccDateZone = False
End Function

 
'' 读取单据列表项目,缺省只过滤显示项
Public Function LoadSAListItem(sVouchType As String, Rs As ADODB.Recordset, ErrMsg As String, CN As ADODB.Connection) As Boolean
  Dim rst As New ADODB.Recordset
  Dim strSql As String
      LoadSAListItem = False
      strSql = " SELECT CardItemName, FieldName,FieldType,TableName From .InvoiceItem_sal" _
             & " WHERE CardNum = '" & sVouchType & "' AND ShowIt <> 0 ORDER BY CardItemNum "
      rst.ActiveConnection = CN
      rst.CursorLocation = 3
      rst.Open strSql, , 3, 1
      If rst.EOF And rst.BOF Then
         ErrMsg = "没有找到单据列表的定义"
         rst.Close
         Set rst = Nothing
         Exit Function
      Else
         Set Rs = rst
         Set Rs.ActiveConnection = Nothing
         LoadSAListItem = True
      End If
End Function


Public Function getDefName(CN As ADODB.Connection, strFldName As String, Optional cItemName As String) As String
     getDefName = getDefineName(CN, strFldName, cItemName)
End Function

Public Function GetVTID(CN As ADODB.Connection, strCardNum As String, Optional strUserID As String) As Long
    On Error GoTo Err_GetVTID
        Dim strSql As String
        Dim RecTemp As New ADODB.Recordset
        Dim clsAuth As New U8RowAuthsvr.clsRowAuth
        Dim strDJAuth As String
    If bAdmin = True Then
        strDJAuth = ""
    Else
        If clsAuth.Init(Connectstr, CurrentUserID) = False Then
            GetVTID = 0
        End If
        strDJAuth = clsAuth.getAuthString("DJMB", , "R")
    End If
    If strDJAuth = "1=2" Then
        GetVTID = 0
    Else
        If strDJAuth <> "" Then
            strSql = "select vt_id from vouchertemplates a left join vouchers b on (a.vt_cardnumber=b.cardnumber and a.vt_id=b.def_id) " & _
                    " where vt_cardnumber='" & strCardNum & "'and vt_id in (" & strDJAuth & ")and vt_templatemode=0 order by b.cardnumber desc, vt_id"
        Else
            strSql = "select vt_id from vouchertemplates a left join vouchers b on (a.vt_cardnumber=b.cardnumber and a.vt_id=b.def_id) " & _
                    " where vt_cardnumber='" & strCardNum & "'and vt_templatemode=0 order by b.cardnumber desc, vt_id"
        End If
        RecTemp.Open strSql, dbSales, adOpenForwardOnly, adLockReadOnly
        If Not RecTemp.EOF Then
            GetVTID = RecTemp(0)
        Else
            GetVTID = 0
        End If
    End If
    If RecTemp.State = 1 Then RecTemp.Close
    Set RecTemp = Nothing
    Exit Function
Err_GetVTID:
End Function
Public Function getDefaltVTID(CN As ADODB.Connection, strCardNum As String, Optional bPrint As Boolean = False) As String
    Dim strSql As String
    Dim RecTemp As New ADODB.Recordset
    
    On Error GoTo errDo
    strSql = "select def_id,def_id_prn from vouchers where cardnumber='" & strCardNum & "'"
    RecTemp.CursorLocation = adUseClient
    RecTemp.Open strSql, CN, adOpenForwardOnly, adLockReadOnly
    If RecTemp.RecordCount > 0 Then
        If bPrint = True Then
            getDefaltVTID = RecTemp!def_id_prn
        Else
            getDefaltVTID = RecTemp!def_id
        End If
    Else
        getDefaltVTID = 0
    End If
    RecTemp.Close
    Set RecTemp = Nothing
    Exit Function
errDo:
    If RecTemp.State = 1 Then RecTemp.Close
    Set RecTemp = Nothing
    getDefaltVTID = 0
End Function

''取假号
Public Function GetVoucherNO(strCardNum As String, domHead As DOMDocument, strVouchNo As String, ErrMsg As String, Optional DomFormat As DOMDocument, Optional bGetFormatOnly As Boolean, Optional bUseSelfFormat As Boolean) As Boolean
        GetVoucherNO = GetVouchNO(Connectstr, strCardNum, domHead, strVouchNo, ErrMsg, DomFormat, bGetFormatOnly, bUseSelfFormat, , sRemoteID, , , False)
End Function
''取真号
Public Function GetVoucherTrueNO(strVouchType As String, domHead As DOMDocument, strVouchNo As String, ErrMsg As String, Optional DomFormat As DOMDocument, Optional bGetFormatOnly As Boolean, Optional bUseSelfFormat As Boolean) As Boolean
        GetVoucherTrueNO = GetVouchNO(Connectstr, strVouchType, domHead, strVouchNo, ErrMsg, DomFormat, bGetFormatOnly, bUseSelfFormat, , sRemoteID, , , True)
End Function
Public Function GetMDefineName(CN As ADODB.Connection, strFldName As String) As String
    GetMDefineName = getDefineName(CN, strFldName)
End Function
Public Function GetMDefineID(strFldName As String) As String
    GetMDefineID = GetDefineID(strFldName)
End Function


Private Function GetSysDicOption(ByVal SysID As String, ByVal sKey As String, Optional ByVal bNewCollection As Boolean = False, Optional CN As ADODB.Connection) As String
    GetSysDicOption = ""
    GetSysDicOption = DAccInfor(UCase(SysID + "," + sKey))
End Function
Public Function GetSysOption(ByVal SysID As String, ByVal sKey As String, Optional ByVal bNewCollection As Boolean = False, Optional CN As ADODB.Connection) As String
    Dim strSql As String
    Dim rec As New ADODB.Recordset
    If bNewCollection Then
        CN.Open Connectstr  ', , SysPassword
    Else
        If IsMissing(CN) Then
            Set CN = dbSales
        Else
            If CN Is Nothing Then
                Set CN = dbSales
            End If
        End If
    End If
    If CN.State = 1 Then
        strSql = " Select cValue From AccInformation where cSysID='" & Trim(SysID) & "' and cName='" & Trim(sKey) & "' "
        rec.CursorLocation = adUseClient
        rec.ActiveConnection = CN
        rec.Open strSql, , 0, 1 ''仅向前游标,只读方式
        If rec.State = 1 Then
            If Not (rec.EOF And rec.BOF) Then
                If IsNull(rec(0)) Then
                    GetSysOption = ""
                Else
                    GetSysOption = rec(0)
                End If
                rec.Close
                Set rec = Nothing
            End If
        End If
    Else  ''连接未打开
        GetSysOption = ""
    End If
    If bNewCollection Then
        CN.Close
        Set CN = Nothing
    End If
End Function


''读取远程标志号
Private Sub ReadRemoteID(CN As ADODB.Connection)
    Dim rec As New ADODB.Recordset
    Dim strSql As String
     
    strSql = "Select cValue From AccInformation where cSysID='AA' and cName='cID'"
    If CN.State = 1 Then
        rec.Open strSql, CN, 0, 1
        If rec.State = 1 Then
            If Not (rec.EOF And rec.BOF) Then
                If Not IsNull(rec(0)) Then
                    sRemoteID = IIf(rec(0) = "", "00", rec(0))
                Else
                    sRemoteID = "00"
                End If
            Else  ''未找到自动补0
                sRemoteID = "00"
            End If
            rec.Close
            Set rec = Nothing
        End If
    End If
End Sub

'' 功能申请
'' 修改:增加参数 sErrMsg
'' 参数: taskID:功能号, sErrMsg:返回值错误信息
Public Function SAApplyTask(TaskID As String, sErrMsg As String) As Boolean
    On Error GoTo Errhandle
 
    If Not objlogin Is Nothing Then
        objlogin.ClearError
        If objlogin.TaskExec(TaskID, -1, iWorkYear) Then
            SAApplyTask = True
            Exit Function
        Else
            If objlogin.ShareString <> "" Then
                sErrMsg = objlogin.ShareString
            Else
                sErrMsg = "共享(网络)冲突或没有此项操作的权限."
            End If
          
            objlogin.ClearError
            SAApplyTask = False
            Exit Function
        End If
    Else
        sErrMsg = "系统管理或注册服务程序工作异常,不能进行功能申请."
        SAApplyTask = False
        Exit Function
    End If
    Exit Function
Errhandle:
    If err.Number <> 0 Then
        sErrMsg = "功能申请失败,请重试!" '"Function:SAApplyTask " & vbCrLf & err.Number & err.Description
    Else
        sErrMsg = "功能申请失败,请重试!"
    End If
  
End Function

'' 释放功能
'' 修改:增加参数 sErrMsg
'' 参数: taskID:功能号, sErrMsg:返回值错误信息
Public Function SAFreeTask(TaskID As String, sErrMsg As String) As Boolean
 On Error GoTo Errhandle
 
    If Not objlogin Is Nothing Then
       objlogin.ClearError
       If objlogin.TaskExec(TaskID, 0, iWorkYear) Then
          SAFreeTask = True
       Else
          If objlogin.ShareString <> "" Then
             sErrMsg = objlogin.ShareString
          Else
             sErrMsg = "共享(网络)冲突或没有此项操作的权限."
          End If
          
          objlogin.ClearError
          SAFreeTask = False
       End If
    Else
       sErrMsg = "系统管理或注册服务程序工作异常,不能进行功能释放."
       SAFreeTask = False
       Exit Function
    End If
    Exit Function
Errhandle:
    If err.Number <> 0 Then
       sErrMsg = "功能申请失败,请重试!" '"Function:SAFreeTask " & vbCrLf & err.Number & err.Description
    Else
       sErrMsg = "功能申请失败,请重试!"
    End If
End Function
Public Function SACheckAuth(TaskID As String, sErrMsg As String) As Boolean
 On Error GoTo Errhandle
    If Not objlogin Is Nothing Then
       objlogin.ClearError
       If objlogin.TaskExec(TaskID, 1, iWorkYear) Then
          SACheckAuth = True
          Exit Function
       Else
          SACheckAuth = False
          Exit Function
       End If
    Else
       sErrMsg = "系统管理或注册服务程序工作异常,不能进行权限检查."
       SACheckAuth = False
       Exit Function
    End If
    Exit Function
Errhandle:
    If err.Number <> "" Then
       sErrMsg = "Function:SACheckAuth " & vbCrLf & err.Number & err.Description
    End If
    
End Function


''按金额精度格式化
Public Function FormatMoney(ByVal iData As Double) As Double
 Dim StrConv As String
     StrConv = Format(iData, "###,###,###,###,##0.00")
     FormatMoney = CDbl(StrConv)
End Function


''获得汇率换算关系
Public Function IsCal(ByVal WBName As String) As Boolean
    On Error GoTo Errhandle
    
    Dim tmpRst As New ADODB.Recordset
    tmpRst.Open "select bcal from ForeignCurrency where cexch_name='" & WBName & "'", dbSales
    If tmpRst.BOF And tmpRst.EOF Then
        IsCal = True
    Else
        IsCal = CBool(tmpRst!bcal)
    End If
    tmpRst.Close
    Set tmpRst = Nothing
Errhandle:
End Function


Public Function GetExchRate(strName As String, ddate As Date, m_Login As U8Login.clsLogin) As Double '返回汇率；设置变量"Exch"(本币金额=原币金额*Exch)
    Dim dtmTmp As Date
    Dim iPeriod As Integer
    Dim lngI As Long
    Dim rstTmp As New ADODB.Recordset
    Dim rstFile As New ADODB.Recordset
    Dim strDate As String
    Dim strExchName  As String
    strExchName = strName
    strDate = str(ddate)
    If rstTmp.State <> 0 Then rstTmp.Close
    rstTmp.Open "select cValue from accinformation where cname= 'iXchgRateStl'", dbSales
    If InStr(strDate, "T") > 0 Then
        strDate = Left(strDate, InStr(strDate, "T") - 1)
    End If
    If CBool(rstTmp.Fields("cValue")) = True Then
        rstFile.Source = "select t1.cexch_name,t1.cexch_code,ISNULL(t2.nflat,1) AS nflat From " _
                    & "(SELECT cexch_name,cexch_code FROM ForeignCurrency Where cexch_name='" & strExchName & "' Or cexch_Code='" & strExchName & "') as t1 " _
                    & "Left Join " _
                    & "(select cexch_name,nflat from exch WHERE itype= '2' AND cdate='" & GetAccMonth(CDate(strDate), Me) & "') as t2 " _
                    & "on t1.cexch_name=t2.cexch_name"
    Else
        rstFile.Source = "select t1.cexch_name,t1.cexch_code,ISNULL(t2.nflat,1) AS nflat From " _
                    & "(SELECT cexch_name,cexch_code FROM ForeignCurrency Where cexch_name='" & strExchName & "' Or cexch_Code='" & strExchName & "') as t1 " _
                    & "Left Join " _
                    & "(select cexch_name,nflat from exch WHERE itype= '1' AND cdate='" & Format(strDate, "yyyy.mm.dd") & "') as t2 " _
                    & "on t1.cexch_name=t2.cexch_name"
    End If
    rstFile.Open , dbSales
    If Not (rstFile.BOF And rstFile.EOF) Then
        GetExchRate = rstFile.Fields("nflat")
    Else
        GetExchRate = 1
    End If
    rstTmp.Close
    rstFile.Close
    GetExchRate = IIf(GetExchRate = 0, 1, GetExchRate)
End Function
Public Function GetExchRateDec(strName As String) As Integer
    Dim tmpRst As New ADODB.Recordset
    
    On Error GoTo DoErr
    GetExchRateDec = 0
    If strName <> "" Then
        If strName = "人民币" Then
            GetExchRateDec = 8
            Exit Function
        Else
            tmpRst.Open "select isnull(idec,0) as idec from dbo.foreigncurrency where  cexch_name='" & strName & "'", dbSales, adOpenForwardOnly, adLockReadOnly
            If Not tmpRst.EOF Then
                GetExchRateDec = tmpRst(0)
            Else
                GetExchRateDec = 0
            End If
            tmpRst.Close
        End If
    Else
        GetExchRateDec = 0
    End If
    Set tmpRst = Nothing
    Exit Function
DoErr:
    err.Raise vbObjectError + 513, " clsSystem.GetExchRateDec", err.Description
End Function

Public Function CheckDJMBAuth(strVTID As String, strOprate As String) As Boolean
    Dim clsAuth As New U8RowAuthsvr.clsRowAuth
    If Not bAdmin Then
        clsAuth.Init dbSales.ConnectionString, CurrentUserID
        CheckDJMBAuth = clsAuth.IsHoldAuth("DJMB", strVTID, , strOprate)
    Else
        CheckDJMBAuth = True
    End If
    Set clsAuth = Nothing
End Function


''判断对单据是否有读、写权限
''不包括操作员权限
Public Function IsHoldRWAuth(strVouchType As String, lngID As String, ErrMsg As String, strOprator As String) As Boolean
    Dim clsAuth As U8RowAuthsvr.clsRowAuth
    Dim strHTable As String, strBTable As String, sKeyID As String
    Dim strAuth As String, strSql As String, strwhere As String
    Dim rstTmp As ADODB.Recordset
    On Error GoTo DoErr
    IsHoldRWAuth = False
    If bAdmin = True Then
        IsHoldRWAuth = True
        Exit Function
    End If
    
    Select Case strVouchType   '//xzq
        Case "87", "88", "89", "90", "96" '
            sKeyID = "id"
            strHTable = "MT_baseset"
            strBTable = "MT_baseset"
            
    End Select
    strwhere = " where " + strHTable + "." + sKeyID + "=" & lngID
    strSql = "select top 1 " + sKeyID + " from " + strHTable + " " + strwhere
    Set rstTmp = New ADODB.Recordset
    rstTmp.Open strSql, dbSales, adOpenForwardOnly, adLockReadOnly
    If rstTmp.EOF Then
        IsHoldRWAuth = False
        ErrMsg = "当前操作员对客户或者部门或者业务员没有权限，操作无效"
        rstTmp.Close
        Set rstTmp = Nothing
        Set clsAuth = Nothing
        Exit Function
    End If
    strwhere = " where " + strBTable + "." + sKeyID + "=" & lngID
    IsHoldRWAuth = True
    Set clsAuth = Nothing
    Set rstTmp = Nothing
    Exit Function
DoErr:
    IsHoldRWAuth = False
    ErrMsg = err.Description
    Set clsAuth = Nothing
    Set rstTmp = Nothing
End Function

'获得操作员对某基础档案的控制权限
Public Function strAuthDic(strTblName As String) As String
Dim clsAuth As U8RowAuthsvr.clsRowAuth
Dim strAuth As String
    Set clsAuth = New U8RowAuthsvr.clsRowAuth
    clsAuth.Init dbSales.ConnectionString, objlogin.cUserId

    If bAdmin Then
        strAuthDic = ""
        Set clsAuth = Nothing
        Exit Function
    End If
    strAuth = "1=1"
    If strTblName = "USER" Then
        strAuth = clsAuth.getAuthString(strTblName, , "W")
    Else
    End If
    Select Case strTblName
            Case "Inventory"
                        strAuthDic = IIf(strAuth = "", "", " and (inventory.iid is null or inventory.iid in (" & IIf(strAuth = "1=2", "''", strAuth) & "))")
        '    Case "Warehouse"
            
            Case "Department" '是否控制部门权限
                strAuthDic = IIf(strAuth = "", "", " and (department.cDepCode is null or department.cDepCode in (" & IIf(strAuth = "1=2", "''", strAuth) & "))")
   
            Case "Person" '是否控制业务员权限
                strAuthDic = IIf(strAuth = "", "", " and (person.cPersonCode is null or person.cPersonCode in (" & IIf(strAuth = "1=2", "''", strAuth) & "))")
        '    Case "Vendor"
            
            Case "Customer" '是否控制客户权限
                strAuthDic = IIf(strAuth = "", "", " and (Customer.iid is null or Customer.iid in (" & IIf(strAuth = "1=2", "''", strAuth) & "))")
    End Select
    Set clsAuth = Nothing
End Function

Public Property Get bSACreate() As Variant
    bSACreate = IIf(LCase(GetSysOption("ST", "bSACreat")) = "true", True, False)
End Property


Private Sub Class_Terminate()
    Set obj_EA = Nothing
End Sub


Public Function ReSetCuriMonth(Optional ErrMsg As String) As Boolean
    On Error GoTo errDo
    If bARUse = False Then
        iCurMonth = CurrentAccMonth(dbSales)
    End If
    ReSetCuriMonth = True
    Exit Function
errDo:
    ErrMsg = err.Description
End Function
Public Function OrderNumCheck(CN As ADODB.Connection, cSOCode As String) As String
    Dim Rs As New ADODB.Recordset
    Rs.Open "Select cSoCode From so_soMain Where cSoCode='" & cSOCode & "' and isnull(cCloser,'')='' and isnull(cVerifier,'') <>''", CN, adOpenForwardOnly, adLockReadOnly
    If (Rs.BOF And Rs.EOF) Then
        OrderNumCheck = "此单据号不存在"
    Else
        OrderNumCheck = ""
    End If
End Function


Public Function getVouchMainIDName(Cardnumber As String, Optional ByVal cnn As ADODB.Connection) As String
    Dim clsVouchLoads As New clsVouchLoad
    clsVouchLoads.Getvoucherinf cnn, Cardnumber, getVouchMainIDName
    Set clsVouchLoads = Nothing
End Function


Public Function GetVoucherTaskID(sKey As String, sVouchType As String, Optional bRed As Boolean) As String
    Dim strID As String
    If IsMissing(bRed) Then bRed = False
    Select Case sVouchType
        Case "87"
            Select Case sKey
                Case "查询"
                    strID = "KI2001"
                Case "增加", "复制", "删除", "修改"
                    strID = "KI2001"
            End Select
        Case "88"
            Select Case sKey
                Case "查询"
                    strID = "KI2002"
                Case "增加", "复制", "删除", "修改"
                    strID = "KI2002"
            End Select
        Case "89"
            Select Case sKey
                Case "查询"
                    strID = "KI2003"
                Case "增加", "复制", "删除", "修改"
                    strID = "KI2003"
            End Select
        Case "90"
            Select Case sKey
                Case "查询"
                    strID = "KI2004"
                Case "增加", "复制", "删除", "修改"
                    strID = "KI2004"
            End Select
        Case "96"
            Select Case sKey
                Case "查询"
                    strID = "KI2008"
                Case "增加", "复制", "删除", "修改"
                    strID = "KI2008"
            End Select
        Case "91"
            Select Case sKey
                Case "查询"
                    'LDX    2009-05-21  Modify  Beg
'                    strID = "MT300101"
                    strID = "KI300101"
                    'LDX    2009-05-21  Modify  End
                Case "增加", "复制", "删除", "修改"
                    'LDX    2009-05-21  Modify  Beg
'                    strID = "MT300102"
                    strID = "KI300102"
                    'LDX    2009-05-21  Modify  End
                Case "审核", "弃审"
                    'LDX    2009-05-21  Modify  Beg
'                    strID = "MT300103"
                    strID = "KI300103"
                    'LDX    2009-05-21  Modify  End
            End Select
        Case "92"
            Select Case sKey
                Case "查询"
                    'LDX    2009-05-21  Modify  Beg
'                    strID = "MT300201"
                    strID = "KI300201"
                    'LDX    2009-05-21  Modify  End
                Case "增加", "复制", "删除", "修改"
                    'LDX    2009-05-21  Modify  Beg
'                    strID = "MT300202"
                    strID = "KI300202"
                    'LDX    2009-05-21  Modify  End
                Case "审核", "弃审"
                    'LDX    2009-05-21  Modify  End
'                    strID = "MT300203"
                    strID = "KI300203"
                    'LDX    2009-05-21  Modify  End
            End Select
        Case "93"
            Select Case sKey
                Case "查询"
                    'LDX    2009-05-21  Modify  Beg
'                    strID = "MT300301"
                    strID = "KI300301"
                    'LDX    2009-05-21  Modify  End
                Case "增加", "复制", "删除", "修改"
                    'LDX    2009-05-21  Modify  Beg
'                    strID = "MT300302"
                    strID = "KI300302"
                    'LDX    2009-05-21  Modify  End
                Case "审核", "弃审"
                    'LDX    2009-05-21  Modify  Beg
'                    strID = "MT300303"
                    strID = "KI300303"
                    'LDX    2009-05-21  Modify  End
            End Select
        Case "94"
            Select Case sKey
                Case "查询"
                    strID = "KI300601"
                Case "增加", "复制", "删除", "修改"
                    strID = "KI300602"
                Case "审核", "弃审"
                    strID = "KI300603"
            End Select
        Case "95"
            Select Case sKey
                Case "查询"
                    strID = "KI300501"
                Case "增加", "复制", "删除", "修改"
                    strID = "KI300502"
                Case "审核", "弃审"
                    strID = "KI300503"
            End Select
            
    End Select
    GetVoucherTaskID = strID
End Function


'判断是否处于VB调试状态
Public Function IsDebug(Optional ByRef ChangeDebug As Boolean = True) As Boolean
    If ChangeDebug = False Then
        ChangeDebug = True:  IsDebug = True: Exit Function
    End If
    Dim T As Boolean
    T = False
    Debug.Assert IsDebug(T)
    IsDebug = T
End Function

Public Function ExecuteSql(conn As Object, ByVal strSql As String) As Object
    On Error GoTo errExit
    Set ExecuteSql = conn.Execute(strSql)
    Exit Function
errExit:
    OnErr , err.Description & vbCrLf & strSql
    VBA.err.Raise VBA.err.Number, VBA.err.Source, VBA.err.Description
End Function

'写文本文件
Public Function WriteText(ByVal fileName As String, ByVal value As String)
    Dim InFile As Integer   ' 文件的描述符。
    Dim sFile As String:    InFile = FreeFile
    Open fileName For Output As InFile
    Print #InFile, value:    Close InFile
End Function
'写入错误信息
Public Function WriteErrLog(ByVal fileName As String, oErr As Object)
    WriteText fileName, "Error:(" & Now & ")" & vbCrLf & "ErrNumber:" & oErr.Number & vbCrLf & "ErrSource:" & oErr.Source & vbCrLf & "ErrDescription:" & oErr.Description
End Function
'简单错误处理方式，在当前目录下创建err目录，生成日期时间加顺序号的log文件
Public Sub OnErr(Optional oErr As Object = Nothing, Optional strErr As String = "")
    Dim strPath As String: strPath = App.Path & "\Err\"
    'On Error Resume Next
    If Dir(strPath, vbDirectory) = "" Then MkDir strPath
    If oErr Is Nothing Then
        WriteText strPath & Format(Date, "yyyymmddhhMMss") & "_" & CStr(Rnd(1) * 1000) & ".log", "Error:(" & Now & ")" & vbCrLf & strErr
    Else
        WriteErrLog strPath & Format(Date, "yyyymmddhhMMss") & "_" & CStr(Rnd(1) * 1000) & ".log", oErr
    End If
End Sub
